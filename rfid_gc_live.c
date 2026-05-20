// rfid_gc_live.c
// Live-state RFID scanner using two antennas (Source_0 and Source_1).
//
// Both antennas are driven at maximum physical speed: there is NO
// artificial sleep between sweeps, so each antenna calls
// CAENRFID_InventoryTag back-to-back as fast as the reader will reply
// over the serial link. To keep the terminal readable the program
// prints ONE SUMMARY LINE PER SECOND containing:
//   - the measured scan rate per antenna in scans/second:  [S0=N/s S1=N/s]
//   - the most-recent non-empty observation per antenna seen during
//     that 1-second window (latest RSSI, latest phase, latest EPC).
//
// The output uses FIXED SLOTS so tags never shift positions:
//
// Empty (no tags in range during that second — still shows the rate):
//   [S0=137/s S1=137/s] []
//
// Tags visible (slot 0 is always antenna 0, slot 1 is always antenna 1).
// Per-tag RSSI is printed in brackets right after the antenna number
// in real dBm (one decimal place), followed by the backscatter PHASE
// in degrees (also one decimal), exactly as reported by the reader
// (no filtering, no arbitration). Empty slots render as pure whitespace
// so the comma and the other slot never shift columns:
//   [TX=30 mW] [S0=137/s S1=137/s] [(0)(-63.1)(p123.4) E2801160600002054E1A1234,   (1)(-65.0)(p210.8) E2801160600002054E1A5678]
//   [TX=30 mW] [S0=137/s S1=137/s] [(0)(-63.1)(p123.4) E2801160600002054E1A1234,                                              ]
//   [TX=30 mW] [S0=137/s S1=137/s] [                                            ,   (1)(-65.0)(p210.8) E2801160600002054E1A5678]
//
// The reader reports RSSI in tenths of dBm internally (e.g. raw -650
// == -65.0 dBm); we just divide by 10.0 for display.
//
// Phase: the CAEN easy2read protocol returns AVP_PHASE as a 16-bit value
// from the Impinj E310 radio (R3100C Lepton3). On Impinj chips that
// value is a 12-bit unsigned phase angle (0..4095 ↔ 0..360°), so we
// display raw * 360.0 / 4096.0. Unwrapped phase is what tells you the
// fine-grained tag distance (Δd = λ·Δφ / (4π)); λ ≈ 32.8 cm at 915 MHz
// means one full 360° wrap ≈ 16.4 cm of round-trip motion.
//
// Antenna index in YELLOW; Src0 tag EPC in GREEN, Src1 tag EPC in RED.
//
// Usage:
//   ./rfid_gc_live              -> both antennas at default power
//   ./rfid_gc_live <mW>         -> both antennas at <mW> (global power)
//   ./rfid_gc_live -h | --help  -> show usage

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdint.h>
#include <stdbool.h>
#include <unistd.h>
#include <signal.h>
#include <time.h>

#include "CAENRFIDLib_Light.h"
#include "host.h"

// Configuration
#define GC_PORT             "/dev/ttyACM0"
#define GC_BAUDRATE         921600
#define DEFAULT_POWER_MW    30            // sensible default for ~7 cm read zone
#define MIN_POWER_MW        1             // reader rejects below its hardware floor
#define MAX_POWER_MW        316           // R3100C Lepton3 max (25 dBm)
#define GC_RATE_WINDOW_MS   1000          // print one summary line every N ms (scan rate is reported over this window)
#define GC_MAX_TAGS         64            // max tags merged across both antennas per sweep
#define ANTENNA_COUNT       2
#define MAX_ID_LENGTH       64

// ANSI colours
#define GREEN  "\033[0;32m"
#define RED    "\033[0;31m"
#define YELLOW "\033[0;33m"
#define CYAN   "\033[0;36m"
#define RESET  "\033[0m"

volatile int running = 0;

typedef struct {
    char    tag[2 * MAX_ID_LENGTH + 1];
    int     antenna;
    int16_t rssi;        /* tenths of dBm, as reported by the reader */
    int16_t phase;       /* raw AVP_PHASE value (12-bit on Impinj E310) */
} TagEntry;

static void hex_str(uint8_t *bytes, uint16_t len, char *out) {
    for (int i = 0; i < len; i++)
        sprintf(out + (i * 2), "%02X", bytes[i]);
    out[len * 2] = '\0';
}

static void usage(const char *prog) {
    fprintf(stderr,
        "Usage:\n"
        "  %s              both antennas at %d mW (default)\n"
        "  %s <mW>         both antennas at <mW> (global power)\n"
        "  %s -h | --help  show this message\n"
        "\nValid power range: %d..%d mW\n",
        prog, DEFAULT_POWER_MW, prog, prog,
        MIN_POWER_MW, MAX_POWER_MW);
}

static bool parse_power(const char *s, uint32_t *out) {
    if (s == NULL || *s == '\0') return false;
    char *end = NULL;
    long v = strtol(s, &end, 10);
    if (end == s || *end != '\0') return false;
    if (v < MIN_POWER_MW || v > MAX_POWER_MW) return false;
    *out = (uint32_t)v;
    return true;
}

static void handle_sigint(int sig) {
    (void)sig;
    printf("\n" YELLOW "[GC] Stopping..." RESET "\n");
    running = 0;
}

// Width (visible chars, ignoring ANSI colour codes) reserved for each
// antenna slot inside the brackets. Comfortably fits
// "(N)(-XXX.X)(pYYY.Y) " + a 24-char EPC-96 hex string; longer tags
// overflow without truncation.
#define SLOT_WIDTH 46

// Convert the reader's raw 16-bit AVP_PHASE value into degrees.
// The R3100C Lepton3 is built on the Impinj E310 radio, whose phase
// register is 12-bit unsigned: 0..4095 maps linearly to 0..360°. If
// you ever swap to a reader that returns phase in a different scale
// (e.g. centidegrees), just change this single constant.
#define PHASE_RAW_FULLSCALE 4096.0
static inline double phase_deg(int16_t raw)
{
    /* AVP_PHASE is read as a uint16_t; the cast keeps it non-negative
       even though the struct field is int16_t for legacy reasons. */
    return ((uint16_t)raw) * 360.0 / PHASE_RAW_FULLSCALE;
}

// Summary output for one rate-window (default 1 s):
//   - "[S0=N/s S1=N/s] []" when neither antenna saw any tag during the window
//   - otherwise: [TX=…] [S0=N/s S1=N/s] [<slot0>,   <slot1>]
//     Each slot is padded to SLOT_WIDTH. If an antenna missed for the
//     entire window, its slot is pure whitespace (no "(N)"), so the
//     comma and the other slot keep their column positions and nothing
//     shifts. The scans/second figure is measured over the window: it
//     is the count of CAENRFID_InventoryTag calls per antenna divided
//     by the elapsed window time.
static void print_sweep_line(uint32_t power,
                             TagEntry bucket[ANTENNA_COUNT][GC_MAX_TAGS],
                             const int cnt[ANTENNA_COUNT],
                             const int rate[ANTENNA_COUNT])
{
    if (cnt[0] == 0 && cnt[1] == 0) {
        printf(CYAN "[S0=%d/s S1=%d/s]" RESET " []\n", rate[0], rate[1]);
        fflush(stdout);
        return;
    }

    printf(CYAN "[TX=%u mW]" RESET " "
           CYAN "[S0=%d/s S1=%d/s]" RESET " [",
           (unsigned)power, rate[0], rate[1]);

    for (int ant = 0; ant < ANTENNA_COUNT; ant++) {
        if (ant > 0)
            printf(",   "); /* fixed separator between the two slots */

        if (cnt[ant] == 0) {
            printf("%*s", SLOT_WIDTH, "");
            continue;
        }

        /* Visible width of the slot content (excludes ANSI codes) so we
           can right-pad to SLOT_WIDTH and keep the ']' column stable.
           Reader RSSI is in tenths of dBm, so we display it as a real
           dBm value with one decimal (e.g. raw -631 -> "-63.1"). The
           Phase is decoded into degrees with one decimal as well. */
        int visible = 3; /* "(N)" */
        for (int i = 0; i < cnt[ant]; i++) {
            char rbuf[24];
            int rlen = snprintf(rbuf, sizeof rbuf,
                                "(%.1f)(p%.1f) ",
                                bucket[ant][i].rssi / 10.0,
                                phase_deg(bucket[ant][i].phase));
            if (i > 0) visible += 1; /* space between multiple tags */
            visible += rlen + (int)strlen(bucket[ant][i].tag);
        }

        const char *tagcol = (ant == 0) ? GREEN : RED;
        printf(YELLOW "(%d)" RESET, ant);
        for (int i = 0; i < cnt[ant]; i++) {
            if (i > 0) printf(" ");
            printf("(%.1f)(p%.1f) %s%s" RESET,
                   bucket[ant][i].rssi / 10.0,
                   phase_deg(bucket[ant][i].phase),
                   tagcol, bucket[ant][i].tag);
        }
        int pad = SLOT_WIDTH - visible;
        if (pad > 0) printf("%*s", pad, "");
    }
    printf("]\n");
    fflush(stdout);
}

int main(int argc, char **argv) {

    uint32_t power = DEFAULT_POWER_MW;

    if (argc == 2) {
        if (strcmp(argv[1], "-h") == 0 || strcmp(argv[1], "--help") == 0) {
            usage(argv[0]);
            return 0;
        }
        if (!parse_power(argv[1], &power)) {
            usage(argv[0]);
            return 1;
        }
    } else if (argc != 1) {
        usage(argv[0]);
        return 1;
    }

    CAENRFIDErrorCodes ec;
    CAENRFIDReader reader = {
        .connect       = _connect,
        .disconnect    = _disconnect,
        .tx            = _tx,
        .rx            = _rx,
        .clear_rx_data = _clear_rx_data,
        .enable_irqs   = _enable_irqs,
        .disable_irqs  = _disable_irqs
    };

    RS232_params port_params = {
        .com         = GC_PORT,
        .baudrate    = GC_BAUDRATE,
        .dataBits    = 8,
        .stopBits    = 1,
        .parity      = 0,
        .flowControl = 0,
    };

    const char *sources[ANTENNA_COUNT] = { "Source_0", "Source_1" };
    char model[64]  = {0};
    char serial[64] = {0};

    signal(SIGINT, handle_sigint);

    printf(CYAN "===== Dual-Antenna RFID Live Scanner =====" RESET "\n");
    printf("Port      : %s @ %d baud\n", GC_PORT, GC_BAUDRATE);
    printf("Power     : %u mW (both antennas)\n", power);
    printf("Scan rate : maximum (no sleep between sweeps)\n");
    printf("Report    : 1 summary line every %d ms\n", GC_RATE_WINDOW_MS);
    printf("Antennas  : %s, %s\n\n", sources[0], sources[1]);

    printf("[GC] Connecting...\n");
    ec = CAENRFID_Connect(&reader, CAENRFID_RS232, &port_params);
    if (ec != CAENRFID_StatusOK) {
        printf("[GC] ERROR: Could not connect (code %d)\n", ec);
        printf("  - Check USB cable\n");
        printf("  - Try: sudo chmod 666 %s\n", GC_PORT);
        printf("  - Or:  sudo usermod -a -G dialout $USER  (then re-login)\n");
        return -1;
    }

    ec = CAENRFID_GetReaderInfo(&reader, model, serial);
    if (ec == CAENRFID_StatusOK)
        printf("[GC] Reader: %s  Serial: %s\n", model, serial);

    char fwrel[MAX_FWREL_LENGTH + 1] = {0};
    if (CAENRFID_GetFirmwareRelease(&reader, fwrel) == CAENRFID_StatusOK)
        printf("[GC] Firmware: %s\n", fwrel);

    ec = CAENRFID_SetPower(&reader, power);
    if (ec != CAENRFID_StatusOK) {
        printf("[GC] WARNING: SetPower(%u) returned %d -- "
               "value may be below the reader's hardware floor.\n",
               power, ec);
    }
    printf("[GC] Ready. One line per second: [S0=…/s S1=…/s] is the measured\n"
           "[GC] scan rate per antenna; tagged windows also prepend [TX …]. Ctrl+C to stop.\n\n");

    running = 1;

    /* disp_* holds the most-recent non-empty observation per antenna
       during the current 1-second window. It is what eventually gets
       printed; we overwrite a slot only when that antenna actually saw
       tags in a sweep, so a brief miss in the last sweep of the window
       doesn't erase a tag that was visible 50 ms earlier. */
    TagEntry disp_bucket[ANTENNA_COUNT][GC_MAX_TAGS];
    int      disp_cnt[ANTENNA_COUNT] = {0, 0};

    /* scans[ant] counts CAENRFID_InventoryTag calls per antenna inside
       the current window; we divide by the actual elapsed window time
       to get the displayed scans/second. */
    unsigned long scans[ANTENNA_COUNT] = {0, 0};

    struct timespec t_window;
    clock_gettime(CLOCK_MONOTONIC, &t_window);

    while (running) {

        TagEntry sweep_bucket[ANTENNA_COUNT][GC_MAX_TAGS];
        int      sweep_cnt[ANTENNA_COUNT] = {0, 0};

        for (int ant = 0; ant < ANTENNA_COUNT && running; ant++) {
            CAENRFIDTagList *tag_list = NULL, *node;
            uint16_t num_tags = 0;

            ec = CAENRFID_InventoryTag(&reader, (char *)sources[ant],
                                       0, 0, 0,
                                       NULL, 0,
                                       RSSI | PHASE,
                                       &tag_list, &num_tags);

            /* Count every attempted inventory call; this is the figure
               the user wants to maximise. */
            scans[ant]++;

            if (ec == CAENRFID_StatusOK && num_tags > 0) {
                node = tag_list;
                while (node != NULL) {
                    if (sweep_cnt[ant] < GC_MAX_TAGS &&
                        sweep_cnt[0] + sweep_cnt[1] < GC_MAX_TAGS) {
                        hex_str(node->Tag.ID, node->Tag.Length,
                                sweep_bucket[ant][sweep_cnt[ant]].tag);
                        sweep_bucket[ant][sweep_cnt[ant]].antenna = ant;
                        sweep_bucket[ant][sweep_cnt[ant]].rssi    = node->Tag.RSSI;
                        sweep_bucket[ant][sweep_cnt[ant]].phase   = node->Tag.Phase;
                        sweep_cnt[ant]++;
                    }
                    CAENRFIDTagList *next = node->Next;
                    free(node);
                    node = next;
                }
            } else {
                // Free list if returned with non-OK code
                node = tag_list;
                while (node != NULL) {
                    CAENRFIDTagList *next = node->Next;
                    free(node);
                    node = next;
                }
            }
        }

        /* Latch the most-recent non-empty result per antenna into the
           display buffer. We only overwrite a slot when this sweep saw
           tags on that antenna, so the displayed snapshot reflects the
           freshest read within the window. */
        for (int ant = 0; ant < ANTENNA_COUNT; ant++) {
            if (sweep_cnt[ant] > 0) {
                for (int i = 0; i < sweep_cnt[ant]; i++)
                    disp_bucket[ant][i] = sweep_bucket[ant][i];
                disp_cnt[ant] = sweep_cnt[ant];
            }
        }

        /* Has the rate-window elapsed? If so, compute scans/second
           per antenna, print the summary, and reset for the next
           window. No fixed sleep -- the inventory call itself is the
           only thing pacing the loop, which is exactly what gives us
           the maximum physical scan rate. */
        struct timespec now;
        clock_gettime(CLOCK_MONOTONIC, &now);
        double elapsed_s = (now.tv_sec  - t_window.tv_sec) +
                           (now.tv_nsec - t_window.tv_nsec) / 1e9;

        if (elapsed_s * 1000.0 >= (double)GC_RATE_WINDOW_MS) {
            int rate[ANTENNA_COUNT];
            for (int ant = 0; ant < ANTENNA_COUNT; ant++)
                rate[ant] = (elapsed_s > 0.0)
                    ? (int)((double)scans[ant] / elapsed_s + 0.5)
                    : 0;

            print_sweep_line(power, disp_bucket, disp_cnt, rate);

            scans[0] = 0;
            scans[1] = 0;
            disp_cnt[0] = 0;
            disp_cnt[1] = 0;
            t_window = now;
        }
    }

    CAENRFID_Disconnect(&reader);
    printf("[GC] Disconnected.\n");
    return 0;
}
