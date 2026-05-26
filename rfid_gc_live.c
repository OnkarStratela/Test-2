// rfid_gc_live.c
// Dual-antenna RFID live scanner with strict antenna attribution.
//
// Physical setup (the one this arbitrator was tuned for):
//   - Two CAEN R3100C Lepton3 antennas mounted horizontally, 150 mm
//     centre-to-centre. A tag is "placed over" one antenna when it sits
//     roughly above its centre (~5-7 cm from that antenna's face). The
//     opposite antenna is then ~15 cm away, so its received signal for
//     the same tag should be ~9-15 dB weaker than the home antenna
//     (free-space path-loss only; reality usually adds more attenuation
//     from antenna pattern roll-off and the dielectric of the housing).
//
// Goal: over every 2-second window decide *which* antenna actually
// owns each EPC, with ZERO cross-reads. A tag sitting over antenna A
// must NEVER show up in antenna B's slot, even when antenna B picks
// up a faint leakage read of the same tag.
//
// How: we drive both antennas serially at maximum physical speed (no
// sleep between sweeps, exactly as before). For every read returned by
// CAENRFID_InventoryTag we update a per-EPC stats table that tracks,
// per antenna, the read count and the maximum RSSI seen during the
// window. When the 2-second window closes we run an arbitration pass
// that promotes a tag onto an antenna's slot ONLY if that antenna is
// dominant in both signal strength AND read count.
//
// Arbitration rule (applied once per EPC at end of each window):
//   - The "winner" is the antenna with the larger max RSSI.
//   - If the winner's max RSSI is below GC_RSSI_FLOOR_DBM10, the tag
//     is treated as leakage and dropped entirely.
//   - If the winner's read count is below GC_MIN_READS, dropped.
//   - If only the winner saw the tag during the window: attribute it.
//   - If both antennas saw the tag, additionally require BOTH:
//       max_rssi[winner] - max_rssi[loser] >= GC_RSSI_MARGIN_DB10
//       count[winner] >= GC_COUNT_DOMINANCE * count[loser]
//     If either condition fails -> ambiguous, drop the tag entirely.
//
// The reader is asked for RSSI only (no PHASE). The displayed RSSI is
// the maximum value seen by the winning antenna across the window
// (closest to 0 dBm = strongest read = most confident attribution).
// The reader reports RSSI in tenths of dBm; we divide by 10.0 for
// display.
//
// Output (identical to previous version minus the phase column):
//
// Empty (no tag was decisively attributed during the 2 s window):
//   [S0=137/s S1=137/s] []
//
// With attribution -- slot 0 is always antenna 0, slot 1 is always
// antenna 1, and empty slots render as pure whitespace so the comma
// and the other slot never shift columns:
//   [TX=30 mW] [S0=137/s S1=137/s] [(0)(-58.3) E2801160600002054E1A1234,   (1)(-61.7) E2801160600002054E1A5678]
//   [TX=30 mW] [S0=137/s S1=137/s] [(0)(-58.3) E2801160600002054E1A1234,                                      ]
//   [TX=30 mW] [S0=137/s S1=137/s] [                                    ,   (1)(-61.7) E2801160600002054E1A5678]
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
#define GC_RATE_WINDOW_MS   1000          // decision/arbitration window (also the print cadence)
#define GC_MAX_TAGS         64            // max distinct EPCs tracked per window
#define ANTENNA_COUNT       2
#define MAX_ID_LENGTH       64

// Arbitration thresholds (tune for your physical layout).
//
// RSSI values from the reader are in tenths of dBm: -631 -> -63.1 dBm.
// All four thresholds are deliberately conservative -- when in doubt we
// drop the tag (= zero cross-reads), which is the explicit design goal.
#define GC_RSSI_FLOOR_DBM10  (-700)  // -70.0 dBm: reads weaker than this are treated as leakage
#define GC_RSSI_MARGIN_DB10  ( 60)   //   6.0 dB: winner must beat loser by this much max RSSI
#define GC_COUNT_DOMINANCE   ( 2)    //   winner's read count must be >= this * loser's read count
#define GC_MIN_READS         ( 3)    //   ignore EPCs with fewer reads than this on the winning antenna

// ANSI colours
#define GREEN  "\033[0;32m"
#define RED    "\033[0;31m"
#define YELLOW "\033[0;33m"
#define CYAN   "\033[0;36m"
#define RESET  "\033[0m"

volatile int running = 0;

// One entry to display in a slot after arbitration.
typedef struct {
    char    tag[2 * MAX_ID_LENGTH + 1];
    int     antenna;
    int16_t rssi;        /* tenths of dBm, as reported by the reader */
} TagEntry;

// Per-EPC statistics accumulated across one decision window.
// max_rssi[ant] is the strongest RSSI seen by `ant` during the window
// (larger = closer to 0 dBm = stronger); INT16_MIN means "not seen".
typedef struct {
    char     epc[2 * MAX_ID_LENGTH + 1];
    unsigned count[ANTENNA_COUNT];
    int16_t  max_rssi[ANTENNA_COUNT];
} TagStats;

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
// "(N)(-XXX.X) " + a 24-char EPC-96 hex string.
#define SLOT_WIDTH 38

// Returns the index of the entry for `epc` in `stats`, inserting a
// new entry (with max_rssi initialised to INT16_MIN) if needed.
// Returns -1 if the table is already full.
static int stats_find_or_insert(TagStats *stats, int *count, const char *epc)
{
    for (int i = 0; i < *count; i++) {
        if (strcmp(stats[i].epc, epc) == 0) return i;
    }
    if (*count >= GC_MAX_TAGS) return -1;
    int idx = (*count)++;
    memset(&stats[idx], 0, sizeof stats[idx]);
    strncpy(stats[idx].epc, epc, sizeof stats[idx].epc - 1);
    stats[idx].epc[sizeof stats[idx].epc - 1] = '\0';
    stats[idx].max_rssi[0] = INT16_MIN;
    stats[idx].max_rssi[1] = INT16_MIN;
    return idx;
}

// Apply the arbitration rule and return the winning antenna index
// (0 or 1), or -1 if the tag is ambiguous / leakage and should NOT
// be displayed. See the file header for the rule definition.
static int stats_arbitrate(const TagStats *s)
{
    int winner, loser;
    if (s->max_rssi[0] >= s->max_rssi[1]) { winner = 0; loser = 1; }
    else                                  { winner = 1; loser = 0; }

    /* Need a real, strong observation on the winning antenna. */
    if (s->count[winner] < (unsigned)GC_MIN_READS)            return -1;
    if (s->max_rssi[winner] < GC_RSSI_FLOOR_DBM10)            return -1;

    /* Only the winner saw it -> clean attribution. */
    if (s->count[loser] == 0) return winner;

    /* Both antennas saw it -> require dominance in BOTH axes. If
       either margin is missed we drop the tag entirely, which is the
       "zero cross-read" guarantee. */
    int16_t margin = s->max_rssi[winner] - s->max_rssi[loser];
    if (margin < GC_RSSI_MARGIN_DB10)                         return -1;
    if (s->count[winner] < (unsigned)GC_COUNT_DOMINANCE * s->count[loser])
        return -1;

    return winner;
}

// Summary output for one decision window:
//   - "[S0=N/s S1=N/s] []" when arbitration attributed nothing
//   - otherwise: [TX=…] [S0=N/s S1=N/s] [<slot0>,   <slot1>]
//     Each slot is padded to SLOT_WIDTH. If an antenna won no tags in
//     the window, its slot is pure whitespace (no "(N)"), so the comma
//     and the other slot keep their column positions and nothing shifts.
//     The scans/second figure is measured over the window: it is the
//     count of CAENRFID_InventoryTag calls per antenna divided by the
//     elapsed window time.
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

        /* Visible width of the slot content (excludes ANSI codes) so
           we can right-pad to SLOT_WIDTH and keep the ']' column
           stable. The displayed RSSI is the max RSSI seen by the
           winning antenna during the window, in real dBm with one
           decimal place (e.g. raw -583 -> "-58.3"). */
        int visible = 3; /* "(N)" */
        for (int i = 0; i < cnt[ant]; i++) {
            char rbuf[24];
            int rlen = snprintf(rbuf, sizeof rbuf,
                                "(%.1f) ",
                                bucket[ant][i].rssi / 10.0);
            if (i > 0) visible += 1; /* space between multiple tags */
            visible += rlen + (int)strlen(bucket[ant][i].tag);
        }

        const char *tagcol = (ant == 0) ? GREEN : RED;
        printf(YELLOW "(%d)" RESET, ant);
        for (int i = 0; i < cnt[ant]; i++) {
            if (i > 0) printf(" ");
            printf("(%.1f) %s%s" RESET,
                   bucket[ant][i].rssi / 10.0,
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

    printf(CYAN "===== Dual-Antenna RFID Live Scanner (arbitrated) =====" RESET "\n");
    printf("Port      : %s @ %d baud\n", GC_PORT, GC_BAUDRATE);
    printf("Power     : %u mW (both antennas)\n", power);
    printf("Scan rate : maximum (no sleep between sweeps)\n");
    printf("Decision  : 1 attribution + 1 summary line every %d ms\n", GC_RATE_WINDOW_MS);
    printf("Antennas  : %s, %s  (150 mm centre-to-centre)\n", sources[0], sources[1]);
    printf("Arbitration: floor=%.1f dBm, margin=%.1f dB, count>=%dx, min reads=%d\n\n",
           GC_RSSI_FLOOR_DBM10 / 10.0,
           GC_RSSI_MARGIN_DB10 / 10.0,
           GC_COUNT_DOMINANCE,
           GC_MIN_READS);

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
    printf("[GC] Ready. One line every %.1f s: arbitrated [(0) tag, (1) tag];\n"
           "[GC] [S0=…/s S1=…/s] is the measured scan rate per antenna. Ctrl+C to stop.\n\n",
           GC_RATE_WINDOW_MS / 1000.0);

    running = 1;

    /* Per-EPC statistics, accumulated across the entire decision window.
       Reset at the end of every window. */
    TagStats stats[GC_MAX_TAGS];
    int      stats_count = 0;

    /* scans[ant] counts CAENRFID_InventoryTag calls per antenna inside
       the current window; we divide by the actual elapsed window time
       to get the displayed scans/second. */
    unsigned long scans[ANTENNA_COUNT] = {0, 0};

    struct timespec t_window;
    clock_gettime(CLOCK_MONOTONIC, &t_window);

    while (running) {

        for (int ant = 0; ant < ANTENNA_COUNT && running; ant++) {
            CAENRFIDTagList *tag_list = NULL, *node;
            uint16_t num_tags = 0;

            ec = CAENRFID_InventoryTag(&reader, (char *)sources[ant],
                                       0, 0, 0,
                                       NULL, 0,
                                       RSSI,
                                       &tag_list, &num_tags);

            /* Count every attempted inventory call; this is the figure
               we report as scans/s/antenna. */
            scans[ant]++;

            /* Walk the linked list returned by the reader. We always
               free every node; we only fold reads into `stats` when
               the inventory call actually succeeded with tags. */
            node = tag_list;
            while (node != NULL) {
                if (ec == CAENRFID_StatusOK && num_tags > 0) {
                    char epc[2 * MAX_ID_LENGTH + 1];
                    hex_str(node->Tag.ID, node->Tag.Length, epc);
                    int idx = stats_find_or_insert(stats, &stats_count, epc);
                    if (idx >= 0) {
                        stats[idx].count[ant]++;
                        if (node->Tag.RSSI > stats[idx].max_rssi[ant])
                            stats[idx].max_rssi[ant] = node->Tag.RSSI;
                    }
                }
                CAENRFIDTagList *next = node->Next;
                free(node);
                node = next;
            }
        }

        /* Has the decision window elapsed? If so, run arbitration on
           every EPC seen during the window, build the display buffer
           with only the unambiguous attributions, print one summary
           line, and reset for the next window. No fixed sleep -- the
           inventory call itself is the only thing pacing the loop. */
        struct timespec now;
        clock_gettime(CLOCK_MONOTONIC, &now);
        double elapsed_s = (now.tv_sec  - t_window.tv_sec) +
                           (now.tv_nsec - t_window.tv_nsec) / 1e9;

        if (elapsed_s * 1000.0 >= (double)GC_RATE_WINDOW_MS) {
            TagEntry disp_bucket[ANTENNA_COUNT][GC_MAX_TAGS];
            int      disp_cnt[ANTENNA_COUNT] = {0, 0};

            for (int i = 0; i < stats_count; i++) {
                int winner = stats_arbitrate(&stats[i]);
                if (winner < 0) continue;
                if (disp_cnt[winner] >= GC_MAX_TAGS) continue;

                TagEntry *e = &disp_bucket[winner][disp_cnt[winner]++];
                strncpy(e->tag, stats[i].epc, sizeof e->tag - 1);
                e->tag[sizeof e->tag - 1] = '\0';
                e->antenna = winner;
                e->rssi    = stats[i].max_rssi[winner];
            }

            int rate[ANTENNA_COUNT];
            for (int ant = 0; ant < ANTENNA_COUNT; ant++)
                rate[ant] = (elapsed_s > 0.0)
                    ? (int)((double)scans[ant] / elapsed_s + 0.5)
                    : 0;

            print_sweep_line(power, disp_bucket, disp_cnt, rate);

            scans[0] = 0;
            scans[1] = 0;
            stats_count = 0;
            t_window = now;
        }
    }

    CAENRFID_Disconnect(&reader);
    printf("[GC] Disconnected.\n");
    return 0;
}
