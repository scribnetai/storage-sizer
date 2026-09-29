# Storage Sizer — storage capacity planner for VMware refreshes

Size the storage before you quote it. Feed in an RVTools export (or type
pool numbers in manually), tune reduction, protection, replication,
snapshots, overhead, headroom, and growth per pool — and get a
year-by-year raw-capacity plan, per-pool worked math, SE talking points,
and a customer-ready HTML report.

**Live:** https://storage-sizer.scribnet.io/

## How it works

1. **Load demand** — drop an RVTools `.xlsx` (File → Export all to Excel),
   enter pool totals manually, or try the synthetic demo environment.
   RVTools parsing rolls up per-cluster storage demand from `vInfo` and
   reads `vDatastore` totals for capacity validation, `vSnapshot` for
   snapshot context, and `vDisk` for thin-vs-thick counts.
2. **Configure storage** — the tunable mid-step: growth horizon (1–5 yrs),
   compounding %/yr or static TB/yr growth (global, with per-pool rate
   overrides), and per-pool demand basis (used/provisioned), data-reduction
   ratio, local protection (mirror / RAID-5 / RAID-6 / EC 8+3), async
   replication copies, snapshot retention, system overhead, and free-space
   headroom. Raw-capacity previews update live as you tune.
3. **Growth plan** — year-by-year raw projections, per-pool worked math
   (today and horizon year), validation against current usable capacity,
   auto-written findings, and a downloadable standalone HTML briefing
   report.

## Sizing math

Per pool, per year `y` (0 → horizon):

```
logical[y] = basis × (1+g)^y        (compounding)
           = basis + y×gTB          (static)
raw[y] = (logical[y] ÷ reduction)
       × (1 + snapCopies × snapSize%)
       × localProtectionFactor
       × (1 + remoteCopies)
       × (1 + overhead%)
       ÷ (1 − headroom%)
```

Storage counts regardless of VM power state (only templates are
excluded); that's why "used" is the default demand basis and
"provisioned" is the conservative flip.

## Privacy

100% client-side. The spreadsheet parser (vendored SheetJS) runs in the
page, nothing is uploaded, nothing is stored (no localStorage/IndexedDB/
cookies), and the page works fully offline after load.

## Honest limitations

- Capacity sizing, not performance sizing — no IOPS/latency modeling.
- Data-reduction ratios are planning assumptions, not measurements —
  validate with vendor sizing tools.
- Snapshot retention cost is an estimate of change rate.
- Replication bandwidth and RPO aren't modeled — only the replica
  capacity footprint.
- No pricing — output is raw TB per pool per year for your VAR quoting
  tools.

## Files

- `index.html` — landing page + 3-step wizard
- `css/styles.css` — theme (cache-busted via `?v=N`)
- `js/app.js` — parsing, sizing math, wizard, report generator
- `lib/xlsx.full.min.js` — vendored SheetJS (offline parsing)

Not affiliated with Broadcom/VMware.
