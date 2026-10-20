# HWP reader fixtures

`reader.hwp` is a synthetic text/table document created with @rhwp/core 0.8.7. `reader.hwpx` uses the format/style shell of the upstream MIT-licensed `samples/표-텍스트.hwpx`, with all document text replaced by synthetic test strings and scripts/previews removed.

Source and exact upstream revision: see the matching entry in `tests/hwp-samples.json`. Other public samples are downloaded into the ignored `analysis/hwp-validation/` directory, not distributed with the app.

RHWP copyright (c) 2025-2026 Edward Kim. MIT license: https://github.com/edwardkim/rhwp/blob/main/LICENSE

`landscape-pagination.hwpx` reuses the same synthetic style/section shell, with
landscape A4, a repeated two-column header, and 24 numbered table rows alternating
6,000/6,750 HWP units (80/90 px). Header height is 3,000 units. Page margins are
6,000 units left/right and 3,000 top/bottom/header/footer. Every cell has local
line origin zero and 141-unit padding. Text after the table is retained to test
flow into the final page. `landscape-pagination.hwp` was exported from this HWPX
using the **unpatched** upstream 0.8.7 writer. Both produce three pages and clipped
rows 15/16 with the upstream engine, four pages with all 48 numbered cells exactly
once with the patched engine. These files contain only synthetic test text.
