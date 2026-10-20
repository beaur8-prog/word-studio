# HWP reader fixtures

Large conversion fixtures are generated at test time by `tests/large-hwp-fixture.cjs`.
Three incompressible PNGs create real ~21 MB HWP/HWPX files whose editable HTML
exceeds the former 24 MiB cap. A separate ~21 MB PNG covers individual-picture
retention and Word Studio backups above 40 MiB. These generated files stay in
ignored test/cache directories and are not committed.

Passing `4` as the fixture generator's second argument creates 12 pictures,
~85 MB source files and >100 MiB of editable HTML for the expanded limits.
`tests/hwp-scroll-fixture.cjs` generates a 1,000-page text document to verify
continuous scrolling, direct seeking and text-only search with bounded SVG paint.

`reader.hwp` is a synthetic text/table document created with @rhwp/core 0.8.7. `reader.hwpx` uses the format/style shell of the upstream MIT-licensed `samples/표-텍스트.hwpx`, with all document text replaced by synthetic test strings and scripts/previews removed.

Source and exact upstream revision: see the matching entry in `tests/hwp-samples.json`. Other public samples are downloaded into the ignored `analysis/hwp-validation/` directory, not distributed with the app.

RHWP copyright (c) 2025-2026 Edward Kim. MIT license: https://github.com/edwardkim/rhwp/blob/main/LICENSE

`manifest-relative.hwpx` uses the same style shell with two synthetic paragraphs
on A4 paper. Its container points to `Contents/content.hpf`; the manifest uses
`href="section0.xml"` and the only section is `Contents/section0.xml`. There is
no root-level section copy. Scripts, previews and the table are removed, and
version.xml identifies it as a Word Studio synthetic fixture. It tests package
path resolution, independently of page-layout changes. The Node tests generate
root-path controls, a conflicting root copy, subdirectories and missing/unsafe
paths from this fixture.

`landscape-pagination.hwpx` reuses the same synthetic style/section shell, with
landscape A4, a repeated two-column header, and 24 numbered table rows alternating
6,000/6,750 HWP units (80/90 px). Header height is 3,000 units. Page margins are
6,000 units left/right and 3,000 top/bottom/header/footer. Every cell has local
line origin zero and 141-unit padding. Text after the table is retained to test
flow into the final page. `landscape-pagination.hwp` was exported from this HWPX
using the **unpatched** upstream 0.8.7 writer. Both produce three pages and clipped
rows 15/16 with the upstream engine, four pages with all 48 numbered cells exactly
once with the patched engine. These files contain only synthetic test text.
