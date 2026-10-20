# Word Studio build of RHWP

MIT-licensed RHWP 0.8.7, built from
[`1a76570e833917d15817415a53c09ad61ab3203f`](https://github.com/edwardkim/rhwp/tree/1a76570e833917d15817415a53c09ad61ab3203f)
with `pagination.patch` and `manifest-paths.patch` (`0.8.7-wordstudio.2`). This is a
local Word Studio build, not an upstream release.
The package is vendored so installers use the exact same engine offline, without
requiring Rust on users' machines or downloading a replacement during CI.

The patch makes two layout corrections:

1. Short/landscape paper uses the same finite row budget as other paper. The old
   scanner allowed another 36–320 px, placing text outside the physical sheet.
2. For unedited stored cell layout, initial paragraph before-space is removed
   before testing whether the saved line positions have a local origin. This
   preserves the existing source frame-reset handling in long tables.

No file names, document hashes, target page counts, or compensating scale factors
are used by the engine. `tests/hwp-pagination.test.cjs` covers synthetic HWP and
HWPX tables that the original engine rendered in three pages with clipped text;
this build uses four with every numbered cell present once. The public 2025
work-rules document changes from 154 to 163 pages, matching Polaris's page count.
Matching page counts do not imply identical page breaks or typography.

`manifest-paths.patch` resolves embedded manifest paths against
`Contents/content.hpf` before reading sections, master pages or lazy binary
resources. If the relative entry is absent, it retains the archive-root notation
used by existing Hancom files. If both exist, the OPF-relative entry wins. Dot
segments are normalized inside the ZIP; absolute paths, remote URLs and paths
escaping the archive are rejected. External linked resources are not rewritten.
This does not relocate content.hpf/header.xml or change pagination rules.
`tests/hwpx-manifest.test.cjs` checks the two-paragraph relative-path fixture,
unchanged rendering against its root-path control, a conflicting root copy,
subdirectories/spine order, and missing/invalid paths.

Build toolchain: Rust 1.93.1 (`wasm32-unknown-unknown`), wasm-pack 0.15.0,
wasm-bindgen CLI 0.2.127, upstream locked Cargo dependencies and wasm-opt settings.
`node scripts/rebuild-hwp-core.cjs` checks out the pinned source into `.cache/`,
applies this patch and invokes upstream's locked build wrapper. Tools must already
be installed. It does not install tools or alter shell configuration. The official
upstream Docker release build was not run; this artifact was built on macOS arm64.

`SHA256SUMS` records the shipped JavaScript, declarations and optimized WASM.
The original copyright and MIT permission notice are in `LICENSE`.
