# Word Studio pagination build of RHWP

MIT-licensed RHWP 0.8.7, built from
[`1a76570e833917d15817415a53c09ad61ab3203f`](https://github.com/edwardkim/rhwp/tree/1a76570e833917d15817415a53c09ad61ab3203f)
with `pagination.patch`. This is a local Word Studio build, not an upstream release.
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

Build toolchain: Rust 1.93.1 (`wasm32-unknown-unknown`), wasm-pack 0.15.0,
wasm-bindgen CLI 0.2.127, upstream locked Cargo dependencies and wasm-opt settings.
`node scripts/rebuild-hwp-core.cjs` checks out the pinned source into `.cache/`,
applies this patch and invokes upstream's locked build wrapper. Tools must already
be installed. It does not install tools or alter shell configuration. The official
upstream Docker release build was not run; this artifact was built on macOS arm64.

`SHA256SUMS` records the shipped JavaScript, declarations and optimized WASM.
The original copyright and MIT permission notice are in `LICENSE`.
