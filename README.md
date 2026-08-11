# Build OnlyOffice x2t for WebAssembly

Current maintained release: `v9.3.0+4` (ONLYOFFICE core
`v9.3.0.140`, Emscripten `4.0.11`).

The release converter is linked with `ALLOW_MEMORY_GROWTH` and
`EMULATE_FUNCTION_POINTER_CASTS=1`. Its immutable build gate converts the real
legacy OLE DOC and native PivotTable/Slicer XLSX fixtures directly to Canvas
formats (`66 -> 8193` and `257 -> 8194`) with `m_bIsNoBase64=false`. The gate
pins the fixture and resulting `Editor.bin` SHA-256 digests, including embedded
OLE media and workbook media.

The maintained source build's legacy OLE DOC baseline is `DOCY;v5;`, 132,614
bytes, SHA-256
`3487cf192cbdea43816c31712dc7e9a4846bb8b1958a1b0d273d8ecd42fde93d`.
This value is deliberately tied to the clean build produced by this repository,
not to a previously distributed browser-carried converter binary. The gate also
requires `display6image1.bin`, `display6image1.emf`, and
`display6image1.svg`, so changing the digest cannot mask lost embedded media.
The sanitized native PivotTable/Slicer baseline is `XLSY;v2;`, 85,138 bytes,
SHA-256
`dc0acc3071dfdd177e2c45eec6437b8e11b622f5457c22274d591252ac8538e1`,
with `media/image1.png`. Fixture privacy metadata is rejected before either
Canvas golden is evaluated.

The WebAssembly `main1` wrapper resets the document-local identifier generator
to a fixed seed at the start of every conversion. Generated GUIDs remain unique
within each document, while wall-clock time and previous conversions can no
longer alter `Editor.bin`. Each real fixture is converted twice in the same
module and must be byte-for-byte identical before its exact golden is checked.

Release tags are built in GitHub Actions from the Dockerfile, tested against
the repository fixtures, packaged reproducibly, and published with SHA-256,
SHA-512, and GitHub artifact attestations. Verify a downloaded artifact with:

```shell
sha256sum --check onlyoffice-x2t-wasm-v9.3.0+4.tar.gz.sha256
gh attestation verify onlyoffice-x2t-wasm-v9.3.0+4.tar.gz \
  --repo agentbridges-ai/onlyoffice-x2t-wasm
```

## Modifications by CryptPad

This repository contains a modified copy of https://github.com/ONLYOFFICE/core.git in `/core`. These modifications are made to be able to compile `x2t` to WebAssembly.

## Build

This is a Dockerfile building OnlyOffice x2t in WebAssembly using emscripten.
Build it with:

``` shell
./build.sh
```

The local script and CI both use the `ci-output` target. BuildKit compiles the
converter once, runs the regression suite against those exact bytes, then
exports the tested payload and evidence together. GitHub Actions warms a
default-branch cache on every `main` update; a signed release tag on the same
commit is serialized behind that run and restores the same cache. Pull requests
write to isolated cache scopes while being allowed to read the `main` cache.
Base images and source tags are pinned so a floating dependency cannot silently
invalidate every compilation layer between otherwise identical builds.
The manual `no_cache` input provides a cold-build canary whose exported hashes
can be compared with a normal cached run before changing a release baseline.

## Update to a new x2t version

This repository includes a clone of x2t in the `core` directory. You can pull a
new x2t release with:

``` shell
git subtree pull --prefix core https://github.com/ONLYOFFICE/core.git <TAG> --squash
```

Since the clone contains small changes there may be merge conflicts.

## See changes we made to https://github.com/ONLYOFFICE/core.git

``` shell
git fetch --depth=1 https://github.com/ONLYOFFICE/core.git v9.3.0.140
git diff FETCH_HEAD HEAD:core
```
