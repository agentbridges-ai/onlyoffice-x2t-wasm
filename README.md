# Build OnlyOffice x2t for WebAssembly

Current maintained release: `v9.3.0+3` (ONLYOFFICE core
`v9.3.0.140`, Emscripten `4.0.11`).

Release tags are built in GitHub Actions from the Dockerfile, tested against
the repository fixtures, packaged reproducibly, and published with SHA-256,
SHA-512, and GitHub artifact attestations. Verify a downloaded artifact with:

```shell
sha256sum --check onlyoffice-x2t-wasm-v9.3.0+3.tar.gz.sha256
gh attestation verify onlyoffice-x2t-wasm-v9.3.0+3.tar.gz \
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
