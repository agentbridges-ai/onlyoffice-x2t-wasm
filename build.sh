#!/usr/bin/env bash

set -euxo pipefail

rm -rf results build
ci_output="$(mktemp -d)"
trap 'rm -rf "${ci_output}"' EXIT

docker build --target ci-output -o "${ci_output}" .

mv "${ci_output}/results" results
mv "${ci_output}/build" build

test -s results/test.js.log
test -s build/x2t.js
test -s build/x2t.wasm
