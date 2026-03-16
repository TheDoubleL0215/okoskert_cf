#!/bin/bash
cd "$(dirname "$0")"
export PATH="$(pwd)/functions/.venv/bin:$PATH"
firebase emulators:start --only functions "$@"
