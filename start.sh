#!/bin/bash
set -euo pipefail

project=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
exec bash "$project/entrypoint.sh"
