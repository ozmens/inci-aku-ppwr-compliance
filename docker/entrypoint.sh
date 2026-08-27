#!/bin/sh
set -eu

mkdir -p /data/workspace /data/delivery /data/candidates

seed() {
  src="$1"
  dest="$2"
  if [ -d "$src" ] && [ -z "$(ls -A "$dest" 2>/dev/null || true)" ]; then
    echo "Seeding $dest from $src"
    cp -a "$src"/. "$dest"/
  fi
}

seed /opt/ppwr/delivery /data/delivery
seed /opt/ppwr/workspace /data/workspace
seed /opt/ppwr/candidates /data/candidates

exec uvicorn main:app --host 0.0.0.0 --port "${PORT:-10000}"
