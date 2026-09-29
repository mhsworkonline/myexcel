#!/bin/sh
# Dev helper: rebuild and restart the production server on :3100
cd "$(dirname "$0")/.."
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 3100 -State Listen -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id \$_.OwningProcess -Force }" >/dev/null 2>&1
if [ "$1" != "--no-build" ]; then npx next build 2>&1 | grep -E "Compiled|Failed|rror" | head -20; fi
(npx next start -p 3100 > /dev/null 2>&1 &)
for i in 1 2 3 4 5 6 7 8 9 10; do sleep 1; curl -s -o /dev/null http://localhost:3100/ && break; done
echo server up
