#!/usr/bin/env bash
# Renders social-preview.html to social-preview.png (1280x640).
# sample-email.html is a rendered example of the summary email from backup.ps1;
# replace it when the email design changes, then run this script again.
set -euo pipefail
cd "$(dirname "$0")"

npx -y playwright@1.65.1 screenshot \
  --browser chromium \
  --viewport-size "1280,640" \
  --wait-for-timeout 500 \
  "file://$PWD/social-preview.html" \
  social-preview.png
