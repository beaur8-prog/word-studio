#!/usr/bin/env bash
set -euo pipefail

if [[ "${CI:-}" != true || "${RUNNER_OS:-}" != Linux ]]; then
  echo 'Linux dependency setup is restricted to isolated GitHub Actions runners.' >&2
  exit 1
fi
if [[ $# -gt 1 || ( $# -eq 1 && "$1" != browser ) ]]; then
  echo 'Usage: prepare-linux-ci.sh [browser]' >&2
  exit 1
fi

# The hosted runner's Azure mirror can take minutes per font package. Keep
# Ubuntu's authenticated repositories, using their official mirrors directly.
if [[ -f /etc/apt/apt-mirrors.txt ]]; then
  printf '%s\n' 'https://archive.ubuntu.com/ubuntu/' | sudo tee /etc/apt/apt-mirrors.txt >/dev/null
fi
if [[ -f /etc/apt/apt-security-mirrors.txt ]]; then
  printf '%s\n' 'https://security.ubuntu.com/ubuntu/' | sudo tee /etc/apt/apt-security-mirrors.txt >/dev/null
fi
printf '%s\n' 'Acquire::ForceIPv4 "true";' 'Acquire::Retries "2";' \
  'Acquire::http::Timeout "30";' 'Acquire::https::Timeout "30";' \
  | sudo tee /etc/apt/apt.conf.d/80-word-studio-ci >/dev/null
sudo apt-get update
sudo apt-get install -y xvfb fonts-noto-cjk desktop-file-utils
if [[ "${1:-}" == browser ]]; then
  npx playwright install --with-deps chromium
fi
