#!/bin/bash
APP_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$APP_DIR"
exec python3 lux_command_center.py
