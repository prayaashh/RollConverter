#!/bin/bash
# Double-click this file in Finder to start Roll Converter.
cd "$(dirname "$0")" || exit 1

if command -v python3 >/dev/null 2>&1; then
    exec python3 rollconverter.py
fi

if command -v python >/dev/null 2>&1; then
    exec python rollconverter.py
fi

osascript -e 'display dialog "Roll Converter needs Python 3.\n\nInstall it from python.org, then open this file again." with title "Roll Converter" buttons {"OK"} default button "OK" with icon caution' >/dev/null 2>&1
exit 1
