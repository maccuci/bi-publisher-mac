#!/bin/bash
set -euo pipefail
task_root="$(cd "$(dirname "$0")/.." && pwd)"
cd "$task_root"
if [ "$(uname -s)" != Darwin ]; then
  printf '%s\n' 'This assistant requires Word for macOS.' >&2
  exit 1
fi
native_base="$task_root/out/native/BIPTemplateEditor.base.dotm"
python3 - "$native_base" <<'PY'
import sys
from zipfile import ZipFile
with ZipFile(sys.argv[1]) as package:
    if 'word/vbaProject.bin' in package.namelist():
        raise SystemExit('The base already contains VBA. Use bun run build and bun run native:install; a duplicate will not be imported.')
PY
cp -p "$native_base" "$native_base.backup-$(date +%Y%m%d%H%M%S)-$$"
printf '%s\n' 'Importing VBA through the Word editor. macOS may ask for Automation/Accessibility permission for your Terminal.'
if ! native_result="$(python3 scripts/run-native-import.py "$native_base" "$task_root/native/BIPTemplateEditor.bas")"; then
  printf '%s\n' 'Import interrupted; nothing was installed. Check the message above and the Word window.' >&2
  exit 1
fi
if [ "$native_result" != BIP_VBA_IMPORTED ]; then
  printf '%s\n' 'Word did not confirm the import; nothing was installed.' >&2
  exit 1
fi
bash scripts/build-native.sh "$native_base"
printf '%s\n' 'Build complete. Save your other documents and quit Word with Cmd+Q.'
read -r -p 'Press Return after quitting Word to install: ' native_reply
bash scripts/install-native.sh
printf '%s\n' 'Reopen Word with a copy of the test RTF and check the BI Publisher tab.'
