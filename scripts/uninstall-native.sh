#!/bin/bash
# Remove o plugin do Word (template global e arquivos do bridge). Backups *.backup-* sao mantidos.
set -euo pipefail
if pgrep -x 'Microsoft Word' >/dev/null; then
  printf '%s\n' 'Save your documents and quit Word with Cmd+Q before uninstalling.' >&2
  exit 1
fi
native_startup="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
native_scripts="$HOME/Library/Application Scripts/com.microsoft.Word"
for installed in "$native_startup/BIPTemplateEditor.dotm" "$native_scripts/BIPBridge.applescript" \
    "$native_scripts/bip-xml" "$native_scripts/BIPPreview.java"; do
  if [ -f "$installed" ]; then
    rm -f "$installed"
    printf '%s\n' "Removed: $installed"
  fi
done
oracle_copy="$HOME/Library/Application Support/BIPTemplateEditor"
if [ -d "$oracle_copy" ] && [ -t 0 ]; then
  read -r -p "Also remove the Oracle engine copied to $oracle_copy? [y/N] " answer
  case "$answer" in
    y|Y|yes|YES) rm -rf "$oracle_copy"; printf '%s\n' "Removed: $oracle_copy" ;;
  esac
fi
printf '%s\n' 'Plugin uninstalled. Backups of previous versions (*.backup-*) were kept in the same folders.'
