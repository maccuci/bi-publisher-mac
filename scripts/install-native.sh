#!/bin/bash
# Instala o plugin no Word. Funciona no repositorio (usa out/native) e no zip da release
# (como "Install.command", ao lado dos arquivos). Usa so ferramentas do macOS.
set -euo pipefail
script_dir="$(cd "$(dirname "$0")" && pwd)"
if [ -f "$script_dir/BIPTemplateEditor.dotm" ]; then
  native_output="$script_dir"
else
  native_output="$(cd "$script_dir/.." && pwd)/out/native"
fi
if [ "$(uname -s)" != Darwin ]; then
  printf '%s\n' 'This installer requires Word for macOS.' >&2
  exit 1
fi
for native_file in BIPTemplateEditor.dotm BIPBridge.applescript bip-xml BIPPreview.java; do
  if [ ! -f "$native_output/$native_file" ]; then
    printf '%s\n' "Missing $native_file in $native_output. In the repository, follow native/README.md to complete the native build." >&2
    exit 1
  fi
done
# Listagem numa variavel: com pipefail, `unzip | grep -q` falha quando o grep fecha o pipe antes.
package_listing="$(unzip -l "$native_output/BIPTemplateEditor.dotm" 2>/dev/null || true)"
if [[ "$package_listing" != *word/vbaProject.bin* || "$package_listing" != *customUI/customUI.xml* ]]; then
  printf '%s\n' 'The file is not the packaged native plugin. Run native:build before installing.' >&2
  exit 1
fi
if pgrep -x 'Microsoft Word' >/dev/null; then
  printf '%s\n' 'Save your documents and quit Word with Cmd+Q before installing.' >&2
  exit 1
fi
# Arquivos baixados da internet ficam em quarentena e o macOS bloquearia o leitor XML.
# Quem executa este instalador ja decidiu confiar nestes arquivos.
xattr -dr com.apple.quarantine "$native_output" 2>/dev/null || true
native_startup="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
native_scripts="$HOME/Library/Application Scripts/com.microsoft.Word"
mkdir -p "$native_startup" "$native_scripts"
native_stamp="$(date +%Y%m%d%H%M%S)"
install_copy() {
  local native_source="$1" native_destination="$2"
  if [ -L "$native_destination" ]; then
    printf '%s\n' "Destination is a symbolic link; installation refused: $native_destination" >&2
    exit 1
  fi
  if [ -f "$native_destination" ]; then
    cp -p "$native_destination" "$native_destination.backup-$native_stamp"
  fi
  cp "$native_source" "$native_destination"
}
install_copy "$native_output/BIPTemplateEditor.dotm" "$native_startup/BIPTemplateEditor.dotm"
install_copy "$native_output/BIPBridge.applescript" "$native_scripts/BIPBridge.applescript"
install_copy "$native_output/bip-xml" "$native_scripts/bip-xml"
install_copy "$native_output/BIPPreview.java" "$native_scripts/BIPPreview.java"
chmod 755 "$native_scripts/bip-xml"
printf '%s\n' 'Installed. Open Word and look for the BI Publisher tab. Enable macros only for this template when Word asks; security settings were not changed.'
printf '%s\n' 'Preview PDF (optional): install Java (brew install openjdk@21) and import the engine from your licensed Oracle Home.'
