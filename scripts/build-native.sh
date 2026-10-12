#!/bin/bash
set -euo pipefail
task_root="$(cd "$(dirname "$0")/.." && pwd)"
cd "$task_root"
mkdir -p out/native
# Binario universal (Apple Silicon + Intel) para distribuir em releases.
for arch in arm64 x86_64; do
  swiftc -module-cache-path "$task_root/out/native/swift-cache" -target "$arch-apple-macos11" native/XmlCatalog.swift -o "out/native/bip-xml-$arch"
done
lipo -create out/native/bip-xml-arm64 out/native/bip-xml-x86_64 -output out/native/bip-xml
rm -f out/native/bip-xml-arm64 out/native/bip-xml-x86_64
out/native/bip-xml --self-test
swiftc -module-cache-path "$task_root/out/native/swift-cache" native/RibbonIcons.swift -o out/native/ribbon-icons
out/native/ribbon-icons out/native/icons
cp native/BIPBridge.applescript out/native/BIPBridge.applescript
cp native/BIPPreview.java out/native/BIPPreview.java
if [ "$#" -eq 1 ]; then
  if [ native/BIPTemplateEditor.bas -nt "$1" ]; then
    printf '%s\n' "native/BIPTemplateEditor.bas is newer than $1: the template VBA is out of date. Reimport the module in Word (native/README.md, Manual import or update) and save before building." >&2
    exit 1
  fi
  python3 scripts/package-native.py "$1"
else
  printf '%s\n' 'Native XML reader built. To package the .dotm, pass the template with VBA saved by Word: bun run native:build path/BIPTemplateEditor.base.dotm'
fi
