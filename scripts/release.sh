#!/bin/bash
# Monta o zip da release para usuarios finais (anexar a uma GitHub Release).
# Pre-requisito manual: o VBA de native/BIPTemplateEditor.bas importado e salvo pelo Word na base
# (native/README.md, "Manual import or update"); o build recusa base desatualizada.
# Nunca inclui componentes Oracle.
set -euo pipefail
task_root="$(cd "$(dirname "$0")/.." && pwd)"
cd "$task_root"
base="${1:-out/native/BIPTemplateEditor.base.dotm}"
bash scripts/build-native.sh "$base"
python3 scripts/package-native.test.py >/dev/null 2>&1 || { printf '%s\n' 'Packaging tests failed; release cancelled.' >&2; exit 1; }

version="$(sed -n 's/^  "version": "\(.*\)",$/\1/p' package.json)"
name="BIPTemplateEditor-$version"
stage="out/release/$name"
rm -rf "$stage" "out/release/$name.zip"
mkdir -p "$stage"
cp out/native/BIPTemplateEditor.dotm out/native/BIPBridge.applescript out/native/bip-xml out/native/BIPPreview.java "$stage/"
cp scripts/install-native.sh "$stage/Install.command"
cp scripts/uninstall-native.sh "$stage/Uninstall.command"
cp scripts/import-oracle-engine.sh "$stage/Import Oracle Engine.command"
cp LICENSE NOTICE "$stage/"
cp scripts/release-readme.txt "$stage/README.txt"
chmod 755 "$stage"/*.command "$stage/bip-xml"

# Conferencias do pacote final.
[ "$(lipo -archs "$stage/bip-xml")" = "x86_64 arm64" ] || { printf '%s\n' 'bip-xml is not universal.' >&2; exit 1; }
"$stage/bip-xml" --self-test >/dev/null
package_listing="$(unzip -l "$stage/BIPTemplateEditor.dotm")"
[[ "$package_listing" == *word/vbaProject.bin* ]] || { printf '%s\n' 'The .dotm does not contain VBA.' >&2; exit 1; }
if [ -n "$(find "$stage" -iname '*.jar')" ]; then printf '%s\n' 'Jar found in the package; release cancelled.' >&2; exit 1; fi

ditto -c -k --norsrc --noextattr --noqtn --keepParent "$stage" "out/release/$name.zip"
printf '%s\n' "Release ready: out/release/$name.zip"
