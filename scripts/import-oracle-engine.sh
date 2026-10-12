#!/bin/bash
# Copia uma vez os jars da engine BI Publisher de um Oracle Home licenciado para uma pasta local do Mac.
# Depois disso o preview nao depende do Oracle Home original (ex.: pasta compartilhada de uma VM).
# Os jars nao sao do projeto e nao vao para o repositorio.
set -euo pipefail
source_home="${1:-${ORACLE_HOME:-}}"
# Duplo clique (release): sem argumento, pede a pasta num seletor do macOS.
if [ -z "$source_home" ] && [ -t 0 ]; then
  source_home="$(osascript -e 'POSIX path of (choose folder with prompt "Select the Oracle_Home folder (contains bip/lib and jlib)")' 2>/dev/null || true)"
  source_home="${source_home%/}"
fi
if [ -z "$source_home" ]; then
  printf '%s\n' 'Usage: bun run oracle:import "/path/to/Oracle_Home"' >&2
  exit 1
fi
if [ ! -f "$source_home/bip/lib/xdo-core.jar" ] || [ ! -f "$source_home/jlib/orai18n.jar" ]; then
  printf '%s\n' "BI Publisher engine not found in $source_home (expected bip/lib/xdo-core.jar and jlib/orai18n.jar)." >&2
  exit 1
fi
target="$HOME/Library/Application Support/BIPTemplateEditor/oracle-lib"
mkdir -p "$target"
rm -f "$target"/*.jar
cp "$source_home"/bip/lib/*.jar "$target/"
# orai18n: OraCollator (obrigatorio); collation/mapping para ordenacao e charsets de outros idiomas.
for jar in orai18n.jar orai18n-collation.jar orai18n-mapping.jar; do
  if [ -f "$source_home/jlib/$jar" ]; then cp "$source_home/jlib/$jar" "$target/"; fi
done
printf 'Source: %s\nCopied: %s\n' "$source_home" "$(date '+%Y-%m-%d %H:%M')" > "$target/SOURCE.txt"
printf '%s\n' "Oracle engine copied to $target ($(ls "$target"/*.jar | wc -l | tr -d ' ') jars). The original Oracle Home is no longer needed for preview."
