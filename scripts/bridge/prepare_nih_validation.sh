#!/usr/bin/env bash
set -euo pipefail

repo_root="/Users/kosta/LocalDev/Word-MCP-Bridge"
docs_dir="$repo_root/docs"
source_doc="${1:-$docs_dir/NIH_Biocoating_CraftV6.docx}"
truth_doc="$docs_dir/NIH_Biocoating_CraftV6 source-of-truth.docx"
copy_a="$docs_dir/NIH_Biocoating_CraftV6.docx"
copy_b="$docs_dir/NIH_Biocoating_CraftV6 copy.docx"

if [[ ! -f "$source_doc" ]]; then
  echo "Missing source doc: $source_doc" >&2
  exit 1
fi

if [[ ! -f "$truth_doc" ]]; then
  cp "$source_doc" "$truth_doc"
fi

cat <<EOF
NIH validation document roles
- Source of truth: $truth_doc
- Copy A: $copy_a
- Copy B: $copy_b

Recommended lane assignments
- Main lane: Copy A
- Independent validation lane: Copy B
- Offline comparison only: Source of truth
EOF
