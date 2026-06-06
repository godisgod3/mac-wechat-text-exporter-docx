#!/usr/bin/env bash
set -euo pipefail

# Example only. Replace the group name with your own local WeChat group name.
python3 scripts/export_chat_docx.py \
  --chat "群名" \
  --output "exports/群名_完整文字聊天记录.docx" \
  --limit 999999

