#!/bin/zsh
set -e
cd "${0:A:h}"
if [[ ! -x .venv/bin/python ]]; then
  uv venv --python 3.12 .venv
fi
uv pip install --python .venv/bin/python -r requirements.txt
PORT="${PORT:-8765}"
echo "发票整理台：http://127.0.0.1:${PORT}"
exec .venv/bin/python -m uvicorn app:app --host 127.0.0.1 --port "$PORT"
