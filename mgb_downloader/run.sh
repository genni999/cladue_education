#!/usr/bin/env bash
# MGB 다운로더 실행 (macOS / Linux)
# 처음 실행하면 가상환경을 만들고 필요한 것들을 자동으로 설치합니다.
set -e
cd "$(dirname "$0")"

if [ ! -x ".venv/bin/python" ]; then
    echo "[1/2] 처음 실행이라 준비 중입니다. 몇 분 걸릴 수 있어요..."
    python3 -m venv .venv
    .venv/bin/python -m pip install --quiet --disable-pip-version-check -r requirements.txt
fi

echo "[2/2] 앱을 실행합니다. http://localhost:8501 을 브라우저에서 엽니다."
echo "종료하려면 Ctrl+C 를 누르세요."
echo

# 앱은 headless로 뜨므로(.streamlit/config.toml) 브라우저는 여기서 열어 준다.
( sleep 6
  if command -v open >/dev/null 2>&1; then open http://localhost:8501
  elif command -v xdg-open >/dev/null 2>&1; then xdg-open http://localhost:8501
  fi ) &

exec .venv/bin/python -m streamlit run app.py
