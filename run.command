#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

echo "京大書道部パンフレット生成ツール"
echo

if [ ! -f "input/作品情報フォーム.xlsx" ]; then
  echo "[ERROR] input/作品情報フォーム.xlsx が見つかりません。"
  echo "input フォルダに作品情報フォーム.xlsxを置いてください。"
  echo
  read -r -p "Enterキーで終了します..."
  exit 1
fi

if [ ! -f "templates/冬樟展パンフ.docx" ]; then
  echo "[ERROR] templates/冬樟展パンフ.docx が見つかりません。"
  echo "templates フォルダに前回パンフレットのdocxを置いてください。"
  echo
  read -r -p "Enterキーで終了します..."
  exit 1
fi

if ! command -v uv >/dev/null 2>&1; then
  echo "uv が見つからないため、インストールします。"
  echo "初回のみ少し時間がかかります。"
  curl -LsSf https://astral.sh/uv/install.sh | sh
  export PATH="$HOME/.local/bin:$HOME/.cargo/bin:$PATH"
fi

echo
echo "パンフレットを生成します。"
uv run python scripts/generate.py

echo
echo "完了しました。output フォルダを確認してください。"
read -r -p "Enterキーで終了します..."
