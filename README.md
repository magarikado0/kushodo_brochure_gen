# 京大書道部パンフレット生成ツール

`作品情報フォーム.xlsx` から、前回パンフレット `冬樟展パンフ.docx` の書式を使ってパンフレットを生成します。

## 生成されるもの

- `output/パンフレット_テンプレ流し込み.docx`
  - 作品一覧
  - 各作品の詳細テキスト
  - 作品画像を手入力するためのプレースホルダ
- `output/作品一覧.txt`
  - 並び順確認用のテキスト

## 使い方

```powershell
python scripts\generate.py
```

別の Excel ファイルを使う場合:

```powershell
python scripts\generate.py --input path\to\作品情報フォーム.xlsx
```

別の前回パンフレットをテンプレートにする場合:

```powershell
python scripts\generate.py --template path\to\前回パンフ.docx
```

出力先を変える場合:

```powershell
python scripts\generate.py --docx output\パンフレット.docx --list output\作品一覧.txt
```

## 並び順

個人作品は `学年順 -> ふりがな順` で並びます。合作は個人作品の後ろにまとめます。

学年は以下の表記に対応しています。

- `2回生`
- `修士1回生`
- `博士1回生`
- `B2`
- `M3`
- `D1`

## 注意

画像は自動配置しません。生成された `output/パンフレット_テンプレ流し込み.docx` の `【作品画像：ここに手入力で配置】` の位置に、あとから手で配置してください。

作品一覧の右端の数字は仮のページ番号です。画像配置やレイアウト調整後に、必要に応じて手で直してください。

実行時に `確認事項` が表示された場合は、該当作品のフォーム回答に空欄があります。
