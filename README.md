# パンフレット生成 Web アプリ

作品情報フォームをブラウザからアップロードして、生成済み Word ファイルをダウンロードするための Web アプリです。

## 構成

```text
backend/      FastAPI
frontend/     React + Vite
Caddyfile     HTTPS とリバースプロキシ
docker-compose.yml
```

生成ロジックとパンフ鋳型は `backend/app/` 配下に置いています。ユーザーがアップロードするのは作品情報フォーム（xlsx）だけです。
アップロードされたファイルは一時ディレクトリで処理し、レスポンス送信後に削除します。

## ファイルの要件

Web アプリでは、作品情報フォーム（xlsx）だけをアップロードします。パンフ鋳型はサーバー同梱の `backend/app/templates/パンフ鋳型.docx` を使います。

### 作品情報フォーム（xlsx）

フォームには `個人` と `合作` の2シートが必要です。シート名は変えないでください。

列の順番は変わっても大丈夫です。ただし、列名の先頭部分は変えないでください。説明文が後ろに続くのは問題ありません。

個人シートに必要な列:

`氏名`、`ふりがな`、`学年`、`臨書 or 創作`、`書体`、`作品名`、`作品の向き`、`作品サイズ`、`展示場所`、`表装形式`、`釈文`、`作品コメント`

合作シートに必要な列:

`合作参加者全員分`、`臨書 or 創作`、`書体`、`作品名`、`作品の向き`、`作品サイズ`、`展示場所`、`表装形式`、`釈文`、`作品コメント`

学年は `2回生`、`二回生`、`B2`、`修士1回生`、`M1` などに対応しています。出力時は漢数字に変換されます。

合作参加者は、名前の後ろに学年を付けてください。

```text
加藤 杏次郎（B4） 星野 真帆（B4）
```

## ローカル確認

Docker が使える環境で、リポジトリのルートから実行します。

```sh
cp .env.example .env
```

`.env` の `DOMAIN` は、ローカル確認では `localhost` にしておきます。

```env
DOMAIN=localhost
```

起動:

```sh
docker compose up --build
```

ブラウザで `http://localhost` を開きます。

## Oracle Cloud への配置

### 1. VM を用意

Oracle Cloud Always Free で Ubuntu の VM を作成します。外部から `80` と `443` にアクセスできるように、セキュリティリストまたは NSG でポートを開けます。

### 2. ドメインを向ける

利用するドメインの DNS に A レコードを追加し、VM のパブリック IP を指定します。

```text
example.com  A  <Oracle Cloud VM のパブリックIP>
```

### 3. Docker を入れる

VM に SSH して Docker を入れます。

```sh
sudo apt update
sudo apt install -y ca-certificates curl git
curl -fsSL https://get.docker.com | sudo sh
sudo usermod -aG docker $USER
```

一度ログアウトして、再度 SSH してください。

### 4. アプリを配置

リポジトリを VM に置きます。Git を使わない場合は ZIP をアップロードして展開しても構いません。

```sh
cd kushodo_brochure_gen
cp .env.example .env
```

`.env` の `DOMAIN` を自分のドメインに変更します。

```env
DOMAIN=example.com
```

### 5. 起動

```sh
docker compose up -d --build
```

Caddy が自動で HTTPS 証明書を取得します。DNS 反映前に起動すると証明書取得に失敗することがあるので、その場合は少し待ってから再起動してください。

```sh
docker compose restart caddy
```

## 更新

コードを更新したら、リポジトリのルートで再ビルドします。

```sh
docker compose up -d --build
```

## ログ確認

```sh
docker compose logs -f
```

バックエンドだけ見る場合:

```sh
docker compose logs -f backend
```

## 注意

- 認証はまだありません。URL を知っている人は誰でも使えます。
- アップロードできるファイルは1ファイル40MBまでです。
- 作品画像は自動配置しません。生成後の Word ファイルで手入力してください。
