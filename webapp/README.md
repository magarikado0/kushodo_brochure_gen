# パンフレット生成 Web アプリ

`webapp/` は、作品情報フォームとパンフ鋳型をブラウザからアップロードして、生成済み Word ファイルをダウンロードするための Web アプリです。

## 構成

```text
webapp/
  backend/      FastAPI
  frontend/     React + Vite
  Caddyfile     HTTPS とリバースプロキシ
  docker-compose.yml
```

アップロードされたファイルは一時ディレクトリで処理し、レスポンス送信後に削除します。

## ローカル確認

Docker が使える環境で、リポジトリのルートから実行します。

```sh
cd webapp
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
cd kushodo_brochure_gen/webapp
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

コードを更新したら、`webapp` フォルダで再ビルドします。

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
