# 10_shijisho-pen — 指図書 書き込み (GAS Web アプリ)

`00_ai-clerk-core` と同じ構成（番号付きプロジェクトフォルダ ＋ `rootDir = ./src`、
ソースは `src/` 配下、`appsscript.json` は Asia/Tokyo / V8 / STACKDRIVER / Drive v3）に揃えたプロジェクト。

## セットアップ — ローカルで1コマンド

`clasp create` / `push` / `open` は Google アカウントの認証とブラウザが必要なため、
クラウド環境では実行できない。手元のマシンで下記を実行する。

```bash
npm i -g @google/clasp   # 未導入なら
clasp login              # 未ログインなら
bash 10_shijisho-pen/setup.sh
```

`setup.sh` が手順 1〜4 を順に実行する。

1. `clasp create --type webapp --title "指図書 書き込み" --rootDir ./src`
   （`.clasp.json` が既にあればスキップ）
2. `~/Downloads` の `Code.gs` / `Index.html` / `appsscript.json` を `src/` へコピー。
   `Code.gs` は `Code.js` にリネーム。
   Downloads が別の場所なら `DOWNLOADS_DIR=/path/to/dir bash 10_shijisho-pen/setup.sh`
3. `src/appsscript.json` を検査。`webapp.access = DOMAIN` と `dependencies` の Drive v3 を確認し、
   不足・相違があれば補正して内容を表示する（timeZone / runtimeVersion / exceptionLogging も揃える）。
4. `clasp push --force` → エディタを開く

## ファイル構成

```
10_shijisho-pen/
├── .clasp.json        # 手順1で生成（scriptId と rootDir）※ clasp が作る
├── setup.sh
├── README.md
└── src/               # clasp の rootDir
    ├── appsscript.json   # webapp/DOMAIN + Drive v3 のひな形（手順2で Downloads 版に差し替え、手順3で補正）
    ├── Code.js           # 手順2で配置（Downloads の Code.gs）
    └── Index.html        # 手順2で配置
```

`src/appsscript.json` はひな形として先に置いてある。`~/Downloads` に `appsscript.json` が
無い場合はこれがそのまま使われる。

## 注意: clasp のバージョン

clasp v3 で `clasp open` は廃止され `clasp open-script` になった。
`setup.sh` は `clasp --version` を見て v3 以上なら `open-script`、v2 なら `open` を呼ぶ。
手で叩く場合も同様に読み替えること。

| 操作 | clasp v2 | clasp v3 |
|---|---|---|
| エディタを開く | `clasp open` | `clasp open-script` |
| プロジェクト作成 | `clasp create` | `clasp create`（`create-script` の別名・オプション同じ） |
