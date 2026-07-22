# Steel-Booking（スチール予約システム）

LINE LIFF から予約を受け付け、Google Apps Script (GAS) を経由して
Google スプレッドシートに保存する予約システムです。

```
LINE → LIFF(予約画面) → GAS Web API → スプレッドシート
                                          ↑ コード管理: GitHub / clasp
公開: GitHub Pages（LIFF・管理画面）
```

## 構成

| 役割 | 実体 |
|------|------|
| 予約画面 | `liff/index.html`（GitHub Pages で公開） |
| 管理画面 | `admin/index.html`（予約一覧の閲覧） |
| 予約処理 | `gas/Code.gs`（GAS Web API: `doPost`/`doGet`） |
| 予約保存 | Google スプレッドシート |
| 公開 | GitHub Pages（`.github/workflows/pages.yml`） |
| GAS 連携 | clasp |

- 公開URL: https://officialwork-design.github.io/Steel-Booking/
  - 予約画面: `/Steel-Booking/liff/`
  - 管理画面: `/Steel-Booking/admin/`
- LIFF URL: https://liff.line.me/2010792348-o9YZUyTV

## 予約フォームの項目

`ライン名`（必須）／`希望日`（必須）／`希望時間`（必須）／`備考`（任意）。
氏名（表示名）と userId は LIFF プロフィールから自動取得します。

## セットアップ手順

### 1. スプレッドシートを用意
新規スプレッドシートを作成し、URL の `/d/` と `/edit` の間の文字列（スプレッドシートID）を控える。
ヘッダー行は GAS 初回実行時に自動作成されます。

### 2. GAS を配置（clasp）
```bash
npm install -g @google/clasp
clasp login
cd gas
# 既存のGASプロジェクトに紐付ける場合:
cp .clasp.json.example .clasp.json   # scriptId を記入
clasp push
```
（GUIで作る場合は Apps Script エディタに `Code.gs` を貼り付けてもOK）

Apps Script の「プロジェクトの設定 > スクリプト プロパティ」に以下を登録:

| キー | 値 | 必須 |
|------|----|------|
| `SHEET_ID` | 手順1のスプレッドシートID | ✅ |
| `SHEET_NAME` | シート名（既定 `予約`） | 任意 |
| `ADMIN_KEY` | 管理画面の閲覧パスコード | 任意 |

### 3. GAS をウェブアプリとしてデプロイ
「デプロイ > 新しいデプロイ > 種類: ウェブアプリ」
- 実行するユーザー: **自分**
- アクセスできるユーザー: **全員**

発行される `https://script.google.com/macros/s/XXXX/exec` を控える。

### 4. エンドポイントを設定
- `liff/index.html` 冒頭の `CONFIG.GAS_ENDPOINT` に `/exec` URL を設定
- `admin/index.html` 冒頭の `GAS_ENDPOINT` に同じ URL を設定
- `CONFIG.LIFF_ID` は `2010792348-o9YZUyTV` 設定済み

### 5. LINE Developers 側の設定
LIFF アプリのエンドポイント URL を
`https://officialwork-design.github.io/Steel-Booking/liff/` に設定。
- サイズ: Full

### 6. GitHub Pages を有効化
リポジトリ Settings > Pages > Build and deployment の Source を
**GitHub Actions** にする。`main` への push で自動デプロイされます。

## 動作確認
1. `https://liff.line.me/2010792348-o9YZUyTV` を LINE で開く
2. ログイン後、ライン名・日時・備考を入力して予約
3. スプレッドシートに行が追加されることを確認
4. `/Steel-Booking/admin/` で一覧を確認

## 注意
- GAS 側で LINE の `idToken` 検証は未実装です（フォームからは送信済み）。
  本番で厳密な本人確認が必要なら `doPost` に検証処理を追加してください。
- 公開設定のトークン・パスコード類は Git にコミットしないでください。
