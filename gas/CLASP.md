# clasp で GAS を更新する手順

## 初回だけ
> ⚠ 事前に https://script.google.com/home/usersettings で
>   「Google Apps Script API」を **オン** にしておくこと（オフだと push が失敗します）。

```bash
npm install -g @google/clasp   # clasp を入れる
clasp login                    # ブラウザで自分のGoogleにログイン（本人操作が必要）
```

## コードを反映（毎回）
```bash
cd gas
clasp push -f                  # Code.gs / appsscript.json をGASへアップロード
```
`.clasp.json` に scriptId は設定済み（このプロジェクトのGAS）。

## /exec を更新（既存URLを変えずに再デプロイ）
```bash
clasp deployments              # 一覧から /exec のデプロイID(AKfy...で始まる)を確認
clasp deploy -i <デプロイID>    # そのデプロイを新バージョンに更新（URL不変）
```
> `clasp deploy` を ID 指定なしで実行すると **新しいURL** が発行され、
> liff/admin の GAS_ENDPOINT を貼り替える必要が出るので注意。

## スクリプトプロパティ（コードでは持てない設定）
`clasp open` でエディタを開き、⚙️プロジェクトの設定 → スクリプトプロパティ に
`SHEET_ID` を登録する（下記）。
- SHEET_ID = 1neHxod-oOulSHjkbd41hoZ9emkkJPtHvvaQgO6wNvoE
- ADMIN_KEY = （任意）管理画面のパスコード

## 補足
- 公開範囲(全員/自分実行)は `appsscript.json` の webapp 設定で管理済み。
- 初回 push 後の実行時に権限承認を求められたら許可すること。
