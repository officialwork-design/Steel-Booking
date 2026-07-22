#!/usr/bin/env bash
# GASへ反映（コードのpush → 既存デプロイを新バージョンで再デプロイ）
# 使い方: gas フォルダで  bash deploy.sh
set -e
DEPLOY_ID="AKfycbzCZMp7K4BcvoL4Jxd1SC1aB1LXtWT2eSbONSokii6mXVLUfIAMIkWPVI4-Fb7PPsWw7A"
cd "$(dirname "$0")"
echo "▶ clasp push..."
clasp push -f
echo "▶ clasp deploy (URL固定)..."
clasp deploy -i "$DEPLOY_ID"
echo "✅ 反映完了"
