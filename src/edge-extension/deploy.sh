#!/bin/bash

# Shima Edge拡張機能デプロイメントスクリプト
# Copyright (c) 2024 Shima Development Team

set -e

# 設定
EXTENSION_DIR="."
BUILD_DIR="../../dist/edge-extension"
ZIP_NAME="ai-edge-extension-v1.0.0.zip"

echo "🏫 AI Edge拡張機能デプロイメント開始"

# ビルドディレクトリの作成
echo "📁 ビルドディレクトリの準備..."
mkdir -p "$BUILD_DIR"
rm -rf "$BUILD_DIR"/*

# ファイルのコピー
echo "📋 ファイルのコピー..."
cp -r "$EXTENSION_DIR"/* "$BUILD_DIR"/

# 不要ファイルの削除
echo "🧹 不要ファイルのクリーンアップ..."
find "$BUILD_DIR" -name "*.md" -delete
find "$BUILD_DIR" -name ".DS_Store" -delete
find "$BUILD_DIR" -name "Thumbs.db" -delete
rm -f "$BUILD_DIR/deploy.sh"

# マニフェストファイルの検証
echo "🔍 マニフェストファイルの検証..."
if [ ! -f "$BUILD_DIR/manifest.json" ]; then
    echo "❌ エラー: manifest.jsonが見つかりません"
    exit 1
fi

# 必須ファイルの確認
echo "✅ 必須ファイルの確認..."
required_files=(
    "manifest.json"
    "background/background.js"
    "content/content.js"
    "content/content.css"
    "popup/popup.html"
    "popup/popup.js"
    "popup/popup.css"
    "options/options.html"
    "options/options.js"
    "options/options.css"
    "assets/icons/icon16.png"
    "assets/icons/icon32.png"
    "assets/icons/icon48.png"
    "assets/icons/icon128.png"
)

for file in "${required_files[@]}"; do
    if [ ! -f "$BUILD_DIR/$file" ]; then
        echo "❌ エラー: 必須ファイル $file が見つかりません"
        exit 1
    fi
done

# ZIPファイルの作成
echo "📦 拡張機能パッケージの作成..."
cd "$BUILD_DIR"
zip -r "../$ZIP_NAME" . -x "*.DS_Store" "Thumbs.db"
cd - > /dev/null

echo "✅ デプロイメント完了!"
echo "📦 パッケージ: ../../dist/$ZIP_NAME"
echo ""
echo "🚀 インストール手順:"
echo "1. Edgeブラウザで edge://extensions/ を開く"
echo "2. 開発者モードを有効にする"
echo "3. '展開して読み込み' をクリック"
echo "4. ../../dist/edge-extension フォルダを選択"
echo ""
echo "📤 Microsoft Edge Add-ons への提出:"
echo "1. https://partner.microsoft.com/dashboard/microsoftedge にログイン"
echo "2. 新しい拡張機能として ../../dist/$ZIP_NAME をアップロード"
echo "3. 審査プロセスを完了"