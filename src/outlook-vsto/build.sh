#!/bin/bash

# PTA Outlook アドイン用 VSTOビルドスクリプト
# VSTOアドインのビルドスクリプト

echo "🚀 PTA Outlook VSTO アドインのビルドを開始します"
echo "================================================="

# 変数設定
PROJECT_NAME="OutlookPTAAddin"
SOLUTION_FILE="$PROJECT_NAME.sln"
PROJECT_FILE="$PROJECT_NAME.csproj"
OUTPUT_DIR="bin/Release"
PUBLISH_DIR="publish"

# 前提条件チェック
echo "📋 前提条件をチェック中..."

# .NET Framework 4.8 SDK のチェック
if ! command -v msbuild &> /dev/null; then
    echo "❌ MSBuild が見つかりません。Visual Studio または Build Tools for Visual Studio をインストールしてください。"
    exit 1
fi

echo "✅ MSBuild が見つかりました"

# NuGet の復元
echo "📦 NuGet パッケージを復元中..."
nuget restore $SOLUTION_FILE
if [ $? -ne 0 ]; then
    echo "❌ NuGet パッケージの復元に失敗しました"
    exit 1
fi
echo "✅ NuGet パッケージの復元完了"

# デバッグビルド
echo "🔨 デバッグビルドを実行中..."
msbuild $SOLUTION_FILE /p:Configuration=Debug /p:Platform="Any CPU" /verbosity:minimal
if [ $? -ne 0 ]; then
    echo "❌ デバッグビルドに失敗しました"
    exit 1
fi
echo "✅ デバッグビルド完了"

# リリースビルド
echo "🔨 リリースビルドを実行中..."
msbuild $SOLUTION_FILE /p:Configuration=Release /p:Platform="Any CPU" /verbosity:minimal
if [ $? -ne 0 ]; then
    echo "❌ リリースビルドに失敗しました"
    exit 1
fi
echo "✅ リリースビルド完了"

# ClickOnce発行準備
echo "📤 ClickOnce発行を準備中..."

# 発行ディレクトリをクリーンアップ
if [ -d $PUBLISH_DIR ]; then
    rm -rf $PUBLISH_DIR
fi
mkdir -p $PUBLISH_DIR

# ClickOnce発行実行
echo "📤 ClickOnce発行を実行中..."
msbuild $PROJECT_FILE /t:Publish /p:Configuration=Release /p:Platform="Any CPU" /p:PublishUrl="$PUBLISH_DIR/" /verbosity:minimal
if [ $? -ne 0 ]; then
    echo "⚠️ ClickOnce発行に失敗しました（オプション機能）"
else
    echo "✅ ClickOnce発行完了"
fi

# 成果物の確認
echo "📁 成果物の確認中..."
if [ -f "$OUTPUT_DIR/$PROJECT_NAME.exe" ]; then
    echo "✅ アドイン実行ファイル: $OUTPUT_DIR/$PROJECT_NAME.exe"
fi

if [ -f "$OUTPUT_DIR/$PROJECT_NAME.dll" ]; then
    echo "✅ アドインライブラリ: $OUTPUT_DIR/$PROJECT_NAME.dll"
fi

if [ -d $PUBLISH_DIR ]; then
    echo "✅ ClickOnce発行ファイル: $PUBLISH_DIR/"
fi

# ビルド完了
echo "================================================="
echo "🎉 ビルドが正常に完了しました！"
echo ""
echo "📋 次のステップ:"
echo "1. $OUTPUT_DIR/ フォルダの内容をOutlookに配置"
echo "2. または $PUBLISH_DIR/ フォルダからClickOnceインストール"
echo "3. OutlookでVSTOアドインを有効化"
echo ""
echo "📚 詳細は DEPLOYMENT.md を参照してください"