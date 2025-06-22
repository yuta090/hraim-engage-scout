#!/bin/bash

# Lambda関数パッケージング用スクリプト

echo "Starting Lambda function packaging..."

# 作業ディレクトリの設定
WORK_DIR=$(pwd)
BUILD_DIR="$WORK_DIR/build"
DIST_DIR="$WORK_DIR/dist"

# ディレクトリのクリーンアップと作成
rm -rf $BUILD_DIR $DIST_DIR
mkdir -p $BUILD_DIR $DIST_DIR

# 必要なファイルをビルドディレクトリにコピー
echo "Copying source files..."
cp -r index.js lib package.json $BUILD_DIR/

# ビルドディレクトリに移動
cd $BUILD_DIR

# 本番用依存関係のインストール（devDependenciesは除外）
echo "Installing production dependencies..."
npm install --production

# 不要なファイルの削除
echo "Cleaning up unnecessary files..."
find . -name "*.md" -type f -delete
find . -name "*.txt" -type f -delete
find . -name ".DS_Store" -type f -delete
find . -name "test" -type d -exec rm -rf {} +
find . -name "tests" -type d -exec rm -rf {} +
find . -name "docs" -type d -exec rm -rf {} +

# ZIPファイルの作成
echo "Creating deployment package..."
zip -r $DIST_DIR/EngageScoutLambda.zip . -x "*.git*"

# 元のディレクトリに戻る
cd $WORK_DIR

# ファイルサイズの確認
echo "Package created successfully!"
echo "Package size: $(du -h $DIST_DIR/EngageScoutLambda.zip | cut -f1)"

# Lambda Layerの情報を表示
echo ""
echo "=== Required Lambda Layers ==="
echo "1. puppeteer-core-layer"
echo "2. sparticuz-chromium-layer"
echo "3. chromium-fonts-layer (for Japanese text)"
echo ""
echo "Make sure these layers are attached to your Lambda function."