# EngageScoutLambda

AWS Lambda上でPuppeteerを使用してEngageのスカウト処理を自動化する関数です。

## 機能概要

- Engageへの自動ログイン
- 年齢範囲と都道府県による候補者フィルタリング
- スカウトメッセージの自動送信
- Lambda環境とローカル環境の両方で動作

## 入力パラメータ

Lambda関数は以下のJSONパラメータを受け取ります：

```json
{
  "id": "login-id@example.com",
  "pass": "password",
  "min_age": 25,
  "max_age": 35,
  "prefectures": ["東京都", "神奈川県"]
}
```

| パラメータ | 型 | 必須 | 説明 |
|-----------|-----|------|------|
| id | string | ✓ | EngageのログインID |
| pass | string | ✓ | Engageのパスワード |
| min_age | number | - | 最小年齢（デフォルト: 21） |
| max_age | number | - | 最大年齢（デフォルト: 60） |
| prefectures | array | - | 都道府県の配列（デフォルト: 全国） |

## セットアップ

### 1. 依存関係のインストール

```bash
npm install
```

### 2. Lambda Layerの準備

以下のLayerが必要です：
- **puppeteer-core-layer**: Puppeteer本体
- **sparticuz-chromium-layer**: Lambda用Chromiumバイナリ
- **chromium-fonts-layer**: 日本語フォント対応

### 3. パッケージング

```bash
./package_function.sh
```

これにより `dist/EngageScoutLambda.zip` が作成されます。

## ローカルテスト

```bash
# 認証情報を設定してテスト実行
node local-test.js
```

## Lambda設定

### 環境設定
- **Runtime**: Node.js 18.x
- **Memory**: 1024 MB以上推奨
- **Timeout**: 5分以上推奨

### 環境変数
- `NODE_ENV`: production（Lambda環境では自動設定）

## デプロイ

1. AWS Lambdaコンソールで新しい関数を作成
2. 作成したzipファイルをアップロード
3. 必要なLayerをアタッチ
4. メモリとタイムアウトを設定

## 注意事項

- Lambda実行時間の制限により、一度の実行で処理できる候補者数は最大50名に制限
- 認証情報は安全に管理し、環境変数やSecrets Managerの使用を推奨
- 本番環境では適切なレート制限を設定してください

## トラブルシューティング

### ログインエラー
- 認証情報が正しいか確認
- Engageのアカウントがロックされていないか確認

### タイムアウトエラー
- Lambda関数のタイムアウト設定を延長
- 処理する候補者数を減らす

### 日本語文字化け
- chromium-fonts-layerが正しくアタッチされているか確認