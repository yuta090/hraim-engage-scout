# EngageScoutLambda リクエストサンプル

Lambda関数に渡すJSONリクエストのサンプル集です。

## パラメータ説明

| パラメータ | 型 | 必須 | デフォルト値 | 説明 |
|-----------|-----|------|-------------|------|
| `id` | string | ✓ | - | EngageのログインID（メールアドレス） |
| `pass` | string | ✓ | - | Engageのパスワード |
| `min_age` | number | - | 21 | スカウト対象の最小年齢 |
| `max_age` | number | - | 60 | スカウト対象の最大年齢 |
| `prefectures` | array | - | [] | 対象都道府県（空配列の場合は全国） |

## サンプルファイル

### 1. 基本的なリクエスト (`../sample-request.json`)
年齢範囲と関東エリアを指定したサンプル

### 2. 最小限のリクエスト (`minimal-request.json`)
必須パラメータのみ（全国対象、21-60歳）

### 3. 年齢のみ指定 (`age-only-request.json`)
年齢範囲のみを指定（全国対象）

### 4. 都道府県のみ指定 (`prefecture-only-request.json`)
関西エリアのみを指定（21-60歳）

### 5. 全都道府県指定 (`all-prefectures.json`)
全47都道府県を明示的に指定

## 使用例

### AWS CLIでのテスト
```bash
aws lambda invoke \
  --function-name EngageScoutLambda \
  --payload file://sample-request.json \
  response.json
```

### API Gateway経由
```bash
curl -X POST \
  https://your-api-gateway-url/engage-scout \
  -H "Content-Type: application/json" \
  -d @sample-request.json
```

### ローカルテスト
```javascript
// local-test.js で使用
const testEvent = {
  body: JSON.stringify(require('./sample-request.json'))
};
```

## 注意事項

1. **認証情報**: 実際の使用時は適切な認証情報に置き換えてください
2. **都道府県名**: 正確な都道府県名（「県」「都」「府」を含む）を使用してください
3. **年齢範囲**: 現実的な範囲を設定してください（例：18歳未満は通常対象外）