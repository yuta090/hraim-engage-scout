// Lambda関数をローカルでテストするスクリプト
// 各種サンプルリクエストを使用してテスト可能

const fs = require('fs');
const path = require('path');

// Lambda関数のハンドラーを読み込み
const { handler } = require('./index');

// テストケースの定義
const testCases = [
    {
        name: 'Basic Request',
        file: './sample-request.json'
    },
    {
        name: 'Minimal Request',
        file: './sample-requests/minimal-request.json'
    },
    {
        name: 'Age Only Request',
        file: './sample-requests/age-only-request.json'
    },
    {
        name: 'Prefecture Only Request',
        file: './sample-requests/prefecture-only-request.json'
    }
];

// 単一のテストケースを実行
async function runTestCase(testCase) {
    console.log(`\n${'='.repeat(50)}`);
    console.log(`Testing: ${testCase.name}`);
    console.log('='.repeat(50));
    
    try {
        // JSONファイルを読み込み
        const requestData = JSON.parse(fs.readFileSync(testCase.file, 'utf8'));
        console.log('Request:', JSON.stringify(requestData, null, 2));
        
        // Lambda形式のイベントを作成
        const event = {
            body: JSON.stringify(requestData)
        };
        
        // ハンドラーを実行
        const startTime = Date.now();
        const result = await handler(event);
        const duration = Date.now() - startTime;
        
        // 結果を表示
        console.log('\nResponse Status:', result.statusCode);
        console.log('Response Body:', JSON.parse(result.body));
        console.log(`Execution Time: ${duration}ms`);
        
    } catch (error) {
        console.error('Test failed:', error.message);
    }
}

// メイン処理
async function main() {
    console.log('EngageScoutLambda Local Test Runner');
    console.log('=====================================\n');
    
    // コマンドライン引数をチェック
    const args = process.argv.slice(2);
    
    if (args.length > 0) {
        // 特定のテストケースを実行
        const testIndex = parseInt(args[0]) - 1;
        if (testIndex >= 0 && testIndex < testCases.length) {
            await runTestCase(testCases[testIndex]);
        } else {
            console.error('Invalid test case number');
            console.log('Available test cases:');
            testCases.forEach((tc, idx) => {
                console.log(`  ${idx + 1}: ${tc.name}`);
            });
        }
    } else {
        // 対話的にテストケースを選択
        console.log('Select a test case:');
        testCases.forEach((tc, idx) => {
            console.log(`  ${idx + 1}: ${tc.name}`);
        });
        console.log('\nUsage: node test-lambda-locally.js [test-number]');
        console.log('Example: node test-lambda-locally.js 1');
    }
}

// 環境変数の設定（ローカルテスト用）
process.env.NODE_ENV = 'development';

// 実行
if (require.main === module) {
    main().catch(console.error);
}