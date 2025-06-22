// ローカル環境でのテスト用スクリプト
// 実際のEngage認証情報を使用する前に、必ずテスト環境で確認してください

const testScout = async () => {
    // Lambda関数のハンドラーを読み込み
    const { handler } = require('./index');
    
    // テスト用のイベントデータ
    const testEvent = {
        body: JSON.stringify({
            id: 'your-test-id@example.com',  // 実際のIDに置き換えてください
            pass: 'your-test-password',       // 実際のパスワードに置き換えてください
            min_age: 25,
            max_age: 35,
            prefectures: ['東京都']           // テストする都道府県
        })
    };
    
    console.log('Starting local test...');
    console.log('Test parameters:', JSON.parse(testEvent.body));
    
    try {
        // Lambda関数を実行
        const result = await handler(testEvent);
        
        console.log('\n=== Test Result ===');
        console.log('Status Code:', result.statusCode);
        console.log('Response:', JSON.parse(result.body));
        
    } catch (error) {
        console.error('Test failed:', error);
    }
};

// 環境変数の設定（ローカルテスト用）
process.env.NODE_ENV = 'development';

// テストの実行
testScout();