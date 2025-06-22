const puppeteer = require('puppeteer-core');
const chromium = require('@sparticuz/chromium');

// Lambda環境の検出
const isLambdaEnvironment = !!(process.env.AWS_LAMBDA_FUNCTION_NAME || process.env.LAMBDA_RUNTIME_DIR);

// ブラウザ設定の取得
const getBrowserConfig = async () => {
    if (isLambdaEnvironment) {
        return {
            args: chromium.args,
            defaultViewport: chromium.defaultViewport,
            executablePath: await chromium.executablePath(),
            headless: chromium.headless,
        };
    } else {
        // ローカル環境用の設定
        return {
            headless: false,
            executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
            defaultViewport: { width: 1920, height: 1080 },
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-web-security',
                '--disable-features=IsolateOrigins,site-per-process'
            ]
        };
    }
};

// メインのLambdaハンドラー
exports.handler = async (event) => {
    console.log('Received event:', JSON.stringify(event, null, 2));
    
    let browser = null;
    
    try {
        // イベントからパラメータを取得
        const body = typeof event.body === 'string' ? JSON.parse(event.body) : event;
        const {
            id,
            pass,
            min_age,
            max_age,
            prefectures
        } = body;
        
        // パラメータのバリデーション
        if (!id || !pass) {
            return {
                statusCode: 400,
                body: JSON.stringify({
                    error: 'Missing required parameters: id and pass'
                })
            };
        }
        
        // デフォルト値の設定
        const minAge = min_age || 21;
        const maxAge = max_age || 60;
        const targetPrefectures = prefectures || [];
        
        console.log(`Starting scout process with parameters:
            ID: ${id}
            Age Range: ${minAge} - ${maxAge}
            Prefectures: ${targetPrefectures.join(', ')}`);
        
        // ブラウザの起動
        const browserConfig = await getBrowserConfig();
        browser = await puppeteer.launch(browserConfig);
        
        // スカウト処理モジュールの読み込み
        const { engageScout } = require('./lib/engage-scout');
        
        // スカウト処理の実行
        const result = await engageScout(browser, {
            loginId: id,
            password: pass,
            minAge,
            maxAge,
            prefectures: targetPrefectures
        });
        
        return {
            statusCode: 200,
            body: JSON.stringify({
                message: 'Scout process completed successfully',
                result: result
            })
        };
        
    } catch (error) {
        console.error('Error in Lambda handler:', error);
        
        return {
            statusCode: 500,
            body: JSON.stringify({
                error: 'Internal server error',
                message: error.message
            })
        };
        
    } finally {
        // ブラウザのクリーンアップ
        if (browser) {
            await browser.close();
        }
    }
};

// ローカルテスト用
if (!isLambdaEnvironment && require.main === module) {
    const testEvent = {
        body: JSON.stringify({
            id: 'test@example.com',
            pass: 'testpassword',
            min_age: 25,
            max_age: 35,
            prefectures: ['東京都', '神奈川県']
        })
    };
    
    exports.handler(testEvent)
        .then(result => console.log('Test result:', result))
        .catch(error => console.error('Test error:', error));
}