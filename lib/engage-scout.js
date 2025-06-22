const ENGAGE_LOGIN_URL = 'https://en-gage.net/company/manage/';
const SCOUT_LIST_URL = 'https://en-gage.net/company/scout/?PK=';
const SCOUT_SEND_URL = 'https://en-gage.net/company/scout/approach/?type_of_scout=1&PK=';

// ランダムな待機時間を生成
const randomWait = (min, max) => {
    return Math.floor(Math.random() * (max - min + 1) + min) * 1000;
};

// ページの読み込み待機
const waitForPageLoad = async (page, timeout = 30000) => {
    await page.waitForNavigation({ waitUntil: 'networkidle2', timeout }).catch(() => {});
};

// Engageへのログイン処理
const loginToEngage = async (page, loginId, password) => {
    console.log('Navigating to Engage login page...');
    await page.goto(ENGAGE_LOGIN_URL, { waitUntil: 'networkidle2' });
    
    // ログインフォームの入力
    await page.waitForSelector('#loginID', { visible: true });
    await page.type('#loginID', loginId);
    await page.type('#password', password);
    
    // ログインボタンをクリック
    await page.click('#login-button');
    
    // ログインエラーのチェック
    try {
        await page.waitForSelector('#login-error-area', { visible: true, timeout: 3000 });
        throw new Error('Login failed - invalid credentials');
    } catch (error) {
        if (error.message.includes('Login failed')) {
            throw error;
        }
        // エラーが表示されない = ログイン成功
    }
    
    // PKの取得
    await page.waitForSelector('#header a[href*="PK="]', { visible: true });
    const pkLink = await page.$eval('#header a[href*="PK="]', el => el.href);
    const pk = new URL(pkLink).searchParams.get('PK');
    
    console.log('Login successful, PK:', pk);
    return pk;
};

// モーダルの処理
const handleModals = async (page) => {
    const modalSelectors = [
        '.modal-close-btn',
        'button[aria-label="閉じる"]',
        '.close-modal',
        '.modal-dismiss'
    ];
    
    for (const selector of modalSelectors) {
        try {
            const modal = await page.$(selector);
            if (modal) {
                await modal.click();
                await page.waitForTimeout(500);
            }
        } catch (error) {
            // モーダルが存在しない場合は無視
        }
    }
};

// 候補者の年齢をチェック
const checkCandidateAge = async (page, minAge, maxAge) => {
    try {
        // 年齢情報の取得（複数のセレクタパターンを試す）
        const ageSelectors = [
            '.age-info',
            '[data-age]',
            'span:has-text("歳")',
            'td:has-text("年齢")'
        ];
        
        for (const selector of ageSelectors) {
            const ageElement = await page.$(selector);
            if (ageElement) {
                const ageText = await ageElement.textContent();
                const ageMatch = ageText.match(/(\d+)歳/);
                if (ageMatch) {
                    const age = parseInt(ageMatch[1]);
                    return age >= minAge && age <= maxAge;
                }
            }
        }
        
        return true; // 年齢が取得できない場合はスキップしない
    } catch (error) {
        console.error('Error checking age:', error);
        return true;
    }
};

// スカウト送信処理
const sendScout = async (page, pk) => {
    try {
        // スカウト送信URLに遷移
        const sendUrl = `${SCOUT_SEND_URL}${pk}`;
        await page.goto(sendUrl, { waitUntil: 'networkidle2' });
        
        // スカウト送信ボタンをクリック
        const sendButtonSelectors = [
            'button[type="submit"]',
            'input[type="submit"]',
            '.scout-send-btn',
            'button:has-text("送信")'
        ];
        
        for (const selector of sendButtonSelectors) {
            const button = await page.$(selector);
            if (button) {
                await button.click();
                await waitForPageLoad(page);
                console.log('Scout sent successfully');
                return true;
            }
        }
        
        return false;
    } catch (error) {
        console.error('Error sending scout:', error);
        return false;
    }
};

// 都道府県での絞り込み
const filterByPrefecture = async (page, prefectures) => {
    if (!prefectures || prefectures.length === 0) {
        return;
    }
    
    try {
        // 絞り込みボタンをクリック
        await page.click('.filter-btn, button:has-text("絞り込み")');
        await page.waitForTimeout(1000);
        
        // 都道府県の選択
        for (const prefecture of prefectures) {
            const checkbox = await page.$(`input[type="checkbox"][value="${prefecture}"]`);
            if (checkbox) {
                await checkbox.click();
            }
        }
        
        // 絞り込み適用
        await page.click('button:has-text("適用")');
        await waitForPageLoad(page);
    } catch (error) {
        console.error('Error filtering by prefecture:', error);
    }
};

// メインのスカウト処理
const engageScout = async (browser, options) => {
    const { loginId, password, minAge, maxAge, prefectures } = options;
    const page = await browser.newPage();
    
    // ユーザーエージェントの設定
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36');
    
    let scoutCount = 0;
    let processedCount = 0;
    
    try {
        // ログイン
        const pk = await loginToEngage(page, loginId, password);
        
        // スカウトリストページへ遷移
        await page.goto(`${SCOUT_LIST_URL}${pk}`, { waitUntil: 'networkidle2' });
        
        // モーダルの処理
        await handleModals(page);
        
        // 都道府県での絞り込み
        await filterByPrefecture(page, prefectures);
        
        // 候補者リストの処理
        let hasNextCandidate = true;
        
        while (hasNextCandidate) {
            processedCount++;
            
            // 候補者の選択
            const candidateLinks = await page.$$('.candidate-link, a[href*="/candidate/"]');
            if (candidateLinks.length === 0) {
                console.log('No more candidates found');
                break;
            }
            
            // 最初の候補者をクリック
            await candidateLinks[0].click();
            await waitForPageLoad(page);
            
            // 年齢チェック
            const isAgeValid = await checkCandidateAge(page, minAge, maxAge);
            
            if (isAgeValid) {
                // スカウト送信
                const sent = await sendScout(page, pk);
                if (sent) {
                    scoutCount++;
                    console.log(`Scout sent: ${scoutCount}`);
                }
            } else {
                console.log('Skipped candidate due to age criteria');
            }
            
            // ランダムな待機
            await page.waitForTimeout(randomWait(4, 5));
            
            // 次の候補者へ
            try {
                await page.click('.next-candidate-btn, button:has-text("次へ")');
                await waitForPageLoad(page);
            } catch (error) {
                hasNextCandidate = false;
            }
            
            // 処理数の上限チェック（Lambda実行時間を考慮）
            if (processedCount >= 50) {
                console.log('Reached processing limit for this execution');
                break;
            }
        }
        
        return {
            scoutCount,
            processedCount,
            status: 'completed'
        };
        
    } catch (error) {
        console.error('Error in engage scout process:', error);
        return {
            scoutCount,
            processedCount,
            status: 'error',
            error: error.message
        };
    } finally {
        await page.close();
    }
};

module.exports = {
    engageScout
};