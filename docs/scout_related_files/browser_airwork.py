from lib.Utilities import ScoutConfig
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
import time

class AirworkSeleniumAutomation:
    def __init__(self, browser_manager, app, max_age, min_age, job_title_value,login_url='https://en-gage.net/company/manage/',logout_url='https://en-gage.net/company_login/auth/logout'):  # コンストラクタを修正

        """
        コンストラクタ
        :param driver: Selenium WebDriverのインスタンス
        :param app: アプリケーションのインスタンス
        """

        self.app = app
        self.bm = browser_manager

        self.config = ScoutConfig()  # 設定クラスのインスタンス化
        self.config.update_age_range(min_age, max_age)  # 両方の年齢を更新
        self.scout_count = 0
        self.wait = WebDriverWait(self.bm.driver, 5)
        self.logout_url = logout_url  # url変数を取得
        self.login_url = login_url

       
    def login(self):
        """ブラウザセッションの初期設定"""
        self.app.log_add("Airworkのログイン処理開始")
        
        try:
            # ログインページにアクセス
            self.bm.driver.get(self.login_url)
            # ログイン状態判定
            WebDriverWait(self.bm.driver, 5).until(
                EC.presence_of_element_located((By.TAG_NAME, 'body'))
                
            )
            try:
                # ログイン状態の確認：ログアウトボタンの存在チェック
                logout_element = WebDriverWait(self.bm.driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="air-common-header"]/div/div[1]/div/span[3]/a'))
                )
                if logout_element:
                    # ログアウト処理
                    myaccount_button = WebDriverWait(self.bm.driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="air-common-header"]/div/ul/li[1]'))
                    )
                    myaccount_button.click()
                    time.sleep(1)
                    logout_element.click()
                    try:
                        alert_dialog_button = WebDriverWait(self.bm.driver, 3).until(
                            EC.element_to_be_clickable((By.XPATH, '//*[@id="alert-dialog"]/div/div/div/footer/button[2]'))
                        )
                        alert_dialog_button.click()
                        self.app.log_add("アラートダイアログのボタンをクリックしました")
                    except Exception as e:
                        self.app.log_add(f"アラートダイアログのボタンをクリック中にエラー: {str(e)}")
                    self.app.log_add("ログアウトしました")

            except Exception as e:
                pass
            #self.app.log_add(f"ログアウト処理中にエラー: {str(e)}")

            # Airworkログイン処理
            # self.app.airwok_id変数の値をXpathに代入
            self.bm.driver.execute_script("document.getElementById('account').value = arguments[0];", self.app.airwork_id)
            
            # self.app.airwok_pass変数の値をXpathに代入
            self.bm.driver.execute_script("document.getElementById('password').value = arguments[0];", self.app.airwork_pass)
            
            # ログインボタンをクリック
            login_button = self.bm.driver.find_element(By.XPATH, '//*[@id="mainContent"]/div/div[2]/div[4]/input')
            login_button.click()
            try:
                # ログイン完了の確認：特定の要素の存在チェック
                login_complete_element = WebDriverWait(self.bm.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/header/div/nav/ul/li[2]/a'))
                )
                if login_complete_element:
                    login_success = True
            except Exception as e:
                self.app.log_add(f"ログイン完了の確認中にエラー: {str(e)}")
                login_success = False
            
            self.app.log_add(f"ログイン結果: {login_success}")
            
            if not login_success:
                self.app.log_add("ログインに失敗しました")
                return False
                
        except Exception as e:
            self.app.log_add(f"セットアップ中に予期せぬエラー: {str(e)}")
            return False
        return True
    
    def is_age_out_of_range(self,min_age, max_age, age):
        try:
            min_age = int(min_age)
            max_age = int(max_age)
            age = int(age)
            return age < min_age or age > max_age
        except ValueError as e:
            self.app.log_add(f"年齢の判定中にエラー: {str(e)}")
            return True

    def select_prefecture(self, driver, prefecture_name):
        """指定された県名のチェックボックスをクリックする"""
        try:
            # ラベル要素を特定して待機
            label = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, 
                    f"//label[.//span[text()='{prefecture_name}']]"))
            )
            
            # JavaScriptでラベルをクリック
            driver.execute_script("arguments[0].click();", label)
            print(f"{prefecture_name}を選択しました")
            return True
            
        except Exception as e:
            print(f"{prefecture_name}の選択中にエラーが発生しました: {str(e)}")
            return False

    def select_multiple_prefectures(self, driver, prefecture_list):
        """複数の県のチェックボックスをクリックする"""
        try:
            # 「条件で候補者を探す」ボタンをクリック
            try:
                search_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, 
                        "//button[contains(text(), '条件で候補者を探す')]"))
                )
                search_button.click()
                print("条件設定を開きました")
            except (TimeoutException, ElementClickInterceptedException):
                print("条件設定ボタンのクリックに失敗しました")
                return False

            # 「設定する」ラジオボタンをクリック
            try:
                radio_label = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, 
                        "//span[text()='設定する']/ancestor::label"))
                )
                driver.execute_script("arguments[0].click();", radio_label)
                print("「設定する」を選択しました")
            except (TimeoutException, ElementClickInterceptedException):
                print("設定するラジオボタンのクリックに失敗しました")
                return False

            # 場所設定ボタンをクリック
            try:
                setting_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, 
                        "//button[contains(text(), '設定する')]"))
                )
                driver.execute_script("arguments[0].click();", setting_button)
                print("設定画面を開きました")
            except (TimeoutException, ElementClickInterceptedException):
                print("設定ボタンのクリックに失敗しました")
                return False

            # 少し待機して都道府県を選択
            import time
            time.sleep(2)

            # 都道府県を選択
            results = {}
            for prefecture in prefecture_list:
                success = self.select_prefecture(driver, prefecture)
                if not success:
                    return False
                time.sleep(1)

            # 保存ボタンをクリック
            try:
                search_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="__next"]/div[5]/div/form/footer/div/section/div[2]/button'))
                )
                search_button.click()
                print("保存しました")
            except (TimeoutException, ElementClickInterceptedException):
                print("保存ボタンのクリックに失敗しました")
                return False
            try:
                # 検索ボタンをクリック
                search_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="__next"]/div[4]/div/form/footer/div/div/section/div[2]/button'))
                )
                search_button.click()
                self.app.log_add("検索ボタンをクリックしました")
                return True
            except Exception as e:
                self.app.log_add(f"検索ボタンのクリック中にエラー: {str(e)}")
                
        except Exception as e:
            print(f"エラーが発生しました: {str(e)}")
            return False