from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException, TimeoutException, UnexpectedTagNameException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import pandas as pd
import os
import platform


class BrowserManager:
    def __init__(self, logger=None):
        self.logger = logger
        self.driver = self.start_browser()  # ブラウザを起動し、driverに割り当てる

    def start_browser(self):
        """ブラウザを起動し、WebDriverインスタンスを設定する"""
        if self.logger:
            self.logger.log_message("engage　ブラウザ起動中…")

        options = Options()
        options.page_load_strategy = 'eager'
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36")

        # ChromeDriverの最新バージョンを自動的に取得してインストール
        webdriver_path = ChromeDriverManager().install()
        
        # ダウンロードされたディレクトリ内の実行可能ファイルを探す
        webdriver_dir = os.path.dirname(webdriver_path)
        executable_name = 'chromedriver.exe' if platform.system() == 'Windows' else 'chromedriver'
        
        # 実行可能ファイルを探す
        found_executable = False
        for root, dirs, files in os.walk(webdriver_dir):
            for file in files:
                full_path = os.path.join(root, file)
                if file == executable_name:
                    #print(f"Found file: {full_path}")
                    if os.access(full_path, os.X_OK):
                        print(f"File is executable: {full_path}")
                        webdriver_path = full_path
                        found_executable = True
                        break
                    else:
                        print(f"File is not executable: {full_path}")
                        # 実行権限を付与する
                        try:
                            os.chmod(full_path, 0o755)
                            #print(f"Execution permission granted: {full_path}")
                            webdriver_path = full_path
                            found_executable = True
                            break
                        except Exception as e:
                            print(f"Failed to grant execution permission: {e}")
            if found_executable:
                break

        if not found_executable:
            raise FileNotFoundError(f"ChromeDriver executable not found or not executable: {webdriver_dir}")

        service = Service(executable_path=webdriver_path)
        self.driver = webdriver.Chrome(service=service, options=options)
        return self.driver

    def set_radio_button(self, identifier, identifier_label,value, by='id'):
        """
        指定したラジオボタンを選択する
        :param identifier: ラジオボタンのIDまたはname
        :param value: 選択するラジオボタンの値
        :param by: 検索方法 ('id' または 'name')
        :param identifier_label: ラジオボタンのラベルのIDまたはname
        :param by_label: ラベルの検索方法 ('id' または 'name')
        :return: 成功した場合はTrue、失敗した場合はFalse
        """
        try:
            if by == 'id':
                radio_buttons = self.driver.find_elements(By.ID, identifier)
                label_elements = self.driver.find_element(By.ID, identifier_label)
            elif by == 'name':
                radio_buttons = self.driver.find_elements(By.NAME, identifier)
                label_elements = self.driver.find_element(By.NAME, identifier_label)
            else:
                raise ValueError("Invalid 'by' parameter. Use 'id' or 'name'.")


            for index, label_element in enumerate(label_elements):
                if label_element.text == value:
                    radio_buttons[index].click()
                    return True, label_element.text
            return False, label_element  # 指定した値のラジオボタンが見つからなかった場合
        except NoSuchElementException:
            return False, None  # 要素が見つからなかった場合
        except UnexpectedTagNameException:
            return False, None  # 予期しないタグ名の場合
        
    def close_existing_browser(self, app):
        """
        既にBMでブラウザを開いている場合はブラウザを閉じる
        """
        if hasattr(self, 'driver') and self.driver is not None:
            try:
                self.driver.quit()
                app.log_add("既存のブラウザを閉じました。")
            except Exception as e:
                app.log_add(f"既存のブラウザを閉じる際にエラーが発生しました: {e}")
            finally:
                self.driver = None

    def get_selected_radio_button_label(self, identifier, by='id'):
        try:
            if by == 'id':
                radio_buttons = self.driver.find_elements(By.ID, identifier)
            elif by == 'name':
                radio_buttons = self.driver.find_elements(By.NAME, identifier)
            else:
                raise ValueError("Invalid 'by' parameter. Use 'id' or 'name'.")

            for radio_button in radio_buttons:
                if radio_button.is_selected():
                    # ラジオボタンに対応する<label>要素を取得
                    label = self.driver.find_element(By.XPATH, f"//label[@for='{radio_button.get_attribute('id')}']")
                    return label.text

            return None  # ラジオボタンが見つからなかった場合
        except NoSuchElementException:
            return None  # 要素が見つからなかった場合
        except UnexpectedTagNameException:
            return None  # 予期しないタグ名の場合


    def get_selected_option_text(self, identifier, by='id'):
        try:
            if by == 'id':
                select_element = self.driver.find_element(By.ID, identifier)
            elif by == 'name':
                select_element = self.driver.find_element(By.NAME, identifier)
            else:
                raise ValueError("Invalid 'by' parameter. Use 'id' or 'name'.")

            # 要素が<select>タグであることを確認
            if select_element.tag_name.lower() != 'select':
                raise UnexpectedTagNameException(f"Select only works on <select> elements, not on {select_element.tag_name}")

            select_instance = Select(select_element)
            selected_option = select_instance.first_selected_option
            return selected_option.text
        except NoSuchElementException:
            return ""
        except UnexpectedTagNameException :
            return ""
        
    def get_checked_checkboxes(self, identifier: str, by: str = 'id') -> list:
        """
        指定したチェックボックスの中から、チェックのついているチェックボックスの名前の一覧を取得する
        :param identifier: チェックボックスのID、name、またはXPATH
        :param by: 検索方法 ('id', 'name', 'xpath')
        :return: チェックのついているチェックボックスの名前のリスト
        """
        if by == 'id':
            checkboxes = self.driver.find_elements(By.CSS_SELECTOR, f'input[id^="{identifier}"]')
        elif by == 'name':
            checkboxes = self.driver.find_elements(By.CSS_SELECTOR, f'input[name^="{identifier}"]')
        elif by == 'xpath':
            checkboxes = self.driver.find_elements(By.XPATH, identifier)
        else:
            raise ValueError("Invalid 'by' parameter. Use 'id', 'name', or 'xpath'.")
        
        #値が空なら空白を返す
        if not checkboxes:
            return ""
        
        checked_checkboxes = []
        for checkbox in checkboxes:
            if checkbox.is_selected():
                # チェックボックスのラベルを取得
                label = checkbox.find_element(By.XPATH, './following-sibling::label').text
                checked_checkboxes.append(label)

        return checked_checkboxes
    
    def get_input_label_and_value(self, identifier, identifier_type='id'):
        """
        指定されたid、name、またはxpathに基づいて入力欄の値とラベルを取得する関数。
        Args:
            identifier (str): 要素を特定するためのid、name、またはxpath。
            identifier_type (str): identifierのタイプ（'id', 'name', 'xpath'のいずれか）。
        Returns:
            tuple: ラベル名と入力値のタプル（ラベルが見つからない場合はNone）。
        """
        try:
            if identifier_type == 'id':
                input_element = self.driver.find_element(By.ID, identifier)
            elif identifier_type == 'name':
                input_element = self.driver.find_element(By.NAME, identifier)
            elif identifier_type == 'xpath':
                input_element = self.driver.find_element(By.XPATH, identifier)
            else:
                raise ValueError("identifier_typeは'id', 'name', 'xpath'のいずれかである必要があります。")

            # 入力値を取得
            input_value = input_element.get_attribute('value')
            return (input_value)

        except NoSuchElementException:
            print(f"指定された要素が見つかりませんでした: {identifier}")
            return (None, None)
        except TimeoutException:
            print(f"指定された要素の取得がタイムアウトしました: {identifier}")
            return (None, None)
        except Exception as e:
            print(f"エラーが発生しました: {e}")
            return (None, None)
        
    def get_before_content_text(self, identifier, identifier_type='id'):
        """
        指定されたid、name、またはxpathに基づいて要素のBeforeコンテンツテキストを取得するメソッド。
        Args:
            identifier (str): 要素を特定するためのid、name、またはxpath。
            identifier_type (str): identifierのタイプ（'id', 'name', 'xpath'のいずれか）。
        Returns:
            str: Beforeコンテンツのテキスト（見つからない場合はNone）。
        """
        try:
            if identifier_type == 'id':
                element = self.driver.find_element(By.ID, identifier)
            elif identifier_type == 'name':
                element = self.driver.find_element(By.NAME, identifier)
            elif identifier_type == 'xpath':
                element = self.driver.find_element(By.XPATH, identifier)
            else:
                raise ValueError("identifier_typeは'id', 'name', 'xpath'のいずれかである必要があります。")

            before_content = self.driver.execute_script(
                "return window.getComputedStyle(arguments[0], '::before').getPropertyValue('content');",
                element
            )
            return before_content
        except NoSuchElementException:
            print(f"指定された要素が見つかりませんでした: {identifier}")
            return None
        except TimeoutException:
            print(f"指定された要素の取得がタイムアウトしました: {identifier}")
            return None
        except Exception as e:
            print(f"エラーが発生しました: {e}")
            return None
        
    def get_element_text(self, identifier, identifier_type='id'):
        """
        指定されたid、name、またはxpathに基づいて要素のテキストを取得するクラスメソッド。
        Args:
            identifier (str): 要素を特定するためのid、name、またはxpath。
            identifier_type (str): identifierのタイプ（'id', 'name', 'xpath'のいずれか）。
        Returns:
            str: 要素のテキスト（見つからない場合はNone）。
        """
        try:
            if identifier_type == 'id':
                element = self.driver.find_element(By.ID, identifier)
            elif identifier_type == 'name':
                element = self.driver.find_element(By.NAME, identifier)
            elif identifier_type == 'xpath':
                element = self.driver.find_element(By.XPATH, identifier)
            else:
                raise ValueError("identifier_typeは'id', 'name', 'xpath'のいずれかである必要があります。")

            return element.text
        except NoSuchElementException:
            print(f"指定された要素が見つかりませんでした: {identifier}")
            return None
        except TimeoutException:
            print(f"指定された要素の取得がタイムアウトしました: {identifier}")
            return None
        except Exception as e:
            print(f"エラーが発生しました: {e}")
            return None

    def log_image_paths(self, search_type, search_value, iterations=6):
        """
        指定された検索タイプと検索値に基づいて画像パスをログに記録する。
        
        Args:
            search_type (str): 検索タイプ ('id', 'xpath', 'name', 'class')。
            search_value (str): 検索値。
            iterations (int): 繰り返し回数。
        
        Returns:
            list: 画像パスのリスト。
        """
        img_src = []
        try:
            if search_type == 'id':
                elements = self.driver.find_elements(By.ID, search_value)
            elif search_type == 'xpath':
                elements = self.driver.find_elements(By.XPATH, search_value)
            elif search_type == 'name':
                elements = self.driver.find_elements(By.NAME, search_value)
            elif search_type == 'class':
                elements = self.driver.find_elements(By.CLASS_NAME, search_value)
            else:
                raise ValueError("Invalid search_type. Use 'id', 'xpath', 'name', or 'class'.")

            for i, element in enumerate(elements):
                if i >= iterations:
                    break
                try:
                    img_tag = element.find_element(By.TAG_NAME, 'img')
                    img_src.append(img_tag.get_attribute('src'))
                    
                    # 画像パスをログに記録する
                    if img_src[i]:
                        # 画像ファイルにSVGが含まれていれば、画像はアップロードされていないのでNoneを返す。
                        if 'svg' in img_src[i]:
                            img_src[i] = None
                        else:
                            #print(f"画像パス {i+1}: {img_src[i]}")
                            if self.logger:
                                self.logger.log_add(f"画像パス {i+1}: {img_src[i]}")
                    else:
                        print(f"画像パス {i+1}: 画像のsrc属性が見つかりませんでした。")
                        if self.logger:
                            self.logger.log_add(f"画像パス {i+1}: 画像のsrc属性が見つかりませんでした。")
                except NoSuchElementException:
                    print(f"画像パス {i+1}: 要素が見つかりませんでした。")
                    if self.logger:
                        self.logger.log_add(f"画像パス {i+1}: 要素が見つかりませんでした。")
                except Exception as e:
                    print(f"画像パス {i+1}: エラーが発生しました: {e}")
                    if self.logger:
                        self.logger.log_add(f"画像パス {i+1}: エラーが発生しました: {e}")
        except NoSuchElementException:
            print(f"要素が見つかりませんでした: {search_value}")
            if self.logger:
                self.logger.log_add(f"要素が見つかりませんでした: {search_value}")
        except Exception as e:
            print(f"エラーが発生しました: {e}")
            if self.logger:
                self.logger.log_add(f"エラーが発生しました: {e}")
        return img_src
    
    def is_element_clickable(self, driver, element, timeout=10):
        try:
            WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(element))
            return True
        except TimeoutException:
            return False
    def convert_columns_to_object(self,df):
        """
        Convert specified columns of the DataFrame to object (string) type.
        
        Parameters:
        df (pd.DataFrame): The DataFrame whose columns need to be converted.
        
        Returns:
        pd.DataFrame: The DataFrame with specified columns converted to object type.
        """
        columns_to_convert = [
             "掲載状況", "ステータス","■更新フラグ","原稿エラー","原稿更新URL","work_id","求人ID", "雇用形態", "派遣:選択", "仕事No.", "試用期間あり", "正社員登用あり", 
            "職種", "表示用職種名", "仕事内容", "職種カテゴリー", "職種カテゴリー_該当なし", "英語・外国語を使う仕事", 
            "マネージャー・管理職採用", "新規事業", "法人名（正式社名）", "事業内容", "上場企業", "官公庁・学校関連", 
            "ベンチャー企業", "外資系企業", "株式公開準備中", "設立30年以上", "女性管理職登用実績あり", "勤務先名", 
            "勤務先区分", "郵便番号", "都道府県", "市区町村", "以降の住所", "詳細な住所は決まっていない", "最寄り駅", 
            "駅から徒歩5分以内", "転勤なし", "テレワーク・在宅OK", "勤務地：備考", "アクセス", "求人区分", "給与タイプ", 
            "給与（最低額）", "給与（最高額）", "支払方法", "想定年収（最低額）", "想定年収（最高額）", "年収例_1", 
            "入社歴_1", "役職例_1", "年収例_2", "入社歴_2", "役職例_2", "年収例_3", "入社歴_3", "役職例_3", "みなし残業代", 
            "給与：備考", "勤務時間", "想定勤務：開始時間", "想定勤務：開始分", "想定勤務：終了時間", "想定勤務：終了分", 
            "勤務時間：備考", "残業なし", "完全土日祝休み", "1日4h以内OK", "週1日からOK", "週2～3日からOK", "単発・1日のみOK", 
            "短期（1ヶ月以内）", "土日祝のみOK", "10時以降に始業", "夜勤・深夜・早朝（22時～7時）", "16時前までの仕事", 
            "17時以降に始業", "春・夏・冬休み期間限定", "長期（3か月以上）", "最終学歴", "募集職種の経験有無", 
            "その他必要な経験・資格など", "高校生歓迎", "既卒・第二新卒歓迎", "シニア歓迎", "障がい者積極採用", 
            "副業・WワークOK", "英語力不問", "募集人数・募集背景", "10名以上の大量募集", "急募！内定まで2週間", "任意", 
            "欠員補充", "増員", "休みの取り方", "年間休日120日以上", "夏季休暇", "年末年始休暇", "休日休暇", "勤務体制：開始", 
            "勤務体制：体制", "雇用保険", "労災保険", "厚生年金", "健康保険", "交通費支給あり", "資格取得支援・手当あり", 
            "寮・社宅・住宅手当あり", "育児支援・託児所あり", "U・Iターン支援あり", "時短勤務制度あり", "日払い・週払い・即日払いOK", 
            "服装自由", "待遇・福利厚生", "画像ファイル_1", "画像コメント_1", "画像ファイル_2", "画像コメント_2", "画像ファイル_3", 
            "画像コメント_3", "画像ファイル_4", "画像コメント_4", "画像ファイル_5", "画像コメント_5", "画像ファイル_6", "画像コメント_6","_取得日時","page番号","原稿行","engage","engage_update"

        ]

        # データフレームに存在しないカラムを追加
        missing_columns = [col for col in columns_to_convert if col not in df.columns]
        if missing_columns:
            df = pd.concat([df, pd.DataFrame(columns=missing_columns)], axis=1)
        
        # すべてのカラムをオブジェクト型に変換
        df = df.astype({col: "object" for col in columns_to_convert})
        
        return df