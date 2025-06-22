import pandas as pd
import threading
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException

import time
import logging

from tkinter import messagebox
import tkinter as tk

import os
from collections import deque

from datetime import datetime
from lib.Utilities import AddressParser,is_file_open
from lib.auth import EngageIdManager  # 追加

# このファイル専用のロガーを作成
logger = logging.getLogger(__name__)

class EngageSeleniumAutomation:
    def __init__(self, driver, app):
        """
        コンストラクタ
        :param driver: Selenium WebDriverのインスタンス
        :param check_xpath: チェックするXPath
        """
        self.driver = driver
        self.app = app
        self.rowdata_value = None
        self.data = {}
    def login(self, url, login_id, password,xpath_login_error=None):
        """
        Engageにログインするメソッド
        :param url: ログインページのURL
        :param login_id: ログインID
        :param password: パスワード
        """
        self.driver.get(url)
        # パスワード入力フィールドが表示されるまで待機
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.ID, "password")))
        # メールアドレスとパスワードを入力
        self.driver.execute_script("document.getElementById('loginID').value = arguments[0];", login_id)
        self.driver.execute_script("document.getElementById('password').value = arguments[0];", password)
        # ログインボタンを見つけてクリック
        login_button = self.driver.find_element(By.XPATH, '//*[@id="login-button"]')
        login_button.click()

        try:
            # ページの読み込みが完了するまで最大15秒間待機し、指定されたXPathが存在するか確認
            if not xpath_login_error:
                xpath_login_error = '//*[@id="login-error-area"]'
            WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, xpath_login_error))
            )
            # エラーが表示された場合、ログインエラーの処理を返す。
            self.app.show_message("", "error", "ログインエラーです。accounts.xlsxのID,PASSを確認してください。")
            self.driver.quit()
            return False
        except TimeoutException:
            # エラーが表示されなかった場合、ログイン成功とみなす。
            pk_value = ""   #URLに必ず付くpkというセッション変数なのか、公開鍵のURLを取得する。
            self.app.pk_value = ""
            try:
                # 指定されたXPathで要素を検索
                header_link_element = self.driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/a')
                # 要素が見つかった場合、href属性からpkの値を抽出
                href_value = header_link_element.get_attribute('href')
                if 'PK=' in href_value:
                    pk_value = href_value.split('PK=')[1].split('&')[0]
                    self.app.pk_value = pk_value
            except NoSuchElementException:
                # 要素が見つからなかった場合は何も返さない
                pass
            return True ,pk_value
    
    def initialize_data(self, row):
        """
        データフレームからデータを読み込み、内部辞書に格納する。
        """

        # チェックしたい列名のリスト
        indices_to_check = [
            '雇用形態', '職種', '表示用職種名', '仕事内容', '職種カテゴリー', '法人名（正式社名）', '事業内容', '勤務先名', '勤務先区分', '郵便番号', '都道府県', '市区町村',
            '求人区分', '給与タイプ', '給与（最低額）', '給与（最高額）'
        ]

        # 空白の値を持つインデックスを格納するリスト
        empty_indices = []

        # 指定したインデックスに対して値が空白かどうかをチェック
        for index in indices_to_check:
            # インデックスが存在し、値が空白（None、NaN、空文字列）かどうかをチェック
            if index in row and (pd.isnull(row[index]) or str(row[index]).strip() == ''):
                empty_indices.append(index)

        # 空白の値を持つインデックスがある場合、それらを表示
        if empty_indices:
            print("以下のインデックスに空白の値が存在します:")
            for index in empty_indices:
                print(index)

    def click_map(self, column_name):
        """
        マップをクリックするメソッド。
        :param column_name: 列名
        """
        if column_name == "最寄り駅":
            xpath_map = '//*[@id="jobMakeForm"]/div[1]/div[2]/dl/dd[4]/div/div/dl/dd[6]/div[2]/a[2]'
            xpath_map_submit = '/html/body/div[14]/div/div[2]/div[2]/a[1]'
            element = self.driver.find_element(By.XPATH,xpath_map)

            # クリックできるまで30回繰り返す
            for _ in range(30):
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
                    ActionChains(self.driver).move_to_element(element).perform()
                    time.sleep(1)           
                    self.fill_web_form('XPATH', xpath_map, 'click')
                    time.sleep(3)

                    self.fill_web_form('XPATH', xpath_map_submit, 'click')
                    break  # 成功したらループを抜ける
                except Exception as e:
                    print(f"マップをクリックできませんでした。: {e}")
                    if _ == 29:
                        print("30回の試行が完了しましたが、要素の操作が成功しませんでした。")
                        self.app.log_add("マップをクリックできませんでした")
                        return False
                    time.sleep(1)
                    continue
            #フォーカスをマップに合わせる
            time.sleep(1)
            self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
            ActionChains(self.driver).move_to_element(element).perform()
            return True
        
    def fill_address_from_station(self,row, tp, xpath_zip, xpath_auto_addr_input_button, xpath_street):
        """
        最寄り駅から住所と郵便番号を取得し、フォームに入力する関数。
        """
        station_name = None
        station_zip = None

        if not pd.isnull(row['最寄り駅']):
            # 最寄り駅の値からstation.csvを検索し、住所と郵便番号を取得
            station_name = tp.remove_station_name(row['最寄り駅'])
            address, station_zip = tp.search_address_in_csv(station_name, 'lib_data/station.csv')

            # 郵便番号の住所のXpath。
            # 郵便番号に、station.csvで取得した郵便番号を挿入
            self.fill_web_form('XPATH', xpath_zip, 'send_key')

            # 住所の自動入力
            self.fill_web_form('XPATH', xpath_auto_addr_input_button, 'click')

            # 以降の住所から番地を除去する
            parser = AddressParser(address)
            street_number = parser.get_street_number()
            if street_number:
                # 渡した以降の住所から番地を取得し、現在の以降の住所から番地を除去し、テキストを挿入
                address = address.replace(street_number, '').strip()
                self.fill_web_form('XPATH', xpath_street, 'send_key')
                
    def fill_web_form(self, selector_type, selector_value, action_type, input_value=None,max_input_length=None):
        """
        指定されたセレクタとアクションタイプに基づいてWeb要素を操作する。
        :param selector_type: 要素を特定するためのセレクタタイプ ('id', 'name', 'xpath', 'css_selector')
        :param selector_value: セレクタの値
        :param action_type: 実行するアクション ('send_keys', 'click', 'clear', 'select_by_value', 'select_by_index', 'select_by_visible_text')
        :param input_value: 'send_keys', 'select_by_value', 'select_by_index', 'select_by_visible_text' などのアクションで使用する入力値
        """

        try:
            #要素が見つかるまで最大10秒待機
            element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((getattr(By, selector_type.upper()), selector_value))
            )
            # スクロールして要素をビューポートに表示
            self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
            if action_type == 'send_keys' and input_value is not None:
                if input_value is not None and max_input_length is not None:
                    input_value = str(input_value)
                    # 既存のテキストをクリアする
                    element.clear()
                    max_input_length_len = None
                    input_count_value = None
                    max_input_length_len = int(max_input_length)
                    input_count_value = input_value.replace('\n', '')  # 改行を削除
                    input_count_value = input_count_value.replace(' ', '')  # 空白を削除

                    if len(str(input_count_value)) <= max_input_length :
                        time.sleep(0.5)
                        #element.send_keys(input_value)
                        self.set_value_with_js(self.driver,element,input_value)
                    else:
                        if self.app.setting_text_over == '削除':
                            # 入力値が最大文字数を超えている場合、削除する
                            element.clear()
                            self.app.log_add(f"{self.app.column_name}_{action_type}_NG: 入力値が最大文字数を超過のため削除")
                        elif self.app.setting_text_over == '最大文字数にまとめる':
                            # 入力値が最大文字数を超えている場合、設定された処理を実施
                            # 分割して送信するために、入力値をチャンクに分割してソウシン
#                            self.send_keys_in_chunks(element, input_value[:max_input_length], 50)
                            self.set_value_with_js(self.driver,element,input_value)

                            self.app.log_add(f"{self.app.column_name}_{action_type}_NG: 入力値が最大文字数を超えています。{self.app.setting_max_length}に従って処理します。")
                        else:#何も指定がなければ全文を挿入する
#                            element.send_keys(input_value)
                            self.set_value_with_js(self.driver,element,input_value)

            elif action_type == 'click':
                # チェックボックスが既にクリックされているかどうかを判定
                # 前の兄弟要素がinputタグであるかどうかを判定し、クリックされているかどうかを確認
                previous_input_element = self.driver.execute_script("""
                    var element = arguments[0];
                    var previousInput = null;
                    while (element = element.previousElementSibling) {
                        if (element.tagName.toLowerCase() === 'input') {
                            previousInput = element;
                            break;
                        }
                    }
                    return previousInput;
                """, element)

                # 前のinput要素が存在する場合、その要素がクリックされているかどうかを判定
                is_previous_input_clicked = False
                if previous_input_element:
                    is_previous_input_clicked = previous_input_element.get_attribute("aria-pressed") == "true" or previous_input_element.get_attribute("checked") == "true"

                # 現在の要素または前のinput要素がクリックされている場合、クリック処理をスキップ
                if  is_previous_input_clicked == True:
                    print("クリック済みのためパス")
                    pass
                else:
                    # 要素がクリックされていない場合、クリック処理を実行
                    element.click()

            elif action_type == 'clear':
                element.clear()
            elif action_type in ['select_by_value', 'select_by_index', 'select_by_visible_text']:
                select = Select(element)
                #input_valueの値が整数であれば文字列に変換
                if isinstance(input_value, int):
                    input_value = str(input_value)
                if action_type == 'select_by_value':
                    select.select_by_value(input_value)
                elif action_type == 'select_by_index':
                    select.select_by_index(input_value)
                elif action_type == 'select_by_visible_text':
                    select.select_by_visible_text(input_value)
            # その他のSelenium操作を追加可能
        except Exception as e:
            self.app.log_add(f"{self.app.column_name}_{action_type}_NG")
            logger.info(f"{self.app.column_name}_{action_type}_NG_Xpath：{selector_value} \n{e}")
            #return
    def send_keys_in_chunks(self,element, input_value, max_input_length):
        """
        指定された要素に対して、入力値をチャンクに分割して送信する。
        :param element: 文字列を送信する対象のWebElement
        :param input_value: 送信する文字列
        :param max_input_length: 一度に送信する最大文字数
        """
        chunks = [input_value[i:i+max_input_length] for i in range(0, len(input_value), max_input_length)]
        for chunk in chunks:
            element.send_keys(chunk)
            time.sleep(0.1)  # サーバーへの負荷を避けるためにわずかに待機

    def find_xpath_from_series(self, row, df_EngageXpath, column_name,debug_column_name=None):
        """
        rowのSeriesオブジェクトからcolumn_nameに対応する値を取得し、
        df_EngageXpath DataFrameから対応するXPathを検索してインスタンス変数に代入する。
        値が見つからない場合やXPathが見つからない場合はエラーを出力する。
        :param row: 検索する行を含むSeriesオブジェクト
        :param df_EngageXpath: XPathを含むDataFrame
        :param column_name: 値とXPathを検索するための列名
        """
        # rowから特定の列の値を取得し、インスタンス変数に代入
        if column_name in row:
            self.rowdata_value = row[column_name]
        else:
            self.app.log_add(f"エラー: Seriesに'{column_name}'列が存在しません。")

            self.rowdata_value = None
            return  # 早期リターン

        try:
            # デバッグ用の列名が指定されていれば、直接Xpathを指定する。
            if debug_column_name and debug_column_name.strip():
                xpath_value = df_EngageXpath[debug_column_name]
            else:
                # self.rowdata_valueが列名として存在するかどうかは確認が必要です。
                # ここではself.rowdata_valueを列名として使用していると仮定しています。
                xpath_value = df_EngageXpath[self.rowdata_value]
                if xpath_value.empty:
                    xpath_value = df_EngageXpath[column_name]

            if not xpath_value.empty:
                self.xpath = xpath_value.iloc[0]
            else:
                self.app.log_add(f"エラー: '{self.xpath_column_name}'に対応するXPathがdf_EngageXpathに存在しません。")
                self.xpath = None

        except KeyError as e:
            if not debug_column_name:
                column_name = debug_column_name
            self.app.log_add(f"エラー: '{column_name}'列が存在しません。")
            self.xpath = None

        return self.xpath

    def submit_form(self):
        """
        フォームを送信する。
        """
        try:
            submit_button = self.driver.find_element(By.ID, "submitButton")
            submit_button.click()
        except Exception as e:
            print(f"フォームの送信中にエラーが発生しました: {e}")
            raise

    def check_download_file(self, file_path, error_check=None):
        """
        engage-download.xlsxファイルを読み込み、必要に応じて新規作成し、DataFrameを返すメソッド。
        """

        # ファイルの存在チェックと新規作成
        if not os.path.exists(file_path):
            # ファイルが存在しない場合、新規でEXCELファイルを作成
            df_new = pd.DataFrame()  # 空のデータフレームを作成
            # エクセルファイルの中身を削除して開く。
            with pd.ExcelWriter(file_path, mode='w', engine='openpyxl') as writer:
                writer.book.create_sheet('入力リスト')
                writer.sheets['入力リスト'] = writer.book['入力リスト']
                df_new.to_excel(writer, sheet_name='入力リスト', index=False)
            self.app.log_add(f"ファイル {file_path} が存在しなかったため、新規に作成しました。")
        
        try:
            with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:
                df_engage_corporate_download = pd.read_excel(file_path, engine='openpyxl')
                self.app.log_add(f"{file_path} を開きました。")
        except Exception as e:
            self.app.log_add(f"{file_path} を開く際にエラーが発生しました: {e}")
            return None
        
        #ダウンロード保存用のファイルが既に開かれている場合はエラーを返し処理を終了
        if is_file_open(file_path):
            self.app.show_message("", "エラー", f"ファイル {file_path} は既に開かれています。\n閉じてから再度実行してください。")
            self.app.log_add(f"ファイル {file_path} は既に開かれているため処理を終了します。")
            return None
        
        #エラーチェックmodeの時はダイアログは表示せず、バックアップも行わない。
        if error_check is None:
            if not df_engage_corporate_download.empty:
                row_count = len(df_engage_corporate_download)
                # ファイルの最終更新日時を取得
                file_timestamp = os.path.getmtime(file_path)
                # タイムスタンプを日時の形式に変換
                last_modified_date = datetime.fromtimestamp(file_timestamp).strftime('%Y-%m-%d %H:%M:%S')
                message = (
                    f"{file_path}\nに値があります({row_count}行)\n更新日:{last_modified_date}\n\n原稿を取り込んでもよろしいですか？\n"
                    "※原稿は全て削除され新規で取り込まれます。"
                )
                response = self.app.show_message(self.app.page, "確認", message, "askyesno")
                if response:
                    # バックアップディレクトリを作成
                    backup_dir = os.path.join(self.app.company_dir, 'bak')
                    os.makedirs(backup_dir, exist_ok=True)
                    # バックアップディレクトリが存在しない場合は作成
                    if not os.path.exists(backup_dir):
                        os.makedirs(backup_dir)
                    # バックアップファイル名を作成
                    today_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                    backup_file_path = os.path.join(backup_dir, f"{today_str}_engage-download.xlsx")
                    # バックアップを保存
                    df_engage_corporate_download.to_excel(backup_file_path, index=False)
                    self.app.log_add(f"バックアップを {backup_file_path} に保存しました。")
                else:
                    return None
        df_engage_corporate_download = None

        # ダウンロードデータのDFを新規作成
        df_download_data = pd.DataFrame()  # 必要に応じてデータを追加

        
        with pd.ExcelWriter(file_path, mode='w', engine='openpyxl') as writer:
            df_download_data.to_excel(writer, index=False, sheet_name='入力リスト')
            self.app.log_add(f"{file_path} を新規作成モードで開きました。")
        
        return df_download_data
    
    def search_and_click_elements_job_category(self,bm, search_text):
        """
        職種カテゴリーに関する一連のクリック処理"
        候補ボタンをクリックした瞬間、動的に職種カテゴリーが順次読み込まれるため、
        選択ボタンが全て読み込まれたかを判定し、読み込まれない場合はエラー処理handle_job_category_not_found関数を実行する。
        """
        time.sleep(1)
        # 職種カテゴリー：候補ボタンをクリック
        try:
            WebDriverWait(bm.driver, 10).until(EC.element_to_be_clickable((By.ID, 'aiJobSuggestBtn'))).click()
        except:
            print('Element not clickable: By.ID', 'aiJobSuggestBtn')
            return None

        base_xpath = '#setAiJobSuggestTarget > tr'  # 基本となるXPath
        #time.sleep(5)

        #全ての職業カテゴリ選択ボタンが表示されるまで検索
        # 5つ目の選択ボタンが表示されるまで繰り返し探す
        for _ in range(10):
            
            try:
                element =None
                element = bm.driver.find_element(By.CSS_SELECTOR, '#setAiJobSuggestTarget > tr:nth-child(5) > td.data.data--min > a')
                if element.is_displayed():
                    #表示されたらループを抜ける
                    break
                else:
                    time.sleep(1)
                break
            except Exception as e:
                if element is None:
                    time.sleep(1)
                    continue
                self.app.log_add("職種カテゴリ検索中に不明なエラーが見つかりました。")
                self.handle_job_category_not_found()
                return "NG"
        else:
            self.app.log_add("職種カテゴリ候補が5つ見つかりませんでした。")
            #self.handle_job_category_not_found()
            #return "NG"
            
        try:
            # 指定された範囲のtr要素をループ処理
            for i in range(1, 5):  # nth-childは1から始まるため、1から6まで
                # 現在のtr要素のXPath
                tr_xpath = f"{base_xpath}:nth-child({i})"
                # tr要素内のclass="main"を持つすべての要素を取得。
                main_elements = WebDriverWait(bm.driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, f"{tr_xpath} .main")))
                # main_elementsに対応する要素の数を数える
                #main_elements_count = len(main_elements)
                
                # mainクラスを持つ要素の中でテキストを検索
                for element in main_elements:
                    if search_text in element.text:
                        # 次のaタグを探す
                        a_tag_xpath = f"//*[@id='setAiJobSuggestTarget']/tr[{i}]/td[2]/a"
                        a_tag_element = bm.driver.find_element(By.XPATH, a_tag_xpath)
                        a_tag_element.click()

                        print(f"クリックされた要素: {i}行目")
                        return "OK"  # 見つかったらループを抜ける
            else:
                #見つからなければエラー処理へ
                self.handle_job_category_not_found()
                return "NG"
        except Exception as e:
            self.app.log_add(f"職業カテゴリ検索中に不明なエラーが発生しました: {e}")
            self.handle_job_category_not_found()
            return "NG"
        
    def handle_job_category_not_found(self):
        """
        engage_settings.xlsxの「職種カテゴリが見つからない場合」の値の入ったクラス変数
        app.setting_job_category_not_foundの値に基づいて処理を分岐させる関数。
        """
        setting_value = self.app.setting_job_category_not_found

        # "何もしない"の場合
        if setting_value == "何もしない":
            # 閉じるボタンをクリックし、関数を終了
            try:
                close_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'body > div.md_modal.md_modal--suggest.md_modal--show > div > div.close > a'))
                )
                close_button.click()
                self.app.log_add("閉じるボタンをクリックしました。")
            except Exception as e:
                self.app.log_add(f"閉じるボタンをクリックできませんでした: {e}")
                return "閉じるボタンをクリックできませんでした。"
            return

        # "エラーを表示して処理を中断"の場合
        elif setting_value == "エラーを表示して処理を中断":
            # エラーメッセージを表示し、処理を中断
            root = tk.Tk()
            root.attributes("-topmost", True)
            root.withdraw()
            messagebox.showinfo("ERROR","職種カテゴリが見つかりませんでした。手動でクリックしてください", parent=root)
            self.app.log_add("職種カテゴリが見つかりませんでした。手動クリックを実行")
            # 閉じるボタンをクリックし、関数を終了
            try:
                close_button = WebDriverWait(self.driver, 1).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'body > div.md_modal.md_modal--suggest.md_modal--show > div > div.close > a'))
                )
                close_button.click()
                self.app.log_add("閉じるボタンをクリックしました。")
            except Exception as e:
                #self.app.log_add(f"閉じるボタンをクリックできませんでした: {e}")
                return "閉じるボタンをクリックできませんでした。"
            return "職種カテゴリが見つかりませんでした。"

        # "一番上のカテゴリをクリック"の場合
        elif setting_value == "一番上のカテゴリをクリック":
            try:
                first_category_element = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "#setAiJobSuggestTarget > tr:nth-child(1) > td.data.data--min > a"))
                )
                first_category_element.click()
                self.app.log_add("一番上の職種カテゴリをクリックしました。")
            except Exception as e:
                self.app.log_add(f"一番上の職種カテゴリをクリックできませんでした: {e}")
                return "一番上の職種カテゴリをクリックできませんでした。"

        # "その他"の場合
        else:
            # その他の処理を行う
            self.app.log_add("職種カテゴリが見つかりませんでした。その他の処理を行います。")
            # ここにその他の処理を記述
            # 例: self.other_process()
        """
        engage_settings.xlsxの「職種カテゴリが見つからない場合」の値の入ったクラス変数
        app.setting_job_category_not_foundの値に基づいて処理を分岐させる関数。
        """
        setting_value = self.app.setting_job_category_not_found

    def set_value_with_js(self,driver, element, value):
        # Send_keysの値を設定するJavaScriptコード
        set_value_script = """
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """
        # スクリプトを実行して値を設定し、イベントをトリガーする
        driver.execute_script(set_value_script, element, value)
        
    def hover_over_element(self, xpath,img_no,filename):
        """
        指定されたXPathの要素にマウスオーバーし、選択ボタンをクリック、アップロード済みの画像を選択する処理
        :param xpath: マウスオーバーする要素のXPath
        
        """

        #クリック出来るまで30回繰り返す
        for _ in range(30):
            try:
                #画像ｘ枚目のXPATHを取得
                element_to_hover = WebDriverWait(self.driver, 1).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                #画像の表示位置までスクロール
                self.driver.execute_script("arguments[0].scrollIntoView(true);", element_to_hover)
                #該当画像に疑似的にマウスオーバーうまくいかない場合はJavascriptでホバー処理
                try:
                    ActionChains(self.driver).move_to_element(element_to_hover).perform()
                except Exception as e:
                    hover_script = """
                    var event = new MouseEvent('mouseover', {
                        view: window,
                        bubbles: true,
                        cancelable: true
                    });
                    arguments[0].dispatchEvent(event);
                    """
                    self.driver.execute_script(hover_script, element_to_hover)
                #
                #マウスオーバー後に、ファイルを選択をクリック
                self.driver.find_element(By.CSS_SELECTOR, f"a:nth-child({img_no}) .editBtn").click()
                break #クリックが成功すれば処理を抜ける
            except Exception as e:
                print(f"追加するボタンをクリックできませんでした。: {e}")
                if _ == 29:
                    print("30回の試行が完了しましたが、要素の操作が成功しませんでした。")
                    self.app.log_add(f"{img_no}枚目の追加ボタンをクリックできませんでした")
                    return False
                time.sleep(1)
                continue

        for _ in range(2):
            # クラ スが"md_card js_workPictureItem"である全ての要素を取得
            try:
                elements = self.driver.find_elements(By.CSS_SELECTOR, ".fileInput.photo.js_editImage.companyImage")
                print("画像一覧の取得が完了")
                break #画像一覧を取得できれば処理を抜ける。
            except Exception as e:
                print(f"画像一覧の取得中の検索中にエラーが発生しました: {e}")
                time.sleep(1)
                if _ == 1:
                    print("2回の試行が完了しましたが、画像一覧の取得が成功しませんでした。")
                    self.app.log_add("画像一覧の取得エラー")
                    return False
                time.sleep(1)
                continue

        #追加ボタンクリック後、画像選択処理
        
        for index, element in enumerate(elements):
            # 各要素のStyle属性の値を取得
            style_attribute_value = element.get_attribute("style")
            # 特定の文字列が含まれているかをチェックし、条件に合致する場合の処理を行う
            if filename in style_attribute_value:
                print(style_attribute_value)
                # 特定の文字列が含まれている場合の処理を記述。安定化のために30回ループ
                for attempt in range(30):
                    try:
                        # スクロールして要素をビューポートに表示
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
                        # マウスオーバーを実行
                        ActionChains(self.driver).move_to_element(element).perform()
                        # マウスオーバーによって表示される要素を待機して検索
                        clickable_element = self.driver.find_element(By.CSS_SELECTOR, f"#workPictureItem{index} .editBtn")
                        clickable_element.click()
                        break  # クリックが成功したらループを抜ける
                    except Exception as e:
                        logger.error(f"クリック可能な要素が見つからないか、クリックできませんでした: {e}")
                        if attempt == 29:
                            logger.error("30回の試行が完了しましたが、クリック可能な要素の操作が成功しませんでした。")
                            self.app.log_add(f"要素のクリック試行が30回に達しました: {filename}")
                            # 閉じるボタンをクリックしてダイアログを閉じる
                            try:
                                alternate_element = self.driver.find_element(By.XPATH, "/html/body/div[9]/div/div[3]")
                                alternate_element.click()
                            except NoSuchElementException:
                                self.app.log_add("エラー: 閉じるボタンが見つかりませんでした。")
                        #time.sleep(1)  # 次の試行まで少し待機
                break
        else:                                        # 閉じるボタンをクリックしてダイアログを閉じる
            try:
                alternate_element = self.driver.find_element(By.XPATH, "/html/body/div[9]/div/div[3]")
                alternate_element.click()
                print("画像が見つかりませんでした")
            except NoSuchElementException:
                self.app.log_add("エラー: 閉じるボタンが見つかりませんでした。")

    def get_joblistPage_element(self, rows, df_download_data, bm, new_index, x_row, td_index):
        """
        求人一覧ページのWeb要素を取得メソッド。
        :param web_element: 処理するWeb要素
        :param df_download_data: データを保存するDataFrame
        :param bm: ブラウザマネージャー
        :param new_index: DataFrameに新しい行を追加するためのインデックス
        :param x_row: 行インデックス
        :param td_index: セルインデックス
        """

        # 原稿一覧の取得処理開始
        try:

            if len(rows) > x_row:
                row_element = rows[x_row]
                tds = row_element.find_elements(By.TAG_NAME, "td")
                # 原稿のエラーを取得する。
                if 'md_errorRow' in row_element.get_attribute('class').split():
                    df_download_data.loc[new_index, "原稿エラー"] = True
                #表示用職種名取得処理
                try:
                    job_title = tds[0].text
                    display_job_title = job_title.split('\n')[0].strip()
                    df_download_data.loc[new_index, "表示用職種名"] = display_job_title if display_job_title else ""
                except NoSuchElementException:
                    self.app.log_add("表示用職種名の要素が見つかりませんでした")
                    df_download_data.loc[new_index, "表示用職種名"] = None


                # 求人ID取得処理
                if len(tds) > 0:
                    td_value = tds[0]  # 1番目のセルの値を取得
                    try:
                        job_no = td_value.find_element(By.CLASS_NAME, "setData").text
                        df_download_data.loc[new_index, "求人ID"] = str(job_no) if job_no else ""
                        # 表示用職種の取得処理
                        new_index = len(df_download_data)
                    except Exception as e:
                        self.app.log_add(f"求人IDが存在しません")
                        df_download_data.loc[new_index, "求人ID"] = None
                else:
                    print(f"求人IDのセルインデックス {td_index} は範囲外です")

                # 掲載状況の取得処理
                if len(tds) > td_index:
                    if tds[td_index].find_elements(By.TAG_NAME, "select"):
                        select_element = tds[td_index].find_element(By.TAG_NAME, "select")
                        selected_option = select_element.find_element(By.CSS_SELECTOR, "option:checked").text
                        td_value = selected_option.strip()
                    else:
                        td_value = tds[td_index].text.strip()
                        if "求人の修正" in td_value:
                            td_value = "".join(td_value.split())    # 求人の修正という文字が含まれている場合は空白のみ除去
                        else:
                            td_value = td_value.split('\n')[0].strip()  # 1行目を取得し、空白を削除

                    # 掲載状況に代入   
                    df_download_data.loc[new_index, "掲載状況"] = str(td_value) if td_value else ""
                    new_index = len(df_download_data)
                else:
                    print(f"公開状況のセルインデックス {td_index} は範囲外です")
            else:
                print(f"行インデックス {x_row} は範囲外です")

        except Exception as e:
            print(f"エラーが発生しました: {e}")
            return None

        return new_index

    def extract_work_id(self, web_element, df_download_data, new_index):
        """
        Web要素からwork_idを抽出し、DataFrameに保存するメソッド。
        :param web_element: 処理するWeb要素
        :param df_download_data: データを保存するDataFrame
        :param new_index: DataFrameに新しい行を追加するためのインデックス
        """
        try:
            # リンクURLからwork_idの値を抽出
            link_url = web_element.get_attribute('href')
            work_id = None
            if 'work_id=' in link_url:
                work_id = link_url.split('work_id=')[1].split('&')[0]
                # 原稿更新URLを作成
                df_download_data.loc[new_index, "原稿更新URL"] = f"https://en-gage.net/company/job/renew/form/?work_id={work_id}"
            # work_idをDFに保存
            df_download_data.loc[new_index, "work_id"] = work_id if work_id else None
            return True
        except Exception as e:
            self.app.log_add(f"work_idの抽出中にエラーが発生しました: {e}")
            return None

    def search_work_id(self, web_element, df_download_data, new_index):
        """
        Web要素からwork_idを抽出し、DataFrameに保存するメソッド。
        :param web_element: 処理するWeb要素
        :param df_download_data: データを保存するDataFrame
        :param new_index: DataFrameに新しい行を追加するためのインデックス
        """
        try:
            # リンクURLからwork_idの値を抽出
            link_url = web_element.get_attribute('href')
            work_id = None
            if 'work_id=' in link_url:
                work_id = link_url.split('work_id=')[1].split('&')[0]
                # 原稿更新URLを作成
                df_download_data.loc[new_index, "原稿更新URL"] = f"https://en-gage.net/company/job/renew/form/?work_id={work_id}"
            # work_idをDFに保存
            df_download_data.loc[new_index, "work_id"] = work_id if work_id else None
            return True
        except Exception as e:
            self.app.log_add(f"work_idの抽出中にエラーが発生しました: {e}")
            return None
        
    def get_individual_job(self, web_element, df_download_data, new_index, bm, joblist_page_no, web_elements_cnt):
        """
        個別原稿の取得処理を行うメソッド。
        :param web_element: 処理するWeb要素
        :param df_download_data: データを保存するDataFrame
        :param new_index: DataFrameに新しい行を追加するためのインデックス
        :param bm: ブラウザマネージャーのインスタンス
        :param joblist_page_no: 原稿一覧のページ番号
        :param web_elements_cnt: 現在処理中の原稿Noカウント用
        """
        # 要素がクリック可能かチェック
        if bm.is_element_clickable(bm.driver, web_element):
            try:
                # 要素の位置までスクロール
                bm.driver.execute_script("arguments[0].scrollIntoView(true);", web_element)
                # 要素をJavascriptでクリック
                bm.driver.execute_script("arguments[0].click();", web_element)
                # self.app.log_add(f"{web_elements_cnt + 1}つ目の要素をクリックしました。")
            except Exception as e:
                self.app.log_add(f"{web_elements_cnt + 1}つ目の要素がクリックできませんでした: {e}")
        else:
            self.app.log_add(f"{web_elements_cnt + 1}番目の原稿がクリックできません。")

        # 雇用形態取得
        employment_type = bm.get_selected_radio_button_label("employment_status", 'name')
        df_download_data.loc[new_index, "雇用形態"] = str(employment_type)
        # 雇用形態の値が派遣だった場合、専用の処理を追加
        if employment_type == "派遣":
            # 派遣:選択時のセレクトボックス取得
            df_download_data.loc[new_index, "派遣:選択"] = bm.get_selected_option_text("select_agency_staff", "id")
            # 仕事No.取得
            df_download_data.loc[new_index, "仕事No."] = bm.get_input_label_and_value("work_manage_no", 'name')

        # 全てのチェックボックス取得処理
        checked_checkboxes = bm.get_checked_checkboxes("feature_id", 'name')

        # チェックボックスの値を設定する前にブール型にキャスト
        for checkbox in checked_checkboxes:
            df_download_data.loc[new_index, checkbox] = True

        # 各種情報を取得してDataFrameに保存
        fields = [
            ("official_occupation_name", "職種"),
            ("occupation_name", "表示用職種名"),
            ("work_contents", "仕事内容"),
            ("categoryInput", "職種カテゴリー", 'id'),
            ("official_corporate_name", "法人名（正式社名）"),
            ("business_content", "事業内容"),
            ("work_office_name[0]", "勤務先名"),
            ("work_office_division[0]", "勤務先区分", 'name', bm.get_selected_radio_button_label),
            ("work_office_zip_code[0]", "郵便番号"),
            ("work_office[0]", "都道府県", 'name', bm.get_selected_option_text),
            ("municipalities[0]", "市区町村"),
            ("other_address[0]", "以降の住所"),
            ("work_office_station[0]", "最寄り駅"),
            ("work_location", "勤務地：備考"),
            ("access", "アクセス"),
            ("work_division", "求人区分", 'name', bm.get_selected_radio_button_label),
            ("salary_type_selected", "給与タイプ", 'name', bm.get_selected_radio_button_label),
            ("educational_status", "最終学歴", 'name', bm.get_selected_option_text),
            ("occupation_experience", "募集職種の経験有無", 'name', bm.get_selected_radio_button_label),
            ("qualification", "その他必要な経験・資格など"),
            ("recruitment_bg", "募集人数・募集背景"),
            ("holiday_system", "休みの取り方", 'name', bm.get_selected_option_text),
            ("holiday", "休日休暇"),
            ("treatment", "待遇・福利厚生")
        ]

        for field in fields:
            if len(field) == 2:
                value = bm.get_input_label_and_value(field[0], 'name')
            elif len(field) == 3:
                value = bm.get_element_text(field[0], field[2])
            else:
                value = field[3](field[0], field[2])
            df_download_data.loc[new_index, field[1]] = value if value != '' else None

        # ステータス
        df_download_data.loc[new_index, "ステータス"] = "入力内容を保存"

        # 給与（最低額）と給与（最高額）
        salary_types = [
            ('salary_amount_from_5_5', 'salary_amount_to_5_5', "年収の値が空です"),
            ('salary_amount_from_1_2', 'salary_amount_to_1_2', "月給の値が空です"),
            ('salary_amount_from_3_3', 'salary_amount_to_3_3', "日給の値が空です"),
            ('salary_amount_from_4_4', 'salary_amount_to_4_4', "時給の値が空です")
        ]

        for salary_from, salary_to, error_message in salary_types:
            try:
                salary_min = bm.get_input_label_and_value(salary_from, 'name')
                salary_max = bm.get_input_label_and_value(salary_to, 'name')
                if not salary_min:
                    raise ValueError(error_message)
                df_download_data.loc[new_index, "給与（最低額）"] = salary_min
                df_download_data.loc[new_index, "給与（最高額）"] = salary_max
                break
            except (NoSuchElementException, ValueError):
                continue
        else:
            # すべての試行が失敗した場合の処理
            df_download_data.loc[new_index, "給与（最低額）"] = None
            df_download_data.loc[new_index, "給与（最高額）"] = None

        if (df_download_data["給与タイプ"] == '年俸').any():
            try:
                value = bm.get_selected_radio_button_label('payment_method', 'name')
                df_download_data.loc[new_index, "支払方法"] = value if value != '' else None
            except Exception as e:
                df_download_data.loc[new_index, "支払方法"] = None  # デフォルト値を設定

        if (df_download_data["給与タイプ"] == '月給').any():
            value = bm.get_input_label_and_value('salary_amount_from_1_1', 'name')
            df_download_data.loc[new_index, "想定年収（最低額）"] = value if value != '' else None
            value = bm.get_input_label_and_value('salary_amount_to_1_1', 'name')
            df_download_data.loc[new_index, "想定年収（最高額）"] = value if value != '' else None

        # 年収例の処理を挿入
        salary_examples = [
            ('annual_salary_1_amount', '年収例_1', 'annual_salary_1_years', '入社歴_1', 'annual_salary_1_remark', '役職例_1'),
            ('annual_salary_2_amount', '年収例_2', 'annual_salary_2_years', '入社歴_2', 'annual_salary_2_remark', '役職例_2'),
            ('annual_salary_3_amount', '年収例_3', 'annual_salary_3_years', '入社歴_3', 'annual_salary_3_remark', '役職例_3')
        ]

        for example in salary_examples:
            value = bm.get_input_label_and_value(example[0], 'name')
            df_download_data.loc[new_index, example[1]] = value if value != '' else None
            value = bm.get_selected_option_text(example[2], 'name')
            df_download_data.loc[new_index, example[3]] = value if value != '' else None
            value = bm.get_input_label_and_value(example[4], 'name')
            df_download_data.loc[new_index, example[5]] = value if value != '' else None

        # 完全成果報酬：金額
        if (df_download_data["給与タイプ"] == '完全成果報酬').any():
            value = bm.get_input_label_and_value('salary_note', 'name')
            df_download_data.loc[new_index, "完全成果報酬：金額"] = value if value != '' else None

        # みなし残業代
        value = bm.get_input_label_and_value('overtime_fee', 'name')
        if not isinstance(value, str):
            value = None
        else:
            try:
                value = int(value)
            except ValueError:
                value = None
        df_download_data.loc[new_index, "みなし残業代"] = value if value != '' else None

        # 給与：備考の値を取得
        try:
            salary_note = bm.get_input_label_and_value('salary_note', 'name')
            df_download_data.loc[new_index, "給与：備考"] = salary_note
        except Exception:
            salary_note = None

        # 勤務時間
        value = bm.get_selected_radio_button_label('office_hour_style', 'name')
        df_download_data.loc[new_index, "勤務時間"] = value if value != '' else None

        # 勤務時間：備考
        value = bm.get_input_label_and_value('office_hours', 'name')
        df_download_data.loc[new_index, "勤務時間：備考"] = value if value != '' else None

        if (df_download_data["勤務時間"] == "固定時間制").any():
            # 想定勤務：開始時間、終了時間、終了分
            value = bm.get_selected_option_text('office_hour_from_hour', 'name')
            df_download_data.loc[new_index, "想定勤務：開始時間"] = value if value != '' else None
            value = bm.get_selected_option_text('office_hour_from_minute', 'name')
            df_download_data.loc[new_index, "想定勤務：開始分"] = value if value != '' else None
            value = bm.get_selected_option_text('office_hour_to_hour', 'name')
            df_download_data.loc[new_index, "想定勤務：終了時間"] = value if value != '' else None
            value = bm.get_selected_option_text('office_hour_to_minute', 'name')
            df_download_data.loc[new_index, "想定勤務：終了分"] = value if value != '' else None

        # 画像ファイルの一覧を取得
        image_paths = bm.log_image_paths('class', 'js_editImage', 5)
        for i, path in enumerate(image_paths):
            if path:
                # '/'の後の部分を抽出
                extracted_path = path[path.rfind('/') + 1:]
                df_download_data.loc[new_index, f"画像ファイル_{i+1}"] = extracted_path

                # 動的に名前属性を生成
                name_attribute = f'work_picture_text[{i+1}]'

                try:
                    # 要素を取得
                    value = bm.get_input_label_and_value(name_attribute, 'name')
                    df_download_data.loc[new_index, f"画像コメント_{i+1}"] = value if value != '' else None
                except Exception as e:
                    df_download_data.loc[new_index, f"画像コメント_{i+1}"] = None  # デフォルト値を設定
                    self.app.log_add(f"画像コメント_{i+1} の要素が見つかりませんでした: {e}")

        # 現在取得している原稿詳細の取得日時をDownload用のDFに代入
        df_download_data.loc[new_index, "page番号"] = joblist_page_no
        df_download_data.loc[new_index, "_取得日時"] = datetime.now().strftime('%Y/%m/%d %H:%M')
        df_download_data.loc[new_index, "原稿行"] = web_elements_cnt

        return df_download_data

def engage_auth_check(app, PROGRAM_VER, filename='account.xlsx', e=None):
    # account.xlsxの存在チェック、およびアカウントチェック
    if not filename:
        filename = 'account.xlsx'


    # ファイルの存在をチェック
    if not os.path.exists(filename):
        app.log_add(f"{filename} が存在しません。")
        return None, None, None,None,None,None

    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    app.log_add(f"Ver.{PROGRAM_VER} - {current_datetime}")
    
    app.log_add(f"{filename} の存在チェック…ＯＫ")

    app.log_add("Id,PASSをロード中…")
    # Excelファイルから読み込む場合、engage_id2からengage_id10まで
    engage_id_manager_excel = EngageIdManager(filename, 'excel', 1, 10)
    engage_id_manager_excel.load_data()

    # engage_id_manager_excelのデータから1行目のengage_idを取得
    engage_id = engage_id_manager_excel.data.get('engage_id1', [None])[0]

    if engage_id is not None:
        app.log_add(f"1行目のengage_id: {engage_id}")
        app.log_add("ロードが完了しました。")

    else:
        app.log_add("1行目のengage_idが見つかりません。")


    engage_id = ""
    engage_pass = ""
    company_name = ""
    engage_id_list = deque()
    engage_pass_list = deque()
    company_name_list = deque()

    engage_id = engage_id_manager_excel.data.get('engage_id1', [None])[0]
    engage_pass = engage_id_manager_excel.data.get('engage_pass1', [None])[0]
    company_name = engage_id_manager_excel.data.get('company_name1', [None])[0]

    engage_id_list = deque(engage_id_manager_excel.data.get('engage_id1', [None]))
    engage_pass_list = deque(engage_id_manager_excel.data.get('engage_pass1', [None]))
    company_name_list = deque(engage_id_manager_excel.data.get('company_name1', [None]))
    
    return engage_id, engage_pass, company_name, engage_id_list, engage_pass_list, company_name_list


def process_salary_types(self,bm, df_download_data, app):
    salary_types = [
        ('salary_amount_from_5_5', 'salary_amount_to_5_5', "年収の値が空です"),
        ('salary_amount_from_1_2', 'salary_amount_to_1_2', "月給の値が空です"),
        ('salary_amount_from_3_3', 'salary_amount_to_3_3', "日給の値が空です"),
        ('salary_amount_from_4_4', 'salary_amount_to_4_4', "時給の値が空です")
    ]

    for salary_from, salary_to, error_message in salary_types:
        try:
            salary_min = bm.get_input_label_and_value(salary_from, 'name')
            salary_max = bm.get_input_label_and_value(salary_to, 'name')
            if not salary_min:
                raise ValueError(error_message)
            df_download_data["給与（最低額）"] = salary_min
            df_download_data["給与（最高額）"] = salary_max
            break
        except (NoSuchElementException, ValueError):
            continue
    else:
        # すべての試行が失敗した場合の処理
        df_download_data["給与（最低額）"] = "未設定"
        df_download_data["給与（最高額）"] = "未設定"
        app.log_add("給与の値がすべて空です")
    def element_exists(self, xpath):
        try:
            self.driver.find_element(By.XPATH, xpath)
            return True
        except NoSuchElementException:
            return False

