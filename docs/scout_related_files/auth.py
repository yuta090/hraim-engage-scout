import os
import pandas as pd
import json
import tkinter as tk
from tkinter import messagebox

class EngageIdManager:
    def __init__(self, filepath, filetype='excel', start_id=2, end_id=100, logger=None):
        """
        EngageIdManager クラスのコンストラクタ。

        :param filepath: 読み込むファイルのパス。
        :param filetype: 読み込むファイルのタイプ（'excel' または 'json'）。
        :param start_id: 読み込むengage_idの開始番号。
        :param end_id: 読み込むengage_idの終了番号。
        :param logger: ログ出力用のLoggerインスタンス（オプショナル）。
        """
        self.filepath = filepath
        self.filetype = filetype
        self.start_id = start_id
        self.end_id = end_id
        self.logger = logger
        self.data = {}

    def load_data(self):
        """ファイルタイプに応じてデータを読み込む。"""
        if self.filetype == 'excel':
            self._load_from_excel()
        elif self.filetype == 'json':
            self._load_from_json()
        else:
            raise ValueError("Unsupported filetype provided.")

    def _load_from_excel(self):
        """Excelファイルからデータを読み込む。"""
        if not os.path.exists(self.filepath):
            self._show_error(f"{self.filepath}が見つかりません。処理を終了します。")
            raise SystemExit

        df = pd.read_excel(self.filepath)
        self._process_dataframe(df)

    def _load_from_json(self):
        """JSONファイルからデータを読み込む。"""
        if not os.path.exists(self.filepath):
            self._show_error(f"{self.filepath}が見つかりません。")
            raise SystemExit

        with open(self.filepath, 'r') as file:
            data = json.load(file)
            self._process_json(data)

    def _process_dataframe(self, df):
        """DataFrameから指定された範囲のengage_id、engage_pass、およびcompany_nameを読み込む。"""
        for i in range(self.start_id, self.end_id + 1):
            id_column = f'engage_id{i}'
            pass_column = f'engage_pass{i}'
            # company_name列の取得を追加
            company_name_column = f'company_name{i}'
            if id_column in df.columns and pass_column in df.columns and company_name_column in df.columns:
                self.data[id_column] = df[id_column].dropna().tolist()
                self.data[pass_column] = df[pass_column].dropna().tolist()
                # company_name列のデータを取得
                self.data[company_name_column] = df[company_name_column].dropna().tolist()

    def _process_json(self, data):
        """JSONから指定された範囲のengage_idとengage_passを読み込む。"""
        for i in range(self.start_id, self.end_id + 1):
            id_key = f'engage_id{i}'
            pass_key = f'engage_pass{i}'
            if id_key in data and pass_key in data:
                self.data[id_key] = data[id_key]
                self.data[pass_key] = data[pass_key]

    def _show_error(self, message):
        """エラーダイアログを表示する内部メソッド。"""
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("エラー", message)
        root.destroy()
