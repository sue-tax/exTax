'''
色々なパターンに対応可能とする
Created on 2026/11/07

@author: sue-t
'''


import json5
import openpyxl as xl
import csv
import os
# import jaconv
from decimal import Decimal, ROUND_UP
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

import c
import d
import e

class ExTax(object):
    def __init__(self, config_data, excel_file_name, csv_folder_name, kaipage_yokusei):
        self.config_data = config_data
        self.book = xl.load_workbook(excel_file_name, data_only=True)
        self.csv_folder_name = csv_folder_name
        # 改ページ抑制の有無　抑制する場合は合計行を入れない
        self.kaipage_yokusei = kaipage_yokusei
        
        self.format_kubun_kouban = config_data["フォーマット区分項番"]
        self.gyou_kubun_kouban = config_data["行区分項番"]
    
    
    def execute(self):
        for data in self.config_data["勘定科目内訳明細書ファイル"]:
            d.dprint(data["科目名"])
            if len(data['入力シート']) == 1:
                # 単独シートの処理
                sheet_data = data['入力シート'][0]
                d.dprint(" "+sheet_data["シート名"])
            
                if sheet_data["シート名"] \
                        not in self.book.sheetnames:
                    d.dprint("{}シートがないので飛ばす".format(sheet_data["シート名"]))
                    continue
                sheet = self.book[sheet_data["シート名"]]
                csv_file_name = os.path.join(self.csv_folder_name,
                        data["出力CSVファイル"])
                d.dprint(" "+csv_file_name)
                self.sheet_to_csv(sheet_data, sheet, csv_file_name)

            else:
                # 複数シートの処理
                for sheet_data in data['入力シート']:
                    d.dprint(" "+sheet_data["シート名"])
                    if sheet_data["シート名"] \
                            not in self.book.sheetnames:
                        d.dprint("{}シートがないので飛ばす".format(sheet_data["シート名"]))
                        break
                else:
                    # シートが全部あった
                    csv_file_name = os.path.join(self.csv_folder_name,
                            data["出力CSVファイル"])
                    d.dprint(" "+csv_file_name)
                    yakuin_kyuyo = data.get("役員給与", None)
                    if yakuin_kyuyo == None:
                        self.sheet_fukusu_to_csv(data["入力シート"], csv_file_name)
                    else:
                        self.sheet_yakuin_to_csv(data["入力シート"], csv_file_name)

                

    def sheet_to_csv(self, sheet_data, sheet, csv_file_name):
        # d.dprint_method_start()
        csv_data = []
        self.sheet_to_data(sheet_data, sheet, csv_data)
        # d.dprint(csv_data)
        if len(csv_data) != 0:
            self.write_csv_file(csv_file_name, csv_data)
        else:
            d.dprint("データがないので、"+csv_file_name+"は作成しない")
        # d.dprint_method_end()


    def sheet_fukusu_to_csv(self, sheet_datas, csv_file_name):
        # d.dprint_method_start()
        
        csv_data = []
        for sheet_data in sheet_datas:
            sheet = self.book[sheet_data["シート名"]]
            self.sheet_to_data(sheet_data, sheet, csv_data)
        # d.dprint(csv_data)
        if len(csv_data) != 0:
            self.write_csv_file(csv_file_name, csv_data)
        else:
            d.dprint("データがないので、"+csv_file_name+"は作成しない")
        # d.dprint_method_end()


    def sheet_yakuin_to_csv(self, sheet_datas, csv_file_name):
        # d.dprint_method_start()
        csv_data = []
        sheet_data = sheet_datas[0]
        sheet = self.book[sheet_data["シート名"]]
        goukei_row_daihyou = self.sheet_to_data(sheet_data, sheet, csv_data)
        sheet_data = sheet_datas[1]
        sheet = self.book[sheet_data["シート名"]]
        goukei_row_igai = self.sheet_to_data(sheet_data, sheet, csv_data)
        
        csv_data.pop(-1)
        row_data = []
        row_data.append(goukei_row_igai[0])
        goukei_kous = sheet_datas[0].get("合計項", None)
        row_data.append("1")
        for kou in sheet_data["項"]:
            column = kou["項番"]
            if column in goukei_kous:
                # d.dprint(column)
                if goukei_row_daihyou[column-1] == '':
                    goukei_row_daihyou[column-1] = 0 
                if goukei_row_igai[column-1] == '':
                    goukei_row_igai[column-1] = 0 
                goukei = int(goukei_row_daihyou[column-1]) + int(goukei_row_igai[column-1])
            else:
                goukei = ""
            row_data.append(goukei)           
        csv_data.append(row_data)

        sheet_data = sheet_datas[2]
        sheet = self.book[sheet_data["シート名"]]
        self.sheet_to_data(sheet_data, sheet, csv_data)
        
        # d.dprint(csv_data)
        if len(csv_data) != 0:
            self.write_csv_file(csv_file_name, csv_data)
        else:
            d.dprint("データがないので、"+csv_file_name+"は作成しない")
        # d.dprint_method_end()


    def sheet_to_data(self, sheet_data, sheet, csv_data):
        # d.dprint_method_start()
        # d.dprint(sheet)
        goukei_kous = sheet_data.get("合計項", None)
        if goukei_kous != None:
            goukei_dict = {}
            for kou in goukei_kous:
                goukei_dict[kou] = 0

        format_kubun = ""
        flag_empty = True
        row_data = []
        for row in range(sheet_data["タイトル行数"] + 1, sheet.max_row + 1):
            if sheet.cell(row=row, column=self.gyou_kubun_kouban).value \
                    == '1':
                if sheet_data["合計行"] == 0:
                    # 合計行を認めない内訳書なので、合計行を飛ばす
                    d.dprint("合計行を認めないので、合計行を飛ばす")
                    continue
                if sheet_data["改ページ"] == 1:
                    if self.kaipage_yokusei:
                        d.dprint("改ページ抑制のために、合計行を飛ばす")
                        continue
            if sheet.cell(row=row, column=1).value == None:
                d.dprint("空の行を飛ばす")
                continue
            row_data = []
            
            # d.dprint(sheet.cell(row=row, column=1).value)
            format_kubun = sheet.cell(row=row, column=1).value
            # d.dprint(data)
            if format_kubun is None:
                format_kubun = ""
            # d.dprint(data)
            row_data.append(format_kubun)
            # d.dprint(sheet.cell(row=row, column=2).value)
            gyou_kubun = sheet.cell(row=row, column=2).value
            # d.dprint(data)
            if gyou_kubun is None:
                gyou_kubun = ""
            # d.dprint(data)
            row_data.append(gyou_kubun)
            
            # for column in range(1, len(sheet_data["項"])+2 + 1):
            for kou in sheet_data["項"]:
                column = kou["項番"]
                # d.dprint(column)
                # d.dprint(sheet.cell(row=row, column=column).value)
                data = sheet.cell(row=row, column=column).value
                # d.dprint(data)
                
                kijun = kou["入力基準"]
                if data is not None:
                    # 先に文字列型に統一。ただし、floatの「.0」が文字数や桁数判定に悪影響を与えないよう考慮
                    if isinstance(data, (int, float)):
                        # 有価証券や利率など、意図しない指数表記（1e-05など）を防ぐためフォーマット
                        num_str = f"{data}".rstrip('0').rstrip('.') if '.' in f"{data}" else f"{data}"
                    else:
                        num_str = str(data).strip()

                    match kijun["種類"] :
                        case "文字":
                            try:
                                # num_str = jaconv.h2z(num_str)
                                import jaconv
                                num_str = jaconv.h2z(num_str, kana=True, ascii=True, digit=True)
                            except ImportError:
                                pass
                            except Exception as _exception:
                                errMsg = "シート:{},行:{},列:{},データ:{}". \
                                    format(sheet, row, column, sheet.cell(row=row, column=column).value)
                                d.dprint(errMsg)
                                e.eprint("データ変換不能", errMsg)
                                num_str = "データ変換不能"
                            
                            # 最新仕様に基づき、改行コード(\n)はスペース置換ではなくそのまま保持して後にダブルクォーテーションで囲む
                            #（もし仕様上スペース置換で統一したい場合は元の replace('\n', ' ') でも可）
                            if "\n" in num_str:
                                if not (num_str.startswith('"') and num_str.endswith('"')):
                                    num_str = f'"{num_str}"'
                            data = num_str[:kijun["文字数"]]
                    
                        case "半角":
                            try:
                                import jaconv
                                # z2h で全角の英数字・カタカナ・記号をすべて半角に一発変換
                                # kana=True, ascii=True, digit=True を明示して確実に半角化します
                                num_str = jaconv.z2h(num_str, kana=True, ascii=True, digit=True)
                            except ImportError:
                                # jaconvがインストールされていない場合のフォールバック（最低限文字列化）
                                pass
                            except Exception as _exception:
                                # 💡 変換不能時のエラーログ出力処理
                                errMsg = "シート:{},行:{},列:{},データ:{}". \
                                    format(sheet, row, column, sheet.cell(row=row, column=column).value)
                                d.dprint(errMsg)
                                e.eprint("データ半角変換不能", errMsg)
                                num_str = "データ変換不能"
                            
                            # 最新仕様に基づき、文字列の改行コード(\n)はそのまま保持して後にダブルクォーテーションで囲む
                            if "\n" in num_str:
                                if not (num_str.startswith('"') and num_str.endswith('"')):
                                    num_str = f'"{num_str}"'
                                    
                            # 指定された文字数で切り捨て
                            data = num_str[:kijun["文字数"]]
                    
                        case "整数":
                            # 文字列としての長さではなく、数値としての正確な桁数を判定
                            # clean_num_str = num_str.replace("-", "").split(".")[0] # マイナス記号と小数点以下を除外
                            clean_num_str = num_str.split(".")[0] # マイナス記号と小数点以下を除外
                            # マイナスのデータの処理方法は未確認
                            if len(clean_num_str) > kijun["桁数"]:
                                errMsg = "シート:{},行:{},列:{},データ:{}". \
                                    format(sheet, row, column, sheet.cell(row=row, column=column).value)
                                d.dprint(errMsg)
                                e.eprint("桁数オーバー", errMsg)
                                data = "桁数オーバー"
                            else:
                                data = clean_num_str
                                if goukei_kous != None:
                                    if column in goukei_kous:
                                        goukei = goukei_dict[column]
                                        goukei += int(clean_num_str)
                                        goukei_dict[column]  = goukei

                        case "小数":
                            # if len(num_str.replace("-", "").replace(".", "")) > kijun["桁数"]:
                            if len(num_str.replace(".", "")) > kijun["桁数"]:
                                errMsg = "シート:{},行:{},列:{},データ:{}". \
                                    format(sheet, row, column, sheet.cell(row=row, column=column).value)
                                d.dprint(errMsg)
                                e.eprint("桁数オーバー", errMsg)
                                data = "桁数オーバー"
                                continue # 桁数オーバー時は次の処理をスキップ

                            if "." in num_str:
                                decimal_count = len(num_str.split(".")[1])
                            else:
                                decimal_count = 0  # 小数点がない（整数の）場合    


                            # 🌟 小数点以下桁数オーバー時の「切り上げ」自動処理
                            if decimal_count > kijun["小数"]:
                                try:
                                    # 精確な10進数に変換
                                    dec_val = Decimal(num_str)
                                    # 切り上げたい目標桁数のマスクを作成（例: kijun["小数"]が3なら "0.001"）
                                    quantize_mask = Decimal('10') ** -kijun["小数"]
                                    # ROUND_UP（指定の小数点桁数で切り上げ）を実行
                                    raised_val = dec_val.quantize(quantize_mask, rounding=ROUND_UP)
                                    
                                    # 結果を文字列に戻す（e-Tax仕様に合わせて余計な.0は省くか、ゼロ埋めするかをjsonで制御可能にするのが理想）
                                    data = str(raised_val)
                                    
                                    d.dprint(f"切り上げ処理実行: {num_str} -> {data}")
                                except Exception as _exception:
                                    errMsg = "シート:{},行:{},列:{},データ:{}". \
                                        format(sheet, row, column, sheet.cell(row=row, column=column).value)
                                    d.dprint(errMsg)
                                    e.eprint("切り上げ処理失敗", errMsg)
                                    data = "切り上げ処理失敗"
                            else:
                                data = num_str
                else:
                    data = ""
                # d.dprint(data)
                row_data.append(data)
                flag_empty = False
            csv_data.append(row_data)
        if not flag_empty and self.kaipage_yokusei and \
                sheet_data["合計行"] == 1 and goukei_kous != None:
            row_data = []
            row_data.append(format_kubun)
            row_data.append("1")
            for kou in sheet_data["項"]:
                column = kou["項番"]
                if column in goukei_kous:
                    goukei = goukei_dict[column]
                else:
                    goukei = ""
                row_data.append(goukei)           
            csv_data.append(row_data)
        # d.dprint(csv_data)
        # d.dprint_method_end()
        return row_data

    def write_csv_file(self, csv_file_name, data_list):
        '''
        データリストからＣＳＶファイルを作成する。
    
        Parameters
        ----------
        csv_file_mame : str
            作成するCSVファイル名。
        data_list : list of list of object
            CSVファイルに出力するデータ。
    
        Returns
        -------
        '''
        # d.dprint_method_start()
        # d.dprint(csv_file_name)
        # d.dprint(data_list)
        
        # errorsは指定せず、厳格にUnicodeEncodeErrorを発生させる設定にします
        with open(csv_file_name, "w", newline='', encoding="cp932") as f:
            writer = csv.writer(f)
            
            # 各行のインデックス（row_idx）を元に、実際のExcelの行番号を計算できるようにします
            for row_idx, row in enumerate(data_list):
                # d.dprint(row)
                try:
                    writer.writerow(row)
                except UnicodeEncodeError as encode_error:
                    # 💡 不正な文字が含まれている箇所の特定ロジック
                    # 通常、タイトル行を飛ばした後のデータなので、実際のExcelの行を特定する情報をログに流します
                    # row_idx + 1 に sheet_data["タイトル行数"] 相当を足すことでより正確になります
                    error_line_info = f"Excel行インデックス: 約 {row_idx + 1} 行目付近"
                    
                    # 変換に失敗した具体的な文字情報を抜き出す
                    bad_char = encode_error.object[encode_error.start:encode_error.end]
                    
                    # 💡 不正文字の16進数（Unicode / UTF-16）表記を生成 (例: 'U+5高' ではなく 'U+9AD9' など、文字列全体をループ)
                    hex_notation = " ".join([f"U+{ord(c):04X}" for c in bad_char])
                    
                    errMsg = "ファイル:{}, 位置:{}, 不正文字:[{}] (16進数: {}), 原因データ:{}".format(
                        csv_file_name,
                        error_line_info,
                        bad_char,
                        hex_notation,
                        row
                    )
                    
                    d.dprint(errMsg)
                    e.eprint("Shift_JIS変換不能（環境依存文字エラー）", errMsg)
                    
                    # 業務要件に応じて、ここでプログラムを完全に止める（raise）か、
                    # あるいはそのファイルだけスキップして処理を続けるかを選択できます。
                    # raise encode_error
                    
        # d.dprint_method_end()


def read_settei_file( str_file_name ):
    with open(str_file_name, mode='r', encoding='utf-8') as file:
        data = json5.load(file)
    return data



def show_config_dialog():
    """1つの画面で設定ファイル・Excel・出力フォルダ・改ページ抑制を入力させる"""
    result = {}

    root = tk.Tk()
    root.title("e-Tax用 CSV作成設定")
    root.attributes("-topmost", True)  # 最前面に表示

    # デフォルト値の設定（現在の作業フォルダ基準）
    current_dir = os.getcwd()
    var_config = tk.StringVar(value=os.path.join(current_dir, "exTaxConfig.json5"))
    var_excel = tk.StringVar(value=os.path.join(current_dir, "勘定科目内訳明細書.xlsx"))
    var_csv_dir = tk.StringVar(value=current_dir)
    var_kaipage = tk.BooleanVar(value=True)  # デフォルト：抑制する(True)

    # --- 参照ボタン用のコールバック関数 ---
    def browse_config():
        path = filedialog.askopenfilename(
            parent=root,
            title="設定ファイル(json5)を選択",
            filetypes=[("JSON5ファイル", "*.json5 *.json"), ("すべて", "*.*")],
        )
        if path:
            var_config.set(path)

    def browse_excel():
        path = filedialog.askopenfilename(
            parent=root,
            title="Excelファイルを選択",
            filetypes=[("Excelファイル", "*.xlsx *.xlsm"), ("すべて", "*.*")],
        )
        if path:
            var_excel.set(path)

    def browse_csv_dir():
        path = filedialog.askdirectory(parent=root, title="CSV出力先フォルダを選択")
        if path:
            var_csv_dir.set(path)

    # --- 実行・キャンセルボタンの処理 ---
    def on_execute():
        if not os.path.exists(var_config.get()):
            messagebox.showerror(
              "エラー", "設定ファイル(json5)が見つかりません。", parent=root
          )
            return
        if not os.path.exists(var_excel.get()):
            messagebox.showerror("エラー", "Excelファイルが見つかりません。", parent=root)
            return
        if not os.path.isdir(var_csv_dir.get()):
            messagebox.showerror("エラー", "CSV出力先フォルダが存在しません。", parent=root)
            return
    
        result["config_file"] = var_config.get()
        result["excel_file"] = var_excel.get()
        result["csv_folder"] = var_csv_dir.get()
        result["kaipage_yokusei"] = var_kaipage.get()
        root.destroy()

    def on_cancel():
        root.destroy()
        sys.exit()

    # ウィンドウの「×」ボタンで閉じられた場合も安全に終了
    root.protocol("WM_DELETE_WINDOW", on_cancel)

    # --- 画面レイアウト（グリッド配置） ---
    pad_opts = {"padx": 8, "pady": 6}

    # 1行目：設定ファイル
    tk.Label(root, text="設定ファイル (json5):").grid(
      row=0, column=0, sticky="e", **pad_opts
  )
    tk.Entry(root, textvariable=var_config, width=50).grid(
      row=0, column=1, **pad_opts
  )
    tk.Button(root, text="参照...", command=browse_config).grid(
      row=0, column=2, **pad_opts
  )

    # 2行目：Excelファイル
    tk.Label(root, text="入力Excelファイル:").grid(
      row=1, column=0, sticky="e", **pad_opts
  )
    tk.Entry(root, textvariable=var_excel, width=50).grid(
      row=1, column=1, **pad_opts
  )
    tk.Button(root, text="参照...", command=browse_excel).grid(
      row=1, column=2, **pad_opts
  )

    # 3行目：CSV出力フォルダ
    tk.Label(root, text="CSV出力先フォルダ:").grid(
      row=2, column=0, sticky="e", **pad_opts
  )
    tk.Entry(root, textvariable=var_csv_dir, width=50).grid(
      row=2, column=1, **pad_opts
  )
    tk.Button(root, text="参照...", command=browse_csv_dir).grid(
      row=2, column=2, **pad_opts
  )

    # 4行目：改ページ抑制チェックボックス
    tk.Checkbutton(
      root,
      text="改ページを抑制する（合計行をCSVに出力しない）",
      variable=var_kaipage,
  ).grid(row=3, column=1, sticky="w", **pad_opts)

    # 5行目：実行・キャンセルボタン
    btn_frame = tk.Frame(root)
    btn_frame.grid(row=4, column=0, columnspan=3, pady=12)
    tk.Button(
      btn_frame,
      text="CSV作成実行",
      command=on_execute,
      width=15,
      bg="#e1f5fe",
      ).pack(side="left", padx=10)
    tk.Button(btn_frame, text="キャンセル", command=on_cancel, width=12).pack(
      side="left", padx=10
      )

    # 画面中央付近に表示して待機
    root.eval("tk::PlaceWindow . center")
    root.mainloop()
    
    return result


if __name__ == "__main__":
    #
    # 1画面ダイアログから設定を取得
    params = show_config_dialog()

    # 設定ファイルの読み込みと変換実行
    config_data = read_settei_file(params["config_file"])
    exTax = ExTax(
			config_data,
			params["excel_file"],
			params["csv_folder"],
			params["kaipage_yokusei"],
	)
    exTax.execute()

    # 完了メッセージの表示
    done_root = tk.Tk()
    done_root.withdraw()
    done_root.attributes("-topmost", True)
    messagebox.showinfo("完了", "すべてのCSVファイルの作成が完了しました。")
    done_root.destroy()		