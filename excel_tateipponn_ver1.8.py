import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import glob
import datetime
import os
import threading
import ctypes

# --------------------------------------------------------------------
# ■ 高DPI対応設定
# --------------------------------------------------------------------
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# --------------------------------------------------------------------
# ■ カスタムTkinterの設定
# --------------------------------------------------------------------
ctk.set_appearance_mode("Light")  
ctk.set_default_color_theme("dark-blue")
ctk.set_widget_scaling(1.0) 

# --------------------------------------------------------------------
# ■ 処理ロジック
# --------------------------------------------------------------------
def process_files(folder_path, sheet_identifier, columns_to_extract, app_instance):
    try:
        app_instance.update_status("処理を開始します...", "processing")
        
        # 検索対象にcsvを追加
        extensions = ["*.xlsx", "*.xls", "*.csv"]
        all_files_path = []
        for ext in extensions:
            all_files_path.extend(glob.glob(os.path.join(folder_path, ext)))
        
        # 【追加機能1】過去に出力した「一本化_」から始まるファイルを除外する
        all_files_path = [
            path for path in all_files_path 
            if not os.path.basename(path).startswith("一本化_")
        ]

        if not all_files_path:
            raise ValueError("指定されたフォルダに処理対象のExcelまたはCSVファイルが見つかりません。")

        # --- 1. 列名の自動抽出（指定がない場合） ---
        if not columns_to_extract:
            app_instance.update_status("全ファイル共通の列名を検索中...", "processing")
            common_columns_set = None
            first_file_columns_order = []

            for i, file_path in enumerate(all_files_path):
                try:
                    ext = os.path.splitext(file_path)[1].lower()
                    if ext == ".csv":
                        # 【追加機能3】CSVの文字コードフォールバック (ヘッダー抽出時)
                        try:
                            df_header = pd.read_csv(file_path, nrows=0, encoding='utf-8-sig')
                        except UnicodeDecodeError:
                            df_header = pd.read_csv(file_path, nrows=0, encoding='shift_jis')
                    else:
                        df_header = pd.read_excel(file_path, sheet_name=sheet_identifier, nrows=0)
                    
                    current_columns = set(df_header.columns)
                    if common_columns_set is None:
                        common_columns_set = current_columns
                        first_file_columns_order = list(df_header.columns)
                    else:
                        common_columns_set.intersection_update(current_columns)
                except Exception as e:
                    print(f"列抽出スキップ: {os.path.basename(file_path)} - {e}")
                    continue
            
            if not common_columns_set:
                raise ValueError("全ファイルに共通する列が見つかりませんでした。")
            columns_to_extract = [col for col in first_file_columns_order if col in common_columns_set]

        # --- 2. データ抽出処理 ---
        combined_df = pd.DataFrame()
        app_instance.update_status("データ抽出処理中...", "processing")
        processed_files_count = 0

        for file_path in all_files_path:
            try:
                ext = os.path.splitext(file_path)[1].lower()
                if ext == ".csv":
                    # 【追加機能3】CSVの文字コードフォールバック (データ抽出時)
                    try:
                        df = pd.read_csv(file_path, dtype=str, encoding='utf-8-sig')
                    except UnicodeDecodeError:
                        df = pd.read_csv(file_path, dtype=str, encoding='shift_jis')
                else:
                    df = pd.read_excel(file_path, sheet_name=sheet_identifier, dtype=str)
                
                valid_columns = [col for col in columns_to_extract if col in df.columns]
                if not valid_columns:
                    continue

                extracted_df = df[valid_columns]
                combined_df = pd.concat([combined_df, extracted_df], ignore_index=True)
                processed_files_count += 1
            except Exception as e:
                print(f"エラー: {os.path.basename(file_path)} の処理中にエラー: {e}")
        
        if combined_df.empty:
            raise ValueError("有効なデータを抽出できませんでした。")

        # --- 3. 書き出し（Excel制限 1,048,576行の判定） ---
        now = datetime.datetime.now()
        date_time_str = now.strftime("%Y%m%d_%H%M%S")
        record_count = len(combined_df)
        EXCEL_LIMIT = 1048576

        if record_count >= EXCEL_LIMIT:
            # CSVで保存
            file_name = f"一本化_{date_time_str}_sum{record_count}records.csv"
            output_file_path = os.path.join(folder_path, file_name)
            combined_df.to_csv(output_file_path, index=False, encoding='utf-8-sig')
            save_type = "CSV (Excel行数制限超過のため)"
        else:
            # Excelで保存
            file_name = f"一本化_{date_time_str}_sum{record_count}records.xlsx"
            output_file_path = os.path.join(folder_path, file_name)
            with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
                combined_df.to_excel(writer, sheet_name='統合データ', index=False)
                workbook = writer.book
                worksheet = writer.sheets['統合データ']
                text_format = workbook.add_format({'num_format': '@'})
                worksheet.set_column(0, len(combined_df.columns) - 1, None, text_format)
            save_type = "Excel"

        success_message = (
            f"処理完了！\n"
            f"{processed_files_count}個のファイルから合計{record_count}行抽出しました。\n\n"
            f"保存形式: {save_type}\n"
            f"保存先: {output_file_path}"
        )
        messagebox.showinfo("成功", success_message)
        app_instance.update_status("待機中...", "normal")

    except Exception as e:
        messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")
        app_instance.update_status("エラーが発生しました", "error")
        
    finally:
        # 【追加機能2】処理が終わったら（成功・エラー問わず）ボタンを再び有効化する
        app_instance.after(0, lambda: app_instance.execute_button.configure(state="normal"))

# --------------------------------------------------------------------
# ■ アプリケーションクラス
# --------------------------------------------------------------------
class ExcelMergerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Excel/CSV縦一本化ツール_ver1.8")
        self.geometry("620x700") 
        
        self.main_font = ("Yu Gothic UI", 14)
        self.bold_font = ("Yu Gothic UI", 14, "bold")
        self.small_font = ("Yu Gothic UI", 12)
        self.btn_font = ("Yu Gothic UI", 16, "bold")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.main_frame = ctk.CTkScrollableFrame(self, corner_radius=0)
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.create_folder_section()
        self.create_sheet_section()
        self.create_column_section()
        self.create_action_section()
        self.create_info_section()

    def create_folder_section(self):
        frame = ctk.CTkFrame(self.main_frame)
        frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(frame, text="1. 対象フォルダを選択", font=self.bold_font).grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 5))
        self.folder_path_entry = ctk.CTkEntry(frame, placeholder_text="Excel/CSVファイルが入っているフォルダパス", font=self.main_font)
        self.folder_path_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(10, 5), pady=10)
        ctk.CTkButton(frame, text="参照", width=80, command=self.select_folder, font=self.main_font).grid(row=1, column=2, sticky="e", padx=(0, 10), pady=10)

    def create_sheet_section(self):
        frame = ctk.CTkFrame(self.main_frame)
        frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text="2. 読み込むシートの指定 (Excelのみ)", font=self.bold_font).grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        self.sheet_select_var = tk.IntVar(value=2)
        ctk.CTkRadioButton(frame, text="左からの順番で指定 (数値)", variable=self.sheet_select_var, value=1, command=self.toggle_sheet_input, font=self.main_font).grid(row=1, column=0, sticky="w", padx=20, pady=5)
        ctk.CTkRadioButton(frame, text="シート名で指定 (文字)", variable=self.sheet_select_var, value=2, command=self.toggle_sheet_input, font=self.main_font).grid(row=2, column=0, sticky="w", padx=20, pady=5)
        self.sheet_input_label = ctk.CTkLabel(frame, text="シート名を入力:", text_color="gray", font=self.main_font)
        self.sheet_input_label.grid(row=3, column=0, sticky="w", padx=25, pady=(5, 0))
        self.sheet_entry = ctk.CTkEntry(frame, font=self.main_font)
        self.sheet_entry.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 15))
        self.sheet_entry.insert(0, "集計")

    def create_column_section(self):
        frame = ctk.CTkFrame(self.main_frame)
        frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text="3. 抽出する列の指定", font=self.bold_font).grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        self.column_select_var = tk.IntVar(value=2)
        ctk.CTkRadioButton(frame, text="全ファイル共通の列を自動抽出", variable=self.column_select_var, value=1, command=self.toggle_column_input, font=self.main_font).grid(row=1, column=0, sticky="w", padx=20, pady=5)
        ctk.CTkRadioButton(frame, text="指定した列のみ抽出", variable=self.column_select_var, value=2, command=self.toggle_column_input, font=self.main_font).grid(row=2, column=0, sticky="w", padx=20, pady=5)
        self.col_input_label = ctk.CTkLabel(frame, text="列名をカンマ区切りで入力 (例: ID, 名前, 住所):", text_color="gray", font=self.main_font)
        self.col_input_label.grid(row=3, column=0, sticky="w", padx=25, pady=(5, 0))
        self.columns_entry = ctk.CTkEntry(frame, font=self.main_font)
        self.columns_entry.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 15))
        self.columns_entry.insert(0, "ID, 名前, 住所")

    def create_action_section(self):
        self.execute_button = ctk.CTkButton(self.main_frame, text="処理を実行する", command=self.start_processing, height=50, font=self.btn_font, fg_color="#107C41", hover_color="#0C5C30")
        self.execute_button.grid(row=3, column=0, padx=20, pady=20, sticky="ew")
        self.status_label = ctk.CTkLabel(self.main_frame, text="待機中...", font=self.small_font)
        self.status_label.grid(row=4, column=0, padx=20, pady=(0, 10))

    def create_info_section(self):
        info_text = (
            "【概要】\n"
            "複数Excel/CSVファイル中の指定データを、列名選別・並替の上で縦1本化します。\n"
            "アクセスやパワークエリなどありますが、サクッとしたいときに！\n\n"

            "【大規模データ・形式への対応】\n"
            "・CSVファイルも自動で読み込み対象に含めます。（Shift-JISにも対応）\n"
            "・合計が1,048,576行（Excel制限）を超える場合、自動的にCSV形式で書き出します。\n"
            "・保存されるCSVは「BOM付きUTF-8」のため、Excelで開いても文字化けしません。\n\n"

            "【前提条件】\n"
            "・見出しは1行目にあること（複数行見出しは不可）。\n"
            "・左に空列があっても動作に支障なし。\n\n"

            "【使い方】\n"
            "0. 一本化したいファイルを任意の【特定フォルダ】にまとめて格納。\n"
            "1. 対象ファイルが入ったフォルダを選択。\n"
            "2. Excelのシートを指定（CSVはシートに関わらず全読み込み）。\n"
            "3. 必要な列を指定（カンマ区切り または 共通列の自動抽出）。\n"
            "4. 「処理を実行する」ボタンをクリック。"
        )
        
        frame = ctk.CTkFrame(self.main_frame)
        frame.grid(row=5, column=0, padx=20, pady=(10, 30), sticky="ew")
        frame.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(frame, text=info_text, justify="left", anchor="w", font=self.small_font)
        label.grid(row=0, column=0, padx=15, pady=15, sticky="w")

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path_entry.delete(0, tk.END); self.folder_path_entry.insert(0, folder_path)

    def toggle_sheet_input(self):
        if self.sheet_select_var.get() == 1:
            self.sheet_input_label.configure(text="シートが左から何番目か (半角数字):"); self.sheet_entry.delete(0, tk.END); self.sheet_entry.insert(0, "1")
        else:
            self.sheet_input_label.configure(text="シート名を入力 (文字):"); self.sheet_entry.delete(0, tk.END); self.sheet_entry.insert(0, "集計")

    def toggle_column_input(self):
        if self.column_select_var.get() == 1:
            self.columns_entry.configure(state="disabled", fg_color=("gray90", "gray90"))
        else:
            self.columns_entry.configure(state="normal", fg_color=("white", "white"))

    def update_status(self, message, status_type="normal"):
        color = "white" if ctk.get_appearance_mode() == "Dark" else "black"
        if status_type == "error": color = "#ff5555"
        elif status_type == "processing": color = "#3b8ed0"
        self.status_label.configure(text=message, text_color=color)

    def start_processing(self):
        folder = self.folder_path_entry.get()
        sheet_input = self.sheet_entry.get()
        if not folder or not sheet_input:
            messagebox.showwarning("入力エラー", "フォルダパスとシート指定は必須です。")
            return

        if self.sheet_select_var.get() == 1:
            try:
                sheet_identifier = int(sheet_input) - 1
            except ValueError:
                messagebox.showwarning("入力エラー", "数字で入力してください。")
                return
        else:
            sheet_identifier = sheet_input

        columns_list = [col.strip() for col in self.columns_entry.get().split(',')] if self.column_select_var.get() == 2 else []

        # 【追加機能2】処理開始時にボタンを無効化する
        self.execute_button.configure(state="disabled")
        
        threading.Thread(target=process_files, args=(folder, sheet_identifier, columns_list, self)).start()

if __name__ == "__main__":
    app = ExcelMergerApp()
    app.mainloop()