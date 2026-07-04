import streamlit as st
import pandas as pd
import datetime
import io

# --------------------------------------------------------------------
# ■ ページ基本設定
# --------------------------------------------------------------------
st.set_page_config(
    page_title="Excel/CSV縦一本化ツール Web版",
    layout="centered"
)

# 大タイトル（上下の余白を極限までカット）
st.markdown("""
    <div style="margin-top: 5px; margin-bottom: 0px;">
        <h1 style="font-size: 2.0rem; color: var(--text-color); font-weight: 700; border: none; padding: 0; margin: 0;">
            Excel/CSV縦一本化ツール
        </h1>
    </div>
""", unsafe_allow_html=True)
st.caption("複数Excel/CSVファイル中の指定データを、列名選別・並替の上で縦1本化します。")


# --------------------------------------------------------------------
# ■ 1. ファイル選択セクション
# --------------------------------------------------------------------
st.markdown("""
    <div style="border-left: 5px solid #107C41; padding-left: 10px; margin-top: 25px; margin-bottom: 10px;">
        <h2 style="font-size: 1.3rem; color: var(--text-color); font-weight: 700; margin: 0; padding: 0; border: none;">
            1. 対象ファイルのアップロード
        </h2>
    </div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "一本化したいExcel（.xlsx, .xls）またはCSVファイルをまとめて選択、またはドロップしてください",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)


# --------------------------------------------------------------------
# ■ 2. シート指定セクション (Excel用)
# --------------------------------------------------------------------
st.markdown("<hr style='border:0;border-top:1px solid var(--text-color);opacity:0.15;margin:20px 0;'>", unsafe_allow_html=True)
st.markdown("""
    <div style="border-left: 5px solid #107C41; padding-left: 10px; margin-top: 0px; margin-bottom: 10px;">
        <h2 style="font-size: 1.3rem; color: var(--text-color); font-weight: 700; margin: 0; padding: 0; border: none;">
            2. 読み込むシートの指定 (Excelのみ)
        </h2>
    </div>
""", unsafe_allow_html=True)

sheet_select_mode = st.radio(
    "指定方法を選んでください",
    options=["シート名で指定 (文字)", "左からの順番で指定 (数値)"],
    index=0
)

if sheet_select_mode == "シート名で指定 (文字)":
    sheet_input = st.text_input("シート名を入力:", value="集計")
    sheet_identifier = sheet_input
else:
    sheet_input = st.number_input("シートが左から何番目か (1以上の整数):", min_value=1, value=1, step=1)
    sheet_identifier = int(sheet_input) - 1


# --------------------------------------------------------------------
# ■ 3. 抽出列の指定セクション
# --------------------------------------------------------------------
st.markdown("<hr style='border:0;border-top:1px solid var(--text-color);opacity:0.15;margin:20px 0;'>", unsafe_allow_html=True)
st.markdown("""
    <div style="border-left: 5px solid #107C41; padding-left: 10px; margin-top: 0px; margin-bottom: 10px;">
        <h2 style="font-size: 1.3rem; color: var(--text-color); font-weight: 700; margin: 0; padding: 0; border: none;">
            3. 抽出する列の指定
        </h2>
    </div>
""", unsafe_allow_html=True)

column_select_mode = st.radio(
    "抽出方法を選んでください",
    options=["指定した列のみ抽出", "全ファイル共通の列を自動抽出"],
    index=0
)

if column_select_mode == "指定した列のみ抽出":
    columns_input = st.text_input("列名をカンマ区切りで入力:", value="ID, 名前, 住所")
    columns_to_extract = [col.strip() for col in columns_input.split(',')] if columns_input else []
else:
    columns_to_extract = []


# --------------------------------------------------------------------
# ■ 4. 処理実行 & ダウンロード
# --------------------------------------------------------------------
st.markdown("<div style='margin-top: 25px;'></div>", unsafe_allow_html=True)

if st.button("処理を実行する", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("ファイルがアップロードされていません。")
    else:
        try:
            with st.spinner("処理を実行中..."):
                
                # --- 3-1. 列名の自動抽出 ---
                if not columns_to_extract:
                    common_columns_set = None
                    first_file_columns_order = []

                    for file in uploaded_files:
                        file.seek(0)
                        ext = file.name.split('.')[-1].lower()
                        try:
                            if ext == "csv":
                                try:
                                    df_header = pd.read_csv(file, nrows=0, encoding='utf-8-sig')
                                except UnicodeDecodeError:
                                    file.seek(0)
                                    df_header = pd.read_csv(file, nrows=0, encoding='shift_jis')
                            else:
                                df_header = pd.read_excel(file, sheet_name=sheet_identifier, nrows=0)
                            
                            current_columns = set(df_header.columns)
                            if common_columns_set is None:
                                common_columns_set = current_columns
                                first_file_columns_order = list(df_header.columns)
                            else:
                                common_columns_set.intersection_update(current_columns)
                        except Exception as e:
                            continue
                    
                    if not common_columns_set:
                        raise ValueError("全ファイルに共通する列が見つかりませんでした。")
                    columns_to_extract = [col for col in first_file_columns_order if col in common_columns_set]

                # --- 3-2. データ抽出処理 ---
                combined_df = pd.DataFrame()
                processed_files_count = 0

                for file in uploaded_files:
                    file.seek(0)
                    ext = file.name.split('.')[-1].lower()
                    try:
                        if ext == "csv":
                            try:
                                df = pd.read_csv(file, dtype=str, encoding='utf-8-sig')
                            except UnicodeDecodeError:
                                file.seek(0)
                                df = pd.read_csv(file, dtype=str, encoding='shift_jis')
                        else:
                            df = pd.read_excel(file, sheet_name=sheet_identifier, dtype=str)
                        
                        valid_columns = [col for col in columns_to_extract if col in df.columns]
                        if not valid_columns:
                            continue

                        extracted_df = df[valid_columns]
                        combined_df = pd.concat([combined_df, extracted_df], ignore_index=True)
                        processed_files_count += 1
                    except Exception as e:
                        st.error(f"エラー: {file.name} の処理中にエラーが発生しました: {e}")
                
                if combined_df.empty:
                    raise ValueError("有効なデータを抽出できませんでした。")

                # --- 3-3. メモリ上への書き出し ---
                record_count = len(combined_df)
                EXCEL_LIMIT = 1048576
                
                now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

                if record_count >= EXCEL_LIMIT:
                    csv_buffer = io.StringIO()
                    combined_df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                    download_data = csv_buffer.getvalue().encode('utf-8-sig')
                    filename = f"一本化_{now}_sum{record_count}records.csv"
                    mime_type = "text/csv"
                    save_type = "CSV (Excel行数制限超過のため)"
                else:
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                        combined_df.to_excel(writer, sheet_name='統合データ', index=False)
                        workbook = writer.book
                        worksheet = writer.sheets['統合データ']
                        text_format = workbook.add_format({'num_format': '@'})
                        worksheet.set_column(0, len(combined_df.columns) - 1, None, text_format)
                    download_data = excel_buffer.getvalue()
                    filename = f"一本化_{now}_sum{record_count}records.xlsx"
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    save_type = "Excel"

                st.success(f"処理完了！ {processed_files_count}個のファイルから合計 {record_count:,} 行を抽出しました。（保存形式: {save_type}）")
                
                st.download_button(
                    label="統合データをダウンロードする",
                    data=download_data,
                    file_name=filename,
                    mime=mime_type,
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"処理中にエラーが発生しました:\n{e}")

# --------------------------------------------------------------------
# ■ 説明インフォメーション
# --------------------------------------------------------------------
st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
with st.expander("詳細な仕様・前提条件を確認する"):
    st.markdown("""
    **【大規模データ・形式への対応】**
    * CSVファイルも自動で読み込み対象に含めます。（Shift-JIS/BOM付きUTF-8に対応）
    * 合計が1,048,576行（Excel制限）を超える場合、自動的にCSV形式に切り替えてダウンロード用データを生成します。
    * 生成されるCSVは「BOM付きUTF-8」のため、Excelでそのまま開いても文字化けしません。

    **【前提条件】**
    * 見出しは1行目にあること（複数行見出しは不可）。
    * 左に空列があっても動作に支障なし。
    """)