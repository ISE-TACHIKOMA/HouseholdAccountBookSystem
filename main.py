#このプログラムは、クレジットカードの利用明細をCSV形式で読み込み、支出をカテゴリ別に集計し、円グラフで表示するものです。
# 使い方:
# 1. CSVファイルを選択します。
# 2. 読み込みボタンを押すと、CSVファイルが読み込まれ、利用日、商品名、料金が表示されます。
# 3. カテゴリを選択し、選択項目にカテゴリ割当ボタンを押すと、選択した行にカテゴリが割り当てられます。
# 4. 円グラフ表示ボタンを押すと、カテゴリ別の円グラフが表示されます。

import PySimpleGUI as sg
import pandas as pd
import matplotlib.pyplot as plt
import difflib
# import japanize_matplotlib
import os
import re

# 日本語フォント対応（matplotlib）
plt.rcParams["font.family"] = "Meiryo"

# 分類カテゴリ
categories = ['食費', '電気代', '交通費', '情報通信費', '日用雑貨', '遊戯', '学業', 'ガス代', '水道代','雑貨','立替金']

# 想定カラム名候補
expected_cols = {
    '利用日': ['利用日', '日付', '購入日', '取引日'],
    '商品名': ['商品名', '内容', '品名', '摘要', '利用店名'],
    '料金': ['料金', '金額', '支出', '金額（円）']
}

pay_month = 0

# 類似列名を検索
def find_closest_column(columns, target_names):
    for target in target_names:
        match = difflib.get_close_matches(target, columns, n=1, cutoff=0.6)
        if match:
            return match[0]
    return None

# CSV読み込み（列名が多少違っても対応）
def load_csv(file_path):
    df = pd.read_csv(file_path)
    col_map = {}
    for std_name, alt_names in expected_cols.items():
        matched = find_closest_column(df.columns, alt_names)
        if matched:
            col_map[std_name] = matched
        else:
            raise ValueError(f"'{std_name}' に対応する列が見つかりません")

    df_clean = df[[col_map['利用日'], col_map['商品名'], col_map['料金']]].copy()
    df_clean.columns = ['利用日', '商品名', '料金']
    df_clean['カテゴリ'] = ''
    df_clean['料金'] = pd.to_numeric(df_clean['料金'], errors='coerce').fillna(0)
    return df_clean

# 円グラフ作成（カテゴリ金額を昇順にして表示）
def plot_pie_chart(df, pay_month):
    summary = df.groupby('カテゴリ')['料金'].sum().sort_values()
    if summary.empty:
        sg.popup("グラフに表示できるデータがありません。")
        return

    labels = [f"{cat}: {int(val):,}円" for cat, val in summary.items()]
    plt.figure(figsize=(6, 6))
    plt.pie(summary, labels=labels, autopct='%1.1f%%', startangle=140)
    plt.title(f'出費のカテゴリ別割合（{pay_month}）合計： {int(summary.sum()):,} 円')
    plt.tight_layout()
    plt.show()

# 支払月の推定（ファイル名から）
def extract_payment_month(file_path):
    fname = os.path.basename(file_path)
    match = re.search(r'(20\d{2})[-_/]?(0[1-9]|1[0-2])', fname)
    if match:
        year, month = match.groups()
        return f"{int(year)}年{int(month)}月分"
    return "支払月：不明"

# 金額集計表示更新
def update_summary_display(df, window):
    if df is None or df.empty:
        window['-SUMMARY-'].update('')
        window['-TOTAL-'].update('')
        return

    summary = df.groupby('カテゴリ')['料金'].sum()
    total = df['料金'].sum()

    lines = [f"{cat}: {int(val):,} 円" for cat, val in summary.items() if cat]
    summary_text = "\n".join(lines)
    window['-SUMMARY-'].update(summary_text)
    window['-TOTAL-'].update(f"総合計: {int(total):,} 円")

# 一時保存ファイル名
temp_file = "temp/temp_output.csv"
# 一時保存ディレクトリ作成
os.makedirs(os.path.dirname(temp_file), exist_ok=True)


# GUIレイアウト
sg.theme('LightBlue')
layout = [
    [sg.Text('CSVファイルを選択:'), sg.Input(), sg.FileBrowse(file_types=(("CSV Files", "*.csv"),))],
    [sg.Button('読み込み'), sg.Button('既定外の一時ファイル読込')],
    [sg.Text('カテゴリ:'), sg.Combo(categories, key='-CAT-', default_value=categories[0])],
    [sg.Button('選択項目にカテゴリ割当'), sg.Button('全選択'), sg.Button('選択解除')],
    [sg.Table(values=[], headings=['利用日', '商品名', '料金', 'カテゴリ'],
              key='-TABLE-', enable_events=True,
              auto_size_columns=False, col_widths=[12, 30, 8, 10],
              select_mode=sg.TABLE_SELECT_MODE_EXTENDED, justification='left',
              size=(None, 15))],
    [sg.Text('', key='-PAYMONTH-', size=(30, 1), font=('Arial', 11))],
    [sg.Text('', key='-TOTAL-', font=('Arial', 12, 'bold'))],
    [sg.Multiline('', size=(30, 6), key='-SUMMARY-', disabled=True, autoscroll=False, font=('Arial', 10))],
    [sg.Button('円グラフ表示'), sg.Button('終了')]
]

window = sg.Window('クレジットカード出費分析', layout, size=(800, 600),finalize=True)
# 初期化

df = None

# 起動時に一時ファイルが存在すれば読み込むか確認
if os.path.exists(temp_file):
    if sg.popup_yes_no("前回の一時ファイルがあります。読み込みますか？") == "Yes":
        try:
            df = pd.read_csv(temp_file)
            pay_month = df['_支払月'].iloc[0] if '_支払月' in df.columns else "不明"
            df = df.drop(columns=['_支払月', '_元ファイル'], errors='ignore')
            df['料金'] = pd.to_numeric(df['料金'], errors='coerce').fillna(0)
            window['-TABLE-'].update(values=df.values.tolist())
            window['-PAYMONTH-'].update(f"支払月：{pay_month}")
            update_summary_display(df, window)
        except:
            sg.popup_error("一時ファイルの読み込みに失敗しました。")

# イベントループ
while True:
    event, values = window.read()

    if event in (sg.WIN_CLOSED, '終了'):
        break

    elif event == '読み込み':
        file_path = values[0]
        try:
            df = load_csv(file_path)
            window['-TABLE-'].update(values=df.values.tolist())
            # 支払月推定
            pay_month = extract_payment_month(file_path)
            window['-PAYMONTH-'].update(f"支払月：{pay_month}")
            update_summary_display(df, window)
        except Exception as e:
            sg.popup_error(f"読み込み失敗: {e}")

    elif event == '一時保存':
        if df is not None:
            df['_支払月'] = pay_month
            df['_元ファイル'] = values[0]
            df.to_csv(temp_file, index=False)
            sg.popup("一時ファイルを保存しました。")

    elif event == '既定外の一時ファイル読込':
        file_path = sg.popup_get_file("CSVファイルを選択", file_types=(("CSV Files", "*.csv"),))
        if file_path:
            try:
                df = pd.read_csv(file_path)
                pay_month = df['_支払月'].iloc[0] if '_支払月' in df.columns else "不明"
                df = df.drop(columns=['_支払月', '_元ファイル'], errors='ignore')
                df['料金'] = pd.to_numeric(df['料金'], errors='coerce').fillna(0)
                window['-TABLE-'].update(values=df.values.tolist())
                window['-PAYMONTH-'].update(f"支払月：{pay_month}")
                update_summary_display(df, window)

            except Exception as e:
                sg.popup_error(f"CSV読み込み失敗: {e}")

    elif event == '-TABLE-':
        pass

    elif event == '選択項目にカテゴリ割当':
        if df is not None and values['-TABLE-']:
            selected_rows = values['-TABLE-']
            for row in selected_rows:
                df.at[row, 'カテゴリ'] = values['-CAT-']
            window['-TABLE-'].update(values=df.values.tolist())
            update_summary_display(df, window)
            if df is not None:
                df['_支払月'] = pay_month
                df['_元ファイル'] = values[0]
                df.to_csv(temp_file, index=False)


    elif event == '全選択':
        if df is not None:
            window['-TABLE-'].update(select_rows=list(range(len(df))))

    elif event == '選択解除':
        window['-TABLE-'].update(select_rows=[])

    elif event == '円グラフ表示':
        if df is not None:
            if (df['カテゴリ'] == '').any():
                sg.popup("すべての行にカテゴリを割り当ててください。")
            else:
                plot_pie_chart(df, pay_month)

window.close()
