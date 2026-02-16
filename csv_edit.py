import csv
import TkEasyGUI as sg

def main():
    while True:
        # CSVファイルを選ぶ
        files = sg.popup_get_file(
            "編集・結合したいCSVファイルを選択（複数可）",
            multiple_files=True,
            no_window=True,
            file_types=(("CSVファイル", "*.csv"),)
        )
        
        if not files:
            break
            
        # 複数のCSVファイルをまとめる
        all_data = []
        header = []
        
        for i, filename in enumerate(files):
            data = read_csv(filename)
            if data is None or not data:
                continue

            if i == 0:
                header = data[0]
                all_data = data # ヘッダーごと追加
            else:
                current_header = data[0]
                current_body = data[1:]
                if current_header == header:
                    all_data += current_body
                else:
                    all_data += data # ヘッダー不一致でも強制追加

        # 編集画面へ
        if edit_and_show_csv(all_data) == False:
            break

def read_csv(filename):
    encodings = ["UTF-8", "CP932", "EUC-JP"]
    for enc in encodings:
        try:
            with open(filename, "r", encoding=enc) as f:
                reader = csv.reader(f)
                return [row for row in reader]
        except:
            pass
    return None

def edit_and_show_csv(data):
    if len(data) == 0:
        data = [["空"],["データなし"]]

    header = data[0]
    table_values = data[1:] # ヘッダー以外のデータ

    # レイアウト
    layout = [
        [sg.Text("行を選択して「選択行を編集」ボタンを押してください", text_color="blue")],
        [sg.Table(
            key="-table-",
            values=table_values,
            headings=header,
            expand_x=True,
            expand_y=True,
            justification='left',
            auto_size_columns=True,
            max_col_width=30,
            font=("Arial", 14),
            enable_events=True, # 選択イベントを有効化
            select_mode=sg.TABLE_SELECT_MODE_BROWSE # 1行選択モード
        )],
        [sg.Button('選択行を編集', button_color="green"), sg.Button('ファイル選択に戻る'), sg.Button('名前を付けて保存'), sg.Button('終了')]
    ]

    window = sg.Window("CSVエディタ", layout, size=(600, 400), resizable=True, finalize=True)
    
    # ダブルクリックも一応バインド
    window["-table-"].bind('<Double-Button-1>', 'Double')

    flag_continue = False
    
    while True:
        event, values = window.read()
        
        if event in [sg.WINDOW_CLOSED, "終了"]:
            break
            
        if event == "ファイル選択に戻る":
            flag_continue = True
            break

        # --- 編集処理（ボタン または ダブルクリック） ---
        if event == "選択行を編集" or event == "-table-Double":
            
            # 行が選択されているかチェック
            if "-table-" in values and len(values["-table-"]) > 0:
                try:
                    # ★【ここを修正しました】★
                    # 選択された値を取得（数字かもしれないし、文字かもしれない）
                    selected_item = values["-table-"][0]
                    
                    selected_index = -1
                    
                    # もし選択されたのが「数字（行番号）」ならそのまま使う
                    if isinstance(selected_item, int):
                        selected_index = selected_item
                    
                    # もし「文字（数字の文字列）」なら数字に変換する
                    elif isinstance(selected_item, str) and selected_item.isdigit():
                         selected_index = int(selected_item)
                         
                    # もし「文字（データの値）」なら、データの中から探す
                    else:
                        # データ（table_values）の中から、1列目が一致する行を探す
                        for i, row in enumerate(table_values):
                            # 念のため文字列にして比較
                            if str(row[0]) == str(selected_item):
                                selected_index = i
                                break
                    
                    # 正しい行番号が見つかった場合のみ編集へ
                    if selected_index != -1 and selected_index < len(table_values):
                        target_row = table_values[selected_index] # その行のデータ
                        
                        # 編集ウィンドウを開く
                        new_row = popup_edit_row(header, target_row)
                        
                        # 更新があった場合
                        if new_row is not None:
                            table_values[selected_index] = new_row
                            window["-table-"].update(values=table_values)
                    else:
                        sg.popup("行を特定できませんでした。")
                        
                except Exception as e:
                    sg.popup_error(f"エラーが発生しました: {e}")
            else:
                sg.popup("編集したい行を選択してください。")

        # --- 保存機能 ---
        if event == "名前を付けて保存":
            filename = sg.popup_get_file(
                "保存先を選択",
                save_as=True,
                no_window=True,
                file_types=(("CSVファイル","*.csv"),),
                default_extension=".csv"
            )
            if filename:
                try:
                    with open(filename, "w", encoding="utf-8", newline="") as f:
                        writer = csv.writer(f)
                        writer.writerows([header] + table_values)
                    sg.popup("保存しました！")
                except Exception as e:
                    sg.popup_error(f"保存エラー: {e}")

    window.close()
    return flag_continue

# 行編集用のポップアップウィンドウ
def popup_edit_row(header, row_data):
    input_rows = []
    # 値が足りない場合のパディング
    safe_row_data = row_data + [""] * (len(header) - len(row_data))

    for h, v in zip(header, safe_row_data):
        input_rows.append([sg.Text(f"{h}:", size=(15,1)), sg.Input(v, key=h, expand_x=True)])
    
    layout = [
        [sg.Text("値を編集してください")],
        # scrollable などのオプションを削除しました
        [sg.Column(input_rows)],
        [sg.Button("更新", key="-OK-"), sg.Button("キャンセル")]
    ]
    
    edit_win = sg.Window("行の編集", layout, modal=True)
    
    new_data = None
    while True:
        e, v = edit_win.read()
        if e in [sg.WINDOW_CLOSED, "キャンセル"]:
            break
        if e == "-OK-":
            new_data = [v[h] for h in header]
            break
            
    edit_win.close()
    return new_data

if __name__ == "__main__":
    main()