import TkEasyGUI as eg
import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
from PIL import Image
import io
import os

# --- グローバル変数 ---
current_file = ""
doc_fitz = None
pdf_states = []     # [{'page_num': 0, 'rotation': 0, 'deleted': False}, ...]
current_page_index = 0

def load_pdf_file(filepath):
    """PDF読み込み"""
    global current_file, doc_fitz, pdf_states, current_page_index
    
    if ";" in filepath: filepath = filepath.split(";")[0]
    filepath = filepath.strip()
    
    if not filepath or not os.path.exists(filepath):
        eg.popup("ファイルが見つかりません: " + filepath)
        return False
        
    try:
        current_file = filepath
        doc_fitz = fitz.open(filepath)
        pdf_states = []
        for i in range(len(doc_fitz)):
            pdf_states.append({'page_num': i, 'rotation': 0, 'deleted': False})
        
        current_page_index = 0
        return True
    except Exception as e:
        eg.popup(f"読み込みエラー: {e}")
        return False

def get_preview_data(page_index, max_size=(400, 500)):
    """プレビュー画像生成"""
    global doc_fitz, pdf_states
    
    if not doc_fitz or page_index < 0 or page_index >= len(pdf_states):
        return None

    state = pdf_states[page_index]
    
    # 削除済みページはグレー表示
    if state['deleted']:
        img = Image.new('RGB', max_size, color='#555555')
    else:
        try:
            page = doc_fitz.load_page(state['page_num'])
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            if state['rotation'] != 0:
                img = img.rotate(-state['rotation'], expand=True)
        except Exception as e:
            return None

    img.thumbnail(max_size)
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    return bio.getvalue()

def parse_page_range(range_str, max_page):
    """'1-3, 5' などの文字列をインデックスのリストに変換"""
    indices = set()
    if not range_str.strip():
        return []
        
    # 全角数字やカンマにも対応
    range_str = range_str.replace("、", ",").replace("ー", "-")
    parts = range_str.split(",")
    
    for part in parts:
        part = part.strip()
        try:
            if "-" in part:
                start, end = map(int, part.split("-"))
                # ユーザー入力は1始まり、内部は0始まり
                for i in range(start - 1, end):
                    if 0 <= i < max_page:
                        indices.add(i)
            elif part.isdigit():
                i = int(part) - 1
                if 0 <= i < max_page:
                    indices.add(i)
        except ValueError:
            continue
    return sorted(list(indices))

def save_current_pdf(output_path):
    """保存処理"""
    global current_file, pdf_states
    try:
        reader = PdfReader(current_file)
        writer = PdfWriter()
        
        count = 0
        for state in pdf_states:
            if state['deleted']: continue
            
            page = reader.pages[state['page_num']]
            if state['rotation'] != 0:
                page.rotate(state['rotation'])
            writer.add_page(page)
            count += 1
            
        with open(output_path, "wb") as f:
            writer.write(f)
        return True, f"{count}ページを保存しました！"
    except Exception as e:
        return False, str(e)

# --- GUI Layout ---

# 操作パネルの定義
control_column = [
    [eg.Frame("1. 対象を選択", [
        [eg.Radio("現在のページのみ", "SCOPE", key="-SCOPE_CURRENT-", default=True)],
        [eg.Radio("すべてのページ", "SCOPE", key="-SCOPE_ALL-")],
        [eg.Radio("範囲指定 (例: 1-3, 5)", "SCOPE", key="-SCOPE_RANGE-"), 
         eg.Input(key="-RANGE_INPUT-", size=(10, 1))]
    ])],
    [eg.Frame("2. 操作を実行", [
        [eg.Button("↶ 左回転", key="-ROT_L-", size=(12, 1)), 
         eg.Button("↷ 右回転", key="-ROT_R-", size=(12, 1))],
        [eg.Button("× 削除/復元", key="-DEL-", button_color=("white", "red"), size=(26, 1))]
    ])],
    [eg.Text("", key="-MSG-", text_color="blue")]
]

layout = [
    [eg.Text("PDF 高機能編集ツール", font=("Arial", 16))],
    
    # ファイル読込エリア
    [eg.Text("ファイル:"), eg.Input(key="-FILE-", readonly=True), 
     eg.FilesBrowse("選択", file_types=(("PDF Files", "*.pdf"),)),
     eg.Button("読込", key="-LOAD-")],
    
    [eg.HSeparator()],

    # メインエリア：左に画像、右に操作パネル
    [
        # 左カラム：画像とページ送り
        eg.Column([
            [eg.Image(key="-IMAGE-", size=(400, 500), background_color="white")],
            # エラー原因だった justification を削除
            [eg.Button("◀ 前へ", key="-PREV-"), 
             eg.Text("0 / 0", key="-STATUS-", size=(15, 1)), 
             eg.Button("次へ ▶", key="-NEXT-")]
        ], vertical_alignment="top"),
        
        # 右カラム：操作パネル
        eg.Column(control_column, vertical_alignment="top")
    ],
    
    [eg.HSeparator()],
    [eg.Button("名前を付けて保存", key="-SAVE-", size=(20, 2)), eg.Button("終了", size=(10, 2))]
]

window = eg.Window("PDF Visual Editor V2", layout)

while True:
    event, values = window.read()
    
    if event in (eg.WINDOW_CLOSED, "終了"):
        break

    # --- 読込処理 ---
    if event == "-LOAD-":
        f_path = values["-FILE-"]
        if not f_path:
            eg.popup("ファイルを選択してください")
            continue
            
        if load_pdf_file(f_path):
            # 初期表示
            window["-IMAGE-"].update(data=get_preview_data(0))
            window["-STATUS-"].update(f"1 / {len(pdf_states)}")
            window["-MSG-"].update("読み込み完了")
        else:
            window["-MSG-"].update("読み込み失敗")

    # --- ページ送り ---
    if event in ("-PREV-", "-NEXT-") and doc_fitz:
        new_index = current_page_index + (-1 if event == "-PREV-" else 1)
        if 0 <= new_index < len(pdf_states):
            current_page_index = new_index
            window["-IMAGE-"].update(data=get_preview_data(current_page_index))
            
            # ステータス更新
            status = f"{current_page_index + 1} / {len(pdf_states)}"
            if pdf_states[current_page_index]['deleted']: status += " [削除済]"
            window["-STATUS-"].update(status)

    # --- 編集操作 (回転・削除) ---
    if event in ("-ROT_L-", "-ROT_R-", "-DEL-") and doc_fitz:
        
        target_indices = []
        
        # 対象範囲の特定
        if values["-SCOPE_CURRENT-"]:
            target_indices = [current_page_index]
        elif values["-SCOPE_ALL-"]:
            target_indices = list(range(len(pdf_states)))
        elif values["-SCOPE_RANGE-"]:
            target_indices = parse_page_range(values["-RANGE_INPUT-"], len(pdf_states))
            if not target_indices:
                window["-MSG-"].update("無効な範囲指定です")
                continue

        # 操作実行
        count = 0
        for idx in target_indices:
            if event == "-ROT_L-":
                pdf_states[idx]['rotation'] = (pdf_states[idx]['rotation'] - 90) % 360
            elif event == "-ROT_R-":
                pdf_states[idx]['rotation'] = (pdf_states[idx]['rotation'] + 90) % 360
            elif event == "-DEL-":
                pdf_states[idx]['deleted'] = not pdf_states[idx]['deleted']
            count += 1
            
        # 画面更新（現在のページが変わった場合のみ再描画）
        if current_page_index in target_indices:
            window["-IMAGE-"].update(data=get_preview_data(current_page_index))
            status = f"{current_page_index + 1} / {len(pdf_states)}"
            if pdf_states[current_page_index]['deleted']: status += " [削除済]"
            window["-STATUS-"].update(status)

        window["-MSG-"].update(f"{count}ページを変更")

    # --- 保存 ---
    if event == "-SAVE-" and doc_fitz:
        save_path = eg.popup_get_file("保存先", save_as=True, file_types=(("PDF Files", "*.pdf"),), default_extension=".pdf")
        if save_path:
            success, msg = save_current_pdf(save_path)
            eg.popup(msg)

window.close()