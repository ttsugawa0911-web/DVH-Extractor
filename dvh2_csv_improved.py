# --- モジュールのインポート ---
# 必要な機能（部品）をプログラムに読み込みます。

import tkinter  # GUI（画面のウィンドウやボタンなど）を作成するための基本モジュール
from tkinter import ttk, filedialog, messagebox  # GUIのより高機能な部品やダイアログボックス
from pathlib import Path  # ファイルやフォルダのパスを簡単に、OSの違いを気にせず扱えるようにするモジュール
import re  # 正規表現（複雑なルールでの文字列検索・置換）を使うためのモジュール
import csv  # CSVファイルを読み書きするためのモジュール
import threading  # 時間のかかる処理をバックグラウンドで行い、アプリが固まるのを防ぐためのモジュール
from collections import defaultdict  # 辞書（データを名前付きで保存する箱）を便利に使うためのモジュール
import queue  # スレッド間で安全にデータをやり取りするためのモジュール
import json  # Pythonのデータ（辞書など）をファイルに保存・復元するためのモジュール

# --- アプリケーションの本体となるクラス ---
# アプリの設計図のようなものです。GUIの見た目や動作をすべてこの中に定義します。

class DvhConverterApp:
    # --- 初期化メソッド ---
    # アプリが起動したときに最初に実行される部分です。
    def __init__(self, root):
        # self: このクラス自身のことを指す印。メソッド内で作った変数をクラス全体で使えるようにする。
        # root: アプリケーションのメインウィンドウのこと。
        self.root = root
        
        # --- 設定ファイルのパスを定義 ---
        # Path.home()はユーザーのホームディレクトリ（例: C:\Users\YourName）を取得します。
        # .dvh_converter_config.json という名前のファイルパスを作成します。
        self.config_file = Path.home() / ".dvh_converter_config.json"
        
        # --- ウィンドウの基本設定 ---
        root.title("DVHデータ抽出ツール")  # ウィンドウのタイトルバーに表示されるテキスト
        root.geometry("600x400")  # ウィンドウの初期サイズ（幅x高さ）
        root.columnconfigure(0, weight=1)  # ウィンドウサイズを変更したときに、中の要素が追従するように設定
        root.rowconfigure(0, weight=1)

        # --- メインフレームの作成 ---
        # ウィジェット（ボタンなど）を配置するための土台となる領域を作成します。
        main_frm = ttk.Frame(root, padding=10)
        main_frm.grid(sticky=tkinter.NSEW)  # フレームをウィンドウ全体に広げる
        main_frm.columnconfigure(1, weight=1) # フレーム内の列がウィンドウサイズに追従するように設定

        # --- GUIと連動する特殊な変数 ---
        # これらの変数の中身が変わると、関連付けられたGUI部品の表示も自動で変わります。
        self.input_folder_path = tkinter.StringVar()  # 入力フォルダのパスを保存する文字列変数
        self.output_folder_path = tkinter.StringVar() # 出力フォルダのパスを保存する文字列変数
        self.is_patient_wise = tkinter.BooleanVar(value=True)  # 「患者ごと」チェックボックスの状態（ON/OFF）
        self.is_structure_wise = tkinter.BooleanVar(value=True) # 「構造体ごと」チェックボックスの状態（ON/OFF）
        
        # --- GUI部品（ウィジェット）の作成と配置 ---
        # ttk.Label: テキストラベル
        # ttk.Entry: 1行のテキスト入力ボックス
        # ttk.Button: クリックできるボタン
        # ttk.Combobox: ドロップダウンリスト
        # ttk.Checkbutton: チェックボックス
        # .grid()や.pack()で、作成したウィジェットを画面上のどこに置くか決めています。

        # 1. 入力フォルダ選択エリア
        ttk.Label(main_frm, text="1. DVHテキストファイルのあるフォルダを選択").grid(column=0, row=0, columnspan=3, sticky=tkinter.W, pady=(0, 5))
        input_folder_box = ttk.Entry(main_frm, textvariable=self.input_folder_path, width=60)
        input_folder_box.grid(column=0, row=1, columnspan=2, sticky=tkinter.EW, padx=(0, 5))
        ttk.Button(main_frm, text="参照...", command=self.ask_input_folder).grid(column=2, row=1)
        
        # 2. 出力先フォルダ選択エリア
        ttk.Label(main_frm, text="2. CSVファイルの保存先フォルダを選択").grid(column=0, row=2, columnspan=3, sticky=tkinter.W, pady=(10, 5))
        output_folder_box = ttk.Entry(main_frm, textvariable=self.output_folder_path, width=60)
        output_folder_box.grid(column=0, row=3, columnspan=2, sticky=tkinter.EW, padx=(0, 5))
        ttk.Button(main_frm, text="参照...", command=self.ask_output_folder).grid(column=2, row=3)

        # 3. オプション設定エリア
        ttk.Label(main_frm, text="3. オプションを設定").grid(column=0, row=4, columnspan=3, sticky=tkinter.W, pady=(10, 5))
        options_frame = ttk.Frame(main_frm)
        options_frame.grid(column=0, row=5, columnspan=3, sticky=tkinter.W)

        ttk.Label(options_frame, text="サンプリング間隔:").pack(side=tkinter.LEFT, padx=(0, 5))
        self.order_comb = ttk.Combobox(options_frame, values=[0.1, 0.5, 1, 5, 10], state='readonly', width=5)
        self.order_comb.current(1) # 初期値をリストの2番目(0.5)に設定
        self.order_comb.pack(side=tkinter.LEFT, padx=5)

        ttk.Label(options_frame, text="線量種別:").pack(side=tkinter.LEFT, padx=(20, 5))
        self.type_comb = ttk.Combobox(options_frame, values=["%", "Gy"], state='readonly', width=5)
        self.type_comb.current(0) # 初期値をリストの1番目(%)に設定
        self.type_comb.pack(side=tkinter.LEFT, padx=5)
        
        # 4. 出力形式選択エリア
        ttk.Label(main_frm, text="4. 出力形式を選択（複数選択可）").grid(column=0, row=6, columnspan=3, sticky=tkinter.W, pady=(10, 5))
        output_type_frame = ttk.Frame(main_frm)
        output_type_frame.grid(column=0, row=7, columnspan=3, sticky=tkinter.W)
        ttk.Checkbutton(output_type_frame, text="患者ごとにCSVを作成", variable=self.is_patient_wise).pack(side=tkinter.LEFT)
        ttk.Checkbutton(output_type_frame, text="構造体ごとにCSVを作成", variable=self.is_structure_wise).pack(side=tkinter.LEFT, padx=20)

        # 5. 実行ボタンとプログレスバー
        self.app_btn = ttk.Button(main_frm, text="実行", command=self.run_conversion)
        self.app_btn.grid(column=0, row=8, columnspan=3, pady=20)

        self.progress_var = tkinter.DoubleVar() # プログレスバーの進捗(0-100)を保存する数値変数
        self.progress_bar = ttk.Progressbar(main_frm, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(column=0, row=9, columnspan=3, sticky=tkinter.EW, pady=10)
        self.status_label = ttk.Label(main_frm, text="準備完了")
        self.status_label.grid(column=0, row=10, columnspan=3, sticky=tkinter.W)
        
        # --- スレッド間通信の準備 ---
        # バックグラウンド処理からGUIへ安全にメッセージを渡すための「パイプ」のようなものを作成
        self.thread_queue = queue.Queue()
        # このパイプを定期的に監視する仕組みを開始
        self.process_queue()
        
        # --- 起動時に前回使ったフォルダを読み込む ---
        self.load_settings()

    # --- 設定をファイルに保存するメソッド ---
    def save_settings(self):
        # 現在の入力・出力フォルダのパスを辞書にまとめる
        settings = {
            'input_folder': self.input_folder_path.get(),
            'output_folder': self.output_folder_path.get()
        }
        try:
            # 設定ファイルを書き込みモードで開く
            with open(self.config_file, 'w') as f:
                # 辞書の内容をJSON形式でファイルに書き込む
                json.dump(settings, f, indent=4)
        except IOError:
            # もしファイル書き込みに失敗しても、エラーメッセージを出すだけでアプリは止めない
            print("設定ファイルの書き込みに失敗しました。")

    # --- 設定をファイルから読み込むメソッド ---
    def load_settings(self):
        try:
            # 設定ファイルが存在するか確認
            if self.config_file.exists():
                # ファイルを読み込みモードで開く
                with open(self.config_file, 'r') as f:
                    # JSONファイルからデータを読み込んで辞書に戻す
                    settings = json.load(f)
                    # 読み込んだパスをGUIの変数にセットする
                    self.input_folder_path.set(settings.get('input_folder', ''))
                    self.output_folder_path.set(settings.get('output_folder', ''))
        except (json.JSONDecodeError, IOError):
            # ファイルが存在しない、壊れている、などの場合は何もしない
            print("設定ファイルの読み込みに失敗しました。")
            
    # --- キューを監視するメソッド ---
    # バックグラウンドからのメッセージを受け取って処理します
    def process_queue(self):
        try:
            # キューからメッセージを1つ取り出そうと試みる（メッセージがなければエラーになる）
            message_type, message_content = self.thread_queue.get(block=False)
            
            # メッセージの種類に応じて、完了またはエラーのダイアログを表示
            if message_type == 'INFO':
                messagebox.showinfo("完了", message_content)
            elif message_type == 'ERROR':
                messagebox.showerror("エラー", message_content)
        except queue.Empty:
            # キューが空っぽだった場合は何もしない
            pass
        finally:
            # 100ミリ秒後に、再度このメソッド自身を呼び出す。これを繰り返して常にキューを監視する。
            self.root.after(100, self.process_queue)

    # --- 「実行」ボタンが押されたときのメソッド ---
    def run_conversion(self):
        # 入力されたパスを取得
        input_dir = self.input_folder_path.get()
        output_dir = self.output_folder_path.get()

        # --- 入力チェック ---
        # フォルダが指定されているかなどを確認し、不備があればエラーメッセージを表示して処理を中断
        if not input_dir or not output_dir:
            messagebox.showerror("エラー", "入力フォルダと出力先フォルダの両方を選択してください。")
            return # returnでこのメソッドの処理をここで終了する
        if not Path(input_dir).is_dir() or not Path(output_dir).is_dir():
            messagebox.showerror("エラー", "指定されたパスが存在しないか、フォルダではありません。")
            return
        if not self.is_patient_wise.get() and not self.is_structure_wise.get():
            messagebox.showerror("エラー", "少なくとも1つの出力形式を選択してください。")
            return

        # --- 処理開始の準備 ---
        self.app_btn.config(state="disabled")  # 処理中にボタンを連打できないように無効化
        self.status_label.config(text="処理中...") # ステータス表示を変更
        self.progress_var.set(0) # プログレスバーをリセット

        # --- バックグラウンド処理の開始 ---
        # conversion_threadメソッドを、メインのGUIとは別の「スレッド」で実行する
        # これにより、重い処理中もGUIが固まらなくなる
        thread = threading.Thread(target=self.conversion_thread, args=(input_dir, output_dir))
        thread.start() # スレッドを開始
        
    # --- バックグラウンドで実行されるメイン処理 ---
    def conversion_thread(self, input_dir, output_dir):
        # try...except...finally構文:
        # tryブロック内の処理を実行し、もしエラー(Exception)が起きたらexceptブロックに飛ぶ。
        # エラーがあってもなくても、最後に必ずfinallyブロックが実行される。
        try:
            # GUIから選択されたオプションの値を取得
            d_interval = float(self.order_comb.get())
            dose_type = self.type_comb.get()
            dose_col_index = 0 if dose_type == "%" else 1

            # 1. 全ファイルの解析
            self.status_label.config(text="ファイルを解析中...")
            all_patients_data, txt_files = self.parse_folder(Path(input_dir), d_interval, dose_col_index)
            
            if not all_patients_data:
                raise ValueError("処理対象のtxtファイルが見つからないか、データがありません。")

            # 2. CSVファイルの作成
            total_tasks = self.is_patient_wise.get() + self.is_structure_wise.get()
            task_count = 0

            # 「患者ごと」がチェックされていれば、そのCSV作成関数を呼ぶ
            if self.is_patient_wise.get():
                task_count += 1
                self.status_label.config(text=f"({task_count}/{total_tasks}) 患者別CSVを作成中...")
                self.write_patient_csvs(all_patients_data, Path(output_dir), dose_type, self.progress_var, 50.0 / total_tasks)

            # 「構造体ごと」がチェックされていれば、そのCSV作成関数を呼ぶ
            if self.is_structure_wise.get():
                task_count += 1
                self.status_label.config(text=f"({task_count}/{total_tasks}) 構造体別CSVを作成中...")
                self.write_structure_csvs(all_patients_data, Path(output_dir), dose_type, self.progress_var, 50.0 / total_tasks * task_count)

            # --- 処理成功 ---
            # 次回起動のために、今回使ったフォルダ設定を保存する
            self.save_settings()
            # 完了メッセージをキューに入れる（GUIスレッドがこれを拾ってダイアログを表示する）
            self.thread_queue.put(('INFO', '処理が完了しました。'))

        except Exception as e:
            # --- 処理失敗 ---
            # エラーメッセージをキューに入れる
            self.thread_queue.put(('ERROR', f"処理中にエラーが発生しました:\n{e}"))
        finally:
            # --- 処理終了後 ---
            # 成功しても失敗しても、ボタンを再度有効化し、ステータス表示を更新する
            self.app_btn.config(state="normal")
            self.status_label.config(text="完了")
            self.progress_var.set(100)

    # --- 「参照...」ボタンのメソッド ---
    def ask_input_folder(self):
        # フォルダ選択ダイアログを開く
        path = filedialog.askdirectory()
        if path: # フォルダが選択された場合（キャンセルされなかった場合）
            self.input_folder_path.set(path) # パスをGUI変数にセットする

    def ask_output_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.output_folder_path.set(path)
            
    # --- フォルダ内の全テキストファイルを解析するメソッド ---
    def parse_folder(self, input_dir, d_interval, dose_col_index):
        all_patients_data = {} # 全患者のデータを保存するからの辞書
        txt_files = list(input_dir.glob('*.txt')) # フォルダ内の.txtファイルをリストアップ
        
        if not txt_files: # ファイルが1つもなければ空のデータを返す
            return {}, []
            
        # ファイルリストをループ処理
        for i, file_path in enumerate(txt_files):
            # 1ファイルずつ解析するメソッドを呼び出す
            patient_data = self.parse_dvh_file(file_path, d_interval, dose_col_index)
            if patient_data and 'Patient ID' in patient_data:
                # 解析結果を、患者IDをキーとして辞書に保存
                all_patients_data[patient_data['Patient ID']] = patient_data
            # プログレスバーを進める
            self.progress_var.set((i + 1) / len(txt_files) * 50)
        
        return all_patients_data, txt_files

    # --- 1つのテキストファイルを解析するメソッド ---
    def parse_dvh_file(self, file_path, d_interval, dose_col_index):
        patient_data = {'structures': {}} # 1患者分のデータを保存する辞書
        current_structure = None # 現在解析中の構造体名を一時的に保存する変数

        # ファイルを開いて1行ずつ読み込む
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            for line in f:
                line = line.strip() # 行の先頭と末尾の空白を削除
                if not line: # 空行ならスキップ
                    continue
                
                # 正規表現で行を「キー」と「値」に分割（例: "Patient Name : hoge" -> "Patient Name", "hoge"）
                parts = re.split(r'\s*:\s*|\s{2,}', line, 1)
                key = parts[0].strip()
                val = parts[1].strip() if len(parts) > 1 else ''

                # キーの種類に応じて、データを辞書に格納していく
                if key == 'Patient Name': patient_data['Patient Name'] = val
                elif key == 'Patient ID': patient_data['Patient ID'] = val
                elif key == 'Prescribed dose [Gy]': patient_data['Prescribed dose [Gy]'] = val
                elif key == 'Structure':
                    current_structure = val # 構造体セクションが始まったことを記録
                    patient_data['structures'][current_structure] = {'dvh': {}}
                elif current_structure: # 構造体セクションの中の行を処理
                    if key in ['Volume [cm³]', 'Min Dose [%]', 'Max Dose [%]', 'Mean Dose [%]']:
                        patient_data['structures'][current_structure][key] = val
                    elif re.match(r'^[0-9.]+$', key): # 行が数字で始まっていたらDVHデータと判断
                        dvh_values = re.split(r'\s+', line.strip())
                        try:
                            dose = float(dvh_values[dose_col_index])
                            volume = float(dvh_values[-1])
                            # サンプリング間隔に合致するデータのみを抽出
                            if abs(dose % d_interval) < 1e-9 or abs(dose % d_interval - d_interval) < 1e-9:
                                patient_data['structures'][current_structure]['dvh'][dose] = volume
                        except (ValueError, IndexError):
                            # 数値に変換できない行などは無視して次に進む
                            pass
        return patient_data

    # --- 患者別CSV（ワイド形式）を作成するメソッド ---
    def write_patient_csvs(self, all_patients_data, output_dir, dose_type, progress_var, progress_offset):
        dose_unit = "[%]" if dose_type == "%" else "[Gy]"

        # 患者データごとにループ
        for patient_id, data in all_patients_data.items():
            filename = output_dir / f"{patient_id}.csv"
            
            structures = data.get('structures', {})
            if not structures: # 構造体データがなければスキップ
                continue

            # --- CSV書き込みのためのデータ準備 ---
            all_doses = set() # 全構造体の線量ポイントを重複なく集めるためのセット
            structure_names = sorted(structures.keys()) # 構造体名をアルファベット順に並べる

            # 全構造体をループして、すべての線量ポイントを集める
            for str_name in structure_names:
                all_doses.update(structures[str_name].get('dvh', {}).keys())

            # 集めた線量ポイントを小さい順に並べる
            sorted_doses = sorted(list(all_doses))

            # --- CSVファイルへの書き込み開始 ---
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)

                # 1. 患者情報の書き出し
                writer.writerow(['Patient Name', data.get('Patient Name', '')])
                writer.writerow(['Patient ID', data.get('Patient ID', '')])
                writer.writerow(['Prescribed dose [Gy]', data.get('Prescribed dose [Gy]', '')])
                writer.writerow([]) # 空行

                # 2. 構造体の統計情報（サマリー）の書き出し
                stats_header = ['Structure Name', 'Volume [cm³]', 'Min Dose [%]', 'Max Dose [%]', 'Mean Dose [%]']
                writer.writerow(stats_header)
                for str_name in structure_names:
                    str_data = structures[str_name]
                    stats_row = [
                        str_name,
                        str_data.get('Volume [cm³]', ''),
                        str_data.get('Min Dose [%]', ''),
                        str_data.get('Max Dose [%]', ''),
                        str_data.get('Mean Dose [%]', '')
                    ]
                    writer.writerow(stats_row)
                writer.writerow([]) # 空行

                # 3. DVHデータ本体（ワイド形式）の書き出し
                # ヘッダー行を作成 (例: Dose [%], BLAD_W_Volume [%], PTV_PSV_Volume [%], ...)
                dvh_header = [f'Dose {dose_unit}']
                for str_name in structure_names:
                    dvh_header.append(f'{str_name}_Volume [%]')
                writer.writerow(dvh_header)

                # 各線量ポイントについて1行ずつデータを作成
                for dose in sorted_doses:
                    row = [dose] # 行の先頭は線量値
                    # 各構造体について、その線量での体積データを取得して行に追加
                    for str_name in structure_names:
                        # .get(dose, '')は、もしその線量のデータがなければ空文字を入れる、という処理
                        volume = structures[str_name].get('dvh', {}).get(dose, '')
                        row.append(volume)
                    writer.writerow(row)

    # --- 構造体別CSVを作成するメソッド ---
    def write_structure_csvs(self, all_patients_data, output_dir, dose_type, progress_var, progress_offset):
        dose_unit = "[%]" if dose_type == "%" else "[Gy]"
        
        # defaultdictは、キーが存在しないときに自動で初期値（ここでは空の辞書）を作ってくれる便利な辞書
        structure_data = defaultdict(lambda: defaultdict(dict))
        all_doses = set()
        patient_ids = list(all_patients_data.keys())

        # 全患者データをループして、構造体ごとにデータを再整理する
        for pid, data in all_patients_data.items():
            for str_name, str_data in data['structures'].items():
                if 'dvh' in str_data:
                    for dose, volume in str_data['dvh'].items():
                        # (構造体名 -> 患者ID -> 線量) の順で体積データを保存
                        structure_data[str_name][pid][dose] = volume
                        all_doses.add(dose)

        sorted_doses = sorted(list(all_doses))

        # 再整理したデータを使って、構造体ごとに1つのCSVファイルを作成
        for str_name, patients in structure_data.items():
            # ファイル名に使えない文字を"_"に置き換える
            safe_str_name = re.sub(r'[\\/:*?"<>|]', '_', str_name)
            filename = output_dir / f"structure_{safe_str_name}.csv"
            
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # ヘッダー行 (Dose, 患者ID1, 患者ID2, ...)
                header = [f'Dose {dose_unit}'] + patient_ids
                writer.writerow(header)
                
                # 各線量ポイントで1行作成
                for dose in sorted_doses:
                    row = [dose]
                    # 各患者について、その線量での体積データを取得して行に追加
                    for pid in patient_ids:
                        volume = patients.get(pid, {}).get(dose, '')
                        row.append(volume)
                    writer.writerow(row)

# --- プログラムの実行開始点 ---
# このファイルが直接実行された場合にのみ、以下のコードが実行されます。
if __name__ == "__main__":
    main_win = tkinter.Tk()          # メインウィンドウを作成
    app = DvhConverterApp(main_win)  # 作成したDvhConverterAppクラスのインスタンス（実体）を作る
    main_win.mainloop()              # GUIアプリケーションを開始し、ユーザーの操作を待ち受ける