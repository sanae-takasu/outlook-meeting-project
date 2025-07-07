import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import pandas as pd
import datetime
import os
import threading
from outlookmeeting import get_meetings 

# プログレスバー付きポップアップウィンドウ
class ProgressPopup:
    def __init__(self, parent):
        # ポップアップウィンドウの作成
        self.popup = tk.Toplevel(parent)
        self.popup.title("Processing...")
        self.popup.transient(parent)
        self.popup.grab_set()

        # ウィンドウサイズを指定
        width = 300
        height = 100

        
        # 画面サイズを取得して中央位置を計算
        screen_width = self.popup.winfo_screenwidth()
        screen_height = self.popup.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)

        # ウィンドウの位置とサイズを設定
        self.popup.geometry(f"{width}x{height}+{x}+{y}")
        self.popup.transient(parent)
        self.popup.grab_set()


        # ラベルの表示
        label = ttk.Label(self.popup, text="Collectiong meeting data...")
        label.pack(pady=10)

        # determinateモードのプログレスバー
        self.progress_bar = ttk.Progressbar(self.popup, mode="determinate", maximum=100)
        self.progress_bar.pack(pady=10, padx=20, fill=tk.X)
        self.progress_bar["value"] = 0

    # 進捗率を更新する関数
    def update_progress(self, value):
        self.progress_bar["value"] = value
        self.popup.update_idletasks()

    # プログレスバーを停止してポップアップを閉じる
    def close(self):
        self.popup.destroy()

# メインアプリケーション
class OutlookMeetingsApp:
    def __init__(self, root):
        # メインウィンドウの設定
        self.root = root
        self.root.title("Outlook Meetings")

        # フレームの作成と配置
        self.frame = ttk.Frame(root, padding="10")
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        # フレーム内のすべての行に対して高さを設定
        for i in range(5):  # 10行分を設定（必要な行数に応じて変更）
            self.frame.grid_rowconfigure(i, minsize=27)  # 各行の高さを20pxに設定

        # 開始日と終了日の入力
        self.start_date_label = ttk.Label(self.frame, text="Start Date (YYYY-MM-DD):")
        self.start_date_label.grid(row=0, column=0, sticky=tk.W)
        self.start_date_entry = DateEntry(self.frame, date_pattern='yyyy-mm-dd')
        self.start_date_entry.grid(row=0, column=1, sticky=tk.W)

        self.end_date_label = ttk.Label(self.frame, text="End Date (YYYY-MM-DD):")
        self.end_date_label.grid(row=1, column=0, sticky=tk.W)
        self.end_date_entry = DateEntry(self.frame, date_pattern='yyyy-mm-dd')
        self.end_date_entry.grid(row=1, column=1, sticky=tk.W)

        # MeetingStatus選択チェックボックス（横並び＋左マージン）
        self.meeting_status_label = ttk.Label(self.frame, text="Meeting Status:")
        self.meeting_status_label.grid(row=2, column=0, sticky=tk.W)
        self.status_vars = {
            0: tk.IntVar(value=1),
            1: tk.IntVar(value=1),
            2: tk.IntVar(value=0),
            3: tk.IntVar(value=1)
        }
        ttk.Checkbutton(self.frame, text="0:Normal", variable=self.status_vars[0]).grid(row=2, column=1, padx=(0, 5), sticky=tk.W)
        ttk.Checkbutton(self.frame, text="1:Meeting", variable=self.status_vars[1]).grid(row=2, column=2, padx=(5, 5), sticky=tk.W)
        ttk.Checkbutton(self.frame, text="2:Canceled", variable=self.status_vars[2]).grid(row=2, column=3, padx=(5, 5), sticky=tk.W)
        ttk.Checkbutton(self.frame, text="3:Request", variable=self.status_vars[3]).grid(row=2, column=4, padx=(5, 5), sticky=tk.W)
        
        # カテゴリフィルターの入力（オプション）
        self.category_label = ttk.Label(self.frame, text="Category Filter:")
        self.category_label.grid(row=3, column=0, sticky=tk.W)
        self.category_entry = ttk.Entry(self.frame, width=80)
        self.category_entry.grid(row=3, column=1, columnspan=3,sticky=tk.W)
        self.exclude_var = tk.IntVar(value=0)
        self.exclude_checkbox = ttk.Checkbutton(self.frame, text="Exclude Category", variable=self.exclude_var)
        self.exclude_checkbox.grid(row=3, column=5, sticky=tk.W)

        # 出力フォルダ選択
        self.output_folder_label = ttk.Label(self.frame, text="Output Folder:")
        self.output_folder_label.grid(row=4, column=0, sticky=tk.W)
        self.output_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        self.output_folder_path = tk.StringVar(value=self.output_folder)
        self.output_folder_entry = ttk.Entry(self.frame, textvariable=self.output_folder_path, state='readonly', width=80 )
        self.output_folder_entry.grid(row=4, column=1, columnspan=4, sticky=tk.W)
        self.output_folder_button = ttk.Button(self.frame, text="Select Folder", command=self.select_output_folder)
        self.output_folder_button.grid(row=4, column=5, sticky=tk.W)

        # 実行ボタンの作成
        self.run_button = ttk.Button(self.frame, text="Run", command=self.run_analysis)
        self.run_button.grid(row=5, columnspan=6)

        # 結果表示用のTreeview
        self.tree = ttk.Treeview(self.frame, columns=("Month", "Subject", "Count", "Total Duration (minutes)", "Categories"), show="headings")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
        self.tree.grid(row=6, column=0, columnspan=6, sticky=(tk.W, tk.E, tk.N, tk.S))

        # スクロールバーの設定
        self.scrollbar = ttk.Scrollbar(self.frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=self.scrollbar.set)
        self.scrollbar.grid(row=6, column=6, sticky=(tk.N, tk.S))

    # 出力フォルダ選択処理
    def select_output_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder_path.set(folder_selected)
            self.output_folder = folder_selected

    # 実行処理（別スレッドで実行）
    def run_analysis(self):
        self.progress_popup = ProgressPopup(self.root)
        thread = threading.Thread(target=self._run_analysis_task)
        thread.start()

    # 実行処理の本体（バックグラウンド）
    def _run_analysis_task(self):
        try:
            start_date = datetime.datetime.strptime(self.start_date_entry.get(), "%Y-%m-%d")
            end_date = datetime.datetime.strptime(self.end_date_entry.get(), "%Y-%m-%d") + datetime.timedelta(days=1)
            meeting_types = [status for status, var in self.status_vars.items() if var.get() == 1]
            category_filter = self.category_entry.get().strip()
            exclude = bool(self.exclude_var.get())

            # 進捗率を更新するコールバック関数
            def progress_callback(value):
                self.root.after(0, self.progress_popup.update_progress, value)

            # 会議データ取得
            file_path = get_meetings(start_date, end_date, self.output_folder, meeting_types, progress_callback,category_filter, exclude)
            df = pd.read_excel(file_path, engine='openpyxl')

            self.root.after(0, self._update_treeview, df, file_path)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            self.root.after(0, self.progress_popup.close)

    # Treeviewの更新処理
    def _update_treeview(self, df, file_path):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=(row["Month"], row["Subject"], row["Count"], row["Total Duration (minutes)"], row["Categories"]))
        self.progress_popup.close()
        messagebox.showinfo("Complete", f"You have successfully saved the data to :\n{file_path}")

# アプリケーションの起動
if __name__ == "__main__":
    root = tk.Tk()
    app = OutlookMeetingsApp(root)
    root.mainloop()

