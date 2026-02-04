# app/ui/gui.py
import datetime
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from tkcalendar import DateEntry
import pandas as pd

from app.services.outlook_service import get_meetings


# プログレスバー付きポップアップウィンドウ
class ProgressPopup:
    def __init__(self, parent):
        # ポップアップウィンドウの作成
        self.popup = tk.Toplevel(parent)
        self.popup.title("Processing...")
        self.popup.transient(parent)   # 親に関連づけ
        self.popup.grab_set()          # モーダルに

        # ウィンドウサイズ
        width = 300
        height = 100

        # 画面サイズを取得して中央位置を計算
        screen_width = self.popup.winfo_screenwidth()
        screen_height = self.popup.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)

        # ウィンドウ位置とサイズ
        self.popup.geometry(f"{width}x{height}+{x}+{y}")

        # ラベル
        label = ttk.Label(self.popup, text="Collecting meeting data...")
        label.pack(pady=10)

        # determinateモードのプログレスバー
        self.progress_bar = ttk.Progressbar(self.popup, mode="determinate", maximum=100)
        self.progress_bar.pack(pady=10, padx=20, fill=tk.X)
        self.progress_bar["value"] = 0

    # 進捗率を更新
    def update_progress(self, value):
        self.progress_bar["value"] = value
        self.popup.update_idletasks()

    # 閉じる
    def close(self):
        try:
            self.popup.grab_release()
        except Exception:
            pass
        self.popup.destroy()


# メインアプリケーション
class OutlookMeetingsApp:
    def __init__(self, root):
        # メインウィンドウ
        self.root = root
        self.root.title("Outlook Meetings")
        self.df = None
        self.last_file_path = None

        # フレーム
        self.frame = ttk.Frame(root, padding="10")
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 行高さ調整（見た目安定化）
        for i in range(8):
            self.frame.grid_rowconfigure(i, minsize=27)

        # Dates
        ttk.Label(self.frame, text="Start Date (YYYY-MM-DD):").grid(row=0, column=0, sticky=tk.W)
        self.start_date_entry = DateEntry(self.frame, date_pattern='yyyy-mm-dd')
        self.start_date_entry.grid(row=0, column=1, sticky=tk.W)

        ttk.Label(self.frame, text="End Date (YYYY-MM-DD):").grid(row=1, column=0, sticky=tk.W)
        self.end_date_entry = DateEntry(self.frame, date_pattern='yyyy-mm-dd')
        self.end_date_entry.grid(row=1, column=1, sticky=tk.W)

        # MeetingStatus
        ttk.Label(self.frame, text="Meeting Status:").grid(row=2, column=0, sticky=tk.W)
        self.status_vars = {
            0: tk.IntVar(value=1),  # olNonMeeting
            1: tk.IntVar(value=1),  # olMeeting
            2: tk.IntVar(value=0),  # olMeetingCancelled
            3: tk.IntVar(value=1),  # olMeetingReceived
        }
        ttk.Checkbutton(self.frame, text="0:Normal",   variable=self.status_vars[0]).grid(row=2, column=1, padx=(0, 5), sticky=tk.W)
        ttk.Checkbutton(self.frame, text="1:Meeting",  variable=self.status_vars[1]).grid(row=2, column=2, padx=(5, 5), sticky=tk.W)
        ttk.Checkbutton(self.frame, text="2:Canceled", variable=self.status_vars[2]).grid(row=2, column=3, padx=(5, 5), sticky=tk.W)
        ttk.Checkbutton(self.frame, text="3:Request",  variable=self.status_vars[3]).grid(row=2, column=4, padx=(5, 5), sticky=tk.W)

        # Category filter
        ttk.Label(self.frame, text="Category Filter:").grid(row=3, column=0, sticky=tk.W)
        self.category_entry = ttk.Entry(self.frame, width=80)
        self.category_entry.grid(row=3, column=1, columnspan=3, sticky=tk.W)

        self.exclude_var = tk.IntVar(value=0)
        ttk.Checkbutton(self.frame, text="Exclude Category", variable=self.exclude_var).grid(row=3, column=5, sticky=tk.W)

        # Output folder
        ttk.Label(self.frame, text="Output Folder:").grid(row=4, column=0, sticky=tk.W)
        self.output_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        self.output_folder_path = tk.StringVar(value=self.output_folder)
        ttk.Entry(self.frame, textvariable=self.output_folder_path, state='readonly', width=80).grid(row=4, column=1, columnspan=4, sticky=tk.W)
        ttk.Button(self.frame, text="Select Folder", command=self.select_output_folder).grid(row=4, column=5, sticky=tk.W)

        # Display unit
        self.display_format = tk.StringVar(value="minutes")
        ttk.Label(self.frame, text="Display format:").grid(row=5, column=0, sticky=tk.W)
        ttk.Radiobutton(self.frame, text="Minutes", variable=self.display_format, value="minutes", command=self.on_display_change).grid(row=5, column=1, sticky=tk.W)
        ttk.Radiobutton(self.frame, text="Hours", variable=self.display_format, value="hours", command=self.on_display_change).grid(row=5, column=2, sticky=tk.W)
        ttk.Radiobutton(self.frame, text="Days", variable=self.display_format, value="days", command=self.on_display_change).grid(row=5, column=3, sticky=tk.W)

        # Run
        ttk.Button(self.frame, text="Run", command=self.run_analysis).grid(row=6, columnspan=6)

        # Tree
        self.tree = ttk.Treeview(
            self.frame,
            columns=("Month", "Subject Categories", "Subject", "Count", "Total Duration (minutes)", "Categories"),
            show="headings"
        )
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
        self.tree.grid(row=7, column=0, columnspan=6, sticky=(tk.W, tk.E, tk.N, tk.S))

        # スクロールバー
        self.scrollbar = ttk.Scrollbar(self.frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.grid(row=7, column=6, sticky=(tk.N, tk.S))

    # 出力フォルダ選択
    def select_output_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder_path.set(folder_selected)
            self.output_folder = folder_selected

    # 実行（別スレッドで）
    def run_analysis(self):
        self.progress_popup = ProgressPopup(self.root)
        thread = threading.Thread(target=self._run_analysis_task, daemon=True)
        thread.start()

    # ラジオ切替時に、既存データがあれば表示のみ切替
    def on_display_change(self):
        if self.df is not None and self.last_file_path is not None:
            display = self.display_format.get()  # 'minutes' / 'hours' / 'days'
            self._update_treeview(self.df, self.last_file_path, display,False)

    # バックグラウンド処理本体
    def _run_analysis_task(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            start_date = datetime.datetime.strptime(self.start_date_entry.get(), "%Y-%m-%d")
            end_date = datetime.datetime.strptime(self.end_date_entry.get(), "%Y-%m-%d") + datetime.timedelta(days=1)

            meeting_types = [status for status, var in self.status_vars.items() if var.get() == 1]
            category_filter = self.category_entry.get().strip()
            exclude = bool(self.exclude_var.get())
            display = self.display_format.get()

            # 進捗更新コールバック（UIスレッドで実行）
            def progress_callback(value):
                self.root.after(0, self.progress_popup.update_progress, value)

            # 会議データ取得
            file_path = get_meetings(
                start_date=start_date,
                end_date=end_date,
                download_folder=self.output_folder,
                meeting_types=meeting_types,
                progress_callback=progress_callback,
                category_filter=category_filter,
                exclude=exclude
            )

            # 読み込み
            df = pd.read_excel(file_path, engine='openpyxl')

            # 最新を保持
            self.df = df
            self.last_file_path = file_path
            
            # Treeview 更新（display を渡す）
            self.root.after(0, self._update_treeview, df, file_path, display)

        except Exception as e:
            # エラー表示とポップアップクローズ
            def show_err():
                try:
                    messagebox.showerror("Error", str(e))
                finally:
                    self.progress_popup.close()
            self.root.after(0, show_err)

    def _update_treeview(self, df, file_path, display="minutes",msg=True):
        # カラム名とラベルを選択
        if display == "minutes":
            col = "Total Duration (minutes)"
            label = "Total Duration (minutes)"
        elif display == "hours":
            col = "Total Duration (hours)"
            label = "Total Duration (hours)"
        else:
            col = "Total Duration (days)"
            label = "Total Duration (days)"

        columns = ("Month", "Subject Categories", "Subject", "Count", col, "Categories")
        self.tree["columns"] = columns
        for c in columns:
            self.tree.heading(c, text=c if c != col else label)

        # 既存行クリア
        for row_id in self.tree.get_children():
            self.tree.delete(row_id)

        # データインサート
        for _, row in df.iterrows():
            self.tree.insert(
                "",
                "end",
                values=(
                    row.get("Month", ""),
                    row.get("Subject Categories", ""),
                    row.get("Subject", ""),
                    row.get("Count", 0),
                    row.get(col, 0),
                    row.get("Categories", "")
                )
            )

        # 完了
        self.progress_popup.close()
        if msg:
            messagebox.showinfo("Complete", f"You have successfully saved the data to :\n{file_path}")


def launch_app():
    root = tk.Tk()
    app = OutlookMeetingsApp(root)
    root.mainloop()
