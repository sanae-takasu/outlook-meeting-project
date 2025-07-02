import tkinter as tk
from tkinter import ttk, filedialog
from tkcalendar import DateEntry
import pandas as pd
import datetime
from src.outlookmeeting import get_meetings  # 修正された関数をインポート
import os

class OutlookMeetingsApp:
    def __init__(self, root):
        # メインウィンドウの設定
        self.root = root
        self.root.title("Outlook Meetings")
        
        # フレームの作成と配置
        self.frame = ttk.Frame(root, padding="10")
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 期間入力ラベルとDateEntryの作成
        self.start_date_label = ttk.Label(self.frame, text="Start Date (YYYY-MM-DD):")
        self.start_date_label.grid(row=0, column=0, sticky=tk.W)
        self.start_date_entry = DateEntry(self.frame, date_pattern='yyyy-mm-dd')
        self.start_date_entry.grid(row=0, column=1, sticky=tk.W)
        
        self.end_date_label = ttk.Label(self.frame, text="End Date (YYYY-MM-DD):")
        self.end_date_label.grid(row=1, column=0, sticky=tk.W)
        self.end_date_entry = DateEntry(self.frame, date_pattern='yyyy-mm-dd')
        self.end_date_entry.grid(row=1, column=1, sticky=tk.W)
        
        # 出力先選択ボタンの作成
        self.output_folder_label = ttk.Label(self.frame, text="Output Folder:")
        self.output_folder_label.grid(row=2, column=0, sticky=tk.W)
        
        # デフォルトでダウンロードフォルダを設定
        self.output_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        
        self.output_folder_path = tk.StringVar(value=self.output_folder)
        self.output_folder_entry = ttk.Entry(self.frame, textvariable=self.output_folder_path, state='readonly', width=50)
        self.output_folder_entry.grid(row=2, column=1, sticky=tk.W)
        
        self.output_folder_button = ttk.Button(self.frame, text="Select Folder", command=self.select_output_folder)
        self.output_folder_button.grid(row=2, column=2, sticky=tk.W)
        
        # 実行ボタンの作成
        self.run_button = ttk.Button(self.frame, text="Run", command=self.run_analysis)
        self.run_button.grid(row=3, columnspan=3)
        
        # Treeviewウィジェットの作成と設定
        self.tree = ttk.Treeview(self.frame, columns=("Month", "Subject", "Count", "Total Duration (minutes)"), show="headings")
        self.tree.heading("Month", text="Month")
        self.tree.heading("Subject", text="Subject")
        self.tree.heading("Count", text="Count")
        self.tree.heading("Total Duration (minutes)", text="Total Duration (minutes)")
        self.tree.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), columnspan=3)
        
        # スクロールバーの作成と設定
        self.scrollbar = ttk.Scrollbar(self.frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=self.scrollbar.set)
        self.scrollbar.grid(row=4, column=3, sticky=(tk.N, tk.S))
        
    def select_output_folder(self):
        # 出力先フォルダを選択するダイアログを表示
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder_path.set(folder_selected)
            self.output_folder = folder_selected
    
    def run_analysis(self):
        # 入力された期間を取得
        start_date_str = self.start_date_entry.get()
        end_date_str = self.end_date_entry.get()
        
        try:
            start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d")
            end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d")
            
            # 出力先フォルダが選択されているか確認
            if hasattr(self, 'output_folder'):
                output_folder = self.output_folder
                
                # outlookmeeting.pyの関数を呼び出してデータを取得
                file_path = get_meetings(start_date, end_date, output_folder)
                df = pd.read_excel(file_path)
                
                # Treeviewをクリアして新しいデータを挿入
                for row in self.tree.get_children():
                    self.tree.delete(row)
                
                for _, row in df.iterrows():
                    self.tree.insert("", "end", values=(row["Month"], row["Subject"], row["Count"], row["Total Duration (minutes)"]))
            else:
                print("Output folder not selected.")
        
        except ValueError:
            print("Invalid date format. Please enter dates in YYYY-MM-DD format.")
    
if __name__ == "__main__":
    # アプリケーションの開始
    root = tk.Tk()
    app = OutlookMeetingsApp(root)
    root.mainloop()
