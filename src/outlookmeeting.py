import win32com.client
import datetime
from collections import defaultdict
import pandas as pd

def get_meetings(start_date, end_date, download_folder):
    # Outlookアプリケーションを初期化
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # カレンダーフォルダを取得
    calendar_folder = namespace.GetDefaultFolder(9)  # 9はカレンダーフォルダを指します

    # カレンダーアイテムを取得
    calendar_items = calendar_folder.Items

    # カレンダーアイテムを開始日時でソートし、繰り返しの会議を含むように設定
    calendar_items.Sort("[Start]")
    calendar_items.IncludeRecurrences = True

    # 会議の期間を定義
    restriction = "[Start] >= '" + start_date.strftime("%m/%d/%Y %H:%M %p") + "' AND [End] <= '" + end_date.strftime("%m/%d/%Y %H:%M %p") + "'"
    restricted_items = calendar_items.Restrict(restriction)

    # 各月ごとの会議の詳細を格納する辞書を初期化
    meetings_by_month = defaultdict(lambda: defaultdict(lambda: {"count": 0, "total_duration": datetime.timedelta()}))

    # 制限されたアイテムを反復処理して会議の詳細を抽出
    for item in restricted_items:
        if item.MeetingStatus in [1, 3]:  # 1は会議、3は会議リクエストを指します
            meeting_month = item.Start.strftime("%Y/%m")  # YYYY/MM形式にフォーマット
            meeting_subject = item.Subject
            meeting_duration = datetime.timedelta(minutes=item.Duration)
            
            meetings_by_month[meeting_month][meeting_subject]["count"] += 1
            meetings_by_month[meeting_month][meeting_subject]["total_duration"] += meeting_duration

    # エクセルファイルに書き込むためのデータフレームを作成
    data = []
    for month, meetings in meetings_by_month.items():
        for subject, details in meetings.items():
            data.append({
                "Month": month,  # ここでYYYY/MM形式の月を使用
                "Subject": subject,
                "Count": details["count"],
                "Total Duration (minutes)": details["total_duration"].total_seconds() / 60
            })

    df = pd.DataFrame(data)

    # 現在の日時を取得してファイル名に追加
    now = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    
    # エクセルファイルに書き込み（指定されたフォルダに保存）
    file_path = download_folder + f"\\outlook_meetings_{now}.xlsx"
    df.to_excel(file_path, index=False)

    return file_path
