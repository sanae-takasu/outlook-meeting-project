import win32com.client
import datetime
from collections import defaultdict
import pandas as pd
import os
import time

def get_meetings(start_date, end_date, download_folder, meeting_types,progress_callback,category_filter,exclude):
    
    import pythoncom
    pythoncom.CoInitialize() 


    # Outlookアプリケーションを初期化
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # カレンダーフォルダを取得（9はカレンダーを指す定数）
    calendar_folder = namespace.GetDefaultFolder(9)

    # カレンダーアイテムを取得
    calendar_items = calendar_folder.Items

    # カレンダーアイテムを開始日時でソートし、繰り返しの会議を含むように設定
    calendar_items.Sort("[Start]")
    calendar_items.IncludeRecurrences = True

    # 終了日を含めるために1日加算（GUI側でも加算しているが、念のため）
    adjusted_end_date = end_date

    # 会議の期間を定義（Outlookのフィルター形式に合わせる）
    restriction = "[Start] >= '" + start_date.strftime("%m/%d/%Y %H:%M %p") + "' AND [End] <= '" + adjusted_end_date.strftime("%m/%d/%Y %H:%M %p") + "'"
    restricted_items = calendar_items.Restrict(restriction)

    # 各月ごとの会議の詳細を格納する辞書を初期化
    meetings_by_month = defaultdict(lambda: defaultdict(lambda: {"count": 0, "total_duration": datetime.timedelta()}))

    # アイテム数を取得（進捗率計算用）
    
    restricted_items_list = []
    for item in restricted_items:
        restricted_items_list.append(item)


    total_items = len(restricted_items_list)
    processed_items = 0

    # 制限されたアイテムを反復処理して会議の詳細を抽出
    for item in restricted_items:
        try:
  
            # MeetingStatusが指定された種類に含まれているか確認
            if item.MeetingStatus in meeting_types:
                meeting_month = item.Start.strftime("%Y/%m")  # YYYY/MM形式にフォーマット
                meeting_subject = item.Subject
                meeting_duration = datetime.timedelta(minutes=item.Duration)
                meeting_categories = item.Categories if item.Categories else "None"  # カテゴリを取得（未設定の場合は "None"）
                #　カテゴリーリストに分割
                category_list = [cat.strip() for cat in category_filter.split(",") if cat.strip()]

                if category_list:
                    matched = any(cat in meeting_categories for cat in category_list)
                    if exclude and matched:
                        continue  # 除外対象ならスキップ
                    elif not exclude and not matched:
                        continue  # 絞り込み対象でなければスキップ
                
                
                # 件名に「:」または「：」がある場合、その前で分類
                if ":" in meeting_subject:
                    meeting_subject_categories = meeting_subject.split(":")[0].strip()
                elif "：" in meeting_subject:
                    meeting_subject_categories = meeting_subject.split("：")[0].strip()
                else:
                    meeting_subject_categories = "None"

                # 集計データに追加
                meetings_by_month[meeting_month][meeting_subject]["count"] += 1
                meetings_by_month[meeting_month][meeting_subject]["total_duration"] += meeting_duration
                meetings_by_month[meeting_month][meeting_subject]["categories"] = meeting_categories  # カテゴリを追加
                meetings_by_month[meeting_month][meeting_subject]["meeting_subject_categories"] = meeting_subject_categories # 件名カテゴリを追加

        except Exception as e:
            # Outlookアイテムにアクセスできない場合などの例外処理
            print(f"Error processing item: {e}")

        # 進捗率を更新（callbackが指定されていれば呼び出す）
        processed_items += 1
        if progress_callback and total_items > 0:
            progress = int((processed_items / total_items) * 100)
            progress_callback(progress)


    # エクセルファイルに書き込むためのデータフレームを作成
    data = []
    for month, meetings in meetings_by_month.items():
        for subject, details in meetings.items():
            total_minutes = details["total_duration"].total_seconds() / 60
            total_hours = total_minutes / 60
            total_days = total_hours / 7.75  # 1日=7.75時間換算
            data.append({
                "Month": month,
                "Subject Categories": details["meeting_subject_categories"],
                "Subject": subject,
                "Count": details["count"],
                "Total Duration (minutes)": round(total_minutes, 2),
                "Total Duration (hours)": round(total_hours, 2),
                "Total Duration (days)": round(total_days, 2),
                "Categories": details.get("categories", "None")
            })



    df = pd.DataFrame(data)

    # 現在の日時を取得してファイル名に追加
    now = datetime.datetime.now().strftime("%Y%m%d%H%M%S")

    # ファイルパスを安全に結合（OS依存を避ける）
    file_path = os.path.join(download_folder, f"outlook_meetings_{now}.xlsx")

    # エクセルファイルに書き込み（指定されたフォルダに保存）
    df.to_excel(file_path, index=False)

    return file_path

