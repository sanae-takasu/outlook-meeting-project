# app/services/outlook_service.py
import datetime
from collections import defaultdict
import os
import pandas as pd

def get_meetings(start_date, end_date, download_folder, meeting_types, progress_callback=None,
                 category_filter="", exclude=False):
    """
    Outlook カレンダーの予定を集計し、Excel へ出力してそのファイルパスを返します。

    Args:
        start_date (datetime): 取得開始日時（含む）
        end_date (datetime):   取得終了日時（含む想定。GUI 側で +1 日済み）
        download_folder (str): 出力フォルダ
        meeting_types (list[int]): Outlook の MeetingStatus を表す数値（例: [0,1,3]）
        progress_callback (callable|None): 進捗コールバック。0-100 を受け取る
        category_filter (str): カンマ区切りのカテゴリ指定
        exclude (bool): True の場合はカテゴリ一致を除外、False の場合は一致のみ対象

    Returns:
        str: 出力した Excel ファイルパス
    """
    import pythoncom
    import win32com.client

    # COM をこのスレッドで初期化
    pythoncom.CoInitialize()

    # Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar_folder = namespace.GetDefaultFolder(9)  # 9=olFolderCalendar
    calendar_items = calendar_folder.Items

    # 開始日時でソート + 繰り返し会議を展開
    calendar_items.Sort("[Start]")
    calendar_items.IncludeRecurrences = True

    # Restrict フィルターを作成（※ 演算子は >= / <=）
    restriction = (
        "[Start] >= '{start}' AND [End] <= '{end}'"
        .format(
            start=start_date.strftime("%m/%d/%Y %I:%M %p"),
            end=end_date.strftime("%m/%d/%Y %I:%M %p")
        )
    )
    restricted_items = calendar_items.Restrict(restriction)

    # 一旦リスト化（Items は一度きりの列挙になりがち）
    restricted_items_list = [item for item in restricted_items]
    total_items = len(restricted_items_list)
    processed = 0

    meetings_by_month = defaultdict(lambda: defaultdict(lambda: {
        "count": 0,
        "total_duration": datetime.timedelta(),
        "categories": "None",
        "meeting_subject_categories": "None",
    }))

    # カテゴリ指定をリスト化
    category_list = [cat.strip() for cat in category_filter.split(",") if cat.strip()]

    for item in restricted_items_list:
        try:
            # MeetingStatus フィルタ
            if item.MeetingStatus not in meeting_types:
                continue

            meeting_month = item.Start.strftime("%Y/%m")
            meeting_subject = item.Subject or ""
            meeting_duration = datetime.timedelta(minutes=item.Duration)
            meeting_categories = item.Categories if item.Categories else "None"

            # カテゴリフィルタ（除外 or 絞り込み）
            if category_list:
                matched = any(cat in meeting_categories for cat in category_list)
                if exclude and matched:
                    continue
                if not exclude and not matched:
                    continue

            # 件名の「: / ：」より前をカテゴリとして抽出
            if ":" in meeting_subject:
                meeting_subject_categories = meeting_subject.split(":")[0].strip()
            elif "：" in meeting_subject:
                meeting_subject_categories = meeting_subject.split("：")[0].strip()
            else:
                meeting_subject_categories = "None"

            # 集計
            agg = meetings_by_month[meeting_month][meeting_subject]
            agg["count"] += 1
            agg["total_duration"] += meeting_duration
            agg["categories"] = meeting_categories
            agg["meeting_subject_categories"] = meeting_subject_categories

        except Exception as e:
            # Outlook アイテムアクセスエラー等はスキップ
            print(f"Error processing item: {e}")

        finally:
            processed += 1
            if progress_callback and total_items > 0:
                progress = int((processed / total_items) * 100)
                progress_callback(progress)

    # DataFrame へ
    rows = []
    for month, meetings in meetings_by_month.items():
        for subject, details in meetings.items():
            total_minutes = details["total_duration"].total_seconds() / 60
            total_hours = total_minutes / 60
            total_days = total_hours / 7.75  # 1 日=7.75h 換算（元仕様踏襲）
            rows.append({
                "Month": month,
                "Subject Categories": details["meeting_subject_categories"],
                "Subject": subject,
                "Count": details["count"],
                "Total Duration (minutes)": round(total_minutes, 2),
                "Total Duration (hours)": round(total_hours, 2),
                "Total Duration (days)": round(total_days, 2),
                "Categories": details.get("categories", "None"),
            })

    df = pd.DataFrame(rows)

    # 出力
    now = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    os.makedirs(download_folder, exist_ok=True)
    file_path = os.path.join(download_folder, f"outlook_meetings_{now}.xlsx")
    df.to_excel(file_path, index=False)  # engine は拡張子から自動選択（openpyxl 必要）

    return file_path
