# data/holidays_initializer.py

import os
import json
from datetime import datetime, timedelta
import jpholiday

def initialize_holidays():
    """
    各工場ごとに日本の土日祝日情報をJSONファイルとして生成・更新
    """
    # 工場リスト
    factories = ['yuki', 'kumagaya', 'shizuoka', 'kyoto', 'chiba']
    
    # 現在の年
    current_year = datetime.now().year
    
    # 生成する年の範囲（現在の5年前から20年後まで）
    start_year = current_year - 5
    end_year = current_year + 20
    
    # 休日情報を格納するディレクトリのパス
    holidays_dir = os.path.join(os.path.dirname(__file__), 'holidays')
    
    # ディレクトリが存在しない場合は作成
    if not os.path.exists(holidays_dir):
        os.makedirs(holidays_dir)
    
    for factory in factories:
        # 各工場のJSONファイルパス
        json_path = os.path.join(holidays_dir, f"{factory}.json")
        
        # 既存のデータを読み込む（存在する場合）
        if os.path.exists(json_path):
            with open(json_path, 'r', encoding='utf-8') as f:
                holidays_data = json.load(f)
        else:
            holidays_data = {}
        
        # 各年ごとに休日情報を追加
        for year in range(start_year, end_year + 1):
            year_str = str(year)
            if year_str not in holidays_data:
                holidays_data[year_str] = {}
            
            for month in range(1, 13):
                month_str = str(month)
                if month_str not in holidays_data[year_str]:
                    holidays_data[year_str][month_str] = []
                
                # 月の初日と最終日を取得
                first_day = datetime(year, month, 1)
                if month == 12:
                    next_month = datetime(year + 1, 1, 1)
                else:
                    next_month = datetime(year, month + 1, 1)
                last_day = next_month - timedelta(days=1)
                
                # 月内の全日をチェック
                for day in range(1, last_day.day + 1):
                    date = datetime(year, month, day)
                    if jpholiday.is_holiday(date) or date.weekday() >= 5:  # 土日祝日
                        holidays_data[year_str][month_str].append(day)
                
                # 重複を排除し、ソート
                holidays_data[year_str][month_str] = sorted(list(set(holidays_data[year_str][month_str])))
        
        # JSONファイルに書き込む
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(holidays_data, f, ensure_ascii=False, indent=4)

if __name__ == "__main__":
    initialize_holidays()
    print("工場の休日情報をリセットしました。")
