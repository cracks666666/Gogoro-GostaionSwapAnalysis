# -*- coding: utf-8 -*-

import os
import re
import pandas as pd
import matplotlib.pyplot as plt
from collections import defaultdict

try:
    import fitz
except ImportError:
    print("錯誤：缺少必要的 'PyMuPDF' 函式庫。")
    print("請在您的終端機或命令提示字元中執行以下指令來安裝：")
    print("pip install PyMuPDF")
    exit()

def extract_swap_stations_from_pdf(pdf_path):
    stations = []
    try:
        doc = fitz.open(pdf_path)
        station_pattern = re.compile(r'([\u4e00-\u9fa5\w\s\d-]+(?:站|店|門市|公所|中心|停車場))')
        time_pattern = re.compile(r'\d{2}:\d{2}:\d{2}')

        for page in doc:
            if "電池服務明細表" not in page.get_text("text"):
                continue

            words = page.get_text("words")
            if not words:
                continue

            rows = defaultdict(list)
            for word in words:
                y0 = round(word[1])
                rows[y0].append(word[4]) 
            
            for y_coord in sorted(rows.keys()):
                row_text = "".join(rows[y_coord])

                if not (time_pattern.search(row_text) and "(安時)" in row_text):
                    continue

                found_stations = station_pattern.findall(row_text)
                
                for station in found_stations:
                    cleaned_station = station.strip()
                    if "換電免費時段折抵" not in cleaned_station and "計費數量" not in cleaned_station:
                        cleaned_station = re.sub(r'\s+[A-D]$', '', cleaned_station).strip()
                        if cleaned_station:
                            stations.append(cleaned_station)

    except Exception as e:
        print(f"處理檔案 '{pdf_path}' 時發生錯誤: {e}")
        
    return stations

def analyze_bills_in_folder(folder_path='.'):
    all_swaps = []
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

    if not pdf_files:
        print("錯誤：在指定資料夾中找不到任何 PDF 檔案。")
        return 0, pd.Series()

    print(f"找到 {len(pdf_files)} 個 PDF 檔案，開始分析...")

    for pdf_file in pdf_files:
        print(f"  - 正在處理: {pdf_file}")
        file_path = os.path.join(folder_path, pdf_file)
        swaps_in_file = extract_swap_stations_from_pdf(file_path)
        all_swaps.extend(swaps_in_file)
    
    print("分析完成！")

    if not all_swaps:
        print("警告：未能在 PDF 檔案中提取到任何換電紀錄。")
        return 0, pd.Series()

    station_series = pd.Series(all_swaps)
    total_swaps = len(station_series)
    station_counts = station_series.value_counts()
    
    return total_swaps, station_counts

def plot_and_save_analysis(station_counts, output_filename='Gogoro換電分析報告.png'):
    if station_counts.empty:
        print("沒有資料可供繪圖。")
        return

    frequent_stations = station_counts[station_counts > 1]

    try:
        plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
        plt.rcParams['axes.unicode_minus'] = False
    except Exception:
        print("警告：找不到 'Microsoft JhengHei' 字體，圖表中的中文可能無法正常顯示。")
        print("您可以修改程式碼中的字體設定。")
        plt.rcParams['font.sans-serif'] = ['sans-serif']

    plt.figure(figsize=(12, 10))
    plot_data = frequent_stations.sort_values(ascending=True)
    bars = plot_data.plot(kind='barh', color='#28b4c0', zorder=2)
    
    plt.title('Gogoro 電池交換站點分析 (交換次數 > 1)', fontsize=18, pad=20)
    plt.xlabel('交換次數', fontsize=12)
    plt.ylabel('電池交換站', fontsize=12)
    plt.grid(axis='x', linestyle='--', alpha=0.6, zorder=1)

    for bar in bars.patches:
        plt.text(bar.get_width() + 0.2,
                 bar.get_y() + bar.get_height() / 2,
                 f'{int(bar.get_width())}',
                 va='center',
                 ha='left',
                 fontsize=10)

    plt.tight_layout()
    plt.savefig(output_filename)
    print(f"\n分析圖表已儲存為 '{output_filename}'")


if __name__ == "__main__":
    total_swaps, station_counts = analyze_bills_in_folder()
    
    if total_swaps > 0:
        print("\n--- Gogoro 電池交換分析結果 ---")
        print(f"\n總交換次數: {total_swaps} 次")
        print("\n各站點交換次數統計:")
        print(station_counts.to_string())
        
        plot_and_save_analysis(station_counts)
