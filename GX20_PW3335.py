# SAMPO GX20/GX30/PW3335 Data Collection 0_1
#-------------------------------------------------------------------------------
#GX20 info: Yokogawa GX20 Paperless Recorder
#       document : IM04L51B01-17EN 
#PW3335 info : GW Instek PW3335 Programmable DC Power Meter
#       document : PW_Communicator_zh / 2018 年1月出版 (改定1.60版)
#-------------------------------------------------------------------------------
#Rev 1_0 2025/5/6 紀錄GX20與PW3335數據
#         1. 讀取GX20的溫度數據
#         2. 讀取PW3335的電壓/電流/功率數據
#         3. 繪製溫度與功率的圖表
#         4. 儲存數據到CSV檔案
#         5. 支援多工位數據收集,計算等功能
#Rev 1_1 2025/5/7 修正報告功能,新增能耗參數進行計算
#-------------------------------------------------------------------------------
import socket
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox  # 修正：添加 messagebox 的導入
import csv
from datetime import datetime, timedelta  # 修正：添加 timedelta 的導入
import pandas as pd  # 修正：添加 pandas 的導入
import threading
import os,sys
from ctypes import windll
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
from matplotlib import rcParams
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
import tkinter.font as tkfont
import matplotlib.dates as mdates

import tempfile

# 確保 LOG 檔案儲存到執行檔所在目錄或臨時目錄
if getattr(sys, 'frozen', False):  # 如果是 pyinstaller 打包的執行檔
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

LOG_PATH = os.path.join(APP_DIR, "Gx20_Pw3335.log")

# 如果無法寫入執行檔目錄，則使用臨時目錄
if not os.access(APP_DIR, os.W_OK):
    LOG_PATH = os.path.join(tempfile.gettempdir(), "Gx20_Pw3335.log")

# 設定 matplotlib 使用的字體
rcParams['font.sans-serif'] = ['Microsoft JhengHei']  # 使用微軟正黑體
rcParams['axes.unicode_minus'] = False  # 解決負號無法顯示的問題
rcParams['font.size'] = 10  # 直接指定matplotlib全局字體大小
rcParams['axes.titlesize'] = 12      # 座標軸標題字體大小
rcParams['axes.labelsize'] = 12      # 座標軸標籤字體大小
rcParams['xtick.labelsize'] = 12     # X軸刻度字體大小
rcParams['ytick.labelsize'] = 12     # Y軸刻度字體大小
rcParams['legend.fontsize'] = 10     # 圖例字體大小
rcParams['figure.titlesize'] = 16    # 圖形標題字體大小

def log_to_file(message):
    """將訊息寫入 LOG 檔案"""
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(message + "\n")
    except Exception as e:
        print(f"無法寫入 LOG 檔案: {e}")
def log_error(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    msg = f"[ERROR] {timestamp} - {message}"
    print(msg)
    log_to_file(msg)

def log_info(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    msg = f"[INFO] {timestamp} - {message}"
    print(msg)
    log_to_file(msg)

class GX20:
    def __init__(self, host="192.168.1.1", port=34434):
        self.valid_data = {}  # 有效數據
        self.gsRemoteHost = host
        self.gnRemotePort = port

    def parse_scientific_notation(self, value_str):
        """解析科學記號格式的數值，非數字或大於999時回傳 None"""
        try:
            if 'E' in value_str:
                base, exp = value_str.split('E')
                value = float(base) * (10 ** int(exp))
                    # 檢查值是否大於 999
                return None if value > 999 else value
            
        except (ValueError, TypeError):
            return None

    def parse_channel_data(self, line):
        """解析頻道數據
        格式: 31字元
        - 第1字元: 資料狀態 (N/B)
        - 第3-6字元: 頻道號碼
        - 第11-18字元: 單位
        - 第19字元: 正負號
        - 第20-31字元: 科學符號數值
        """
        if len(line) != 31:
            return None
            
        data_type = line[0]  # 資料狀態
        channel = line[2:6]  # 頻道號碼
        unit = line[10:18].strip()  # 單位
        sign = line[18]  # 正負號
        value_str = sign + line[19:31]  # 科學符號數值
        
        return {
            "type": data_type,
            "channel": channel,
            "unit": unit,
            "value_str": value_str
        }

    def GX20GetData(self):
        try:
            # 建立 TCP 連線
            with socket.create_connection((self.gsRemoteHost, self.gnRemotePort), timeout=3) as s:
                # 送出指令
                s.sendall(b"FData,0,0001,1210\r\n")
                # 暫停0.5秒
                time.sleep(0.5)
                # 接收資料
                data = s.recv(10240).decode("ascii", errors="ignore")
                #print("Raw data:", repr(data))
                
                # 初始化數據字典
                self.valid_data = {}
                
                # 解析每一行數據
                lines = data.split('\r\n')
                # 從第5筆開始處理頻道數據
                for line in lines[4:]:
                    if len(line) == 31:  # 確保是頻道數據行
                        #print(line)
                        parsed = self.parse_channel_data(line)
                            # 解析數值
                        value = self.parse_scientific_notation(parsed["value_str"])
                        self.valid_data[parsed["channel"]] = {
                            "value": value,
                            "unit": parsed["unit"]
                        }
                
        except Exception as e:
            self.valid_data = {"Error": f"Error-Code: {str(e)}"}

    def decode_temperature(self, channels: list[str]) -> list[float]:
        """
        根據 self.valid_data 取出指定 channels 的溫度值，沒有資料則回傳 None。
        """
        return [
            self.valid_data.get(ch, {}).get("value", None)
            for ch in channels
        ]

class PW3335:
    def __init__(self, ip_address, port=3300):
        self.ip_address = ip_address
        self.port = port
        self.sock = None

    def connect(self):
        """Establish a TCP connection to the power meter."""
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.connect((self.ip_address, self.port))

    def disconnect(self):
        """Close the TCP connection."""
        if self.sock:
            self.sock.close()
            self.sock = None

    def query_data(self):
        """Query voltage, current, power, and accumulated power."""
        if not self.sock:
            raise ConnectionError("Socket is not connected to the power meter.")
        self.sock.sendall(b':MEAS? U,I,P,WH\n')
        response = self.sock.recv(1024).decode('ascii').strip()
        try:
            # Parse the response format: "U +110.14E+0;I +0.0000E+0;P +000.00E+0;WP +00.0000E+0"
            data = response.split(';')
            if len(data) == 4:
                parsed_data = [float(item.split(' ')[-1].replace('E+0', '')) for item in data]
                return parsed_data
            else:
                raise ValueError(f"Unexpected response format: {response}")
        except Exception as e:
            print(f"Error parsing response: {response}, Exception: {e}")
            raise ValueError(f"Failed to parse response: {response}")

class EnergyCalculator:
    def __init__(self):
        pass
    def current_ef_thresholds(self,energy_allowance,fridge_type):
        if fridge_type == 5:
            #IF(fridge_type=5,ROUND(N4*1.72,1),ROUND(N4*1.6,1))
            threshold_lv1 = round(energy_allowance * 1.72,1)
            threshold_lv2 = round(energy_allowance * 1.54,1)
            threshold_lv3 = round(energy_allowance * 1.36,1)
            threshold_lv4 = round(energy_allowance * 1.18,1)
        else:
            #IF(fridge_type=5,ROUND(N4*1.72,1),ROUND(N4*1.6,1))
            threshold_lv1 = round(energy_allowance * 1.6,1)
            threshold_lv2 = round(energy_allowance * 1.45,1)
            threshold_lv3 = round(energy_allowance * 1.3,1)
            threshold_lv4 = round(energy_allowance * 1.15,1)
        return[ threshold_lv1, threshold_lv2, threshold_lv3, threshold_lv4 ]

    def future_ef_thresholds(self,energy_allowance,fridge_type):
        if fridge_type == 5:
            #IF(fridge_type=5,ROUND(N4*1.72,1),ROUND(N4*1.6,1))
            threshold_lv1 = round(energy_allowance * 1.294,1)
            threshold_lv2 = round(energy_allowance * 1.221,1)
            threshold_lv3 = round(energy_allowance * 1.147,1)
            threshold_lv4 = round(energy_allowance * 1.074,1)
        else:
            #IF(fridge_type=5,ROUND(N4*1.72,1),ROUND(N4*1.6,1))
            threshold_lv1 = round(energy_allowance * 1.308,1)
            threshold_lv2 = round(energy_allowance * 1.231,1)
            threshold_lv3 = round(energy_allowance * 1.154,1)
            threshold_lv4 = round(energy_allowance * 1.077,1)
        return[ threshold_lv1, threshold_lv2, threshold_lv3, threshold_lv4 ]


    def calculate(self, VF, VR, daily_consumption, fridge_temp, freezer_temp, fan_type):
        """
        計算冰箱能耗相關指標
        
        參數:
            VR: 冷藏室容積(L)
            VF: 冷凍室容積(L)
            daily_consumption: 日耗電量(kWh/日)
            fridge_temp: 冷藏室溫度(°C), 預設3.0
            freezer_temp: 冷凍室溫度(°C), 預設-18.0
        
        返回:
            包含所有計算結果的字典
        """
        results = {}
        
        # 1. 計算K值 (溫度係數)
        K = self.calculate_K_value(freezer_temp, fridge_temp)
        #print(f"K值: {K}")
        # 2. 計算等效內容積
        equivalent_volume = self.calculate_equivalent_volume(VR, VF, K)
        
        # 3. 確定冰箱型式
        fridge_type = self.determine_fridge_type(equivalent_volume, VR, VF, fan_type)
        #print(f"冰箱型式: {fridge_type}")
        # 4. 計算容許耗用能源基準 (每月)
        energy_allowance = self.calculate_energy_allowance(equivalent_volume, fridge_type)
        
        # 5. 計算2027容許耗用能源基準
        future_energy_allowance = self.calculate_future_energy_allowance(equivalent_volume, fridge_type)
        
        # 6. 計算耗電量基準 (每月)
        benchmark_consumption = self.calculate_benchmark_consumption(equivalent_volume, energy_allowance)
        
        # 7. 計算2027耗電量基準
        future_benchmark_consumption = self.calculate_future_benchmark_consumption(equivalent_volume, future_energy_allowance)
        
        # 8. 計算實測月耗電量
        monthly_consumption = round(daily_consumption * 30,1)
        
        # 9. 計算EF值 (能效因子)
        ef_value = round(equivalent_volume / monthly_consumption,1)
        
        # 9.1 計算現有效率基準百分比和等級
        current_ef_thresholds = self.current_ef_thresholds(energy_allowance, fridge_type)

        # 10. 計算現有效率基準百分比和等級
        current_percent, current_grade = self.calculate_current_efficiency(ef_value, current_ef_thresholds)
        
        # 10.1 計算2027新效率基準百分比和等級
        future_ef_thresholds = self.future_ef_thresholds(future_energy_allowance, fridge_type)

        # 11. 計算2027新效率基準百分比和等級
        future_percent, future_grade = self.calculate_future_efficiency(ef_value, future_ef_thresholds)
        
        # 整理所有結果
        results.update({
            'K值': K,
            'VF(L)': VF,
            'VR(L)': VR,
            '等效內容積(L)': equivalent_volume,
            '冰箱型式': fridge_type,
            '容許耗用能源基準(L/kWh/月)': energy_allowance,
            '2027容許耗用能源基準(L/kWh/月)': future_energy_allowance,
            '耗電量基準(kWh/月)': benchmark_consumption,
            '2027耗電量基準(kWh/月)': future_benchmark_consumption,
            '實測月耗電量(kWh/月)': monthly_consumption,
            'EF值': ef_value,
            '現有效率基準百分比(%)': current_percent,
            '現有效率等級': current_grade,
            '2027新效率基準百分比(%)': future_percent,
            '2027新效率等級': future_grade
        })
        
        return results
    
    def calculate_K_value(self, freezer_temp, fridge_temp):
        """計算K值 (溫度係數)"""
        # 根據公式 K = (30 - 冷凍庫溫度) / (30 - 冷藏庫溫度)
        print(f"冷凍庫溫度: {freezer_temp}, 冷藏庫溫度: {fridge_temp}")        
        return round((30 - freezer_temp) / (30 - fridge_temp), 2)
    
    def calculate_equivalent_volume(self, VR, VF, K):
        """計算等效內容積"""
        return round(VR + (K * VF), 1)
    
    def determine_fridge_type(self, equivalent_volume, VR, VF, fan_type):
        """確定冰箱型式"""
        if VF == 0:  # 只有冷藏室
            return 5
        elif equivalent_volume < 400 and fan_type == 1:
            return 1  # 假設是風冷式(實際應根據具體設計)
        elif equivalent_volume >= 400 and fan_type == 1:
            return 2
        elif equivalent_volume < 400 and fan_type == 0:
            return 3
        else:
            return 4  # 假設是風冷式(實際應根據具體設計)
    
    def calculate_energy_allowance(self, equivalent_volume, fridge_type):
        """計算容許耗用能源基準"""
        # 根據公式，ROUND(IFS(fridge_type=1,equivalent_volume/(0.037*equivalent_volume+24.3),fridge_type=2,equivalent_volume/(0.031*M4+21),fridge_type=3,equivalent_volume/(0.033*equivalent_volume+19.7),fridge_type=4,equivalent_volume/(0.029*equivalent_volume+17),fridge_type=5,equivalent_volume/(0.033*equivalent_volume+15.8)),1)
        if fridge_type == 1:
            return round( equivalent_volume / (0.037 * equivalent_volume + 24.3), 1)
        elif fridge_type == 2:
            return round( equivalent_volume / (0.031 * equivalent_volume + 21), 1)
        elif fridge_type == 3:
            return round( equivalent_volume / (0.033 * equivalent_volume + 19.7), 1)
        elif fridge_type == 4:
            return round( equivalent_volume / (0.029 * equivalent_volume + 17), 1)
        else:
            return round( equivalent_volume / (0.033 * equivalent_volume + 15.8), 1)
    
    def calculate_future_energy_allowance(self, equivalent_volume, fridge_type):
        """計算2027年容許耗用能源基準"""
        # 公式:=ROUND(IFS(F4=1,1.3*M4/(0.037*M4+24.3),F4=2,1.3*M4/(0.031*M4+21),F4=3,1.3*M4/(0.033*M4+19.7),F4=4,1.3*M4/(0.029*M4+17),F4=5,1.36*M4/(0.033*M4+15.8)),1)
        # F4 = fridge_type, M4 = equivalent_volume
        if fridge_type == 1:
            return round( 1.3 * equivalent_volume / (0.037 * equivalent_volume + 24.3), 1)
        elif fridge_type == 2:
            return round( 1.3 * equivalent_volume / (0.031 * equivalent_volume + 21), 1)
        elif fridge_type == 3:
            return round(1.3 * equivalent_volume / (0.033 * equivalent_volume + 19.7), 1)
        elif fridge_type == 4:
            return round(1.3 * equivalent_volume / (0.029 * equivalent_volume + 17), 1)
        else:
            return round(1.36 * equivalent_volume / (0.033 * equivalent_volume + 15.8), 1)

    def calculate_benchmark_consumption(self, equivalent_volume, energy_allowance):
        """計算耗電量基準"""
        # 根據公式:ROUND(equivalent_volume / energy_allowance, 1)
        return round(equivalent_volume / energy_allowance, 1)
    
    def calculate_future_benchmark_consumption(self, equivalent_volume, future_energy_allowance):
        """計算2027耗電量基準"""
        return round(equivalent_volume / future_energy_allowance, 1)
    
    def calculate_current_efficiency(self, ef_value, thresholds):
        # 確定等級
        if ef_value >= thresholds[0]:
            grade = "1級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[0] * 0.95:
            grade = "1*級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[1]:
            grade = "2級"
            final_percent = round(ef_value / thresholds[1] * 100, 1)
        elif ef_value >= thresholds[2]:
            grade = "3級"
            final_percent = round(ef_value / thresholds[2] * 100, 1)
        elif ef_value >= thresholds[3]:
            grade = "4級"
            final_percent = round(ef_value / thresholds[3] * 100, 1)
        else :
            grade = "5級"
            final_percent = round(ef_value / thresholds[3] * 100, 1)
        
        return final_percent, grade
    
    def calculate_future_efficiency(self, ef_value, thresholds):
        # 確定等級
        if ef_value >= thresholds[0]:
            grade = "1級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[0] * 0.95:
            grade = "1*級"
            final_percent = round(ef_value / thresholds[0] * 100, 1)
        elif ef_value >= thresholds[1]:
            grade = "2級"
            final_percent = round(ef_value / thresholds[1] * 100, 1)
        elif ef_value >= thresholds[2]:
            grade = "3級"
            final_percent = round(ef_value / thresholds[2] * 100, 1)
        elif ef_value >= thresholds[3]:
            grade = "4級"
            final_percent = round(ef_value / thresholds[3] * 100, 1)
        else :
            grade = "5級"
            final_percent = round(ef_value / thresholds[3] * 100, 1)
        
        return final_percent, grade


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("SAMPO GX20/PW3335 Data Collection")  # 初始標題
        self.file_path = ""  # 初始化 file_path 屬性
        self.collecting = {}  # 初始化 collecting 屬性，用於跟蹤正在收集數據的設備
        self.gx20_ip = "192.168.1.1"
        self.gx20_instance = GX20(self.gx20_ip)
        self.pw3335_instances = {}
        self.time_data = []
        self.temperature_data = []
        self.power_data = []
        self.pause_plot = False  # 新增變數，用於控制圖表更新的暫停/恢復
        self.gx20_data_dict = {}  # {channel: value}
        self.gx20_connected = False
        self.plot_channel_labels = {}

        # 新增：儲存各工位的頻道對應
        self.channel_number = {
            "工位1": "0001-0010,0101-0110",
            "工位2": "0201-0210,0301-0310",
            "工位3": "0401-0410,1001-1010",
            "工位4": "0701-0710,0801-0810",
            "工位5": "0501-0510,0601-0610",
            "工位6": "1101-1110,1201-1210"
        }
        # 為每個工位創建獨立的數據存儲
        self.station_data = {
            f"工位{i}": {
                "time_data": [],
                "temperature_data": [],
                "power_data": [],
            }
            for i in range(1, 7)
        }
        
        # 初始化 Notebook（頁面容器）
        self.notebook = ttk.Notebook(root)
        self.notebook.place(x=5, y=5, width=ws-10, height=hs-10)

        # 創建 6 個頁面（工位 1 到工位 6）
        self.frames = {}
        for i in range(1, 7):
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=f"   工位{i}   ", padding=5)
            self.frames[f"工位{i}"] = frame
            # 在每個頁面中添加控件
            self.setup_station_page(frame, f"工位{i}")

        # 綁定窗口關閉事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 啟動 GX20 連線與資料更新執緒
        threading.Thread(target=self.gx20_data_updater, daemon=True).start()

        # 初始化 PW3335 實例
        self.pw3335_instances = {}
        for i in range(1, 7):
            pw_ip = f"192.168.1.{i + 1}"
            try:
                pw = PW3335(pw_ip)
                pw.connect()
                self.pw3335_instances[pw_ip] = pw
            except Exception as e:
                print(f"PW3335 {pw_ip} 連線失敗: {e}")
                log_error(f"App.init:PW3335 {pw_ip} 連線失敗: {e}")


    def setup_station_page(self, frame, station_name):
        """設置每個工位頁面的控件"""
        station_name = station_name.replace(" ", "")  # 去除空格，統一名稱格式

        # 創建子頁面的 Notebook
        station_notebook = ttk.Notebook(frame)
        station_notebook.grid(row=0, column=0, sticky="nsew")
        
        # 創建四個子頁面
        file_frame = ttk.Frame(station_notebook)
        channel_frame = ttk.Frame(station_notebook)
        plot_frame = ttk.Frame(station_notebook)
        report_frame = ttk.Frame(station_notebook)
        
        # 將子頁面加入 Notebook
        station_notebook.add(file_frame, text="  FILE  ")
        station_notebook.add(channel_frame, text="  CHANNEL  ")
        station_notebook.add(plot_frame, text="  PLOT  ")
        station_notebook.add(report_frame, text="  REPORT  ")
        
        # 設置 frame 的網格權重，使其可以填滿整個空間
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
        # 在各個子頁面中設置控件
        self.setup_file_page(file_frame, station_name)
        self.setup_channel_page(channel_frame, station_name)
        self.setup_plot_page(plot_frame, station_name)
        self.setup_report_page(report_frame, station_name)

    def setup_file_page(self, frame, station_name):
        """設置 FILE 頁面的控件"""
        # File path selection
        ttk.Label(frame, text="儲存路徑:").grid(row=0, column=0, padx=5, pady=5)
        file_path_var = tk.StringVar(value="D:/sampo")  # 設定預設路徑
        file_path_entry = ttk.Entry(frame, textvariable=file_path_var, width=30)
        file_path_entry.grid(row=0, column=1, padx=5, pady=5)
        browse_button = ttk.Button(frame, text="Browse", command=lambda: self.browse_file(file_path_var))
        browse_button.grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(frame, text="檔名:").grid(row=1, column=0, padx=5, pady=5)
        file_name_var = tk.StringVar(value="*.csv") 
        file_name_entry = ttk.Entry(frame, textvariable=file_name_var, width=30, state="readonly")
        file_name_entry.grid(row=1, column=1, padx=5, pady=5)

        # Frequency selection
        ttk.Label(frame, text="記錄頻率(sec):").grid(row=2, column=0, padx=5, pady=5)
        frequency_var = tk.IntVar(value=10)
        frequency_menu = ttk.Combobox(frame, textvariable=frequency_var, state="readonly")
        frequency_menu['values'] = [10, 60, 180, 300]
        frequency_menu.grid(row=2, column=1, padx=5, pady=5)

        # Start, Stop buttons
        start_button = ttk.Button(frame, text="Start", command=lambda: self.start_collection(station_name), state="normal")
        start_button.grid(row=1, column=2, padx=5, pady=5)
        stop_button = ttk.Button(frame, text="Stop", command=lambda: self.stop_collection(station_name), state="disabled")
        stop_button.grid(row=2, column=2, padx=5, pady=5)

        # 分割線
        ttk.Separator(frame, orient="horizontal").grid(row=3, column=0, columnspan=3, sticky="ew", pady=10)

        # 新增一個frame,設定冰箱規格
        ttk.Label(frame, text="冰箱規格:").grid(row=4, column=0, padx=5, pady=5)
        ttk.Label(frame, text="機種:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        model_entry_var = tk.StringVar(value="NA")
        model_entry = ttk.Entry(frame, width=30, textvariable=model_entry_var)
        model_entry.grid(row=5, column=1, padx=5, pady=5)
        ttk.Label(frame, text="冷凍庫容量(L):").grid(row=6, column=0, padx=5, pady=5, sticky="w")
        vf_entry_var = tk.StringVar(value="150")
        vf_entry = ttk.Entry(frame, width=10, textvariable=vf_entry_var)
        vf_entry.grid(row=6, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(frame, text="冷藏庫容量:").grid(row=7, column=0, padx=5, pady=5, sticky="w")
        vr_entry_var = tk.StringVar(value="350")
        vr_entry = ttk.Entry(frame, width=10, textvariable=vr_entry_var)
        vr_entry.grid(row=7, column=1, padx=5, pady=5, sticky="w")
        # Fan type checkbox
        fan_type_var = tk.IntVar(value=1)  # 0: unchecked, 1: checked
        fan_type_checkbox = ttk.Checkbutton(frame, text="風扇式", variable=fan_type_var)
        fan_type_checkbox.grid(row=8, column=0, columnspan=2, padx=5, pady=5, sticky="w")

        # Save references
        setattr(self, f"{station_name}_file_path_var", file_path_var)
        setattr(self, f"{station_name}_file_path_entry", file_path_entry)
        setattr(self, f"{station_name}_file_name_var", file_name_var)
        setattr(self, f"{station_name}_file_name_entry", file_name_entry)
        setattr(self, f"{station_name}_Browse_button", browse_button)
        setattr(self, f"{station_name}_frequency_var", frequency_var)
        setattr(self, f"{station_name}_frequency_menu", frequency_menu)
        setattr(self, f"{station_name}_start_button", start_button)
        setattr(self, f"{station_name}_stop_button", stop_button)
        setattr(self, f"{station_name}_model_entry", model_entry)
        setattr(self, f"{station_name}_model_entry_var", model_entry_var)
        setattr(self, f"{station_name}_vf_entry", vf_entry)
        setattr(self, f"{station_name}_vf_entry_var", vf_entry_var)
        setattr(self, f"{station_name}_vr_entry", vr_entry)
        setattr(self, f"{station_name}_vr_entry_var", vr_entry_var)
        setattr(self, f"{station_name}_fan_type_var", fan_type_var)
        setattr(self, f"{station_name}_fan_type_checkbox", fan_type_checkbox)

    def browse_file(self, file_path_var):
        file_path = filedialog.askdirectory()
        file_path_var.set(file_path)
        self.file_path = file_path  # 將選擇的路徑保存到 self.file_path

    def setup_channel_page(self, frame, station_name):
        """設置 CHANNEL 頁面的控件"""
        # 頻道設定標題
        ttk.Label(frame, text="頻道設定").grid(row=0, column=0, columnspan=6, pady=5)  # 改為6欄

        # 按鈕框架
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=1, column=0, columnspan=6, pady=5)  # 改為6欄

        select_all_btn = ttk.Button(
            button_frame, 
            text="全選", 
            command=lambda: self.update_all_checkboxes(station_name, True)
        )
        select_all_btn.grid(row=0, column=0, padx=5)

        deselect_all_btn = ttk.Button(
            button_frame, 
            text="全取消", 
            command=lambda: self.update_all_checkboxes(station_name, False)
        )
        deselect_all_btn.grid(row=0, column=1, padx=5)

        # 頻道列表 (10行，每行6個元件)
        channels = self.parse_channels(self.channel_number[station_name])
        checkboxes = {}
        ch_aliases = {}

        for i in range(10):  # 10 行
            row = i + 2  # 從第2行開始

            # 左側 CH01-CH10
            ch_num_left = channels[i] if i < len(channels) else None
            if ch_num_left:
                # 左側頻道號碼
                ttk.Label(frame, text=f"{ch_num_left}").grid(
                    row=row, 
                    column=0, 
                    padx=2, 
                    sticky='e'
                )

                # 左側 Checkbox
                var_left = tk.IntVar(value=1)
                checkbox_left = ttk.Checkbutton(frame, variable=var_left)
                checkbox_left.grid(row=row, column=1, padx=2)
                checkboxes[ch_num_left] = var_left

                # 左側別名輸入框
                alias_entry_left = ttk.Entry(frame, width=8)
                alias_entry_left.grid(row=row, column=2, padx=2, sticky='ew')
                ch_aliases[ch_num_left] = alias_entry_left

            # 右側 CH11-CH20
            ch_num_right = channels[i + 10] if i + 10 < len(channels) else None
            if ch_num_right:
                # 右側頻道號碼
                ttk.Label(frame, text=f"{ch_num_right}").grid(
                    row=row, 
                    column=3, 
                    padx=2, 
                    sticky='e'
                )

                # 右側 Checkbox
                var_right = tk.IntVar(value=1)
                checkbox_right = ttk.Checkbutton(frame, variable=var_right)
                checkbox_right.grid(row=row, column=4, padx=2)
                checkboxes[ch_num_right] = var_right

                # 右側別名輸入框
                alias_entry_right = ttk.Entry(frame, width=8)
                alias_entry_right.grid(row=row, column=5, padx=2, sticky='ew')
                ch_aliases[ch_num_right] = alias_entry_right

        # 保存引用
        setattr(self, f"{station_name}_checkboxes", checkboxes)
        setattr(self, f"{station_name}_ch_aliases", ch_aliases)
    
    def update_all_checkboxes(self, station_name, value: bool):
        """更新指定工位的所有 checkbox 狀態
        Args:
            station_name: 工位名稱
            value: True 為全選，False 為全取消
        """
        checkboxes = getattr(self, f"{station_name}_checkboxes", {})
        for var in checkboxes.values():
            var.set(1 if value else 0)
        # 更新圖表顯示
        #self.update_plot(station_name=station_name)  # 傳遞 station_name

    def setup_plot_page(self, frame, station_name):
        self.start_date = tk.StringVar()
        self.start_time = tk.StringVar()
        self.end_date = tk.StringVar()
        self.end_time = tk.StringVar()
        """設置 PLOT 頁面的控件"""
        # X 軸範圍選擇
        ttk.Label(frame, text="X軸區間:", width=8).grid(row=0, column=0, padx=1, pady=5)
        x_axis_range_var = tk.StringVar(value="30min")
        x_axis_range_menu = ttk.Combobox(frame, textvariable=x_axis_range_var, state="readonly", width=6)
        x_axis_range_menu['values'] = ["30min", "3hrs", "12hrs", "24hrs"]
        x_axis_range_menu.grid(row=1, column=0, padx=1, pady=5)
        
        ttk.Label(frame, text=station_name, width=8).grid(row=0, column=1, padx=1, pady=5)

        # Pause/Resume button
        pause_button = ttk.Button(frame, text="暫停", command=lambda: self.toggle_pause_plot(station_name), 
                                state="disabled", width=6)
        pause_button.grid(row=2, column=0, padx=5, pady=5)
        
        # 設置20個頻道讀值標籤
        channel_labels = []  # 新增：儲存當前工位的頻道標籤
        
        # 解析該工位的頻道設定
        channels = self.parse_channels(self.channel_number[station_name])
        
        # 創建頻道標籤
        for i, ch in enumerate(channels):
            if i < 10:
                row, col = 3 + i, 0
            else:
                row, col = 3 + (i-10), 1
            
            ch_label = ttk.Label(frame, text=f"{ch}", width=4, relief="solid", anchor="center")
            ch_label.grid(row=row, column=col, padx=0.5, pady=5)
            channel_labels.append((ch, ch_label))  # 儲存頻道號碼和對應的標籤
        
        def increment_date_time(var, increment, unit):
            try:
                current_value = pd.to_datetime(var.get())
                if unit == "day":
                    new_value = current_value + pd.Timedelta(days=increment)
                elif unit == "hour":
                    new_value = current_value + pd.Timedelta(hours=increment)
                var.set(new_value.strftime('%Y-%m-%d' if unit == "day" else '%H:%M'))
            except Exception:
                messagebox.showerror("錯誤", "無效的日期或時間格式！")

        def bind_increment(widget, var, unit):
            def on_key(event):
                if event.state & 0x4:  # 檢查是否按下 CTRL 鍵
                    if event.keysym == "Up":
                        increment_date_time(var, 1, unit)
                    elif event.keysym == "Down":
                        increment_date_time(var, -1, unit)
            widget.bind("<KeyPress-Up>", on_key)
            widget.bind("<KeyPress-Down>", on_key)


        start_date_entry = ttk.Entry(frame, textvariable=self.start_date, width=10)
        start_date_entry.grid(row=13, column=0, padx=1, pady=1)
        bind_increment(start_date_entry, self.start_date, "day")

        start_time_entry = ttk.Entry(frame, textvariable=self.start_time, width=10)
        start_time_entry.grid(row=14, column=0, padx=1, pady=1)
        bind_increment(start_time_entry, self.start_time, "hour")

        end_date_entry = ttk.Entry(frame, textvariable=self.end_date, width=10)
        end_date_entry.grid(row=13, column=1, padx=1, pady=1)
        bind_increment(end_date_entry, self.end_date, "day")

        end_time_entry = ttk.Entry(frame, textvariable=self.end_time, width=10)
        end_time_entry.grid(row=14, column=1)
        bind_increment(end_time_entry, self.end_time, "hour")

        #計算結果顯示區
        calculate_text = tk.Text(frame, height=6, width=20, wrap="word")
        calculate_text.grid(row=13, rowspan=2, column=2)
        calculate_text.insert(tk.END, "計算結果顯示區\n")
        # calculate button
        calculate_button = ttk.Button(frame, text="平均", command=lambda: self.calculate_avg_temp(station_name), width=6)
        calculate_button.grid(row=13, column=3)
        # report button
        report_button = ttk.Button(frame, text="報告", command=lambda: self.report_calculate(station_name), width=6)
        report_button.grid(row=14, column=3)
        
        # 儲存該工位的頻道標籤
        self.plot_channel_labels[station_name] = channel_labels

        figure = plt.Figure(figsize=(15, 7), dpi=85)
        canvas = FigureCanvasTkAgg(figure, master=frame)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.grid(row=0, rowspan= 13, column=2, columnspan=5, padx=5, pady=5)
        ax_temp = figure.add_subplot(211, facecolor='lightcyan')
        ax_power = figure.add_subplot(212, sharex=ax_temp, facecolor='lightyellow')
        # 減少左右空白
        figure.subplots_adjust(left=0.05, right=0.95, top=0.92, bottom=0.10, hspace=0.30)

        ax_temp.set_ylabel("Temperature (°C)")
        ax_temp.get_xaxis().set_visible(False)
        ax_temp.grid(True)
        ax_power.set_xlabel("Time")
        ax_power.set_ylabel("Power (W)")
        ax_power.grid(True)

        # 新增：創建一個 Frame 來容納工具欄，並使用標準工具欄
        toolbar_frame = ttk.Frame(frame)
        toolbar_frame.grid(row=0, column=2, columnspan= 6, padx=5, pady=5)
        toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)  # 使用標準工具欄
        #toolbar.grid(row=0, column=0)  # 在 toolbar_frame 中使用 grid
        toolbar.update()

        # Memo text box
        memo_text = tk.Text(frame, height=6, width=80, wrap="word")
        memo_text.grid(row=13, rowspan=2, column=4, padx=5, pady=5)
        memo_text.insert(tk.END, "備註:\n")
        
        # Save references
        setattr(self, f"{station_name}_figure", figure)
        setattr(self, f"{station_name}_canvas", canvas)
        setattr(self, f"{station_name}_ax_temp", ax_temp)
        setattr(self, f"{station_name}_ax_power", ax_power)
        setattr(self, f"{station_name}_x_axis_range_var", x_axis_range_var)
        setattr(self, f"{station_name}_pause_button", pause_button)
        setattr(self, f"{station_name}_x_axis_range_var", x_axis_range_var)
        setattr(self, f"{station_name}_toolbar", toolbar)
        setattr(self, f"{station_name}_toolbar_frame", toolbar_frame)

        setattr(self, f"{station_name}_start_date_entry", start_date_entry)
        setattr(self, f"{station_name}_start_time_entry", start_time_entry)
        setattr(self, f"{station_name}_end_date_entry", end_date_entry)
        setattr(self, f"{station_name}_end_time_entry", end_time_entry)
        setattr(self, f"{station_name}_calculate_text", calculate_text)
        setattr(self, f"{station_name}_calculate_button", calculate_button)

        setattr(self, f"{station_name}_channel_labels", channel_labels)  # 儲存頻道標籤

        setattr(self, f"{station_name}_memo_text", memo_text)  # 儲存備註文本框

        # 啟動 FuncAnimation
        anim = FuncAnimation(
                            figure,
                            self.update_plot,
                            fargs=(station_name,),
                            interval=5000,  # 每 5000 毫秒更新一次
                            cache_frame_data=False,  # 禁用幀緩存
                            blit=False  # 禁用 blit 以避免可能的繪圖問題
                            )

        # 保存動畫引用，避免被垃圾回收
        setattr(self, f"{station_name}_animation", anim)

    def setup_report_page(self, frame, station_name):
        """設置 REPORT 頁面的控件"""
        report_text = tk.Text(frame, height=35, width=100, wrap="word")
        report_text.grid(row=0, column=0, padx=5, pady=5)
        report_text.insert(tk.END, "NA\n")
        save_button = ttk.Button(frame, text="儲存", command=lambda: self.save_report(station_name), width=10)
        save_button.grid(row=1, column=0, padx=5, pady=5)

        setattr(self, f"{station_name}_report_text", report_text)
  
    
    def save_report(self, station_name):
        """儲存報告"""
        file_path = filedialog.askdirectory()
        if not file_path:
            self.show_error_dialog("路徑錯誤", "請先選擇儲存路徑")
            return
        report_text = getattr(self, f"{station_name}_report_text", None)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if report_text:
            report_content = report_text.get("1.0", tk.END)
            file_name = f"{station_name}_report_{timestamp}.txt"
            with open(f"{file_path}/{file_name}", "w") as f:
                f.write(report_content)
            messagebox.showinfo("儲存成功", f"報告已儲存到 {file_path}/{file_name}")
        else:
            self.show_error_dialog("報告錯誤", "無法獲取報告內容")

    def parse_channels(self, channel_str: str) -> list[str]:
        channels = []
        parts = channel_str.split(',')
        for part in parts:
            part = part.strip()
            if '-' in part:
                start, end = part.split('-')
                start, end = start.strip(), end.strip()
                prefix_start, num_start = start[:2], int(start[2:])
                prefix_end, num_end = end[:2], int(end[2:])
                if prefix_start != prefix_end:
                    raise ValueError("範圍必須在同一區段，例如 0201-0207")
                for n in range(num_start, num_end + 1):
                    channels.append(f"{prefix_start}{n:02d}")
            else:
                if len(part) != 4 or not part.isdigit():
                    raise ValueError("頻道格式錯誤，必須為4位數字")
                channels.append(part)
        return sorted(channels)

    def start_collection(self, station_name):
        """啟動數據收集"""
        if not self.gx20_connected:
            self.handle_gx20_connection_error()
            return
        
        try:
            # 檢查檔案路徑
            file_path_var = getattr(self, f"{station_name}_file_path_var", None)
            if not file_path_var or not file_path_var.get():
                self.show_error_dialog("路徑錯誤", "請先選擇儲存路徑")
                return
            self.file_path = file_path_var.get()

            # **檢查至少有一個頻道被勾選**
            checkboxes = getattr(self, f"{station_name}_checkboxes", {})
            if not any(var.get() == 1 for var in checkboxes.values()):
                self.show_error_dialog("頻道選擇錯誤", "請至少勾選一個頻道！")
                return

            # 啟用 stop 按鈕，禁用 start 按鈕
            start_button = getattr(self, f"{station_name}_start_button", None)
            stop_button = getattr(self, f"{station_name}_stop_button", None)
            pause_button = getattr(self, f"{station_name}_pause_button", None)
            frequency_menu = getattr(self, f"{station_name}_frequency_menu", None)
            
            if start_button:
                start_button.config(state="disabled")
            if stop_button:
                stop_button.config(state="normal")
            if pause_button:
                pause_button.config(state="normal")
            if frequency_menu:
                frequency_menu.config(state="disabled")
            
            # 初始化該工位的數據儲存結構
            self.station_data[station_name] = {
                "time_data": [],
                "temperature_data": [],
                "power_data": []
            }
            
            # 確保圖表初始化
            ax_temp = getattr(self, f"{station_name}_ax_temp", None)
            ax_power = getattr(self, f"{station_name}_ax_power", None)
            if ax_temp and ax_power:
                ax_temp.clear()
                ax_power.clear()


            # 設置收集狀態
            self.collecting[station_name] = True

            # ===== 新增：頁籤加上 [] =====
            for idx in range(self.notebook.index("end")):
                tab_text = self.notebook.tab(idx, "text").replace(" ", "")
                if tab_text == station_name:
                    self.notebook.tab(idx, text=f"[{tab_text}]")
                    break
            # ===========================
            
            # 取得工位對應的頻道和 IP
            channels = self.parse_channels(self.channel_number[station_name])
            station_index = int(station_name[-1]) - 1
            pw_ip = f"192.168.1.{station_index + 2}"
            
            # 初始化該工位的數據儲存結構
            self.station_data[station_name] = {
                "time_data": [],
                "temperature_data": [],
                "power_data": []
            }
            
            # 啟動數據收集執行緒
            collection_thread = threading.Thread(
                target=self.collect_data,
                args=(pw_ip, channels, station_name),
                daemon=True
            )
            collection_thread.start()
            log_info(f"{station_name} 開始收集數據")
            # 啟動圖表更新
            self.start_plot_update(station_name)
        except Exception as e:
            self.handle_data_collection_error(station_name, str(e))
            self.stop_collection(station_name)

    def stop_collection(self, station_name):
        """停止指定工位的數據收集"""
        self.collecting[station_name] = False
        # 清除數據
        self.station_data[station_name] = {
            "time_data": [],
            "temperature_data": [],
            "power_data": []
        }
        
        # 清除圖表
        ax_temp = getattr(self, f"{station_name}_ax_temp", None)
        ax_power = getattr(self, f"{station_name}_ax_power", None)
        if ax_temp and ax_power:
            ax_temp.clear()
            ax_power.clear()
            canvas = getattr(self, f"{station_name}_canvas", None)
            if canvas:
                canvas.draw()

        # ===== 新增：頁籤移除 [] =====
        for idx in range(self.notebook.index("end")):
            tab_text = self.notebook.tab(idx, "text")
            if tab_text.replace("[", "").replace("]", "").replace(" ", "") == station_name:
                self.notebook.tab(idx, text=station_name)
                break
        # ===========================
        
        # 啟用 start 按鈕，禁用 stop 按鈕和 pause 按鈕
        start_button = getattr(self, f"{station_name}_start_button", None)
        stop_button = getattr(self, f"{station_name}_stop_button", None)
        pause_button = getattr(self, f"{station_name}_pause_button", None)
        frequency_menu = getattr(self, f"{station_name}_frequency_menu", None)
        
        if start_button:
            start_button.config(state="normal")
        if stop_button:
            stop_button.config(state="disabled")
        if pause_button:
            pause_button.config(state="disabled")
            pause_button.config(text="暫停")  # 重設暫停按鈕文字
        if frequency_menu:
            frequency_menu.config(state="normal")
        
        # 關閉正在寫入的 CSV 檔案
        file_name_entry = getattr(self, f"{station_name}_file_name_entry", None)
        if file_name_entry:
            file_name = file_name_entry.get()
            if file_name:
                file_path = os.path.join(self.file_path, file_name)
            if os.path.exists(file_path):
                try:
                    with open(file_path, 'r+'):  # Open and close to ensure it's not locked
                        pass
                except Exception as e:
                    log_error(f"Error closing CSV file for {station_name}: {e}")
        # 重設暫停狀態
        self.pause_plot = False
        log_info(f"{station_name} 停止收集數據")

    def collect_data(self, pw_ip, channels, station_name):
        """收集數據並保存到 CSV 文件"""
        try:
            # 檢查 PW3335 連線
            pw = self.pw3335_instances.get(pw_ip)
            if not pw:
                try:
                    pw = PW3335(pw_ip)
                    pw.connect()
                    self.pw3335_instances[pw_ip] = pw
                except Exception as e:
                    self.show_error_dialog("設備錯誤", f"{station_name} 的 PW3335 未連線")
                    self.stop_collection_with_error(station_name, "PW3335 未連線")
                    return

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name = f"{self.file_path}/{timestamp}_{station_name}.csv"
            
            file_name_entry = getattr(self, f"{station_name}_file_name_entry", None)
            file_name_entry.config(state="normal")
            file_name_entry.delete(0, tk.END)
            file_name_entry.insert(0, os.path.basename(file_name))
            file_name_entry.config(state="readonly")
            
            # 取得該工位的 alias 設定
            ch_aliases = getattr(self, f"{station_name}_ch_aliases", {})
            
            with open(file_name, mode="a", newline="", buffering=1) as file:
                writer = csv.writer(file)
                # 寫入機種資訊行
                model_entry = getattr(self, f"{station_name}_model_entry", None)
                model_value = model_entry.get() if model_entry else "NA"
                writer.writerow([f"DateTime: {timestamp}"])
                writer.writerow([f"Model: {model_value}"])
                # 寫入標題行：使用 alias 或預設頻道名稱
                header = ["Date", "Time"]
                for i, ch in enumerate(channels):
                    # 取得對應的 alias entry
                    alias_entry = ch_aliases.get(ch)
                    if alias_entry and alias_entry.get().strip():
                        # 如果有設定 alias 且不為空，使用 alias
                        header.append(f"{alias_entry.get().strip()}")
                    else:
                        # 否則使用預設的頻道名稱
                        header.append(f"CH{ch}")
                
                # 加入電力頻道
                header.extend(["U(V)", "I(A)", "P(W)", "WP(Wh)"])
                writer.writerow(header)

                frequency_var = getattr(self, f"{station_name}_frequency_var", None)
                if not frequency_var:
                    raise AttributeError(f"Frequency variable for {station_name} is not defined.")

                # 取得該工位的 checkbox 狀態（僅用於圖表顯示）
                #checkboxes = getattr(self, f"{station_name}_checkboxes", {})

                while self.collecting.get(station_name, False):
                    now = datetime.now()
                    temperatures = []
                    for ch in channels:
                        value = self.gx20_data_dict.get(ch, {}).get("value")
                        temperatures.append(value if value is not None and value <= 999 else None)

                    power_data = [None] * 4
                    try:
                        power_data = pw.query_data()[:4]
                    except Exception as e:
                        log_error(f"collect_data.pw.query_data()發生錯誤: {e} at {pw_ip}")

                    # 寫入所有數據：時間 + 20個溫度 + 4個電力值
                    writer.writerow([now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")] + 
                                temperatures + power_data)
                    file.flush()

                    # 更新數據存儲
                    self.station_data[station_name]["temperature_data"].append((now, {ch: temp for ch, temp in zip(channels, temperatures)}))
                    self.station_data[station_name]["power_data"].append((now, power_data[2]))  # P(W)

                    time.sleep(frequency_var.get())
        except Exception as e:
            self.handle_data_collection_error(station_name, str(e))
            self.stop_collection_with_error(station_name, str(e))
            #print(f"Data collection error for {station_name}: {e}")
            log_error(f"collect_data error for {station_name}: {e}")
            # 關閉 CSV 檔案
            file.close()


    def on_closing(self):
        """關閉程式時斷開所有 PW3335 連線"""
        # 檢查是否有工位正在啟動
        active_stations = [station for station, is_collecting in self.collecting.items() if is_collecting]
        if active_stations:
            tk.messagebox.showwarning(
                "警告", 
                f"以下工位正在收集數據，請先停止數據收集再退出程序：\n{', '.join(active_stations)}"
            )
            log_info(f"以下工位正在收集數據，請先停止數據收集再退出程序：\n{', '.join(active_stations)}")
        else:
            # 斷開所有 PW3335 連線
            for pw in self.pw3335_instances.values():
                try:
                    pw.disconnect()
                except:
                    pass
            # 關閉所有開啟的 CSV 檔案
            for station_name in self.station_data.keys():
                try:
                    file_name_entry = getattr(self, f"{station_name}_file_name_entry", None)
                    if file_name_entry:
                        file_name = file_name_entry.get()
                        if file_name:
                            file_path = os.path.join(self.file_path, file_name)
                            if os.path.exists(file_path):
                                with open(file_path, 'r+'):  # Open and close to ensure it's not locked
                                    pass
                except Exception as e:
                    log_error(f"Error closing CSV file for {station_name}: {e}")
            self.root.destroy()
            log_info("程式已關閉")

    def gx20_data_updater(self):
        """持續連線GX20並更新所有頻道數據到 self.gx20_data_dict，並即時更新各工位溫度顯示"""
        while True:
            try:
                self.gx20_instance.GX20GetData()
                self.gx20_data_dict = self.gx20_instance.valid_data.copy()
                self.gx20_connected = True
                

                
                # 依照每個工位的頻道設定，更新溫度顯示
                for i in range(1, 7):
                    station_name = f"工位{i}"

                    # 新增：更新 PLOT 頁面的頻道讀值
                    self.root.after(0, self.update_plot_channel_values, station_name)
                    
            except Exception as e:
                self.handle_gx20_connection_error()
                print(f"GX20 connection error: {e}")
                log_error(f"GX20 connection error: {e}")
            time.sleep(2)

    # 新增：更新 PLOT 頁面頻道名稱的方法
    def update_plot_channel_values(self, station_name):
        checkboxes = getattr(self, f"{station_name}_checkboxes", None)
        
        """更新 PLOT 頁面的頻道讀值顯示"""
        if station_name not in self.plot_channel_labels:
            return
            
        for channel, label in self.plot_channel_labels[station_name]:
            value = self.gx20_data_dict.get(channel, {}).get("value")
            if value is not None and value <= 999:
                label.config(text=f"{value:.1f}")
                #checkboxes[channel].set(1)  # 將對應的 checkbox 設為enabled
            else:
                label.config(text="--")
                checkboxes[channel].set(0)  # 將對應的 checkbox 設為未選中
                
    def show_error_dialog(self, title: str, message: str):
        """顯示錯誤對話框"""
        tk.messagebox.showerror(title, message)
        print(f"{title}: {message}")
        log_error(f"{title}: {message}")

    def handle_gx20_connection_error(self):
        """處理 GX20 連線錯誤"""
        self.gx20_connected = False
        self.root.title("SAMPO GX20/PW3335 Data Collection - GX20無法連線")
        self.show_error_dialog("連線錯誤","GX20 紀錄器連線失敗。")

    def handle_data_collection_error(self, station_name: str, error_msg: str):
        """處理數據收集錯誤"""
        self.collecting[station_name] = False
        self.show_error_dialog("數據收集錯誤", f"工位 {station_name} 數據收集錯誤：{error_msg}")

    def stop_collection_with_error(self, station_name, error_msg):
        """停止數據收集並恢復按鈕狀態"""
        # 停止數據收集
        self.collecting[station_name] = False
        
        # 恢復按鈕狀態
        start_button = getattr(self, f"{station_name}_start_button", None)
        stop_button = getattr(self, f"{station_name}_stop_button", None)
        pause_button = getattr(self, f"{station_name}_pause_button", None)
        
        if start_button:
            start_button.config(state="normal")
        if stop_button:
            stop_button.config(state="disabled")
        if pause_button:
            pause_button.config(state="disabled")
            pause_button.config(text="暫停")
        
        # 重設暫停狀態
        self.pause_plot = False
        
        # 顯示錯誤訊息
        self.show_error_dialog("數據收集停止", f"{station_name} 數據收集已停止：{error_msg}")

    def update_plot(self, frame=None, station_name=None):
        """更新圖表數據"""
        # 檢查 station_name
        if station_name is None:
            if hasattr(frame, 'station_name'):
                station_name = frame.station_name
            else:
                return  # 靜默返回，不顯示錯誤信息
        
        if self.pause_plot:
            return

        # 檢查數據是否已經開始收集
        if not self.collecting.get(station_name, False):
            return  # 如果尚未開始收集數據，直接返回
            
        # 檢查並獲取數據
        try:
            temp_line = self.station_data[station_name]["temperature_data"]
            power_line = self.station_data[station_name]["power_data"]
            
            # 檢查是否有數據
            if not temp_line or not power_line:
                return  # 靜默返回，不顯示錯誤信息
                
        except (KeyError, AttributeError) as e:
            return  # 靜默返回，不顯示錯誤信息

        # 取得頻率與 X 軸範圍
        frequency_var = getattr(self, f"{station_name}_frequency_var", None)
        x_axis_range_var = getattr(self, f"{station_name}_x_axis_range_var", None)
        if not frequency_var or not x_axis_range_var:
            return

        # 設定 X 軸範圍
        axis_range = {
            "30min": timedelta(minutes=30),
            "3hrs": timedelta(hours=3),
            "12hrs": timedelta(hours=12),
            "24hrs": timedelta(hours=24),
        }.get(x_axis_range_var.get(), timedelta(minutes=30))

        # 檢查並獲取數據
        try:
            temp_line = self.station_data[station_name]["temperature_data"]
            power_line = self.station_data[station_name]["power_data"]
        except (KeyError, AttributeError) as e:
            self.show_error_dialog("數據錯誤",f"無法訪問 {station_name} 的數據，請檢查數據收集狀態。")
            return

        # 檢查是否有數據
        if not temp_line or not power_line:
            self.show_error_dialog("數據錯誤", f"No data in temp_line or power_line for {station_name}")
            return

        # 設定 X 軸範圍
        try:
            latest_time = temp_line[-1][0]
            start_time = latest_time - axis_range
        except IndexError:
            print(f"No valid time data for {station_name}")
            log_error(f"app.update_plot: No valid time data for {station_name}")
            return

        # 獲取 ax_temp 和 ax_power
        ax_temp = getattr(self, f"{station_name}_ax_temp", None)
        ax_power = getattr(self, f"{station_name}_ax_power", None)
        if not ax_temp or not ax_power:
            print(f"Error: ax_temp or ax_power not initialized for {station_name}")
            log_error(f"app.update_plot: ax_temp or ax_power not initialized for {station_name}")
            return

        # 清除圖表
        ax_temp.clear()
        ax_power.clear()

        # 更新溫度折線圖
        checkboxes = getattr(self, f"{station_name}_checkboxes", {})
        ch_aliases = getattr(self, f"{station_name}_ch_aliases", {})

        legend_labels = []
        for ch, var in checkboxes.items():
                if var.get() == 1:  # 如果該頻道被選中
                    alias_entry = ch_aliases.get(ch)
                    alias = alias_entry.get().strip() if alias_entry and alias_entry.get().strip() else f"CH{ch}"
                    legend_labels.append(alias)
                    times = [data[0] for data in temp_line if data[0] >= start_time]
                    values = [data[1].get(ch, None) for data in temp_line if data[0] >= start_time]
                    values = [v for v in values if v is not None]  # 過濾掉 None
                    if times and values:  # 確保有數據
                        ax_temp.plot(times, values, label=alias)
                    else:
                        print(f"No valid data for {ch} in {station_name}")
                        log_error(f"app.ipdate_plot: No valid data for {ch} in {station_name}")
                        continue  # 如果沒有數據，則跳過該頻道      
        # 添加 legend
        ax_temp.legend(legend_labels, loc="upper left")

        # 設置 Y 軸標籤和網格
        ax_temp.set_ylabel("Temperature (°C)")
        ax_temp.grid(True)

        # 更新功率折線圖
        times = [data[0] for data in power_line if data[0] >= start_time]
        powers = [data[1] for data in power_line if data[0] >= start_time]
        ax_power.plot(times, powers, color="red")
        ax_power.set_xlabel("Time")
        ax_power.set_ylabel("Power (W)")
        ax_power.grid(True)

        # 設置 X 軸範圍和格式
        ax_temp.set_xlim(start_time, latest_time)
        ax_power.set_xlim(start_time, latest_time)
        ax_power.xaxis.set_major_formatter(mdates.DateFormatter("%d-%H:%M"))
        ax_power.tick_params(axis="x", rotation=45)

        # 更新圖表
        canvas = getattr(self, f"{station_name}_canvas", None)
        if canvas:
            canvas.draw()

    def toggle_pause_plot(self, station_name):
        """暫停或恢復圖表更新"""
        pause_button = getattr(self, f"{station_name}_pause_button", None)
        self.pause_plot = not self.pause_plot
        if self.pause_plot:
            pause_button.config(text="恢復")

            # 填入目前 X 軸的資料到 start_date, start_time, end_date, end_time
            try:
                x_start, x_end = self.get_x_axis_range(station_name)
                start_date_entry = getattr(self, f"{station_name}_start_date_entry", None)
                start_time_entry = getattr(self, f"{station_name}_start_time_entry", None)
                end_date_entry = getattr(self, f"{station_name}_end_date_entry", None)
                end_time_entry = getattr(self, f"{station_name}_end_time_entry", None)

                if start_date_entry and start_time_entry and end_date_entry and end_time_entry:
                    start_date_entry.delete(0, tk.END)
                    start_date_entry.insert(0, x_start.strftime('%Y-%m-%d'))
                    start_time_entry.delete(0, tk.END)
                    start_time_entry.insert(0, x_start.strftime('%H:%M'))
                    end_date_entry.delete(0, tk.END)
                    end_date_entry.insert(0, x_end.strftime('%Y-%m-%d'))
                    end_time_entry.delete(0, tk.END)
                    end_time_entry.insert(0, x_end.strftime('%H:%M'))
            except AttributeError as e:
                print(f"讀不到 start/end date/time 欄位: {e}")
                self.show_error_dialog("錯誤", f"讀不到 start/end date/time 欄位: {e}")
        else:
            pause_button.config(text="暫停")
            self.update_plot(station_name)  # 恢復時立即更新一次圖表
    
    def get_x_axis_range(self, station_name):
            """根據選擇的 X 軸範圍返回時間範圍"""
            now = datetime.now()
            range_mapping = {
                "30min": timedelta(minutes=30),
                "3hrs": timedelta(hours=3),
                "12hrs": timedelta(hours=12),
                "24hrs": timedelta(hours=24),
            }
            x_axis_range_var = getattr(self, f"{station_name}_x_axis_range_var", None)
            if x_axis_range_var is None:
                raise AttributeError(f"x_axis_range_var for {station_name} is not defined.")
            selected_range = range_mapping.get(x_axis_range_var.get(), timedelta(minutes=30))
            return now - selected_range, now
    
    
    def calculate_avg_temp(self, station_name):
        """計算指定時間範圍內的平均溫度，僅針對選定的頻道"""
        try:
            # 動態獲取對應工位的日期和時間輸入框
            start_date_entry = getattr(self, f"{station_name}_start_date_entry", None)
            start_time_entry = getattr(self, f"{station_name}_start_time_entry", None)
            end_date_entry = getattr(self, f"{station_name}_end_date_entry", None)
            end_time_entry = getattr(self, f"{station_name}_end_time_entry", None)
            calculate_text = getattr(self, f"{station_name}_calculate_text", None)
            checkboxes = getattr(self, f"{station_name}_checkboxes", {})

            if not all([start_date_entry, start_time_entry, end_date_entry, end_time_entry, calculate_text]):
                raise AttributeError(f"One or more required widgets for {station_name} are not defined.")

            # 獲取開始和結束時間
            start_datetime = pd.to_datetime(f"{start_date_entry.get()} {start_time_entry.get()}")
            end_datetime = pd.to_datetime(f"{end_date_entry.get()} {end_time_entry.get()}")

            if start_datetime >= end_datetime:
                tk.messagebox.showerror("錯誤", "開始時間必須早於結束時間")
                return

            # 篩選在指定範圍內的溫度數據
            temp_data = self.station_data[station_name]
            filtered_temps = [
                temps for time, temps in temp_data["temperature_data"]
                if start_datetime <= time <= end_datetime
            ]

            if not filtered_temps:
                tk.messagebox.showinfo("提示", "指定範圍內沒有溫度數據")
                return

            # 僅計算選定的頻道
            selected_channels = [ch for ch, var in checkboxes.items() if var.get() == 1]
            if not selected_channels:
                tk.messagebox.showinfo("提示", "未選擇任何頻道")
                return

            # 計算每個選定頻道的平均溫度
            avg_temps = {}
            for ch in selected_channels:
                channel_temps = [temps.get(ch) for temps in filtered_temps if temps.get(ch) is not None]
                if channel_temps:
                    avg_temps[ch] = sum(channel_temps) / len(channel_temps)
                else:
                    avg_temps[ch] = None

            # 顯示結果
            calculate_text.config(state="normal")
            calculate_text.delete("1.0", tk.END)
            for ch, avg_temp in avg_temps.items():
                if avg_temp is not None:
                    calculate_text.insert(tk.END, f"{ch}: {avg_temp:.1f}°C\n")
                else:
                    calculate_text.insert(tk.END, f"{ch}: --°C\n")
            calculate_text.config(state="disabled")

        except Exception as e:
            tk.messagebox.showerror("錯誤", f"計算平均溫度時發生錯誤: {e}")
            print(f"計算平均溫度時發生錯誤: {e}")


    def report_calculate(self, station_name):
        """計算報告"""
        try:
            # 動態獲取對應工位的日期和時間輸入框
            start_date_entry = getattr(self, f"{station_name}_start_date_entry", None)
            start_time_entry = getattr(self, f"{station_name}_start_time_entry", None)
            end_date_entry = getattr(self, f"{station_name}_end_date_entry", None)
            end_time_entry = getattr(self, f"{station_name}_end_time_entry", None)
            report_text = getattr(self, f"{station_name}_report_text", None)
            checkboxes = getattr(self, f"{station_name}_checkboxes", {})

            if not all([start_date_entry, start_time_entry, end_date_entry, end_time_entry, report_text]):
                raise AttributeError(f"One or more required widgets for {station_name} are not defined.")

            # 獲取開始和結束時間
            start_datetime = pd.to_datetime(f"{start_date_entry.get()} {start_time_entry.get()}")
            end_datetime = pd.to_datetime(f"{end_date_entry.get()} {end_time_entry.get()}")

            if start_datetime >= end_datetime:
                tk.messagebox.showerror("錯誤", "開始時間必須早於結束時間")
                return

            # 篩選在指定範圍內的數據
            temp_data = self.station_data[station_name]
            filtered_temps = [
                temps for time, temps in temp_data["temperature_data"]
                if start_datetime <= time <= end_datetime
            ]

            if not filtered_temps:
                tk.messagebox.showinfo("提示", "指定範圍內沒有數據")
                return

            # 僅計算選定的頻道
            selected_channels = [ch for ch, var in checkboxes.items() if var.get() == 1]
            if not selected_channels:
                tk.messagebox.showinfo("提示", "未選擇任何頻道")
                return

            # 計算每個選定頻道的平均值和標準差
            avg_temps = {}
            std_temps = {}
            for ch in selected_channels:
                channel_temps = [temps.get(ch) for temps in filtered_temps if temps.get(ch) is not None]
                if channel_temps:
                    avg_temp = sum(channel_temps) / len(channel_temps)
                    std_temp = (sum([(temp - avg_temp) ** 2 for temp in channel_temps]) / len(channel_temps)) ** 0.5
                    avg_temps[ch] = avg_temp
                    std_temps[ch] = std_temp
                else:
                    avg_temps[ch] = None
                    std_temps[ch] = None
            
            # 讀取正在紀錄的 CSV 檔案
            try:
                file_path_var = getattr(self, f"{station_name}_file_path_entry", None)
                file_name_var = getattr(self, f"{station_name}_file_name_entry", None)
                csv_file = file_path_var.get() + "/" + file_name_var.get()
                if not os.path.exists(csv_file):
                    tk.messagebox.showerror("錯誤", "報告計算用的報告計算用的CSV 檔案不存在，請確認儲存路徑是否正確")
                    log_error("報告計算用的CSV 檔案不存在，請確認儲存路徑是否正確")
                    return

                # 讀取 CSV 檔案
                df = pd.read_csv(csv_file, encoding='utf-8',skiprows=2)
            except Exception as e:
                tk.messagebox.showerror("錯誤", f"讀取報告計算用的 CSV 檔案時發生錯誤: {e}")
                print(f"讀取報告計算用的 CSV 檔案時發生錯誤: {e}")
                log_error(f"讀取報告計算用的 CSV 檔案時發生錯誤: {e}")

            # 合併 Date 和 Time 欄位為 datetime
            df['datetime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])
            
            # 檢查是否有 U(V), I(A), P(W), WP(Wh) 欄位，若無則補上並填入 0
            required_columns = ['U(V)', 'I(A)', 'P(W)', 'WP(Wh)']
            for col in required_columns:
                if col not in df.columns:
                    df[col] = 0
            # 計算 start 和 end 之間的分鐘數
            minutes_difference = int((end_datetime - start_datetime).total_seconds() / 60)
            
            # 過濾指定的日期時間範圍，並創建副本
            filtered_df = df[(df['datetime'] >= start_datetime) & (df['datetime'] <= end_datetime)].copy()

            if filtered_df.empty:
                messagebox.showinfo("結果", "指定範圍內沒有資料！")
                return

            # 計算平均值
            averages = filtered_df.mean(numeric_only=True)

            # 計算電力啟停周期
            power_column = 'P(W)'
            if power_column in filtered_df.columns:
                # 確保正確建立 power_on 欄位
                filtered_df.loc[:, 'power_on'] = filtered_df[power_column] >= 3

                # 計算啟停周期次數
                power_cycles = int(filtered_df['power_on'].astype(int).diff().fillna(0).abs().sum() // 2)

                # 計算大於等於3W和小於3W的週期數，排除頭尾兩個周期
                mask = filtered_df['power_on']
                groups = (mask != mask.shift()).cumsum()
                segments = pd.DataFrame({
                    '狀態': mask,
                    '區段編號': groups,
                    '時間': filtered_df['datetime']
                }).groupby(['區段編號', '狀態']).agg({'時間': ['min', 'max']}).reset_index()

                # 排除頭尾兩個周期
                if len(segments) > 2:
                    segments = segments.iloc[1:-1]

                # 計算每個區段的持續時間
                segments['持續時間'] = (segments[('時間', 'max')] - segments[('時間', 'min')]).dt.total_seconds()

                # 分別計算大於等於3W和小於3W的週期數與平均時間
                above_segments = segments[segments['狀態']]
                below_segments = segments[~segments['狀態']]

                above_count = len(above_segments)
                below_count = len(below_segments)

                above_avg_time = (above_segments['持續時間'].mean() / 60) if above_count > 0 else 0
                below_avg_time = (below_segments['持續時間'].mean() / 60) if below_count > 0 else 0

                # 計算百分比
                if above_avg_time + below_avg_time > 0:
                    above_percentage = (above_avg_time / (above_avg_time + below_avg_time)) * 100
                else:
                    above_percentage = 0
            else:
                power_cycles = "無法計算，缺少 P(W) 欄位"
                
            # 計算 WP(Wh) 欄位的差值
            wp_column = 'WP(Wh)'
            if wp_column in filtered_df.columns:
                wp_difference = filtered_df[wp_column].iloc[-1] - filtered_df[wp_column].iloc[0]
                
                # 使用線性法推算 24 小時的差值
                total_seconds = (filtered_df['datetime'].iloc[-1] - filtered_df['datetime'].iloc[0]).total_seconds()
                if (total_seconds > 0):
                    wp_24h_difference = (wp_difference / total_seconds) * (24 * 3600)
                else:
                    wp_24h_difference = "無法計算，時間範圍不足"
            else:
                wp_difference = "無法計算，缺少 WP(Wh) 欄位"
                wp_24h_difference = "無法計算，缺少 WP(Wh) 欄位"
            
            # 計算能耗
            vf_entry = getattr(self, f"{station_name}_vf_entry", None) # 冷凍室容積(L)
            vr_entry = getattr(self, f"{station_name}_vr_entry", None) # 冷藏室容積(L)
            fan_type = getattr(self, f"{station_name}_fan_type_var", {}).get()  # 取得風扇類型的狀態
            vf = float(vf_entry.get()) if vf_entry and vf_entry.get().isdigit() else 0
            vr = float(vr_entry.get()) if vr_entry and vr_entry.get().isdigit() else 0
            fridge_temp = 3 # 冷藏室溫度
            freezer_temp = -18 # 冷凍室溫度
            energy_calculator = EnergyCalculator()
            if isinstance(wp_24h_difference, (int, float)) and vf > 0 and vr > 0:
                daily_consumption = wp_24h_difference / 1000  # 將 Wh 轉換為 kWh
                # 計算
                results = energy_calculator.calculate(vf, vr, daily_consumption, fridge_temp, freezer_temp, fan_type)
                # 提取結果
                if results:
                    # 打印結果
                    print("冰箱能耗計算結果:")
                    for key, value in results.items():
                        print(f"{key}: {value}")
            else:
                results = None
                print("無耗電量數據,無法計算能耗")

                
            

            # 顯示結果
            report_text.delete(1.0, tk.END)  # 清空文字框
            report_text.insert(tk.END, f"統計範圍：{start_datetime} ~ {end_datetime}\n")
            report_text.insert(tk.END, "平均值計算：\n")
            for column, avg in averages.items():
                report_text.insert(tk.END, f"{column}: {avg:.2f}\n")
            report_text.insert(tk.END, f"\nON / Off 周期次數：{power_cycles}\n")
            report_text.insert(tk.END, f"On 的平均時間: {above_avg_time:.1f} 分\n" if above_count > 0 else "P(W) >= 3 的平均時間: 無資料\n")
            report_text.insert(tk.END, f"Off 的平均時間: {below_avg_time:.1f} 分\n" if below_count > 0 else "P(W) < 3 的平均時間: 無資料\n")
            report_text.insert(tk.END, f"On / Off 百分比: {above_percentage:.2f}%\n")
            report_text.insert(tk.END, f"\n電力消耗：{wp_difference:.2f} w / {minutes_difference} 分\n")
            report_text.insert(tk.END, f"24 小時電力消耗：{wp_24h_difference:.1f} w\n")
            report_text.insert(tk.END, f"\n能耗計算：\n")
            if results:
                for key, value in results.items():
                    report_text.insert(tk.END, f"{key}: {value}\n")
            else:
                report_text.insert(tk.END, "無法計算能耗，請檢查數據\n")

        except Exception as e:
            tk.messagebox.showerror("錯誤", f"計算報告時發生錯誤: {e}")
            #print(f"計算報告時發生錯誤: {e}")
            log_error(f"計算報告時發生錯誤: {e}")

    def start_plot_update(self, station_name):
        """啟動圖表更新的定時器"""
        if self.collecting.get(station_name, False) and not self.pause_plot:
            self.update_plot(station_name)
        # 每5秒調用一次
        self.root.after(5000, self.start_plot_update, station_name)

if __name__ == "__main__":
    log_info("程式啟動")
    # 支援 pyinstaller 打包時找資源
    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS  # PyInstaller 打包後會存在這個暫存路徑
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
    
    now = datetime.now()
    specific_date = datetime(2025, 12, 31)
    if now > specific_date:
        tk.messagebox.showinfo("Info", " SAMPO GX20/PW3335 Data Collection !!")
        sys.exit()

    root = tk.Tk()

    # 全局字體加大 2 號，並指定為微軟正黑體
    default_font = tkfont.nametofont("TkDefaultFont")
    default_font.configure(family="Microsoft JhengHei", size=default_font.cget("size") + 2)
    text_font = tkfont.nametofont("TkTextFont")
    text_font.configure(family="Microsoft JhengHei", size=text_font.cget("size") + 2)
    fixed_font = tkfont.nametofont("TkFixedFont")
    fixed_font.configure(family="Microsoft JhengHei", size=fixed_font.cget("size") + 2)

    # 啟動時最大化
    root.state('zoomed')

    # 指定視窗大小並置中（這段可保留，最大化時會被覆蓋）
    #w, h = 1920, 1080
    #root.geometry(f'{w}x{h}')
    root.update_idletasks()
    ws = root.winfo_screenwidth()
    hs = root.winfo_screenheight()


    AppTitle = "SAMPO GX20/PW3335 Data Collection 1_0"
    ModelID = AppTitle + now.strftime("%Y%m%d_%H%M%S")
    # 讓工作列圖示正確顯示
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(ModelID)
    # 設定執行檔的路徑變數
    exe_path = os.path.dirname(os.path.abspath(__file__))
    ico_file = exe_path + "\\GX20_PW3335.ico"
    if os.path.exists(ico_file):
        ico_path = resource_path(ico_file)
        root.iconbitmap(ico_path)  # 設定視窗圖示
    else:
        #print(f"Icon file not found: {ico_file}")
        log_info(f"Icon file not found: {ico_file}")

    root.title(AppTitle)
    app = App(root)
    root.mainloop()