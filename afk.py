import os
import sys

# Add single instance check at the very beginning
try:
    from single_instance import SingleInstance
    instance = SingleInstance("AFKTracker")
except ImportError:
    # If import fails, continue running
    pass

import time
import threading
import datetime
import getpass
import platform
from pynput import mouse, keyboard
try:
    import pygetwindow as gw
except ImportError:
    # 如果在 macOS 或 Linux 上運行，提供替代方案
    gw = None

# 導入 MongoDB 配置
from database.mongo_config import get_database
from logger_config import setup_logger
import os

# 設置記錄器
logger = setup_logger('AFKTracker')

# 記錄啟動信息
logger.info('AFK Tracking service starting...')
logger.info(f'User: {os.environ.get("USERNAME")}')
logger.info(f'Computer name: {os.environ.get("COMPUTERNAME")}')

class AFK:
    #180
    def __init__(self, idle_time=300):
        """
        初始化活動監控器
        
        參數:
            idle_time (int): 判定為離開(AFK)的閒置秒數，預設為300秒
        """
        self.idle_time = idle_time
        self.last_activity_time = time.time()
        self.is_afk = False
        self.afk_start_time = None
        self.work_start_time = time.time()
        self.current_window = self._get_current_window()
        self.username = getpass.getuser()
        self.sessions = []
        self.running = False
        self.mouse_listener = None
        self.keyboard_listener = None
        self.monitor_thread = None
        
        # 連接 MongoDB 並創建索引
        try:
            self.db = get_database()
            if 'afk' not in self.db.list_collection_names():
                self.db.create_collection('afk')
                print("已創建 'afk' collection")
                
            # 創建複合索引
            self.db.afk.create_index([
                ("timestamp", 1),
                ("username", 1),
                ("date", 1),
                ("status", 1)
            ], name="activity_tracking_index")
            
            self.mongo_connected = True
            # print("已成功連接到 MongoDB 並創建索引")
        except Exception as e:
            print(f"無法連接到 MongoDB: {e}")
            print("將只在本地存儲會話數據")
            self.mongo_connected = False
        
    def _get_current_window(self):
        """獲取當前活動視窗名稱"""
        try:
            if platform.system() == "Windows" and gw:
                active_window = gw.getActiveWindow()
                return active_window.title if active_window else "Unknown"
            else:
                return "Unknown (非Windows系統)"
        except Exception:
            return "Unknown"
    
    def _save_to_mongodb(self, session_data):
        """將會話數據存儲到 MongoDB"""
        if not self.mongo_connected:
            return
            
        try:
            # 添加時間戳用於排序和查詢
            session_data['timestamp'] = datetime.datetime.now()
            # 插入數據到 MongoDB
            self.db.afk.insert_one(session_data)
        except Exception as e:
            print(f"保存數據到 MongoDB 時出錯: {e}")
    
    def on_activity(self):
        """當檢測到活動時呼叫"""
        current_time = time.time()
        self.last_activity_time = current_time
        
        # 檢查活動狀態是否從AFK變為非AFK
        if self.is_afk:
            self.is_afk = False
            afk_end_time = datetime.datetime.now()
            afk_duration = current_time - self.afk_start_time
            
            # 記錄AFK會話
            session_data = {
                'date': datetime.datetime.now().strftime('%Y-%m-%d'),
                'username': self.username,
                'window': self.current_window,
                'type': 'afk',
                'start_time': datetime.datetime.fromtimestamp(self.afk_start_time).strftime('%H:%M:%S'),
                'end_time': afk_end_time.strftime('%H:%M:%S'),
                'duration': self._format_duration(afk_duration)
            }
            
            self.sessions.append(session_data)
            # 保存到 MongoDB
            self._save_to_mongodb(session_data)
            
            # 開始新的工作會話
            self.work_start_time = current_time
            self.current_window = self._get_current_window()
    
    def on_mouse_move(self, x, y):
        self.on_activity()
    
    def on_mouse_click(self, x, y, button, pressed):
        if pressed:  # 僅在按下時觸發，而不是釋放時
            self.on_activity()
    
    def on_mouse_scroll(self, x, y, dx, dy):
        self.on_activity()
    
    def on_key_press(self, key):
        self.on_activity()
    
    def check_afk_status(self):
        """檢查使用者是否已離開(AFK)"""
        while self.running:
            current_time = time.time()
            idle_duration = current_time - self.last_activity_time
            current_window = self._get_current_window()
            
            # 每秒記錄一次活動狀態
            if not self.is_afk:  # 如果用戶不是AFK狀態
                interim_end_time = datetime.datetime.now()
                interim_duration = 5  # 固定為1秒
                
                session_data = {
                    'date': datetime.datetime.now().strftime('%Y-%m-%d'),
                    'username': self.username,
                    'user_name': self.username,
                    'window': self.current_window,
                    'type': 'work',
                    'status': 'Work',
                    'start_time': datetime.datetime.fromtimestamp(current_time - 5).strftime('%H:%M:%S'),
                    'end_time': interim_end_time.strftime('%H:%M:%S'),
                    'duration': self._format_duration(interim_duration),
                    'is_heartbeat': True
                }
                
                # 保存到 MongoDB
                self._save_to_mongodb(session_data)
            
            # 檢查是否已閒置超過閾值
            if not self.is_afk and idle_duration >= self.idle_time:
                self.is_afk = True
                self.afk_start_time = current_time
                
                # 記錄開始AFK狀態
                session_data = {
                    'date': datetime.datetime.now().strftime('%Y-%m-%d'),
                    'username': self.username,
                    'user_name': self.username,
                    'window': self.current_window,
                    'type': 'afk',
                    'status': 'AFK',
                    'start_time': datetime.datetime.fromtimestamp(current_time).strftime('%H:%M:%S'),
                    'end_time': datetime.datetime.fromtimestamp(current_time + 1).strftime('%H:%M:%S'),
                    'duration': self._format_duration(1),
                    'is_heartbeat': True
                }
                
                self._save_to_mongodb(session_data)
            elif self.is_afk:  # 如果用戶處於AFK狀態
                # 每秒記錄AFK狀態
                session_data = {
                    'date': datetime.datetime.now().strftime('%Y-%m-%d'),
                    'username': self.username,
                    'user_name': self.username,
                    'window': self.current_window,
                    'type': 'afk',
                    'status': 'AFK',
                    'start_time': datetime.datetime.fromtimestamp(current_time).strftime('%H:%M:%S'),
                    'end_time': datetime.datetime.fromtimestamp(current_time + 1).strftime('%H:%M:%S'),
                    'duration': self._format_duration(1),
                    'is_heartbeat': True
                }
                
                self._save_to_mongodb(session_data)
            
            time.sleep(5)  # 每秒檢查一次
    
    def start(self):
        """開始監控使用者活動"""
        if self.running:
            return
            
        self.running = True
        
        # 啟動鍵盤和滑鼠監聽器
        self.mouse_listener = mouse.Listener(
            on_move=self.on_mouse_move,
            on_click=self.on_mouse_click,
            on_scroll=self.on_mouse_scroll
        )
        self.keyboard_listener = keyboard.Listener(on_press=self.on_key_press)
        
        self.mouse_listener.start()
        self.keyboard_listener.start()
        
        # 啟動監控線程
        self.monitor_thread = threading.Thread(target=self.check_afk_status)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
        
        # print(f"活動監控已啟動。閒置{self.idle_time}秒後將判定為AFK。")
    
    def stop(self):
        """停止監控使用者活動"""
        if not self.running:
            return
            
        self.running = False
        
        # 停止監聽器
        if self.mouse_listener:
            self.mouse_listener.stop()
        if self.keyboard_listener:
            self.keyboard_listener.stop()
        
        # 記錄最後一個會話
        current_time = time.time()
        end_time = datetime.datetime.now()
        
        if self.is_afk:
            # 記錄AFK會話
            afk_duration = current_time - self.afk_start_time
            session_data = {
                'date': datetime.datetime.now().strftime('%Y-%m-%d'),
                'username': self.username,
                'window': self.current_window,
                'type': 'afk',
                'start_time': datetime.datetime.fromtimestamp(self.afk_start_time).strftime('%H:%M:%S'),
                'end_time': end_time.strftime('%H:%M:%S'),
                'duration': self._format_duration(afk_duration)
            }
            
            self.sessions.append(session_data)
            # 保存到 MongoDB
            self._save_to_mongodb(session_data)
        else:
            # 記錄工作會話
            work_duration = current_time - self.work_start_time
            session_data = {
                'date': datetime.datetime.now().strftime('%Y-%m-%d'),
                'username': self.username,
                'window': self.current_window,
                'type': 'work',
                'start_time': datetime.datetime.fromtimestamp(self.work_start_time).strftime('%H:%M:%S'),
                'end_time': end_time.strftime('%H:%M:%S'),
                'duration': self._format_duration(work_duration)
            }
            
            self.sessions.append(session_data)
            # 保存到 MongoDB
            self._save_to_mongodb(session_data)
        
        print("活動監控已停止。")
    
    def get_sessions(self):
        """獲取所有記錄的會話"""
        return self.sessions
    
    def is_user_afk(self):
        """檢查使用者目前是否離開"""
        return self.is_afk

    def _format_duration(self, seconds):
        """將秒數轉換為 HH:MM:SS 格式字串"""
        hours = int(seconds) // 3600
        minutes = (int(seconds) % 3600) // 60
        seconds = int(seconds) % 60
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


# 使用示例
if __name__ == "__main__":
    # 建立活動監控器，設定閒置時間為60秒
    monitor = AFK()
    try:
        monitor.start()
        while True:
            time.sleep(60)
            # print(f"目前狀態: {'AFK' if monitor.is_user_afk() else '工作中'}")
    except KeyboardInterrupt:
        monitor.stop()
        # print("\n會話記錄:")
        # for session in monitor.get_sessions():
            # print(f"{session['type'].upper()}: {session['start_time']} - {session['end_time']} ({session['duration']}秒) - {session['window']}")