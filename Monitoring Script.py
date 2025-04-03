import os
import sys

# Add single instance check at the very beginning
try:
    from single_instance import SingleInstance
    instance = SingleInstance("ActivityMonitor")
except ImportError:
    # If import fails, continue running
    pass

import time
import os
import psutil
import sqlite3
import pandas as pd
import zipfile
from datetime import datetime, timedelta
import win32gui
import win32process
import win32api
import win32con
import sys
import traceback
import subprocess
from database.mongo_config import get_database
from logger_config import setup_logger
import os

# 設置記錄器
logger = setup_logger('ActivityMonitor')

# 記錄啟動信息
logger.info('Activity monitoring service starting...')
logger.info(f'User: {os.environ.get("USERNAME")}')
logger.info(f'Computer name: {os.environ.get("COMPUTERNAME")}')

# Add this function at the beginning of your file (after imports)
def restart_script():
    """
    Restarts the script when an unhandled exception occurs
    """
    print("Restarting script due to error...")
    python_executable = sys.executable
    script_path = os.path.abspath(__file__)
    subprocess.Popen([python_executable, script_path])
    sys.exit(0)  # Exit the current process

def cleanup_old_records():
    """
    Maintains database hygiene by removing records older than 7 days.
    Only keeps the most recent 7 days of records in the database.
    """
    try:
        db = get_database()
        
        # Calculate the date 7 days ago
        seven_days_ago = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        
        # Delete older records from activities collection
        activities_result = db.activities.delete_many({'date': {'$lt': seven_days_ago}})
        
        # Delete older records from user_idle_times collection
        idle_times_result = db.user_idle_times.delete_many({'date': {'$lt': seven_days_ago}})
        
        # If you have any other collections that need cleaning, add them here
        
        print(f"Database cleanup complete: Removed {activities_result.deleted_count} activity records and "
              f"{idle_times_result.deleted_count} idle time records older than {seven_days_ago}")
        
    except Exception as e:
        print(f"Error during database cleanup: {e}")
        logger.error(f"Database cleanup error: {e}")

# Function to get workstation name
def get_workstation_name():
    return os.environ.get('COMPUTERNAME', 'unknown')

# Function to get user name
def get_user_name():
    return psutil.users()[0].name

# Function to get logon time
def get_logon_time():
    sessions = psutil.users()
    if sessions:
        return datetime.fromtimestamp(sessions[0].started)
    else:
        return datetime.now()

# Function to get current time
def get_current_time():
    return datetime.now()  # Return datetime object

# Function to get idle time
def get_idle_time(user_name, previous_max_idle="00:00:00"):
    """
    Calculate system idle time based on last input for specific user.
    Continues counting from previous max idle time.
    
    Args:
        user_name (str): Windows用戶名稱
        previous_max_idle (str): 該用戶上一次記錄的最大閒置時間 (HH:MM:SS格式)
        
    Returns:
        str: 閒置時間，格式為 HH:MM:SS
    """
    try:
        # 讀取使用者閒置時間記錄
        idle_records = load_user_idle_times()
        
        # 如果有該用戶的記錄，使用其記錄的閒置時間
        if user_name in idle_records:
            previous_max_idle = idle_records[user_name]
        
        # 轉換previous max idle time為秒數
        prev_h, prev_m, prev_s = map(int, previous_max_idle.split(':'))
        previous_idle_seconds = prev_h * 3600 + prev_m * 60 + prev_s
        
        # 獲取最後輸入資訊（毫秒）
        last_input = win32api.GetLastInputInfo()
        current_tick = win32api.GetTickCount()
        
        # 計算目前閒置時間（秒）
        current_idle_seconds = (current_tick - last_input) / 1000.0
        
        # 只有超過30秒才計為閒置
        if current_idle_seconds < 30:
            return previous_max_idle if previous_idle_seconds > 0 else "00:00:00"
        
        # 從上次閒置時間繼續計數
        total_idle_seconds = previous_idle_seconds + 1
        
        # 轉換為timedelta以正確格式化
        idle_duration = timedelta(seconds=int(total_idle_seconds))
        
        # 格式化為 HH:MM:SS
        idle_time = str(idle_duration)
        if '.' in idle_time:
            idle_time = idle_time.split('.')[0]
            
        # 更新使用者閒置時間記錄
        save_user_idle_time(user_name, idle_time)
            
        return idle_time
        
    except Exception as e:
        print(f"Error calculating idle time for user {user_name}: {e}")
        return previous_max_idle if previous_max_idle != "00:00:00" else "00:00:00"

def load_user_idle_times():
    """
    從MongoDB載入所有使用者的閒置時間記錄，每天重置
    """
    try:
        db = get_database()
        
        # Get today's date
        today = datetime.now().strftime('%Y-%m-%d')
        
        # Query MongoDB for today's idle records
        records = db.user_idle_times.find({'date': today})
        
        # Convert to dictionary format {user_name: idle_time}
        return {doc['user_name']: doc['idle_time'] for doc in records}
        
    except Exception as e:
        print(f"Error loading user idle times from MongoDB: {e}")
        return {}

def save_user_idle_time(user_name, idle_time):
    """
    保存使用者的閒置時間到MongoDB，按日期保存
    """
    try:
        db = get_database()
        today = datetime.now().strftime('%Y-%m-%d')
        
        # Update or insert user idle time document
        db.user_idle_times.update_one(
            {
                'user_name': user_name,
                'date': today
            },
            {
                '$set': {
                    'idle_time': idle_time,
                    'last_updated': datetime.now()
                }
            },
            upsert=True
        )
        
    except Exception as e:
        print(f"Error saving idle time to MongoDB for user {user_name}: {e}")

# Function to get active application info
def get_active_application_info():
    try:
        hwnd = win32gui.GetForegroundWindow()
        # 檢查窗口句柄是否有效
        if (hwnd == 0):
            return "System_Locked", "Windows鎖定畫面", "系統", datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        # 檢查PID是否為負數或無效值
        if (pid <= 0):
            return "System_Locked", "Windows鎖定畫面", "系統", datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
        process = psutil.Process(pid)
        return process.name(), win32gui.GetWindowText(hwnd), process.exe(), datetime.fromtimestamp(process.create_time()).strftime("%Y-%m-%d %H:%M:%S")
    except (psutil.NoSuchProcess, ValueError, Exception) as e:
        print(f"無法獲取活動應用程式信息: {e}")
        return "Unknown", "Unknown", "Unknown", datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Function to get computer boot time
def get_boot_time():
    try:
        boot_time_timestamp = psutil.boot_time()
        boot_time = datetime.fromtimestamp(boot_time_timestamp)
        return boot_time.strftime("%Y-%m-%d %H:%M:%S")
    except Exception as e:
        print(f"Error getting boot time: {e}")
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Fallback to current time

# Function to log data to the database
def log_to_database(workstation, user, logon_time, logoff_time, idle_time, active_time, 
                   app_name, app_title, app_path, total_time, boot_time, app_start_time,
                   sum_time, system_working_time):
    try:
        db = get_database()
        
        # Create activity document
        activity = {
            'workstation_name': workstation,
            'user_name': user,
            'logon_time': logon_time,
            'logoff_time': logoff_time,
            'idle_time': idle_time,
            'active_time': active_time,
            'app_name': app_name,
            'app_title': app_title,
            'app_path': app_path,
            'total_time': total_time,
            'boot_time': boot_time,
            'app_start_time': app_start_time,
            'sum_time': sum_time, 
            'system_working_time': system_working_time,
            'date': datetime.now().strftime('%Y-%m-%d'),
            'created_at': datetime.now()
        }
        
        # Insert into MongoDB
        db.activities.insert_one(activity)
        
    except Exception as e:
        print(f"MongoDB error: {e}")

# Function to log data to an Excel file
# def log_to_excel(workstation, user, logon_time, logoff_time, idle_time, active_time, app_name, app_title, app_path, total_time, boot_time, app_start_time, sum_time, system_working_time):
#     try:
#         log_data = {
#             'Workstation Name': [workstation],
#             'User Name': [user],
#             'Logon Time': [logon_time],
#             'Logoff Time': [logoff_time],
#             'Idle Time': [idle_time],
#             'Active Time': [active_time],
#             'App Name': [app_name],
#             'App Title': [app_title],
#             'App Path': [app_path],
#             'Total Time': [total_time],
#             'Boot Time': [boot_time],
#             'App Start Time': [app_start_time],
#             'Sum Time': [sum_time],
#             'System Working Time': [system_working_time]
#         }
#         df = pd.DataFrame(log_data)
        
#         # Define Excel file path
#         excel_dir = os.path.join(os.path.dirname(__file__), 'data')
#         os.makedirs(excel_dir, exist_ok=True)
#         file_name = os.path.join(excel_dir, f"log_file_{datetime.now().strftime('%Y_%m_%d')}.xlsx")
        
#         if os.path.exists(file_name):
#             try:
#                 # Try to read existing file
#                 existing_df = pd.read_excel(file_name, engine='openpyxl')
#                 # Append new data
#                 df = pd.concat([existing_df, df], ignore_index=True)
#             except Exception as read_error:
#                 print(f"Error reading Excel file: {str(read_error)}")
#                 # If file is corrupted, rename it and create a new one
#                 corrupted_file = os.path.join(excel_dir, f"corrupted_log_{datetime.now().strftime('%Y_%m_%d_%H%M%S')}.xlsx")
#                 try:
#                     os.rename(file_name, corrupted_file)
#                     print(f"Corrupted file renamed to {corrupted_file}")
#                 except Exception as rename_error:
#                     print(f"Failed to rename corrupted file: {str(rename_error)}")
#                     # If rename fails, create a unique filename
#                     file_name = os.path.join(excel_dir, f"log_file_{datetime.now().strftime('%Y_%m_%d_%H%M%S')}.xlsx")
        
#         # Save the dataframe with explicit engine specification
#         df.to_excel(file_name, index=False, engine='openpyxl')
            
#     except Exception as e:
#         print(f"Error logging to Excel: {str(e)}")
#         # Consider logging to a backup CSV file if Excel fails completely
#         try:
#             csv_file = os.path.join(excel_dir, f"log_file_{datetime.now().strftime('%Y_%m_%d')}.csv")
#             df.to_csv(csv_file, index=False)
#             print(f"Data logged to CSV backup: {csv_file}")
#         except Exception as csv_error:
#             print(f"Failed to create CSV backup: {str(csv_error)}")

def load_existing_app_usage(table_name):
    """Load existing app usage records for today and get maximum cumulative time from MongoDB"""
    try:
        db = get_database()
        
        # Get today's date
        today = datetime.now().strftime('%Y-%m-%d')
        
        # Query MongoDB for today's records
        pipeline = [
            {
                '$match': {
                    'date': today
                }
            },
            {
                '$group': {
                    '_id': {
                        'app_name': '$app_name',
                        'app_title': '$app_title',
                        'app_path': '$app_path'
                    },
                    'max_sum_time': {'$max': '$sum_time'}
                }
            }
        ]
        
        results = db.activities.aggregate(pipeline)
        
        existing_usage = {}
        for record in results:
            app_name = record['_id']['app_name']
            app_title = record['_id']['app_title']
            app_path = record['_id']['app_path']
            max_sum_time = record['max_sum_time']
            
            if max_sum_time:
                # Convert time string to seconds
                time_parts = max_sum_time.split(':')
                total_seconds = int(time_parts[0]) * 3600 + int(time_parts[1]) * 60 + int(time_parts[2])
                
                existing_usage[app_name] = {
                    'total_time': total_seconds,
                    'title': app_title,
                    'path': app_path
                }
        
        return existing_usage
        
    except Exception as e:
        print(f"Error loading existing records from MongoDB: {e}")
        return {}

def load_existing_idle_time(table_name):
    """載入當天已存在的idle time記錄，並獲取最大累計時間"""
    try:
        db = get_database()
        today = datetime.now().strftime('%Y-%m-%d')
        
        # Query MongoDB for today's records
        pipeline = [
            {
                '$match': {
                    'date': today,
                    'idle_time': {'$ne': '00:00:00'}
                }
            },
            {
                '$group': {
                    '_id': None,
                    'max_idle': {'$max': '$idle_time'},
                    'all_idle_times': {'$push': '$idle_time'}
                }
            }
        ]
        
        result = list(db.activities.aggregate(pipeline))
        
        if not result:
            return {"max_idle": "00:00:00", "total_idle": "00:00:00"}
            
        max_idle = result[0]['max_idle']
        idle_times = result[0]['all_idle_times']
        
        # Calculate total idle time
        total_idle_seconds = 0
        last_idle_seconds = 0
        
        for idle_time in idle_times:
            if idle_time and idle_time != "00:00:00":
                h, m, s = map(int, idle_time.split(':'))
                current_idle_seconds = h * 3600 + m * 60 + s
                
                if current_idle_seconds > last_idle_seconds:
                    increment = current_idle_seconds - last_idle_seconds
                    total_idle_seconds += increment
                    last_idle_seconds = current_idle_seconds
        
        total_idle = str(timedelta(seconds=int(total_idle_seconds)))
        
        return {
            "max_idle": max_idle,
            "total_idle": total_idle
        }
        
    except Exception as e:
        print(f"Error loading idle times from MongoDB: {e}")
        return {"max_idle": "00:00:00", "total_idle": "00:00:00"}
    
# def log_user_duration(user_name, app_name, start_time, end_time, idle_time):
#     """
#     Log user application duration data to MongoDB.
    
#     Args:
#         user_name (str): Name of the user
#         app_name (str): Name of the application
#         start_time (datetime): Start time of app usage 
#         end_time (datetime): End time of app usage
#         idle_time (str): Idle time in HH:MM:SS format
#     """
#     try:
#         db = get_database()
        
#         # Calculate duration
#         duration = end_time - start_time
#         duration_str = str(duration).split('.')[0]  # Convert to HH:MM:SS format
        
#         # Create duration document
#         duration_doc = {
#             'user_name': user_name,
#             'app_name': app_name,
#             'start_time': start_time.strftime("%Y-%m-%d %H:%M:%S"),
#             'end_time': end_time.strftime("%Y-%m-%d %H:%M:%S"),
#             'duration': duration_str,
#             'app_idle': idle_time,
#             'date': datetime.now().strftime('%Y-%m-%d'),
#             'created_at': datetime.now()
#         }
        
#         # Insert into MongoDB user_duration collection
#         result = db.user_duration.insert_one(duration_doc)
        
#         # Create index if not exists
#         db.user_duration.create_index([
#             ('user_name', 1),
#             ('app_name', 1),
#             ('date', 1)
#         ])
        
#         return result.inserted_id
        
#     except Exception as e:
#         print(f"Error logging user duration: {e}")
#         return None    

def is_system_locked():
    """檢查系統是否處於鎖定狀態"""
    try:
        # 嘗試獲取前景窗口
        hwnd = win32gui.GetForegroundWindow()
        if hwnd == 0:
            return True
            
        # 嘗試獲取窗口標題
        title = win32gui.GetWindowText(hwnd)
        # 鎖定畫面通常沒有窗口標題
        if not title:
            return True
            
        return False
    except Exception:
        # 如果出錯，保守地假設系統已鎖定
        return True

# Main function to run the monitoring script
def main():
    
    cleanup_old_records()

    logon_time = get_logon_time()
    active_app = None
    start_time = time.time()
    
    
    # 載入當天既有的記錄
    table_name = f"activity_{datetime.now().strftime('%Y_%m_%d')}"
    app_usage_times = load_existing_app_usage(table_name)
    
    # 每天重新開始計算
    current_max_idle = "00:00:00"
    total_idle_time = "00:00:00"
    
    # 追蹤累計的idle time
    idle_times = load_existing_idle_time(table_name)
    
    boot_time_str = get_boot_time()
    boot_time = datetime.strptime(boot_time_str, "%Y-%m-%d %H:%M:%S")

    

    while True:
        workstation = get_workstation_name()
        user = get_user_name()
        current_time = get_current_time()
        
        # 使用前一次的最大idle time來計算新的idle time
        idle_time = get_idle_time(user, current_max_idle)
        
        active_time_seconds = time.time() - psutil.boot_time()
        active_time = str(timedelta(seconds=int(active_time_seconds)))
        current_app_name, current_app_title, current_app_path, app_start_time = get_active_application_info()

        # 在main函數中增加以下處理
        if is_system_locked():
            # 如果系統已鎖定，使用特殊處理邏輯
            # 例如減少日誌記錄頻率或標記此時間為系統鎖定
            current_app_name = "System_Locked"
            current_app_title = "Windows鎖定畫面"
            current_app_path = "系統"

        # Calculate system working time
        system_working_time = current_time - boot_time
        system_working_time_str = str(system_working_time).split('.')[0]

        # Check if the active application has changed
        if active_app != current_app_name:
            # If there was a previous active app, log its usage
            if (active_app and active_app in app_usage_times):
                end_time = time.time()
                app_usage_times[active_app]['total_time'] += end_time - start_time

                # Calculate total usage time for the app
                total_time_seconds = app_usage_times[active_app]['total_time']
                total_time_hms = str(timedelta(seconds=int(total_time_seconds)))

                logon_time_str = logon_time.strftime("%Y-%m-%d %H:%M:%S")
                current_time_str = current_time.strftime("%Y-%m-%d %H:%M:%S")

                log_to_database(workstation, user, logon_time_str, current_time_str, idle_time, active_time, active_app, app_usage_times[active_app]['title'], app_usage_times[active_app]['path'], total_time_hms, boot_time_str, app_start_time, total_time_hms, system_working_time_str)
                # log_to_excel(workstation, user, logon_time_str, current_time_str, idle_time, active_time, active_app, app_usage_times[active_app]['title'], app_usage_times[active_app]['path'], total_time_hms, boot_time_str, app_start_time, total_time_hms, system_working_time_str)

                # Log the duration
                start_time_dt = datetime.fromtimestamp(start_time)
                # log_user_duration(
                #     user_name=user,
                #     app_name=active_app,
                #     start_time=start_time_dt,
                #     end_time=current_time,
                #     idle_time=idle_time
                # )

            # Reset the start time and update the active app
            start_time = time.time()
            logon_time = current_time  # set new logon time
            active_app = current_app_name

            # Initialize usage time for the new application
            if current_app_name not in app_usage_times:
                app_usage_times[current_app_name] = {'total_time': 0, 'title': current_app_title, 'path': current_app_path}

        # If the active app hasn't changed, still update its total time
        else:
            end_time = time.time()
            if current_app_name in app_usage_times:
                app_usage_times[current_app_name]['total_time'] += end_time - start_time
                total_time_seconds = app_usage_times[current_app_name]['total_time']
                total_time_hms = str(timedelta(seconds=int(total_time_seconds)))
                logon_time_str = logon_time.strftime("%Y-%m-%d %H:%M:%S")
                current_time_str = current_time.strftime("%Y-%m-%d %H:%M:%S")
                log_to_database(workstation, user, logon_time_str, current_time_str, idle_time, active_time, current_app_name, app_usage_times[current_app_name]['title'], app_usage_times[current_app_name]['path'], total_time_hms, boot_time_str, app_start_time,total_time_hms, system_working_time_str)
            #     log_to_excel(workstation, user, logon_time_str, current_time_str, idle_time, active_time, current_app_name, app_usage_times[current_app_name]['title'], app_usage_times[current_app_name]['path'], total_time_hms, boot_time_str, app_start_time,total_time_hms, system_working_time_str)
            # start_time = time.time()

        # 更新idle time記錄
        if idle_time > current_max_idle:
            current_max_idle = idle_time
            total_idle_time = idle_time  # 直接使用新的 idle time

        time.sleep(1)  # Sleep for 1 second before logging again

# Modify the main function to include exception handling
if __name__ == "__main__":
    # try:
    #     # Add logging for script start
    #     print(f"Script started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
    #     # Run the main function
        main()
    # except KeyboardInterrupt:
    #     # Allow clean exit with Ctrl+C
    #     print("Script terminated by user.")
    #     sys.exit(0)
    # except Exception as e:
    #     # Log the error
    #     error_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    #     error_message = f"Fatal error at {error_time}: {str(e)}"
    #     error_traceback = traceback.format_exc()
        
    #     print(error_message)
    #     print(error_traceback)
        
    #     # Log error to file
    #     log_dir = os.path.join(os.path.dirname(__file__), 'logs')
    #     os.makedirs(log_dir, exist_ok=True)
    #     error_log_file = os.path.join(log_dir, f"error_log_{datetime.now().strftime('%Y_%m_%d')}.txt")
        
    #     with open(error_log_file, 'a') as f:
    #         f.write(f"{error_message}\n{error_traceback}\n\n")
        
    #     # Restart the script
        # restart_script()