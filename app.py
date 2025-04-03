import os
import sys
import time

# Add single instance check at the very beginning
# try:
#     from single_instance import SingleInstance
#     instance = SingleInstance("ActivityTrackerAPI")
# except ImportError:
#     # If import fails, continue running (development environment fallback)
#     pass

from flask import Flask, render_template, jsonify, request, session
from flask_cors import CORS
from functools import wraps
import hashlib
from datetime import timedelta
import pandas as pd
import os
from datetime import datetime
import logging
from contextlib import contextmanager
from database.mongo_config import get_database
from bson import ObjectId
from logger_config import setup_logger
from config import CONFIG

# 設置記錄器
logger = setup_logger('ActivityTrackerAPI')

# 記錄啟動信息
logger.info('API Service starting...')
logger.info(f'User: {os.environ.get("USERNAME")}')
logger.info(f'Computer name: {os.environ.get("COMPUTERNAME")}')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# 添加一個函數用於獲取應用程序根目錄
def get_app_root():
    """獲取應用程序根目錄，處理打包和非打包情況"""
    if getattr(sys, 'frozen', False):
        # 打包後的可執行文件路徑
        return os.path.dirname(sys.executable)
    else:
        # 開發環境
        return os.path.dirname(os.path.abspath(__file__))

# 修改 ensure_data_directory 函數
def ensure_data_directory():
    data_dir = os.path.join(get_app_root(), 'data')
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)
    return data_dir

def format_timedelta(td):
    if pd.isnull(td):
        return '00:00:00'
    seconds = int(td.total_seconds())
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
        # if 'user name' not in session:
            return jsonify({"error": "Authentication required"}), 401
        return f(*args, **kwargs)
    return decorated_function

app = Flask(__name__)
CORS(app, 
     supports_credentials=CONFIG['CORS']['SUPPORTS_CREDENTIALS'],
     origins=CONFIG['CORS']['ORIGINS'],
     allow_headers=CONFIG['CORS']['ALLOW_HEADERS'],
     methods=CONFIG['CORS']['METHODS'],
     expose_headers=CONFIG['CORS']['EXPOSE_HEADERS'])

app.secret_key = CONFIG['SECRET_KEY']
app.permanent_session_lifetime = timedelta(days=CONFIG['SESSION_LIFETIME_DAYS'])

# Add this after the app initialization but before any routes
def init_app():
    """初始化應用，連接 MongoDB 並執行清理工作"""
    logger.info("開始應用程序初始化...")
    try:
        # 嘗試連接 MongoDB 並在打包環境中提供更明確的錯誤消息
        try:
            db = get_database()
            # 簡單測試連接
            db.command('ping')
            logger.info("MongoDB 連接成功初始化")
            if getattr(sys, 'frozen', False):
                print("MongoDB 連接正常")
        except Exception as db_err:
            error_msg = f"MongoDB 連接錯誤: {str(db_err)}"
            logger.error(error_msg)
            if getattr(sys, 'frozen', False):
                print(f"警告: {error_msg}")
                print("請確保 MongoDB 服務正在運行")
                print("應用程序將嘗試繼續運行，但可能會出現數據問題")
                time.sleep(2)  # 給用戶時間閱讀警告
        
        # 執行 Excel 文件清理
        try:
            deleted_count, deleted_files = cleanup_excel_files()
            if deleted_count > 0:
                logger.info(f"啟動清理: 移除了 {deleted_count} 個舊 Excel 文件")
                logger.info(f"已刪除文件: {', '.join(deleted_files)}")
            else:
                logger.info("啟動時無需清理 Excel 文件")
        except Exception as cleanup_err:
            logger.error(f"Excel 清理錯誤: {str(cleanup_err)}")
    except Exception as e:
        logger.error(f"啟動過程中出錯: {str(e)}")

@app.route('/api/register', methods=['POST'])
def register():
    try:
        data = request.get_json()
        username = data.get('username')
        password = data.get('password')
        
        if not username or not password:
            return jsonify({"error": "Username and password required"}), 400
            
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        
        # Use MongoDB instead of SQLite
        db = get_database()
        try:
            db.users.insert_one({
                'username': username,
                'password_hash': password_hash,
                'created_at': datetime.now()
            })
            return jsonify({"message": "User registered successfully"}), 201
        except Exception as e:
            if 'duplicate key error' in str(e):
                return jsonify({"error": "Username already exists"}), 409
            raise
                
    except Exception as e:
        logger.error(f"Registration error: {str(e)}")
        return jsonify({"error": "Registration failed"}), 500

# Initialize MongoDB connection
db = get_database()

@app.route('/api/login', methods=['POST'])
def login():
    try:
        data = request.get_json()
        username = data.get('username')
        password = data.get('password')
        
        if not username or not password:
            return jsonify({"error": "Username and password required"}), 400
            
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        
        # Find user in MongoDB
        user = db.users.find_one({
            'username': username,
            'password_hash': password_hash
        })
        
        if user:
            session.permanent = True
            session['user_id'] = str(user['_id'])
            session['username'] = user['username']
            return jsonify({
                "message": "Login successful",
                "user": {"id": str(user['_id']), "username": user['username']}
            }), 200
            
        return jsonify({"error": "Invalid credentials"}), 401
        
    except Exception as e:
        logger.error(f"Login error: {str(e)}")
        return jsonify({"error": "Login failed"}), 500

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({"message": "Logged out successfully"}), 200

@app.route('/api/activities')
# @login_required
def get_data():
    try:
        # 從請求參數獲取日期範圍，默認為當天
        start_date = request.args.get('start_date')
        end_date = request.args.get('end_date')
        
        # 如果沒有提供日期參數，默認只返回當天記錄
        if not start_date:
            start_date = datetime.now().strftime('%Y-%m-%d')
        if not end_date:
            end_date = datetime.now().strftime('%Y-%m-%d')
            
        logger.info(f"Fetching activities from {start_date} to {end_date}")
        
        # 查詢 MongoDB 獲取指定日期範圍的活動記錄
        activities = list(db.activities.find({
            'date': {
                '$gte': start_date,
                '$lte': end_date
            }
        }))
        
        # 轉換 ObjectId 為字串以便 JSON 序列化
        for activity in activities:
            activity['_id'] = str(activity['_id'])
        
        # 對記錄進行預排序，按照創建時間降序，以便後面處理時最新的記錄會覆蓋舊的
        activities.sort(key=lambda x: (
            x['date'],
            x.get('user_name', ''),
            x.get('app_name', ''),
            x.get('app_start_time', ''),
            x.get('created_at', '')
        ), reverse=True)  # 降序排序，最新的記錄在前
        
        # 合併相同應用程序的記錄，保留具有最大 logoff_time 的記錄
        merged_activities = {}
        for activity in activities:
            # 創建唯一鍵，用於識別相同的活動記錄
            key = (
                activity.get('date', ''),
                activity.get('user_name', ''),
                activity['workstation_name'],  # 改為直接訪問 workstation_name
                activity.get('app_name', ''),
                activity.get('logon_time', '') or activity.get('app_start_time', '')  # 優先使用 logon_time
            )
            
            # 如果鍵不存在或當前記錄的 logoff_time 更大，則更新
            current_logoff = activity.get('logoff_time', '')
            if key not in merged_activities or (current_logoff and current_logoff > merged_activities[key].get('logoff_time', '')):
                merged_activities[key] = activity
        
        # 將字典轉換回列表
        unique_activities = list(merged_activities.values())
        
        # 重新計算每個活動的 total_time
        for activity in unique_activities:
            logon_time = activity.get('logon_time', '') or activity.get('app_start_time', '')
            logoff_time = activity.get('logoff_time', '')
            
            if logon_time and logoff_time:
                try:
                    # 嘗試提取完整的日期時間格式
                    if ' ' in logon_time and ' ' in logoff_time:
                        # 完整格式: '2025-02-28 17:24:56'
                        logon_dt = datetime.strptime(logon_time, '%Y-%m-%d %H:%M:%S')
                        logoff_dt = datetime.strptime(logoff_time, '%Y-%m-%d %H:%M:%S')
                    else:
                        # 只有時間格式: 'HH:MM:SS'
                        logon_dt = datetime.strptime(logon_time, '%H:%M:%S')
                        logoff_dt = datetime.strptime(logoff_time, '%H:%M:%S')
                        
                        # 處理跨日的情況（如果結束時間小於開始時間，視為下一天）
                        if logoff_dt < logon_dt:
                            logoff_dt = logoff_dt + timedelta(days=1)
                    
                    # 計算時間差
                    time_diff = logoff_dt - logon_dt
                    
                    # 格式化為 HH:MM:SS
                    hours, remainder = divmod(time_diff.seconds, 3600)
                    minutes, seconds = divmod(remainder, 60)
                    activity['total_time'] = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                    
                    # 添加調試日誌
                    logger.debug(f"計算 total_time: {activity['total_time']} (logon: {logon_time}, logoff: {logoff_time})")
                    
                except ValueError as e:
                    logger.warning(f"無法計算 total_time: {str(e)} - logon: {logon_time}, logoff: {logoff_time}")
                    activity['total_time'] = activity.get('total_time', '00:00:00')
            else:
                activity['total_time'] = activity.get('total_time', '00:00:00')
        
        # 計算使用時間摘要 - 按用戶、日期和應用程式分組
        usage_time_summary = []
        
        # 建立按用戶、日期、應用程式分組的時間統計
        usage_time_dict = {}
        for activity in unique_activities:
            key = (
                activity.get('date', ''),
                activity.get('user_name', ''),
                activity.get('app_name', '')
            )
            
            try:
                # 解析時間字串並轉換為秒數
                time_parts = activity.get('total_time', '00:00:00').split(':')
                total_seconds = int(time_parts[0]) * 3600 + int(time_parts[1]) * 60 + int(time_parts[2])
                
                if key not in usage_time_dict:
                    usage_time_dict[key] = {
                        'date': activity.get('date', ''),
                        'user_name': activity.get('user_name', ''),
                        'app_name': activity.get('app_name', ''),
                        'total_seconds': total_seconds,
                        'session_count': 1
                    }
                else:
                    usage_time_dict[key]['total_seconds'] += total_seconds
                    usage_time_dict[key]['session_count'] += 1
            except (ValueError, IndexError) as e:
                logger.warning(f"解析時間出錯: {str(e)} - {activity.get('total_time')}")
        
        # 轉換為列表並格式化總時間
        for stats in usage_time_dict.values():
            hours, remainder = divmod(stats['total_seconds'], 3600)
            minutes, seconds = divmod(remainder, 60)
            stats['total_time'] = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            del stats['total_seconds']  # 移除不需要的欄位
            usage_time_summary.append(stats)
        
        # 按日期和使用時間排序
        usage_time_summary.sort(key=lambda x: (x['date'], x['user_name'], x['total_time']), reverse=True)
        
        logger.info(f"生成了 {len(usage_time_summary)} 個使用時間摘要記錄")
        
        # 返回結果時使用實際查詢的日期範圍
        return jsonify({
            'total_records': len(unique_activities),
            'activities': unique_activities,
            'usagetime': usage_time_summary,
            'date_range': {
                'from': start_date,
                'to': end_date
            }
        })
        
    except Exception as e:
        logger.error(f"Error fetching activities: {str(e)}")
        return jsonify({
            'error': str(e),
            'type': type(e).__name__
        }), 500

@app.route('/api/usage')
# login_required  
def get_app_usage_stats():
    try:
        all_stats = []
        three_days_ago = (datetime.now() - timedelta(days=3)).strftime('%Y-%m-%d')
        
        try:
            db = get_database()
            pipeline = [
                {
                    '$match': {
                        'date': {'$gte': three_days_ago}
                    }
                },
                {
                    '$group': {
                        '_id': {
                            'user_name': '$user_name',
                            'app_name': '$app_name',
                            'date': '$date'
                        },
                        'usage_count': {'$sum': 1},
                        # 計算時間字串轉換為秒數
                        'total_seconds': {
                            '$sum': {
                                '$function': {
                                    'body': '''function(time) {
                                        if (!time) return 0;
                                        const parts = time.split(':');
                                        if (parts.length !== 3) return 0;
                                        
                                        // 安全轉換為數字
                                        const hours = parseInt(parts[0]) || 0;
                                        const minutes = parseInt(parts[1]) || 0;
                                        const seconds = parseInt(parts[2]) || 0;
                                        
                                        // 驗證時間合理性（不超過24小時）
                                        if (hours > 24 || minutes > 59 || seconds > 59) {
                                            return 0;
                                        }
                                        
                                        return (hours * 3600) + (minutes * 60) + seconds;
                                    }''',
                                    'args': ['$total_time'],
                                    'lang': 'js'
                                }
                            }
                        },
                        'max_seconds': {
                            '$max': {
                                '$function': {
                                    'body': '''function(time) {
                                        if (!time) return 0;
                                        const parts = time.split(':');
                                        if (parts.length !== 3) return 0;
                                        
                                        // 安全轉換為數字
                                        const hours = parseInt(parts[0]) || 0;
                                        const minutes = parseInt(parts[1]) || 0;
                                        const seconds = parseInt(parts[2]) || 0;
                                        
                                        // 驗證時間合理性（不超過24小時）
                                        if (hours > 24 || minutes > 59 || seconds > 59) {
                                            return 0;
                                        }
                                        
                                        return (hours * 3600) + (minutes * 60) + seconds;
                                    }''',
                                    'args': ['$total_time'],
                                    'lang': 'js'
                                }
                            }
                        },
                        'max_time': {'$max': '$total_time'}  # 保留最長單次使用時間
                    }
                },
                {
                    '$project': {
                        '_id': 0,
                        'user_name': '$_id.user_name',
                        'app_name': '$_id.app_name',
                        'date': '$_id.date',
                        'usage_count': 1,
                        'max_time': 1,
                        'total_time': {
                            '$function': {
                                'body': '''function(seconds) {
                                    // 限制合理範圍 - 一天最多24小時
                                    seconds = Math.min(seconds, 86400);
                                    
                                    const hours = Math.floor(seconds / 3600);
                                    const minutes = Math.floor((seconds % 3600) / 60);
                                    const secs = seconds % 60;
                                    return (hours < 10 ? '0' : '') + hours + ':' +
                                           (minutes < 10 ? '0' : '') + minutes + ':' +
                                           (secs < 10 ? '0' : '') + secs;
                                }''',
                                'args': ['$total_seconds'],
                                'lang': 'js'
                            }
                        }
                    }
                },
                {'$sort': {'date': -1, 'total_seconds': -1}}  # 改為按總使用時間排序
            ]
            
            all_stats = list(db.activities.aggregate(pipeline))
            logger.info(f"Retrieved {len(all_stats)} app usage statistics")
            
        except Exception as mongo_err:
            logger.error(f"MongoDB stats error: {str(mongo_err)}")
            return jsonify({'error': f"Database error: {str(mongo_err)}"}), 500
            
        return jsonify({
            'total_records': len(all_stats),
            'stats': all_stats,
            'date_range': {
                'from': three_days_ago,
                'to': datetime.now().strftime('%Y-%m-%d')
            }
        })
        
    except Exception as e:
        logger.error(f"Error getting usage stats: {str(e)}")
        return jsonify({'error': str(e)}), 500

from datetime import datetime, timedelta
@app.route('/api/afk')
def get_afk_stats():
    try:
        db = get_database()
        # 取得查詢參數，預設為最近三天
        days = request.args.get('days', default=7, type=int)
        username = request.args.get('username')
        
        # 計算過濾日期
        three_days_ago = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
        
        # 構建查詢條件
        query = {'date': {'$gte': three_days_ago}}
        if username:
            query['username'] = username
        
        # 使用 MongoDB 排序功能
        afk_records = list(db.afk.find(query).sort([
            ("username", 1),
            ("date", 1),
            ("start_time", 1)
        ]))
        
        if not afk_records:
            return jsonify({
                "error": "No data found",
                "details": "No AFK statistics available"
            }), 404

        # 格式化記錄
        formatted_stats = []
        for record in afk_records:
            formatted_stats.append({
                'user_name': record['username'],
                'Status': record['type'],
                'date': record['date'],
                'duration': record['duration'],
                'window': record.get('window', 'Unknown'),
                'start_time': record.get('start_time', ''),
                'end_time': record.get('end_time', '')
            })
        
        # 合併連續相同視窗的記錄
        merged_stats = []
        if formatted_stats:
            current = formatted_stats[0]
            
            for i in range(1, len(formatted_stats)):
                next_record = formatted_stats[i]
                
                # 檢查是否可以合併（相同使用者、狀態、視窗，且時間連續）
                can_merge = (
                    current['user_name'] == next_record['user_name'] and
                    current['Status'] == next_record['Status'] and
                    current['date'] == next_record['date'] and
                    current['window'] == next_record['window'] and
                    current['end_time'] == next_record['start_time']
                )
                
                if can_merge:
                    # 更新結束時間和持續時間
                    current['end_time'] = next_record['end_time']
                    
                    # 重新計算合併後的持續時間
                    try:
                        start_dt = datetime.strptime(current['start_time'], '%H:%M:%S')
                        end_dt = datetime.strptime(current['end_time'], '%H:%M:%S')
                        
                        # 處理跨日情況
                        if end_dt < start_dt:
                            end_dt += timedelta(days=1)
                            
                        duration_seconds = (end_dt - start_dt).total_seconds()
                        hours, remainder = divmod(int(duration_seconds), 3600)
                        minutes, seconds = divmod(remainder, 60)
                        current['duration'] = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                    except ValueError:
                        logger.warning(f"無法計算合併記錄的持續時間: {current['start_time']} to {current['end_time']}")
                else:
                    # 無法合併，將當前記錄添加到結果中，開始新記錄
                    merged_stats.append(current)
                    current = next_record
            
            # 不要忘記最後一筆記錄
            merged_stats.append(current)
        
        logger.info(f"Retrieved {len(afk_records)} AFK records, consolidated to {len(merged_stats)} records")
        
        return jsonify({
            'total_records': len(merged_stats),
            'afk_stats': merged_stats,
            'date_range': {
                'from': three_days_ago,
                'to': datetime.now().strftime('%Y-%m-%d')
            }
        })

    except Exception as e:
        logger.error(f"Error processing AFK statistics: {str(e)}")
        return jsonify({
            'error': str(e),
            'type': type(e).__name__,
            'details': 'Error processing AFK statistics'
        }), 500

@app.route('/api/afk/summary')
def get_afk_summary():
    try:
        db = get_database()
        three_days_ago = (datetime.now() - timedelta(days=3)).strftime('%Y-%m-%d')
        
        # 獲取用戶名篩選（可選）
        username = request.args.get('username')
        
        # 構建匹配條件 
        match_criteria = {
            'date': {'$gte': three_days_ago},
            'is_heartbeat': {'$ne': True}  # 排除心跳資料
        }
        
        if username:
            match_criteria['username'] = username
            
        # 用聚合管道分析每位使用者每天的 AFK 時間
        pipeline = [
            {'$match': match_criteria},
            {'$group': {
                '_id': {
                    'username': '$username',
                    'date': '$date',
                    'type': '$type'
                },
                'total_records': {'$sum': 1},
                'total_duration_str': {'$sum': {
                    '$cond': [
                        {'$eq': ['$type', 'afk']},
                        {'$function': {
                            'body': '''function(duration) {
                                const parts = duration.split(':');
                                return (parseInt(parts[0]) * 3600) + (parseInt(parts[1]) * 60) + (parseInt(parts[2]));
                            }''',
                            'args': ['$duration'],
                            'lang': 'js'
                        }},
                        0
                    ]
                }}
            }},
            {'$project': {
                '_id': 0,
                'username': '$_id.username',
                'date': '$_id.date',
                'type': '$_id.type',
                'total_records': 1,
                'total_duration_str': {
                    '$function': {
                        'body': '''function(seconds) {
                            const hours = Math.floor(seconds / 3600);
                            const minutes = Math.floor((seconds % 3600) / 60);
                            const secs = seconds % 60;
                            return (hours < 10 ? '0' : '') + hours + ':' +
                                   (minutes < 10 ? '0' : '') + minutes + ':' +
                                   (secs < 10 ? '0' : '') + secs;
                        }''',
                        'args': ['$total_duration_str'],
                        'lang': 'js'
                    }
                }
            }},
            {'$sort': {'date': -1, 'username': 1}}
        ]
        
        summary_data = list(db.afk.aggregate(pipeline))
        
        # 回傳摘要結果
        return jsonify({
            'total_records': len(summary_data),
            'summary': summary_data
        })
        
    except Exception as e:
        logger.error(f"Error generating AFK summary: {str(e)}")
        return jsonify({
            'error': str(e),
            'type': type(e).__name__,
            'details': 'Error generating AFK summary'
        }), 500

@app.route('/api/cleanup', methods=['POST'])
@login_required
def trigger_cleanup():
    """Manually trigger Excel file cleanup"""
    try:
        deleted_count, deleted_files = cleanup_excel_files()
        return jsonify({
            "message": f"Cleanup completed. Deleted {deleted_count} files",
            "deleted_files": deleted_files
        })
    except Exception as e:
        logger.error(f"Cleanup error: {str(e)}")
        return jsonify({"error": "Cleanup failed"}), 500

# Add after the app initialization
@app.teardown_appcontext
def shutdown_session(exception=None):
    """Clean up resources when app context ends"""
    try:
        db = get_database()
        if db:
            db.client.close()
    except Exception as e:
        logger.error(f"Error closing MongoDB connection: {e}")

# 更新主程序部分，完善錯誤處理
if __name__ == "__main__":
    try:
        print("初始化 Activity Tracker API...")
        
        # 確保數據目錄存在
        ensure_data_directory()
        
        # 初始化應用
        init_app()
        
        # 檢查是否以打包方式運行
        if getattr(sys, 'frozen', False):
            # 生產模式 - 使用 waitress
            from waitress import serve
            
            # 輸出清晰的啟動信息
            # print("=" * 50)
            # print(" Activity Tracker API 服務已啟動")
            # print(" 訪問地址: http://127.0.0.1:5000")
            # print("=" * 50)
            
            # 啟動生產服務器
            serve(app, host='0.0.0.0', port=5000)
            # 127.0.0.1
        else:
            debug_mode = os.environ.get('FLASK_DEBUG', 'False').lower() == 'true'
            if debug_mode:
                print("警告: 調試模式已啟用，這會帶來安全風險")
                print("僅用於開發環境，切勿在生產環境中啟用")
            # 開發模式 - 使用 Flask 開發服務器
            app.run(
                host='0.0.0.0',
                port=5000,
                debug=False,  
                use_reloader=False
            )
    except OSError as e:
        # 處理端口被占用等情況
        error_msg = f"啟動服務器時出錯: {str(e)}"
        logger.error(error_msg)
        print(f"錯誤: {error_msg}")
        
        if "Address already in use" in str(e):
            print("端口 5000 已被佔用，可能是另一個 Activity Tracker API 實例正在運行")
            
        # 在打包環境中保持窗口開啟
        if getattr(sys, 'frozen', False):
            input("按 Enter 鍵退出...")
    except Exception as e:
        # 處理其他所有異常
        error_msg = f"應用程序啟動錯誤: {str(e)}"
        logger.error(error_msg)
        print(f"錯誤: {error_msg}")
        print(f"詳細信息: {type(e).__name__}")
        
        # 顯示完整的異常信息
        import traceback
        traceback.print_exc()
        
        # 在打包環境中保持窗口開啟
        if getattr(sys, 'frozen', False):
            input("按 Enter 鍵退出...")