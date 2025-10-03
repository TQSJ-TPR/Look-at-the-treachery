import time
import json
import sys
import platform
import os
import psutil
import ctypes
import ctypes.wintypes as wt
import re
import subprocess
import threading
from flask import Flask, Response, send_from_directory, jsonify, request

# 全局变量用于跟踪当前窗口和激活时间
_current_window_info = {
    'last_window_title': None,
    'window_activation_time': None
}

# Windows API 初始化
user32   = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
psapi    = ctypes.windll.psapi

# 音乐检测相关常量
MUSIC_EXES = {
    'cloudmusic.exe', 'spotify.exe', 'qqmusic.exe', 'kwmusic.exe',
    'kugou.exe', 'coriander_player.exe', 'itunes.exe', 'applemusic.exe'
}

VIDEO_KEYS = {'youtube', 'bilibili', 'video', '电影', '剧集', '预告', 'trailer', 'mv'}
NOISE_TITLES = {'default ime', 'msctfime ui', 'desktop lyrics', '桌面歌词'}

app = Flask(__name__, static_folder='static', static_url_path='')

# ===== 音乐检测功能 =====

def _get_window_title(hwnd):
    length = user32.GetWindowTextLengthW(hwnd)
    if length == 0:
        return ''
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value

def _enum_process_windows(pid):
    hwnds = []
    def _cb(hwnd, _):
        found_pid = wt.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(found_pid))
        if found_pid.value == pid:
            hwnds.append(hwnd)
        return True
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wt.HWND, wt.LPARAM)
    user32.EnumWindows(WNDENUMPROC(_cb), 0)
    return hwnds

def _get_main_orch_window(pid):
    """网易云主窗口类名：OrpheusBrowserHost"""
    hwnds = []
    def _cb(hwnd, _):
        found_pid = wt.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(found_pid))
        if found_pid.value != pid:
            return True
        buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, buf, 256)
        if buf.value == 'OrpheusBrowserHost':
            hwnds.append(hwnd)
        return True
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wt.HWND, wt.LPARAM)
    user32.EnumWindows(WNDENUMPROC(_cb), 0)
    return hwnds[0] if hwnds else None

def get_song_info():
    """获取当前播放歌曲信息"""
    for proc in psutil.process_iter(['pid', 'name']):
        exe = proc.info['name'].lower()
        if exe not in MUSIC_EXES:
            continue
        pid = proc.info['pid']

        # 1) 先枚举所有顶层窗口
        for hwnd in _enum_process_windows(pid):
            raw = _get_window_title(hwnd)
            if not raw or len(raw) < 4:
                continue
            low = raw.lower()
            if any(n in low for n in NOISE_TITLES):
                continue
            if any(k in low for k in VIDEO_KEYS):
                continue

            clean = re.sub(
                r'(\s*[-–—]\s*)?(网易云音乐|CloudMusic|QQ音乐|Spotify|酷狗|酷我|Coriander Player)\s*$',
                '', raw, flags=re.I
            ).strip()
            if ' - ' in clean:
                left, right = clean.rsplit(' - ', 1)
                title, artist = (left, right) if len(left) >= len(right) else (right, left)
            else:
                title, artist = clean, ''
            if title:
                return {
                    'status': 'playing',
                    'title': title,
                    'artist': artist,
                    'album': '',
                    'progress_str': '',
                    'remaining': '',
                    'source': exe,
                    'thumbnail_available': False,
                }

        # 2) 兜底：直接抓网易云主窗口标题
        if exe == 'cloudmusic.exe':
            hwnd_main = _get_main_orch_window(pid)
            if hwnd_main:
                title = _get_window_title(hwnd_main)
                if title and len(title) > 4 and title.lower() not in NOISE_TITLES:
                    clean = re.sub(
                        r'(\s*[-–—]\s*)?网易云音乐\s*$', '', title, flags=re.I
                    ).strip()
                    if ' - ' in clean:
                        left, right = clean.rsplit(' - ', 1)
                        title, artist = (left, right) if len(left) >= len(right) else (right, left)
                    else:
                        title, artist = clean, ''
                    if title:
                        return {
                            'status': 'playing',
                            'title': title,
                            'artist': artist,
                            'album': '',
                            'progress_str': '',
                            'remaining': '',
                            'source': 'cloudmusic.exe',
                            'thumbnail_available': False,
                        }

    return {
        'status': 'idle',
        'title': None,
        'artist': None,
        'album': '',
        'progress_str': '',
        'remaining': '',
        'source': None,
        'thumbnail_available': False,
    }

# ===== 窗口监控功能 =====

def get_active_window():
    """获取当前活动窗口信息"""
    if platform.system() != 'Windows':
        return {'title': 'N/A (non-Windows)', 'boot_time': None}
    try:
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return {'title': '', 'boot_time': None}

        length = 512
        buffer = ctypes.create_unicode_buffer(length)
        if user32.GetWindowTextW(hwnd, buffer, length) == 0:
            raw_title = ''
        else:
            raw_title = buffer.value

        title = os.path.basename(raw_title) if isinstance(raw_title, str) else ''
        if raw_title == 'Program Manager':
            title = '摸鱼～～～'
        else:
            if ' - ' in title:
                title = title.split(' - ', 1)[0]
            title = os.path.splitext(title)[0]
            
        # 如果窗口标题为空，显示"摸鱼～～～"
        if not title or title.strip() == '':
            title = '摸鱼～～～'

        pid = wt.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value == 0:
            proc_name = ''
        else:
            PROCESS_QUERY_INFORMATION = 0x0400
            PROCESS_VM_READ = 0x0010
            handle = kernel32.OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, pid.value)
            if handle:
                exe = ctypes.create_unicode_buffer(260)
                if psapi.GetModuleBaseNameW(handle, None, exe, 260):
                    proc_name = exe.value
                else:
                    proc_name = ''
                kernel32.CloseHandle(handle)
            else:
                proc_name = ''

        if isinstance(proc_name, str) and proc_name:
            proc_name = os.path.basename(proc_name)
        
        if raw_title == 'Program Manager' or not proc_name or proc_name.strip() == '':
            proc_name = '挂机ing...'

        boot_time = None
        try:
            boot_epoch = psutil.boot_time()
            if isinstance(boot_epoch, (int, float)):
                boot_time = float(time.time() - float(boot_epoch))
        except Exception:
            boot_time = None

        if boot_time is None:
            try:
                tick = kernel32.GetTickCount64()
                boot_time = tick / 1000.0
            except Exception:
                boot_time = None

        # 计算当前窗口的已打开时间
        active_duration = None
        if _current_window_info['last_window_title'] != title:
            # 窗口切换，重置激活时间
            _current_window_info['last_window_title'] = title
            _current_window_info['window_activation_time'] = time.time()
        elif _current_window_info['window_activation_time'] is not None:
            # 计算从激活到现在的时间（秒）
            active_duration = time.time() - _current_window_info['window_activation_time']

        return {
            'title': title, 
            'boot_time': boot_time,
            'active_duration': active_duration
        }
    except Exception:
        return {'title': '', 'boot_time': None, 'active_duration': None}

# ===== 手机应用检测功能 (MacroDroid LAN) =====

# 存储从MacroDroid接收到的手机应用数据
_mobile_apps_data = {
    'apps': [],
    'last_update': None,
    'status': 'waiting',
    'message': '等待MacroDroid数据...',
    'current_app': None
}

def get_mobile_apps():
    """获取手机当前打开的应用 (通过MacroDroid)"""
    global _mobile_apps_data
    
    # 如果有接收到的应用数据，始终显示最后一个应用
    if _mobile_apps_data['current_app']:
        return {
            'status': 'success',
            'message': f'当前应用: {_mobile_apps_data["current_app"]["name"]}',
            'apps': [_mobile_apps_data['current_app']],
            'last_update': _mobile_apps_data['last_update']
        }
    
    # 初始状态
    return {
        'status': 'waiting',
        'message': '等待MacroDroid数据...',
        'apps': [],
        'last_update': None
    }

def update_mobile_apps_from_macrodroid(data):
    """从MacroDroid更新手机应用数据 - 始终保存最新应用"""
    global _mobile_apps_data
    
    try:
        current_app = None
        
        # 应用名称过滤和替换规则
        def process_app_name(name):
            if not name:
                return None
            
            # 过滤讯飞输入法
            if '讯飞输入法' in str(name) or '讯飞' in str(name):
                return None
            
            # 替换熄屏显示
            if str(name) == '熄屏显示':
                return '手机熄屏ing'
            
            # 替换华为桌面为手机熄屏中
            if str(name) == '华为桌面':
                return '手机熄屏ing'

            return str(name)
        
        # 处理MacroDroid发送的数据格式，提取最新应用
        if isinstance(data, dict):
            if 'app_name' in data:
                processed_name = process_app_name(data.get('app_name'))
                if processed_name:
                    current_app = {
                        'name': processed_name,
                        'package': data.get('package_name', 'unknown'),
                        'timestamp': time.time()
                    }
            elif 'name' in data:
                processed_name = process_app_name(data.get('name'))
                if processed_name:
                    current_app = {
                        'name': processed_name,
                        'package': data.get('package', 'unknown'),
                        'timestamp': time.time()
                    }
            elif 'apps' in data and data['apps']:
                # 如果有应用列表，取第一个作为当前应用
                app = data['apps'][0]
                if isinstance(app, dict):
                    processed_name = process_app_name(app.get('name', str(app)))
                    if processed_name:
                        current_app = {
                            'name': processed_name,
                            'package': app.get('package', 'unknown'),
                            'timestamp': time.time()
                        }
                else:
                    processed_name = process_app_name(str(app))
                    if processed_name:
                        current_app = {
                            'name': processed_name,
                            'package': str(app).lower().replace(' ', '_'),
                            'timestamp': time.time()
                        }
        elif isinstance(data, str) and data.strip():
            # 简单字符串格式
            processed_name = process_app_name(data)
            if processed_name:
                current_app = {
                    'name': processed_name,
                    'package': data.lower().replace(' ', '_'),
                    'timestamp': time.time()
                }
        elif isinstance(data, list) and data:
            # 应用列表，取第一个
            item = data[0]
            if isinstance(item, dict):
                processed_name = process_app_name(item.get('name', str(item)))
                if processed_name:
                    current_app = {
                        'name': processed_name,
                        'package': item.get('package', 'unknown'),
                        'timestamp': time.time()
                    }
            else:
                processed_name = process_app_name(str(item))
                if processed_name:
                    current_app = {
                        'name': processed_name,
                        'package': str(item).lower().replace(' ', '_'),
                        'timestamp': time.time()
                    }
        
        if current_app:
            _mobile_apps_data = {
                'current_app': current_app,
                'last_update': time.time(),
                'status': 'success',
                'message': f'当前应用: {current_app["name"]}'
            }
            return True
        
        return False
        
    except Exception as e:
        # 保持之前的应用数据，不重置
        _mobile_apps_data['status'] = 'error'
        _mobile_apps_data['message'] = f'数据格式错误: {str(e)}'
        return False

# ===== Flask 路由 =====

@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/stream')
def stream():
    """实时窗口信息流"""
    def event_stream():
        while True:
            data = get_active_window()
            if not isinstance(data, dict):
                data = {'title': '', 'boot_time': None}
            else:
                data = dict(data)
                data.setdefault('title', '')
                data.setdefault('boot_time', None)
            yield f"data: {json.dumps(data)}\n\n"
            time.sleep(1)
    return Response(event_stream(), mimetype="text/event-stream")

@app.route('/song')
def get_song():
    """独立获取当前播放歌曲的接口"""
    return jsonify(get_song_info())

@app.route('/song/stream')
def song_stream():
    """实时歌曲信息流"""
    def song_event_stream():
        while True:
            song_data = get_song_info()
            yield f"data: {json.dumps(song_data)}\n\n"
            time.sleep(2)
    return Response(song_event_stream(), mimetype="text/event-stream")

@app.route('/mobile')
def get_mobile():
    """独立获取手机应用信息的接口"""
    return jsonify(get_mobile_apps())

@app.route('/mobile/stream')
def mobile_stream():
    """实时手机应用信息流"""
    def mobile_event_stream():
        while True:
            mobile_data = get_mobile_apps()
            yield f"data: {json.dumps(mobile_data)}\n\n"
            time.sleep(1)
    return Response(mobile_event_stream(), mimetype="text/event-stream")

@app.route('/macrodroid', methods=['POST'])
def receive_macrodroid():
    """接收MacroDroid发送的手机应用信息"""
    try:
        data = request.get_json()
        if not data:
            # 尝试从表单数据获取
            data = {
                'app_name': request.form.get('app_name', request.form.get('name')),
                'package_name': request.form.get('package_name', request.form.get('package')),
                'timestamp': request.form.get('timestamp', time.time())
            }
        
        success = update_mobile_apps_from_macrodroid(data)
        
        return jsonify({
            'status': 'success' if success else 'error',
            'message': '数据已接收' if success else '数据格式错误',
            'received_data': data
        })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 400

@app.route('/macrodroid/test')
def test_macrodroid():
    """测试MacroDroid接口"""
    return jsonify({
        'message': 'MacroDroid接口测试',
        'example_request': {
            'method': 'POST',
            'url': '/macrodroid',
            'content_type': 'application/json',
            'data': {
                'app_name': '微信',
                'package_name': 'com.tencent.mm'
            }
        },
        'current_data': get_mobile_apps()
    })

@app.route('/macrodroid/simple', methods=['GET', 'POST'])
def simple_macrodroid():
    """简化的MacroDroid接口，支持GET和POST"""
    try:
        if request.method == 'GET':
            app_name = request.args.get('app', request.args.get('name', '未知应用'))
        else:
            if request.is_json:
                data = request.get_json()
                app_name = data.get('app', data.get('name', '未知应用'))
            else:
                app_name = request.form.get('app', request.form.get('name', '未知应用'))
        
        update_mobile_apps_from_macrodroid({
            'name': app_name,
            'package': app_name.lower().replace(' ', '_')
        })
        
        return jsonify({
            'status': 'success',
            'message': f'已更新应用: {app_name}'
        })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 400

if __name__ == '__main__':
    import webbrowser
    import threading
    
    def _open_browser():
        time.sleep(1.5)
        try:
            webbrowser.open('http://localhost:5000/')
        except Exception:
            pass
    
    t = threading.Thread(target=_open_browser)
    t.daemon = True
    t.start()
    
    print("启动中...")
    print("访问: http://localhost:5000")
    app.run(debug=False, threaded=True, host='0.0.0.0', port=5000)