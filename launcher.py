"""
NRL Lookup Tool - Launcher
Mở trình duyệt tự động và chạy server Flask
"""
import sys
import os
import webbrowser
import threading
import time
import socket

# Đảm bảo có thể import từ thư mục hiện tại
if getattr(sys, 'frozen', False):
    # Chạy từ exe
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Chạy từ script
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

os.chdir(BASE_DIR)

# Set environment variables
os.environ['EXCEL_FILE'] = os.path.join(BASE_DIR, 'nrl.xlsx')

def find_free_port():
    """Tìm port trống"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        s.listen(1)
        port = s.getsockname()[1]
    return port

def open_browser(port):
    """Mở trình duyệt sau 1.5 giây"""
    time.sleep(1.5)
    webbrowser.open(f'http://127.0.0.1:{port}')

def main():
    # Import Flask app
    from app import app
    
    port = 5000
    
    # Thử tìm port trống nếu 5000 đã dùng
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('127.0.0.1', port))
    except OSError:
        port = find_free_port()
    
    print("="*50)
    print("   NRL LOOKUP TOOL - Tra cứu điểm rèn luyện")
    print("="*50)
    print(f"\n🌐 Server đang chạy tại: http://127.0.0.1:{port}")
    print("📂 File Excel: nrl.xlsx (đặt cùng thư mục)")
    print("\n⚠️  KHÔNG ĐÓNG CỬA SỔ NÀY khi đang sử dụng!")
    print("    Nhấn Ctrl+C để tắt server\n")
    print("="*50)
    
    # Mở browser trong thread riêng
    browser_thread = threading.Thread(target=open_browser, args=(port,))
    browser_thread.daemon = True
    browser_thread.start()
    
    # Chạy Flask server (production mode, không debug)
    from werkzeug.serving import run_simple
    run_simple('127.0.0.1', port, app, use_reloader=False, use_debugger=False)

if __name__ == '__main__':
    main()
