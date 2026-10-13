"""
रजत पीडीएफ कन्वर्टर (Rajat PDF Converter) - डेस्कटॉप लॉन्चर
यह स्क्रिप्ट बैकग्राउंड में लोकल सर्वर चलाकर एक खूबसूरत डेस्कटॉप विंडो में सॉफ्टवेयर खोलती है।
"""

import os
import sys
import time
import socket
import subprocess
import webbrowser
import threading

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

def find_available_port(start_port=5050):
    port = start_port
    while is_port_in_use(port):
        port += 1
    return port

def launch_desktop_window(url):
    """
    Tries to open in Microsoft Edge or Google Chrome App Mode (frameless web app).
    Falls back to default browser if not found.
    """
    edge_paths = [
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    ]
    
    chrome_paths = [
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe")
    ]
    
    # Check Edge
    for ep in edge_paths:
        if os.path.exists(ep):
            try:
                subprocess.Popen([ep, f"--app={url}", "--window-size=1240,840"])
                return
            except Exception:
                pass
                
    # Check Chrome
    for cp in chrome_paths:
        if os.path.exists(cp):
            try:
                subprocess.Popen([cp, f"--app={url}", "--window-size=1240,840"])
                return
            except Exception:
                pass
                
    # Fallback to default browser
    webbrowser.open(url)

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)
    sys.path.insert(0, base_dir)

    from app import app
    
    port = 5050
    if is_port_in_use(port):
        # Already running or port occupied
        url = f"http://127.0.0.1:{port}"
        print(f"सॉफ्टवेयर पहले से सक्रिय है: {url}")
        launch_desktop_window(url)
        return

    url = f"http://127.0.0.1:{port}"

    # Launch browser after a short delay
    def open_browser():
        time.sleep(1.2)
        print(f"ओपनिंग डेस्कटॉप इंटरफेस: {url}")
        launch_desktop_window(url)

    threading.Thread(target=open_browser, daemon=True).start()

    print("=" * 60)
    print(" 🚀 रजत पीडीएफ कन्वर्टर (Rajat PDF Converter Pro)")
    print(f" विंडो प्रारंभ हो रही है: {url}")
    print(" विंडो बंद करने के लिए Ctrl+C दबाएं।")
    print("=" * 60)

    app.run(host="127.0.0.1", port=port, debug=False)

if __name__ == "__main__":
    main()
