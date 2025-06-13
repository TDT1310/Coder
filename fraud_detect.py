
import threading
import webview
from main_app import app  # Import your Flask app object

def start_flask():
    app.run()

if __name__ == '__main__':
    # Start Flask in a separate thread
    flask_thread = threading.Thread(target=start_flask)
    flask_thread.daemon = True
    flask_thread.start()

    # Open the webview window
    webview.create_window("Financial Fraud Detection", "http://127.0.0.1:5000")
    webview.start()