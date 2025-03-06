import streamlit as st
import main
import threading
from datetime import datetime
import queue
import time
import uuid

class ThreadSafeLogger:
    def __init__(self):
        self.log_queue = queue.Queue()
        self._lock = threading.Lock()
        self.is_stopped = False
        self.active_thread = None
        self.log_history = []
        self.processing_complete = threading.Event()
        self.session_id = str(uuid.uuid4())[:8]  # Unique session ID for debugging

    def add_log(self, message):
        """Add a log message to both queue and history"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        with self._lock:
            self.log_history.append(log_entry)
            self.log_queue.put(log_entry)
        print(f"Added log: {log_entry}")  # Debug print

    def get_new_logs(self):
        """Get new logs from queue without removing them from history"""
        new_logs = []
        try:
            while not self.log_queue.empty():
                new_logs.append(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        return new_logs

    def get_all_logs(self):
        """Get complete log history"""
        with self._lock:
            return self.log_history.copy()

    def clear_logs(self):
        """Clear all logs"""
        with self._lock:
            self.log_history.clear()
            # Clear queue
            while not self.log_queue.empty():
                try:
                    self.log_queue.get_nowait()
                except queue.Empty:
                    break

    def stop_processing(self):
        """Stop the processing thread"""
        self.is_stopped = True
        if self.active_thread and self.active_thread.is_alive():
            self.add_log("🛑 Forcefully stopping process...")
            self.processing_complete.set()
            self.active_thread = None

    def should_stop(self):
        """Check if processing should stop"""
        return self.is_stopped

    def set_current_thread(self, thread):
        """Set the current processing thread"""
        self.active_thread = thread

    def is_thread_active(self):
        """Check if processing thread is active"""
        return self.active_thread and self.active_thread.is_alive()

def process_emails(logger):
    """Process emails in a separate thread"""
    try:
        logger.add_log("🚀 Starting email processing...")
        if not logger.should_stop():
            # Pass the logger.add_log function directly to main.main
            main.main(logger.add_log)
    except Exception as e:
        logger.add_log(f"❌ Error during processing: {str(e)}")
    finally:
        logger.add_log("✅ Email processing completed.")
        st.session_state.is_processing = False

def display_logs():
    """Display logs in the UI"""
    # Get all logs from history
    logs = st.session_state.logger.get_all_logs()
    
    # Display logs
    if logs:
        # Create a container with fixed height and styling
        st.markdown("""
        <style>
        .log-container {
            height: 400px;
            overflow-y: auto;
            border: 1px solid #444;
            padding: 10px;
            font-family: monospace;
            background-color: #1E1E1E;
            color: #E0E0E0;
            white-space: pre-wrap;
            margin-bottom: 10px;
        }
        .log-line {
            margin: 0;
            padding: 2px 0;
            border-bottom: 1px solid #333;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # Create log lines HTML
        log_lines_html = ""
        for log in logs:
            log_lines_html += f'<div class="log-line">{log}</div>'
        
        # Display logs with auto-scroll
        st.markdown(f"""
        <div class="log-container" id="log-container">
            {log_lines_html}
        </div>
        <script>
            // Force scroll to bottom
            document.addEventListener('DOMContentLoaded', function() {{
                var container = document.getElementById('log-container');
                if (container) container.scrollTop = container.scrollHeight;
                
                // Set up MutationObserver to keep scrolling to bottom when content changes
                var observer = new MutationObserver(function() {{
                    if (container) container.scrollTop = container.scrollHeight;
                }});
                
                if (container) {{
                    observer.observe(container, {{ childList: true, subtree: true }});
                }}
            }});
        </script>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="height: 400px; border: 1px solid #444; padding: 10px; 
             font-family: monospace; background-color: #1E1E1E; color: #E0E0E0;">
            No logs yet...
        </div>
        """, unsafe_allow_html=True)

def main_ui():
    """Main UI function"""
    
    # Initialize session state
    if 'logger' not in st.session_state:
        st.session_state.logger = ThreadSafeLogger()
    if 'is_processing' not in st.session_state:
        st.session_state.is_processing = False
    if 'refresh_counter' not in st.session_state:
        st.session_state.refresh_counter = 0

    # Increment refresh counter on each run
    st.session_state.refresh_counter += 1

    st.title("📧 Email Analytics Dashboard")
    st.markdown("---")

    # Create columns for buttons
    col1, col2, col3 = st.columns([1, 1, 1])
    thread_active = st.session_state.logger.is_thread_active()

    with col1:
        if st.button("🚀 Run Email Processing", 
                     disabled=thread_active or st.session_state.is_processing,
                     use_container_width=True):
            st.session_state.logger.is_stopped = False
            st.session_state.is_processing = True
            processing_thread = threading.Thread(
                target=process_emails, 
                args=(st.session_state.logger,), 
                daemon=True
            )
            st.session_state.logger.set_current_thread(processing_thread)
            processing_thread.start()

    with col2:
        if st.button("🛑 Stop Processing", 
                     disabled=not (thread_active or st.session_state.is_processing),
                     use_container_width=True):
            st.session_state.logger.stop_processing()
            st.session_state.is_processing = False
            st.warning("🛑 Process stopped by user")

    with col3:
        if st.button("🧹 Clear Logs", use_container_width=True):
            if not thread_active:
                st.session_state.logger.clear_logs()

    # Show processing status
    if st.session_state.is_processing:
        st.info("⏳ Processing emails... Please wait...")

    # Logs section
    st.markdown("### 📝 Process Logs")
    
    # Display logs
    display_logs()
    
    # Get any new logs that might have been added
    new_logs = st.session_state.logger.get_new_logs()
    if new_logs:
        # Force a rerun to display the new logs
        time.sleep(0.1)
        st.rerun()
    
    # Auto-refresh during processing
    if thread_active or st.session_state.is_processing:
        time.sleep(0.5)  # Refresh every 0.5 seconds
        st.rerun()

if __name__ == "__main__":
    st.set_page_config(page_title="Email Analytics Dashboard", page_icon="📧", layout="wide")
    main_ui()