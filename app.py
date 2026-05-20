import streamlit as st
import pandas as pd
import re
import os
import time
from io import BytesIO
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from googleapiclient.errors import HttpError
import pdfplumber
from PyPDF2 import PdfReader, PdfWriter
from dotenv import load_dotenv
import requests
import json
import extra_streamlit_components as stx  # REQUIRED FOR PERSISTENCE
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Load environment variables
load_dotenv()
# Add this right after loading .env
import os
from typing import Optional

#configuration
QIK_URL = os.getenv("QIK_URL")
QIK_AUTH_TOKEN = os.getenv("QIK_AUTH_TOKEN")
QIK_SENDER = os.getenv("QIK_SENDER")
QIK_TEMPLATE_ID = os.getenv("QIK_TEMPLATE_ID")
QIK_SERVICE = os.getenv("QIK_SERVICE")
QIK_SHORTEN_URL = os.getenv("QIK_SHORTEN_URL")
HR_USERNAME = os.getenv("HR_USERNAME")
HR_PASSWORD = os.getenv("HR_PASSWORD")
SHARED_DRIVE_ID = os.getenv("SHARED_DRIVE_ID")

# Email Configuration (Add these to your .env file)
EMAIL_SMTP_SERVER = os.getenv("EMAIL_SMTP_SERVER", "smtp.gmail.com")
EMAIL_SMTP_PORT = int(os.getenv("EMAIL_SMTP_PORT", "587"))
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
REPORT_RECIPIENT = os.getenv("REPORT_RECIPIENT", "rajiv@wedtree.com")

# --- CACHED RESOURCES ---
@st.cache_resource
def get_cached_drive_service():
    """
    Singleton connection to Google Drive.
    Prevents reconnecting on every button click.
    """
    try:
        credentials_dict = {
            "type": os.getenv("GOOGLE_SERVICE_ACCOUNT_TYPE"),
            "project_id": os.getenv("GOOGLE_PROJECT_ID"),
            "private_key_id": os.getenv("GOOGLE_PRIVATE_KEY_ID"),
            "private_key": os.getenv("GOOGLE_PRIVATE_KEY", "").replace('\\n', '\n'),
            "client_email": os.getenv("GOOGLE_CLIENT_EMAIL"),
            "client_id": os.getenv("GOOGLE_CLIENT_ID"),
            "auth_uri": os.getenv("GOOGLE_AUTH_URI"),
            "token_uri": os.getenv("GOOGLE_TOKEN_URI"),
            "auth_provider_x509_cert_url": os.getenv("GOOGLE_AUTH_PROVIDER_CERT_URL"),
            "client_x509_cert_url": os.getenv("GOOGLE_CLIENT_CERT_URL")
        }
        
        # Check for critical keys
        if not credentials_dict["private_key"]:
            return None

        creds = service_account.Credentials.from_service_account_info(
            credentials_dict,
            scopes=['https://www.googleapis.com/auth/drive']
        )
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        print(f"Drive Connection Error: {e}")
        return None

class PayslipDistributorStreamlit:
    def __init__(self):
        self.shared_drive_id = SHARED_DRIVE_ID
        
        self.cookie_manager = stx.CookieManager()
        
        self.initialize_session_state()
        self.setup_page()

    def initialize_session_state(self):
        # --- FIRST restore from cookies ---
        cookie_auth = self.cookie_manager.get("hr_portal_auth")
        cookie_user = self.cookie_manager.get("hr_portal_user")
        cookie_page = self.cookie_manager.get("hr_portal_page")

        # Restore authentication BEFORE anything else
        if 'authenticated' not in st.session_state:
            if cookie_auth == "valid_token" and cookie_user:
                st.session_state.authenticated = True
                st.session_state.username = cookie_user
            else:
                st.session_state.authenticated = False
                st.session_state.username = ""

        # Restore last visited page
        if 'current_page' not in st.session_state:
            st.session_state.current_page = cookie_page if cookie_page else "payslips"

        # --- THEN initialize remaining variables ---
        if 'initialized' not in st.session_state:
            st.session_state.initialized = True
            defaults = {
                'sent_numbers': {},
                'log_entries': [],
                'current_folder': None,
                'folder_stack': [],
                'drive_initialized': False,
                'shared_drive_id': self.shared_drive_id,
                'files_processed': False,
                'uan_pages': {},
                'df': None,
                'pdf_count': 0,
                'excel_count': 0,
                'processing_complete': False,
                'results': {},
                'updated_excel_buffer': None,
                'current_path': 'Root',
                'selected_items': [],
                'drive_page': 0,
                'items_per_page': 20,
                'force_rerun': False,
                'last_sms_report_data': None  # Store last SMS report for email sending
            }

            for key, val in defaults.items():
                if key not in st.session_state:
                    st.session_state[key] = val

    def setup_page(self):
        st.set_page_config(
            page_title="Payslip Distribution System",
            page_icon="📄",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        
        # Keeping your original comprehensive CSS
        st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
        
        * {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }
        
        .main {
            background: linear-gradient(135deg, #f5f7fa 0%, #e4e8f0 100%);
        }
        
        [data-testid="stSidebar"] {
            background-color: #ffffff;
            border-right: 1px solid #e5e7eb;
        }
        
        [data-testid="stSidebar"] .stMarkdown {
            color: #374151;
        }

        .sidebar-header {
            padding: 1.5rem 1rem;
            text-align: center;
            border-bottom: 2px solid #f3f4f6;
            margin-bottom: 1.5rem;
        }
        
        .sidebar-logo {
            font-size: 3rem;
            margin-bottom: 0.5rem;
        }
        
        .sidebar-title {
            font-size: 1.3rem;
            font-weight: 700;
            color: #111827;
            margin-bottom: 0.3rem;
        }
        
        .sidebar-subtitle {
            font-size: 0.85rem;
            color: #6b7280;
        }
        
        .user-info {
            background: #f9fafb;
            padding: 1rem;
            border-radius: 12px;
            margin: 1rem;
            border: 1px solid #e5e7eb;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }
        
        .user-name {
            font-size: 0.95rem;
            font-weight: 600;
            color: #1f2937;
            margin-bottom: 0.3rem;
        }
        
        .user-status {
            font-size: 0.8rem;
            color: #10b981;
            font-weight: 500;
        }

        .nav-title {
            font-size: 0.75rem;
            font-weight: 600;
            color: #6b7280;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 0.5rem;
            padding: 0 0.5rem;
            margin-top: 1rem;
        }

        .page-header {
            background: white;
            padding: 2rem;
            border-radius: 16px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
            margin-bottom: 2rem;
            animation: fadeIn 0.5s ease-out;
        }
        
        .page-title {
            font-size: 2.5rem;
            font-weight: 700;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0.5rem;
        }
        
        .page-subtitle {
            font-size: 1.1rem;
            color: #6b7280;
            font-weight: 500;
        }
        
        .section-header {
            font-size: 1.3rem;
            font-weight: 600;
            color: #1f2937;
            margin: 2rem 0 1.5rem 0;
            padding: 1rem 1.5rem;
            background: white;
            border-left: 4px solid #667eea;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        
        .metric-card {
            background: white;
            padding: 1.8rem;
            border-radius: 16px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
            border: 1px solid #e5e7eb;
            transition: all 0.3s ease;
            height: 100%;
        }
        
        .metric-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 12px 24px rgba(102, 126, 234, 0.15);
        }
        
        .metric-icon {
            font-size: 2.5rem;
            margin-bottom: 1rem;
        }
        
        .metric-value {
            font-size: 2.8rem;
            font-weight: 800;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin: 0.5rem 0;
        }
        
        .metric-label {
            font-size: 0.9rem;
            color: #6b7280;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .log-container {
            background: white;
            border-radius: 12px;
            padding: 1.5rem;
            margin-top: 2rem;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        }
        
        .log-header {
            font-size: 1.2rem;
            font-weight: 600;
            color: #1f2937;
            margin-bottom: 1rem;
            padding-bottom: 0.75rem;
            border-bottom: 2px solid #e5e7eb;
        }
        
        .log-box {
            background: #f8fafc;
            color: #334155;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 1.5rem;
            max-height: 400px;
            overflow-y: auto;
            font-family: 'Courier New', monospace;
            font-size: 0.85rem;
            line-height: 1.6;
        }
        
        .stButton > button {
            border-radius: 8px;
            font-weight: 500;
            transition: all 0.3s ease;
            border: 2px solid #89cff0 !important; 
            font-size: 0.95rem;
        }
        
        .stButton > button:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 12px rgba(137, 207, 240, 0.4);
            border-color: #4fb3f7 !important;
        }
        
        .status-badge {
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            padding: 0.5rem 1rem;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        .status-connected { background: #d4edda; color: #155724; border: 2px solid #28a745; }
        .status-disconnected { background: #f8d7da; color: #721c24; border: 2px solid #dc3545; }
        
        .sidebar-status-box {
            padding: 0.75rem; 
            background: #f1f5f9; 
            border-radius: 8px; 
            margin: 0.5rem 0;
            border: 1px solid #e2e8f0;
        }

        .breadcrumb-container {
            background: white;
            padding: 1rem 1.5rem;
            border-radius: 12px;
            margin-bottom: 1.5rem;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .breadcrumb-item {
            color: #4b5563;
            font-weight: 500;
            font-size: 0.95rem;
        }

        .breadcrumb-separator {
            color: #9ca3af;
            margin: 0 0.5rem;
        }

        .drive-card {
            background: white;
            border-radius: 12px;
            padding: 1.5rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            margin-bottom: 1.5rem;
            transition: all 0.3s ease;
        }

        .drive-card:hover {
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }

        .file-item {
            padding: 1rem;
            border-bottom: 1px solid #f3f4f6;
            transition: background 0.2s ease;
        }

        .file-item:hover {
            background: #f9fafb;
        }

        .file-item:last-child {
            border-bottom: none;
        }

        .file-icon {
            font-size: 1.5rem;
            margin-right: 0.75rem;
        }

        .file-name {
            font-weight: 500;
            color: #1f2937;
            font-size: 0.95rem;
        }

        .file-meta {
            color: #6b7280;
            font-size: 0.85rem;
        }

        .pagination-bar {
            background: white;
            padding: 1rem;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        </style>
        """, unsafe_allow_html=True)

    # --- EMAIL FUNCTIONALITY ---
    def send_email_report(self, report_data, report_filename, report_buffer):
        """
        Send SMS report via email with attachment
        """
        try:
            if not EMAIL_SENDER or not EMAIL_PASSWORD:
                self.log_message("⚠️ Email credentials not configured. Skipping email report.")
                return False, "Email credentials not configured"
            
            # Create email
            msg = MIMEMultipart()
            msg['From'] = EMAIL_SENDER
            msg['To'] = REPORT_RECIPIENT
            msg['Subject'] = f"SMS Distribution Report - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            
            # Email body
            body = f"""
            <html>
            <body>
                <h2>SMS Distribution Report</h2>
                <p><strong>Generated on:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                
                <h3>Summary Statistics:</h3>
                <table border="1" cellpadding="5" style="border-collapse: collapse;">
                    <tr style="background-color: #f2f2f2;">
                        <th>Metric</th>
                        <th>Count</th>
                    </tr>
                    <tr><td>Total SMS Attempted</td><td>{report_data.get('total_attempted', 0)}</td></tr>
                    <tr><td>Successfully Sent</td><td>{report_data.get('sent_count', 0)}</td></tr>
                    <tr><td>API Status - DELIVERED</td><td>{report_data.get('delivered_count', 0)}</td></tr>
                    <tr><td>API Status - PROCESSED</td><td>{report_data.get('processed_count', 0)}</td></tr>
                    <tr><td>Failed to Send</td><td>{report_data.get('failed_count', 0)}</td></tr>
                    <tr><td>Skipped (No Drive Link)</td><td>{report_data.get('skipped_no_link', 0)}</td></tr>
                    <tr><td>Skipped (No Phone Number)</td><td>{report_data.get('skipped_no_phone', 0)}</td></tr>
                </table>
                
                <p><strong>Report file attached:</strong> {report_filename}</p>
                
                <hr>
                <p style="color: #666; font-size: 12px;">This is an automated report from the Payslip Distribution System.</p>
            </body>
            </html>
            """
            
            msg.attach(MIMEText(body, 'html'))
            
            # Attach the report file
            report_buffer.seek(0)
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(report_buffer.read())
            encoders.encode_base64(attachment)
            attachment.add_header(
                'Content-Disposition',
                f'attachment; filename={report_filename}'
            )
            msg.attach(attachment)
            
            # Send email
            with smtplib.SMTP(EMAIL_SMTP_SERVER, EMAIL_SMTP_PORT) as server:
                server.starttls()
                server.login(EMAIL_SENDER, EMAIL_PASSWORD)
                server.send_message(msg)
            
            self.log_message(f"📧 Email report sent successfully to {REPORT_RECIPIENT}")
            return True, "Email sent successfully"
            
        except Exception as e:
            error_msg = f"Failed to send email: {str(e)}"
            self.log_message(f"❌ {error_msg}")
            return False, error_msg

    # --- CALLBACKS & ACTIONS ---
    
    def cb_navigate_folder(self, folder_id, folder_name):
        """Callback: Open a folder"""
        st.session_state.folder_stack.append({
            'id': st.session_state.current_folder,
            'name': st.session_state.current_path
        })
        st.session_state.current_folder = folder_id
        st.session_state.current_path = folder_name
        st.session_state.drive_page = 0
        st.session_state.force_rerun = True
        self.log_message(f"📂 Opened folder: {folder_name}")

    def cb_navigate_back(self):
        """Callback: Go back one level"""
        if st.session_state.folder_stack:
            prev = st.session_state.folder_stack.pop()
            st.session_state.current_folder = prev['id']
            st.session_state.current_path = prev['name']
            st.session_state.drive_page = 0
            st.session_state.force_rerun = True
            self.log_message(f"⬅️ Navigated back to: {prev['name']}")
        else:
            self.cb_navigate_root()

    def cb_navigate_root(self):
        """Callback: Go to root"""
        st.session_state.current_folder = None
        st.session_state.current_path = 'Root'
        st.session_state.folder_stack = []
        st.session_state.drive_page = 0
        st.session_state.force_rerun = True
        self.log_message("🏠 Navigated to root")

    def cb_refresh_drive(self):
        st.session_state.drive_page = 0
        st.session_state.force_rerun = True
        self.log_message("🔄 Drive view refreshed")

    def cb_switch_page(self, page_name):
        st.session_state.current_page = page_name
        # Persist the current page to cookie so refresh stays here
        self.cookie_manager.set("hr_portal_page", page_name, key="cookie_set_page")
        st.session_state.force_rerun = True

    def cb_logout(self):

        # Safely delete cookies
        for ck in ["hr_portal_auth", "hr_portal_user", "hr_portal_page"]:
            try:
                self.cookie_manager.delete(ck)
            except KeyError:
                pass
            except Exception:
                pass

        # Reset session state (safe)
        st.session_state.authenticated = False
        st.session_state.username = ''
        st.session_state.drive_initialized = False
        st.session_state.current_folder = None
        st.session_state.folder_stack = []
        st.session_state.current_page = 'payslips'
        st.session_state.force_rerun = True

        self.log_message("👋 User logged out")

    def cb_disconnect_drive(self):
        """Disconnect drive connection manually"""
        st.session_state.drive_initialized = False
        st.session_state.force_rerun = True
        self.log_message("🔌 Drive disconnected manually")

    def cb_send_email_report(self):
        """Callback to send email report manually"""
        if st.session_state.last_sms_report_data:
            with st.spinner("📧 Sending email report..."):
                success, message = self.send_email_report(
                    st.session_state.last_sms_report_data['summary'],
                    st.session_state.last_sms_report_data['filename'],
                    st.session_state.last_sms_report_data['buffer']
                )
                if success:
                    st.success(f"✅ {message}")
                else:
                    st.error(f"❌ {message}")
        else:
            st.warning("No report data available to send")

    # --- RENDER METHODS ---

    def render_welcome_screen(self):
        """Render the welcome screen in the main area when not logged in"""
        st.markdown("""
        <div style="text-align: center; padding: 4rem 2rem;">
            <div style="font-size: 5rem; margin-bottom: 1rem;">👋</div>
            <h1 style="color: #1f2937; font-weight: 800; margin-bottom: 1rem;">Welcome Back!</h1>
            <p style="font-size: 1.2rem; color: #6b7280; max-width: 600px; margin: 0 auto;">
                Please log in using the form in the sidebar to access the Payslip Distribution System and HR Portal tools.
            </p>
            <div style="margin-top: 3rem; padding: 2rem; background: white; border-radius: 16px; display: inline-block; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                <div style="color: #667eea; font-weight: 600; margin-bottom: 0.5rem;">System Status</div>
                <div style="display: flex; gap: 2rem; justify-content: center;">
                    <div>
                        <div style="font-size: 1.5rem;">🔒</div>
                        <div style="font-size: 0.9rem; color: #6b7280;">Secure Access</div>
                    </div>
                    <div>
                        <div style="font-size: 1.5rem;">⚡</div>
                        <div style="font-size: 0.9rem; color: #6b7280;">Fast Processing</div>
                    </div>
                    <div>
                        <div style="font-size: 1.5rem;">☁️</div>
                        <div style="font-size: 0.9rem; color: #6b7280;">Cloud Storage</div>
                    </div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    def render_sidebar(self):
        """Render the sidebar navigation and login"""
        with st.sidebar:
            st.markdown(f"""
            <div class="sidebar-header">
                <div class="sidebar-logo">📄</div>
                <div class="sidebar-title">Payslip System</div>
                <div class="sidebar-subtitle">HR Portal</div>
            </div>
            """, unsafe_allow_html=True)

            if not st.session_state.authenticated:
                st.markdown('<div class="nav-title">🔐 Login Required</div>', unsafe_allow_html=True)
                
                with st.form("sidebar_login_form", clear_on_submit=False):
                    username = st.text_input("Username", placeholder="Enter ID")
                    password = st.text_input("Password", type="password", placeholder="Enter Password")
                    
                    submit = st.form_submit_button("🚀 Sign In", use_container_width=True, type="primary")

                    if submit:
                        if username == HR_USERNAME and password == HR_PASSWORD:
                            st.session_state.authenticated = True
                            st.session_state.username = username
                            
                            # --- SET COOKIES ON SUCCESSFUL LOGIN ---
                            self.cookie_manager.set("hr_portal_auth", "valid_token", key="set_cookie_auth")
                            self.cookie_manager.set("hr_portal_user", username, key="set_cookie_user")
                            # ---------------------------------------

                            st.session_state.force_rerun = True
                            self.log_message(f"✅ User logged in: {username}")
                            st.success("Login successful!")
                        else:
                            st.error("Invalid credentials")
                
                st.info("Please enter your HR credentials to access the system.")

            else:
                st.markdown(f"""
                <div class="user-info">
                    <div class="user-name">👤 {st.session_state.username}</div>
                    <div class="user-status">● Online</div>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown('<div class="nav-title">📍 Navigation</div>', unsafe_allow_html=True)
                
                # NAVIGATION BUTTONS
                st.button("📤 Send Payslips", use_container_width=True, 
                          type="primary" if st.session_state.current_page == 'payslips' else "secondary",
                          on_click=self.cb_switch_page, args=('payslips',))
                
                st.button("📁 Drive Explorer", use_container_width=True,
                          type="primary" if st.session_state.current_page == 'drive' else "secondary",
                          on_click=self.cb_switch_page, args=('drive',))
                
                st.button("📨 Send SMS", use_container_width=True,
                          type="primary" if st.session_state.current_page == 'sms' else "secondary",
                          on_click=self.cb_switch_page, args=('sms',))
                
                st.markdown("---")
                st.markdown('<div class="nav-title">⚡ Quick Actions</div>', unsafe_allow_html=True)
                
                if st.button("🔄 Refresh Session", use_container_width=True):
                    st.session_state.force_rerun = True
                
                if st.button("🗑️ Clear Data", use_container_width=True):
                    self.reset_session()
                
                st.markdown("---")
                # Logout calls cb_logout to clear cookies
                st.button("🚪 Logout", use_container_width=True, type="secondary", on_click=self.cb_logout)

    def run(self):
        self.render_sidebar()
        
        # Check if we need to force a rerun due to state changes
        if st.session_state.get('force_rerun', False):
            st.session_state.force_rerun = False
            st.rerun()
        
        if not st.session_state.authenticated:
            self.render_welcome_screen()
            return
        
        if st.session_state.current_page == 'payslips':
            self.render_payslips_page()
        elif st.session_state.current_page == 'drive':
            self.render_drive_page()
        elif st.session_state.current_page == 'sms':
            self.render_sms_page()
        
        self.render_activity_log()

    # --- PAGES ---

    def render_payslips_page(self):
        """Render the Send Payslips page"""
        st.markdown("""
        <div class="page-header">
            <div class="page-title">📤 Send Payslips</div>
            <div class="page-subtitle">Upload and distribute employee payslips via Google Drive</div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown('<div class="section-header">🔗 Google Drive Connection</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([2, 2, 1])
        with col1:
            # ONLY SHOW INITIALIZE IF NOT CONNECTED
            if not st.session_state.drive_initialized:
                if st.button("🚀 Initialize Connection", type="primary", use_container_width=True):
                    self.initialize_drive()
            else:
                st.success("✅ Drive is already connected")

        with col2:
            if st.session_state.drive_initialized:
                st.markdown('''
                <div class="status-badge status-connected">
                    <div style="width: 8px; height: 8px; background: #28a745; border-radius: 50%;"></div>
                    <span>Connected</span>
                </div>
                ''', unsafe_allow_html=True)
            else:
                st.markdown('''
                <div class="status-badge status-disconnected">
                    <div style="width: 8px; height: 8px; background: #dc3545; border-radius: 50%;"></div>
                    <span>Disconnected</span>
                </div>
                ''', unsafe_allow_html=True)
        with col3:
            if st.session_state.drive_initialized:
                # Disconnect logic now handled strictly via manual callback
                st.button("🛑 Disconnect", use_container_width=True, on_click=self.cb_disconnect_drive)

        if not st.session_state.drive_initialized:
            st.warning("⚠️ Please initialize Google Drive connection before proceeding.")
            return

        st.markdown('<div class="section-header">📤 Upload Files</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**📄 PDF Payslip File**")
            pdf_file = st.file_uploader(
                "Upload consolidated payslip PDF",
                type=["pdf"],
                key="pdf_upload",
                help="Upload the PDF file containing all employee payslips"
            )
            if pdf_file:
                st.success(f"✅ {pdf_file.name} ({pdf_file.size / 1024:.1f} KB)")
        
        with col2:
            st.markdown("**📊 Employee Contact File**")
            excel_file = st.file_uploader(
                "Upload employee data Excel",
                type=["xlsx", "xls"],
                key="excel_upload",
                help="Excel file with columns: Employee Name, Employee no, UAN"
            )
            if excel_file:
                st.success(f"✅ {excel_file.name} ({excel_file.size / 1024:.1f} KB)")

        st.markdown("---")
        
        if st.button("🔍 Process & Analyze Files", type="primary", use_container_width=True):
            if not pdf_file or not excel_file:
                st.error("❌ Please upload both PDF and Excel files")
            else:
                self.process_files(pdf_file, excel_file)

        if st.session_state.files_processed:
            st.markdown('<div class="section-header">📊 File Analysis Results</div>', unsafe_allow_html=True)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-icon">📄</div>
                    <div class="metric-label">UANs Found</div>
                    <div class="metric-value">{st.session_state.pdf_count}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-icon">👥</div>
                    <div class="metric-label">Employees</div>
                    <div class="metric-value">{st.session_state.excel_count}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                match_rate = (min(st.session_state.pdf_count, st.session_state.excel_count) / 
                             max(st.session_state.excel_count, 1) * 100)
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-icon">🎯</div>
                    <div class="metric-label">Match Rate</div>
                    <div class="metric-value">{match_rate:.0f}%</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('<div class="section-header">🚀 Upload Payslips to Drive</div>', unsafe_allow_html=True)
            
            col1, col2 = st.columns([3, 1])
            with col1:
                confirm = st.checkbox(
                    "✅ I confirm I want to process and upload payslips for all employees",
                    key="confirm_checkbox",
                    help="This will upload individual payslips to Google Drive"
                )
            with col2:
                if st.button("📤 Start Upload", type="primary", disabled=not confirm, use_container_width=True):
                    self.process_payslips()

        if st.session_state.processing_complete:
            st.markdown('<div class="section-header">📊 Upload Results</div>', unsafe_allow_html=True)
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-icon">✅</div>
                    <div class="metric-label">Uploaded</div>
                    <div class="metric-value" style="color: #28a745;">{st.session_state.results.get('uploaded', 0)}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-icon">❌</div>
                    <div class="metric-label">Failed</div>
                    <div class="metric-value" style="color: #dc3545;">{st.session_state.results.get('failed', 0)}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-icon">⭐</div>
                    <div class="metric-label">Skipped</div>
                    <div class="metric-value" style="color: #ffc107;">{st.session_state.results.get('skipped', 0)}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                total = sum([st.session_state.results.get(k, 0) for k in ['uploaded', 'failed', 'skipped']])
                success_rate = (st.session_state.results.get('uploaded', 0) / max(total, 1) * 100)
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-icon">📈</div>
                    <div class="metric-label">Success Rate</div>
                    <div class="metric-value" style="color: #667eea;">{success_rate:.0f}%</div>
                </div>
                """, unsafe_allow_html=True)

            if st.session_state.updated_excel_buffer:
                st.markdown("---")
                st.download_button(
                    label="📥 Download Updated Excel File with Drive Links",
                    data=st.session_state.updated_excel_buffer,
                    file_name=f"payslip_links_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

    def render_drive_page(self):
        """Render the Drive Explorer page"""
        st.markdown("""
        <div class="page-header">
            <div class="page-title">📁 Drive Explorer</div>
            <div class="page-subtitle">Browse, manage and download files from Google Drive</div>
        </div>
        """, unsafe_allow_html=True)
        
        if not st.session_state.drive_initialized:
            st.warning("⚠️ Please initialize Google Drive connection in the 'Send Payslips' page first.")
            return

        # Breadcrumb Navigation
        breadcrumb_parts = ["🏠 Root"]
        if st.session_state.current_path != 'Root':
            breadcrumb_parts.append(st.session_state.current_path)
        
        breadcrumb = " <span class='breadcrumb-separator'>›</span> ".join(breadcrumb_parts)
        st.markdown(f'<div class="breadcrumb-container"><span class="breadcrumb-item">{breadcrumb}</span></div>', unsafe_allow_html=True)

        col1, col2, col3, col4 = st.columns([2, 2, 2, 3])
        
        with col1:
            st.button("🏠 Go to Root", use_container_width=True, 
                      disabled=not st.session_state.current_folder,
                      on_click=self.cb_navigate_root)
        
        with col2:
            st.button("⬅️ Go Back", use_container_width=True, 
                      disabled=not st.session_state.folder_stack,
                      on_click=self.cb_navigate_back)
        
        with col3:
            st.button("🔄 Refresh", use_container_width=True,
                      on_click=self.cb_refresh_drive)
        
        with col4:
            items_per_page = st.selectbox(
                "Items per page:",
                options=[10, 20, 30, 50],
                index=[10, 20, 30, 50].index(st.session_state.items_per_page),
                key="items_selector"
            )
            if items_per_page != st.session_state.items_per_page:
                st.session_state.items_per_page = items_per_page
                st.session_state.drive_page = 0
                st.session_state.force_rerun = True
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Display Drive Contents
        self.display_drive_contents()

    def render_sms_page(self):
        """Render the SMS Distribution page"""
        st.markdown("""
        <div class="page-header">
            <div class="page-title">📨 SMS Distribution</div>
            <div class="page-subtitle">Send payslip notifications to employees via SMS</div>
        </div>
        """, unsafe_allow_html=True)
        
        st.info("📋 Upload the updated Excel file (with Drive Links) from the Send Payslips page to send SMS notifications to employees.")
        
        if not st.session_state.drive_initialized:
            st.warning("⚠️ Please initialize Google Drive connection first.")
            return

        uploaded = st.file_uploader(
            "📊 Upload Updated Excel File",
            type=["xlsx", "xls"],
            key="sms_excel_upload",
            help="Excel file must contain: Employee Name, Employee no, UAN, Drive Link"
        )
        
        if uploaded:
            st.success(f"✅ File uploaded: {uploaded.name}")
            
            try:
                preview_df = pd.read_excel(uploaded, dtype=str).fillna("")
                st.markdown("**📋 Data Preview (First 5 rows)**")
                st.dataframe(preview_df.head(), use_container_width=True)

                has_links = 0
                if 'Drive Link' in preview_df.columns:
                    has_links = (preview_df['Drive Link'].str.strip() != "").sum()
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-icon">👥</div>
                        <div class="metric-label">Total Employees</div>
                        <div class="metric-value">{len(preview_df)}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-icon">🔗</div>
                        <div class="metric-label">With Drive Links</div>
                        <div class="metric-value">{has_links}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-icon">📨</div>
                        <div class="metric-label">Ready to Send</div>
                        <div class="metric-value">{has_links}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"Error previewing file: {str(e)}")

        st.markdown("---")
        
        if st.button("📤 Send SMS to All Employees with Drive Links", 
                     type="primary", 
                     use_container_width=True,
                     disabled=not uploaded):
            if uploaded:
                self.process_and_send_sms(uploaded)
        
        # Add email report button if report data exists
        if st.session_state.last_sms_report_data:
            st.markdown("---")
            st.markdown('<div class="section-header">📧 Email Report</div>', unsafe_allow_html=True)
            
            col1, col2 = st.columns([3, 1])
            with col1:
                st.info(f"Report will be sent to: {REPORT_RECIPIENT}")
            with col2:
                st.button("📧 Send Report via Email", 
                         use_container_width=True, 
                         type="primary",
                         on_click=self.cb_send_email_report)

    def render_activity_log(self):
        """Render activity log"""
        st.markdown("""
        <div class="log-container">
            <div class="log-header">📜 Activity Log</div>
        """, unsafe_allow_html=True)
        
        if st.session_state.log_entries:
            log_html = '<div class="log-box">'
            for entry in reversed(st.session_state.log_entries[-100:]):
                log_html += f'<div>{entry}</div>'
            log_html += '</div>'
            st.markdown(log_html, unsafe_allow_html=True)
        else:
            st.info("💤 No activity yet. Start by initializing Google Drive connection.")
        
        st.markdown('</div>', unsafe_allow_html=True)

    # --- LOGIC IMPLEMENTATIONS ---

    def initialize_drive(self):
        with st.spinner("🔄 Connecting to Google Drive..."):
            try:
                service = get_cached_drive_service()
                if service:
                    service.files().list(pageSize=1, supportsAllDrives=True).execute()
                    # SET THE FLAG TO TRUE
                    st.session_state.drive_initialized = True
                    st.session_state.force_rerun = True
                    st.success("✅ Google Drive connected successfully!")
                    self.log_message("✅ Google Drive connection initialized")
                else:
                    st.session_state.drive_initialized = False
                    st.error("❌ Failed to connect to Google Drive (Check env vars)")
                    self.log_message("❌ Failed to initialize Google Drive")
            except Exception as e:
                st.session_state.drive_initialized = False
                st.error(f"❌ Connection error: {str(e)}")
                self.log_message(f"❌ Drive connection error: {str(e)}")

    def process_files(self, pdf_file, excel_file):
        progress_container = st.empty()
        
        with progress_container:
            with st.spinner("🔄 Processing files..."):
                try:
                    st.info("📄 Extracting UAN information from PDF...")
                    uan_pages = self.process_pdf(pdf_file)
                    st.session_state.pdf_count = len(uan_pages)
                    
                    st.info("📊 Loading employee data from Excel...")
                    df = pd.read_excel(excel_file, dtype=str).fillna("")
                    
                    if 'UAN/member ID' in df.columns and 'UAN' not in df.columns:
                        df.rename(columns={'UAN/member ID': 'UAN'}, inplace=True)
                    
                    if 'UAN' not in df.columns:
                        st.error("❌ Excel must have 'UAN' or 'UAN/member ID' column")
                        return
                    
                    if 'Employee Name' not in df.columns:
                        if 'Name' in df.columns:
                            df.rename(columns={'Name': 'Employee Name'}, inplace=True)
                        else:
                            df['Employee Name'] = ""
                    
                    if 'Employee no' not in df.columns:
                        if 'Employee No' in df.columns:
                            df.rename(columns={'Employee No': 'Employee no'}, inplace=True)
                        else:
                            st.error("❌ Excel must have 'Employee no' column")
                            return
                    
                    st.session_state.excel_count = len(df)
                    st.session_state.uan_pages = uan_pages
                    st.session_state.df = df
                    st.session_state.files_processed = True
                    st.session_state.processing_complete = False
                    # Store file objects for later
                    st.session_state.pdf_file_obj = pdf_file
                    st.session_state.excel_file_obj = excel_file
                    
                    progress_container.empty()
                    st.success("✅ Files processed successfully!")
                    self.log_message(f"📊 PDF: {st.session_state.pdf_count} UANs | Excel: {st.session_state.excel_count} employees")
                    
                except Exception as e:
                    st.error(f"❌ Error processing files: {str(e)}")
                    self.log_message(f"❌ Processing error: {str(e)}")

    def process_payslips(self):
        if not st.session_state.files_processed or st.session_state.df is None:
            st.error("❌ Please process files first")
            return
        
        if not st.session_state.drive_initialized:
            st.error("❌ Google Drive not connected")
            return

        today = datetime.now()
        first_day_of_current_month = today.replace(day=1)
        last_month = first_day_of_current_month - timedelta(days=1)
        current_month = last_month.strftime("%B %Y")

        progress_placeholder = st.empty()
        status_placeholder = st.empty()
        
        service = get_cached_drive_service()

        with st.spinner(f"🚀 Uploading payslips for {current_month}..."):
            try:
                folder_id = self.get_monthly_folder_id(
                    service,
                    current_month,
                    st.session_state.shared_drive_id
                )
                
                if not folder_id:
                    st.error("❌ Failed to create monthly folder")
                    return

                results = {'uploaded': 0, 'failed': 0, 'skipped': 0}
                updated_df = st.session_state.df.copy()
                
                if 'Drive Link' not in updated_df.columns:
                    updated_df['Drive Link'] = ""

                total = len(updated_df)
                
                # Retrieve the PDF file object from session state
                pdf_file_obj = st.session_state.pdf_file_obj
                
                for idx, (_, row) in enumerate(updated_df.iterrows()):
                    uan = str(row.get('UAN', '')).strip()
                    emp_name = str(row.get('Employee Name', '')).strip()
                    
                    progress = (idx + 1) / total
                    progress_placeholder.progress(progress)
                    status_placeholder.info(f"⏳ Processing {idx + 1}/{total}: {emp_name} (UAN: {uan})")
                    
                    if not uan:
                        self.log_message(f"⚠️ Skipped row {idx+1}: Missing UAN")
                        results['skipped'] += 1
                        continue
                    
                    if uan in st.session_state.uan_pages:
                        try:
                            pdf_file_obj.seek(0)
                            page_num = st.session_state.uan_pages[uan]
                            pdf_buffer = self.extract_individual_payslip(
                                pdf_file_obj,
                                uan,
                                page_num
                            )
                            
                            emp_no = str(row.get("Employee no", "")).strip()
                            safe_name = emp_name.replace("/", "-").replace("\\", "-")
                            filename = f"{safe_name}_{emp_no}_{current_month.replace(' ', '_')}.pdf"
                            drive_link, file_id = self.upload_to_drive(
                                service,
                                pdf_buffer,
                                filename,
                                folder_id
                            )
                            
                            if drive_link:
                                match_index = updated_df[updated_df['UAN'].astype(str) == str(uan)].index
                                if not match_index.empty:
                                    updated_df.loc[match_index, 'Drive Link'] = drive_link
                                
                                st.session_state.sent_numbers[uan] = {
                                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    'drive_link': drive_link,
                                    'file_id': file_id
                                }
                                
                                self.log_message(f"✅ Uploaded: {emp_name} (UAN: {uan})")
                                results['uploaded'] += 1
                            else:
                                self.log_message(f"❌ Upload failed: {emp_name} (UAN: {uan})")
                                results['failed'] += 1
                                
                        except Exception as e:
                            self.log_message(f"❌ Error processing {emp_name}: {str(e)}")
                            results['failed'] += 1
                    else:
                        self.log_message(f"⚠️ No payslip found for {emp_name} (UAN: {uan})")
                        results['skipped'] += 1
                
                final_df = updated_df.copy()
                for col in ['Employee Name', 'Employee no', 'UAN', 'Drive Link']:
                    if col not in final_df.columns:
                        final_df[col] = ""
                
                final_df = final_df[['Employee Name', 'Employee no', 'UAN', 'Drive Link']]
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Payslips')
                
                st.session_state.updated_excel_buffer = output.getvalue()
                st.session_state.results = results
                st.session_state.processing_complete = True
                
                progress_placeholder.empty()
                status_placeholder.empty()
                
                st.success(f"🎉 Upload complete! {results['uploaded']} uploaded, {results['failed']} failed, {results['skipped']} skipped")
                self.log_message(f"🎉 Processing complete: {results}")
                
            except Exception as e:
                st.error(f"❌ Processing error: {str(e)}")
                self.log_message(f"❌ Processing error: {str(e)}")

    def process_pdf(self, pdf_file):
        uan_pages = {}
        try:
            pdf_file.seek(0)
            with pdfplumber.open(pdf_file) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    match = re.search(r"(?:UAN\/?MEMBER ID|UAN|UAN MEMBER ID)[:\s]*([A-Za-z0-9\-]+)", text, re.I)
                    if match:
                        uan = match.group(1).strip()
                        uan_pages[uan] = page_num
            return uan_pages
        except Exception as e:
            st.error(f"❌ PDF processing error: {str(e)}")
            return {}

    def extract_individual_payslip(self, pdf_file, uan, page_num):
        pdf_reader = PdfReader(pdf_file)
        pdf_writer = PdfWriter()
        pdf_writer.add_page(pdf_reader.pages[page_num])
        output_buffer = BytesIO()
        pdf_writer.write(output_buffer)
        output_buffer.seek(0)
        return output_buffer

    def get_monthly_folder_id(self, service, month_year, shared_drive_id):
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                query = f"name='{month_year}' and mimeType='application/vnd.google-apps.folder' and trashed=false and '{shared_drive_id}' in parents"
                results = service.files().list(
                    q=query,
                    fields="files(id, name)",
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True,
                    corpora='drive',
                    driveId=shared_drive_id
                ).execute()
                
                folders = results.get('files', [])
                if folders:
                    return folders[0]['id']
                
                folder_metadata = {
                    'name': month_year,
                    'mimeType': 'application/vnd.google-apps.folder',
                    'parents': [shared_drive_id]
                }
                folder = service.files().create(
                    body=folder_metadata,
                    fields='id',
                    supportsAllDrives=True
                ).execute()
                
                self.log_message(f"📁 Created folder: {month_year}")
                return folder['id']
                
            except HttpError as error:
                if hasattr(error, 'resp') and error.resp.status == 503:
                    retry_count += 1
                    time.sleep(5)
                    continue
                st.error(f"❌ Drive API error: {error}")
                return None
            except Exception as e:
                st.error(f"❌ Folder creation error: {str(e)}")
                return None
        
        return None

    def upload_to_drive(self, service, file_buffer, filename, folder_id):
        try:
            # 1️⃣ Check for existing file
            query = (
                f"name='{filename}' and '{folder_id}' in parents "
                f"and trashed=false"
            )

            existing = service.files().list(
                q=query,
                fields="files(id)",
                supportsAllDrives=True
            ).execute()

            # 2️⃣ If exists → delete it
            if existing.get("files"):
                old_id = existing["files"][0]["id"]
                service.files().delete(
                    fileId=old_id,
                    supportsAllDrives=True
                ).execute()
                self.log_message(f"♻️ Replaced existing file: {filename}")

            # 3️⃣ Upload new file
            file_buffer.seek(0)
            media = MediaIoBaseUpload(file_buffer, mimetype='application/pdf', resumable=True)
            file_metadata = {'name': filename, 'parents': [folder_id]}

            file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id',
                supportsAllDrives=True
            ).execute()

            permission = {'type': 'anyone', 'role': 'reader', 'allowFileDiscovery': False}
            service.permissions().create(
                fileId=file['id'], body=permission, supportsAllDrives=True
            ).execute()

            file_link = f"https://drive.google.com/uc?export=download&id={file['id']}"
            return file_link, file['id']

        except Exception as e:
            st.error(f"❌ Upload error: {str(e)}")
            return None, None

    def list_drive_contents(self, service, shared_drive_id, folder_id=None):
        max_retries = 3

        for attempt in range(max_retries):
            try:
                target_id = folder_id if folder_id else shared_drive_id
                q = f"'{target_id}' in parents and trashed=false"
                
                results = service.files().list(
                    q=q,
                    fields="files(id, name, mimeType, size, createdTime, modifiedTime, webViewLink)",
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True,
                    orderBy="folder,name",
                    corpora='drive',
                    driveId=shared_drive_id
                ).execute()
                
                return results.get('files', [])

            except Exception as e:
                # SSL/TLS errors → force reconnect
                err = str(e)
                if "SSL" in err or "DECRYPTION_FAILED" in err:
                    self.log_message("⚠️ SSL error, reconnecting Google Drive...")
                    time.sleep(1)
                    service = get_cached_drive_service()
                    continue

                st.error(f"❌ Error listing contents: {str(e)}")
                return []

        st.error("❌ Failed after multiple retries.")
        return []

    def display_drive_contents(self):
        """Enhanced drive contents display with better UX"""
        try:
            service = get_cached_drive_service()
            contents = self.list_drive_contents(
                service,
                st.session_state.shared_drive_id,
                st.session_state.current_folder
            )
            
            if not contents:
                st.info("📂 This folder is empty.")
                return
            
            # Pagination setup
            items_per_page = st.session_state.items_per_page
            total_items = len(contents)
            total_pages = (total_items + items_per_page - 1) // items_per_page
            current_page = st.session_state.drive_page
            
            if current_page >= total_pages:
                st.session_state.drive_page = max(0, total_pages - 1)
                current_page = st.session_state.drive_page
            
            start_idx = current_page * items_per_page
            end_idx = min(start_idx + items_per_page, total_items)
            page_contents = contents[start_idx:end_idx]
            
            # Summary Info
            folders_count = sum(1 for item in contents if item['mimeType'] == 'application/vnd.google-apps.folder')
            files_count = total_items - folders_count
            
            st.markdown(f"""
            <div class="drive-card">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem;">
                    <div>
                        <span style="font-weight: 600; color: #1f2937;">📊 Total Items: {total_items}</span>
                        <span style="margin-left: 1.5rem; color: #6b7280;">📁 Folders: {folders_count}</span>
                        <span style="margin-left: 1rem; color: #6b7280;">📄 Files: {files_count}</span>
                    </div>
                    <div style="color: #6b7280; font-size: 0.9rem;">
                        Showing {start_idx + 1}-{end_idx} of {total_items}
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            
            for item in page_contents:
                is_folder = item['mimeType'] == 'application/vnd.google-apps.folder'
                file_icon = "📁" if is_folder else "📄"
                modified = item.get('modifiedTime', '')[:10] if 'modifiedTime' in item else "-"
                
                st.markdown('<div class="file-item">', unsafe_allow_html=True)
                
                col1, col2, col3 = st.columns([6, 3, 2])
                
                with col1:
                    st.markdown(f"""
                    <div style="display: flex; align-items: center;">
                        <span class="file-icon">{file_icon}</span>
                        <div>
                            <div class="file-name">{item['name']}</div>
                            <div class="file-meta">Modified: {modified}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    if is_folder:
                        st.button("📂 Open Folder", key=f"open_{item['id']}", 
                                  use_container_width=True,
                                  on_click=self.cb_navigate_folder, args=(item['id'], item['name']))
                    else:
                        st.button("⬇️ Download", key=f"dl_{item['id']}", 
                                  use_container_width=True,
                                  on_click=self.download_file, args=(item['id'], item['name']))
                
                with col3:
                    if st.button("🗑️ Delete", key=f"del_{item['id']}", use_container_width=True, type="secondary"):
                        self.confirm_delete(item['id'], item['name'], is_folder)
                
                st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Pagination controls
            if total_pages > 1:
                st.markdown('<div class="pagination-bar">', unsafe_allow_html=True)
                
                col1, col2, col3, col4, col5 = st.columns([1, 1, 3, 1, 1])
                
                with col1:
                    if st.button("⏮️ First", disabled=current_page == 0, use_container_width=True, key="first_page"):
                        st.session_state.drive_page = 0
                        st.session_state.force_rerun = True
                
                with col2:
                    if st.button("◀️ Prev", disabled=current_page == 0, use_container_width=True, key="prev_page"):
                        st.session_state.drive_page = max(0, current_page - 1)
                        st.session_state.force_rerun = True
                
                with col3:
                    st.markdown(f'<div style="text-align: center; font-weight: 600; color: #4b5563; padding-top: 0.5rem;">Page {current_page + 1} of {total_pages}</div>', unsafe_allow_html=True)
                
                with col4:
                    if st.button("Next ▶️", disabled=current_page >= total_pages - 1, use_container_width=True, key="next_page"):
                        st.session_state.drive_page = min(total_pages - 1, current_page + 1)
                        st.session_state.force_rerun = True
                
                with col5:
                    if st.button("Last ⏭️", disabled=current_page >= total_pages - 1, use_container_width=True, key="last_page"):
                        st.session_state.drive_page = total_pages - 1
                        st.session_state.force_rerun = True
                
                st.markdown('</div>', unsafe_allow_html=True)
                    
        except Exception as e:
            st.error(f"❌ Error displaying contents: {str(e)}")
            self.log_message(f"❌ Display error: {str(e)}")

    def confirm_delete(self, file_id, file_name, is_folder):
        item_type = "folder" if is_folder else "file"
        
        @st.dialog(f"Delete {item_type}?")
        def delete_dialog():
            st.warning(f"⚠️ Are you sure you want to delete **{file_name}**?")
            if is_folder:
                st.error("🚨 This will delete the folder and all its contents permanently!")
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ Yes, Delete", type="primary", use_container_width=True):
                    self.delete_item(file_id, file_name, is_folder)
                    st.session_state.force_rerun = True
            with col2:
                if st.button("❌ Cancel", use_container_width=True):
                    pass
        
        delete_dialog()

    def delete_item(self, file_id, file_name, is_folder):
        try:
            service = get_cached_drive_service()
            service.files().delete(
                fileId=file_id,
                supportsAllDrives=True
            ).execute()
            
            item_type = "folder" if is_folder else "file"
            st.success(f"✅ Successfully deleted {item_type}: {file_name}")
            self.log_message(f"🗑️ Deleted {item_type}: {file_name}")
            
        except Exception as e:
            st.error(f"❌ Failed to delete: {str(e)}")
            self.log_message(f"❌ Delete failed: {file_name} - {str(e)}")

    def download_file(self, file_id, filename):
        with st.spinner(f"⬇️ Downloading {filename}..."):
            try:
                service = get_cached_drive_service()
                request = service.files().get_media(
                    fileId=file_id,
                    supportsAllDrives=True
                )
                file_buffer = BytesIO()
                downloader = MediaIoBaseDownload(file_buffer, request)
                
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                
                file_buffer.seek(0)
                st.download_button(
                    label=f"💾 Save {filename}",
                    data=file_buffer.getvalue(),
                    file_name=filename,
                    mime="application/octet-stream",
                    key=f"save_{file_id}",
                    use_container_width=True
                )
                self.log_message(f"📥 Downloaded: {filename}")
                
            except Exception as e:
                st.error(f"❌ Download failed: {str(e)}")
                self.log_message(f"❌ Download failed: {filename}")

    def format_phone_number(self, phone_str):
        if pd.isna(phone_str) or str(phone_str).strip() == "":
            return ""
        
        digits = re.sub(r'\D', '', str(phone_str))
        
        if len(digits) == 10:
            return f"+91{digits}"
        elif len(digits) == 11 and digits.startswith('0'):
            return f"+91{digits[1:]}"
        elif digits.startswith('91') and len(digits) == 12:
            return f"+{digits}"
        elif str(phone_str).strip().startswith("+"):
            return str(phone_str).strip()
        
        return digits

    def send_sms_via_qik(self, name, phone, download_link, previous_month):
        """Send SMS and return delivery tracking info"""
        if not phone:
            return False, "Empty phone", None
        
        template_message = (
            f"Hello! {name},\n\n"
            f"Your payslip for {previous_month} is ready.\n\n"
            f"Download Link: {download_link}\n\n"
            "This link will expire in 30 days.\n\n"
            "Regards\n"
            "HR Team - PRASHANTI"
        )
        
        payload = {
            "to": phone,
            "sender": QIK_SENDER,
            "service": QIK_SERVICE,
            "template_id": QIK_TEMPLATE_ID,
            "shorten_url": QIK_SHORTEN_URL,
            "message": template_message
        }
        
        headers = {
            "Authorization": QIK_AUTH_TOKEN,
            "Content-Type": "application/json"
        }
        
        try:
            resp = requests.post(QIK_URL, headers=headers, json=payload, timeout=30)
            
            # Extract message_id from response
            message_id = None
            response_data = None
            full_response = resp.text
            
            if resp.status_code in (200, 201):
                try:
                    response_data = resp.json()
                    # Extract message_id from the data array
                    if response_data and 'data' in response_data:
                        data_array = response_data['data']
                        if data_array and len(data_array) > 0:
                            # Get message_id from first item in data array
                            message_id = data_array[0].get('message_id')
                    
                    # If not found in data array, try other common locations
                    if not message_id:
                        message_id = (
                            response_data.get('message_id') or 
                            response_data.get('messageId') or 
                            response_data.get('id')
                        )
                except json.JSONDecodeError:
                    pass
            
            return (resp.status_code in (200, 201)), full_response, message_id
            
        except Exception as e:
            return False, str(e), None
            
    
        
    def get_sms_status(self, message_ids):
        """
        Fetch SMS delivery status for multiple message IDs
        Returns raw API response without any mapping
        """
        if not message_ids:
            return {}
        
        # Handle single ID or list
        if isinstance(message_ids, str):
            message_ids = [message_ids]
        
        # Filter out None or empty values
        message_ids = [mid for mid in message_ids if mid]
        
        if not message_ids:
            return {}
        
        headers = {
            "Authorization": QIK_AUTH_TOKEN,
            "Content-Type": "application/json"
        }
        
        status_url = "https://rest.qikberry.ai/v1/sms/status"
        
        results = {}
        
        # Make individual calls for each message_id
        for message_id in message_ids:
            params = {
                "message_ids": message_id
            }
            
            try:
                self.log_message(f"🔍 Fetching status for message_id: {message_id}")
                
                resp = requests.get(status_url, headers=headers, params=params, timeout=30)
                
                self.log_message(f"📡 Status API Response Code: {resp.status_code} for {message_id}")
                
                if resp.status_code == 200:
                    try:
                        data = resp.json()
                        # Store the full response as JSON string
                        full_response_json = json.dumps(data, indent=2)
                        
                        # Parse the response to extract all fields
                        if isinstance(data, dict) and 'data' in data and isinstance(data['data'], list):
                            for item in data['data']:
                                mid = item.get('message_id')
                                if mid == message_id:
                                    # Store ALL raw data from the API without any mapping
                                    results[message_id] = {
                                        # Store the complete original response
                                        'full_response': full_response_json,
                                        # Store all individual fields as they come from API
                                        'raw_status': item.get('status', ''),
                                        'raw_message_id': item.get('message_id', ''),
                                        'raw_phone': item.get('phone', ''),
                                        'raw_isoCode': item.get('isoCode', ''),
                                        'raw_service': item.get('service', ''),
                                        'raw_message': item.get('message', ''),
                                        'raw_sender': item.get('sender', ''),
                                        'raw_template_id': item.get('template_id', ''),
                                        'raw_flash': item.get('flash', ''),
                                        'raw_unicode': item.get('unicode', ''),
                                        'raw_length': item.get('length', ''),
                                        'raw_units': item.get('units', ''),
                                        'raw_charges': item.get('charges', ''),
                                        'raw_source': item.get('source', ''),
                                        'raw_submit_time': item.get('submit_time', ''),
                                        'raw_submit_at': item.get('submit_at', ''),
                                        'raw_deliv_time': item.get('deliv_time', ''),
                                        'raw_deliv_at': item.get('deliv_at', ''),
                                        # Additional metadata from response wrapper
                                        'api_message': data.get('message', ''),
                                        'api_page': data.get('page', ''),
                                        'api_length': data.get('length', '')
                                    }
                                    self.log_message(f"✅ Raw status for {message_id}: {item.get('status', 'UNKNOWN')}")
                                    break
                        else:
                            # Store unexpected response format as is
                            results[message_id] = {
                                'full_response': full_response_json,
                                'raw_status': 'unknown_format',
                                'error_message': 'Unexpected response format'
                            }
                            
                    except json.JSONDecodeError as e:
                        self.log_message(f"❌ JSON decode error for {message_id}: {str(e)}")
                        results[message_id] = {
                            'full_response': resp.text,
                            'raw_status': 'json_error',
                            'error_message': f'Invalid JSON: {str(e)}'
                        }
                        
                elif resp.status_code == 401:
                    self.log_message(f"⚠️ SMS Status API Authentication Failed for {message_id}")
                    results[message_id] = {
                        'full_response': resp.text,
                        'raw_status': 'auth_failed',
                        'error_message': 'Authentication failed'
                    }
                else:
                    self.log_message(f"⚠️ SMS Status API returned {resp.status_code} for {message_id}")
                    results[message_id] = {
                        'full_response': resp.text,
                        'raw_status': f'http_{resp.status_code}',
                        'error_message': resp.text[:200]
                    }
                    
            except requests.exceptions.Timeout:
                self.log_message(f"⚠️ SMS Status API timeout for {message_id}")
                results[message_id] = {
                    'full_response': '',
                    'raw_status': 'timeout',
                    'error_message': 'Request timeout'
                }
            except Exception as e:
                self.log_message(f"⚠️ SMS Status API error for {message_id}: {str(e)}")
                results[message_id] = {
                    'full_response': '',
                    'raw_status': 'error',
                    'error_message': str(e)
                }
            
            # Small delay to avoid rate limiting
            time.sleep(0.1)
        
        return results

    def process_and_send_sms(self, excel_file):
        """Enhanced SMS sending with raw API response capture and email report"""
        try:
            df = pd.read_excel(excel_file, dtype=str).fillna("")
            
            required = ['Employee Name', 'Employee no', 'UAN', 'Drive Link']
            for col in required:
                if col not in df.columns:
                    st.error(f"❌ Missing column: {col}")
                    return
            
            today = datetime.now()
            first_day_of_current_month = today.replace(day=1)
            last_month = first_day_of_current_month - timedelta(days=1)
            prev_month_str = last_month.strftime("%B %Y")
            
            total = len(df)
            progress = st.progress(0)
            status = st.empty()
            
            sent_count = 0
            failed_count = 0
            skipped_count = 0
            results = []
            
            # Store message_ids for later status lookup
            message_ids_list = []
            
            # First pass: Send SMS and collect message_ids
            for i, row in df.iterrows():
                name = row['Employee Name']
                emp_no_raw = str(row['Employee no']).strip()
                drive_link = str(row['Drive Link']).strip()
                
                status.info(f"⏳ Sending SMS {i+1}/{total}: {name}")
                
                if not drive_link:
                    skipped_count += 1
                    results.append({
                        'row': i+1,
                        'name': name,
                        'phone': emp_no_raw,
                        'send_status': 'skipped_no_link',
                        'send_response': 'No Drive Link available',
                        'message_id': '',
                        # Status API fields (empty for skipped)
                        'api_raw_status': '',
                        'api_full_response': '',
                        'api_phone': '',
                        'api_isoCode': '',
                        'api_service': '',
                        'api_message_preview': '',
                        'api_sender': '',
                        'api_template_id': '',
                        'api_flash': '',
                        'api_unicode': '',
                        'api_length': '',
                        'api_units': '',
                        'api_charges': '',
                        'api_source': '',
                        'api_submit_time': '',
                        'api_submit_at': '',
                        'api_deliv_time': '',
                        'api_deliv_at': '',
                        'api_response_message': '',
                        'api_page': '',
                        'api_length_count': ''
                    })
                    self.log_message(f"⭐ Skipped {name} - No Drive Link")
                    progress.progress((i+1)/total)
                    continue
                
                if not emp_no_raw:
                    skipped_count += 1
                    results.append({
                        'row': i+1,
                        'name': name,
                        'phone': '',
                        'send_status': 'skipped_no_phone',
                        'send_response': 'No phone number',
                        'message_id': '',
                        # Status API fields (empty for skipped)
                        'api_raw_status': '',
                        'api_full_response': '',
                        'api_phone': '',
                        'api_isoCode': '',
                        'api_service': '',
                        'api_message_preview': '',
                        'api_sender': '',
                        'api_template_id': '',
                        'api_flash': '',
                        'api_unicode': '',
                        'api_length': '',
                        'api_units': '',
                        'api_charges': '',
                        'api_source': '',
                        'api_submit_time': '',
                        'api_submit_at': '',
                        'api_deliv_time': '',
                        'api_deliv_at': '',
                        'api_response_message': '',
                        'api_page': '',
                        'api_length_count': ''
                    })
                    self.log_message(f"⭐ Skipped {name} - No phone")
                    progress.progress((i+1)/total)
                    continue
                
                phone = self.format_phone_number(emp_no_raw)
                success, resp_text, message_id = self.send_sms_via_qik(name, phone, drive_link, prev_month_str)
                
                if success:
                    sent_count += 1
                    result_entry = {
                        'row': i+1,
                        'name': name,
                        'phone': phone,
                        'send_status': 'sent',
                        'send_response': resp_text[:500] if len(resp_text) > 500 else resp_text,
                        'message_id': message_id if message_id else '',
                        # Status API fields (to be filled later)
                        'api_raw_status': '',
                        'api_full_response': '',
                        'api_phone': '',
                        'api_isoCode': '',
                        'api_service': '',
                        'api_message_preview': '',
                        'api_sender': '',
                        'api_template_id': '',
                        'api_flash': '',
                        'api_unicode': '',
                        'api_length': '',
                        'api_units': '',
                        'api_charges': '',
                        'api_source': '',
                        'api_submit_time': '',
                        'api_submit_at': '',
                        'api_deliv_time': '',
                        'api_deliv_at': '',
                        'api_response_message': '',
                        'api_page': '',
                        'api_length_count': ''
                    }
                    results.append(result_entry)
                    
                    if message_id:
                        message_ids_list.append(message_id)
                        self.log_message(f"✅ SMS sent: {name} ({phone}) - Message ID: {message_id}")
                    else:
                        self.log_message(f"✅ SMS sent: {name} ({phone}) - No message_id in response")
                else:
                    failed_count += 1
                    results.append({
                        'row': i+1,
                        'name': name,
                        'phone': phone,
                        'send_status': 'failed',
                        'send_response': resp_text[:500] if len(resp_text) > 500 else resp_text,
                        'message_id': '',
                        # Status API fields (empty for failed)
                        'api_raw_status': '',
                        'api_full_response': '',
                        'api_phone': '',
                        'api_isoCode': '',
                        'api_service': '',
                        'api_message_preview': '',
                        'api_sender': '',
                        'api_template_id': '',
                        'api_flash': '',
                        'api_unicode': '',
                        'api_length': '',
                        'api_units': '',
                        'api_charges': '',
                        'api_source': '',
                        'api_submit_time': '',
                        'api_submit_at': '',
                        'api_deliv_time': '',
                        'api_deliv_at': '',
                        'api_response_message': '',
                        'api_page': '',
                        'api_length_count': ''
                    })
                    self.log_message(f"❌ SMS failed: {name} - {resp_text[:100]}")
                
                time.sleep(0.25)  # Rate limiting
                progress.progress((i+1)/total)
            
            # ========== ADDED DELAY BEFORE STATUS CHECK ==========
            if message_ids_list:
                self.log_message(f"⏳ Waiting 10 seconds before checking delivery status for {len(message_ids_list)} messages...")
                status.info(f"⏳ Waiting 10 seconds before checking delivery status...")
                time.sleep(10)  # Add delay to allow API to update status to DELIVERED
            # ====================================================
            
            # Second pass: Fetch delivery status for all sent messages that have message_ids
            if message_ids_list:
                status.info(f"📊 Fetching delivery status for {len(message_ids_list)} messages...")
                
                # Get status for all message_ids
                status_results = self.get_sms_status(message_ids_list)
                
                # Update results with raw API data
                for result in results:
                    if result.get('message_id') and result['message_id'] in status_results:
                        api_data = status_results[result['message_id']]
                        
                        # Populate all fields with raw API data
                        result['api_raw_status'] = api_data.get('raw_status', '')
                        result['api_full_response'] = api_data.get('full_response', '')
                        result['api_phone'] = api_data.get('raw_phone', '')
                        result['api_isoCode'] = api_data.get('raw_isoCode', '')
                        result['api_service'] = api_data.get('raw_service', '')
                        result['api_message_preview'] = api_data.get('raw_message', '')[:200] if api_data.get('raw_message') else ''
                        result['api_sender'] = api_data.get('raw_sender', '')
                        result['api_template_id'] = api_data.get('raw_template_id', '')
                        result['api_flash'] = api_data.get('raw_flash', '')
                        result['api_unicode'] = api_data.get('raw_unicode', '')
                        result['api_length'] = api_data.get('raw_length', '')
                        result['api_units'] = api_data.get('raw_units', '')
                        result['api_charges'] = api_data.get('raw_charges', '')
                        result['api_source'] = api_data.get('raw_source', '')
                        result['api_submit_time'] = api_data.get('raw_submit_time', '')
                        result['api_submit_at'] = api_data.get('raw_submit_at', '')
                        result['api_deliv_time'] = api_data.get('raw_deliv_time', '')
                        result['api_deliv_at'] = api_data.get('raw_deliv_at', '')
                        result['api_response_message'] = api_data.get('api_message', '')
                        result['api_page'] = api_data.get('api_page', '')
                        result['api_length_count'] = api_data.get('api_length', '')
                        
                        self.log_message(f"📬 Status update for {result['name']}: Raw Status = {result['api_raw_status']}")
                    elif result.get('message_id'):
                        result['api_raw_status'] = 'status_unavailable'
                        result['api_full_response'] = 'Status API returned no data for this message_id'
                        self.log_message(f"⚠️ No status data for {result['name']} (ID: {result['message_id']})")
            
            progress.empty()
            status.empty()
            
            # Calculate stats from raw status
            delivered_count = sum(1 for r in results if r.get('api_raw_status') == 'DELIVERED')
            processed_count = sum(1 for r in results if r.get('api_raw_status') == 'PROCESSED')
            sent_count_actual = sum(1 for r in results if r.get('send_status') == 'sent')
            
            # Display Results
            st.markdown('<div class="section-header">📊 SMS Distribution Results</div>', unsafe_allow_html=True)
            
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-icon">✅</div>
                    <div class="metric-label">Sent</div>
                    <div class="metric-value" style="color: #28a745;">{sent_count_actual}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-icon">📬</div>
                    <div class="metric-label">DELIVERED</div>
                    <div class="metric-value" style="color: #667eea;">{delivered_count}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-icon">⚙️</div>
                    <div class="metric-label">PROCESSED</div>
                    <div class="metric-value" style="color: #ffc107;">{processed_count}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                other_status = sent_count_actual - delivered_count - processed_count
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-icon">⏳</div>
                    <div class="metric-label">Other Status</div>
                    <div class="metric-value" style="color: #ff9800;">{other_status}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col5:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-icon">❌</div>
                    <div class="metric-label">Failed/Skipped</div>
                    <div class="metric-value" style="color: #dc3545;">{failed_count + skipped_count}</div>
                </div>
                """, unsafe_allow_html=True)
            
            # Create detailed report DataFrame with all raw fields
            report_df = pd.DataFrame(results)
            
            # Define column order with all raw API fields
            column_order = [
                'row', 'name', 'phone', 
                'send_status', 'send_response', 'message_id',
                'api_raw_status', 'api_submit_time', 'api_deliv_time',
                'api_units', 'api_charges', 'api_phone', 'api_sender',
                'api_message_preview', 'api_full_response'
            ]
            
            # Add all other api_ columns that might exist
            api_columns = [col for col in report_df.columns if col.startswith('api_') and col not in column_order]
            column_order.extend(api_columns)
            
            # Ensure all columns exist
            column_order = [c for c in column_order if c in report_df.columns]
            report_df = report_df[column_order]
            
            # Display summary table
            st.markdown("**📋 Detailed Report Preview**")
            st.dataframe(report_df.head(10), use_container_width=True)
            
            # Prepare report for email
            report_filename = f"sms_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            excel_buf = BytesIO()
            with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
                # Main report sheet with all data
                report_df.to_excel(writer, index=False, sheet_name='SMS Report')
                
                # Status summary sheet
                summary_data = {
                    'Metric': [
                        'Total SMS Attempted', 
                        'Successfully Sent',
                        'API Status - DELIVERED',
                        'API Status - PROCESSED',
                        'API Status - Other',
                        'Failed to Send',
                        'Skipped (No Drive Link)', 
                        'Skipped (No Phone Number)'
                    ],
                    'Count': [
                        sent_count + failed_count,
                        sent_count,
                        delivered_count,
                        processed_count,
                        sent_count - delivered_count - processed_count,
                        failed_count,
                        sum(1 for r in results if r.get('send_status') == 'skipped_no_link'),
                        sum(1 for r in results if r.get('send_status') == 'skipped_no_phone')
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, index=False, sheet_name='Summary')
                
                # Raw API responses sheet (only if there's data)
                api_responses = report_df[['row', 'name', 'message_id', 'api_raw_status', 'api_full_response']].copy()
                api_responses = api_responses[api_responses['api_full_response'].str.len() > 0]
                if not api_responses.empty:
                    api_responses.to_excel(writer, index=False, sheet_name='Raw API Responses')
            
            excel_buf.seek(0)
            
            # Store report data for email sending
            st.session_state.last_sms_report_data = {
                'filename': report_filename,
                'buffer': excel_buf,
                'summary': {
                    'total_attempted': sent_count + failed_count,
                    'sent_count': sent_count,
                    'delivered_count': delivered_count,
                    'processed_count': processed_count,
                    'failed_count': failed_count,
                    'skipped_no_link': sum(1 for r in results if r.get('send_status') == 'skipped_no_link'),
                    'skipped_no_phone': sum(1 for r in results if r.get('send_status') == 'skipped_no_phone')
                }
            }
            
            # Download buttons
            st.markdown("---")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # CSV Download (without full API response to keep it manageable)
                csv_df = report_df.copy()
                if 'api_full_response' in csv_df.columns:
                    csv_df['api_full_response'] = csv_df['api_full_response'].str[:200]
                csv_buf = csv_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "📥 Download SMS Report (CSV)",
                    data=csv_buf,
                    file_name=f"sms_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    type="primary"
                )
            
            with col2:
                # Excel Download with full data
                excel_buf_download = BytesIO(excel_buf.getvalue())
                st.download_button(
                    "📊 Download SMS Report (Excel)",
                    data=excel_buf_download,
                    file_name=report_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="secondary"
                )
            
            with col3:
                # Email report button
                if st.button("📧 Send Report via Email", use_container_width=True, type="secondary"):
                    with st.spinner("📧 Sending email report..."):
                        success, message = self.send_email_report(
                            st.session_state.last_sms_report_data['summary'],
                            report_filename,
                            BytesIO(excel_buf.getvalue())
                        )
                        if success:
                            st.success(f"✅ {message}")
                        else:
                            st.error(f"❌ {message}")
            
            # Show success message with raw status counts
            st.success(f"✅ SMS processing complete! Sent: {sent_count_actual} | Raw Status - DELIVERED: {delivered_count}, PROCESSED: {processed_count}, Other: {sent_count_actual - delivered_count - processed_count}")
            
            self.log_message(f"📊 SMS Summary - Sent: {sent_count_actual}, DELIVERED: {delivered_count}, PROCESSED: {processed_count}, Failed: {failed_count}, Skipped: {skipped_count}")
            
            # Auto-send email if configured
            if EMAIL_SENDER and EMAIL_PASSWORD:
                with st.spinner("📧 Auto-sending report via email..."):
                    success, message = self.send_email_report(
                        st.session_state.last_sms_report_data['summary'],
                        report_filename,
                        BytesIO(excel_buf.getvalue())
                    )
                    if success:
                        st.info(f"📧 Report auto-sent to {REPORT_RECIPIENT}")
                        self.log_message(f"📧 Auto-email sent to {REPORT_RECIPIENT}")
                    else:
                        self.log_message(f"⚠️ Auto-email failed: {message}")
            
        except Exception as e:
            st.error(f"❌ SMS processing error: {str(e)}")
            self.log_message(f"❌ SMS error: {str(e)}")
            
    def reset_session(self):
        st.session_state.files_processed = False
        st.session_state.processing_complete = False
        st.session_state.results = {}
        st.session_state.updated_excel_buffer = None
        st.session_state.uan_pages = {}
        st.session_state.df = None
        st.session_state.pdf_count = 0
        st.session_state.excel_count = 0
        st.session_state.force_rerun = True
        st.success("✅ Session reset successfully!")
        self.log_message("🔄 Session reset")

    def log_message(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        st.session_state.log_entries.append(log_entry)
        
        if len(st.session_state.log_entries) > 500:
            st.session_state.log_entries = st.session_state.log_entries[-500:]

if __name__ == "__main__":
    app = PayslipDistributorStreamlit()
    app.run()
