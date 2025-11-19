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

# Load environment variables
load_dotenv()

QIK_URL = os.getenv("QIK_URL")
QIK_AUTH_TOKEN = os.getenv("QIK_AUTH_TOKEN")
QIK_SENDER = os.getenv("QIK_SENDER")
QIK_TEMPLATE_ID = os.getenv("QIK_TEMPLATE_ID")
QIK_SERVICE = os.getenv("QIK_SERVICE")
QIK_SHORTEN_URL = os.getenv("QIK_SHORTEN_URL")

class PayslipDistributorStreamlit:
    def __init__(self):
        self.shared_drive_id = os.getenv("SHARED_DRIVE_ID")
        self.initialize_session_state()
        self.setup_page()

    def initialize_session_state(self):
        """Initialize session state variables"""
        defaults = {
            'sent_numbers': {},
            'log_entries': [],
            'current_folder': None,
            'folder_stack': [],
            'drive_service': None,
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
            'show_delete_confirm': False,
            'drive_page': 0,  # Add pagination state
            'items_per_page': 10  # Add items per page
        }
        for key, value in defaults.items():
            if key not in st.session_state:
                st.session_state[key] = value

    def setup_page(self):
        st.set_page_config(
            page_title="Payslip Distribution System",
            page_icon="📄",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        
        # Enhanced modern CSS styling
        st.markdown("""
        <style>
        /* Import Google Fonts */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        
        /* Global Styles */
        .main {
            font-family: 'Inter', sans-serif;
        }
        
        /* Header Styles */
        .main-header {
            font-size: 3rem;
            font-weight: 700;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            text-align: center;
            margin-bottom: 0.5rem;
            padding: 1rem 0;
        }
        
        .sub-header {
            font-size: 1.1rem;
            color: #6b7280;
            text-align: center;
            margin-bottom: 2.5rem;
            font-weight: 500;
        }
        
        .section-header {
            font-size: 1.4rem;
            font-weight: 600;
            color: #1f2937;
            margin-top: 2rem;
            margin-bottom: 1.5rem;
            padding: 0.75rem 1rem;
            background: linear-gradient(135deg, #667eea15 0%, #764ba215 100%);
            border-left: 4px solid #667eea;
            border-radius: 8px;
        }
        
        /* Card Styles */
        .metric-card {
            background: white;
            padding: 1.5rem;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.07);
            border: 1px solid #e5e7eb;
            transition: transform 0.2s, box-shadow 0.2s;
        }
        
        .metric-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 12px rgba(0,0,0,0.12);
        }
        
        .metric-value {
            font-size: 2.5rem;
            font-weight: 700;
            color: #667eea;
            margin: 0.5rem 0;
        }
        
        .metric-label {
            font-size: 0.9rem;
            color: #6b7280;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        /* Status Boxes */
        .success-box {
            background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
            border: 2px solid #28a745;
            border-radius: 12px;
            padding: 1.25rem;
            margin: 1rem 0;
            box-shadow: 0 4px 6px rgba(40, 167, 69, 0.1);
        }
        
        .error-box {
            background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%);
            border: 2px solid #dc3545;
            border-radius: 12px;
            padding: 1.25rem;
            margin: 1rem 0;
            box-shadow: 0 4px 6px rgba(220, 53, 69, 0.1);
        }
        
        .info-box {
            background: linear-gradient(135deg, #d1ecf1 0%, #bee5eb 100%);
            border: 2px solid #17a2b8;
            border-radius: 12px;
            padding: 1.25rem;
            margin: 1rem 0;
            box-shadow: 0 4px 6px rgba(23, 162, 184, 0.1);
        }
        
        .warning-box {
            background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%);
            border: 2px solid #ffc107;
            border-radius: 12px;
            padding: 1.25rem;
            margin: 1rem 0;
            box-shadow: 0 4px 6px rgba(255, 193, 7, 0.1);
        }
        
        /* Log Box */
        .log-box {
            background: #1e293b;
            color: #e2e8f0;
            border-radius: 12px;
            padding: 1.5rem;
            margin: 1rem 0;
            max-height: 400px;
            overflow-y: auto;
            font-family: 'Courier New', monospace;
            font-size: 0.85rem;
            box-shadow: 0 8px 16px rgba(0,0,0,0.2);
        }
        
        .log-box::-webkit-scrollbar {
            width: 8px;
        }
        
        .log-box::-webkit-scrollbar-track {
            background: #0f172a;
            border-radius: 4px;
        }
        
        .log-box::-webkit-scrollbar-thumb {
            background: #475569;
            border-radius: 4px;
        }
        
        /* Drive Item Styles */
        .drive-item {
            background: white;
            padding: 1rem;
            margin: 0.5rem 0;
            border-radius: 8px;
            border: 1px solid #e5e7eb;
            transition: all 0.2s;
        }
        
        .drive-item:hover {
            background: #f9fafb;
            border-color: #667eea;
            box-shadow: 0 2px 4px rgba(102, 126, 234, 0.1);
        }
        
        .drive-folder {
            background: linear-gradient(135deg, #667eea15 0%, #764ba215 100%);
        }
        
        /* Button Enhancements */
        .stButton > button {
            border-radius: 8px;
            font-weight: 500;
            transition: all 0.2s;
            border: none;
        }
        
        .stButton > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        }
        
        /* Status Badge */
        .status-badge {
            display: inline-block;
            padding: 0.35rem 0.85rem;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .status-connected {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        
        .status-disconnected {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        
        /* Progress Bar Enhancement */
        .stProgress > div > div > div {
            background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        }
        
        /* Tab Styling */
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
        }
        
        .stTabs [data-baseweb="tab"] {
            border-radius: 8px 8px 0 0;
            font-weight: 500;
        }
        
        /* File uploader enhancement */
        .uploadedFile {
            border-radius: 8px;
        }
        
        /* Breadcrumb */
        .breadcrumb {
            background: #f3f4f6;
            padding: 0.75rem 1rem;
            border-radius: 8px;
            margin-bottom: 1rem;
            font-size: 0.9rem;
            color: #4b5563;
        }
        
        .breadcrumb-item {
            display: inline-block;
            margin: 0 0.5rem;
        }
        
        /* Animation */
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        .animated {
            animation: fadeIn 0.5s ease-out;
        }
        
        /* Delete button styling */
        .delete-btn {
            background: #dc3545 !important;
            color: white !important;
        }
        
        .delete-btn:hover {
            background: #c82333 !important;
        }
        </style>
        """, unsafe_allow_html=True)

    def run(self):
        # Animated header
        st.markdown('<div class="main-header animated">📄 Payslip Distribution System</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header animated">Fast, Secure & Efficient Employee Payslip Distribution via Google Drive</div>', unsafe_allow_html=True)

        # Tabs with icons
        tab1, tab2, tab3, tab4 = st.tabs([
            "📤 Send Payslips", 
            "📁 Drive Explorer", 
            "📨 Send SMS", 
            "⚙️ Settings"
        ])

        with tab1:
            self.render_send_tab()
        with tab2:
            self.render_drive_tab()
        with tab3:
            self.render_sms_tab()
        with tab4:
            self.render_settings_tab()

        # Enhanced activity log
        st.markdown('<div class="section-header">📜 Activity Log</div>', unsafe_allow_html=True)
        if st.session_state.log_entries:
            log_html = '<div class="log-box">'
            for entry in reversed(st.session_state.log_entries[-100:]):
                log_html += f'<div>{entry}</div>'
            log_html += '</div>'
            st.markdown(log_html, unsafe_allow_html=True)
        else:
            st.info("💤 No activity yet. Start by initializing Google Drive connection.")

    def render_send_tab(self):
        st.markdown('<div class="section-header">🔐 Google Drive Connection</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([2, 2, 1])
        with col1:
            if st.button("🚀 Initialize Google Drive Connection", type="primary", use_container_width=True):
                self.initialize_drive()
        with col2:
            if st.session_state.drive_initialized:
                st.markdown('<span class="status-badge status-connected">✅ Connected</span>', unsafe_allow_html=True)
            else:
                st.markdown('<span class="status-badge status-disconnected">❌ Not Connected</span>', unsafe_allow_html=True)
        with col3:
            if st.session_state.drive_initialized:
                if st.button("🔄 Reconnect", use_container_width=True):
                    self.initialize_drive()

        if not st.session_state.drive_initialized:
            st.warning("⚠️ Please initialize Google Drive connection before proceeding.")
            return

        # File upload section with enhanced UI
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

        # Display analysis with metrics
        if st.session_state.files_processed:
            st.markdown('<div class="section-header">📊 File Analysis Results</div>', unsafe_allow_html=True)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">📄 UANs Found</div>
                    <div class="metric-value">{st.session_state.pdf_count}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">👥 Employees</div>
                    <div class="metric-value">{st.session_state.excel_count}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                match_rate = (min(st.session_state.pdf_count, st.session_state.excel_count) / 
                             max(st.session_state.excel_count, 1) * 100)
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">🎯 Match Rate</div>
                    <div class="metric-value">{match_rate:.0f}%</div>
                </div>
                """, unsafe_allow_html=True)

            # Process payslips section
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

        # Results display with enhanced metrics
        if st.session_state.processing_complete:
            st.markdown('<div class="section-header">📊 Upload Results</div>', unsafe_allow_html=True)
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">✅ Uploaded</div>
                    <div class="metric-value" style="color: #28a745;">{st.session_state.results.get('uploaded', 0)}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">❌ Failed</div>
                    <div class="metric-value" style="color: #dc3545;">{st.session_state.results.get('failed', 0)}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">⏭️ Skipped</div>
                    <div class="metric-value" style="color: #ffc107;">{st.session_state.results.get('skipped', 0)}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                total = sum([st.session_state.results.get(k, 0) for k in ['uploaded', 'failed', 'skipped']])
                success_rate = (st.session_state.results.get('uploaded', 0) / max(total, 1) * 100)
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">📈 Success Rate</div>
                    <div class="metric-value" style="color: #667eea;">{success_rate:.0f}%</div>
                </div>
                """, unsafe_allow_html=True)

            if st.session_state.updated_excel_buffer:
                st.markdown("---")
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.download_button(
                        label="📥 Download Updated Excel File with Drive Links",
                        data=st.session_state.updated_excel_buffer,
                        file_name=f"payslip_links_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )

        # Reset section
        st.markdown("---")
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            if st.button("🔄 Reset Session", type="secondary", use_container_width=True):
                self.reset_session()

    def render_drive_tab(self):
        st.markdown('<div class="section-header">📁 Google Drive Explorer</div>', unsafe_allow_html=True)
        
        if not st.session_state.drive_initialized:
            st.warning("⚠️ Please initialize Google Drive connection in the 'Send Payslips' tab first.")
            return

        # Breadcrumb navigation
        breadcrumb = "📁 Root"
        if st.session_state.current_folder:
            breadcrumb += " > " + st.session_state.current_path
        st.markdown(f'<div class="breadcrumb">{breadcrumb}</div>', unsafe_allow_html=True)

        # Navigation controls with enhanced styling
        col1, col2, col3, col4, col5 = st.columns([1, 1, 1, 1, 2])
        with col1:
            if st.button("🏠 Root", use_container_width=True):
                self.go_to_root()
        with col2:
            if st.button("⬅️ Back", use_container_width=True, disabled=not st.session_state.folder_stack):
                self.go_back()
        with col3:
            if st.button("🔄 Refresh", use_container_width=True):
                self.refresh_drive()
        with col4:
            if st.button("➕ New Folder", use_container_width=True):
                st.session_state.show_new_folder = True
        with col5:
            st.empty()

        # New folder creation
        if st.session_state.get('show_new_folder', False):
            with st.form("new_folder_form"):
                folder_name = st.text_input("Folder Name")
                col1, col2 = st.columns(2)
                with col1:
                    if st.form_submit_button("Create", use_container_width=True):
                        if folder_name:
                            self.create_folder(folder_name)
                            st.session_state.show_new_folder = False
                            st.rerun()
                with col2:
                    if st.form_submit_button("Cancel", use_container_width=True):
                        st.session_state.show_new_folder = False
                        st.rerun()

        st.markdown("---")
        
        # Display drive contents with delete functionality
        self.display_drive_contents()

    def render_sms_tab(self):
        st.markdown('<div class="section-header">📨 SMS Distribution</div>', unsafe_allow_html=True)
        
        st.info("📋 Upload the updated Excel file (with Drive Links) from the Send Payslips tab to send SMS notifications to employees.")
        
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
            
            # Preview data
            try:
                preview_df = pd.read_excel(uploaded, dtype=str).fillna("")
                st.markdown("**📋 Data Preview (First 5 rows)**")
                st.dataframe(preview_df.head(), use_container_width=True)
                
                # Check for Drive Links
                has_links = preview_df['Drive Link'].notna().sum() if 'Drive Link' in preview_df.columns else 0
                st.info(f"📊 {has_links} employees have Drive Links and will receive SMS")
                
            except Exception as e:
                st.error(f"Error previewing file: {str(e)}")

        st.markdown("---")
        
        col1, col2 = st.columns([2, 1])
        with col1:
            if st.button("📤 Send SMS to All Employees with Drive Links", 
                        type="primary", 
                        use_container_width=True,
                        disabled=not uploaded):
                if uploaded:
                    self.process_and_send_sms(uploaded)

    def render_settings_tab(self):
        st.markdown('<div class="section-header">⚙️ System Configuration</div>', unsafe_allow_html=True)
        
        # Drive settings
        st.markdown("**🔧 Google Drive Settings**")
        drive_id = st.text_input(
            "Shared Drive ID",
            value=st.session_state.shared_drive_id or self.shared_drive_id,
            help="Enter your Google Shared Drive ID",
            key="drive_id_input"
        )
        
        if st.button("💾 Save Drive Settings", type="primary", use_container_width=True):
            st.session_state.shared_drive_id = drive_id
            st.success("✅ Settings saved successfully!")
            self.log_message("✅ Saved shared drive ID configuration")

        st.markdown("---")
        
        # System status
        st.markdown('<div class="section-header">📊 System Status</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            status = "Connected" if st.session_state.drive_initialized else "Disconnected"
            status_color = "#28a745" if st.session_state.drive_initialized else "#dc3545"
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Google Drive</div>
                <div class="metric-value" style="color: {status_color}; font-size: 1.5rem;">{status}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Active Sessions</div>
                <div class="metric-value" style="font-size: 1.5rem;">{len(st.session_state.sent_numbers)}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Log Entries</div>
                <div class="metric-value" style="font-size: 1.5rem;">{len(st.session_state.log_entries)}</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        
        # SMS Configuration
        st.markdown("**📨 SMS Service Configuration**")
        col1, col2 = st.columns(2)
        with col1:
            qik_status = "Configured" if QIK_URL and QIK_AUTH_TOKEN else "Not Configured"
            st.info(f"Status: {qik_status}")
        with col2:
            if QIK_SENDER:
                st.info(f"Sender: {QIK_SENDER}")

        st.markdown("---")
        
        # Danger zone
        st.markdown('<div class="section-header">⚠️ Danger Zone</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🗑️ Clear Session Data", type="secondary", use_container_width=True):
                self.reset_session()
                st.success("✅ Session data cleared")
        
        with col2:
            if st.button("💣 Clear All Data", type="secondary", use_container_width=True):
                self.clear_session_data()

    # Core functionality methods
    def initialize_drive(self):
        """Initialize Google Drive connection"""
        with st.spinner("🔄 Connecting to Google Drive..."):
            try:
                service = self.get_drive_service()
                if service:
                    # Test the connection
                    service.files().list(pageSize=1, supportsAllDrives=True).execute()
                    st.session_state.drive_service = service
                    st.session_state.drive_initialized = True
                    st.success("✅ Google Drive connected successfully!")
                    self.log_message("✅ Google Drive connection initialized")
                else:
                    st.session_state.drive_initialized = False
                    st.error("❌ Failed to connect to Google Drive")
                    self.log_message("❌ Failed to initialize Google Drive")
            except Exception as e:
                st.session_state.drive_initialized = False
                st.error(f"❌ Connection error: {str(e)}")
                self.log_message(f"❌ Drive connection error: {str(e)}")

    def process_files(self, pdf_file, excel_file):
        """Process uploaded PDF and Excel files"""
        progress_container = st.empty()
        
        with progress_container:
            with st.spinner("🔄 Processing files..."):
                try:
                    # Process PDF
                    st.info("📄 Extracting UAN information from PDF...")
                    uan_pages = self.process_pdf(pdf_file)
                    st.session_state.pdf_count = len(uan_pages)
                    
                    # Process Excel
                    st.info("📊 Loading employee data from Excel...")
                    df = pd.read_excel(excel_file, dtype=str).fillna("")
                    
                    # Column handling
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
                    st.session_state.pdf_file = pdf_file
                    st.session_state.excel_file = excel_file
                    
                    progress_container.empty()
                    st.success("✅ Files processed successfully!")
                    self.log_message(f"📊 PDF: {st.session_state.pdf_count} UANs | Excel: {st.session_state.excel_count} employees")
                    time.sleep(1)
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ Error processing files: {str(e)}")
                    self.log_message(f"❌ Processing error: {str(e)}")

    def process_payslips(self):
        """Extract and upload individual payslips to Drive"""
        if not st.session_state.files_processed or st.session_state.df is None:
            st.error("❌ Please process files first")
            return
        
        if not st.session_state.drive_initialized:
            st.error("❌ Google Drive not connected")
            return

        # Get previous month
        today = datetime.now()
        first_day_of_current_month = today.replace(day=1)
        last_month = first_day_of_current_month - timedelta(days=1)
        current_month = last_month.strftime("%B %Y")

        progress_placeholder = st.empty()
        status_placeholder = st.empty()
        
        with st.spinner(f"🚀 Uploading payslips for {current_month}..."):
            try:
                # Create monthly folder
                folder_id = self.get_monthly_folder_id(
                    st.session_state.drive_service,
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
                
                for idx, (_, row) in enumerate(updated_df.iterrows()):
                    uan = str(row.get('UAN', '')).strip()
                    emp_name = str(row.get('Employee Name', '')).strip()
                    
                    # Update progress
                    progress = (idx + 1) / total
                    progress_placeholder.progress(progress)
                    status_placeholder.info(f"⏳ Processing {idx + 1}/{total}: {emp_name} (UAN: {uan})")
                    
                    if not uan:
                        self.log_message(f"⚠️ Skipped row {idx+1}: Missing UAN")
                        results['skipped'] += 1
                        continue
                    
                    if uan in st.session_state.uan_pages:
                        try:
                            # Extract individual payslip
                            st.session_state.pdf_file.seek(0)
                            page_num = st.session_state.uan_pages[uan]
                            pdf_buffer = self.extract_individual_payslip(
                                st.session_state.pdf_file,
                                uan,
                                page_num
                            )
                            
                            # Upload to Drive
                            filename = f"Payslip_{uan}_{current_month.replace(' ', '_')}.pdf"
                            drive_link, file_id = self.upload_to_drive(
                                st.session_state.drive_service,
                                pdf_buffer,
                                filename,
                                folder_id,
                                st.session_state.shared_drive_id
                            )
                            
                            if drive_link:
                                # Update DataFrame
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
                
                # Prepare final Excel
                final_df = updated_df.copy()
                for col in ['Employee Name', 'Employee no', 'UAN', 'Drive Link']:
                    if col not in final_df.columns:
                        final_df[col] = ""
                
                final_df = final_df[['Employee Name', 'Employee no', 'UAN', 'Drive Link']]
                
                # Save to buffer
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
                time.sleep(2)
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ Processing error: {str(e)}")
                self.log_message(f"❌ Processing error: {str(e)}")

    def process_pdf(self, pdf_file):
        """Extract UAN to page mapping from PDF"""
        uan_pages = {}
        try:
            pdf_file.seek(0)
            with pdfplumber.open(pdf_file) as pdf:
                total_pages = len(pdf.pages)
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
        """Extract single page as PDF"""
        pdf_reader = PdfReader(pdf_file)
        pdf_writer = PdfWriter()
        pdf_writer.add_page(pdf_reader.pages[page_num])
        output_buffer = BytesIO()
        pdf_writer.write(output_buffer)
        output_buffer.seek(0)
        return output_buffer

    # Google Drive helper methods
    def get_google_credentials(self):
        """Load Google service account credentials"""
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
            
            for key, value in credentials_dict.items():
                if not value:
                    st.error(f"❌ Missing credential: {key}")
                    return None
            
            return service_account.Credentials.from_service_account_info(
                credentials_dict,
                scopes=['https://www.googleapis.com/auth/drive']
            )
        except Exception as e:
            st.error(f"❌ Credential error: {str(e)}")
            return None

    def get_drive_service(self):
        """Initialize Google Drive service"""
        try:
            credentials = self.get_google_credentials()
            if credentials:
                return build('drive', 'v3', credentials=credentials)
            return None
        except Exception as e:
            st.error(f"❌ Drive service error: {str(e)}")
            return None

    def get_monthly_folder_id(self, service, month_year, shared_drive_id):
        """Get or create monthly folder with retry logic"""
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                # Search for existing folder
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
                
                # Create new folder
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
                    if retry_count < max_retries:
                        self.log_message(f"⚠️ Retry {retry_count}/{max_retries} for folder creation")
                        time.sleep(5)
                        continue
                st.error(f"❌ Drive API error: {error}")
                return None
            except Exception as e:
                st.error(f"❌ Folder creation error: {str(e)}")
                return None
        
        return None

    def upload_to_drive(self, service, file_buffer, filename, folder_id, shared_drive_id):
        """Upload file to Drive with retry logic"""
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                file_buffer.seek(0)
                media = MediaIoBaseUpload(file_buffer, mimetype='application/pdf', resumable=True)
                file_metadata = {
                    'name': filename,
                    'parents': [folder_id]
                }
                
                # Upload file
                file = service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id',
                    supportsAllDrives=True
                ).execute()
                
                # Set public permissions
                permission = {
                    'type': 'anyone',
                    'role': 'reader',
                    'allowFileDiscovery': False
                }
                service.permissions().create(
                    fileId=file['id'],
                    body=permission,
                    supportsAllDrives=True
                ).execute()
                
                file_link = f"https://drive.google.com/uc?export=download&id={file['id']}"
                return file_link, file['id']
                
            except HttpError as error:
                if hasattr(error, 'resp') and error.resp.status == 503:
                    retry_count += 1
                    if retry_count < max_retries:
                        self.log_message(f"⚠️ Retry {retry_count}/{max_retries} for {filename}")
                        time.sleep(5)
                        continue
                st.error(f"❌ Upload error: {error}")
                return None, None
            except Exception as e:
                st.error(f"❌ Upload error: {str(e)}")
                return None, None
        
        return None, None

    def list_drive_contents(self, service, shared_drive_id, folder_id=None):
        """List Drive contents with error handling"""
        try:
            query = "trashed=false"
            
            if folder_id:
                query += f" and '{folder_id}' in parents"
                results = service.files().list(
                    q=query,
                    fields="files(id, name, mimeType, size, createdTime, modifiedTime, webViewLink)",
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True,
                    orderBy="folder,name"
                ).execute()
            else:
                query += f" and '{shared_drive_id}' in parents"
                results = service.files().list(
                    q=query,
                    fields="files(id, name, mimeType, size, createdTime, modifiedTime, webViewLink)",
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True,
                    corpora='drive',
                    driveId=shared_drive_id,
                    orderBy="folder,name"
                ).execute()
            
            return results.get('files', [])
        except Exception as e:
            st.error(f"❌ Error listing contents: {str(e)}")
            return []

    def display_drive_contents(self):
        """Display Drive contents with delete functionality and pagination"""
        try:
            contents = self.list_drive_contents(
                st.session_state.drive_service,
                st.session_state.shared_drive_id,
                st.session_state.current_folder
            )
            
            if not contents:
                st.info("📭 No files or folders found in this location.")
                return
            
            # Calculate pagination
            items_per_page = st.session_state.items_per_page
            total_items = len(contents)
            total_pages = (total_items + items_per_page - 1) // items_per_page
            current_page = st.session_state.drive_page
            
            # Ensure current page is within bounds
            if current_page >= total_pages:
                st.session_state.drive_page = max(0, total_pages - 1)
                current_page = st.session_state.drive_page
            
            # Get items for current page
            start_idx = current_page * items_per_page
            end_idx = min(start_idx + items_per_page, total_items)
            page_contents = contents[start_idx:end_idx]
            
            # Display pagination info and controls at top
            col1, col2, col3, col4, col5 = st.columns([2, 1, 2, 1, 2])
            
            with col1:
                st.markdown(f"**📊 Showing {start_idx + 1}-{end_idx} of {total_items} items**")
            
            with col2:
                if st.button("⏮️ First", disabled=current_page == 0, use_container_width=True):
                    st.session_state.drive_page = 0
                    st.rerun()
            
            with col3:
                if st.button("◀️ Previous", disabled=current_page == 0, use_container_width=True):
                    st.session_state.drive_page = max(0, current_page - 1)
                    st.rerun()
            
            with col4:
                if st.button("Next ▶️", disabled=current_page >= total_pages - 1, use_container_width=True):
                    st.session_state.drive_page = min(total_pages - 1, current_page + 1)
                    st.rerun()
            
            with col5:
                if st.button("Last ⏭️", disabled=current_page >= total_pages - 1, use_container_width=True):
                    st.session_state.drive_page = total_pages - 1
                    st.rerun()
            
            # Page number display
            st.markdown(f"<div style='text-align: center; margin: 10px 0;'><strong>Page {current_page + 1} of {total_pages}</strong></div>", unsafe_allow_html=True)
            
            st.markdown("---")
            
            # Display header
            col1, col2, col3, col4, col5, col6 = st.columns([4, 1, 1, 2, 1.5, 1.5])
            with col1:
                st.markdown("**📝 Name**")
            with col2:
                st.markdown("**📂 Type**")
            with col3:
                st.markdown("**📊 Size**")
            with col4:
                st.markdown("**📅 Modified**")
            with col5:
                st.markdown("**⚡ Actions**")
            with col6:
                st.markdown("**🗑️ Delete**")
            
            st.markdown("---")
            
            # Display each item on current page
            for item in page_contents:
                is_folder = item['mimeType'] == 'application/vnd.google-apps.folder'
                item_type = "📁" if is_folder else "📄"
                size = f"{int(item.get('size', 0)) / 1024:.1f} KB" if 'size' in item else "-"
                modified = item.get('modifiedTime', '')[:10] if 'modifiedTime' in item else "-"
                
                col1, col2, col3, col4, col5, col6 = st.columns([4, 1, 1, 2, 1.5, 1.5])
                
                with col1:
                    st.text(item['name'][:40] + "..." if len(item['name']) > 40 else item['name'])
                with col2:
                    st.text(item_type)
                with col3:
                    st.text(size)
                with col4:
                    st.text(modified)
                with col5:
                    if is_folder:
                        if st.button("📂 Open", key=f"open_{item['id']}", use_container_width=True):
                            self.navigate_to_folder(item['id'], item['name'])
                    else:
                        if st.button("⬇️ Download", key=f"dl_{item['id']}", use_container_width=True):
                            self.download_file(item['id'], item['name'])
                with col6:
                    if st.button("🗑️", key=f"del_{item['id']}", help=f"Delete {item['name']}", use_container_width=True):
                        self.confirm_delete(item['id'], item['name'], is_folder)
                
                st.markdown("---")
            
            # Pagination controls at bottom (optional, for convenience)
            st.markdown("---")
            col1, col2, col3, col4, col5 = st.columns([2, 1, 2, 1, 2])
            
            with col1:
                st.markdown(f"**📊 Page {current_page + 1} of {total_pages}**")
            
            with col2:
                if st.button("⏮️", key="first_bottom", disabled=current_page == 0, use_container_width=True):
                    st.session_state.drive_page = 0
                    st.rerun()
            
            with col3:
                if st.button("◀️", key="prev_bottom", disabled=current_page == 0, use_container_width=True):
                    st.session_state.drive_page = max(0, current_page - 1)
                    st.rerun()
            
            with col4:
                if st.button("▶️", key="next_bottom", disabled=current_page >= total_pages - 1, use_container_width=True):
                    st.session_state.drive_page = min(total_pages - 1, current_page + 1)
                    st.rerun()
            
            with col5:
                if st.button("⏭️", key="last_bottom", disabled=current_page >= total_pages - 1, use_container_width=True):
                    st.session_state.drive_page = total_pages - 1
                    st.rerun()
                    
        except Exception as e:
            st.error(f"❌ Error displaying contents: {str(e)}")
    def create_folder(self, folder_name):
        """Create a new folder in current location"""
        try:
            parent_id = st.session_state.current_folder or st.session_state.shared_drive_id
            
            folder_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_id]
            }
            
            folder = st.session_state.drive_service.files().create(
                body=folder_metadata,
                fields='id',
                supportsAllDrives=True
            ).execute()
            
            st.success(f"✅ Folder '{folder_name}' created successfully!")
            self.log_message(f"📁 Created folder: {folder_name}")
            
        except Exception as e:
            st.error(f"❌ Failed to create folder: {str(e)}")
            self.log_message(f"❌ Folder creation failed: {str(e)}")

    def confirm_delete(self, file_id, file_name, is_folder):
        """Show delete confirmation dialog"""
        item_type = "folder" if is_folder else "file"
        
        @st.dialog(f"Delete {item_type}?")
        def delete_dialog():
            st.warning(f"⚠️ Are you sure you want to delete '{file_name}'?")
            if is_folder:
                st.error("🚨 This will delete the folder and all its contents!")
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ Confirm Delete", type="primary", use_container_width=True):
                    self.delete_item(file_id, file_name, is_folder)
                    st.rerun()
            with col2:
                if st.button("❌ Cancel", use_container_width=True):
                    st.rerun()
        
        delete_dialog()

    def delete_item(self, file_id, file_name, is_folder):
        """Delete a file or folder from Drive"""
        try:
            st.session_state.drive_service.files().delete(
                fileId=file_id,
                supportsAllDrives=True
            ).execute()
            
            item_type = "folder" if is_folder else "file"
            st.success(f"✅ {item_type.capitalize()} '{file_name}' deleted successfully!")
            self.log_message(f"🗑️ Deleted {item_type}: {file_name}")
            
        except Exception as e:
            st.error(f"❌ Failed to delete: {str(e)}")
            self.log_message(f"❌ Delete failed: {file_name} - {str(e)}")

    def navigate_to_folder(self, folder_id, folder_name=None):
        """Navigate into a folder"""
        if st.session_state.current_folder:
            st.session_state.folder_stack.append({
                'id': st.session_state.current_folder,
                'name': st.session_state.current_path
            })
        st.session_state.current_folder = folder_id
        st.session_state.current_path = folder_name or "Folder"
        st.session_state.drive_page = 0  # Reset to first page when navigating
        st.rerun()

    def download_file(self, file_id, filename):
        """Download a file from Drive"""
        with st.spinner(f"⬇️ Downloading {filename}..."):
            try:
                request = st.session_state.drive_service.files().get_media(
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
                    label=f"📥 Save {filename}",
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

    def go_to_root(self):
        """Navigate to root folder"""
        st.session_state.current_folder = None
        st.session_state.current_path = 'Root'
        st.session_state.folder_stack = []
        st.session_state.drive_page = 0  # Reset to first page
        st.rerun()

    def go_back(self):
        """Go back to previous folder"""
        if st.session_state.folder_stack:
            prev = st.session_state.folder_stack.pop()
            st.session_state.current_folder = prev['id']
            st.session_state.current_path = prev['name']
            st.session_state.drive_page = 0  # Reset to first page
            st.rerun()
        else:
            self.go_to_root()

    def refresh_drive(self):
        """Refresh current view"""
        st.session_state.drive_page = 0  # Reset to first page on refresh
        st.rerun()

    def format_phone_number(self, phone_str):
        """Format phone to +91XXXXXXXXXX"""
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
        """Send SMS via Qik API"""
        if not phone:
            return False, "Empty phone"
        
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
            return (resp.status_code in (200, 201)), resp.text
        except Exception as e:
            return False, str(e)

    def process_and_send_sms(self, excel_file):
        """Process Excel and send SMS to employees with Drive Links"""
        try:
            df = pd.read_excel(excel_file, dtype=str).fillna("")
            
            required = ['Employee Name', 'Employee no', 'UAN', 'Drive Link']
            for col in required:
                if col not in df.columns:
                    st.error(f"❌ Missing column: {col}")
                    return
            
            # Get previous month
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
            
            for i, row in df.iterrows():
                name = row['Employee Name']
                emp_no_raw = str(row['Employee no']).strip()
                drive_link = str(row['Drive Link']).strip()
                
                status.info(f"⏳ Processing {i+1}/{total}: {name}")
                
                # Only send if Drive Link exists
                if not drive_link:
                    skipped_count += 1
                    results.append({
                        'row': i+1,
                        'name': name,
                        'phone': emp_no_raw,
                        'status': 'skipped_no_link'
                    })
                    self.log_message(f"⏭️ Skipped {name} - No Drive Link")
                    progress.progress((i+1)/total)
                    continue
                
                if not emp_no_raw:
                    skipped_count += 1
                    results.append({
                        'row': i+1,
                        'name': name,
                        'phone': '',
                        'status': 'skipped_no_phone'
                    })
                    self.log_message(f"⏭️ Skipped {name} - No phone")
                    progress.progress((i+1)/total)
                    continue
                
                phone = self.format_phone_number(emp_no_raw)
                success, resp_text = self.send_sms_via_qik(name, phone, drive_link, prev_month_str)
                
                if success:
                    sent_count += 1
                    results.append({
                        'row': i+1,
                        'name': name,
                        'phone': phone,
                        'status': 'sent',
                        'response': resp_text
                    })
                    self.log_message(f"✅ SMS sent: {name} ({phone})")
                else:
                    failed_count += 1
                    results.append({
                        'row': i+1,
                        'name': name,
                        'phone': phone,
                        'status': 'failed',
                        'response': resp_text
                    })
                    self.log_message(f"❌ SMS failed: {name} - {resp_text}")
                
                time.sleep(0.25)
                progress.progress((i+1)/total)
            
            progress.empty()
            status.empty()
            
            # Display results
            col1, col2, col3 = st.columns(3)
            with col1:
                st.success(f"✅ Sent: {sent_count}")
            with col2:
                st.error(f"❌ Failed: {failed_count}")
            with col3:
                st.warning(f"⏭️ Skipped: {skipped_count}")
            
            self.log_message(f"📊 SMS Summary - Sent: {sent_count}, Failed: {failed_count}, Skipped: {skipped_count}")
            
            # Download report
            report_df = pd.DataFrame(results)
            csv_buf = report_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                "📥 Download SMS Report (CSV)",
                data=csv_buf,
                file_name=f"sms_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"❌ SMS processing error: {str(e)}")
            self.log_message(f"❌ SMS error: {str(e)}")

    def reset_session(self):
        """Reset processing session"""
        st.session_state.files_processed = False
        st.session_state.processing_complete = False
        st.session_state.results = {}
        st.session_state.updated_excel_buffer = None
        st.session_state.uan_pages = {}
        st.session_state.df = None
        st.session_state.pdf_count = 0
        st.session_state.excel_count = 0
        st.success("✅ Session reset successfully!")
        self.log_message("🔄 Session reset")
        time.sleep(1)
        st.rerun()

    def clear_session_data(self):
        """Clear all session data including logs"""
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        self.initialize_session_state()
        st.success("✅ All data cleared successfully!")
        time.sleep(1)
        st.rerun()

    def log_message(self, message):
        """Add message to activity log with timestamp"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        st.session_state.log_entries.append(log_entry)
        
        # Keep only last 500 entries
        if len(st.session_state.log_entries) > 500:
            st.session_state.log_entries = st.session_state.log_entries[-500:]


# Main execution
if __name__ == "__main__":
    app = PayslipDistributorStreamlit()
    app.run()
