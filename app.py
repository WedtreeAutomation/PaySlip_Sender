# payslip_distributor_app.py
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
        # default shared drive id
        self.shared_drive_id = os.getenv("SHARED_DRIVE_ID")
        self.initialize_session_state()
        self.setup_page()

    def initialize_session_state(self):
        """Initialize session state variables used by the app"""
        if 'sent_numbers' not in st.session_state:
            st.session_state.sent_numbers = {}
        if 'log_entries' not in st.session_state:
            st.session_state.log_entries = []
        if 'current_folder' not in st.session_state:
            st.session_state.current_folder = None
        if 'folder_stack' not in st.session_state:
            st.session_state.folder_stack = []
        if 'drive_service' not in st.session_state:
            st.session_state.drive_service = None
        if 'drive_initialized' not in st.session_state:
            st.session_state.drive_initialized = False
        if 'shared_drive_id' not in st.session_state:
            st.session_state.shared_drive_id = self.shared_drive_id
        if 'files_processed' not in st.session_state:
            st.session_state.files_processed = False
        if 'uan_pages' not in st.session_state:
            st.session_state.uan_pages = {}
        if 'df' not in st.session_state:
            st.session_state.df = None
        if 'pdf_count' not in st.session_state:
            st.session_state.pdf_count = 0
        if 'excel_count' not in st.session_state:
            st.session_state.excel_count = 0
        if 'processing_complete' not in st.session_state:
            st.session_state.processing_complete = False
        if 'results' not in st.session_state:
            st.session_state.results = {}
        if 'updated_excel_buffer' not in st.session_state:
            st.session_state.updated_excel_buffer = None

    def setup_page(self):
        st.set_page_config(
            page_title="Payslip Distribution System",
            page_icon="📄",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        # Keep your exact CSS for consistent look
        st.markdown("""
        <style>
        .main-header {
            font-size: 2.5rem;
            font-weight: bold;
            color: #1f77b4;
            text-align: center;
            margin-bottom: 1rem;
        }
        .sub-header {
            font-size: 1.2rem;
            color: #666;
            text-align: center;
            margin-bottom: 2rem;
        }
        .section-header {
            font-size: 1.5rem;
            font-weight: bold;
            color: #1f77b4;
            margin-top: 2rem;
            margin-bottom: 1rem;
            padding-bottom: 0.5rem;
            border-bottom: 2px solid #1f77b4;
        }
        .success-box {
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            border-radius: 5px;
            padding: 1rem;
            margin: 1rem 0;
        }
        .error-box {
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
            border-radius: 5px;
            padding: 1rem;
            margin: 1rem 0;
        }
        .info-box {
            background-color: #d1ecf1;
            border: 1px solid #bee5eb;
            border-radius: 5px;
            padding: 1rem;
            margin: 1rem 0;
        }
        .log-box {
            background-color: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 5px;
            padding: 1rem;
            margin: 1rem 0;
            max-height: 300px;
            overflow-y: auto;
        }
        .stButton button {
            width: 100%;
        }
        </style>
        """, unsafe_allow_html=True)

    def run(self):
        # Header
        st.markdown('<div class="main-header">📄 Payslip Distribution System</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header">Fast, Secure & Efficient Employee Payslip Distribution</div>', unsafe_allow_html=True)

        # Tabs: 4 tabs as requested: Send Payslips, Drive Explorer, Send SMS, Settings
        tab1, tab2, tab3, tab4 = st.tabs(["📤 Send Payslips", "📁 Drive Explorer", "📨 Send SMS", "⚙️ Settings"])

        with tab1:
            self.render_send_tab()
        with tab2:
            self.render_drive_tab()
        with tab3:
            self.render_sms_tab()
        with tab4:
            self.render_settings_tab()

        # Activity log at bottom
        st.markdown('<div class="section-header">📜 Activity Log</div>', unsafe_allow_html=True)
        if st.session_state.log_entries:
            for entry in reversed(st.session_state.log_entries[-50:]):
                st.text(entry)
        else:
            st.info("No activity yet.")

    # ---------------- UI renderers ----------------
    def render_send_tab(self):
        st.markdown('<div class="section-header">🔐 Google Drive Setup</div>', unsafe_allow_html=True)
        col1, col2 = st.columns([3, 1])
        with col1:
            # Initialize button - user must click
            if st.button("Initialize Google Drive Connection", type="primary", key="init_drive"):
                self.initialize_drive()
        with col2:
            drive_status = "✅ Connected" if st.session_state.drive_initialized else "❌ Not connected"
            st.info(f"Status: {drive_status}")

        # If not initialized, block further operations
        if not st.session_state.drive_initialized:
            st.warning("Please click 'Initialize Google Drive Connection' before uploading or processing files.")
            # show file upload but disable process
            col1, col2 = st.columns(2)
            with col1:
                st.file_uploader("PDF Payslip File (disabled until Drive init)", type=["pdf"], key="pdf_upload_disabled")
            with col2:
                st.file_uploader("Employee Contacts File (disabled until Drive init)", type=["xlsx", "xls"], key="excel_upload_disabled")
            return

        # Drive initialized -> allow uploads and processing
        st.markdown('<div class="section-header">📤 Upload Files</div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            pdf_file = st.file_uploader("PDF Payslip File", type=["pdf"], key="pdf_upload")
        with col2:
            excel_file = st.file_uploader("Employee Contacts File", type=["xlsx", "xls"], key="excel_upload")

        if st.button("Process Files", type="primary", key="process_files"):
            if not pdf_file or not excel_file:
                st.error("Please upload both PDF and Excel files")
            else:
                # process only when drive initialized
                if not st.session_state.drive_initialized:
                    st.error("Google Drive not initialized. Initialize first.")
                else:
                    self.process_files(pdf_file, excel_file)

        # Display analysis if processed
        if st.session_state.files_processed:
            st.markdown('<div class="section-header">📊 File Analysis</div>', unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                st.success(f"PDF Analysis: {st.session_state.pdf_count} UANs found")
            with col2:
                st.success(f"Employee Data: {st.session_state.excel_count} employees")

        # Send payslips to drive & update excel
        if st.session_state.files_processed:
            st.markdown('<div class="section-header">🚀 Process Payslips</div>', unsafe_allow_html=True)
            confirm = st.checkbox("I confirm I want to process payslips for all employees", key="confirm_checkbox")
            if st.button("📤 Upload Payslips to Drive & Update Excel", type="primary", key="upload_payslips", disabled=not confirm):
                if not st.session_state.drive_initialized:
                    st.error("Google Drive not initialized. Initialize first.")
                else:
                    self.process_payslips()

        # Results and download
        if st.session_state.processing_complete:
            st.markdown('<div class="section-header">📊 Results</div>', unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            with col1:
                st.success(f"✅ Uploaded: {st.session_state.results.get('uploaded', 0)}")
            with col2:
                st.error(f"❌ Failed: {st.session_state.results.get('failed', 0)}")
            with col3:
                st.warning(f"⏭️ Skipped: {st.session_state.results.get('skipped', 0)}")

            if st.session_state.updated_excel_buffer:
                st.download_button(
                    label="📥 Download Updated Excel File",
                    data=st.session_state.updated_excel_buffer,
                    file_name="send_payslip_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel"
                )

        # Reset button
        if st.button("🔄 Reset Session", type="secondary", key="reset_session"):
            self.reset_session()

    def render_drive_tab(self):
        st.markdown('<div class="section-header">📁 Google Drive Explorer</div>', unsafe_allow_html=True)
        # Navigation controls
        col1, col2, col3, col4 = st.columns([2,2,2,4])
        with col1:
            if st.button("🏠 Root Folder", key="root_folder"):
                self.go_to_root()
        with col2:
            if st.button("⬅️ Go Back", key="go_back"):
                self.go_back()
        with col3:
            if st.button("🔄 Refresh", key="refresh_drive"):
                self.refresh_drive()
        with col4:
            current_path = "Root" if st.session_state.current_folder is None else "Current Folder"
            st.text(f"Current: {current_path}")

        # Require drive initialized
        if not st.session_state.drive_initialized:
            st.warning("Please initialize Google Drive connection first")
            return

        self.display_drive_contents()

    def render_sms_tab(self):
        st.markdown('<div class="section-header">📨 Send SMS to Employees</div>', unsafe_allow_html=True)
        st.markdown("Upload the updated Excel (Employee Name, Employee no, UAN, Drive Link) and the app will send SMS only to rows that have a Drive Link.")

        if not st.session_state.drive_initialized:
            st.warning("Please initialize Google Drive connection first")
            return

        uploaded = st.file_uploader("Upload updated Excel (from Tab1)", type=["xlsx", "xls"], key="sms_excel_upload")
        if st.button("Send SMS to employees with Drive Link", key="send_sms_btn"):
            if not uploaded:
                st.error("Please upload the updated Excel from Tab 1 first.")
            else:
                self.process_and_send_sms(uploaded)

    def render_settings_tab(self):
        st.markdown('<div class="section-header">⚙️ System Settings</div>', unsafe_allow_html=True)
        st.subheader("Configuration")
        drive_id = st.text_input(
            "Shared Drive ID",
            value=st.session_state.shared_drive_id or self.shared_drive_id,
            help="Enter your Google Shared Drive ID",
            key="drive_id_input"
        )
        if st.button("💾 Save Settings", key="save_settings"):
            st.session_state.shared_drive_id = drive_id
            st.success("Settings saved successfully!")
            self.log_message("Saved shared drive id")

        st.subheader("📊 System Status")
        col1, col2 = st.columns(2)
        with col1:
            drive_status = "✅ Connected" if st.session_state.drive_initialized else "❌ Disconnected"
            st.info(f"Google Drive: {drive_status}")
        with col2:
            st.info(f"Shared Drive ID: {st.session_state.shared_drive_id or 'Not set'}")

        st.subheader("⚠️ Danger Zone")
        if st.button("🗑️ Clear All Session Data", type="secondary", key="clear_session"):
            self.clear_session_data()

    # ---------------- Core actions ----------------
    def initialize_drive(self):
        """User must click this. On success it sets drive_service and drive_initialized True"""
        with st.spinner("Connecting to Google Drive..."):
            service = self.get_drive_service()
            if service:
                st.session_state.drive_service = service
                st.session_state.drive_initialized = True
                st.success("✅ Google Drive is ready for uploads!")
                self.log_message("✅ Google Drive connection initialized successfully")
                # Keep UI showing initialized until page refresh (session_state persists)
            else:
                st.session_state.drive_initialized = False
                st.error("❌ Failed to connect to Google Drive")
                self.log_message("❌ Failed to initialize Google Drive connection")

    def process_files(self, pdf_file, excel_file):
        """Process uploaded files: extract UAN pages, read excel, store in session state"""
        if not st.session_state.drive_initialized:
            st.error("Initialize Google Drive first")
            return

        with st.spinner("Processing files..."):
            try:
                # Process PDF
                uan_pages = self.process_pdf(pdf_file)
                st.session_state.pdf_count = len(uan_pages)

                # Process Excel - ensure Employee Name present, rename UAN column if needed
                df = pd.read_excel(excel_file, dtype=str).fillna("")
                # Accept both 'UAN/member ID' or 'UAN'
                if 'UAN/member ID' in df.columns and 'UAN' not in df.columns:
                    df.rename(columns={'UAN/member ID': 'UAN'}, inplace=True)
                if 'UAN' not in df.columns:
                    st.error("Input Excel must have 'UAN' or 'UAN/member ID' column.")
                    return

                # Ensure Employee Name column exists
                if 'Employee Name' not in df.columns:
                    if 'Name' in df.columns:
                        df.rename(columns={'Name': 'Employee Name'}, inplace=True)
                    else:
                        df['Employee Name'] = ""

                # Ensure Employee no column present
                if 'Employee no' not in df.columns and 'Employee No' in df.columns:
                    df.rename(columns={'Employee No': 'Employee no'}, inplace=True)
                if 'Employee no' not in df.columns:
                    st.error("Input Excel must have 'Employee no' column (10-digit number).")
                    return

                # No formatted phone column stored in df (we will format only when sending SMS)
                st.session_state.excel_count = len(df)

                # Store in session for later
                st.session_state.uan_pages = uan_pages
                st.session_state.df = df
                st.session_state.files_processed = True
                st.session_state.processing_complete = False
                st.session_state.pdf_file = pdf_file
                st.session_state.excel_file = excel_file

                st.success("Files processed successfully!")
                self.log_message(f"📊 Processed PDF: {st.session_state.pdf_count} UANs found")
                self.log_message(f"📊 Processed Excel: {st.session_state.excel_count} employees")

                # Rerun to update UI
                st.rerun()
            except Exception as e:
                st.error(f"Failed to process files: {str(e)}")
                self.log_message(f"❌ Error processing files: {str(e)}")

    def process_payslips(self):
        """Extract individual payslips, upload to drive; update df with Drive Link and prepare output Excel"""
        if not st.session_state.files_processed or st.session_state.df is None:
            st.error("Please process files first")
            return
        if not st.session_state.drive_initialized:
            st.error("Please initialize Google Drive first")
            return

        with st.spinner("Processing payslips..."):
            try:
                # previous month
                today = datetime.now()
                first_day_of_current_month = today.replace(day=1)
                last_month = first_day_of_current_month - timedelta(days=1)
                current_month = last_month.strftime("%B %Y")

                folder_id = self.get_monthly_folder_id(
                    st.session_state.drive_service,
                    current_month,
                    st.session_state.shared_drive_id
                )
                if not folder_id:
                    st.error("Failed to create monthly folder")
                    return

                results = {'uploaded': 0, 'failed': 0, 'skipped': 0}
                updated_df = st.session_state.df.copy()

                # Ensure Drive Link column exists and final column order
                if 'Drive Link' not in updated_df.columns:
                    updated_df['Drive Link'] = ""

                progress_bar = st.progress(0)
                status_text = st.empty()
                total = len(updated_df)

                for idx, (_, row) in enumerate(updated_df.iterrows()):
                    # Use UAN column
                    uan = str(row.get('UAN', '')).strip()
                    emp_no = str(row.get('Employee no', '')).strip()

                    progress = (idx + 1) / max(1, total)
                    progress_bar.progress(progress)
                    status_text.text(f"Processing {idx + 1} of {total}: UAN {uan}")

                    # Skip if already processed this session based on UAN (or phone)
                    if uan == "":
                        self.log_message(f"⚠️ Skipped row {idx+1}: missing UAN")
                        results['skipped'] += 1
                        continue

                    if uan in st.session_state.uan_pages:
                        try:
                            st.session_state.pdf_file.seek(0)
                            page_num = st.session_state.uan_pages[uan]
                            pdf_buffer = self.extract_individual_payslip(
                                st.session_state.pdf_file,
                                uan,
                                page_num
                            )
                            filename = f"Payslip_{uan}_{current_month.replace(' ', '')}.pdf"
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
                                self.log_message(f"✅ Successfully uploaded for UAN: {uan} - Link: {drive_link}")
                                results['uploaded'] += 1
                            else:
                                self.log_message(f"❌ Failed to upload payslip for UAN {uan}")
                                results['failed'] += 1
                        except Exception as e:
                            self.log_message(f"❌ Failed to process UAN {uan}: {str(e)}")
                            results['failed'] += 1
                    else:
                        self.log_message(f"⚠️ No payslip found for UAN {uan}")
                        results['skipped'] += 1

                # Finalize the output DataFrame with exact columns: Employee Name, Employee no, UAN, Drive Link
                final_df = updated_df.copy()
                # Ensure these columns exist and in order
                for col in ['Employee Name', 'Employee no', 'UAN', 'Drive Link']:
                    if col not in final_df.columns:
                        final_df[col] = ""

                final_df = final_df[['Employee Name', 'Employee no', 'UAN', 'Drive Link']]

                # Write to BytesIO
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Payslips')
                st.session_state.updated_excel_buffer = output.getvalue()
                st.session_state.results = results
                st.session_state.processing_complete = True

                progress_bar.empty()
                status_text.empty()

                st.success("Payslip processing completed!")
                self.log_message("🎉 All payslips processed (attempted)!")
                st.rerun()
            except Exception as e:
                st.error(f"Error during payslip processing: {str(e)}")
                self.log_message(f"❌ Processing error: {str(e)}")

    def process_pdf(self, pdf_file):
        """Extract UAN -> page mapping using pdfplumber"""
        uan_pages = {}
        try:
            pdf_file.seek(0)
            with pdfplumber.open(pdf_file) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    # match common variations
                    match = re.search(r"(?:UAN\/?MEMBER ID|UAN|UAN MEMBER ID)[:\s]*([A-Za-z0-9\-]+)", text, re.I)
                    if match:
                        uan = match.group(1).strip()
                        uan_pages[uan] = page_num
            return uan_pages
        except Exception as e:
            st.error(f"Error processing PDF: {str(e)}")
            return {}

    def extract_individual_payslip(self, pdf_file, uan, page_num):
        """Return BytesIO with single-page PDF for given page_num"""
        pdf_reader = PdfReader(pdf_file)
        pdf_writer = PdfWriter()
        pdf_writer.add_page(pdf_reader.pages[page_num])
        output_buffer = BytesIO()
        pdf_writer.write(output_buffer)
        output_buffer.seek(0)
        return output_buffer

    # ---------------- Google Drive helpers ----------------
    def get_google_credentials(self):
        """Get Google credentials from environment variables"""
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
                    st.error(f"Missing environment variable for: {key}")
                    return None
            return service_account.Credentials.from_service_account_info(
                credentials_dict,
                scopes=['https://www.googleapis.com/auth/drive']
            )
        except Exception as e:
            st.error(f"Failed to load Google credentials: {str(e)}")
            return None

    def get_drive_service(self):
        """Initialize Drive client"""
        try:
            credentials = self.get_google_credentials()
            if credentials:
                return build('drive', 'v3', credentials=credentials)
            return None
        except Exception as e:
            st.error(f"Failed to initialize Google Drive service: {str(e)}")
            return None

    def get_monthly_folder_id(self, service, month_year, shared_drive_id):
        """Get or create monthly folder"""
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
            return folder['id']
        except HttpError as error:
            if hasattr(error, 'resp') and error.resp.status == 503:
                self.log_message("⚠️ Transient error, retrying in 5 seconds...")
                time.sleep(5)
                return self.get_monthly_folder_id(service, month_year, shared_drive_id)
            else:
                st.error(f"Google Drive API error: {error}")
                return None
        except Exception as e:
            st.error(f"Error creating folder: {str(e)}")
            return None

    def upload_to_drive(self, service, file_buffer, filename, folder_id, shared_drive_id):
        """Upload file to Drive and make it publicly readable"""
        max_retries = 3
        retry_count = 0
        while retry_count < max_retries:
            try:
                file_buffer.seek(0)
                media = MediaIoBaseUpload(file_buffer, mimetype='application/pdf')
                file_metadata = {'name': filename, 'parents': [folder_id]}
                file = service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id',
                    supportsAllDrives=True
                ).execute()
                permission = {'type': 'anyone', 'role': 'reader', 'allowFileDiscovery': False}
                service.permissions().create(
                    fileId=file['id'],
                    body=permission,
                    supportsAllDrives=True
                ).execute()
                file_link = f"https://drive.google.com/uc?export=download&id={file['id']}"
                return file_link, file['id']
            except HttpError as error:
                if hasattr(error, 'resp') and error.resp.status == 503 and retry_count < max_retries:
                    self.log_message(f"⚠️ Transient error uploading {filename}, retrying ({retry_count+1}/{max_retries})...")
                    retry_count += 1
                    time.sleep(5)
                    continue
                else:
                    st.error(f"Google Drive API error: {error}")
                    return None, None
            except Exception as e:
                st.error(f"Error uploading to Drive: {str(e)}")
                return None, None
        return None, None

    def list_drive_contents(self, service, shared_drive_id, folder_id=None):
        """List contents of shared drive or folder"""
        try:
            query = "trashed=false"
            if folder_id:
                query += f" and '{folder_id}' in parents"
                results = service.files().list(
                    q=query,
                    fields="files(id, name, mimeType, size, createdTime, modifiedTime, webViewLink)",
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True,
                    orderBy="name"
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
                    orderBy="name"
                ).execute()
            return results.get('files', [])
        except Exception as e:
            st.error(f"Error listing drive contents: {str(e)}")
            return []

    def display_drive_contents(self):
        """Render drive listing UI"""
        try:
            contents = self.list_drive_contents(
                st.session_state.drive_service,
                st.session_state.shared_drive_id,
                st.session_state.current_folder
            )
            if not contents:
                st.info("No files or folders found in this location.")
                return
            for item in contents:
                item_type = "📁 Folder" if item['mimeType'] == 'application/vnd.google-apps.folder' else "📄 File"
                size = f"{int(item.get('size', 0)) / 1024:.1f} KB" if 'size' in item else ""
                modified = item.get('modifiedTime', '')[:16].replace('T', ' ') if 'modifiedTime' in item else ""
                col1, col2, col3, col4, col5 = st.columns([4,1,1,2,2])
                with col1:
                    st.text(item['name'])
                with col2:
                    st.text(item_type)
                with col3:
                    st.text(size)
                with col4:
                    st.text(modified)
                with col5:
                    if item_type == "📁 Folder":
                        if st.button("Open", key=f"open_{item['id']}"):
                            self.navigate_to_folder(item['id'])
                    else:
                        if st.button("Download", key=f"dl_{item['id']}"):
                            self.download_file(item['id'], item['name'])
        except Exception as e:
            st.error(f"Error displaying drive contents: {str(e)}")

    def navigate_to_folder(self, folder_id):
        if st.session_state.current_folder:
            st.session_state.folder_stack.append(st.session_state.current_folder)
        st.session_state.current_folder = folder_id
        st.rerun()

    def download_file(self, file_id, filename):
        with st.spinner(f"Downloading {filename}..."):
            try:
                request = st.session_state.drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
                file_buffer = BytesIO()
                downloader = MediaIoBaseDownload(file_buffer, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                file_buffer.seek(0)
                st.download_button(
                    label=f"📥 Download {filename}",
                    data=file_buffer.getvalue(),
                    file_name=filename,
                    mime="application/octet-stream",
                    key=f"download_{file_id}"
                )
                self.log_message(f"📥 Downloaded: {filename}")
            except Exception as e:
                st.error(f"Failed to download file: {str(e)}")
                self.log_message(f"❌ Download failed: {filename} - {str(e)}")

    def go_to_root(self):
        st.session_state.current_folder = None
        st.session_state.folder_stack = []
        st.rerun()

    def go_back(self):
        if st.session_state.folder_stack:
            st.session_state.current_folder = st.session_state.folder_stack.pop()
            st.rerun()
        else:
            self.go_to_root()

    def refresh_drive(self):
        st.rerun()

    def reset_session(self):
        # Reset only processing-related session state
        st.session_state.files_processed = False
        st.session_state.processing_complete = False
        st.session_state.results = {}
        st.session_state.updated_excel_buffer = None
        st.success("Session data cleared!")
        st.rerun()

    def clear_session_data(self):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.success("All session data cleared!")
        st.rerun()

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        st.session_state.log_entries.append(log_entry)

    # ---------------- SMS sending ----------------
    def format_phone_number(self, phone_str):
        """Format phone number to +91XXXXXXXXXX as requested (excel contains 10 digits)"""
        if pd.isna(phone_str) or str(phone_str).strip() == "":
            return ""
        digits = re.sub(r'\D', '', str(phone_str))
        if len(digits) == 10:
            return f"+91{digits}"
        if len(digits) == 11 and digits.startswith('0'):
            return f"+91{digits[1:]}"
        if digits.startswith('91') and len(digits) == 12:
            return f"+{digits}"
        # fallback: if already has + keep it
        if str(phone_str).strip().startswith("+"):
            return str(phone_str).strip()
        return digits

    def send_sms_via_qik(self, name, phone, download_link, previous_month):
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
            return (resp.status_code in (200,201)), resp.text
        except Exception as e:
            return False, str(e)

    def process_and_send_sms(self, excel_file):
        """Send SMS only to rows that have Drive Link"""
        try:
            df = pd.read_excel(excel_file, dtype=str).fillna("")
            required = ['Employee Name', 'Employee no', 'UAN', 'Drive Link']
            for col in required:
                if col not in df.columns:
                    st.error(f"Uploaded file missing required column: {col}")
                    self.log_message(f"Upload missing column: {col}")
                    return

            # previous month
            today = datetime.now()
            first_day_of_current_month = today.replace(day=1)
            last_month = first_day_of_current_month - timedelta(days=1)
            prev_month_str = last_month.strftime("%B %Y")

            total = len(df)
            progress = st.progress(0)
            sent_count = 0
            failed_count = 0
            skipped_count = 0
            results = []

            for i, row in df.iterrows():
                name = row['Employee Name']
                emp_no_raw = str(row['Employee no']).strip()
                drive_link = str(row['Drive Link']).strip()

                # RULE: Only send when Drive Link present
                if not drive_link:
                    skipped_count += 1
                    results.append({'row': i+1, 'name': name, 'phone': emp_no_raw, 'status': 'skipped_no_drive_link'})
                    self.log_message(f"Skipped {name} — missing Drive Link")
                    progress.progress((i+1)/max(1,total))
                    continue

                if not emp_no_raw:
                    skipped_count += 1
                    results.append({'row': i+1, 'name': name, 'phone': emp_no_raw, 'status': 'skipped_no_phone'})
                    self.log_message(f"Skipped {name} — missing phone")
                    progress.progress((i+1)/max(1,total))
                    continue

                phone = self.format_phone_number(emp_no_raw)
                success, resp_text = self.send_sms_via_qik(name, phone, drive_link, prev_month_str)
                if success:
                    sent_count += 1
                    results.append({'row': i+1, 'name': name, 'phone': phone, 'status': 'sent', 'response': resp_text})
                    self.log_message(f"SMS sent to {name} ({phone})")
                else:
                    failed_count += 1
                    results.append({'row': i+1, 'name': name, 'phone': phone, 'status': 'failed', 'response': resp_text})
                    self.log_message(f"SMS failed for {name} ({phone}): {resp_text}")

                time.sleep(0.25)
                progress.progress((i+1)/max(1,total))

            st.success(f"SMS send complete — Sent: {sent_count}, Failed: {failed_count}, Skipped: {skipped_count}")
            self.log_message(f"SMS summary - Sent:{sent_count} Failed:{failed_count} Skipped:{skipped_count}")

            report_df = pd.DataFrame(results)
            csv_buf = report_df.to_csv(index=False).encode('utf-8')
            st.download_button("Download SMS report (CSV)", data=csv_buf, file_name="sms_send_report.csv", mime="text/csv")
        except Exception as e:
            st.error(f"Error sending SMS: {e}")
            self.log_message(f"Tab3 processing error: {e}")

# ---------------- Run the app ----------------
if __name__ == "__main__":
    app = PayslipDistributorStreamlit()
    app.run()
