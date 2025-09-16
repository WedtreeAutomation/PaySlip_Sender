import streamlit as st
import pandas as pd
import re
import os
import time
from io import BytesIO
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError
import pdfplumber
from PyPDF2 import PdfReader, PdfWriter
from dotenv import load_dotenv
import json

# Load environment variables
load_dotenv()

# ---------------------------------------------------
# Google Drive Service Functions
# ---------------------------------------------------
def get_google_credentials():
    """Get Google credentials from environment variables"""
    try:
        credentials_dict = {
            "type": os.getenv("GOOGLE_SERVICE_ACCOUNT_TYPE"),
            "project_id": os.getenv("GOOGLE_PROJECT_ID"),
            "private_key_id": os.getenv("GOOGLE_PRIVATE_KEY_ID"),
            "private_key": os.getenv("GOOGLE_PRIVATE_KEY").replace('\\n', '\n'),
            "client_email": os.getenv("GOOGLE_CLIENT_EMAIL"),
            "client_id": os.getenv("GOOGLE_CLIENT_ID"),
            "auth_uri": os.getenv("GOOGLE_AUTH_URI"),
            "token_uri": os.getenv("GOOGLE_TOKEN_URI"),
            "auth_provider_x509_cert_url": os.getenv("GOOGLE_AUTH_PROVIDER_CERT_URL"),
            "client_x509_cert_url": os.getenv("GOOGLE_CLIENT_CERT_URL")
        }
        
        # Validate that all required environment variables are present
        for key, value in credentials_dict.items():
            if not value:
                raise ValueError(f"Missing environment variable for: {key}")
                
        return service_account.Credentials.from_service_account_info(
            credentials_dict,
            scopes=['https://www.googleapis.com/auth/drive']
        )
    except Exception as e:
        st.error(f"Failed to load Google credentials: {str(e)}")
        return None

@st.cache_resource
def get_drive_service():
    """Initialize Google Drive service using credentials from environment variables"""
    try:
        credentials = get_google_credentials()
        if credentials:
            return build('drive', 'v3', credentials=credentials)
        return None
    except Exception as e:
        st.error(f"Failed to initialize Google Drive service: {str(e)}")
        return None

# ---------------------------------------------------
# Helper Functions
# ---------------------------------------------------
def format_phone_number(phone_str):
    """Format phone number to +91XXXXXXXXXX"""
    digits = re.sub(r'\D', '', str(phone_str))
    if len(digits) == 10:
        return f"+91{digits}"
    elif len(digits) == 11 and digits.startswith('0'):
        return f"+91{digits[1:]}"
    elif digits.startswith('91') and len(digits) == 12:
        return f"+{digits}"
    else:
        return f"+{digits}" if not digits.startswith('+') else digits

def extract_individual_payslip(pdf_file, uan, page_num):
    """Extract individual payslip as PDF for a specific UAN"""
    pdf_reader = PdfReader(pdf_file)
    pdf_writer = PdfWriter()
    
    # Add the specific page to the new PDF
    pdf_writer.add_page(pdf_reader.pages[page_num])
    
    # Create a bytes buffer for the new PDF
    output_buffer = BytesIO()
    pdf_writer.write(output_buffer)
    output_buffer.seek(0)
    
    return output_buffer

def get_monthly_folder_id(service, month_year, shared_drive_id):
    """Get or create monthly folder in Google Drive Shared Drive"""
    try:
        # Check if folder already exists in Shared Drive
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
        
        # Create new folder in Shared Drive
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
        st.error(f"Google Drive API error: {error}")
        return None
    except Exception as e:
        st.error(f"Error creating folder: {str(e)}")
        return None

def upload_to_drive(service, file_buffer, filename, folder_id, shared_drive_id):
    """Upload file to Google Drive Shared Drive and make it publicly accessible"""
    try:
        # Upload file to Shared Drive
        media = MediaIoBaseUpload(file_buffer, mimetype='application/pdf')
        file_metadata = {
            'name': filename,
            'parents': [folder_id]
        }
        
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id',
            supportsAllDrives=True
        ).execute()
        
        # Make file publicly accessible with download permission
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
        
        # Get direct download link (add force download parameter)
        file_link = f"https://drive.google.com/uc?export=download&id={file['id']}"
        
        return file_link, file['id']
    
    except HttpError as error:
        st.error(f"Google Drive API error: {error}")
        return None, None
    except Exception as e:
        st.error(f"Error uploading to Drive: {str(e)}")
        return None, None

def process_pdf(pdf_file):
    """Extract UAN -> page mapping from PDF"""
    uan_pages = {}
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                match = re.search(r"UAN/MEMBER ID:\s*([^/\s]+)", text)
                if match:
                    uan = match.group(1).strip()
                    uan_pages[uan] = page_num
        return uan_pages
    except Exception as e:
        st.error(f"Error processing PDF: {str(e)}")
        return {}

def list_drive_contents(service, shared_drive_id, folder_id=None):
    """List all files and folders in the shared drive or specific folder"""
    try:
        query = f"trashed=false"
        if folder_id:
            query += f" and '{folder_id}' in parents"
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

def delete_file_or_folder(service, file_id):
    """Delete a file or folder from Google Drive"""
    try:
        service.files().delete(
            fileId=file_id,
            supportsAllDrives=True
        ).execute()
        return True
    except Exception as e:
        st.error(f"Error deleting file: {str(e)}")
        return False

def get_folder_tree(service, shared_drive_id, parent_id=None, level=0):
    """Get hierarchical folder structure"""
    folders = []
    try:
        query = f"mimeType='application/vnd.google-apps.folder' and trashed=false"
        if parent_id:
            query += f" and '{parent_id}' in parents"
        else:
            query += f" and '{shared_drive_id}' in parents"
        
        results = service.files().list(
            q=query,
            fields="files(id, name, mimeType)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            corpora='drive',
            driveId=shared_drive_id
        ).execute()
        
        for folder in results.get('files', []):
            folders.append({
                'id': folder['id'],
                'name': folder['name'],
                'level': level,
                'children': get_folder_tree(service, shared_drive_id, folder['id'], level + 1)
            })
        
        return folders
    except Exception as e:
        st.error(f"Error getting folder tree: {str(e)}")
        return []

# ---------------------------------------------------
# Streamlit App
# ---------------------------------------------------
st.set_page_config(
    page_title="Payslip Distributor", 
    layout="wide",
    page_icon="📄"
)

# Custom CSS for styling
st.markdown("""
    <style>
    .main-header {
        font-size: 3rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #2c9f2c;
        border-bottom: 2px solid #ff7f0e;
        padding-bottom: 0.5rem;
        margin-top: 1.5rem;
    }
    .success-box {
        background-color: #d4edda;
        color: #155724;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #28a745;
        margin-bottom: 15px;
    }
    .info-box {
        background-color: #d1ecf1;
        color: #0c5460;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #17a2b8;
        margin-bottom: 15px;
    }
    .warning-box {
        background-color: #fff3cd;
        color: #856404;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #ffc107;
        margin-bottom: 15px;
    }
    .error-box {
        background-color: #f8d7da;
        color: #721c24;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #dc3545;
        margin-bottom: 15px;
    }
    .stButton>button {
        background: linear-gradient(to right, #4CAF50, #2E8B57);
        color: white;
        border-radius: 8px;
        border: none;
        padding: 12px 28px;
        font-weight: bold;
        transition: all 0.3s;
        width: 100%;
    }
    .stButton>button:hover {
        background: linear-gradient(to right, #45a049, #26734d);
        box-shadow: 0 6px 12px 0 rgba(0,0,0,0.2);
        transform: translateY(-2px);
    }
    .reset-button>button {
        background: linear-gradient(to right, #ff7f0e, #e67300);
    }
    .reset-button>button:hover {
        background: linear-gradient(to right, #e67300, #cc6600);
    }
    .danger-button>button {
        background: linear-gradient(to right, #dc3545, #c82333);
    }
    .danger-button>button:hover {
        background: linear-gradient(to right, #c82333, #a71e2a);
    }
    .card {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 8px 0 rgba(0,0,0,0.1);
        margin-bottom: 20px;
        border: 1px solid #e9ecef;
    }
    .footer {
        text-align: center;
        margin-top: 30px;
        color: #6c757d;
        font-size: 0.9rem;
    }
    .log-entry {
        padding: 10px;
        margin: 5px 0;
        border-radius: 5px;
        border-left: 4px solid;
    }
    .file-item {
        padding: 10px;
        margin: 5px 0;
        border-radius: 5px;
        background-color: white;
        border: 1px solid #dee2e6;
    }
    .folder-item {
        background-color: #e8f4f8;
        border-color: #bee5eb;
    }
    </style>
    """, unsafe_allow_html=True)

# Initialize session state
if 'sent_numbers' not in st.session_state:
    st.session_state.sent_numbers = {}
if 'processing' not in st.session_state:
    st.session_state.processing = False
if 'log_entries' not in st.session_state:
    st.session_state.log_entries = []
if 'drive_service' not in st.session_state:
    st.session_state.drive_service = None
if 'current_folder' not in st.session_state:
    st.session_state.current_folder = None
if 'shared_drive_id' not in st.session_state:
    st.session_state.shared_drive_id = os.getenv("SHARED_DRIVE_ID")

# Header
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.markdown('<h1 class="main-header">📄 Payslip Distribution System</h1>', unsafe_allow_html=True)
    st.markdown("### Fast, Secure & Efficient Employee Payslip Distribution")

# Create tabs
tab1, tab2, tab3 = st.tabs(["📤 Send Payslips", "📁 Drive Explorer", "⚙️ Settings"])

with tab1:
    # ---------------------------------------------------
    # Google Drive Setup Section
    # ---------------------------------------------------
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">🔐 Google Drive Setup</p>', unsafe_allow_html=True)

    if st.button("Initialize Google Drive Connection"):
        with st.spinner("Connecting to Google Drive..."):
            service = get_drive_service()
            if service:
                st.session_state.drive_service = service
                st.success("✅ Successfully connected to Google Drive!")
            else:
                st.error("❌ Failed to connect to Google Drive. Please check your .env file.")

    if st.session_state.drive_service:
        st.info("✅ Google Drive is ready for uploads")
    else:
        st.warning("⚠️ Please initialize Google Drive connection first")

    st.markdown('</div>', unsafe_allow_html=True)

    # ---------------------------------------------------
    # File Upload Section
    # ---------------------------------------------------
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">📤 Upload Files</p>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        uploaded_pdf = st.file_uploader("**PDF Payslip File**", type="pdf", help="Upload the PDF containing all employee payslips")
    with col2:
        uploaded_excel = st.file_uploader("**Employee Contacts File**", type="xlsx", help="Upload Excel file with UAN and phone numbers")

    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_pdf and uploaded_excel and st.session_state.drive_service:
        with st.spinner("🔍 Processing files... This may take a moment"):
            uan_pages = process_pdf(uploaded_pdf)
            df = pd.read_excel(uploaded_excel)
            df['Formatted_Phone'] = df['Employee no'].apply(format_phone_number)

        # Display file processing results
        col1, col2 = st.columns(2)
        with col1:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown(f"<h3 style='color: #1f77b4;'>📊 PDF Analysis</h3>", unsafe_allow_html=True)
            st.markdown(f"<h2 style='color: #ff7f0e; text-align: center;'>{len(uan_pages)}</h2>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center;'>UANs found in PDF</p>", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown(f"<h3 style='color: #1f77b4;'>👥 Employee Data</h3>", unsafe_allow_html=True)
            st.markdown(f"<h2 style='color: #2c9f2c; text-align: center;'>{len(df)}</h2>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center;'>Employees in Excel</p>", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # ---------------------------------------------------
        # Send Payslips Section
        # ---------------------------------------------------
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<p class="sub-header">🚀 Send Payslips</p>', unsafe_allow_html=True)
        
        # Get current month-year for folder name
        current_month = datetime.now().strftime("%B %Y")
        
        # Add a confirmation checkbox to prevent accidental sends
        confirm_send = st.checkbox("I confirm I want to send payslips to all employees")
        
        if confirm_send and st.button("📤 Send Payslip Links via WhatsApp", use_container_width=True, disabled=st.session_state.processing):
            st.session_state.processing = True
            sent_count, failed_count, skipped_count = 0, 0, 0
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            results_placeholder = st.empty()
            
            # Create monthly folder
            with st.spinner(f"Creating folder for {current_month}..."):
                folder_id = get_monthly_folder_id(st.session_state.drive_service, current_month, st.session_state.shared_drive_id)
            
            if not folder_id:
                st.error("Failed to create monthly folder. Aborting.")
                st.session_state.processing = False
                st.stop()
            
            # Process each employee
            for idx, (_, row) in enumerate(df.iterrows()):
                phone = row['Formatted_Phone']
                uan = str(row['UAN/member ID'])
                
                # Update progress
                progress = (idx + 1) / len(df)
                progress_bar.progress(progress)
                status_text.text(f"📱 Processing {idx + 1} of {len(df)}: {phone}")
                
                # Skip if already sent in this session
                if phone in st.session_state.sent_numbers:
                    log_entry = f"⏭️ Already sent to {phone} (UAN: {uan})"
                    st.session_state.log_entries.append(log_entry)
                    skipped_count += 1
                    continue
                
                if uan in uan_pages:
                    try:
                        # Extract individual payslip
                        uploaded_pdf.seek(0)  # Reset file pointer
                        pdf_buffer = extract_individual_payslip(uploaded_pdf, uan, uan_pages[uan])
                        
                        # Upload to Google Drive
                        filename = f"Payslip_{uan}_{current_month.replace(' ', '_')}.pdf"
                        drive_link, file_id = upload_to_drive(
                            st.session_state.drive_service,
                            pdf_buffer,
                            filename,
                            folder_id,
                            st.session_state.shared_drive_id
                        )
                        
                        if drive_link:
                            # Send WhatsApp message with download link
                            message = f"Hello! Your payslip for {current_month} is ready.\n\n📄 Download Link: {drive_link}\n\nThis link will expire in 30 days."
                            
                            # Using pywhatkit to send message
                            import pywhatkit
                            pywhatkit.sendwhatmsg_instantly(
                                phone_no=phone,
                                message=message,
                                wait_time=15,
                                tab_close=False
                            )
                            
                            # Mark as sent with timestamp
                            st.session_state.sent_numbers[phone] = {
                                'uan': uan,
                                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                'drive_link': drive_link,
                                'file_id': file_id
                            }
                            
                            log_entry = f"✅ Successfully sent to {phone} (UAN: {uan}) - Link: {drive_link}"
                            st.session_state.log_entries.append(log_entry)
                            sent_count += 1
                            
                            # Wait to prevent rate limiting
                            time.sleep(10)
                        else:
                            log_entry = f"❌ Failed to upload payslip for UAN {uan}"
                            st.session_state.log_entries.append(log_entry)
                            failed_count += 1
                        
                    except Exception as e:
                        log_entry = f"❌ Failed to send to {phone}: {str(e)}"
                        st.session_state.log_entries.append(log_entry)
                        failed_count += 1
                else:
                    log_entry = f"⚠️ No payslip found for UAN {uan} ({phone})"
                    st.session_state.log_entries.append(log_entry)
                    failed_count += 1
            
            # Clear progress indicators
            progress_bar.empty()
            status_text.empty()
            st.session_state.processing = False
            
            # Display results
            with results_placeholder.container():
                st.markdown('<p class="sub-header">📊 Results</p>', unsafe_allow_html=True)
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f'<div class="success-box"><h3>✅ Sent: {sent_count}</h3></div>', unsafe_allow_html=True)
                with col2:
                    st.markdown(f'<div class="error-box"><h3>❌ Failed: {failed_count}</h3></div>', unsafe_allow_html=True)
                with col3:
                    st.markdown(f'<div class="info-box"><h3>⏭️ Skipped: {skipped_count}</h3></div>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)
        
        # Display log entries
        if st.session_state.log_entries:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<p class="sub-header">📜 Activity Log</p>', unsafe_allow_html=True)
            
            # Show only the last 20 log entries to avoid clutter
            for log in st.session_state.log_entries[-20:]:
                if log.startswith("✅"):
                    st.markdown(f'<div class="log-entry" style="border-color: #28a745; background-color: #f8fff9;">{log}</div>', unsafe_allow_html=True)
                elif log.startswith("❌"):
                    st.markdown(f'<div class="log-entry" style="border-color: #dc3545; background-color: #fffafa;">{log}</div>', unsafe_allow_html=True)
                elif log.startswith("⚠️"):
                    st.markdown(f'<div class="log-entry" style="border-color: #ffc107; background-color: #fffef5;">{log}</div>', unsafe_allow_html=True)
                elif log.startswith("⏭️"):
                    st.markdown(f'<div class="log-entry" style="border-color: #17a2b8; background-color: #f8fdff;">{log}</div>', unsafe_allow_html=True)
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Reset button
        if st.session_state.sent_numbers:
            st.markdown("---")
            if st.button("🔄 Reset Sent History", key="reset", use_container_width=True):
                st.session_state.sent_numbers = {}
                st.session_state.log_entries = []
                st.success("✅ Sent history has been cleared!")
                st.rerun()

    elif uploaded_pdf and uploaded_excel and not st.session_state.drive_service:
        st.error("❌ Please initialize Google Drive connection first")
    else:
        st.markdown('<div class="info-box">Please upload both PDF and Excel files to continue</div>', unsafe_allow_html=True)

with tab2:
    # ---------------------------------------------------
    # Drive Explorer Section
    # ---------------------------------------------------
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">📁 Google Drive Explorer</p>', unsafe_allow_html=True)
    
    if not st.session_state.drive_service:
        st.warning("⚠️ Please initialize Google Drive connection first in the 'Send Payslips' tab")
    else:
        # Navigation controls
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            if st.button("🏠 Root Folder", use_container_width=True):
                st.session_state.current_folder = None
        with col2:
            if st.session_state.current_folder and st.button("⬅️ Go Back", use_container_width=True):
                # For simplicity, we'll just go back to root
                st.session_state.current_folder = None
        with col3:
            if st.button("🔄 Refresh", use_container_width=True):
                st.rerun()
        
        # List contents
        with st.spinner("Loading drive contents..."):
            contents = list_drive_contents(
                st.session_state.drive_service,
                st.session_state.shared_drive_id,
                st.session_state.current_folder
            )
        
        if contents:
            st.markdown(f"### 📂 Contents of {'Root' if not st.session_state.current_folder else 'Current Folder'}")
            
            # Separate folders and files
            folders = [item for item in contents if item['mimeType'] == 'application/vnd.google-apps.folder']
            files = [item for item in contents if item['mimeType'] != 'application/vnd.google-apps.folder']
            
            # Display folders
            if folders:
                st.markdown("#### 📁 Folders")
                for folder in folders:
                    col1, col2, col3 = st.columns([3, 1, 1])
                    with col1:
                        st.markdown(f"**{folder['name']}**")
                    with col2:
                        if st.button("📂 Open", key=f"open_{folder['id']}"):
                            st.session_state.current_folder = folder['id']
                    with col3:
                        if st.button("🗑️ Delete", key=f"del_{folder['id']}", type="secondary"):
                            if delete_file_or_folder(st.session_state.drive_service, folder['id']):
                                st.success(f"Deleted folder: {folder['name']}")
                                st.rerun()
            
            # Display files
            if files:
                st.markdown("#### 📄 Files")
                for file in files:
                    col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
                    with col1:
                        st.markdown(f"**{file['name']}**")
                        if 'size' in file:
                            st.caption(f"Size: {int(file['size']) / 1024:.1f} KB")
                        if 'modifiedTime' in file:
                            st.caption(f"Modified: {file['modifiedTime'][:10]}")
                    with col2:
                        download_link = f"https://drive.google.com/uc?export=download&id={file['id']}"
                        st.markdown(f"[⬇️ Download]({download_link})")
                    with col3:
                        if 'webViewLink' in file:
                            st.markdown(f"[👀 View]({file['webViewLink']})")
                    with col4:
                        if st.button("🗑️ Delete", key=f"del_file_{file['id']}", type="secondary"):
                            if delete_file_or_folder(st.session_state.drive_service, file['id']):
                                st.success(f"Deleted file: {file['name']}")
                                st.rerun()
            
            # Show empty state if no contents
            if not folders and not files:
                st.info("This folder is empty")
        
        else:
            st.info("No files or folders found in the shared drive")
    
    st.markdown('</div>', unsafe_allow_html=True)

with tab3:
    # ---------------------------------------------------
    # Settings Section
    # ---------------------------------------------------
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">⚙️ System Settings</p>', unsafe_allow_html=True)
    
    # Shared Drive ID configuration
    shared_drive_id = st.text_input(
        "Shared Drive ID",
        value=st.session_state.shared_drive_id or "0AMGgmW1LqUijUk9PVA",
        help="Enter the ID of your Google Shared Drive"
    )
        
    if st.button("💾 Save Settings"):
        st.session_state.shared_drive_id = shared_drive_id
        st.success("Settings saved successfully!")
    
    # System information
    st.markdown("### 📊 System Status")
    col1, col2 = st.columns(2)
    with col1:
        st.info(f"**Google Drive:** {'✅ Connected' if st.session_state.drive_service else '❌ Disconnected'}")
    with col2:
        st.info(f"**Shared Drive ID:** {st.session_state.shared_drive_id or 'Not set'}")
    
    # Danger zone
    st.markdown("### ⚠️ Danger Zone")
    if st.button("🗑️ Clear All Session Data", type="secondary"):
        st.session_state.sent_numbers = {}
        st.session_state.log_entries = []
        st.session_state.current_folder = None
        st.success("Session data cleared!")
        st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------------------------------
# Footer
# ---------------------------------------------------
st.markdown("---")
st.markdown("""
    <div class="footer">
        <p>Payslip Distribution System • Uses Google Drive & WhatsApp • Secure PDF Links</p>
        <p>© 2025 HR Management System</p>
    </div>
    """, unsafe_allow_html=True)
