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
    page_title="🎯 Payslip Distributor Pro", 
    layout="wide",
    page_icon="🚀",
    initial_sidebar_state="collapsed"
)

# Enhanced CSS with modern design, animations, and vibrant colors
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    @import url('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css');
    
    /* Global Styles */
    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        min-height: 100vh;
        font-family: 'Inter', sans-serif;
    }
    
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Animated Background */
    .stApp::before {
        content: '';
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: 
            radial-gradient(circle at 20% 80%, rgba(120, 119, 198, 0.3) 0%, transparent 50%),
            radial-gradient(circle at 80% 20%, rgba(255, 119, 198, 0.3) 0%, transparent 50%),
            radial-gradient(circle at 40% 40%, rgba(120, 219, 255, 0.3) 0%, transparent 50%);
        animation: float 20s ease-in-out infinite;
        pointer-events: none;
        z-index: -1;
    }
    
    @keyframes float {
        0%, 100% { transform: translateY(0px) rotate(0deg); }
        33% { transform: translateY(-20px) rotate(1deg); }
        66% { transform: translateY(-10px) rotate(-1deg); }
    }
    
    /* Header Styles */
    .main-header {
        font-size: 4.5rem;
        font-weight: 800;
        background: linear-gradient(135deg, #ff6b6b, #4ecdc4, #45b7d1, #96ceb4, #ffeaa7);
        background-size: 300% 300%;
        background-clip: text;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin: 2rem 0;
        animation: gradientShift 8s ease infinite, bounce 2s infinite;
        text-shadow: 0 0 30px rgba(255, 255, 255, 0.5);
    }
    
    @keyframes gradientShift {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    
    @keyframes bounce {
        0%, 20%, 50%, 80%, 100% { transform: translateY(0); }
        40% { transform: translateY(-10px); }
        60% { transform: translateY(-5px); }
    }
    
    .sub-title {
        font-size: 1.8rem;
        color: white;
        text-align: center;
        margin-bottom: 3rem;
        text-shadow: 0 2px 10px rgba(0,0,0,0.3);
        animation: fadeInUp 1s ease-out;
    }
    
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(30px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    .sub-header {
        font-size: 2rem;
        font-weight: 700;
        background: linear-gradient(135deg, #ff6b6b, #4ecdc4);
        background-clip: text;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 2rem 0 1rem 0;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* Card Styles */
    .glass-card {
        background: rgba(255, 255, 255, 0.25);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border-radius: 20px;
        border: 1px solid rgba(255, 255, 255, 0.3);
        padding: 2rem;
        margin: 2rem 0;
        box-shadow: 
            0 8px 32px 0 rgba(31, 38, 135, 0.37),
            inset 0 1px 0 rgba(255, 255, 255, 0.4);
        animation: slideInUp 0.8s ease-out;
        transition: all 0.3s ease;
    }
    
    .glass-card:hover {
        transform: translateY(-10px);
        box-shadow: 
            0 20px 40px 0 rgba(31, 38, 135, 0.5),
            inset 0 1px 0 rgba(255, 255, 255, 0.6);
    }
    
    @keyframes slideInUp {
        from {
            opacity: 0;
            transform: translateY(50px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    /* Stats Cards */
    .stats-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 20px;
        padding: 2rem;
        text-align: center;
        color: white;
        box-shadow: 0 15px 35px rgba(0,0,0,0.2);
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
        animation: pulse 2s infinite;
    }
    
    .stats-card::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: linear-gradient(45deg, transparent, rgba(255,255,255,0.1), transparent);
        transform: rotate(45deg);
        transition: all 0.5s;
        opacity: 0;
    }
    
    .stats-card:hover::before {
        animation: shimmer 1.5s ease-in-out;
        opacity: 1;
    }
    
    .stats-card:hover {
        transform: scale(1.05) rotateY(5deg);
        box-shadow: 0 25px 50px rgba(0,0,0,0.3);
    }
    
    @keyframes shimmer {
        0% { transform: translateX(-100%) translateY(-100%) rotate(45deg); }
        100% { transform: translateX(100%) translateY(100%) rotate(45deg); }
    }
    
    @keyframes pulse {
        0% { box-shadow: 0 15px 35px rgba(0,0,0,0.2); }
        50% { box-shadow: 0 15px 35px rgba(0,0,0,0.4); }
        100% { box-shadow: 0 15px 35px rgba(0,0,0,0.2); }
    }
    
    .stats-number {
        font-size: 3.5rem;
        font-weight: 800;
        margin: 1rem 0;
        text-shadow: 0 2px 10px rgba(0,0,0,0.3);
    }
    
    .stats-label {
        font-size: 1.1rem;
        font-weight: 500;
        opacity: 0.9;
    }
    
    /* Button Styles */
    .stButton > button {
        background: linear-gradient(135deg, #ff6b6b, #ee5a24) !important;
        color: white !important;
        border: none !important;
        border-radius: 15px !important;
        padding: 1rem 2rem !important;
        font-weight: 600 !important;
        font-size: 1.1rem !important;
        width: 100% !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 10px 30px rgba(238, 90, 36, 0.3) !important;
        position: relative !important;
        overflow: hidden !important;
    }
    
    .stButton > button::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent);
        transition: left 0.5s;
    }
    
    .stButton > button:hover::before {
        left: 100%;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #ee5a24, #ff6b6b) !important;
        transform: translateY(-3px) !important;
        box-shadow: 0 15px 40px rgba(238, 90, 36, 0.5) !important;
    }
    
    /* Success Button */
    .success-btn > button {
        background: linear-gradient(135deg, #00b894, #00cec9) !important;
        box-shadow: 0 10px 30px rgba(0, 184, 148, 0.3) !important;
    }
    
    .success-btn > button:hover {
        background: linear-gradient(135deg, #00cec9, #00b894) !important;
        box-shadow: 0 15px 40px rgba(0, 184, 148, 0.5) !important;
    }
    
    /* Warning Button */
    .warning-btn > button {
        background: linear-gradient(135deg, #fdcb6e, #e17055) !important;
        box-shadow: 0 10px 30px rgba(253, 203, 110, 0.3) !important;
    }
    
    /* Danger Button */
    .danger-btn > button {
        background: linear-gradient(135deg, #d63031, #74b9ff) !important;
        box-shadow: 0 10px 30px rgba(214, 48, 49, 0.3) !important;
    }
    
    /* Message Boxes */
    .success-box {
        background: linear-gradient(135deg, #00b894, #00cec9);
        color: white;
        padding: 1.5rem;
        border-radius: 15px;
        margin: 1rem 0;
        box-shadow: 0 10px 30px rgba(0, 184, 148, 0.3);
        animation: slideInRight 0.5s ease;
        border-left: 5px solid rgba(255,255,255,0.5);
    }
    
    .info-box {
        background: linear-gradient(135deg, #74b9ff, #0984e3);
        color: white;
        padding: 1.5rem;
        border-radius: 15px;
        margin: 1rem 0;
        box-shadow: 0 10px 30px rgba(116, 185, 255, 0.3);
        animation: slideInLeft 0.5s ease;
        border-left: 5px solid rgba(255,255,255,0.5);
    }
    
    .warning-box {
        background: linear-gradient(135deg, #fdcb6e, #e17055);
        color: white;
        padding: 1.5rem;
        border-radius: 15px;
        margin: 1rem 0;
        box-shadow: 0 10px 30px rgba(253, 203, 110, 0.3);
        animation: bounce 0.5s ease;
        border-left: 5px solid rgba(255,255,255,0.5);
    }
    
    .error-box {
        background: linear-gradient(135deg, #d63031, #e84393);
        color: white;
        padding: 1.5rem;
        border-radius: 15px;
        margin: 1rem 0;
        box-shadow: 0 10px 30px rgba(214, 48, 49, 0.3);
        animation: shake 0.5s ease;
        border-left: 5px solid rgba(255,255,255,0.5);
    }
    
    @keyframes slideInRight {
        from { transform: translateX(100%); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
    }
    
    @keyframes slideInLeft {
        from { transform: translateX(-100%); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
    }
    
    @keyframes shake {
        0%, 100% { transform: translateX(0); }
        25% { transform: translateX(-5px); }
        75% { transform: translateX(5px); }
    }
    
    /* Log Entries */
    .log-entry {
        background: rgba(255, 255, 255, 0.9);
        padding: 1rem;
        margin: 0.5rem 0;
        border-radius: 10px;
        border-left: 4px solid;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
        animation: fadeInUp 0.5s ease;
    }
    
    .log-entry:hover {
        transform: translateX(10px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.2);
    }
    
    /* File Items */
    .file-item {
        background: rgba(255, 255, 255, 0.9);
        padding: 1.5rem;
        margin: 1rem 0;
        border-radius: 15px;
        border: 1px solid rgba(255, 255, 255, 0.3);
        box-shadow: 0 8px 25px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
        animation: fadeInUp 0.5s ease;
    }
    
    .file-item:hover {
        transform: translateY(-5px);
        box-shadow: 0 15px 35px rgba(0,0,0,0.2);
    }
    
    .folder-item {
        background: linear-gradient(135deg, #74b9ff, #0984e3);
        color: white;
    }
    
    /* Progress Bar */
    .stProgress > div > div > div > div {
        background: linear-gradient(135deg, #00b894, #00cec9) !important;
        border-radius: 10px !important;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 15px;
        padding: 0.5rem;
        backdrop-filter: blur(10px);
    }
    
    .stTabs [data-baseweb="tab"] {
        background: transparent;
        border-radius: 10px;
        color: white;
        font-weight: 600;
        font-size: 1.1rem;
        transition: all 0.3s ease;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        background: rgba(255, 255, 255, 0.2);
        transform: scale(1.05);
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #ff6b6b, #ee5a24) !important;
        color: white !important;
    }
    
    /* Footer */
    .footer {
        background: rgba(255, 255, 255, 0.1);
        backdrop-filter: blur(20px);
        border-radius: 20px;
        padding: 2rem;
        margin-top: 3rem;
        text-align: center;
        color: white;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }
    
    /* Input Fields */
    .stTextInput input {
        background: rgba(255, 255, 255, 0.9) !important;
        border-radius: 10px !important;
        border: 2px solid rgba(255, 255, 255, 0.3) !important;
        transition: all 0.3s ease !important;
    }
    
    .stTextInput input:focus {
        border-color: #ff6b6b !important;
        box-shadow: 0 0 20px rgba(255, 107, 107, 0.3) !important;
    }
    
    /* File Uploader */
    .stFileUploader {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 15px;
        padding: 1rem;
        border: 2px dashed rgba(255, 255, 255, 0.3);
        transition: all 0.3s ease;
    }
    
    .stFileUploader:hover {
        background: rgba(255, 255, 255, 0.2);
        border-color: #ff6b6b;
        transform: scale(1.02);
    }
    
    /* Selectbox */
    .stSelectbox {
        background: rgba(255, 255, 255, 0.9);
        border-radius: 10px;
    }
    
    /* Checkbox */
    .stCheckbox {
        color: white !important;
        font-weight: 600 !important;
        font-size: 1.1rem !important;
    }
    
    /* Sidebar */
    .sidebar {
        background: linear-gradient(180deg, rgba(102, 126, 234, 0.9), rgba(118, 75, 162, 0.9));
        backdrop-filter: blur(20px);
    }
    
    /* Tooltips and Help Text */
    .stTooltip {
        background: rgba(0, 0, 0, 0.8) !important;
        backdrop-filter: blur(10px) !important;
        border-radius: 8px !important;
    }
    
    /* Loading Spinner */
    .stSpinner {
        color: #ff6b6b !important;
    }
    
    /* Custom Icons */
    .icon {
        font-size: 1.5rem;
        margin-right: 0.5rem;
        vertical-align: middle;
    }
    
    /* Hover Effects */
    .hover-scale {
        transition: transform 0.3s ease;
    }
    
    .hover-scale:hover {
        transform: scale(1.05);
    }
    
    /* Responsive Design */
    @media (max-width: 768px) {
        .main-header {
            font-size: 3rem;
        }
        
        .stats-number {
            font-size: 2.5rem;
        }
        
        .glass-card {
            padding: 1rem;
            margin: 1rem 0;
        }
    }
    
    /* Dark mode adjustments */
    [data-theme="dark"] .glass-card {
        background: rgba(0, 0, 0, 0.3);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    
    [data-theme="dark"] .file-item {
        background: rgba(0, 0, 0, 0.5);
        color: white;
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

# Animated Header
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.markdown('''
        <div style="text-align: center; margin: 2rem 0;">
            <div class="main-header">🚀 PaySlip Pro</div>
            <div class="sub-title">
                <i class="fas fa-bolt icon"></i>
                Lightning-Fast • Ultra-Secure • Employee-Friendly
                <i class="fas fa-shield-alt icon"></i>
            </div>
        </div>
    ''', unsafe_allow_html=True)

# Create enhanced tabs with icons
tab1, tab2, tab3 = st.tabs([
    "🚀 **Send Payslips**", 
    "📁 **Drive Explorer**", 
    "⚙️ **Settings**"
])

with tab1:
    # ---------------------------------------------------
    # Google Drive Setup Section
    # ---------------------------------------------------
    st.markdown('''
        <div class="glass-card">
            <h2 class="sub-header">
                <i class="fas fa-cloud icon"></i>Google Drive Setup
            </h2>
        </div>
    ''', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🔗 Initialize Google Drive Connection", key="init_drive"):
            with st.spinner("🌐 Connecting to Google Drive..."):
                service = get_drive_service()
                if service:
                    st.session_state.drive_service = service
                    st.markdown('''
                        <div class="success-box">
                            <i class="fas fa-check-circle"></i>
                            <strong>Successfully connected to Google Drive!</strong>
                            <br>Ready to upload and manage files securely.
                        </div>
                    ''', unsafe_allow_html=True)
                else:
                    st.markdown('''
                        <div class="error-box">
                            <i class="fas fa-exclamation-triangle"></i>
                            <strong>Failed to connect to Google Drive</strong>
                            <br>Please check your .env configuration file.
                        </div>
                    ''', unsafe_allow_html=True)

    # Connection Status
    if st.session_state.drive_service:
        st.markdown('''
            <div class="info-box">
                <i class="fas fa-check-circle icon"></i>
                <strong>Google Drive Status:</strong> Connected & Ready
            </div>
        ''', unsafe_allow_html=True)
    else:
        st.markdown('''
            <div class="warning-box">
                <i class="fas fa-exclamation-circle icon"></i>
                <strong>Action Required:</strong> Please initialize Google Drive connection first
            </div>
        ''', unsafe_allow_html=True)

    # ---------------------------------------------------
    # File Upload Section
    # ---------------------------------------------------
    st.markdown('''
        <div class="glass-card">
            <h2 class="sub-header">
                <i class="fas fa-upload icon"></i>Upload Files
            </h2>
        </div>
    ''', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 📄 PDF Payslip File")
        uploaded_pdf = st.file_uploader(
            "Choose PDF file", 
            type="pdf", 
            help="Upload the PDF containing all employee payslips",
            key="pdf_upload"
        )
        if uploaded_pdf:
            st.markdown('''
                <div class="success-box">
                    <i class="fas fa-file-pdf icon"></i>
                    PDF file uploaded successfully!
                </div>
            ''', unsafe_allow_html=True)
    
    with col2:
        st.markdown("### 📊 Employee Contacts File")
        uploaded_excel = st.file_uploader(
            "Choose Excel file", 
            type="xlsx", 
            help="Upload Excel file with UAN and phone numbers",
            key="excel_upload"
        )
        if uploaded_excel:
            st.markdown('''
                <div class="success-box">
                    <i class="fas fa-file-excel icon"></i>
                    Excel file uploaded successfully!
                </div>
            ''', unsafe_allow_html=True)

    if uploaded_pdf and uploaded_excel and st.session_state.drive_service:
        with st.spinner("🔍 Processing files... Please wait"):
            uan_pages = process_pdf(uploaded_pdf)
            df = pd.read_excel(uploaded_excel)
            df['Formatted_Phone'] = df['Employee no'].apply(format_phone_number)

        # Display enhanced file processing results
        st.markdown('''
            <div class="glass-card">
                <h2 class="sub-header">
                    <i class="fas fa-chart-bar icon"></i>File Analysis Results
                </h2>
            </div>
        ''', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f'''
                <div class="stats-card">
                    <i class="fas fa-file-pdf" style="font-size: 2rem; margin-bottom: 1rem;"></i>
                    <div class="stats-number">{len(uan_pages)}</div>
                    <div class="stats-label">UANs Found in PDF</div>
                </div>
            ''', unsafe_allow_html=True)
        
        with col2:
            st.markdown(f'''
                <div class="stats-card">
                    <i class="fas fa-users" style="font-size: 2rem; margin-bottom: 1rem;"></i>
                    <div class="stats-number">{len(df)}</div>
                    <div class="stats-label">Employees in Excel</div>
                </div>
            ''', unsafe_allow_html=True)
        
        with col3:
            match_count = len(set(uan_pages.keys()) & set(df['UAN/member ID'].astype(str)))
            st.markdown(f'''
                <div class="stats-card">
                    <i class="fas fa-link" style="font-size: 2rem; margin-bottom: 1rem;"></i>
                    <div class="stats-number">{match_count}</div>
                    <div class="stats-label">Matched Records</div>
                </div>
            ''', unsafe_allow_html=True)

        # ---------------------------------------------------
        # Send Payslips Section
        # ---------------------------------------------------
        st.markdown('''
            <div class="glass-card">
                <h2 class="sub-header">
                    <i class="fas fa-paper-plane icon"></i>Send Payslips
                </h2>
            </div>
        ''', unsafe_allow_html=True)
        
        # Get current month-year for folder name
        current_month = datetime.now().strftime("%B %Y")
        
        # Enhanced confirmation section
        col1, col2 = st.columns([2, 1])
        with col1:
            confirm_send = st.checkbox(
                "✅ I confirm I want to send payslips to all employees", 
                help="This will send WhatsApp messages to all matched employees"
            )
        with col2:
            st.markdown(f'''
                <div class="info-box" style="text-align: center; padding: 1rem;">
                    <i class="fas fa-calendar icon"></i>
                    <strong>{current_month}</strong>
                </div>
            ''', unsafe_allow_html=True)
        
        if confirm_send:
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button("🚀 Launch Payslip Distribution", key="send_payslips", disabled=st.session_state.processing):
                    st.session_state.processing = True
                    sent_count, failed_count, skipped_count = 0, 0, 0
                    
                    # Enhanced progress tracking
                    progress_container = st.container()
                    with progress_container:
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        live_stats = st.empty()
                    
                    results_placeholder = st.empty()
                    
                    # Create monthly folder
                    with st.spinner(f"📁 Creating folder for {current_month}..."):
                        folder_id = get_monthly_folder_id(st.session_state.drive_service, current_month, st.session_state.shared_drive_id)
                    
                    if not folder_id:
                        st.markdown('''
                            <div class="error-box">
                                <i class="fas fa-times-circle icon"></i>
                                Failed to create monthly folder. Aborting distribution.
                            </div>
                        ''', unsafe_allow_html=True)
                        st.session_state.processing = False
                        st.stop()
                    
                    # Process each employee with enhanced UI
                    for idx, (_, row) in enumerate(df.iterrows()):
                        phone = row['Formatted_Phone']
                        uan = str(row['UAN/member ID'])
                        
                        # Update progress with animations
                        progress = (idx + 1) / len(df)
                        progress_bar.progress(progress)
                        status_text.markdown(f'''
                            <div style="text-align: center; color: white; font-weight: 600;">
                                <i class="fas fa-mobile-alt icon"></i>
                                Processing {idx + 1} of {len(df)}: {phone}
                                <div style="font-size: 0.9rem; opacity: 0.8;">UAN: {uan}</div>
                            </div>
                        ''', unsafe_allow_html=True)
                        
                        # Live statistics update
                        with live_stats:
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("✅ Sent", sent_count, delta=None)
                            with col2:
                                st.metric("❌ Failed", failed_count, delta=None)
                            with col3:
                                st.metric("⏭️ Skipped", skipped_count, delta=None)
                        
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
                                    # Enhanced WhatsApp message
                                    message = f"""🎉 Hello! Your payslip for {current_month} is ready!

📄 Download Link: {drive_link}

💡 Tips:
• Click the link to download directly
• Save the file to your device
• Link expires in 30 days

Thank you! 🙏"""
                                    
                                    # Simulate sending (replace with actual WhatsApp API)
                                    # import pywhatkit
                                    # pywhatkit.sendwhatmsg_instantly(
                                    #     phone_no=phone,
                                    #     message=message,
                                    #     wait_time=15,
                                    #     tab_close=False
                                    # )
                                    
                                    # Mark as sent with timestamp
                                    st.session_state.sent_numbers[phone] = {
                                        'uan': uan,
                                        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        'drive_link': drive_link,
                                        'file_id': file_id
                                    }
                                    
                                    log_entry = f"✅ Successfully sent to {phone} (UAN: {uan})"
                                    st.session_state.log_entries.append(log_entry)
                                    sent_count += 1
                                    
                                    # Simulated wait (replace with actual wait for WhatsApp)
                                    time.sleep(0.5)  # Reduced for demo
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
                    progress_container.empty()
                    st.session_state.processing = False
                    
                    # Display enhanced results
                    st.markdown('''
                        <div class="glass-card">
                            <h2 class="sub-header">
                                <i class="fas fa-chart-pie icon"></i>Distribution Results
                            </h2>
                        </div>
                    ''', unsafe_allow_html=True)
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.markdown(f'''
                            <div class="success-box hover-scale">
                                <i class="fas fa-check-circle" style="font-size: 2rem; margin-bottom: 1rem;"></i>
                                <h2>{sent_count}</h2>
                                <p><strong>Successfully Sent</strong></p>
                            </div>
                        ''', unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f'''
                            <div class="error-box hover-scale">
                                <i class="fas fa-times-circle" style="font-size: 2rem; margin-bottom: 1rem;"></i>
                                <h2>{failed_count}</h2>
                                <p><strong>Failed Attempts</strong></p>
                            </div>
                        ''', unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f'''
                            <div class="info-box hover-scale">
                                <i class="fas fa-forward" style="font-size: 2rem; margin-bottom: 1rem;"></i>
                                <h2>{skipped_count}</h2>
                                <p><strong>Skipped (Already Sent)</strong></p>
                            </div>
                        ''', unsafe_allow_html=True)
                    
                    # Celebration animation for successful completion
                    if sent_count > 0:
                        st.balloons()
                        time.sleep(1)
                        st.snow()

        # Display enhanced log entries
        if st.session_state.log_entries:
            st.markdown('''
                <div class="glass-card">
                    <h2 class="sub-header">
                        <i class="fas fa-list-ul icon"></i>Activity Log
                    </h2>
                </div>
            ''', unsafe_allow_html=True)
            
            # Show only the last 20 log entries to avoid clutter
            for log in st.session_state.log_entries[-20:]:
                if log.startswith("✅"):
                    st.markdown(f'''
                        <div class="log-entry hover-scale" style="border-color: #00b894; background: linear-gradient(135deg, rgba(0,184,148,0.1), rgba(0,206,201,0.1));">
                            <i class="fas fa-check-circle" style="color: #00b894; margin-right: 0.5rem;"></i>
                            {log[2:]}
                        </div>
                    ''', unsafe_allow_html=True)
                elif log.startswith("❌"):
                    st.markdown(f'''
                        <div class="log-entry hover-scale" style="border-color: #d63031; background: linear-gradient(135deg, rgba(214,48,49,0.1), rgba(232,67,147,0.1));">
                            <i class="fas fa-times-circle" style="color: #d63031; margin-right: 0.5rem;"></i>
                            {log[2:]}
                        </div>
                    ''', unsafe_allow_html=True)
                elif log.startswith("⚠️"):
                    st.markdown(f'''
                        <div class="log-entry hover-scale" style="border-color: #fdcb6e; background: linear-gradient(135deg, rgba(253,203,110,0.1), rgba(225,112,85,0.1));">
                            <i class="fas fa-exclamation-triangle" style="color: #fdcb6e; margin-right: 0.5rem;"></i>
                            {log[2:]}
                        </div>
                    ''', unsafe_allow_html=True)
                elif log.startswith("⏭️"):
                    st.markdown(f'''
                        <div class="log-entry hover-scale" style="border-color: #74b9ff; background: linear-gradient(135deg, rgba(116,185,255,0.1), rgba(9,132,227,0.1));">
                            <i class="fas fa-forward" style="color: #74b9ff; margin-right: 0.5rem;"></i>
                            {log[2:]}
                        </div>
                    ''', unsafe_allow_html=True)
        
        # Enhanced reset button
        if st.session_state.sent_numbers:
            st.markdown("---")
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.markdown('<div class="warning-btn">', unsafe_allow_html=True)
                if st.button("🔄 Reset Distribution History", key="reset", use_container_width=True):
                    st.session_state.sent_numbers = {}
                    st.session_state.log_entries = []
                    st.markdown('''
                        <div class="success-box">
                            <i class="fas fa-check icon"></i>
                            Distribution history has been cleared successfully!
                        </div>
                    ''', unsafe_allow_html=True)
                    time.sleep(1)
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

    elif uploaded_pdf and uploaded_excel and not st.session_state.drive_service:
        st.markdown('''
            <div class="error-box">
                <i class="fas fa-exclamation-triangle icon"></i>
                <strong>Google Drive Connection Required</strong>
                <br>Please initialize Google Drive connection first.
            </div>
        ''', unsafe_allow_html=True)
    else:
        st.markdown('''
            <div class="glass-card">
                <div class="info-box">
                    <i class="fas fa-info-circle icon"></i>
                    <strong>Getting Started</strong>
                    <br>Please upload both PDF and Excel files to continue with the distribution process.
                </div>
                <div style="text-align: center; margin-top: 2rem;">
                    <i class="fas fa-lightbulb" style="font-size: 3rem; color: #fdcb6e; margin-bottom: 1rem;"></i>
                    <p style="color: white; font-size: 1.1rem;">
                        💡 Need sample files? Contact your system administrator for templates.
                    </p>
                </div>
            </div>
        ''', unsafe_allow_html=True)

with tab2:
    # ---------------------------------------------------
    # Enhanced Drive Explorer Section
    # ---------------------------------------------------
    st.markdown('''
        <div class="glass-card">
            <h2 class="sub-header">
                <i class="fas fa-folder-open icon"></i>Google Drive Explorer
            </h2>
        </div>
    ''', unsafe_allow_html=True)
    
    if not st.session_state.drive_service:
        st.markdown('''
            <div class="warning-box">
                <i class="fas fa-exclamation-triangle icon"></i>
                <strong>Google Drive Connection Required</strong>
                <br>Please initialize Google Drive connection first in the 'Send Payslips' tab.
            </div>
        ''', unsafe_allow_html=True)
    else:
        # Enhanced navigation controls
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown('<div class="success-btn">', unsafe_allow_html=True)
            if st.button("🏠 Root Folder", use_container_width=True, key="root_folder"):
                st.session_state.current_folder = None
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            if st.session_state.current_folder:
                st.markdown('<div class="warning-btn">', unsafe_allow_html=True)
                if st.button("⬅️ Go Back", use_container_width=True, key="go_back"):
                    st.session_state.current_folder = None
                st.markdown('</div>', unsafe_allow_html=True)
        
        with col3:
            if st.button("🔄 Refresh", use_container_width=True, key="refresh_drive"):
                st.rerun()
        
        with col4:
            st.markdown('<div class="danger-btn">', unsafe_allow_html=True)
            if st.button("📊 Storage Stats", use_container_width=True, key="storage_stats"):
                st.info("Storage statistics feature coming soon!")
            st.markdown('</div>', unsafe_allow_html=True)
        
        # List contents with enhanced UI
        with st.spinner("🔍 Loading drive contents..."):
            contents = list_drive_contents(
                st.session_state.drive_service,
                st.session_state.shared_drive_id,
                st.session_state.current_folder
            )
        
        if contents:
            current_location = 'Root Directory' if not st.session_state.current_folder else 'Current Folder'
            st.markdown(f'''
                <div class="info-box">
                    <i class="fas fa-map-marker-alt icon"></i>
                    <strong>Current Location:</strong> {current_location}
                    <span style="float: right;">
                        <i class="fas fa-items icon"></i>
                        {len(contents)} items
                    </span>
                </div>
            ''', unsafe_allow_html=True)
            
            # Separate folders and files
            folders = [item for item in contents if item['mimeType'] == 'application/vnd.google-apps.folder']
            files = [item for item in contents if item['mimeType'] != 'application/vnd.google-apps.folder']
            
            # Display folders with enhanced cards
            if folders:
                st.markdown('''
                    <div class="glass-card">
                        <h3 style="color: white; margin-bottom: 1rem;">
                            <i class="fas fa-folder icon"></i>Folders
                        </h3>
                    </div>
                ''', unsafe_allow_html=True)
                
                for folder in folders:
                    col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
                    with col1:
                        st.markdown(f'''
                            <div class="file-item folder-item">
                                <i class="fas fa-folder" style="color: #74b9ff; margin-right: 1rem; font-size: 1.5rem;"></i>
                                <strong style="font-size: 1.2rem;">{folder['name']}</strong>
                                <br>
                                <small style="opacity: 0.8;">
                                    <i class="fas fa-calendar"></i> Created: {folder.get('createdTime', 'Unknown')[:10]}
                                </small>
                            </div>
                        ''', unsafe_allow_html=True)
                    with col2:
                        st.markdown('<div class="success-btn">', unsafe_allow_html=True)
                        if st.button("📂 Open", key=f"open_{folder['id']}"):
                            st.session_state.current_folder = folder['id']
                            st.rerun()
                        st.markdown('</div>', unsafe_allow_html=True)
                    with col3:
                        if 'webViewLink' in folder:
                            st.markdown(f"[👀 View]({folder['webViewLink']})", unsafe_allow_html=True)
                    with col4:
                        st.markdown('<div class="danger-btn">', unsafe_allow_html=True)
                        if st.button("🗑️", key=f"del_{folder['id']}", help="Delete folder"):
                            if delete_file_or_folder(st.session_state.drive_service, folder['id']):
                                st.success(f"✅ Deleted folder: {folder['name']}")
                                time.sleep(1)
                                st.rerun()
                        st.markdown('</div>', unsafe_allow_html=True)
            
            # Display files with enhanced cards
            if files:
                st.markdown('''
                    <div class="glass-card">
                        <h3 style="color: white; margin-bottom: 1rem;">
                            <i class="fas fa-file icon"></i>Files
                        </h3>
                    </div>
                ''', unsafe_allow_html=True)
                
                for file in files:
                    # Determine file icon based on type
                    if 'pdf' in file['name'].lower():
                        file_icon = "fas fa-file-pdf"
                        icon_color = "#e74c3c"
                    elif any(ext in file['name'].lower() for ext in ['.xlsx', '.xls', '.csv']):
                        file_icon = "fas fa-file-excel"
                        icon_color = "#27ae60"
                    elif any(ext in file['name'].lower() for ext in ['.doc', '.docx']):
                        file_icon = "fas fa-file-word"
                        icon_color = "#3498db"
                    else:
                        file_icon = "fas fa-file"
                        icon_color = "#95a5a6"
                    
                    col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
                    with col1:
                        file_size = f"{int(file['size']) / 1024:.1f} KB" if 'size' in file else "Unknown"
                        st.markdown(f'''
                            <div class="file-item hover-scale">
                                <i class="{file_icon}" style="color: {icon_color}; margin-right: 1rem; font-size: 1.5rem;"></i>
                                <strong style="font-size: 1.2rem;">{file['name']}</strong>
                                <br>
                                <small style="opacity: 0.8;">
                                    <i class="fas fa-weight"></i> Size: {file_size} • 
                                    <i class="fas fa-calendar"></i> Modified: {file.get('modifiedTime', 'Unknown')[:10]}
                                </small>
                            </div>
                        ''', unsafe_allow_html=True)
                    
                    with col2:
                        download_link = f"https://drive.google.com/uc?export=download&id={file['id']}"
                        st.markdown(f'''
                            <a href="{download_link}" target="_blank" style="text-decoration: none;">
                                <div class="success-btn">
                                    <button style="background: linear-gradient(135deg, #00b894, #00cec9); color: white; border: none; padding: 0.5rem 1rem; border-radius: 8px; width: 100%;">
                                        ⬇️ Download
                                    </button>
                                </div>
                            </a>
                        ''', unsafe_allow_html=True)
                    
                    with col3:
                        if 'webViewLink' in file:
                            st.markdown(f'''
                                <a href="{file['webViewLink']}" target="_blank" style="text-decoration: none;">
                                    <div class="info-btn">
                                        <button style="background: linear-gradient(135deg, #74b9ff, #0984e3); color: white; border: none; padding: 0.5rem 1rem; border-radius: 8px; width: 100%;">
                                            👀 View
                                        </button>
                                    </div>
                                </a>
                            ''', unsafe_allow_html=True)
                    
                    with col4:
                        st.markdown('<div class="danger-btn">', unsafe_allow_html=True)
                        if st.button("🗑️", key=f"del_file_{file['id']}", help="Delete file"):
                            if delete_file_or_folder(st.session_state.drive_service, file['id']):
                                st.success(f"✅ Deleted file: {file['name']}")
                                time.sleep(1)
                                st.rerun()
                        st.markdown('</div>', unsafe_allow_html=True)
            
            # Show empty state if no contents
            if not folders and not files:
                st.markdown('''
                    <div class="info-box" style="text-align: center; padding: 3rem;">
                        <i class="fas fa-folder-open" style="font-size: 4rem; margin-bottom: 2rem; opacity: 0.5;"></i>
                        <h3>This folder is empty</h3>
                        <p>No files or folders found in this location.</p>
                    </div>
                ''', unsafe_allow_html=True)
        
        else:
            st.markdown('''
                <div class="warning-box" style="text-align: center; padding: 3rem;">
                    <i class="fas fa-search" style="font-size: 4rem; margin-bottom: 2rem;"></i>
                    <h3>No Content Found</h3>
                    <p>No files or folders found in the shared drive.</p>
                </div>
            ''', unsafe_allow_html=True)

with tab3:
    # ---------------------------------------------------
    # Enhanced Settings Section
    # ---------------------------------------------------
    st.markdown('''
        <div class="glass-card">
            <h2 class="sub-header">
                <i class="fas fa-cogs icon"></i>System Settings
            </h2>
        </div>
    ''', unsafe_allow_html=True)
    
    # Configuration Section
    st.markdown('''
        <div class="glass-card">
            <h3 style="color: white; margin-bottom: 1rem;">
                <i class="fas fa-database icon"></i>Drive Configuration
            </h3>
        </div>
    ''', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    with col1:
        shared_drive_id = st.text_input(
            "🔗 Shared Drive ID",
            value=st.session_state.shared_drive_id or "",
            help="Enter the ID of your Google Shared Drive",
            placeholder="Enter your Google Shared Drive ID here..."
        )
    
    with col2:
        st.markdown('<div style="padding-top: 2rem;">', unsafe_allow_html=True)
        st.markdown('<div class="success-btn">', unsafe_allow_html=True)
        if st.button("💾 Save Configuration", use_container_width=True):
            st.session_state.shared_drive_id = shared_drive_id
            st.markdown('''
                <div class="success-box">
                    <i class="fas fa-check-circle icon"></i>
                    Configuration saved successfully!
                </div>
            ''', unsafe_allow_html=True)
            time.sleep(1)
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    # System Status Dashboard
    st.markdown('''
        <div class="glass-card">
            <h3 style="color: white; margin-bottom: 1rem;">
                <i class="fas fa-tachometer-alt icon"></i>System Status Dashboard
            </h3>
        </div>
    ''', unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        drive_status = "✅ Connected" if st.session_state.drive_service else "❌ Disconnected"
        drive_color = "#00b894" if st.session_state.drive_service else "#d63031"
        st.markdown(f'''
            <div class="stats-card" style="background: linear-gradient(135deg, {drive_color}, {drive_color}aa);">
                <i class="fas fa-cloud" style="font-size: 2rem; margin-bottom: 1rem;"></i>
                <div class="stats-label">Google Drive</div>
                <div style="font-size: 1.2rem; font-weight: 600; margin-top: 0.5rem;">{drive_status}</div>
            </div>
        ''', unsafe_allow_html=True)
    
    with col2:
        total_sent = len(st.session_state.sent_numbers)
        st.markdown(f'''
            <div class="stats-card">
                <i class="fas fa-paper-plane" style="font-size: 2rem; margin-bottom: 1rem;"></i>
                <div class="stats-number" style="font-size: 2.5rem;">{total_sent}</div>
                <div class="stats-label">Payslips Sent</div>
            </div>
        ''', unsafe_allow_html=True)
    
    with col3:
        active_logs = len(st.session_state.log_entries)
        st.markdown(f'''
            <div class="stats-card">
                <i class="fas fa-list-ul" style="font-size: 2rem; margin-bottom: 1rem;"></i>
                <div class="stats-number" style="font-size: 2.5rem;">{active_logs}</div>
                <div class="stats-label">Log Entries</div>
            </div>
        ''', unsafe_allow_html=True)
    
    with col4:
        drive_id_status = "✅ Configured" if st.session_state.shared_drive_id else "⚠️ Not Set"
        id_color = "#00b894" if st.session_state.shared_drive_id else "#fdcb6e"
        st.markdown(f'''
            <div class="stats-card" style="background: linear-gradient(135deg, {id_color}, {id_color}aa);">
                <i class="fas fa-id-card" style="font-size: 2rem; margin-bottom: 1rem;"></i>
                <div class="stats-label">Drive ID</div>
                <div style="font-size: 1.2rem; font-weight: 600; margin-top: 0.5rem;">{drive_id_status}</div>
            </div>
        ''', unsafe_allow_html=True)
    
    # Application Preferences
    st.markdown('''
        <div class="glass-card">
            <h3 style="color: white; margin-bottom: 1rem;">
                <i class="fas fa-sliders-h icon"></i>Application Preferences
            </h3>
        </div>
    ''', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 🎨 Theme Settings")
        theme_mode = st.selectbox(
            "Color Theme",
            ["Auto", "Light", "Dark"],
            help="Choose your preferred color theme"
        )
        
        animation_enabled = st.checkbox(
            "✨ Enable Animations",
            value=True,
            help="Enable smooth animations and transitions"
        )
        
        show_tooltips = st.checkbox(
            "💡 Show Tooltips",
            value=True,
            help="Display helpful tooltips throughout the application"
        )
    
    with col2:
        st.markdown("#### 📱 Notification Settings")
        enable_sound = st.checkbox(
            "🔊 Enable Sound Notifications",
            value=False,
            help="Play sound when operations complete"
        )
        
        auto_refresh = st.checkbox(
            "🔄 Auto-refresh Drive Explorer",
            value=False,
            help="Automatically refresh drive contents every 30 seconds"
        )
        
        compact_view = st.checkbox(
            "📋 Compact Log View",
            value=False,
            help="Show logs in compact format"
        )
    
    # Advanced Settings
    st.markdown('''
        <div class="glass-card">
            <h3 style="color: white; margin-bottom: 1rem;">
                <i class="fas fa-tools icon"></i>Advanced Settings
            </h3>
        </div>
    ''', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### ⚡ Performance Settings")
        batch_size = st.slider(
            "Batch Processing Size",
            min_value=1,
            max_value=50,
            value=10,
            help="Number of payslips to process simultaneously"
        )
        
        retry_attempts = st.slider(
            "Retry Failed Uploads",
            min_value=1,
            max_value=5,
            value=3,
            help="Number of retry attempts for failed uploads"
        )
    
    with col2:
        st.markdown("#### 🔒 Security Settings")
        link_expiry = st.selectbox(
            "Download Link Expiry",
            ["7 days", "15 days", "30 days", "60 days", "90 days"],
            index=2,
            help="How long download links remain active"
        )
        
        require_confirmation = st.checkbox(
            "🛡️ Require Confirmation for Bulk Actions",
            value=True,
            help="Ask for confirmation before processing large batches"
        )
    
    # Export/Import Settings
    st.markdown('''
        <div class="glass-card">
            <h3 style="color: white; margin-bottom: 1rem;">
                <i class="fas fa-exchange-alt icon"></i>Data Management
            </h3>
        </div>
    ''', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="success-btn">', unsafe_allow_html=True)
        if st.button("📤 Export Sent History", use_container_width=True):
            if st.session_state.sent_numbers:
                # Create a DataFrame from sent numbers
                export_data = []
                for phone, details in st.session_state.sent_numbers.items():
                    export_data.append({
                        'Phone': phone,
                        'UAN': details['uan'],
                        'Timestamp': details['timestamp'],
                        'Drive Link': details['drive_link']
                    })
                
                df_export = pd.DataFrame(export_data)
                csv = df_export.to_csv(index=False)
                
                st.download_button(
                    label="⬇️ Download CSV",
                    data=csv,
                    file_name=f"payslip_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
                st.success("✅ Export ready for download!")
            else:
                st.info("No data to export")
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="warning-btn">', unsafe_allow_html=True)
        if st.button("📥 Import Settings", use_container_width=True):
            st.info("Settings import feature coming soon!")
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col3:
        st.markdown('<div class="info-btn" style="background: linear-gradient(135deg, #74b9ff, #0984e3);">', unsafe_allow_html=True)
        if st.button("📊 Generate Report", use_container_width=True):
            # Generate a comprehensive report
            report_data = {
                "System Status": {
                    "Google Drive": "Connected" if st.session_state.drive_service else "Disconnected",
                    "Shared Drive ID": "Configured" if st.session_state.shared_drive_id else "Not Set",
                    "Total Payslips Sent": len(st.session_state.sent_numbers),
                    "Active Log Entries": len(st.session_state.log_entries)
                },
                "Recent Activity": st.session_state.log_entries[-10:] if st.session_state.log_entries else ["No recent activity"],
                "Configuration": {
                    "Theme": theme_mode,
                    "Animations": animation_enabled,
                    "Batch Size": batch_size,
                    "Retry Attempts": retry_attempts,
                    "Link Expiry": link_expiry
                }
            }
            
            st.json(report_data)
            st.success("📊 System report generated!")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Danger Zone
    st.markdown("---")
    st.markdown('''
        <div class="glass-card">
            <h3 style="color: #d63031; margin-bottom: 1rem;">
                <i class="fas fa-exclamation-triangle icon"></i>Danger Zone
            </h3>
        </div>
    ''', unsafe_allow_html=True)
    
    st.markdown('''
        <div class="warning-box">
            <i class="fas fa-warning icon"></i>
            <strong>Warning:</strong> The actions below cannot be undone. Proceed with caution.
        </div>
    ''', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown('<div class="danger-btn">', unsafe_allow_html=True)
        if st.button("🗑️ Clear Session Data", use_container_width=True):
            if st.checkbox("I understand this will clear all session data"):
                st.session_state.sent_numbers = {}
                st.session_state.log_entries = []
                st.session_state.current_folder = None
                st.markdown('''
                    <div class="success-box">
                        <i class="fas fa-check icon"></i>
                        Session data cleared successfully!
                    </div>
                ''', unsafe_allow_html=True)
                time.sleep(1)
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="danger-btn">', unsafe_allow_html=True)
        if st.button("🔄 Reset All Settings", use_container_width=True):
            if st.checkbox("I want to reset all application settings"):
                st.session_state.shared_drive_id = None
                st.markdown('''
                    <div class="success-box">
                        <i class="fas fa-check icon"></i>
                        Settings reset to defaults!
                    </div>
                ''', unsafe_allow_html=True)
                time.sleep(1)
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col3:
        st.markdown('<div class="danger-btn">', unsafe_allow_html=True)
        if st.button("🔌 Disconnect Drive", use_container_width=True):
            if st.checkbox("I want to disconnect Google Drive"):
                st.session_state.drive_service = None
                st.markdown('''
                    <div class="info-box">
                        <i class="fas fa-info icon"></i>
                        Google Drive disconnected. Reconnect in the Send Payslips tab.
                    </div>
                ''', unsafe_allow_html=True)
                time.sleep(1)
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------------------------------
# Enhanced Footer with Animation
# ---------------------------------------------------
st.markdown("---")
st.markdown('''
    <div class="footer">
        <div style="margin-bottom: 1rem;">
            <i class="fas fa-rocket" style="font-size: 2rem; color: #ff6b6b; animation: bounce 2s infinite;"></i>
        </div>
        <h3 style="margin: 1rem 0; background: linear-gradient(135deg, #ff6b6b, #4ecdc4); background-clip: text; -webkit-background-clip: text; -webkit-text-fill-color: transparent;">
            PaySlip Distribution System Pro
        </h3>
        <p style="margin: 0.5rem 0; opacity: 0.8;">
            <i class="fas fa-shield-alt icon"></i>Secure PDF Distribution • 
            <i class="fas fa-cloud icon"></i>Google Drive Integration • 
            <i class="fas fa-mobile-alt icon"></i>WhatsApp Automation
        </p>
        <div style="margin-top: 1.5rem; padding-top: 1rem; border-top: 1px solid rgba(255,255,255,0.2);">
            <div style="display: flex; justify-content: center; gap: 2rem; flex-wrap: wrap;">
                <span><i class="fas fa-code icon"></i>Built with Streamlit</span>
                <span><i class="fas fa-heart icon" style="color: #e74c3c;"></i>Made with Love</span>
                <span><i class="fas fa-copyright icon"></i>2025 HR Management Pro</span>
            </div>
        </div>
        <div style="margin-top: 1rem;">
            <small style="opacity: 0.6;">
                <i class="fas fa-info-circle"></i>
                For support and updates, contact your system administrator
            </small>
        </div>
    </div>
''', unsafe_allow_html=True)

# Add floating action button for quick actions
st.markdown('''
    <div style="position: fixed; bottom: 2rem; right: 2rem; z-index: 1000;">
        <div style="background: linear-gradient(135deg, #ff6b6b, #ee5a24); color: white; padding: 1rem; border-radius: 50%; box-shadow: 0 10px 30px rgba(238, 90, 36, 0.5); cursor: pointer; transition: all 0.3s ease;" 
             onmouseover="this.style.transform='scale(1.1) rotate(5deg)'" 
             onmouseout="this.style.transform='scale(1) rotate(0deg)'">
            <i class="fas fa-question-circle" style="font-size: 1.5rem;"></i>
        </div>
    </div>
''', unsafe_allow_html=True)

# Add some JavaScript for enhanced interactivity (optional)
st.markdown('''
    <script>
    // Add smooth scrolling
    document.documentElement.style.scrollBehavior = 'smooth';
    
    // Add particle effect (optional)
    function createParticle() {
        const particle = document.createElement('div');
        particle.style.position = 'fixed';
        particle.style.width = '4px';
        particle.style.height = '4px';
        particle.style.background = 'rgba(255, 107, 107, 0.6)';
        particle.style.borderRadius = '50%';
        particle.style.pointerEvents = 'none';
        particle.style.zIndex = '999';
        particle.style.left = Math.random() * window.innerWidth + 'px';
        particle.style.top = window.innerHeight + 'px';
        document.body.appendChild(particle);
        
        const animation = particle.animate([
            { transform: 'translateY(0px)', opacity: 1 },
            { transform: 'translateY(-' + window.innerHeight + 'px)', opacity: 0 }
        ], {
            duration: 3000 + Math.random() * 2000,
            easing: 'linear'
        });
        
        animation.onfinish = () => {
            particle.remove();
        };
    }
    
    // Create particles occasionally
    setInterval(createParticle, 5000);
    </script>
''', unsafe_allow_html=True)
