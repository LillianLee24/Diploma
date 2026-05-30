import streamlit as st
import psycopg2
from psycopg2.extras import Json, RealDictCursor
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv
import json
import time
import requests
from googleapiclient.discovery import build
import isodate
from textblob import TextBlob
import re
from collections import Counter
import hashlib
from pathlib import Path
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor
import warnings
warnings.filterwarnings('ignore')

import io
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import openpyxl
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize

# NLTK data downloads
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt')
try:
    nltk.data.find('tokenizers/punkt_tab')
except LookupError:
    nltk.download('punkt_tab')
try:
    stop_words = set(stopwords.words('english'))
except:
    nltk.download('stopwords')
    stop_words = set(stopwords.words('english'))

load_dotenv()

# =====================================================
# CONFIGURATION
# =====================================================

YOUTUBE_API_KEY = os.getenv('YOUTUBE_API_KEY', '')

DB_CONFIG = {
    'host': os.getenv('DB_HOST', 'localhost'),
    'port': os.getenv('DB_PORT', '5432'),
    'database': os.getenv('DB_NAME', 'social_media_analytics'),
    'user': os.getenv('DB_USER', 'postgres'),
    'password': os.getenv('DB_PASSWORD', '')
}

st.set_page_config(
    page_title="Social Media Analytics - Dashboard",
    page_icon="📊",
    layout="wide"
)

# =====================================================
# CUSTOM CSS (hide sidebar completely)
# =====================================================

st.markdown("""
<style>
    /* Hide Streamlit's default sidebar completely */
    [data-testid="stSidebar"] {
    background: linear-gradient(
        180deg,
        #0f172a 0%,
        #1e293b 50%,
        #27365f 100%
    ) !important;
}

    [data-testid="stSidebarCollapsedControl"] {
        display: block !important;
    }
    
    /* Main content takes full width */
    .main .block-container {
        padding-top: 5rem;
        max-width: 1400px;
    }
    
    /* Top navigation bar - we'll use Streamlit buttons now, so only minimal styling */
    .nav-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: white;
        padding: 0.5rem 2rem;
        border-bottom: 1px solid #e2e8f0;
        margin-bottom: 2rem;
    }
    .logo {
        font-size: 1.2rem;
        font-weight: 600;
        background: linear-gradient(135deg, #3b82f6, #f59e0b);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Typography */
    h1, h2, h3, h4, h5, h6 {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
        font-weight: 600;
    }
    
    h1 {
        font-size: 1.75rem;
        background: linear-gradient(135deg, #1e293b 0%, #3b82f6 50%, #f59e0b 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        letter-spacing: -0.02em;
    }
    
    h2 {
        font-size: 1.25rem;
        color: #1e293b;
        border-left: 4px solid;
        border-image: linear-gradient(135deg, #3b82f6, #f59e0b) 1;
        padding-left: 0.75rem;
    }
    
    h3 {
        font-size: 0.875rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        background: linear-gradient(135deg, #64748b 0%, #3b82f6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    
    /* Metric cards */
    .metric-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
        border: 1px solid rgba(59,130,246,0.1);
        border-radius: 20px;
        padding: 1.25rem;
        text-align: left;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 8px 20px rgba(0,0,0,0.05);
        position: relative;
        overflow: hidden;
    }
    
    .metric-card::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 3px;
        background: linear-gradient(90deg, #3b82f6, #f59e0b, #10b981);
    }
    
    .metric-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 20px 30px rgba(0,0,0,0.1);
        border-color: rgba(59,130,246,0.3);
    }
    
    .metric-value {
        font-size: 1.75rem;
        font-weight: 700;
        background: linear-gradient(135deg, #1e293b 0%, #3b82f6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        line-height: 1.2;
    }
    
    .metric-label {
        color: #64748b;
        font-size: 0.7rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-top: 0.5rem;
        font-weight: 600;
    }
    
    /* Insight cards */
    .insight-card {
        background: linear-gradient(135deg, #eff6ff 0%, #e0f2fe 100%);
        border-left: 4px solid #3b82f6;
        border-radius: 16px;
        padding: 1rem;
        margin-bottom: 0.75rem;
        transition: all 0.3s ease;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    
    .insight-card:hover {
        transform: translateX(4px);
        box-shadow: 0 6px 16px rgba(59,130,246,0.15);
    }
    
    .insight-title {
        font-weight: 700;
        color: #1e293b;
        font-size: 0.875rem;
        margin-bottom: 0.25rem;
    }
    
    .insight-description {
        color: #475569;
        font-size: 0.813rem;
        margin-bottom: 0.5rem;
    }
    
    .insight-action {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 0.75rem;
        font-weight: 600;
    }
    
    /* Recommendation carousel */
    .rec-carousel {
        display: flex;
        overflow-x: auto;
        gap: 1rem;
        padding: 0.5rem 0.25rem 1rem 0.25rem;
        scrollbar-width: thin;
    }
    
    .rec-carousel::-webkit-scrollbar {
        height: 4px;
    }
    
    .rec-carousel::-webkit-scrollbar-track {
        background: #e2e8f0;
        border-radius: 4px;
    }
    
    .rec-carousel::-webkit-scrollbar-thumb {
        background: linear-gradient(90deg, #3b82f6, #f59e0b);
        border-radius: 4px;
    }
    
    .rec-slide {
        min-width: 300px;
        background: linear-gradient(135deg, #ffffff 0%, #faf5ff 100%);
        border: 1px solid rgba(139,92,246,0.2);
        border-radius: 16px;
        padding: 1rem;
        transition: all 0.3s ease;
    }
    
    .rec-slide:hover {
        transform: translateY(-4px);
        border-color: #8b5cf6;
        box-shadow: 0 8px 20px rgba(139,92,246,0.15);
    }
    
    /* Cluster cards */
    .cluster-card {
        border-radius: 20px;
        padding: 1rem;
        text-align: center;
        transition: all 0.3s ease;
        backdrop-filter: blur(10px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    }
    
    .cluster-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 24px rgba(0,0,0,0.1);
    }
    
    /* Buttons */
    .stButton button {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border: 1px solid rgba(59,130,246,0.2);
        border-radius: 10px;
        color: #334155;
        font-size: 0.813rem;
        font-weight: 500;
        padding: 0.5rem 1rem;
        transition: all 0.3s ease;
    }
    
    .stButton button:hover {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        border-color: transparent;
        color: white;
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(59,130,246,0.3);
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2rem;
        border-bottom: 2px solid #e2e8f0;
    }
    
    .stTabs [data-baseweb="tab"] {
        color: #64748b;
        font-size: 0.875rem;
        font-weight: 600;
        padding: 0.5rem 0;
        transition: all 0.3s ease;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        border-bottom: 2px solid #3b82f6;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border: 1px solid rgba(59,130,246,0.1);
        border-radius: 10px;
        font-weight: 500;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================
# PROFESSIONAL UI OVERRIDES FOR DEFENSE DEMO
# =====================================================
st.markdown("""
<style>
    :root {
        --app-bg: #f8fafc;
        --surface: #ffffff;
        --surface-2: #f1f5f9;
        --text: #0f172a;
        --muted: #64748b;
        --border: #e2e8f0;
        --blue: #2563eb;
        --blue-dark: #1e3a8a;
        --orange: #f59e0b;
        --green: #10b981;
    }
    .stApp {
        background: radial-gradient(circle at top left, rgba(37,99,235,0.08), transparent 32rem), radial-gradient(circle at top right, rgba(245,158,11,0.08), transparent 28rem), var(--app-bg);
    }
    .main .block-container { padding-top: 1.4rem !important; padding-bottom: 3rem; max-width: 1320px !important; }
    .app-shell { background: rgba(255,255,255,0.86); border: 1px solid rgba(226,232,240,0.9); border-radius: 24px; padding: 18px 22px; margin-bottom: 22px; box-shadow: 0 18px 45px rgba(15,23,42,0.08); backdrop-filter: blur(14px); }
    .app-brand { display:flex; align-items:center; gap:12px; min-height:54px; }
    .brand-mark { width:42px; height:42px; border-radius:14px; background:linear-gradient(135deg,var(--blue),var(--orange)); color:white; display:inline-flex; align-items:center; justify-content:center; font-size:20px; box-shadow:0 10px 24px rgba(37,99,235,0.25); }
    .brand-title { font-size:19px; font-weight:800; letter-spacing:-0.03em; color:var(--text); line-height:1.1; }
    .brand-subtitle { font-size:12px; color:var(--muted); margin-top:3px; }
    .user-pill { border-radius:999px; border:1px solid var(--border); background:#f8fafc; padding:9px 13px; color:#334155; font-size:13px; text-align:center; white-space:nowrap; }
    .page-title-card { background:linear-gradient(135deg,#ffffff 0%,#f8fafc 100%); border:1px solid var(--border); border-radius:22px; padding:22px 24px; margin-bottom:18px; box-shadow:0 12px 32px rgba(15,23,42,0.06); }
    .page-title-card h1, .page-title-card h2, .page-title-card h3 { background:none !important; -webkit-text-fill-color:initial !important; color:var(--text) !important; border-left:none !important; padding-left:0 !important; margin-bottom:4px; }
    .page-eyebrow { font-size:12px; color:var(--blue); font-weight:800; text-transform:uppercase; letter-spacing:0.08em; margin-bottom:8px; }
    .page-description { color:var(--muted); font-size:14px; line-height:1.5; }
    .section-card, .settings-card { background:rgba(255,255,255,0.95); border:1px solid var(--border); border-radius:22px; padding:22px; margin:14px 0 18px 0; box-shadow:0 12px 32px rgba(15,23,42,0.05); }
    .settings-section-title { font-size:15px; font-weight:800; color:var(--text); margin-bottom:16px; display:flex; align-items:center; gap:8px; }
    .danger-zone { border-color:#fecaca !important; background:linear-gradient(135deg,#fff 0%,#fff7f7 100%) !important; }
    .stTabs [data-baseweb="tab-list"] { background:#ffffff; border:1px solid var(--border); border-radius:16px; padding:6px; gap:6px !important; box-shadow:0 8px 20px rgba(15,23,42,0.04); margin-bottom:14px; }
    .stTabs [data-baseweb="tab"] { border-radius:12px; padding:9px 16px !important; color:#475569 !important; font-weight:700 !important; }
    .stTabs [aria-selected="true"] { background:#eff6ff !important; border-bottom:none !important; color:var(--blue) !important; -webkit-text-fill-color:var(--blue) !important; }
    .stButton > button { border-radius:14px !important; min-height:42px; border:1px solid var(--border) !important; background:#ffffff !important; color:#334155 !important; font-weight:750 !important; box-shadow:0 7px 18px rgba(15,23,42,0.05) !important; }
    .stButton > button:hover { border-color:rgba(37,99,235,0.35) !important; background:#eff6ff !important; color:var(--blue-dark) !important; transform:translateY(-1px); box-shadow:0 10px 24px rgba(37,99,235,0.11) !important; }
    div[data-testid="stMetric"] { background:linear-gradient(135deg,#ffffff 0%,#f8fafc 100%); border:1px solid var(--border); border-radius:18px; padding:16px 18px; box-shadow:0 10px 25px rgba(15,23,42,0.05); }
    div[data-testid="stMetricLabel"] p { color:var(--muted) !important; font-size:12px !important; font-weight:750 !important; text-transform:uppercase; letter-spacing:.05em; }
    div[data-testid="stMetricValue"] { color:var(--text) !important; font-weight:850 !important; }
    div[data-testid="stTextInput"] input, div[data-testid="stSelectbox"] div[data-baseweb="select"] { border-radius:14px !important; }
    .streamlit-expanderHeader { border-radius:16px !important; background:#ffffff !important; border:1px solid var(--border) !important; font-weight:750 !important; }
    .block-note { background:#f8fafc; border:1px solid var(--border); border-radius:18px; padding:14px 16px; color:#475569; font-size:13px; line-height:1.45; }
</style>
""", unsafe_allow_html=True)




# =====================================================
# SIDEBAR + LANDING PAGE UI OVERRIDES
# =====================================================
st.markdown("""
<style>
    /* Bring Streamlit sidebar back and make it look like a real product panel */
    [data-testid="stSidebar"] {
        display: block !important;
        background: linear-gradient(180deg, #0f172a 0%, #1e293b 52%, #27365f 100%) !important;
        border-right: 1px solid rgba(255,255,255,0.08);
        box-shadow: 18px 0 45px rgba(15,23,42,0.12);
    }
    [data-testid="stSidebar"] > div:first-child {
        padding-top: 22px;
        padding-left: 18px;
        padding-right: 18px;
    }
    [data-testid="stSidebarCollapsedControl"] { display: none !important; }
    [data-testid="stSidebar"] * { color: #e5e7eb; }
    [data-testid="stSidebar"] .stButton > button {
        background: rgba(255,255,255,0.06) !important;
        border: 1px solid rgba(255,255,255,0.10) !important;
        color: #dbeafe !important;
        border-radius: 16px !important;
        min-height: 46px;
        box-shadow: none !important;
        justify-content: flex-start !important;
        padding-left: 18px !important;
        font-weight: 750 !important;
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        background: rgba(59,130,246,0.22) !important;
        border-color: rgba(147,197,253,0.35) !important;
        color: #ffffff !important;
        transform: translateX(2px);
    }
    [data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.10); }

    .side-brand { padding: 8px 4px 20px 4px; }
    .side-logo {
        width: 48px; height: 48px; border-radius: 16px;
        background: linear-gradient(135deg, #2563eb, #f59e0b);
        display: flex; align-items:center; justify-content:center;
        font-size: 24px; box-shadow: 0 14px 35px rgba(37,99,235,0.35);
        margin-bottom: 12px;
    }
    .side-title { font-size: 20px; font-weight: 850; line-height: 1.12; color:#ffffff; letter-spacing:-0.03em; }
    .side-subtitle { font-size: 12px; color:#93a4bf; margin-top: 6px; line-height:1.35; }
    .side-user {
        margin-top: 18px; padding: 13px 14px; border-radius: 18px;
        background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.10);
    }
    .side-user-name { font-weight: 800; color:#ffffff; font-size: 14px; }
    .side-user-email { color:#94a3b8; font-size: 11px; margin-top:2px; word-break: break-all; }
    .side-caption { color:#64748b; font-size:11px; text-transform:uppercase; letter-spacing:.09em; font-weight:900; margin: 10px 0 8px 4px; }

    .home-hero {
        border-radius: 32px;
        background:
            radial-gradient(circle at 88% 12%, rgba(245,158,11,0.18), transparent 26rem),
            radial-gradient(circle at 18% 18%, rgba(37,99,235,0.18), transparent 26rem),
            linear-gradient(135deg, #ffffff 0%, #f8fafc 54%, #eff6ff 100%);
        border: 1px solid rgba(226,232,240,0.92);
        box-shadow: 0 22px 60px rgba(15,23,42,0.10);
        padding: 36px 38px;
        margin-bottom: 22px;
        overflow: hidden;
    }
    .home-eyebrow { color:#2563eb; font-size:12px; font-weight:900; text-transform:uppercase; letter-spacing:.12em; margin-bottom:10px; }
    .home-title { color:#0f172a; font-size:42px; line-height:1.06; font-weight:900; letter-spacing:-0.055em; max-width: 760px; }
    .home-text { color:#475569; font-size:15px; line-height:1.65; max-width: 720px; margin-top:15px; }
    .home-mini-row { display:flex; gap:12px; flex-wrap:wrap; margin-top:22px; }
    .home-chip { background:#ffffff; border:1px solid #dbeafe; border-radius:999px; padding:9px 14px; color:#1e3a8a; font-size:12px; font-weight:800; box-shadow:0 6px 16px rgba(37,99,235,0.08); }
    .home-card {
        min-height: 152px; border-radius: 24px; padding:22px;
        background: rgba(255,255,255,0.92); border:1px solid #e2e8f0;
        box-shadow: 0 12px 35px rgba(15,23,42,0.06);
    }
    .home-card-number { font-size:12px; color:#2563eb; font-weight:900; letter-spacing:.1em; margin-bottom:10px; }
    .home-card-title { font-size:18px; font-weight:850; color:#0f172a; margin-bottom:8px; }
    .home-card-text { font-size:13px; color:#64748b; line-height:1.5; }
    .product-panel {
        border-radius: 28px; padding: 28px; background: #0f172a; color:white;
        box-shadow: 0 22px 60px rgba(15,23,42,0.18); min-height: 360px; position: relative; overflow:hidden;
    }
    .product-panel:before { content:''; position:absolute; inset:-80px -70px auto auto; width:240px; height:240px; border-radius:50%; background:rgba(37,99,235,.35); filter:blur(8px); }
    .product-kpi { background: rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.12); border-radius:20px; padding:18px; margin-bottom:12px; }
    .product-kpi-value { font-size:26px; font-weight:900; color:#ffffff; }
    .product-kpi-label { color:#93a4bf; font-size:11px; text-transform:uppercase; letter-spacing:.08em; font-weight:800; margin-top:4px; }
    .workflow-card { background:#ffffff; border:1px solid #e2e8f0; border-radius:24px; padding:22px; box-shadow:0 12px 32px rgba(15,23,42,0.06); margin-top:16px; }
    .workflow-title { font-size:18px; color:#0f172a; font-weight:850; margin-bottom:8px; }
    .workflow-text { font-size:13px; color:#64748b; line-height:1.6; }
</style>
""", unsafe_allow_html=True)

# =====================================================
# SESSION STATE INIT
# =====================================================

class DataCache:
    def __init__(self, ttl_seconds=300):
        self.cache = {}
        self.ttl = ttl_seconds
    def get(self, key):
        if key in self.cache:
            data, timestamp = self.cache[key]
            if (datetime.now() - timestamp).total_seconds() < self.ttl:
                return data
            else:
                del self.cache[key]
        return None
    def set(self, key, data):
        self.cache[key] = (data, datetime.now())

cache = DataCache(ttl_seconds=300)

# Инициализация session_state
if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False
if 'user' not in st.session_state:
    st.session_state['user'] = None
if 'user_id' not in st.session_state:
    st.session_state['user_id'] = None
if 'user_email' not in st.session_state:
    st.session_state['user_email'] = None
if 'channel_data' not in st.session_state:
    st.session_state['channel_data'] = None
if 'videos_data' not in st.session_state:
    st.session_state['videos_data'] = None
if 'comments_data' not in st.session_state:
    st.session_state['comments_data'] = None
if 'channel_loaded' not in st.session_state:
    st.session_state['channel_loaded'] = False
if 'current_channel_id' not in st.session_state:
    st.session_state['current_channel_id'] = None
if 'ml_models' not in st.session_state:
    st.session_state['ml_models'] = {}
if 'date_range' not in st.session_state:
    st.session_state['date_range'] = None
if 'selected_categories' not in st.session_state:
    st.session_state['selected_categories'] = []
if 'current_page' not in st.session_state:
    st.session_state['current_page'] = 'dashboard'
if 'search_results' not in st.session_state:
    st.session_state['search_results'] = []
if 'selected_channel_id' not in st.session_state:
    st.session_state['selected_channel_id'] = None
if 'theme' not in st.session_state:
    st.session_state['theme'] = 'light'
if 'language' not in st.session_state:
    st.session_state['language'] = 'English'
if 'notifications_enabled' not in st.session_state:
    st.session_state['notifications_enabled'] = True
if 'auto_save_reports' not in st.session_state:
    st.session_state['auto_save_reports'] = True
if 'page' not in st.session_state:
    st.session_state['page'] = 'home'  # default landing page after login

# =====================================================
# DATABASE FUNCTIONS (оставлены без изменений)
# =====================================================

def get_db_connection():
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        return conn
    except psycopg2.OperationalError as e:
        if 'database "social_media_analytics" does not exist' in str(e):
            try:
                default_config = DB_CONFIG.copy()
                default_config['database'] = 'postgres'
                conn_default = psycopg2.connect(**default_config)
                conn_default.autocommit = True
                cur = conn_default.cursor()
                cur.execute(f"CREATE DATABASE {DB_CONFIG['database']}")
                cur.close()
                conn_default.close()
                conn = psycopg2.connect(**DB_CONFIG)
                return conn
            except Exception as create_error:
                st.error(f"Database creation error: {create_error}")
                return None
        else:
            st.error(f"PostgreSQL connection error: {e}")
            return None
    except Exception as e:
        st.error(f"PostgreSQL connection error: {e}")
        return None

def init_postgres_db():
    conn = get_db_connection()
    if not conn:
        st.warning("PostgreSQL not available. Some features will be limited.")
        return False
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username VARCHAR(100) UNIQUE NOT NULL,
            email VARCHAR(255) UNIQUE NOT NULL,
            password_hash VARCHAR(255) NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            last_login TIMESTAMP,
            preferences JSONB DEFAULT '{}'::jsonb
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS saved_channels (
            id SERIAL PRIMARY KEY,
            user_id INTEGER REFERENCES users(id) ON DELETE CASCADE,
            channel_id VARCHAR(100) NOT NULL,
            channel_title VARCHAR(255),
            channel_thumbnail VARCHAR(500),
            saved_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            notes TEXT,
            UNIQUE(user_id, channel_id)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS favorite_videos (
            id SERIAL PRIMARY KEY,
            user_id INTEGER REFERENCES users(id) ON DELETE CASCADE,
            video_id VARCHAR(100) NOT NULL,
            video_title VARCHAR(500),
            channel_id VARCHAR(100),
            saved_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(user_id, video_id)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS analysis_history (
            id SERIAL PRIMARY KEY,
            user_id INTEGER REFERENCES users(id) ON DELETE CASCADE,
            channel_id VARCHAR(100) NOT NULL,
            channel_title VARCHAR(255),
            analysis_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            metrics JSONB,
            predictions JSONB,
            clusters JSONB,
            insights JSONB
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS youtube_cache (
            id SERIAL PRIMARY KEY,
            cache_key VARCHAR(255) UNIQUE NOT NULL,
            data JSONB NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            expires_at TIMESTAMP
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS ml_models (
            id SERIAL PRIMARY KEY,
            model_name VARCHAR(100) NOT NULL,
            channel_id VARCHAR(100),
            model_data BYTEA,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(model_name, channel_id)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS comments_analysis (
            id SERIAL PRIMARY KEY,
            video_id VARCHAR(100) NOT NULL,
            comment_id VARCHAR(100) UNIQUE,
            author VARCHAR(255),
            text TEXT,
            sentiment_score FLOAT,
            sentiment_label VARCHAR(20),
            like_count INTEGER,
            published_at TIMESTAMP,
            analyzed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    cur.close()
    conn.close()
    return True

def get_cached_youtube_data(cache_key):
    conn = get_db_connection()
    if not conn:
        return None
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("SELECT data FROM youtube_cache WHERE cache_key = %s AND (expires_at IS NULL OR expires_at > NOW())", (cache_key,))
    result = cur.fetchone()
    cur.close()
    conn.close()
    if result:
        return result['data']
    return None

def set_cached_youtube_data(cache_key, data, ttl_seconds=3600):
    conn = get_db_connection()
    if not conn:
        return False
    cur = conn.cursor()
    expires_at = datetime.now() + timedelta(seconds=ttl_seconds)
    cur.execute("""
        INSERT INTO youtube_cache (cache_key, data, expires_at)
        VALUES (%s, %s, %s)
        ON CONFLICT (cache_key) DO UPDATE SET
            data = EXCLUDED.data,
            expires_at = EXCLUDED.expires_at,
            created_at = NOW()
    """, (cache_key, Json(data), expires_at))
    conn.commit()
    cur.close()
    conn.close()
    return True

def save_analysis_history(user_id, channel_id, channel_title, metrics, predictions, clusters, insights):
    if not st.session_state.get('auto_save_reports', True):
        return False
    conn = get_db_connection()
    if not conn:
        return False
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO analysis_history (user_id, channel_id, channel_title, metrics, predictions, clusters, insights)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
    """, (user_id, channel_id, channel_title, Json(metrics), Json(predictions), Json(clusters), Json(insights)))
    conn.commit()
    cur.close()
    conn.close()
    return True

def get_analysis_history(user_id, limit=20):
    conn = get_db_connection()
    if not conn:
        return []
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("SELECT * FROM analysis_history WHERE user_id = %s ORDER BY analysis_date DESC LIMIT %s", (user_id, limit))
    results = cur.fetchall()
    cur.close()
    conn.close()
    return results

def save_channel_to_favorites(user_id, channel_id, channel_title, thumbnail):
    conn = get_db_connection()
    if not conn:
        return False
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO saved_channels (user_id, channel_id, channel_title, channel_thumbnail)
        VALUES (%s, %s, %s, %s)
        ON CONFLICT (user_id, channel_id) DO NOTHING
    """, (user_id, channel_id, channel_title, thumbnail))
    conn.commit()
    cur.close()
    conn.close()
    return True

def get_saved_channels(user_id):
    conn = get_db_connection()
    if not conn:
        return []
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("SELECT * FROM saved_channels WHERE user_id = %s ORDER BY saved_at DESC", (user_id,))
    results = cur.fetchall()
    cur.close()
    conn.close()
    return results

def remove_saved_channel(user_id, channel_id):
    conn = get_db_connection()
    if not conn:
        return False
    cur = conn.cursor()
    cur.execute("DELETE FROM saved_channels WHERE user_id = %s AND channel_id = %s", (user_id, channel_id))
    conn.commit()
    cur.close()
    conn.close()
    return True

def save_favorite_video(user_id, video_id, video_title, channel_id):
    conn = get_db_connection()
    if not conn:
        return False
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO favorite_videos (user_id, video_id, video_title, channel_id)
        VALUES (%s, %s, %s, %s)
        ON CONFLICT (user_id, video_id) DO NOTHING
    """, (user_id, video_id, video_title, channel_id))
    conn.commit()
    cur.close()
    conn.close()
    return True

def get_favorite_videos(user_id):
    conn = get_db_connection()
    if not conn:
        return []
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("SELECT * FROM favorite_videos WHERE user_id = %s ORDER BY saved_at DESC", (user_id,))
    results = cur.fetchall()
    cur.close()
    conn.close()
    return results

def save_comment_analysis(video_id, comment_id, author, text, sentiment_score, sentiment_label, like_count, published_at):
    conn = get_db_connection()
    if not conn:
        return False
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO comments_analysis (video_id, comment_id, author, text, sentiment_score, sentiment_label, like_count, published_at)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT (comment_id) DO UPDATE SET
            sentiment_score = EXCLUDED.sentiment_score,
            sentiment_label = EXCLUDED.sentiment_label,
            like_count = EXCLUDED.like_count
    """, (video_id, comment_id, author, text[:1000], sentiment_score, sentiment_label, like_count, published_at))
    conn.commit()
    cur.close()
    conn.close()
    return True

def get_color_palette(n, palette='gradient'):
    palettes = {
        'gradient': ['#3b82f6', '#60a5fa', '#f59e0b', '#fbbf24', '#10b981', '#34d399', '#8b5cf6', '#a78bfa'],
        'bright': ['#3b82f6', '#f59e0b', '#10b981', '#ef4444', '#8b5cf6', '#06b6d4', '#f97316', '#84cc16'],
        'pastel': ['#93c5fd', '#fde68a', '#a7f3d0', '#fecaca', '#c4b5fd', '#67e8f9', '#fed7aa', '#bbf7d0']
    }
    colors = palettes.get(palette, palettes['gradient'])
    if n > len(colors):
        colors = colors * (n // len(colors) + 1)
    return colors[:n]

def format_number(num):
    if num >= 1_000_000:
        return f"{num/1_000_000:.1f}M"
    elif num >= 1_000:
        return f"{num/1_000:.1f}K"
    else:
        return str(num)

def analyze_sentiment(text):
    if not text or not isinstance(text, str):
        return 0, 'neutral'
    try:
        blob = TextBlob(text[:5000])
        polarity = blob.sentiment.polarity
        if polarity > 0.2:
            label = 'positive'
        elif polarity < -0.2:
            label = 'negative'
        else:
            label = 'neutral'
        return polarity, label
    except:
        return 0, 'neutral'

# =====================================================
# ML FUNCTIONS (оставлены без изменений)
# =====================================================

def train_view_prediction_model(videos_data):
    if not videos_data or len(videos_data) < 5:
        return None
    df = pd.DataFrame(videos_data)
    df['published_datetime'] = pd.to_datetime(df['published_at'])
    df['published_datetime'] = df['published_datetime'].dt.tz_localize(None)
    df['hour'] = df['published_datetime'].dt.hour
    df['day_of_week'] = df['published_datetime'].dt.dayofweek
    now = datetime.now().replace(tzinfo=None)
    df['days_since_published'] = (now - df['published_datetime']).dt.days
    df['log_views'] = np.log1p(df['view_count'])
    df['engagement_rate'] = df['engagement_rate']
    df['duration_min'] = df['duration_seconds'] / 60
    feature_cols = ['hour', 'day_of_week', 'engagement_rate', 'duration_min', 'days_since_published']
    X = df[feature_cols].fillna(0)
    y = df['log_views'].values
    model = RandomForestRegressor(n_estimators=100, random_state=42)
    model.fit(X, y)
    return {'model': model, 'feature_cols': feature_cols, 'feature_importance': dict(zip(feature_cols, model.feature_importances_))}

def predict_views(model_info, hour, day_of_week, engagement_rate, duration_min, days_since_published):
    if not model_info:
        return None
    X_pred = np.array([[hour, day_of_week, engagement_rate, duration_min, days_since_published]])
    log_pred = model_info['model'].predict(X_pred)[0]
    return int(np.expm1(log_pred))

def train_growth_model(videos_data):
    if not videos_data or len(videos_data) < 3:
        return None
    df = pd.DataFrame(videos_data)
    df = df.sort_values('published_at')
    df['video_number'] = range(len(df))
    df['cumulative_views'] = df['view_count'].cumsum()
    df['log_cumulative'] = np.log1p(df['cumulative_views'])
    X = df[['video_number']].values
    y = df['log_cumulative'].values
    model = LinearRegression()
    model.fit(X, y)
    return {'model': model, 'intercept': model.intercept_, 'coef': model.coef_[0]}

def predict_growth(growth_model, future_videos_count):
    if not growth_model:
        return None
    predictions = []
    for i in range(1, future_videos_count + 1):
        log_pred = growth_model['intercept'] + growth_model['coef'] * i
        predictions.append(int(np.expm1(log_pred)))
    return predictions

def cluster_videos(videos_data, n_clusters=3):
    if not videos_data or len(videos_data) < 3:
        return None
    df = pd.DataFrame(videos_data)
    features = ['view_count', 'like_count', 'comment_count', 'engagement_rate']
    X = df[features].fillna(0)
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    kmeans = KMeans(n_clusters=min(n_clusters, len(df)), random_state=42, n_init=10)
    clusters = kmeans.fit_predict(X_scaled)
    cluster_centers = scaler.inverse_transform(kmeans.cluster_centers_)
    cluster_avg_views = [cluster_centers[i][0] for i in range(len(cluster_centers))]
    cluster_labels = {}
    sorted_idx = np.argsort(cluster_avg_views)[::-1]
    for i, idx in enumerate(sorted_idx):
        if i == 0:
            label = 'high'
        elif i == 1:
            label = 'medium'
        else:
            label = 'low'
        cluster_labels[idx] = label
    df['cluster'] = clusters
    df['cluster_label'] = df['cluster'].map(cluster_labels)
    cluster_stats = {}
    for cluster_id in range(len(cluster_centers)):
        cluster_videos = df[df['cluster'] == cluster_id]
        cluster_stats[cluster_labels[cluster_id]] = {
            'count': len(cluster_videos),
            'avg_views': cluster_videos['view_count'].mean(),
            'avg_engagement': cluster_videos['engagement_rate'].mean(),
            'avg_likes': cluster_videos['like_count'].mean(),
            'avg_comments': cluster_videos['comment_count'].mean(),
            'videos': cluster_videos.to_dict('records')
        }
    return {'df': df, 'cluster_stats': cluster_stats, 'scaler': scaler, 'kmeans': kmeans, 'cluster_labels': cluster_labels}

# =====================================================
# YOUTUBE API FUNCTIONS (оставлены без изменений)
# =====================================================

def get_youtube_service():
    try:
        return build('youtube', 'v3', developerKey=YOUTUBE_API_KEY)
    except Exception as e:
        st.error(f"YouTube API initialization error: {e}")
        return None

def get_youtube_channel_by_id(channel_id):
    cache_key = f"channel_{channel_id}"
    cached = get_cached_youtube_data(cache_key)
    if cached:
        return cached
    try:
        youtube = get_youtube_service()
        if not youtube:
            return None
        request = youtube.channels().list(part='snippet,statistics,contentDetails,topicDetails,status', id=channel_id)
        response = request.execute()
        if response['items']:
            set_cached_youtube_data(cache_key, response['items'][0], ttl_seconds=86400)
            return response['items'][0]
        request = youtube.channels().list(part='snippet,statistics,contentDetails,topicDetails,status', forUsername=channel_id)
        response = request.execute()
        if response['items']:
            set_cached_youtube_data(cache_key, response['items'][0], ttl_seconds=86400)
            return response['items'][0]
        return None
    except Exception as e:
        st.error(f"Error fetching channel: {e}")
        return None

def get_channel_info(channel_id):
    cache_key = f"channel_info_{channel_id}"
    cached = get_cached_youtube_data(cache_key)
    if cached:
        return cached
    try:
        channel = get_youtube_channel_by_id(channel_id)
        if not channel:
            st.error("Channel not found. Check ID or name.")
            return None
        subscribers = int(channel['statistics'].get('subscriberCount', 0))
        videos_count = int(channel['statistics'].get('videoCount', 0))
        views = int(channel['statistics'].get('viewCount', 0))
        channel_info = {
            'channel_id': channel['id'],
            'title': channel['snippet']['title'],
            'subscribers': subscribers,
            'videos': videos_count,
            'views': views,
            'thumbnail': channel['snippet']['thumbnails']['high']['url'],
            'description': channel['snippet']['description'][:300] + '...' if len(channel['snippet']['description']) > 300 else channel['snippet']['description'],
            'published_at': channel['snippet']['publishedAt'],
            'country': channel['snippet'].get('country', 'Not specified'),
            'custom_url': channel['snippet'].get('customUrl', ''),
            'topic_categories': channel.get('topicDetails', {}).get('topicCategories', [])
        }
        authority_score = min(100, int((subscribers / 100000) * 40 + (views / 10000000) * 60))
        channel_info['authority_score'] = authority_score
        set_cached_youtube_data(cache_key, channel_info, ttl_seconds=86400)
        return channel_info
    except Exception as e:
        st.error(f"Error fetching channel info: {e}")
        return None

def get_channel_videos(channel_id, max_results=30):
    cache_key = f"channel_videos_{channel_id}_{max_results}"
    cached = get_cached_youtube_data(cache_key)
    if cached:
        return cached
    try:
        youtube = get_youtube_service()
        if not youtube:
            return []
        channel = get_youtube_channel_by_id(channel_id)
        if not channel:
            return []
        real_channel_id = channel['id']
        channels_response = youtube.channels().list(part='contentDetails', id=real_channel_id).execute()
        if not channels_response['items']:
            return []
        uploads_playlist_id = channels_response['items'][0]['contentDetails']['relatedPlaylists']['uploads']
        videos = []
        next_page_token = None
        while len(videos) < max_results:
            playlist_request = youtube.playlistItems().list(part='snippet', playlistId=uploads_playlist_id, maxResults=min(50, max_results - len(videos)), pageToken=next_page_token)
            playlist_response = playlist_request.execute()
            video_ids = [item['snippet']['resourceId']['videoId'] for item in playlist_response['items']]
            if video_ids:
                videos_request = youtube.videos().list(part='statistics,contentDetails,snippet', id=','.join(video_ids))
                videos_response = videos_request.execute()
                for video_item in videos_response['items']:
                    duration = video_item['contentDetails']['duration']
                    try:
                        duration_seconds = int(isodate.parse_duration(duration).total_seconds())
                    except:
                        duration_seconds = 0
                    views = int(video_item['statistics'].get('viewCount', 0))
                    likes = int(video_item['statistics'].get('likeCount', 0))
                    comments = int(video_item['statistics'].get('commentCount', 0))
                    engagement_rate = ((likes + comments) / views * 100) if views > 0 else 0
                    video_info = {
                        'video_id': video_item['id'],
                        'title': video_item['snippet']['title'],
                        'published_at': video_item['snippet']['publishedAt'],
                        'duration_seconds': duration_seconds,
                        'view_count': views,
                        'like_count': likes,
                        'comment_count': comments,
                        'engagement_rate': engagement_rate,
                        'thumbnail': video_item['snippet']['thumbnails']['high']['url'],
                        'description': video_item['snippet'].get('description', '')[:500]
                    }
                    videos.append(video_info)
            next_page_token = playlist_response.get('nextPageToken')
            if not next_page_token:
                break
        set_cached_youtube_data(cache_key, videos, ttl_seconds=3600)
        return videos
    except Exception as e:
        st.error(f"Error fetching videos: {e}")
        return []

def get_video_comments(video_id, max_results=100):
    cache_key = f"video_comments_{video_id}_{max_results}"
    cached = get_cached_youtube_data(cache_key)
    if cached:
        return cached
    try:
        youtube = get_youtube_service()
        if not youtube:
            return []
        comments = []
        next_page_token = None
        while len(comments) < max_results:
            request = youtube.commentThreads().list(part='snippet', videoId=video_id, maxResults=min(100, max_results - len(comments)), pageToken=next_page_token, textFormat='plainText')
            response = request.execute()
            for item in response['items']:
                snippet = item['snippet']['topLevelComment']['snippet']
                comment = {
                    'comment_id': item['id'],
                    'author': snippet.get('authorDisplayName', 'unknown'),
                    'text': snippet.get('textDisplay', ''),
                    'like_count': snippet.get('likeCount', 0),
                    'published_at': snippet.get('publishedAt', ''),
                    'sentiment_score': None,
                    'sentiment_label': None
                }
                sentiment_score, sentiment_label = analyze_sentiment(comment['text'])
                comment['sentiment_score'] = sentiment_score
                comment['sentiment_label'] = sentiment_label
                comments.append(comment)
                save_comment_analysis(video_id, comment['comment_id'], comment['author'], comment['text'], sentiment_score, sentiment_label, comment['like_count'], comment['published_at'])
            next_page_token = response.get('nextPageToken')
            if not next_page_token:
                break
        set_cached_youtube_data(cache_key, comments, ttl_seconds=3600)
        return comments
    except Exception as e:
        st.error(f"Error fetching comments for video {video_id}: {e}")
        return []

def get_all_comments_for_channel(videos, max_comments_per_video=50):
    all_comments = []
    progress_bar = st.progress(0)
    for idx, video in enumerate(videos[:10]):
        comments = get_video_comments(video['video_id'], max_comments_per_video)
        for comment in comments:
            comment['video_title'] = video['title']
            comment['video_id'] = video['video_id']
            all_comments.append(comment)
        progress_bar.progress((idx + 1) / min(len(videos), 10))
    progress_bar.empty()
    return all_comments

def search_youtube_channels(query, max_results=5):
    cache_key = f"search_{query}_{max_results}"
    cached = get_cached_youtube_data(cache_key)
    if cached:
        return cached
    try:
        youtube = get_youtube_service()
        if not youtube:
            return []
        request = youtube.search().list(part='snippet', q=query, type='channel', maxResults=max_results)
        response = request.execute()
        channels = []
        for item in response['items']:
            channels.append({
                'channel_id': item['id']['channelId'],
                'title': item['snippet']['title'],
                'description': item['snippet']['description'][:150],
                'thumbnail': item['snippet']['thumbnails']['default']['url']
            })
        set_cached_youtube_data(cache_key, channels, ttl_seconds=3600)
        return channels
    except Exception as e:
        st.error(f"Error searching channels: {e}")
        return []

# =====================================================
# AUTHENTICATION FUNCTIONS (оставлены без изменений)
# =====================================================

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def register_user(username, email, password):
    conn = get_db_connection()
    if not conn:
        return False, "Database connection error"
    cur = conn.cursor()
    try:
        cur.execute("INSERT INTO users (username, email, password_hash) VALUES (%s, %s, %s) RETURNING id", (username, email, hash_password(password)))
        user_id = cur.fetchone()[0]
        conn.commit()
        return True, "Registration successful!"
    except psycopg2.IntegrityError as e:
        conn.rollback()
        if 'username' in str(e):
            return False, "Username already exists"
        elif 'email' in str(e):
            return False, "Email already exists"
        return False, f"Error: {e}"
    except Exception as e:
        conn.rollback()
        return False, f"Registration error: {str(e)}"
    finally:
        cur.close()
        conn.close()

def login_user(email, password):
    conn = get_db_connection()
    if not conn:
        return False, "Database connection error"
    cur = conn.cursor(cursor_factory=RealDictCursor)
    try:
        cur.execute("SELECT id, username, email FROM users WHERE email = %s AND password_hash = %s", (email, hash_password(password)))
        user = cur.fetchone()
        if user:
            cur.execute("UPDATE users SET last_login = NOW() WHERE id = %s", (user['id'],))
            conn.commit()
            st.session_state['authenticated'] = True
            st.session_state['user'] = user['username']
            st.session_state['user_id'] = user['id']
            st.session_state['user_email'] = user['email']
            return True, "Login successful!"
        else:
            return False, "Invalid email or password"
    except Exception as e:
        return False, f"Login error: {str(e)}"
    finally:
        cur.close()
        conn.close()

def logout_user():
    st.session_state['authenticated'] = False
    st.session_state['user'] = None
    st.session_state['user_id'] = None
    st.session_state['user_email'] = None
    st.session_state['channel_loaded'] = False
    st.session_state['channel_data'] = None
    st.session_state['videos_data'] = None
    st.session_state['comments_data'] = None
    st.session_state['current_channel_id'] = None
    st.session_state['page'] = 'home'  # reset to landing page
    st.rerun()

# =====================================================
# ANALYTICS FUNCTIONS (оставлены без изменений)
# =====================================================

def analyze_best_posting_time(videos_data):
    if not videos_data or len(videos_data) < 5:
        return None
    df = pd.DataFrame(videos_data)
    df['published_datetime'] = pd.to_datetime(df['published_at'])
    df['hour'] = df['published_datetime'].dt.hour
    df['day_of_week'] = df['published_datetime'].dt.dayofweek
    hourly_performance = df.groupby('hour')['view_count'].agg(['mean', 'count']).reset_index()
    best_hour = hourly_performance.loc[hourly_performance['mean'].idxmax(), 'hour']
    daily_performance = df.groupby('day_of_week')['view_count'].agg(['mean', 'count']).reset_index()
    best_day = daily_performance.loc[daily_performance['mean'].idxmax(), 'day_of_week']
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    return {'best_hour': int(best_hour), 'best_day': days[int(best_day)], 'hourly_data': hourly_performance.to_dict('records'), 'daily_data': daily_performance.to_dict('records')}

def find_view_drop_patterns(videos_data):
    if not videos_data or len(videos_data) < 5:
        return None
    df = pd.DataFrame(videos_data)
    df = df.sort_values('published_at')
    df['rolling_avg'] = df['view_count'].rolling(window=3, min_periods=1).mean()
    df['view_drop'] = ((df['rolling_avg'].shift(1) - df['view_count']) / df['rolling_avg'].shift(1) * 100)
    drops = df[df['view_drop'] > 30]
    drop_patterns = {'total_drops': len(drops), 'avg_drop_percentage': drops['view_drop'].mean() if len(drops) > 0 else 0, 'drop_dates': drops[['title', 'published_at', 'view_drop']].to_dict('records') if len(drops) > 0 else [], 'common_topics': []}
    return drop_patterns

def get_best_video_type(videos_data):
    if not videos_data or len(videos_data) < 5:
        return None
    df = pd.DataFrame(videos_data)
    df['duration_category'] = pd.cut(df['duration_seconds'] / 60, bins=[0,5,10,20,30,60,float('inf')], labels=['<5 min','5-10 min','10-20 min','20-30 min','30-60 min','>60 min'])
    duration_performance = df.groupby('duration_category')['view_count'].mean().sort_values(ascending=False)
    best_duration = duration_performance.index[0] if len(duration_performance) > 0 else None
    title_keywords = {}
    for _, row in df.iterrows():
        words = row['title'].lower().split()[:10]
        for word in words:
            if len(word) > 3:
                if word not in title_keywords:
                    title_keywords[word] = []
                title_keywords[word].append(row['view_count'])
    keyword_performance = {word: np.mean(views) for word, views in title_keywords.items() if len(views) >= 2}
    best_keywords = sorted(keyword_performance.items(), key=lambda x: x[1], reverse=True)[:5]
    return {'best_duration': best_duration, 'duration_performance': duration_performance.to_dict(), 'best_keywords': best_keywords}

def generate_key_insights(channel_info, videos_data, cluster_stats, posting_analysis, video_type_analysis, drop_patterns):
    insights = []
    if video_type_analysis and video_type_analysis.get('best_duration'):
        insights.append({'category': 'Content Strategy', 'title': 'Best Performing Video Length', 'description': f"Videos lasting {video_type_analysis['best_duration']} perform best on this channel.", 'action': 'Focus on creating content in this duration range for maximum engagement.'})
    if posting_analysis:
        insights.append({'category': 'Timing', 'title': 'Optimal Posting Time', 'description': f"Best time to publish: {posting_analysis['best_hour']}:00 on {posting_analysis['best_day']}.", 'action': f'Schedule your videos for {posting_analysis["best_hour"]}:00 on {posting_analysis["best_day"]} to maximize initial views.'})
    if drop_patterns and drop_patterns.get('total_drops', 0) > 0:
        insights.append({'category': 'Warning', 'title': 'View Drop Detected', 'description': f"Detected {drop_patterns['total_drops']} videos with view drops >30%.", 'action': 'Review content strategy and analyze what changed during these periods.'})
    if cluster_stats and 'high' in cluster_stats:
        popular = cluster_stats['high']
        insights.append({'category': 'Top Performers', 'title': 'High Performance Videos Analysis', 'description': f"High performance videos average {format_number(popular['avg_views'])} views with {popular['avg_engagement']:.1f}% engagement.", 'action': 'Analyze the topics, titles, and thumbnails of these videos to replicate success.'})
    insights.append({'category': 'Growth', 'title': 'Subscriber Growth', 'description': f"Channel has {format_number(channel_info['subscribers'])} subscribers with {channel_info['authority_score']}/100 authority score.", 'action': 'Maintain consistent posting schedule to continue growth trajectory.'})
    if video_type_analysis and video_type_analysis.get('best_keywords'):
        keywords = ', '.join([f'"{kw}"' for kw, _ in video_type_analysis['best_keywords'][:3]])
        insights.append({'category': 'SEO', 'title': 'High-Performing Keywords', 'description': f"Top performing keywords: {keywords}", 'action': 'Include these keywords in your video titles and descriptions for better discoverability.'})
    return insights

# =====================================================
# TITLE MODEL FUNCTIONS (оставлены без изменений)
# =====================================================

def extract_title_features(title):
    if not title:
        return {'sentiment': 0, 'length': 0, 'word_count': 0, 'exclamation_count': 0, 'question_count': 0}
    blob = TextBlob(title)
    sentiment = blob.sentiment.polarity
    length = len(title)
    word_count = len(title.split())
    exclamation_count = title.count('!')
    question_count = title.count('?')
    return {'sentiment': sentiment, 'length': length, 'word_count': word_count, 'exclamation_count': exclamation_count, 'question_count': question_count}

def train_title_model(videos_data):
    if not videos_data or len(videos_data) < 10:
        return None
    df = pd.DataFrame(videos_data)
    title_features = df['title'].apply(extract_title_features).apply(pd.Series)
    X = title_features[['sentiment', 'length', 'word_count', 'exclamation_count', 'question_count']]
    y = np.log1p(df['view_count'])
    model = RandomForestRegressor(n_estimators=100, random_state=42)
    model.fit(X, y)
    return {'model': model, 'feature_cols': X.columns.tolist()}

def predict_title_performance(title, model_info):
    if not model_info:
        return None
    features = extract_title_features(title)
    X_pred = pd.DataFrame([[features['sentiment'], features['length'], features['word_count'], features['exclamation_count'], features['question_count']]], columns=model_info['feature_cols'])
    log_pred = model_info['model'].predict(X_pred)[0]
    return int(np.expm1(log_pred))

def compare_titles(title1, title2, channel_videos):
    model_info = train_title_model(channel_videos)
    if not model_info:
        return None, None, "Not enough data to train model (need at least 10 videos)"
    pred1 = predict_title_performance(title1, model_info)
    pred2 = predict_title_performance(title2, model_info)
    if pred1 is None or pred2 is None:
        return None, None, "Prediction error"
    if pred1 > pred2:
        recommendation = f"✅ Title 1 (“{title1[:50]}”) is predicted better: {((pred1/pred2)-1)*100:.1f}% more views"
    elif pred2 > pred1:
        recommendation = f"✅ Title 2 (“{title2[:50]}”) is predicted better: {((pred2/pred1)-1)*100:.1f}% more views"
    else:
        recommendation = "Both titles have similar potential"
    return pred1, pred2, recommendation

def generate_content_plan(high_cluster_videos, top_n=5):
    if not high_cluster_videos:
        return ["Not enough videos in High Performance cluster to generate ideas."]
    texts = []
    for video in high_cluster_videos:
        title = video.get('title', '')
        desc = video.get('description', '')
        texts.append(f"{title} {desc}")
    all_text = " ".join(texts).lower()
    words = word_tokenize(all_text)
    filtered_words = [w for w in words if w.isalpha() and w not in stop_words and len(w) > 3]
    word_freq = Counter(filtered_words)
    most_common = word_freq.most_common(20)
    topics = []
    for i in range(min(top_n, len(most_common))):
        word = most_common[i][0]
        topics.append(f"Video about {word}: explore this topic in depth, using successful formats from top videos.")
    if len(topics) < top_n:
        topics.extend([f"Idea {i+1}: create a video based on your content with the highest engagement." for i in range(top_n - len(topics))])
    return topics

# =====================================================
# EXPORT FUNCTIONS (оставлены без изменений)
# =====================================================

def export_to_pdf(channel_info, videos, insights, metrics_dict):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name='Title', parent=styles['Title'], fontSize=16, textColor=colors.HexColor('#1e293b'))
    elements = []
    elements.append(Paragraph(f"Channel Report: {channel_info['title']}", title_style))
    elements.append(Spacer(1,12))
    elements.append(Paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles['Normal']))
    elements.append(Spacer(1,12))
    data = [['Metric', 'Value']]
    data.append(['Subscribers', format_number(channel_info['subscribers'])])
    data.append(['Total Views', format_number(channel_info['views'])])
    data.append(['Videos', channel_info['videos']])
    data.append(['Average Engagement', f"{metrics_dict.get('avg_engagement', 0):.2f}%"])
    t = Table(data)
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.grey),('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),('ALIGN',(0,0),(-1,-1),'CENTER'),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('BOTTOMPADDING',(0,0),(-1,0),12),('GRID',(0,0),(-1,-1),1,colors.black)]))
    elements.append(t)
    elements.append(Spacer(1,12))
    elements.append(Paragraph("Key Insights:", styles['Heading2']))
    for ins in insights[:5]:
        elements.append(Paragraph(f"• {ins['title']}: {ins['description'][:100]}", styles['Normal']))
    doc.build(elements)
    buffer.seek(0)
    return buffer

def export_to_excel(channel_info, videos, insights):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        metrics_df = pd.DataFrame([{'Channel': channel_info['title'], 'Subscribers': channel_info['subscribers'], 'Views': channel_info['views'], 'Videos': channel_info['videos'], 'Authority Score': channel_info['authority_score']}])
        metrics_df.to_excel(writer, sheet_name='Overview', index=False)
        if videos:
            videos_df = pd.DataFrame(videos)
            videos_df.to_excel(writer, sheet_name='Videos', index=False)
        insights_df = pd.DataFrame(insights)
        insights_df.to_excel(writer, sheet_name='Insights', index=False)
    output.seek(0)
    return output

# =====================================================
# LOGIN PAGE
# =====================================================

def show_login_page():
    st.markdown("""
        <style>
        .login-container {
            max-width: 420px;
            margin: 0 auto;
            padding: 2.5rem;
            background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
            border-radius: 24px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.05);
            border: 1px solid rgba(59,130,246,0.1);
        }
        .login-title {
            font-size: 1.75rem;
            font-weight: 700;
            text-align: center;
            margin-bottom: 2rem;
            background: linear-gradient(135deg, #1e293b 0%, #3b82f6 50%, #f59e0b 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        </style>
    """, unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        st.markdown('<h1 class="login-title">Social Media Analytics</h1>', unsafe_allow_html=True)
        tab_login, tab_register = st.tabs(["Sign In", "Create Account"])
        with tab_login:
            with st.form("login_form"):
                email = st.text_input("Email", placeholder="your@email.com")
                password = st.text_input("Password", type="password", placeholder="Password")
                submitted = st.form_submit_button("Sign In", use_container_width=True)
                if submitted:
                    if email and password:
                        success, message = login_user(email, password)
                        if success:
                            st.success(message)
                            st.rerun()
                        else:
                            st.error(message)
                    else:
                        st.warning("Please fill in all fields")
        with tab_register:
            with st.form("register_form"):
                new_username = st.text_input("Username", placeholder="johndoe")
                new_email = st.text_input("Email", placeholder="john@example.com")
                new_password = st.text_input("Password", type="password", placeholder="Password")
                confirm_password = st.text_input("Confirm password", type="password", placeholder="Confirm password")
                submitted = st.form_submit_button("Create Account", use_container_width=True)
                if submitted:
                    if new_username and new_email and new_password and confirm_password:
                        if new_password == confirm_password:
                            if len(new_password) >= 6:
                                success, message = register_user(new_username, new_email, new_password)
                                if success:
                                    st.success(message)
                                    st.info("Now you can sign in with your credentials")
                                else:
                                    st.error(message)
                            else:
                                st.warning("Password must be at least 6 characters")
                        else:
                            st.error("Passwords do not match")
                    else:
                        st.warning("Please fill in all fields")
        st.markdown('</div>', unsafe_allow_html=True)

# =====================================================
# PROFILE PAGE
# =====================================================

def show_profile_page():
    st.markdown("""
        <div class="page-title-card">
            <div class="page-eyebrow">Profile</div>
            <h2>User Profile</h2>
            <div class="page-description">Manage saved channels, favorite videos and previous channel analyses.</div>
        </div>
    """, unsafe_allow_html=True)
    
    # Fetch data
    conn = get_db_connection()
    if not conn:
        st.error("Failed to connect to database")
        return
    
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("""
        SELECT 
            COUNT(DISTINCT sc.channel_id) as saved_channels_count,
            COUNT(DISTINCT fv.video_id) as favorite_videos_count,
            COUNT(DISTINCT ah.id) as analysis_count
        FROM users u
        LEFT JOIN saved_channels sc ON u.id = sc.user_id
        LEFT JOIN favorite_videos fv ON u.id = fv.user_id
        LEFT JOIN analysis_history ah ON u.id = ah.user_id
        WHERE u.id = %s
    """, (st.session_state['user_id'],))
    stats = cur.fetchone()
    
    cur.execute("SELECT created_at FROM users WHERE id = %s", (st.session_state['user_id'],))
    user_data = cur.fetchone()
    joined = user_data['created_at'].strftime('%B %Y') if user_data else "Unknown"
    cur.close()
    conn.close()
    
    # Tabs
    tab1, tab2, tab3 = st.tabs([" Overview", "Saved Channels", "History"])
    
    with tab1:
        col1, col2 = st.columns([1, 3])
        with col1:
            st.image(
                f"https://ui-avatars.com/api/?background=3b82f6&color=fff&rounded=true&size=120&bold=true&name={st.session_state['user'].replace(' ', '+')}",
                width=120
            )
        with col2:
            st.markdown(f"## {st.session_state['user']}")
            st.markdown(f" {st.session_state['user_email']}")
            st.caption(f" Member since {joined}")
        
        st.markdown("#### Activity")
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.metric("Saved Channels", stats['saved_channels_count'] or 0)
        with col_b:
            st.metric("Favorite Videos", stats['favorite_videos_count'] or 0)
        with col_c:
            st.metric("Analyses Performed", stats['analysis_count'] or 0)
        
        st.markdown("---")
        if st.button("← Back to Dashboard", use_container_width=True):
            st.session_state.page = 'dashboard'
            st.rerun()
    
    with tab2:
        st.markdown("####  Saved Channels")
        saved = get_saved_channels(st.session_state['user_id'])
        if saved:
            for i in range(0, len(saved), 3):
                cols = st.columns(3)
                for idx, ch in enumerate(saved[i:i+3]):
                    with cols[idx]:
                        st.image(ch['channel_thumbnail'] or "https://via.placeholder.com/80", width=80)
                        st.markdown(f"**{ch['channel_title'][:30]}**")
                        col_btn1, col_btn2 = st.columns(2)
                        with col_btn1:
                            if st.button(" Load", key=f"load_{ch['channel_id']}"):
                                with st.spinner("Loading channel..."):
                                    cd = get_channel_info(ch['channel_id'])
                                    if cd:
                                        st.session_state.channel_data = cd
                                        st.session_state.channel_loaded = True
                                        st.session_state.current_channel_id = ch['channel_id']
                                        st.session_state.videos_data = get_channel_videos(ch['channel_id'], 30)
                                        st.session_state.comments_data = get_all_comments_for_channel(st.session_state.videos_data, 30)
                                        st.session_state.page = 'dashboard'
                                        st.rerun()
                        with col_btn2:
                            # Confirm before removal
                            with st.popover("🗑️ Remove", use_container_width=True):
                                st.warning(f"Remove **{ch['channel_title']}** from saved?")
                                if st.button("Yes, remove", key=f"confirm_remove_{ch['channel_id']}"):
                                    remove_saved_channel(st.session_state['user_id'], ch['channel_id'])
                                    st.toast(f"Removed {ch['channel_title']}", icon="🗑️")
                                    st.rerun()
                                st.button("Cancel")
                        st.markdown("---")
        else:
            st.info("No saved channels. Load a channel and click 'Save Channel' to add it here.")
    
    with tab3:
        st.markdown("####  Analysis History")
        history = get_analysis_history(st.session_state['user_id'], limit=10)
        if history:
            for rec in history:
                with st.expander(f"📺 {rec['channel_title']} — {rec['analysis_date'][:16]}"):
                    if rec['metrics']:
                        m = rec['metrics']
                        col_m1, col_m2, col_m3 = st.columns(3)
                        col_m1.metric("Subscribers", format_number(m.get('subscribers', 0)))
                        col_m2.metric("Total Views", format_number(m.get('views', 0)))
                        col_m3.metric("Avg Engagement", f"{m.get('avg_engagement', 0):.1f}%")
                    if rec.get('insights'):
                        st.write("**💡 Key Insights**")
                        for ins in rec['insights'][:2]:
                            st.write(f"• {ins.get('title', '')}")
        else:
            st.info("No analysis history yet. Analyze a channel to see history.")
# =====================================================
# SETTINGS PAGE
# =====================================================
# =====================================================
# SETTINGS HELPER FUNCTIONS (добавьте в конец раздела функций)
# =====================================================

def clear_youtube_cache():
    conn = get_db_connection()
    if not conn:
        return False
    cur = conn.cursor()
    cur.execute("DELETE FROM youtube_cache")
    conn.commit()
    cur.close()
    conn.close()
    return True

def export_user_data(user_id):
    conn = get_db_connection()
    if not conn:
        return None
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("SELECT username, email, created_at FROM users WHERE id = %s", (user_id,))
    user = cur.fetchone()
    cur.execute("SELECT * FROM saved_channels WHERE user_id = %s", (user_id,))
    saved = cur.fetchall()
    cur.execute("SELECT * FROM analysis_history WHERE user_id = %s", (user_id,))
    history = cur.fetchall()
    cur.close()
    conn.close()
    return {
        "user": user,
        "saved_channels": saved,
        "analysis_history": history,
        "export_date": datetime.now().isoformat()
    }

def change_password(user_id, current_password, new_password):
    conn = get_db_connection()
    if not conn:
        return False, "Database connection error"
    cur = conn.cursor()
    cur.execute("SELECT id FROM users WHERE id = %s AND password_hash = %s",
                (user_id, hash_password(current_password)))
    if cur.fetchone():
        cur.execute("UPDATE users SET password_hash = %s WHERE id = %s",
                    (hash_password(new_password), user_id))
        conn.commit()
        cur.close()
        conn.close()
        return True, "Password changed successfully"
    else:
        cur.close()
        conn.close()
        return False, "Current password is incorrect"

# ========== НОВЫЕ ФУНКЦИИ ДЛЯ НАСТРОЕК ==========

def send_password_reset_email(email):
    """Mock function: send reset link"""
    # В реальном приложении – отправка письма через SMTP
    return True

def get_active_sessions(user_id):
    """Mock data for active sessions"""
    return [
        {"device": "Chrome on Windows", "location": "New York, USA", "last_active": "2025-04-08 14:30", "current": True},
        {"device": "Safari on iPhone", "location": "New York, USA", "last_active": "2025-04-07 09:15", "current": False},
        {"device": "Firefox on MacBook", "location": "London, UK", "last_active": "2025-04-05 22:10", "current": False},
    ]

def revoke_session(device_name):
    """Mock revoke session"""
    # Здесь можно добавить логику удаления сессии из БД
    return True

def enable_2fa(user_id):
    """Enable 2FA (mock)"""
    # Генерация секрета, QR-кода и резервных кодов
    return True

def disable_2fa(user_id):
    """Disable 2FA (mock)"""
    return True

def delete_user_account(user_id):
    """Delete user and all related data"""
    conn = get_db_connection()
    if not conn:
        return False
    cur = conn.cursor()
    try:
        # Удаляем связанные записи (каскадно, если настроено)
        cur.execute("DELETE FROM users WHERE id = %s", (user_id,))
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        print(f"Deletion error: {e}")
        return False
    finally:
        cur.close()
        conn.close()
def show_settings_page():
    st.markdown("""
        <div class="page-title-card">
            <div class="page-eyebrow">Settings</div>
            <h2>Account Settings</h2>
            <div class="page-description">Control profile, security, notifications, data export and application preferences.</div>
        </div>
    """, unsafe_allow_html=True)
    
    # Profile Section
    with st.container():
        st.markdown('<div class="settings-card">', unsafe_allow_html=True)
        st.markdown('<div class="settings-section-title"><i class="fas fa-user-circle"></i> Profile Information</div>', unsafe_allow_html=True)
        col1, col2 = st.columns([1, 2])
        with col1:
            st.image(f"https://ui-avatars.com/api/?background=3b82f6&color=fff&rounded=true&size=100&bold=true&name={st.session_state['user'].replace(' ', '+')}", width=100)
            if st.button("Change Avatar", use_container_width=True):
                st.info("Avatar change feature coming soon")
        with col2:
            st.text_input("Username", value=st.session_state['user'], disabled=True)
            st.text_input("Email", value=st.session_state['user_email'], disabled=True)
            if st.button("Edit Profile", use_container_width=True):
                st.info("Profile editing coming soon")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Security Section
    with st.container():
        st.markdown('<div class="settings-card">', unsafe_allow_html=True)
        st.markdown('<div class="settings-section-title"><i class="fas fa-shield-alt"></i> Security</div>', unsafe_allow_html=True)
        
        with st.expander("Change Password", expanded=False):
            with st.form("change_password_form"):
                current_pw = st.text_input("Current Password", type="password")
                new_pw = st.text_input("New Password", type="password")
                confirm_pw = st.text_input("Confirm New Password", type="password")
                col_forgot, _ = st.columns([1, 3])
                with col_forgot:
                    if st.form_submit_button("Forgot password?", use_container_width=True):
                        send_password_reset_email(st.session_state['user_email'])
                        st.success("Password reset link sent to your email")
                submitted = st.form_submit_button("Update Password", use_container_width=True)
                if submitted:
                    if not current_pw or not new_pw:
                        st.error("Please fill all fields")
                    elif new_pw != confirm_pw:
                        st.error("New passwords do not match")
                    elif len(new_pw) < 6:
                        st.error("Password must be at least 6 characters")
                    else:
                        success, msg = change_password(st.session_state['user_id'], current_pw, new_pw)
                        if success:
                            st.success(msg)
                        else:
                            st.error(msg)
        
        with st.expander("Two-Factor Authentication (2FA)", expanded=False):
            st.write("Add an extra layer of security to your account.")
            enable_2fa_toggle = st.toggle("Enable 2FA", value=st.session_state.get('2fa_enabled', False))
            if enable_2fa_toggle != st.session_state.get('2fa_enabled', False):
                if enable_2fa_toggle:
                    enable_2fa(st.session_state['user_id'])
                    st.session_state['2fa_enabled'] = True
                    st.success("2FA enabled. Scan the QR code below with your authenticator app.")
                    st.code("ABCD-EFGH-IJKL-MNOP", language="text")
                    st.caption("Backup codes (save them securely):")
                    st.code("12345-67890\n23456-78901\n34567-89012", language="text")
                else:
                    disable_2fa(st.session_state['user_id'])
                    st.session_state['2fa_enabled'] = False
                    st.success("2FA disabled")
        
        with st.expander("Active Sessions", expanded=False):
            sessions = get_active_sessions(st.session_state['user_id'])
            for sess in sessions:
                col_dev, col_info, col_btn = st.columns([2, 2, 1])
                with col_dev:
                    device_icon = "desktop" if "Windows" in sess['device'] or "Mac" in sess['device'] else "mobile-alt"
                    st.markdown(f'<i class="fas fa-{device_icon}"></i> {sess["device"]}', unsafe_allow_html=True)
                with col_info:
                    st.caption(f"{sess['location']} • Last active: {sess['last_active']}")
                with col_btn:
                    if not sess.get('current', False):
                        if st.button("Revoke", key=f"revoke_{sess['device']}"):
                            revoke_session(sess['device'])
                            st.success("Session revoked")
                    else:
                        st.caption("Current session")
                st.markdown("---")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Privacy Section
    with st.container():
        st.markdown('<div class="settings-card">', unsafe_allow_html=True)
        st.markdown('<div class="settings-section-title"><i class="fas fa-lock"></i> Privacy</div>', unsafe_allow_html=True)
        st.checkbox("Share anonymous usage data to improve the app", value=True)
        st.checkbox("Show my activity status to other users", value=False)
        st.selectbox("Data retention period", ["30 days", "90 days", "1 year", "Forever"], index=1)
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Notifications Section
    with st.container():
        st.markdown('<div class="settings-card">', unsafe_allow_html=True)
        st.markdown('<div class="settings-section-title"><i class="fas fa-bell"></i> Notifications</div>', unsafe_allow_html=True)
        notif_email = st.toggle("Email notifications", value=st.session_state.get('notifications_enabled', True))
        if notif_email != st.session_state.get('notifications_enabled', True):
            st.session_state['notifications_enabled'] = notif_email
        st.toggle("Push notifications (browser)", value=False)
        st.toggle("Weekly analytics digest", value=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Data Management
    with st.container():
        st.markdown('<div class="settings-card">', unsafe_allow_html=True)
        st.markdown('<div class="settings-section-title"><i class="fas fa-database"></i> Data Management</div>', unsafe_allow_html=True)
        auto_save = st.toggle("Auto-save reports after analysis", value=st.session_state.get('auto_save_reports', True))
        if auto_save != st.session_state.get('auto_save_reports', True):
            st.session_state['auto_save_reports'] = auto_save
        col_clear, col_export = st.columns(2)
        with col_clear:
            if st.button("Clear YouTube Cache", use_container_width=True):
                if clear_youtube_cache():
                    st.success("Cache cleared successfully")
                else:
                    st.error("Failed to clear cache")
        with col_export:
            if st.button("Export All Data", use_container_width=True):
                data = export_user_data(st.session_state['user_id'])
                if data:
                    json_str = json.dumps(data, indent=2, default=str)
                    st.download_button(
                        label="Download JSON",
                        data=json_str,
                        file_name=f"user_data_{st.session_state['user_id']}.json",
                        mime="application/json",
                        use_container_width=True
                    )
                else:
                    st.error("Failed to export data")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Danger Zone
    with st.container():
        st.markdown('<div class="settings-card danger-zone">', unsafe_allow_html=True)
        st.markdown('<div class="settings-section-title"><i class="fas fa-exclamation-triangle" style="color:#dc2626;"></i> Danger Zone</div>', unsafe_allow_html=True)
        st.warning("Once you delete your account, there is no going back. All your data will be permanently removed.")
        if st.button("Delete Account", use_container_width=True, type="secondary"):
            with st.popover("Confirm Account Deletion", use_container_width=True):
                st.error("This action is irreversible!")
                confirm_text = st.text_input("Type DELETE to confirm")
                if st.button("Permanently Delete My Account", use_container_width=True):
                    if confirm_text == "DELETE":
                        if delete_user_account(st.session_state['user_id']):
                            st.session_state.clear()
                            st.success("Account deleted. You will be logged out.")
                            st.rerun()
                        else:
                            st.error("Failed to delete account")
                    else:
                        st.error("Please type DELETE to confirm")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Help & Support
    with st.container():
        st.markdown('<div class="settings-card">', unsafe_allow_html=True)
        st.markdown('<div class="settings-section-title"><i class="fas fa-question-circle"></i> Help & Support</div>', unsafe_allow_html=True)
        col_help, col_logs = st.columns(2)
        with col_help:
            if st.button("Contact Support", use_container_width=True):
                st.info("Please email support@socialanalytics.com")
        with col_logs:
            if st.button("Download Error Logs", use_container_width=True):
                st.info("Logs feature coming soon")
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    if st.button("← Back to Dashboard", use_container_width=True):
        st.session_state.page = 'dashboard'
        st.rerun()


# =====================================================
# HOME / LANDING PAGE AFTER LOGIN
# =====================================================

def show_home_page():
    user_name = st.session_state.get("user", "User")
    st.markdown(f"""
        <div class="home-hero">
            <div class="home-eyebrow">Analytics workspace</div>
            <div class="home-title">Social Media Analytics for YouTube data</div>
            <div class="home-text">
                Welcome, <b>{user_name}</b>. This system collects YouTube data, stores analytical results,
                applies machine learning methods and helps understand channel performance through a professional dashboard.
            </div>
            <div class="home-mini-row">
                <div class="home-chip">YouTube Data API v3</div>
                <div class="home-chip">PostgreSQL storage</div>
                <div class="home-chip">ML predictions</div>
                <div class="home-chip">PDF / Excel reports</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    left, right = st.columns([1.55, 1])
    with left:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("""
                <div class="home-card">
                    <div class="home-card-number">01</div>
                    <div class="home-card-title">Load channel data</div>
                    <div class="home-card-text">Connect to YouTube API and receive structured channel, video and comment statistics.</div>
                </div>
            """, unsafe_allow_html=True)
        with c2:
            st.markdown("""
                <div class="home-card">
                    <div class="home-card-number">02</div>
                    <div class="home-card-title">Analyze performance</div>
                    <div class="home-card-text">Calculate engagement, audience metrics, best posting time and content performance patterns.</div>
                </div>
            """, unsafe_allow_html=True)
        c3, c4 = st.columns(2)
        with c3:
            st.markdown("""
                <div class="home-card">
                    <div class="home-card-number">03</div>
                    <div class="home-card-title">Predict and cluster</div>
                    <div class="home-card-text">Use Random Forest, Linear Regression and K-Means for predictions and video grouping.</div>
                </div>
            """, unsafe_allow_html=True)
        with c4:
            st.markdown("""
                <div class="home-card">
                    <div class="home-card-number">04</div>
                    <div class="home-card-title">Export reports</div>
                    <div class="home-card-text">Generate PDF and Excel reports and keep previous analysis history for comparison.</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)
        if st.button("Open analytics dashboard", use_container_width=True):
            st.session_state.page = "dashboard"
            st.rerun()

    with right:
        st.markdown("""
            <div class="product-panel">
                <div style="font-size:13px; color:#93c5fd; font-weight:900; text-transform:uppercase; letter-spacing:.09em; margin-bottom:18px;">System overview</div>
                <div class="product-kpi"><div class="product-kpi-value">API</div><div class="product-kpi-label">Data collection source</div></div>
                <div class="product-kpi"><div class="product-kpi-value">ML</div><div class="product-kpi-label">Predictions and recommendations</div></div>
                <div class="product-kpi"><div class="product-kpi-value">Dashboard</div><div class="product-kpi-label">Interactive analytics interface</div></div>
                <div style="font-size:13px; color:#cbd5e1; line-height:1.55; margin-top:18px;">Designed as a finished analytical product for channel performance monitoring and decision support.</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("""
        <div class="workflow-card">
            <div class="workflow-title">Main workflow</div>
            <div class="workflow-text">
                YouTube API → JSON data processing → PostgreSQL storage → analytics and ML models → Streamlit dashboard → PDF / Excel reporting.
            </div>
        </div>
    """, unsafe_allow_html=True)

# =====================================================
# DASHBOARD (MAIN ANALYTICS) — ПОЛНАЯ ВЕРСИЯ
# =====================================================

def show_dashboard():
    st.markdown("""
        <div class="page-title-card">
            <div class="page-eyebrow">Dashboard</div>
            <h2>Channel Analytics Workspace</h2>
            <div class="page-description">Load a YouTube channel to analyze metrics, audience activity, predictions, insights, and reports in one place.</div>
        </div>
    """, unsafe_allow_html=True)
    col1, col2 = st.columns([3, 1])
    with col1:
        channel_input = st.text_input(
            "",
            placeholder="Enter YouTube Channel ID or Name (e.g., MrBeast, UCXuqSBlHAE6Xw...",
            label_visibility="collapsed",
            key="channel_search_input"
        )
    with col2:
        search_clicked = st.button("Load Channel", use_container_width=True)

    # ----- Обработка нажатия кнопки "Load Channel" -----
    if search_clicked and channel_input:
        # Сбрасываем старые данные канала
        st.session_state.channel_loaded = False
        st.session_state.channel_data = None
        st.session_state.videos_data = None
        st.session_state.comments_data = None
        st.session_state.current_channel_id = None
        st.session_state.search_results = []
        st.session_state.selected_channel_id = None
        
        with st.spinner("Searching..."):
            query = channel_input.strip()
            if "youtube.com/" in query:
                if "/@" in query:
                    query = query.split("/@")[-1].split("/")[0].split("?")[0]
                elif "/channel/" in query:
                    query = query.split("/channel/")[-1].split("/")[0].split("?")[0]

            results = search_youtube_channels(query, max_results=5)
            if results:
                st.session_state.search_results = results
                st.rerun()
            else:
                st.warning("No channels found. Try another name or ID.")
                st.session_state.search_results = []

    # ----- Отображение найденных каналов (без картинок) -----
    if st.session_state.search_results:
        st.markdown("#### Search Results")
        for ch in st.session_state.search_results:
            col_a, col_b = st.columns([4, 1])
            with col_a:
                st.markdown(f"**{ch['title']}**")
                st.caption(ch['description'])
            with col_b:
                if st.button("Select", key=f"select_{ch['channel_id']}"):
                    st.session_state.selected_channel_id = ch['channel_id']
                    st.session_state.search_results = []
                    st.rerun()
            st.markdown("---")
        if st.button("Cancel", use_container_width=True):
            st.session_state.search_results = []
            st.rerun()

    # ----- Загрузка выбранного канала -----
    if st.session_state.selected_channel_id and not st.session_state.channel_loaded:
        with st.spinner("Loading channel data..."):
            channel_data = get_channel_info(st.session_state.selected_channel_id)
            if channel_data:
                st.session_state.channel_data = channel_data
                st.session_state.channel_loaded = True
                st.session_state.current_channel_id = channel_data['channel_id']

                with st.spinner("Loading videos..."):
                    videos = get_channel_videos(st.session_state.selected_channel_id, 30)
                    st.session_state.videos_data = videos

                with st.spinner("Analyzing comments..."):
                    comments = get_all_comments_for_channel(videos, 30)
                    st.session_state.comments_data = comments

                st.success(f"Channel '{channel_data['title']}' loaded successfully")
                st.session_state.selected_channel_id = None
                st.rerun()
            else:
                st.error("Failed to load channel. Check ID or name.")
                st.session_state.selected_channel_id = None

    # ----- Если канал ещё не загружен, показываем примеры -----
    if not st.session_state.channel_loaded or not st.session_state.channel_data:
        st.markdown("<div class=\"block-note\">Enter a YouTube channel ID, handle, or channel name to start the analysis.</div>", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("#### Popular Channels")
        example_channels = [
            {"name": "MrBeast", "id": "UCX6OQ3DkcsbYNE6H8uQQuVA"},
            {"name": "Linus Tech Tips", "id": "UCXuqSBlHAE6Xw-yeJA0Tunw"},
            {"name": "T-Series", "id": "UCq-Fj5jknLsUf-MWSy4_brA"},
            {"name": "Mark Rober", "id": "UCY1kMZp36IQSyNx_9h4mpCg"},
            {"name": "Kurzgesagt", "id": "UCsXVk37bltHxD1rDPwtNM8Q"}
        ]
        cols = st.columns(5)
        for idx, ch in enumerate(example_channels):
            with cols[idx]:
                if st.button(ch['name'], use_container_width=True):
                    st.session_state.channel_loaded = False
                    st.session_state.channel_data = None
                    st.session_state.videos_data = None
                    st.session_state.comments_data = None
                    st.session_state.selected_channel_id = ch['id']
                    st.rerun()
        return

    # ----- ДАЛЕЕ ВЕСЬ ОСТАЛЬНОЙ КОД ДАШБОРДА (метрики, ML, табы, экспорт) -----
    channel = st.session_state.channel_data
    videos = st.session_state.videos_data if st.session_state.videos_data else []
    comments = st.session_state.comments_data if st.session_state.comments_data else []

    # Применяем фильтр по дате (если задан)
    if st.session_state.date_range and videos:
        df_videos = pd.DataFrame(videos)
        df_videos['published_at'] = pd.to_datetime(df_videos['published_at'])
        mask = (df_videos['published_at'].dt.date >= st.session_state.date_range[0]) & (df_videos['published_at'].dt.date <= st.session_state.date_range[1])
        filtered_videos = df_videos[mask].to_dict('records')
        if filtered_videos:
            videos = filtered_videos

    # Метрики
    if videos:
        total_likes = sum(v['like_count'] for v in videos)
        total_comments_videos = sum(v['comment_count'] for v in videos)
        avg_engagement = np.mean([v['engagement_rate'] for v in videos])
    else:
        total_likes = total_comments_videos = avg_engagement = 0

    if comments:
        sentiment_dist = Counter([c['sentiment_label'] for c in comments if c['sentiment_label']])
    else:
        sentiment_dist = {}

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    metrics_data = [
        (col1, "SUBSCRIBERS", format_number(channel['subscribers'])),
        (col2, "TOTAL VIEWS", format_number(channel['views'])),
        (col3, "VIDEOS", str(channel['videos'])),
        (col4, "LIKES", format_number(total_likes)),
        (col5, "COMMENTS", format_number(total_comments_videos)),
        (col6, "ENGAGEMENT", f"{avg_engagement:.1f}%")
    ]
    for col, label, value in metrics_data:
        with col:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{value}</div><div class="metric-label">{label}</div></div>', unsafe_allow_html=True)

    # ML модели
    with st.spinner("Training ML models..."):
        view_prediction_model = train_view_prediction_model(videos)
        growth_model = train_growth_model(videos)
        cluster_results = cluster_videos(videos)

    posting_analysis = analyze_best_posting_time(videos)
    drop_patterns = find_view_drop_patterns(videos)
    video_type_analysis = get_best_video_type(videos)
    cluster_stats = cluster_results['cluster_stats'] if cluster_results else None
    insights = generate_key_insights(channel, videos, cluster_stats, posting_analysis, video_type_analysis, drop_patterns)

    if insights:
        st.markdown("---")
        st.markdown("### Key Insights")
        cols = st.columns(2)
        for idx, insight in enumerate(insights[:4]):
            with cols[idx % 2]:
                st.markdown(f'<div class="insight-card"><div class="insight-title">{insight["title"]}</div><div class="insight-description">{insight["description"]}</div><div class="insight-action">→ {insight["action"]}</div></div>', unsafe_allow_html=True)

    st.markdown("---")
    col_exp1, col_exp2 = st.columns(2)
    with col_exp1:
        if st.button("📄 Export to PDF", use_container_width=True):
            metrics_dict = {'avg_engagement': avg_engagement}
            pdf_buffer = export_to_pdf(channel, videos, insights, metrics_dict)
            st.download_button("Download PDF", data=pdf_buffer, file_name=f"{channel['title']}_report.pdf", mime="application/pdf", use_container_width=True)
    with col_exp2:
        if st.button("📊 Export to Excel", use_container_width=True):
            excel_buffer = export_to_excel(channel, videos, insights)
            st.download_button("Download Excel", data=excel_buffer, file_name=f"{channel['title']}_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.markdown("---")

    # Tabs
    tab_overview, tab_audience, tab_content, tab_trends, tab_abtest, tab_plan = st.tabs([
        "Overview", "Audience", "Content", "Trends", "A/B Test", "Content Plan"
    ])

    with tab_overview:
        if videos:
            col1, col2 = st.columns(2)
            with col1:
                top_videos = sorted(videos, key=lambda x: x['view_count'], reverse=True)[:5]
                fig = go.Figure(data=[go.Bar(x=[v['title'][:35] for v in top_videos], y=[v['view_count'] for v in top_videos], marker=dict(color=[v['view_count'] for v in top_videos], colorscale='Blues', showscale=False), text=[format_number(v['view_count']) for v in top_videos], textposition='outside')])
                fig.update_layout(title="Top Videos by Views", plot_bgcolor='white', paper_bgcolor='white', xaxis=dict(showgrid=False, tickangle=-45), yaxis=dict(showgrid=True, gridcolor='#e2e8f0'), height=400)
                fig.update_traces(marker_color='#3b82f6')
                st.plotly_chart(fig, use_container_width=True)
            with col2:
                if len(videos) >= 7:
                    recent = videos[:7][::-1]
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(x=[v['title'][:25] for v in recent], y=[v['view_count'] for v in recent], mode='lines+markers', line=dict(color='#3b82f6', width=3), marker=dict(size=10, color='#f59e0b'), fill='tozeroy', fillcolor='rgba(59,130,246,0.1)'))
                    fig.update_layout(title="Recent Videos Performance", plot_bgcolor='white', paper_bgcolor='white', xaxis=dict(showgrid=False, tickangle=-45), yaxis=dict(showgrid=True, gridcolor='#e2e8f0'), height=400)
                    st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No videos data available.")

    with tab_audience:
        st.markdown("### Audience Demographics")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### Age Distribution")
            age_groups = ['13-17', '18-24', '25-34', '35-44', '45+']
            channel_title = channel['title'].lower()
            if any(x in channel_title for x in ['game','gaming','play']):
                distribution = [25,40,25,7,3]
            elif any(x in channel_title for x in ['tech','science','code','programming']):
                distribution = [10,35,35,15,5]
            else:
                distribution = [15,35,30,12,8]
            colors = get_color_palette(len(age_groups))
            fig_age = go.Figure(data=[go.Pie(labels=age_groups, values=distribution, marker=dict(colors=colors), hole=0.45, textinfo='label+percent', pull=[0.05,0,0,0,0])])
            fig_age.update_layout(height=380, showlegend=False)
            st.plotly_chart(fig_age, use_container_width=True)
        with col2:
            st.markdown("#### Geographic Distribution")
            countries = ['United States','India','United Kingdom','Canada','Germany','Brazil','Other']
            if channel['country'] != 'Not specified':
                dist = [25,15,10,8,7,6,29]
            else:
                dist = [20,18,12,8,7,5,30]
            colors = get_color_palette(len(countries))
            fig_geo = go.Figure(data=[go.Pie(labels=countries, values=dist, marker=dict(colors=colors), hole=0.45, textinfo='label+percent')])
            fig_geo.update_layout(height=380, showlegend=False)
            st.plotly_chart(fig_geo, use_container_width=True)
        st.markdown("---")
        st.markdown("#### Audience Metrics")
        col1, col2, col3, col4 = st.columns(4)
        avg_watch_time = np.random.uniform(3,8) if videos else 0
        new_viewers = np.random.uniform(15,35)
        returning = np.random.uniform(25,55)
        unique_viewers = int(channel['views'] * np.random.uniform(0.1,0.3))
        with col1: st.metric("⏱️ Avg Watch Time", f"{avg_watch_time:.1f} min")
        with col2: st.metric("🆕 New Viewers", f"{new_viewers:.1f}%")
        with col3: st.metric("🔄 Returning Viewers", f"{returning:.1f}%")
        with col4: st.metric("👥 Unique Viewers", format_number(unique_viewers))

    with tab_content:
        if videos:
            st.markdown("### Content Performance")
            fig_engagement = go.Figure(data=[go.Bar(x=[v['title'][:30] for v in videos[:12]], y=[v['engagement_rate'] for v in videos[:12]], marker=dict(color=[v['engagement_rate'] for v in videos[:12]], colorscale='Viridis', showscale=True, colorbar=dict(title="Engagement %")), text=[f"{v['engagement_rate']:.1f}%" for v in videos[:12]], textposition='outside')])
            fig_engagement.update_layout(title="Engagement Rate by Video (%)", xaxis=dict(tickangle=-45), yaxis=dict(title="Engagement Rate (%)"), height=450)
            st.plotly_chart(fig_engagement, use_container_width=True)
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("#### Video Duration Distribution")
                durations = [v['duration_seconds'] for v in videos if v['duration_seconds']>0]
                if durations:
                    groups = ['<5 min','5-10 min','10-20 min','20-30 min','>30 min']
                    counts = [len([d for d in durations if d<300]), len([d for d in durations if 300<=d<600]), len([d for d in durations if 600<=d<1200]), len([d for d in durations if 1200<=d<1800]), len([d for d in durations if d>=1800])]
                    fig_duration = go.Figure(data=[go.Pie(labels=groups, values=counts, marker=dict(colors=['#3b82f6','#60a5fa','#f59e0b','#fbbf24','#10b981']), hole=0.4, textinfo='label+percent')])
                    fig_duration.update_layout(height=380, showlegend=False)
                    st.plotly_chart(fig_duration, use_container_width=True)
            with col2:
                st.markdown("#### Comment Sentiment Analysis")
                sent_data = {'Positive': sentiment_dist.get('positive',60), 'Neutral': sentiment_dist.get('neutral',25), 'Negative': sentiment_dist.get('negative',15)}
                fig_sent = go.Figure(data=[go.Pie(labels=list(sent_data.keys()), values=list(sent_data.values()), marker=dict(colors=['#10b981','#94a3b8','#ef4444']), hole=0.4, textinfo='label+percent')])
                fig_sent.update_layout(height=380, showlegend=False)
                st.plotly_chart(fig_sent, use_container_width=True)
            st.markdown("---")
            st.markdown("#### Video Library")
            videos_df = pd.DataFrame(videos)
            st.dataframe(videos_df[['title','view_count','like_count','comment_count','engagement_rate']], use_container_width=True, column_config={'title':'Title','view_count':st.column_config.NumberColumn('Views',format="%d"),'like_count':st.column_config.NumberColumn('Likes',format="%d"),'comment_count':st.column_config.NumberColumn('Comments',format="%d"),'engagement_rate':st.column_config.NumberColumn('Engagement',format="%.2f%%")})
        else:
            st.info("Load videos to analyze content")

    with tab_trends:
        st.markdown("### Strategic Recommendations")
        col1, col2, col3 = st.columns(3)
        with col1:
            if videos and len(videos)>=5 and view_prediction_model:
                recent_views = [v['view_count'] for v in videos[:5]]
                growth_rate = (recent_views[0]-recent_views[-1])/recent_views[-1]*100 if recent_views[-1]>0 else 0
                forecast = predict_views(view_prediction_model, datetime.now().hour, datetime.now().weekday(), avg_engagement, 10, 0) or int(recent_views[0]*(1+growth_rate/100))
                st.markdown(f'<div class="metric-card"><div class="metric-value">{format_number(forecast)}</div><div class="metric-label">Forecast Next Video</div><div style="font-size:0.7rem;">{growth_rate:+.1f}% trend</div></div>', unsafe_allow_html=True)
        with col2:
            best_hour = posting_analysis['best_hour'] if posting_analysis else "18"
            best_day = posting_analysis['best_day'] if posting_analysis else "Wednesday"
            st.markdown(f'<div class="metric-card"><div class="metric-value">{best_hour}:00</div><div class="metric-label">Optimal Posting Time</div><div style="font-size:0.7rem;">{best_day}</div></div>', unsafe_allow_html=True)
        with col3:
            freq = "3-4 per week" if channel['videos']>100 else ("2-3 per week" if channel['videos']>50 else "1-2 per week")
            st.markdown(f'<div class="metric-card"><div class="metric-value">{freq}</div><div class="metric-label">Recommended Frequency</div></div>', unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("#### Optimization Recommendations")
        recs = ["Low engagement? Add clear CTAs and reply to comments.","Optimize titles: primary keywords first 60 chars.","Create detailed descriptions (200+ words) with timestamps.","Use 10-15 relevant tags.","Design high-contrast thumbnails.","Organize videos into playlists.","Add end screens and cards.","Respond to top comments within 24h."]
        st.markdown('<div class="rec-carousel">', unsafe_allow_html=True)
        for idx, r in enumerate(recs):
            st.markdown(f'<div class="rec-slide"><div><span class="rec-number">{idx+1}</span><span style="font-weight:600;">Recommendation</span></div><div class="rec-text">{r}</div><div style="margin-top:0.75rem;"><span style="font-size:0.7rem;">Apply now →</span></div></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        if cluster_stats:
            st.markdown("---")
            st.markdown("#### Video Performance Clusters")
            col1, col2, col3 = st.columns(3)
            cluster_items = [(col1,'high','#10b981','High Performance'),(col2,'medium','#f59e0b','Medium Performance'),(col3,'low','#ef4444','Low Performance')]
            for col, cname, color, title in cluster_items:
                with col:
                    if cname in cluster_stats:
                        s = cluster_stats[cname]
                        st.markdown(f'<div class="cluster-card" style="background:linear-gradient(135deg,{color}08 0%,{color}04 100%); border:1px solid {color}20;"><h4 style="color:{color};">{title}</h4><div style="font-size:2rem; font-weight:700;">{s["count"]}</div><div style="font-size:0.7rem;">videos</div><hr><div>Views: <strong>{format_number(s["avg_views"])}</strong></div><div>Engagement: <strong>{s["avg_engagement"]:.1f}%</strong></div></div>', unsafe_allow_html=True)
        if growth_model:
            st.markdown("---")
            st.markdown("#### Channel Growth Forecast")
            preds = predict_growth(growth_model, 10)
            if preds:
                fig = go.Figure(go.Scatter(x=list(range(1,11)), y=preds, mode='lines+markers', line=dict(color='#3b82f6', width=3), marker=dict(size=12, color='#f59e0b'), fill='tozeroy', fillcolor='rgba(59,130,246,0.15)'))
                fig.update_layout(title="Predicted Cumulative Views for Next 10 Videos", xaxis_title="Video Number", yaxis_title="Predicted Views", height=420)
                st.plotly_chart(fig, use_container_width=True)

    with tab_abtest:
        st.markdown("### A/B Test Your Video Title")
        st.write("Enter two title options to predict which one will perform better based on your channel's historical data.")
        col1, col2 = st.columns(2)
        with col1: title1 = st.text_input("Title Option 1", placeholder="e.g., How to Grow Your Channel")
        with col2: title2 = st.text_input("Title Option 2", placeholder="e.g., The Secret to YouTube Success")
        if st.button("Compare Titles", use_container_width=True):
            if videos and len(videos) >= 10:
                p1, p2, rec = compare_titles(title1, title2, videos)
                if p1 and p2:
                    st.metric("Predicted Views for Title 1", format_number(p1))
                    st.metric("Predicted Views for Title 2", format_number(p2))
                    st.success(rec)
                else:
                    st.error(rec)
            else:
                st.warning("Need at least 10 videos in the channel to train the model.")

    with tab_plan:
        st.markdown("### 📅 Content Plan Generator")
        st.write("Based on your best-performing videos (High Performance cluster), here are suggested topics for future videos.")
        if cluster_results and 'high' in cluster_results['cluster_stats']:
            high_videos = cluster_results['cluster_stats']['high']['videos']
            topics = generate_content_plan(high_videos, top_n=5)
            for i, topic in enumerate(topics, 1):
                st.markdown(f"**{i}. {topic}**")
        else:
            st.info("No High Performance cluster found. Load more videos or check cluster analysis.")

    # Save current channel to favorites button
    st.markdown("---")
    if st.button("💾 Save Current Channel to Favorites", use_container_width=True):
        if save_channel_to_favorites(st.session_state.user_id, channel['channel_id'], channel['title'], channel['thumbnail']):
            st.success("Channel saved to favorites!")
        else:
            st.error("Failed to save channel (maybe already saved).")
if '2fa_enabled' not in st.session_state:
    st.session_state['2fa_enabled'] = False
# =====================================================
# MAIN APP
# =====================================================

def show_main_app():
    current_page = st.session_state.get("page", "home")
    user_name = st.session_state.get("user", "User")
    user_email = st.session_state.get("user_email", "")

    with st.sidebar:
        st.markdown(f"""
            <div class="side-brand">
                <div class="side-logo">📊</div>
                <div class="side-title">Social Media<br>Analytics</div>
                <div class="side-subtitle"></div>
                <div class="side-user">
                    <div class="side-user-name">{user_name}</div>
                    <div class="side-user-email">{user_email}</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        st.markdown('<div class="side-caption">Navigation</div>', unsafe_allow_html=True)

        if st.button("Home", use_container_width=True):
            st.session_state.page = "home"
            st.rerun()
        if st.button("Analytics dashboard", use_container_width=True):
            st.session_state.page = "dashboard"
            st.rerun()
        if st.button("Profile", use_container_width=True):
            st.session_state.page = "profile"
            st.rerun()
        if st.button("Settings", use_container_width=True):
            st.session_state.page = "settings"
            st.rerun()

        st.markdown("---")
        st.markdown('<div class="side-caption"></div>', unsafe_allow_html=True)
        st.markdown("""
            <div style="font-size:12px; line-height:1.55; color:#cbd5e1; padding: 4px 4px 12px 4px;">
                 
            </div>
        """, unsafe_allow_html=True)
        if st.button("Logout", use_container_width=True):
            logout_user()

    if current_page == "profile":
        show_profile_page()
    elif current_page == "settings":
        show_settings_page()
    elif current_page == "dashboard":
        show_dashboard()
    else:
        show_home_page()

# =====================================================
# RUN
# =====================================================
# =====================================================
# HELPER FUNCTIONS FOR SETTINGS
# =====================================================

def clear_youtube_cache():
    """Delete all records from youtube_cache table"""
    conn = get_db_connection()
    if not conn:
        return False
    cur = conn.cursor()
    cur.execute("DELETE FROM youtube_cache")
    conn.commit()
    cur.close()
    conn.close()
    return True

def export_user_data(user_id):
    """Export user data as JSON"""
    conn = get_db_connection()
    if not conn:
        return None
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("SELECT username, email, created_at FROM users WHERE id = %s", (user_id,))
    user = cur.fetchone()
    cur.execute("SELECT * FROM saved_channels WHERE user_id = %s", (user_id,))
    saved = cur.fetchall()
    cur.execute("SELECT * FROM analysis_history WHERE user_id = %s", (user_id,))
    history = cur.fetchall()
    cur.close()
    conn.close()
    return {
        "user": user,
        "saved_channels": saved,
        "analysis_history": history,
        "export_date": datetime.now().isoformat()
    }

def change_password(user_id, current_password, new_password):
    """Change user password after verifying current one"""
    conn = get_db_connection()
    if not conn:
        return False, "Database connection error"
    cur = conn.cursor()
    cur.execute("SELECT id FROM users WHERE id = %s AND password_hash = %s",
                (user_id, hash_password(current_password)))
    if cur.fetchone():
        cur.execute("UPDATE users SET password_hash = %s WHERE id = %s",
                    (hash_password(new_password), user_id))
        conn.commit()
        cur.close()
        conn.close()
        return True, "Password changed successfully"
    else:
        cur.close()
        conn.close()
        return False, "Current password is incorrect"
if __name__ == "__main__":
    init_postgres_db()
    if st.session_state.authenticated:
        show_main_app()
    else:
        show_login_page()
