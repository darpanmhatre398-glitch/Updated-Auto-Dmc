import os
import re
import json
import shutil
import logging
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
from datetime import datetime
from bs4 import BeautifulSoup
import docx
import requests

# --- CONFIGURATION ---
OLLAMA_API_URL = "http://localhost:11434/api/generate"
OLLAMA_MODEL = "llama3.1:8b"
DOCS_DIRECTORY = "documents_to_process"
DATA_DIRECTORY = "Lake"
LOGS_DIRECTORY = "logs"
OUTPUT_DIRECTORY = "output"

# --- USER-FIXED DMC COMPONENTS ---
USER_MODEL_IDENT_CODE = "USERMODEL"
USER_SYSTEM_DIFF_CODE = "A"
USER_ASSY_CODE = "0000"
USER_ITEM_LOCATION_CODE = "D"

# --- DEFAULTS ---
DEFAULT_SYSTEM_CODE = "00"
DEFAULT_INFO_CODE = "000"

# --- SETUP LOGGING ---
if not os.path.exists(LOGS_DIRECTORY):
    os.makedirs(LOGS_DIRECTORY)
if not os.path.exists(OUTPUT_DIRECTORY):
    os.makedirs(OUTPUT_DIRECTORY)


# --- DATA PARSING FUNCTIONS ---

def parse_sns_json(file_path):
    """Parses SNS JSON files with various formats."""
    sns_data = {}
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        if isinstance(data, list):
            for item in data:
                if 'tables' in item:
                    for table in item['tables']:
                        code = table.get('system_code', '')
                        title = table.get('title', '')
                        definition = table.get('definition', '')
                        if code:
                            sns_data[code] = {
                                'title': title,
                                'definition': definition,
                                'subsystems': {}
                            }
                            for sub in table.get('subsystems', []):
                                sub_code = sub.get('subsystem_code', '').replace('-', '')
                                if sub_code and 'thru' not in sub_code:
                                    sns_data[code]['subsystems'][sub_code] = {
                                        'title': sub.get('title', ''),
                                        'definition': sub.get('definition', '')
                                    }
        
        elif isinstance(data, dict):
            for key, value in data.items():
                if isinstance(value, list):
                    for item in value:
                        code = item.get('System', '')
                        title = item.get('Title', '')
                        definition = item.get('Definition', '')
                        if code:
                            sns_data[code] = {
                                'title': title,
                                'definition': definition,
                                'subsystems': {}
                            }
                            for sub in item.get('Subsystems', []):
                                sub_code = sub.get('Subsystem', sub.get('System', '')).replace('-', '')
                                if sub_code and 'thru' not in sub_code:
                                    sns_data[code]['subsystems'][sub_code] = {
                                        'title': sub.get('Title', ''),
                                        'definition': sub.get('Definition', '')
                                    }
            
            if 'System_categories' in data:
                for category in data['System_categories']:
                    cat_code = category.get('System', '')
                    cat_title = category.get('Title', '')
                    if cat_code:
                        sns_data[cat_code] = {
                            'title': cat_title,
                            'definition': '',
                            'subsystems': {}
                        }
                        for sub in category.get('Subsystems', []):
                            sub_code = sub.get('System', '')
                            if sub_code:
                                sns_data[sub_code] = {
                                    'title': sub.get('Title', ''),
                                    'definition': sub.get('Definition', ''),
                                    'subsystems': {}
                                }
        
        return sns_data
    except Exception as e:
        return {}


def parse_info_codes_json(file_path):
    """Parses the info codes JSON file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}


def parse_info_codes_txt(file_path):
    """Parses the info codes text file."""
    info_codes = {}
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                if not line.strip():
                    continue
                match = re.match(r'^([0-9A-Z]{3})\s+([a-z]+)\s+(.*)', line.strip())
                if match:
                    code, type_, desc = match.groups()
                    info_codes[code] = {'type': type_, 'description': desc}
        return info_codes
    except:
        return {}


def prepare_context_for_llm(sns_data, info_codes):
    """Creates context string for LLM including subsystems."""
    sns_lines = []
    for code, data in sorted(sns_data.items()):
        title = data.get('title', '')
        definition = data.get('definition', '')
        if title:
            # Include definition for better context
            if definition:
                sns_lines.append(f"{code}: {title} - {definition[:100]}")
            else:
                sns_lines.append(f"{code}: {title}")
            subsystems = data.get('subsystems', {})
            for sub_code, sub_data in sorted(subsystems.items()):
                sub_title = sub_data.get('title', '')
                if sub_title and sub_code not in ['00', '0']:
                    sns_lines.append(f"  {code}-{sub_code}: {sub_title}")
    
    sns_context = "VALID SYSTEM CODES AND SUBSYSTEMS:\n" + "\n".join(sns_lines)
    
    info_by_type = {}
    for code, data in sorted(info_codes.items()):
        code_type = data.get('type', 'other')
        desc = data.get('description', '')
        if code_type not in info_by_type:
            info_by_type[code_type] = []
        # Increased from 15 to 30 for more comprehensive coverage
        if desc and len(info_by_type[code_type]) < 30:
            info_by_type[code_type].append(f"{code}: {desc}")
    
    info_lines = []
    for code_type in ['proced', 'descript', 'fault', 'process', 'sched']:
        if code_type in info_by_type:
            info_lines.append(f"\n[{code_type.upper()}]")
            info_lines.extend(info_by_type[code_type])
    
    info_context = "VALID INFO CODES:" + "\n".join(info_lines)
    
    return sns_context, info_context


def extract_text_from_docx(file_path):
    """Extracts text from a .docx file."""
    try:
        doc = docx.Document(file_path)
        headings, body = [], []
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            if para.style and para.style.name.lower().startswith('heading'):
                headings.append(text)
            else:
                body.append(text)
        return '\n'.join(headings), '\n'.join(body)
    except:
        return None, None


def generate_dmc_with_llm(headings_text, body_text, sns_context, info_context, available_sns_codes, available_info_codes):
    """Uses Ollama to determine the DMC."""
    # Use ALL headings and as much body content as practical for LLM context
    # Most LLMs can handle 8000+ chars comfortably while staying accurate
    headings_preview = headings_text if headings_text else "No headings."
    body_preview = body_text[:8000] if body_text else "No content."
    
    # Calculate how much content we're analyzing
    total_chars = len(headings_preview) + len(body_preview)
    
    prompt = f"""You are an expert in S1000D documentation standards. Analyze this COMPLETE technical document carefully and select the MOST APPROPRIATE DMC codes.

DOCUMENT TITLE/HEADINGS (COMPLETE):
{headings_preview}

DOCUMENT CONTENT (Full text - {len(body_preview)} characters):
{body_preview}

{sns_context}

{info_context}

INSTRUCTIONS:
1. Read the ENTIRE document title and content carefully
2. Identify the main system/component being discussed throughout the document
3. Determine the document type (procedure, description, fault isolation, etc.)
4. Match to the MOST SPECIFIC system code and subsystem from the valid codes above
5. Select the info code that best matches the document type and purpose
6. Consider the WHOLE document, not just the beginning

CODE SELECTION RULES:
- systemCode: 2-digit code for the main system (e.g., 20=Air Conditioning, 24=Electrical Power, 34=Navigation)
- subSystemCode: If you find a subsystem match like "24-10", use the full "10" as subSystemCode
- subSubSystemCode: Usually "0" unless a more specific sub-subsystem is identified
- infoCode: 3-character code matching document type (e.g., 000=General, 040=Description, 520=Procedure, 720=Fault Isolation)
- confidence: Your confidence level (0-100) in this selection based on the entire document
- reasoning: Brief explanation of why you chose these codes

Return ONLY this JSON (no other text):
{{"systemCode": "XX", "subSystemCode": "XX", "subSubSystemCode": "0", "infoCode": "XXX", "disassyCode": "00", "disassyCodeVariant": "A", "confidence": 85, "reasoning": "Brief explanation"}}"""
    
    try:
        payload = {
            "model": OLLAMA_MODEL,
            "prompt": prompt,
            "stream": False,
            "format": "json",
            "options": {"temperature": 0.2, "num_predict": 300}
        }
        response = requests.post(OLLAMA_API_URL, json=payload, timeout=180)
        response.raise_for_status()
        
        raw_response = response.json().get('response', '')
        if not raw_response.strip():
            return None
        
        json_text = raw_response.strip()
        if not json_text.endswith('}'):
            last_brace = json_text.rfind('}')
            if last_brace > 0:
                json_text = json_text[:last_brace + 1]
        
        if not json_text.startswith('{'):
            start = json_text.find('{')
            if start >= 0:
                json_text = json_text[start:]
        
        dmc_parts = json.loads(json_text)
        
        final_parts = {
            'systemCode': str(dmc_parts.get('systemCode', DEFAULT_SYSTEM_CODE)),
            'infoCode': str(dmc_parts.get('infoCode', DEFAULT_INFO_CODE)),
            'subSystemCode': str(dmc_parts.get('subSystemCode', '0')),
            'subSubSystemCode': str(dmc_parts.get('subSubSystemCode', '0')),
            'disassyCode': str(dmc_parts.get('disassyCode', '00')),
            'disassyCodeVariant': str(dmc_parts.get('disassyCodeVariant', 'A')),
            'confidence': dmc_parts.get('confidence', 0),
            'reasoning': dmc_parts.get('reasoning', '')
        }
        
        if available_sns_codes and final_parts['systemCode'] not in available_sns_codes:
            final_parts['systemCode'] = DEFAULT_SYSTEM_CODE
            final_parts['confidence'] = max(0, final_parts['confidence'] - 20)  # Reduce confidence if code was invalid
        
        if available_info_codes and final_parts['infoCode'] not in available_info_codes:
            final_parts['infoCode'] = DEFAULT_INFO_CODE
            final_parts['confidence'] = max(0, final_parts['confidence'] - 20)  # Reduce confidence if code was invalid
        
        return final_parts
    except:
        return None


def generate_dmc_with_fallback(headings_text, body_text, sns_data, info_codes):
    """Fallback mechanism when LLM fails."""
    full_text_lower = (headings_text + " " + body_text).lower()
    
    CATEGORY_KEYWORDS = {
        'proced': ['procedure', 'step', 'task', 'perform', 'install', 'remove', 'assemble', 'prepare', 'unpack'],
        'descript': ['description', 'overview', 'introduction', 'component', 'feature', 'specification'],
        'fault': ['fault', 'troubleshooting', 'symptom', 'remedy', 'isolation', 'failure'],
    }
    
    doc_category, max_score = None, 0
    for category, keywords in CATEGORY_KEYWORDS.items():
        score = sum(1 for kw in keywords if kw in full_text_lower)
        if score > max_score:
            max_score, doc_category = score, category
    
    filtered_info = {c: d for c, d in info_codes.items() if d.get('type') == doc_category} if doc_category else info_codes
    if not filtered_info:
        filtered_info = info_codes
    
    info_scores = []
    for code, data in filtered_info.items():
        score = sum(1 for kw in data.get('description', '').lower().split() if kw in full_text_lower)
        if score > 0:
            info_scores.append((score, code))
    
    sns_scores = []
    for code, data in sns_data.items():
        score = sum(1 for kw in data.get('title', '').lower().split() if kw in full_text_lower)
        if score > 0:
            sns_scores.append((score, code))
    
    info_scores.sort(reverse=True)
    sns_scores.sort(reverse=True)
    
    return {
        "systemCode": sns_scores[0][1] if sns_scores else DEFAULT_SYSTEM_CODE,
        "infoCode": info_scores[0][1] if info_scores else DEFAULT_INFO_CODE,
        "subSystemCode": "0",
        "subSubSystemCode": "0",
        "disassyCode": "00",
        "disassyCodeVariant": "A"
    }


def format_dmc(parts):
    return (f'DMC-{USER_MODEL_IDENT_CODE}-{USER_SYSTEM_DIFF_CODE}-{parts.get("systemCode", DEFAULT_SYSTEM_CODE)}-'
            f'{str(parts.get("subSystemCode", "0"))}{str(parts.get("subSubSystemCode", "0"))}-{USER_ASSY_CODE}-'
            f'{parts.get("disassyCode", "00")}{parts.get("disassyCodeVariant", "A")}-{parts.get("infoCode", DEFAULT_INFO_CODE)}A-{USER_ITEM_LOCATION_CODE}')


# --- GUI APPLICATION ---

class DMCAutomationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DMC Automation Tool")
        self.root.geometry("1000x850")
        self.root.configure(bg='#2b2b2b')
        self.root.resizable(True, True)
        
        # Directory paths
        self.data_directory = DATA_DIRECTORY
        self.docs_directory = DOCS_DIRECTORY
        self.output_directory = OUTPUT_DIRECTORY
        
        # DMC Component Configuration (user-editable)
        self.model_ident_code = USER_MODEL_IDENT_CODE
        self.system_diff_code = USER_SYSTEM_DIFF_CODE
        self.assy_code = USER_ASSY_CODE
        self.item_location_code = USER_ITEM_LOCATION_CODE
        
        # Data storage
        self.sns_data = {}
        self.info_codes = {}
        self.selected_sns_files = []
        self.available_sns_files = []
        self.processing = False
        self.ollama_connected = False
        
        self.setup_styles()
        self.create_widgets()
        self.check_ollama_connection()
        self.load_available_files()
    
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('TFrame', background='#2b2b2b')
        style.configure('TLabel', background='#2b2b2b', foreground='#ffffff', font=('Segoe UI', 10))
        style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'), foreground='#4fc3f7')
        style.configure('Path.TLabel', background='#2b2b2b', foreground='#aaaaaa', font=('Consolas', 9))
        style.configure('TButton', font=('Segoe UI', 10), padding=10)
        style.configure('Small.TButton', font=('Segoe UI', 9), padding=5)
        style.configure('TCheckbutton', background='#2b2b2b', foreground='#ffffff', font=('Segoe UI', 10))
        style.configure('TLabelframe', background='#2b2b2b', foreground='#ffffff')
        style.configure('TLabelframe.Label', background='#2b2b2b', foreground='#4fc3f7', font=('Segoe UI', 11, 'bold'))
        style.configure('TEntry', font=('Consolas', 10))
        style.configure('Connected.TLabel', background='#2b2b2b', foreground='#4caf50', font=('Segoe UI', 10, 'bold'))
        style.configure('Disconnected.TLabel', background='#2b2b2b', foreground='#f44336', font=('Segoe UI', 10, 'bold'))
    
    def create_widgets(self):
        # Create a canvas with scrollbar
        canvas = tk.Canvas(self.root, bg='#2b2b2b', highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        
        # Create the main frame inside the canvas
        main_frame = ttk.Frame(canvas, padding=20)
        
        # Configure the canvas
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Create window in canvas
        canvas_frame = canvas.create_window((0, 0), window=main_frame, anchor="nw")
        
        # Configure scroll region when frame changes size
        def configure_scroll_region(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Update canvas window width to match canvas width
            canvas_width = canvas.winfo_width()
            canvas.itemconfig(canvas_frame, width=canvas_width)
        
        main_frame.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", configure_scroll_region)
        
        # Bind mousewheel for scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # Store canvas reference for later use
        self.canvas = canvas
        self.main_frame = main_frame
        
        # Title and status bar
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        title_label = ttk.Label(header_frame, text="🔧 DMC Automation Tool", style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        
        # Ollama connection status
        status_frame = ttk.Frame(header_frame)
        status_frame.pack(side=tk.RIGHT, padx=(20, 0))
        
        ttk.Label(status_frame, text="Ollama:").pack(side=tk.LEFT, padx=(0, 5))
        self.ollama_status_label = ttk.Label(status_frame, text="● Checking...", style='Disconnected.TLabel')
        self.ollama_status_label.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(status_frame, text="Test", style='Small.TButton', command=self.check_ollama_connection).pack(side=tk.LEFT)
        
        # --- DMC COMPONENT CONFIGURATION SECTION ---
        dmc_config_frame = ttk.LabelFrame(main_frame, text="DMC Component Configuration", padding=10)
        dmc_config_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Create a grid layout for DMC components
        dmc_grid = ttk.Frame(dmc_config_frame)
        dmc_grid.pack(fill=tk.X)
        
        # Row 1: Model Ident Code and System Diff Code
        row1 = ttk.Frame(dmc_grid)
        row1.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(row1, text="Model Ident Code:", width=18).pack(side=tk.LEFT)
        self.model_ident_var = tk.StringVar(value=self.model_ident_code)
        model_entry = ttk.Entry(row1, textvariable=self.model_ident_var, width=15)
        model_entry.pack(side=tk.LEFT, padx=(5, 20))
        
        ttk.Label(row1, text="System Diff Code:", width=18).pack(side=tk.LEFT)
        self.system_diff_var = tk.StringVar(value=self.system_diff_code)
        system_diff_entry = ttk.Entry(row1, textvariable=self.system_diff_var, width=15)
        system_diff_entry.pack(side=tk.LEFT, padx=(5, 0))
        
        # Row 2: Assembly Code and Item Location Code
        row2 = ttk.Frame(dmc_grid)
        row2.pack(fill=tk.X)
        
        ttk.Label(row2, text="Assembly Code:", width=18).pack(side=tk.LEFT)
        self.assy_code_var = tk.StringVar(value=self.assy_code)
        assy_entry = ttk.Entry(row2, textvariable=self.assy_code_var, width=15)
        assy_entry.pack(side=tk.LEFT, padx=(5, 20))
        
        ttk.Label(row2, text="Item Location Code:", width=18).pack(side=tk.LEFT)
        self.item_location_var = tk.StringVar(value=self.item_location_code)
        item_location_entry = ttk.Entry(row2, textvariable=self.item_location_var, width=15)
        item_location_entry.pack(side=tk.LEFT, padx=(5, 0))
        
        # --- FOLDER SELECTION SECTION ---
        folders_frame = ttk.LabelFrame(main_frame, text="Folder Settings", padding=10)
        folders_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Input folder
        input_row = ttk.Frame(folders_frame)
        input_row.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(input_row, text="Input Folder:", width=12).pack(side=tk.LEFT)
        self.input_path_var = tk.StringVar(value=os.path.abspath(self.docs_directory))
        self.input_entry = ttk.Entry(input_row, textvariable=self.input_path_var, width=60)
        self.input_entry.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        ttk.Button(input_row, text="Browse...", style='Small.TButton', command=self.browse_input_folder).pack(side=tk.LEFT)
        
        # Output folder
        output_row = ttk.Frame(folders_frame)
        output_row.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(output_row, text="Output Folder:", width=12).pack(side=tk.LEFT)
        self.output_path_var = tk.StringVar(value=os.path.abspath(self.output_directory))
        self.output_entry = ttk.Entry(output_row, textvariable=self.output_path_var, width=60)
        self.output_entry.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        ttk.Button(output_row, text="Browse...", style='Small.TButton', command=self.browse_output_folder).pack(side=tk.LEFT)
        
        # Data folder (SNS files)
        data_row = ttk.Frame(folders_frame)
        data_row.pack(fill=tk.X)
        
        ttk.Label(data_row, text="Data Folder:", width=12).pack(side=tk.LEFT)
        self.data_path_var = tk.StringVar(value=os.path.abspath(self.data_directory))
        self.data_entry = ttk.Entry(data_row, textvariable=self.data_path_var, width=60)
        self.data_entry.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        ttk.Button(data_row, text="Browse...", style='Small.TButton', command=self.browse_data_folder).pack(side=tk.LEFT)
        
        # --- FILE SELECTION SECTION ---
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Left - SNS Files
        sns_frame = ttk.LabelFrame(top_frame, text="SNS Data Files", padding=10)
        sns_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        self.sns_listbox = tk.Listbox(sns_frame, selectmode=tk.MULTIPLE, height=8, 
                                       bg='#3c3c3c', fg='#ffffff', font=('Consolas', 10),
                                       selectbackground='#4fc3f7', selectforeground='#000000')
        self.sns_listbox.pack(fill=tk.BOTH, expand=True)
        
        btn_frame = ttk.Frame(sns_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(btn_frame, text="Select All", command=self.select_all_sns).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Clear", command=self.clear_sns_selection).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Refresh", command=self.load_available_files).pack(side=tk.LEFT)
        
        # Right - Documents
        docs_frame = ttk.LabelFrame(top_frame, text="Documents to Process", padding=10)
        docs_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.docs_listbox = tk.Listbox(docs_frame, height=8, 
                                        bg='#3c3c3c', fg='#ffffff', font=('Consolas', 10))
        self.docs_listbox.pack(fill=tk.BOTH, expand=True)
        
        ttk.Button(docs_frame, text="Refresh", command=self.load_documents).pack(pady=(10, 0))
        
        # Status section
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding=10)
        status_frame.pack(fill=tk.X, pady=10)
        
        self.status_label = ttk.Label(status_frame, text="Ready. Select SNS files and click 'Start Processing'")
        self.status_label.pack(anchor=tk.W)
        
        self.progress = ttk.Progressbar(status_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(10, 0))
        
        # Log section
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, 
                                                   bg='#1e1e1e', fg='#00ff00',
                                                   font=('Consolas', 9), insertbackground='#ffffff')
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Buttons section
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.start_btn = ttk.Button(button_frame, text="▶ Start Processing", command=self.start_processing)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="📂 Open Output Folder", command=self.open_output_folder).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="🗑 Clear Log", command=self.clear_log).pack(side=tk.LEFT)
        
        # Info section
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.info_label = ttk.Label(info_frame, text="Info Codes: Not loaded | SNS Systems: 0")
        self.info_label.pack(anchor=tk.W)
    
    def load_available_files(self):
        """Load available SNS files from data directory."""
        self.sns_listbox.delete(0, tk.END)
        self.available_sns_files = []
        
        data_dir = self.data_directory
        if os.path.exists(data_dir):
            for f in os.listdir(data_dir):
                if f.endswith('.json') and f != 'info_codes.json':
                    self.available_sns_files.append(f)
                    self.sns_listbox.insert(tk.END, f"  {f}")
        
        self.load_documents()
        self.log(f"Found {len(self.available_sns_files)} SNS files in '{data_dir}'")
    
    def load_documents(self):
        """Load documents from input documents directory."""
        self.docs_listbox.delete(0, tk.END)
        
        docs_dir = self.docs_directory
        if os.path.exists(docs_dir):
            docs = [f for f in os.listdir(docs_dir) if f.endswith('.docx')]
            for doc in docs:
                self.docs_listbox.insert(tk.END, f"  {doc}")
            self.log(f"Found {len(docs)} documents in '{docs_dir}'")
    
    def select_all_sns(self):
        self.sns_listbox.select_set(0, tk.END)
    
    def clear_sns_selection(self):
        self.sns_listbox.selection_clear(0, tk.END)
    
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
    
    def open_output_folder(self):
        output_dir = self.output_path_var.get()
        if os.path.exists(output_dir):
            os.startfile(output_dir)
        else:
            messagebox.showwarning("Warning", f"Output folder '{output_dir}' does not exist.")
    
    def browse_input_folder(self):
        """Browse for input documents folder."""
        folder = filedialog.askdirectory(
            title="Select Input Documents Folder",
            initialdir=self.input_path_var.get()
        )
        if folder:
            self.input_path_var.set(folder)
            self.docs_directory = folder
            self.load_documents()
            self.log(f"Input folder changed to: {folder}")
    
    def browse_output_folder(self):
        """Browse for output folder."""
        folder = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=self.output_path_var.get()
        )
        if folder:
            self.output_path_var.set(folder)
            self.output_directory = folder
            self.log(f"Output folder changed to: {folder}")
    
    def browse_data_folder(self):
        """Browse for data folder containing SNS files."""
        folder = filedialog.askdirectory(
            title="Select Data Folder (SNS Files)",
            initialdir=self.data_path_var.get()
        )
        if folder:
            self.data_path_var.set(folder)
            self.data_directory = folder
            self.load_available_files()
            self.log(f"Data folder changed to: {folder}")
    
    def check_ollama_connection(self):
        """Check if Ollama is running and accessible."""
        def check():
            try:
                self.ollama_status_label.config(text="● Checking...", style='Disconnected.TLabel')
                
                # First try a simple GET to check if server is running
                test_url = OLLAMA_API_URL.replace('/api/generate', '/api/tags')
                response = requests.get(test_url, timeout=15)
                
                if response.status_code == 200:
                    self.ollama_connected = True
                    self.ollama_status_label.config(text="● Connected", style='Connected.TLabel')
                    self.log(f"✓ Ollama connected (Model: {OLLAMA_MODEL})")
                else:
                    self.ollama_connected = False
                    self.ollama_status_label.config(text="● Disconnected", style='Disconnected.TLabel')
                    self.log(f"✗ Ollama returned status {response.status_code}")
            except requests.exceptions.Timeout:
                self.ollama_connected = False
                self.ollama_status_label.config(text="● Timeout", style='Disconnected.TLabel')
                self.log("✗ Ollama connection timeout - server may be slow or unresponsive")
            except requests.exceptions.ConnectionError:
                self.ollama_connected = False
                self.ollama_status_label.config(text="● Not Running", style='Disconnected.TLabel')
                self.log(f"✗ Cannot connect to Ollama at {OLLAMA_API_URL}")
            except Exception as e:
                self.ollama_connected = False
                self.ollama_status_label.config(text="● Error", style='Disconnected.TLabel')
                self.log(f"✗ Ollama connection error: {e}")
        
        threading.Thread(target=check, daemon=True).start()
    
    def format_dmc(self, parts):
        """Format DMC using current GUI configuration values."""
        model_ident = self.model_ident_var.get()
        system_diff = self.system_diff_var.get()
        assy_code = self.assy_code_var.get()
        item_location = self.item_location_var.get()
        
        # Handle subsystem code properly
        sub_sys = str(parts.get("subSystemCode", "0"))
        sub_sub_sys = str(parts.get("subSubSystemCode", "0"))
        
        # If subSystemCode is already 2 digits (e.g., "10"), use it as-is
        # Otherwise, concatenate subSystemCode and subSubSystemCode
        if len(sub_sys) >= 2:
            subsystem_code = sub_sys[:2]  # Take only first 2 characters
        else:
            subsystem_code = sub_sys + sub_sub_sys
        
        return (f'DMC-{model_ident}-{system_diff}-{parts.get("systemCode", DEFAULT_SYSTEM_CODE)}-'
                f'{subsystem_code}-{assy_code}-'
                f'{parts.get("disassyCode", "00")}{parts.get("disassyCodeVariant", "A")}-{parts.get("infoCode", DEFAULT_INFO_CODE)}A-{item_location}')
    
    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def start_processing(self):
        if self.processing:
            return
        
        # Check Ollama connection
        if not self.ollama_connected:
            response = messagebox.askyesno(
                "Ollama Not Connected",
                "Ollama appears to be disconnected. Processing may fail or use fallback methods only.\n\nDo you want to continue anyway?"
            )
            if not response:
                return
        
        # Get selected SNS files
        selected_indices = self.sns_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning", "Please select at least one SNS file.")
            return
        
        self.selected_sns_files = [self.available_sns_files[i] for i in selected_indices]
        
        # Start processing in a thread
        self.processing = True
        self.start_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process_documents, daemon=True).start()
    
    def process_documents(self):
        try:
            # Get current folder paths
            data_dir = self.data_directory
            docs_dir = self.docs_directory
            output_dir = self.output_directory
            
            # Create output directory if it doesn't exist
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Load info codes
            self.log("Loading info codes...")
            self.update_status("Loading info codes...")
            
            info_json = os.path.join(data_dir, "info_codes.json")
            info_txt = os.path.join(data_dir, "info_codes.txt")
            
            if os.path.exists(info_json):
                self.info_codes = parse_info_codes_json(info_json)
                self.log(f"✓ Loaded {len(self.info_codes)} info codes from JSON")
            elif os.path.exists(info_txt):
                self.info_codes = parse_info_codes_txt(info_txt)
                self.log(f"✓ Loaded {len(self.info_codes)} info codes from TXT")
            else:
                self.log("✗ No info codes file found!")
            
            # Load SNS data
            self.log("Loading SNS files...")
            self.sns_data = {}
            
            for sns_file in self.selected_sns_files:
                file_path = os.path.join(data_dir, sns_file)
                file_data = parse_sns_json(file_path)
                if file_data:
                    self.sns_data.update(file_data)
                    self.log(f"✓ Loaded {len(file_data)} systems from: {sns_file}")
            
            self.log(f"Total: {len(self.sns_data)} systems, {len(self.info_codes)} info codes")
            self.info_label.config(text=f"Info Codes: {len(self.info_codes)} | SNS Systems: {len(self.sns_data)}")
            
            # Prepare context
            sns_context, info_context = prepare_context_for_llm(self.sns_data, self.info_codes)
            available_sns = set(self.sns_data.keys())
            available_info = set(self.info_codes.keys())
            
            # Get documents
            docs = [f for f in os.listdir(docs_dir) if f.endswith('.docx')]
            if not docs:
                self.log("No documents found to process!")
                return
            
            self.progress['maximum'] = len(docs)
            self.progress['value'] = 0
            
            log_data = {
                "data_sources": {
                    "sns_files": self.selected_sns_files,
                    "total_systems": len(self.sns_data),
                    "total_info_codes": len(self.info_codes)
                },
                "successful": [],
                "failed": []
            }
            
            # Process each document
            for i, filename in enumerate(docs):
                self.update_status(f"Processing {i+1}/{len(docs)}: {filename}")
                self.log(f"\n--- Processing: {filename} ---")
                
                filepath = os.path.join(docs_dir, filename)
                headings, body = extract_text_from_docx(filepath)
                
                if not headings and not body:
                    self.log(f"✗ Could not read file")
                    log_data["failed"].append({"file": filename, "issue": "Could not read"})
                    continue
                
                # Show document statistics
                headings_len = len(headings) if headings else 0
                body_len = len(body) if body else 0
                self.log(f"📄 Document size: {headings_len} chars (headings) + {body_len} chars (body) = {headings_len + body_len} total")
                
                # Try LLM
                self.log("Querying LLM with FULL document content...")
                dmc_parts = generate_dmc_with_llm(headings, body, sns_context, info_context, available_sns, available_info)
                
                if not dmc_parts:
                    self.log("LLM failed, using fallback...")
                    dmc_parts = generate_dmc_with_fallback(headings, body, self.sns_data, self.info_codes)
                
                if dmc_parts:
                    final_dmc = self.format_dmc(dmc_parts)
                    new_filename = f"{final_dmc}.docx"
                    output_path = os.path.join(output_dir, new_filename)
                    
                    # Handle duplicate DMC codes by appending a counter
                    if os.path.exists(output_path):
                        counter = 1
                        while True:
                            new_filename = f"{final_dmc}__{counter:03d}.docx"
                            output_path = os.path.join(output_dir, new_filename)
                            if not os.path.exists(output_path):
                                break
                            counter += 1
                        self.log(f"⚠ Duplicate DMC detected! Appending counter: __{counter:03d}")
                    
                    try:
                        shutil.copy2(filepath, output_path)
                        self.log(f"✓ Assigned: {final_dmc}")
                        self.log(f"  System: {dmc_parts['systemCode']}, SubSys: {dmc_parts['subSystemCode']}, Info: {dmc_parts['infoCode']}")
                        
                        # Show confidence and reasoning if available
                        if 'confidence' in dmc_parts and dmc_parts['confidence'] > 0:
                            confidence = dmc_parts['confidence']
                            confidence_icon = "🟢" if confidence >= 80 else "🟡" if confidence >= 60 else "🔴"
                            self.log(f"  {confidence_icon} Confidence: {confidence}%")
                        
                        if 'reasoning' in dmc_parts and dmc_parts['reasoning']:
                            self.log(f"  💡 Reasoning: {dmc_parts['reasoning']}")
                        
                        self.log(f"  Saved as: {new_filename}")
                        log_data["successful"].append({
                            "file": filename,
                            "assigned_dmc": final_dmc,
                            "output_file": new_filename,
                            "dmc_parts": dmc_parts
                        })
                    except Exception as e:
                        self.log(f"✗ Failed to save: {e}")
                        log_data["failed"].append({"file": filename, "issue": str(e)})
                else:
                    self.log(f"✗ Could not determine DMC")
                    log_data["failed"].append({"file": filename, "issue": "Failed to determine DMC"})
                
                self.progress['value'] = i + 1
            
            # Save log
            log_filename = os.path.join(LOGS_DIRECTORY, f"dmc_processing_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
            with open(log_filename, 'w', encoding='utf-8') as f:
                json.dump(log_data, f, indent=4)
            
            # Summary
            self.log(f"\n{'='*50}")
            self.log(f"PROCESSING COMPLETE")
            self.log(f"  Successful: {len(log_data['successful'])} files")
            self.log(f"  Failed: {len(log_data['failed'])} files")
            self.log(f"  Log saved: {log_filename}")
            self.log(f"{'='*50}")
            
            self.update_status(f"Complete! {len(log_data['successful'])} successful, {len(log_data['failed'])} failed")
            
            messagebox.showinfo("Complete", 
                f"Processing complete!\n\n"
                f"Successful: {len(log_data['successful'])}\n"
                f"Failed: {len(log_data['failed'])}\n\n"
                f"Output folder: {os.path.abspath(OUTPUT_DIRECTORY)}")
            
        except Exception as e:
            self.log(f"ERROR: {e}")
            self.update_status(f"Error: {e}")
            messagebox.showerror("Error", str(e))
        
        finally:
            self.processing = False
            self.start_btn.config(state=tk.NORMAL)


def main():
    root = tk.Tk()
    app = DMCAutomationGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
