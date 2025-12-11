import os
import re
import json
import shutil
import logging
from datetime import datetime
from bs4 import BeautifulSoup
import docx
import requests

# --- CONFIGURATION ---
OLLAMA_API_URL = "http://localhost:11434/api/generate"
# IMPORTANT: Use a reliable, instruction-following model like "llama3" or "mistral"
OLLAMA_MODEL = "llama3.1:8b" 
DOCS_DIRECTORY = "documents_to_process"
DATA_DIRECTORY = "Lake"
LOGS_DIRECTORY = "logs"
OUTPUT_DIRECTORY = "output"

# --- SNS JSON FILES TO LOAD ---
SNS_JSON_FILES = [
    "Maintained SNS - Generic.json",
    "maintained_sns_ordanance.json",
    "maintained_sns_support.json",
    "general_air_vehicles.json",
    "genral_surface_vehicles.json",
    "Maintained SNS - General communications.json"
]

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
log_filename = os.path.join(LOGS_DIRECTORY, f"dmc_processing_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# --- DATA PARSING AND PREPARATION ---

def parse_sns_json(file_path):
    """
    Parses SNS JSON files with various formats.
    Returns a dictionary of system codes with their titles and definitions.
    """
    sns_data = {}
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        filename = os.path.basename(file_path).lower()
        
        # Handle different JSON structures
        if isinstance(data, list):
            # Format: general_air_vehicles.json - array of system objects
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
            # Format: Maintained SNS - Generic.json or maintained_sns_ordanance.json
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
            
            # Also handle System_categories format (maintained_sns_support.json)
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
        
        logging.info(f"Loaded {len(sns_data)} systems from '{os.path.basename(file_path)}'")
        return sns_data
    except Exception as e:
        logging.error(f"Error parsing SNS JSON file {file_path}: {e}")
        return {}


def parse_info_codes_json(file_path):
    """Parses the info codes JSON file directly."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            info_codes = json.load(f)
        logging.info(f"Loaded {len(info_codes)} info codes from JSON")
        return info_codes
    except Exception as e:
        logging.error(f"Error parsing info codes JSON file {file_path}: {e}")
        return {}


def parse_sns_xml(file_path):
    """
    Parses an S1000D SNS XML file to extract system codes and titles.
    Handles files with or without a root <sns> tag.
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            soup = BeautifulSoup(f, 'xml')
        sns_data = {}
        
        systems_found = soup.find_all('snsSystem')
        
        if not systems_found:
            logging.warning(f"No <snsSystem> tags were found in '{os.path.basename(file_path)}'. The file might be empty or malformed.")
            return {}

        for system in systems_found:
            sys_code_tag, sys_title_tag = system.find('snsCode'), system.find('snsTitle')
            if not sys_code_tag or not sys_title_tag or not sys_code_tag.text.strip(): continue
            sys_code, sys_title = sys_code_tag.text.strip(), sys_title_tag.text.strip()
            sns_data[sys_code] = {'title': sys_title, 'subsystems': {}}
            for subsys in system.find_all('snsSubSystem'):
                sub_code_tag, sub_title_tag = subsys.find('snsCode'), subsys.find('snsTitle')
                if not sub_code_tag or not sub_title_tag or not sub_code_tag.text.strip(): continue
                sub_code, sub_title = sub_code_tag.text.strip(), sub_title_tag.text.strip()
                sns_data[sys_code]['subsystems'][sub_code] = {'title': sub_title, 'subsubsystems': {}}
                for subsubsys in subsys.find_all('snsSubSubSystem'):
                    subsub_code_tag, subsub_title_tag = subsubsys.find('snsCode'), subsubsys.find('snsTitle')
                    if not subsub_code_tag or not subsub_title_tag or not subsub_code_tag.text.strip(): continue
                    subsub_code, subsub_title = subsub_code_tag.text.strip(), subsub_title_tag.text.strip()
                    sns_data[sys_code]['subsystems'][sub_code]['subsubsystems'][subsub_code] = {'title': subsub_title}
        return sns_data
    except Exception as e:
        logging.error(f"Error parsing SNS XML file {file_path}: {e}")
        return {}

def parse_info_codes(file_path):
    """Parses the info codes text file."""
    info_codes = {}
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                if not line.strip(): continue
                match = re.match(r'^([0-9A-Z]{3})\s+([a-z]+)\s+(.*)', line.strip())
                if match:
                    code, type, desc = match.groups()
                    info_codes[code] = {'type': type, 'description': desc}
        return info_codes
    except Exception as e:
        logging.error(f"Error parsing info codes file {file_path}: {e}")
        return {}

def prepare_context_for_llm(sns_data, info_codes):
    """Creates a balanced string representation of the data for the LLM prompt, including subsystems."""
    # Include ALL system codes with titles AND their subsystems
    sns_lines = []
    for code, data in sorted(sns_data.items()):
        title = data.get('title', '')
        if title:
            sns_lines.append(f"{code}: {title}")
            # Add subsystems
            subsystems = data.get('subsystems', {})
            for sub_code, sub_data in sorted(subsystems.items()):
                sub_title = sub_data.get('title', '')
                if sub_title and sub_code not in ['00', '0']:  # Skip general subsystems
                    sns_lines.append(f"  {code}-{sub_code}: {sub_title}")
    
    sns_context = "VALID SYSTEM CODES AND SUBSYSTEMS:\n" + "\n".join(sns_lines)
    
    # Include key info codes - grouped by common types
    info_by_type = {}
    for code, data in sorted(info_codes.items()):
        code_type = data.get('type', 'other')
        desc = data.get('description', '')
        if code_type not in info_by_type:
            info_by_type[code_type] = []
        if desc and len(info_by_type[code_type]) < 15:  # Limit per type
            info_by_type[code_type].append(f"{code}: {desc}")
    
    info_lines = []
    for code_type in ['proced', 'descript', 'fault', 'process', 'sched']:
        if code_type in info_by_type:
            info_lines.append(f"\n[{code_type.upper()}]")
            info_lines.extend(info_by_type[code_type])
    
    info_context = "VALID INFO CODES (use one of these exactly):" + "\n".join(info_lines)
    
    return sns_context, info_context

# --- DOCUMENT PROCESSING ---

def extract_text_from_docx(file_path):
    """Extracts text from a .docx file, separating headings from body paragraphs."""
    try:
        doc = docx.Document(file_path)
        headings, body = [], []
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text: continue
            if para.style and para.style.name.lower().startswith('heading'):
                headings.append(text)
            else:
                body.append(text)
        return '\n'.join(headings), '\n'.join(body)
    except Exception as e:
        logging.error(f"Could not read docx file {file_path}: {e}")
        return None, None


# --- CORE LOGIC: LLM AND FALLBACK ---

def generate_dmc_with_llm(headings_text, body_text, sns_context, info_context, available_sns_codes, available_info_codes):
    """Uses Ollama with an optimized compact prompt to determine the DMC."""
    
    # Truncate body text to reduce prompt size
    body_preview = body_text[:1500] if body_text else "No content."
    headings_preview = headings_text[:400] if headings_text else "No headings."
    
    prompt = f"""Analyze this document and select the best S1000D DMC codes.

DOCUMENT TITLE/HEADINGS:
{headings_preview}

DOCUMENT EXCERPT:
{body_preview[:800]}

{sns_context}

{info_context}

INSTRUCTIONS:
- systemCode: Pick the 2-character system code (e.g., 20, 21, 24, 34)
- subSystemCode: Pick the subsystem digit (e.g., if 24-10 matches, use subSystemCode="1")
- subSubSystemCode: Usually "0" unless more specific
- infoCode: Pick a 3-character code (e.g., 000, 040, 520, 720)

Return ONLY this JSON:
{{"systemCode": "XX", "subSystemCode": "X", "subSubSystemCode": "0", "infoCode": "XXX", "disassyCode": "00", "disassyCodeVariant": "A"}}"""

    try:
        logging.info("Querying LLM...")
        payload = {
            "model": OLLAMA_MODEL, 
            "prompt": prompt, 
            "stream": False, 
            "format": "json",
            "options": {
                "temperature": 0.1,
                "num_predict": 200  # Increased to avoid truncation
            }
        }
        response = requests.post(OLLAMA_API_URL, json=payload, timeout=180)
        response.raise_for_status()
        
        raw_llm_response_text = response.json().get('response', '')
        logging.info(f"LLM response: {raw_llm_response_text[:300]}")

        if not raw_llm_response_text.strip():
            logging.error("LLM returned an empty response.")
            return None

        # Try to parse JSON, with cleanup for common issues
        json_text = raw_llm_response_text.strip()
        
        # Try to fix incomplete JSON
        if not json_text.endswith('}'):
            # Find last complete brace
            last_brace = json_text.rfind('}')
            if last_brace > 0:
                json_text = json_text[:last_brace + 1]
                logging.warning(f"Fixed truncated JSON response")
        
        # Try to extract JSON from response if there's extra text
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
            'disassyCodeVariant': str(dmc_parts.get('disassyCodeVariant', 'A'))
        }
        
        # VALIDATION: Check if codes exist in loaded data
        if available_sns_codes and final_parts['systemCode'] not in available_sns_codes:
            logging.warning(f"LLM systemCode '{final_parts['systemCode']}' not in loaded data. Using default.")
            final_parts['systemCode'] = DEFAULT_SYSTEM_CODE
        
        if available_info_codes and final_parts['infoCode'] not in available_info_codes:
            logging.warning(f"LLM infoCode '{final_parts['infoCode']}' not in loaded data. Using default.")
            final_parts['infoCode'] = DEFAULT_INFO_CODE

        logging.info(f"LLM returned codes: {final_parts}")
        return final_parts
        
    except json.JSONDecodeError as e:
        logging.error(f"Failed to parse LLM JSON response: {e}")
        return None
    except requests.exceptions.Timeout:
        logging.error("LLM request timed out after 120 seconds.")
        return None
    except Exception as e:
        logging.error(f"LLM processing failed: {e}")
        return None

def generate_dmc_with_fallback(headings_text, body_text, sns_data, info_codes):
    """A context-aware fallback that scores based on document category."""
    logging.warning("Executing context-aware fallback...")
    full_text_lower = (headings_text + " " + body_text).lower()

    CATEGORY_KEYWORDS = {
        'proced': ['procedure', 'step', 'task', 'perform', 'install', 'remove', 'assemble', 'disassemble', 'prepare', 'unpack', 'setup', 'execute', 'how to'],
        'descript': ['description', 'overview', 'introduction', 'component', 'feature', 'specification', 'what is', 'theory'],
        'fault': ['fault', 'troubleshooting', 'symptom', 'remedy', 'isolation', 'failure', 'error code', 'diagnose'],
    }

    doc_category, max_category_score = None, 0
    for category, keywords in CATEGORY_KEYWORDS.items():
        score = sum(1 for keyword in keywords if keyword in full_text_lower)
        if score > max_category_score:
            max_category_score, doc_category = score, category
    
    if doc_category:
        logging.info(f"Fallback: Detected document category as '{doc_category}' with score {max_category_score}.")
    else:
        logging.warning("Fallback: Could not determine a strong document category.")

    filtered_info_codes = {code: data for code, data in info_codes.items() if data['type'] == doc_category} if doc_category else info_codes
    if doc_category and not filtered_info_codes:
        logging.warning(f"No info codes of the detected category '{doc_category}' were found. Considering all info codes.")
        filtered_info_codes = info_codes

    info_scores = []
    if filtered_info_codes:
        for code, data in filtered_info_codes.items():
            score = sum(1 for keyword in data['description'].lower().split() if keyword in full_text_lower)
            if score > 0:
                info_scores.append((score, code, data.get('description')))

    sns_scores = []
    if sns_data:
        for code, data in sns_data.items():
            score = sum(10 for keyword in data['title'].lower().split() if keyword in headings_text.lower())
            score += sum(1 for keyword in data['title'].lower().split() if keyword in body_text.lower())
            if score > 0:
                sns_scores.append((score, code, data.get('title')))

    logging.info("--- Fallback Scoring Report ---")
    if sns_scores:
        sns_scores.sort(key=lambda x: x[0], reverse=True)
        logging.info("Top SNS Candidates:")
        for score, code, title in sns_scores[:3]: logging.info(f"  - Score: {score}, Code: {code}, Title: {title}")
    
    if info_scores:
        info_scores.sort(key=lambda x: x[0], reverse=True)
        logging.info("Top Info Code Candidates (from detected category):")
        for score, code, desc in info_scores[:3]: logging.info(f"  - Score: {score}, Code: {code}, Description: {desc}")
    
    best_sns = sns_scores[0][1] if sns_scores else DEFAULT_SYSTEM_CODE
    best_info = info_scores[0][1] if info_scores else DEFAULT_INFO_CODE
    
    dmc_parts = {"systemCode": best_sns, "infoCode": best_info, "subSystemCode": "0", "subSubSystemCode": "0", "disassyCode": "00", "disassyCodeVariant": "A"}
    logging.info(f"Fallback mechanism selected: {dmc_parts}")
    return dmc_parts

# --- Main execution block and other functions remain the same ---
def format_dmc(parts):
    return (f'DMC-{USER_MODEL_IDENT_CODE}-{USER_SYSTEM_DIFF_CODE}-{parts.get("systemCode", DEFAULT_SYSTEM_CODE)}-'
            f'{str(parts.get("subSystemCode", "0"))}{str(parts.get("subSubSystemCode", "0"))}-{USER_ASSY_CODE}-'
            f'{parts.get("disassyCode", "00")}{parts.get("disassyCodeVariant", "A")}-{parts.get("infoCode", DEFAULT_INFO_CODE)}A-{USER_ITEM_LOCATION_CODE}')


def select_sns_files():
    """Interactive menu to select which SNS JSON files to use."""
    print("\n" + "="*60)
    print("SELECT SNS DATA FILES TO USE")
    print("="*60)
    
    # Find available SNS files in the Lake directory
    available_files = []
    for f in os.listdir(DATA_DIRECTORY):
        if f.endswith('.json') and f != 'info_codes.json':
            available_files.append(f)
    
    if not available_files:
        print("No SNS JSON files found in Lake directory!")
        return []
    
    print("\nAvailable SNS files:")
    print("-" * 40)
    for i, filename in enumerate(available_files, 1):
        print(f"  [{i}] {filename}")
    
    print(f"\n  [A] Select ALL files")
    print(f"  [Q] Quit")
    print("-" * 40)
    
    while True:
        choice = input("\nEnter your choice (number, 'A' for all, or 'Q' to quit): ").strip().upper()
        
        if choice == 'Q':
            print("Exiting...")
            return None
        
        if choice == 'A':
            print(f"\n✓ Selected ALL {len(available_files)} files")
            return available_files
        
        # Handle multiple selections (e.g., "1,2,3" or "1 2 3")
        try:
            if ',' in choice:
                indices = [int(x.strip()) for x in choice.split(',')]
            elif ' ' in choice:
                indices = [int(x.strip()) for x in choice.split()]
            else:
                indices = [int(choice)]
            
            selected = []
            for idx in indices:
                if 1 <= idx <= len(available_files):
                    selected.append(available_files[idx - 1])
                else:
                    print(f"Invalid selection: {idx}")
            
            if selected:
                print(f"\n✓ Selected {len(selected)} file(s):")
                for f in selected:
                    print(f"    - {f}")
                return selected
                
        except ValueError:
            print("Invalid input. Enter number(s), 'A' for all, or 'Q' to quit.")


def main():
    print("\n" + "="*60)
    print("       DMC AUTOMATION PROCESS")
    print("="*60)
    
    # Select SNS files interactively
    selected_sns_files = select_sns_files()
    if selected_sns_files is None:
        return
    
    logging.info("--- Starting DMC Automation Process ---")
    try:
        sns_data, info_codes = {}, {}
        
        # Load info codes from JSON (preferred) or TXT - ALWAYS LOADED
        info_codes_json_path = os.path.join(DATA_DIRECTORY, "info_codes.json")
        info_codes_txt_path = os.path.join(DATA_DIRECTORY, "info_codes.txt")
        
        print("\n[INFO CODES] Loading automatically...")
        if os.path.exists(info_codes_json_path):
            logging.info(f"Loading info codes from JSON: {info_codes_json_path}")
            info_codes = parse_info_codes_json(info_codes_json_path)
            print(f"  ✓ Loaded {len(info_codes)} info codes from: info_codes.json")
        elif os.path.exists(info_codes_txt_path):
            logging.info(f"Loading info codes from TXT: {info_codes_txt_path}")
            info_codes = parse_info_codes(info_codes_txt_path)
            print(f"  ✓ Loaded {len(info_codes)} info codes from: info_codes.txt")
        else:
            logging.warning("No info_codes.json or info_codes.txt found.")
            print("  ✗ No info codes file found!")
        
        # Load SNS data from SELECTED JSON files
        print("\n[SNS FILES] Loading selected files...")
        loaded_sns_files = []
        for sns_file in selected_sns_files:
            file_path = os.path.join(DATA_DIRECTORY, sns_file)
            if os.path.exists(file_path):
                file_sns_data = parse_sns_json(file_path)
                if file_sns_data:
                    sns_data.update(file_sns_data)
                    loaded_sns_files.append(sns_file)
                    print(f"  ✓ Loaded {len(file_sns_data)} systems from: {sns_file}")
                    logging.info(f"  ✓ Loaded {len(file_sns_data)} systems from: {sns_file}")
            else:
                print(f"  ✗ SNS file not found: {sns_file}")
                logging.debug(f"  ✗ SNS file not found (skipping): {sns_file}")

        # Summary of loaded files
        print(f"\n{'='*50}")
        print(f"DATA SOURCES LOADED:")
        print(f"  Info Codes: {len(info_codes)} codes")
        print(f"  SNS Files: {len(loaded_sns_files)} file(s)")
        print(f"  Total SNS systems: {len(sns_data)}")
        print(f"{'='*50}")
        
        logging.info(f"\n{'='*50}")
        logging.info(f"DATA SOURCES LOADED:")
        logging.info(f"  Info Codes: {info_codes_json_path if os.path.exists(info_codes_json_path) else info_codes_txt_path}")
        logging.info(f"  SNS Files ({len(loaded_sns_files)}):")
        for f in loaded_sns_files:
            logging.info(f"    - {f}")
        logging.info(f"  Total SNS systems: {len(sns_data)}")
        logging.info(f"  Total Info codes: {len(info_codes)}")
        logging.info(f"{'='*50}\n")

        if not sns_data: 
            logging.warning(f"CRITICAL: No SNS systems were loaded.")
        if not info_codes: 
            logging.warning(f"CRITICAL: No Info Codes were loaded.")
            
    except Exception as e:
        logging.error(f"Critical error during data loading: {e}. Exiting.")
        return

    sns_context_str, info_context_str = prepare_context_for_llm(sns_data, info_codes)
    available_sns_codes, available_info_codes = set(sns_data.keys()), set(info_codes.keys())
    
    logging.info(f"Context prepared - SNS context size: {len(sns_context_str)} chars, Info context size: {len(info_context_str)} chars")
    
    files_to_process = [f for f in os.listdir(DOCS_DIRECTORY) if f.endswith(".docx")]
    if not files_to_process:
        logging.warning(f"No .docx files found in '{DOCS_DIRECTORY}'.")
        return
    
    # Ensure output directory exists
    if not os.path.exists(OUTPUT_DIRECTORY):
        os.makedirs(OUTPUT_DIRECTORY)
        logging.info(f"Created output directory: {OUTPUT_DIRECTORY}")
    
    log_data = {
        "data_sources": {
            "info_codes_file": info_codes_json_path if os.path.exists(info_codes_json_path) else info_codes_txt_path,
            "sns_files_loaded": loaded_sns_files,
            "total_sns_systems": len(sns_data),
            "total_info_codes": len(info_codes)
        },
        "successful": [], 
        "failed": []
    }
    for filename in files_to_process:
        logging.info(f"--- Processing file: {filename} ---")
        filepath = os.path.join(DOCS_DIRECTORY, filename)
        headings_text, body_text = extract_text_from_docx(filepath)
        if not headings_text and not body_text:
            log_data["failed"].append({"file": filename, "issue": "Could not read or extract content."})
            continue
        
        dmc_parts = generate_dmc_with_llm(headings_text, body_text, sns_context_str, info_context_str, available_sns_codes, available_info_codes)
        
        if not dmc_parts:
            logging.warning("LLM failed, attempting context-aware fallback.")
            dmc_parts = generate_dmc_with_fallback(headings_text, body_text, sns_data, info_codes)
            
        if dmc_parts:
            final_dmc = format_dmc(dmc_parts)
            
            # Save file to output directory with new DMC filename
            new_filename = f"{final_dmc}.docx"
            output_path = os.path.join(OUTPUT_DIRECTORY, new_filename)
            try:
                shutil.copy2(filepath, output_path)
                logging.info(f"Saved: {new_filename} -> {OUTPUT_DIRECTORY}/")
            except Exception as e:
                logging.error(f"Failed to save file {new_filename}: {e}")
            
            log_data["successful"].append({
                "file": filename, 
                "assigned_dmc": final_dmc, 
                "output_file": new_filename,
                "dmc_parts": dmc_parts
            })
            logging.info(f"Successfully assigned DMC: {final_dmc}")
        else:
            log_data["failed"].append({"file": filename, "issue": "Failed to determine DMC using all methods."})
            logging.error(f"Could not assign DMC for file: {filename}")
            
    with open(log_filename, 'w', encoding='utf-8') as f:
        json.dump(log_data, f, indent=4)
    
    # Print summary
    logging.info(f"\n{'='*50}")
    logging.info(f"PROCESSING COMPLETE")
    logging.info(f"  Successful: {len(log_data['successful'])} files")
    logging.info(f"  Failed: {len(log_data['failed'])} files")
    logging.info(f"  Output folder: {os.path.abspath(OUTPUT_DIRECTORY)}")
    logging.info(f"  Log file: {log_filename}")
    logging.info(f"{'='*50}")

if __name__ == "__main__":
    main()