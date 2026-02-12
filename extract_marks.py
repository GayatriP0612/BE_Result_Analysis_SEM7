print("DEBUG: Importing os...", flush=True)
import os
print("DEBUG: Importing fitz...", flush=True)
import fitz
print("DEBUG: Importing pytesseract...", flush=True)
import pytesseract
print("DEBUG: Importing PIL...", flush=True)
from PIL import Image, ImageFilter, ImageEnhance
print("DEBUG: Importing io/re/json...", flush=True)
import io
import re
import json
print("DEBUG: Importing pandas...", flush=True)
import pandas as pd
print("DEBUG: Importing genai...", flush=True)
import google.generativeai as genai
print("DEBUG: Imports done.", flush=True)
import time
from itertools import cycle
import warnings

# Suppress all warnings
warnings.filterwarnings("ignore")

# Configuration
# NOTE: Update this path to point to your Tesseract installation
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"



# Multiple Gemini API Keys for rotation
GEMINI_API_KEYS = [
    "AIzaSyCpLQg4g5InyjnIwK3hMAMBQXHlLbeNbx4",
    "AIzaSyBlQW7FCQo1UH0NOVfC_MhN9kUI3fbOUCg",
    "AIzaSyCx7q_UXlI1C8yowMEBMGT_aWerlNuX5yE",
    "AIzaSyB2Stna2jdh0GkNvjW14rZ3P0q6lYWYN8w",
    "AIzaSyCzuZH1EMqbaBo-vgbbPNr2l0n_bKSjnME",
    "AIzaSyCxoIt04uoi1Tfb4IhHE45DCnGZYdeiq_g",
    "AIzaSyBDy81O0ADK1I1mmkdGY3qtf_gPSzi-k6Q",
    "AIzaSyCGZrt_kyaMJjEpjlv6HF3R9K662f9nXeA",
    "AIzaSyCzc_B_FSXc48med90IHzwoBXxEpnlpItg",
    "AIzaSyCviO4U0sYq8CwRbjLrBZXYxcMuyqlgzsw",
    "AIzaSyBJsuuNVPBzMzni4OJsMbjahq1SK_rNIPU",
    "AIzaSyD2EUN2Yzvg8BGlTgSGGD0DMJIljJYQnx4",
    "AIzaSyB6_jTDNmAJ9zUmNxE01bb3g-pUHAlD_6U",
    "AIzaSyD76h3-ui94MNKrajYOUh7KvQDQwPvPuX4",
    "AIzaSyC1SGZ6z3pfPfAA1LANB29m9jFw4uotass",
    "AIzaSyDIHSaswmcJ9a97PIEUzoHOS2C90ZoYxI4"
]
key_pool = cycle(GEMINI_API_KEYS)
current_key = next(key_pool)
genai.configure(api_key=current_key)

# File Paths
INPUT_PDF = os.path.join("input", "BECOMP(1).pdf")
OUTPUT_XLSX = "BE_2025_Final_Results.xlsx"
os.makedirs("ocr_pages", exist_ok=True)

# Updated subjects dictionary
# Updated subjects dictionary
# Updated subjects dictionary - Using Tail Anchor Strategy (Tot% Crd Grd GP CP)
# Pattern: SubjectCode ... Marks ... Tot% Crd Grd GP CP
# Updated subjects dictionary - Robust "Skip Intermediates" Strategy
# Pattern: SubjectCode ... Marks ... [Skip Tot%, Crd] ... Grd GP CP
SUBJECTS = {
    "410241": {
        "name": "DESIGN & ANALYSIS OF ALGO",
        "type": "theory",
        "pattern": r"410241.*?(?:\d{2,3}/\d{2,3}\s+){2}(\d{2,3})/100.*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
    "410242": {
        "name": "MACHINE LEARNING", 
        "type": "theory",
        "pattern": r"410242.*?(?:\d{2,3}/\d{2,3}\s+){2}(\d{2,3})/100.*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
    "410243": {
        "name": "BLOCKCHAIN TECHNOLOGY",
        "type": "theory",
        "pattern": r"410243.*?(?:\d{2,3}/\d{2,3}\s+){2}(\d{2,3})/100.*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
    "410244C": {
        "name": "CYBER SEC. & DIGITAL FORENSICS",
        "type": "theory",
        "pattern": r"410244C.*?(?:\d{2,3}/\d{2,3}\s+){2}(\d{2,3})/100.*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
    "410245A": {
        "name": "INFORMATION RETRIEVAL",
        "type": "theory",
        "pattern": r"410245A.*?(?:\d{2,3}/\d{2,3}\s+){2}(\d{2,3})/100.*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
    "410246": {
        "name": "LABORATORY PRACTICE - III",
        "type": "practical",
        "pattern": r"410246.*?(\d{2,3})/(\d{2,3})\s+(\d{2,3})/(\d{2,3}).*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
    "410247": {
        "name": "LABORATORY PRACTICE - IV",
        "type": "practical",
        "pattern": r"410247.*?(\d{2,3})/(?:\d{2,3})(?:\s+(\d{2,3})/(?:\d{2,3}))?.*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
    "410248": {
        "name": "PROJECT STAGE - I",
        "type": "termwork",
        "pattern": r"410248.*?(\d{2,3})/(\d{2,3}).*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
     "410249A": {
        "name": "MOOC - LEARN NEW SKILLS",
        "type": "grade",
        "pattern": r"410249A.*?\s+(PP|AC|FA|EX|P)"
    },
    # Honors Subjects
    "410501": {
        "name": "HON-MACH. LEARN.& DATA SCI.",
        "type": "theory", 
        "pattern": r"410501\s+HON-MACH\.\s+LEARN\.\&\s+DATA\s+SCI\..*?(?:\d{2,3}/\d{2,3}\s+){2}(\d{2,3})/100.*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
    "410501_PR": {
        "name": "HON-MACH. LEARN.& DATA SCI. (PR)",
        "type": "theory", # Treat as theory to extract single Total mark
        # Use Greedy .* to find the LAST mark component (Total) before Grade
        "pattern": r"410501\s+HON-MACH\.\s+LEARN\.\&\s+DATA\s+SCI\.\(PR\).*(\d{2,3})/\d+.*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
    "410301": {
        "name": "HON-MACHINE LEARNING",
        "type": "theory",
        "pattern": r"410301\s+HON-MACHINE\s+LEARNING.*?(?:\d{2,3}/\d{2,3}\s+){2}(\d{2,3})/100.*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
    "410301_PR": {
        "name": "HON-MACHINE LEARNING (PR)",
        "type": "theory",
        "pattern": r"410301\s+HON-MACHINE\s+LEARNING\s*\(PR\).*(\d{2,3})/\d+.*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
    "410401": {
        "name": "HON-IOT & EMBEDDED SECURITY",
        "type": "theory",
        "pattern": r"410401\s+HON-IOT\s+\&\s+EMBEDDED\s+SECURITY.*?(?:\d{2,3}/\d{2,3}\s+){2}(\d{2,3})/100.*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    },
    "410402": {
        "name": "HON-RISK ASSMNT LABORATORY (PR)",
        "type": "theory",
        "pattern": r"410402\s+HON-RISK\s+ASSMNT\s+LABORATORY\s*\(PR\).*(\d{2,3})/\d+.*?([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)"
    }
}

# Required columns as per user request
REQUIRED_COLUMNS = [
    "Seat No",
    "Name",
    "Mother",
    "PRN",
    "DESIGN & ANALYSIS OF ALGO",
    "DESIGN & ANALYSIS OF ALGO (Grade)",
    "MACHINE LEARNING",
    "MACHINE LEARNING (Grade)",
    "BLOCKCHAIN TECHNOLOGY",
    "BLOCKCHAIN TECHNOLOGY (Grade)",
    "CYBER SEC. & DIGITAL FORENSICS",
    "CYBER SEC. & DIGITAL FORENSICS (Grade)",
    "INFORMATION RETRIEVAL",
    "INFORMATION RETRIEVAL (Grade)",
    "LABORATORY PRACTICE - III (TW)",
    "LABORATORY PRACTICE - III (PR)",
    "LABORATORY PRACTICE - III (Grade)",
    "LABORATORY PRACTICE - IV (TW)",
    "LABORATORY PRACTICE - IV (PR)",
    "LABORATORY PRACTICE - IV (Grade)",
    "PROJECT STAGE - I",
    "PROJECT STAGE - I (Grade)",
    "MOOC - LEARN NEW SKILLS",
    # Honors Columns
    "HON-MACH. LEARN.& DATA SCI.",
    "HON-MACH. LEARN.& DATA SCI. (Grade)",
    "HON-MACH. LEARN.& DATA SCI. (PR)",
    "HON-MACH. LEARN.& DATA SCI. (PR) (Grade)",
    "HON-MACHINE LEARNING",
    "HON-MACHINE LEARNING (Grade)",
    "HON-MACHINE LEARNING (PR)",
    "HON-MACHINE LEARNING (PR) (Grade)",
    "HON-IOT & EMBEDDED SECURITY",
    "HON-IOT & EMBEDDED SECURITY (Grade)",
    "HON-RISK ASSMNT LABORATORY (PR)",
    "HON-RISK ASSMNT LABORATORY (PR) (Grade)",
    "SGPA",
    "TOTAL CREDITS"
]

def rotate_gemini_key():
    """Rotate to the next API key in the pool"""
    global current_key
    current_key = next(key_pool)
    genai.configure(api_key=current_key)
    # print(f"🔑 Rotated to key: {current_key[:8]}...")

def enhance_image(img):
    """Simplified image preprocessing for speed and sufficient quality"""
    try:
        # Convert to grayscale
        img = img.convert("L")
        
        # Increase contrast slightly
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(1.5)
        
        # Binarization (optional, but helpful for clean text)
        # img = img.point(lambda x: 0 if x < 200 else 255)
        
        return img
    except Exception as e:
        print(f"Image enhancement failed: {str(e)}")
        return img

GP_MAP = {
    '10': 'O',
    '09': 'A+',
    '9': 'A+',
    '08': 'A',
    '8': 'A',
    '07': 'B+',
    '7': 'B+',
    '06': 'B',
    '6': 'B',
    '05': 'C',
    '5': 'C',
    '04': 'P',
    '4': 'P',
    '00': 'F',
    '0': 'F'
}

def clean_grade(grade, gp=None):
    """Normalize grade strings to standard formats, using GP if available"""
    if not grade or grade == "NA":
        return "NA"
    
    # If GP Provided, use it as Authority
    if gp and gp in GP_MAP:
        return GP_MAP[gp]

    # Common OCR fixups (Fallback)
    grade = grade.upper().strip()
    
    # Fix 'Bt' -> 'B+'
    grade = grade.replace("BT", "B+")
    
    # Fix 'Ct' -> 'C+'
    grade = grade.replace("CT", "C+")

    # Fix 'A1' -> 'A+'
    grade = grade.replace("A1", "A+")
    
    # Removed unsafe "AT" -> "A+" because it conflicts with "A" being read as "At"
    # Logic: "At" is ambiguous. Without GP, we can't be 100% sure. 
    # But usually, if GP is missing, we might assume A+ if it looks like At?
    # For now, safer to leave it or rely on valid chars.
    
    # Fix 0 -> O (Outstanding)
    if grade == '0':
        grade = 'O'
        
    return grade

def extract_data_regex(text):
    """Fallback regex extraction for when Gemini fails or for validation"""
    # print(f"[DEBUG] Analyzing text for regex (first 300 chars): {text[:300]!r}")
    data = {}
    
    # Basic Info
    seat_match = re.search(r"SEAT\s*N[O0][\.:\s\-]*([A-Z0-9]\d{5,12})", text, re.IGNORECASE)
    data["Seat No"] = seat_match.group(1) if seat_match else "NA"
    
    name_match = re.search(r"NAME\s*[:\.]?\s*([A-Z\s\.]+?)(?:MOTHER|PRN)", text, re.IGNORECASE)
    data["Name"] = name_match.group(1).strip() if name_match else "NA"
    
    mother_match = re.search(r"MOTHER\s*[:\.]?\s*([A-Z\s\.]+?)(?:PRN|CLG)", text, re.IGNORECASE)
    data["Mother"] = mother_match.group(1).strip() if mother_match else "NA"
    
    # Debug PRN area
    prn_index = text.find("PRN")
    if prn_index != -1:
        print(f"[DEBUG] Text around PRN: {text[prn_index:prn_index+50]!r}", flush=True)

    # PRN Extraction Strategy
    # 1. Try context-based match first
    # Match PRN followed by digits/letters/spaces/OCR noise, until CLG/Pict/[ or end
    prn_match = re.search(r"PRN\s*[:\.]?\s*([0-9A-Z\s/\.-]{8,20})\s*(?:CLG|Pict|\[|$)", text, re.IGNORECASE)
    
    found_prn = None
    if prn_match:
        # Clean up: remove spaces and OCR noise
        candidate = re.sub(r"[\s/\.-]", "", prn_match.group(1))
        # Validate candidate (should contain at least 5 digits to be a PRN, typically >7 chars)
        if len(re.findall(r"\d", candidate)) >= 5 and len(candidate) > 7:
            found_prn = candidate
            
    # 2. Fallback: Direct pattern match with spaces allowed
    if not found_prn:
        # PUNE Univ PRNs: starts with 7, approx 8-9 digits, ends with char
        # Regex: 7 followed by 7-10 digits/spaces, ending with a letter (optional)
        direct_match = re.search(r"\b(7[\d\s]{7,12}[A-Z]?)\b", text)
        if direct_match:
            found_prn = direct_match.group(1).replace(" ", "")

    data["PRN"] = found_prn if found_prn else "NA"
    
    # Relaxed Regex for SGPA and Credits
    # Matches SGPA, SGPA1, SCPA (OCR error), etc. with flexible separator
    sgpa_match = re.search(r"(?:SGPA|SCPA|SGPA1|S\.G\.P\.A)\s*\d*\s*[:=,.-]?\s*(\d+\.\d+)", text, re.IGNORECASE)
    data["SGPA"] = sgpa_match.group(1).strip() if sgpa_match else "NA"
    
    # Matches TOTAL CREDITS with flexible separator
    credits_match = re.search(r"TOTAL\s*CREDITS.*?(?:[:=,-]|\s)\s*(\d+)", text, re.IGNORECASE)
    data["TOTAL CREDITS"] = credits_match.group(1).strip() if credits_match else "NA"
    
    # Subject Marks
    for code, info in SUBJECTS.items():
        pattern = info['pattern']
        # Use re.IGNORECASE only; removing re.DOTALL prevents bleeding into next subject lines
        match = re.search(pattern, text, re.IGNORECASE)
        name = info['name']
        
        if match:
             # Identify Grade and GP
             if info['type'] == 'grade':
                 # MOOC case
                 val = match.group(1)
                 data[name] = clean_grade(val, gp=None)
             else:
                 # Robust Pattern: Marks ... [Skip] ... <Grade> <GP> <CP>
                 # We simply take the LAST 3 groups for Grade, GP, CP.
                 # This works because all patterns end with: ([A-Z][\w\+]*|0)\s+(10|0?[0-9])\s+(\d+)
                 
                 groups = match.groups()
                 grade_val = groups[-3]
                 gp_val = groups[-2]
                 # cp_val = groups[-1] # Not used currently
                 
                 data[f"{name} (Grade)"] = clean_grade(grade_val, gp=gp_val)

             if info['type'] == 'theory':
                # Group 1 is Marks
                if match.lastindex and match.lastindex >= 1:
                     val = match.group(1)
                     data[name] = val if val else "NA"
             elif info['type'] == 'practical':
                # Check for TW
                tw_val = match.group(1) if match.lastindex >= 1 else None
                data[f"{name} (TW)"] = tw_val if tw_val else "NA"
                
                # Check for PR
                if "Lab III" in name or "LABORATORY PRACTICE - III" in name:
                     # PR is group 3
                     pr_val = match.group(3) if match.lastindex >= 3 else None
                     data[f"{name} (PR)"] = pr_val if pr_val else "NA"
                else:
                     # Lab IV: PR is group 2
                     pr_val = match.group(2) if match.lastindex >= 2 else None
                     data[f"{name} (PR)"] = pr_val if pr_val else "NA"

             elif info['type'] == 'termwork':
                  val = match.group(1) if match.lastindex >= 1 else None
                  data[name] = val if val else "NA"

        else:
            if info['type'] == 'grade':
                 data[name] = "NA"
            elif info['type'] == 'practical':
                data[f"{name} (TW)"] = "NA"
                data[f"{name} (PR)"] = "NA"
                data[f"{name} (Grade)"] = "NA"
            else:
                data[name] = "NA"
                data[f"{name} (Grade)"] = "NA"
                
    return data

def extract_text_from_image(img):
    """Robust text extraction with multiple OCR attempts"""
    try:
        # Try with different configurations
        configs = [
            r'--psm 6 --oem 1',
            r'--psm 3 --oem 1',
            r'--psm 11 --oem 1'
        ]
        
        best_text = ""
        best_score = 0
        
        for config in configs:
            try:
                text = pytesseract.image_to_string(img, config=config)
                score = sum(c.isalnum() for c in text)
                if score > best_score:
                    best_text = text
                    best_score = score
            except:
                continue
        
        return best_text if best_text else pytesseract.image_to_string(img)
    except Exception as e:
        print(f"OCR failed: {str(e)}")
        return ""

# ... (rest of functions like format_for_gemini, parse_with_gemini, process_student_from_data, split_student_sections are assumed to be below this)

        configs = [
            r'--psm 6 --oem 1',
            r'--psm 3 --oem 1',
            r'--psm 11 --oem 1'
        ]
        
        best_text = ""
        best_score = 0
        
        for config in configs:
            try:
                text = pytesseract.image_to_string(img, config=config)
                score = sum(c.isalnum() for c in text)
                if score > best_score:
                    best_text = text
                    best_score = score
            except:
                continue
        
        return best_text if best_text else pytesseract.image_to_string(img)
    except Exception as e:
        print(f"OCR failed: {str(e)}")
        return ""

def format_for_gemini(student_text):
    """Prompt for Gemini to extract specific fields"""
    return f"""
You are a data extraction assistant. Extract the following fields from the student marksheet text below.
Return the output STRICTLY as a JSON object.

Fields to Extract:
1. "Seat No": The seat number (e.g. B400050314)
2. "Name": The student's full name (e.g. KOSHATWAR VAISHANAVI RANJIT)
3. "Mother": The mother's name (e.g. RENUKA RANJIT KOSHATWAR or just RENUKA). Look for "MOTHER :"
4. "PRN": The PRN number (e.g. 72278407E)
5. "subjects": A nested object containing marks for these exact subjects.
   - For all subjects, extract "Grade" (Grd) AND "GP" (Grade Point, a number from 0-10).
   
   Expected Keys:
    - "DESIGN & ANALYSIS OF ALGO": Total marks
    - "DESIGN & ANALYSIS OF ALGO (Grade)": Grade (e.g. A+, O)
    - "DESIGN & ANALYSIS OF ALGO (GP)": Grade Point (e.g. 09, 10)
    
    - "MACHINE LEARNING": Total marks
    - "MACHINE LEARNING (Grade)": Grade
    - "MACHINE LEARNING (GP)": GP
    
    - "BLOCKCHAIN TECHNOLOGY": Total marks
    - "BLOCKCHAIN TECHNOLOGY (Grade)": Grade
    - "BLOCKCHAIN TECHNOLOGY (GP)": GP
    
    - "CYBER SEC. & DIGITAL FORENSICS": Total marks
    - "CYBER SEC. & DIGITAL FORENSICS (Grade)": Grade
    - "CYBER SEC. & DIGITAL FORENSICS (GP)": GP
    
    - "INFORMATION RETRIEVAL": Total marks
    - "INFORMATION RETRIEVAL (Grade)": Grade
    - "INFORMATION RETRIEVAL (GP)": GP
    
    - "LABORATORY PRACTICE - III (TW)": Term Work marks
    - "LABORATORY PRACTICE - III (PR)": Practical marks
    - "LABORATORY PRACTICE - III (Grade)": Grade
    - "LABORATORY PRACTICE - III (GP)": GP
    
    - "LABORATORY PRACTICE - IV (TW)": Term Work marks
    - "LABORATORY PRACTICE - IV (PR)": Practical marks
    - "LABORATORY PRACTICE - IV (Grade)": Grade
    - "LABORATORY PRACTICE - IV (GP)": GP
    
    - "PROJECT STAGE - I": Term Work marks
    - "PROJECT STAGE - I (Grade)": Grade
    - "PROJECT STAGE - I (GP)": GP
    
    - "MOOC - LEARN NEW SKILLS": Grade or Status

6. "SGPA": The SGPA value (e.g. 7.85). Look for "SGPA1" or "SGPA".
7. "TOTAL CREDITS": The total credits earned (e.g. 20).

JSON Structure:
{{
  "seat_no": "...",
  "name": "...",
  "mother": "...",
  "prn": "...",
  "subjects": {{
    "DESIGN & ANALYSIS OF ALGO": "...",
    "DESIGN & ANALYSIS OF ALGO (Grade)": "...",
    "DESIGN & ANALYSIS OF ALGO (GP)": "...",
    ... other subjects ...
    "MOOC - LEARN NEW SKILLS": "..."
  }},
  "sgpa": "...",
  "total_credits": "..."
}}

If a value is missing or cannot be found, set it to "NA".

Input Text:
\"\"\"
{student_text}
\"\"\"
"""

def parse_with_gemini(prompt, page_num, student_idx):
    """Robust parsing with Gemini including better error handling"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            model = genai.GenerativeModel("models/gemini-2.5-flash")
            response = model.generate_content(prompt)
            text = response.text.strip()
            
            # Clean JSON formatting
            text = text.replace("```json", "").replace("```", "").strip()
            
            # Basic validation
            if not text.startswith("{") or not text.endswith("}"):
                raise ValueError("Invalid JSON format")
                
            data = json.loads(text)
            return data
            
        except Exception as e:
            print(f"⚠️ Attempt {attempt+1} failed: {str(e)[:100]}")
            if attempt < max_retries - 1:
                rotate_gemini_key()
                time.sleep(2)
            else:
                return None

def process_student_from_data(data):
    """Convert Gemini JSON to flat dictionary matching REQUIRED_COLUMNS"""
    if not data:
        return None
        
    student = {
        "Seat No": data.get("seat_no", "NA"),
        "Name": data.get("name", "NA"),
        "Mother": data.get("mother", "NA"),
        "PRN": data.get("prn", "NA"),
        "SGPA": data.get("sgpa", "NA"),
        "TOTAL CREDITS": data.get("total_credits", "NA")
    }
    
    subjects = data.get("subjects", {})
    
    # Map subject marks to student dict
    for col in REQUIRED_COLUMNS:
        if col not in student:
            # If requesting Grade, try to find GP for it to validate
            if "(Grade)" in col:
                subj_base = col.replace(" (Grade)", "")
                raw_grade = subjects.get(col, "NA")
                gp_key = f"{subj_base} (GP)"
                raw_gp = subjects.get(gp_key, None) # Get GP if exists from Gemini
                
                # Apply validation using GP
                student[col] = clean_grade(raw_grade, gp=raw_gp)
            else:
                student[col] = subjects.get(col, "NA")
            
    return student

def split_student_sections(text):
    """Split text by finding all 'SEAT NO' occurrences and slicing"""
    # Find all start indices of "SEAT NO"
    matches = list(re.finditer(r"SEAT\s*NO", text, re.IGNORECASE))
    
    if not matches:
        return [text] if len(text) > 100 else []
        
    sections = []
    for i in range(len(matches)):
        start = matches[i].start()
        # End is start of next match or end of string
        end = matches[i+1].start() if i + 1 < len(matches) else len(text)
        sections.append(text[start:end])
        
    print(f"[DEBUG] Found {len(sections)} students by splitting")
    return sections

def merge_data(regex_data, gemini_data):
    """Merge data, preferring regex for numbers/patterns and Gemini for complex text"""
    if not gemini_data:
        return regex_data
        
    merged = regex_data.copy()
    
    # Use Gemini for clean names if Regex failed or seems short
    if merged["Name"] == "NA" or len(merged["Name"]) < 3:
        merged["Name"] = gemini_data.get("Name", "NA")
        
    if merged["Mother"] == "NA":
        merged["Mother"] = gemini_data.get("Mother", "NA")

    # For marks, if regex found them, trust regex (it extracts exact numbers). 
    # If regex is NA, try Gemini.
    for key in merged:
        if merged[key] == "NA" and key in gemini_data:
            merged[key] = gemini_data[key]
            
    # Post-process MOOC - LEARN NEW SKILLS
    mooc_key = "MOOC - LEARN NEW SKILLS"
    if merged.get(mooc_key) in ["NA", "MOOC", None]:
        merged[mooc_key] = "P"
        
    return merged

def process_page(page, page_num):
    """Process a single PDF page"""
    try:
        print(f"[DEBUG] Processing Page {page_num}...", flush=True)
        # Render page to image efficiently
        pix = page.get_pixmap(dpi=300)
        mode = "RGBA" if pix.alpha else "RGB"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        
        # Enhanced processing
        img = enhance_image(img)
        
        text = extract_text_from_image(img)
        print(f"[DEBUG] Page {page_num} Text Snippet: {text[:100]!r}", flush=True)
        
        student_sections = split_student_sections(text)
        valid_students = []
        
        for i, section in enumerate(student_sections, 1):
            print(f"[DEBUG] Processing student {i} on page {page_num}...", flush=True)
            
            # 1. Regex Extraction
            regex_data = extract_data_regex(section)
            
            # 2. Gemini Extraction
            prompt = format_for_gemini(section)
            gemini_raw = parse_with_gemini(prompt, page_num, i)
            gemini_data = process_student_from_data(gemini_raw) if gemini_raw else {}
            
            # 3. Merge
            final_student = merge_data(regex_data, gemini_data)
            valid_students.append(final_student)
        
        return valid_students, text
    
    except Exception as e:
        print(f"[ERROR] Page {page_num}: {e}")
        return [], ""

def main():
    print("🚀 Starting marksheet extraction...")
    print(f"📂 Input File: {INPUT_PDF}")
    
    if not os.path.exists(INPUT_PDF):
        print(f"❌ Input file not found. Please check: {INPUT_PDF}")
        return

    try:
        doc = fitz.open(INPUT_PDF)
        all_students = []
        total_students_found = 0
        
        print(f"📄 Found {len(doc)} pages.")
        
        for page_index in range(len(doc)):
            page_num = page_index + 1
            page = doc.load_page(page_index)
            
            students, _ = process_page(page, page_num)
            
            if students:
                all_students.extend(students)
                total_students_found += len(students)
                
                # Preview First Student on Page
                s = students[0]
                print(f"[{page_num}] Extracted Student:", flush=True)
                print(f"Seat No: {s.get('Seat No')}", flush=True)
                print(f"Name: {s.get('Name')}", flush=True)
                print(f"Mother: {s.get('Mother')}", flush=True)
                print(f"PRN: {s.get('PRN')}", flush=True)
                
                print("--- Subject Marks ---", flush=True)
                subjects_to_print = [k for k in REQUIRED_COLUMNS if k not in ["Seat No", "Name", "Mother", "PRN", "SGPA", "TOTAL CREDITS"]]
                for subj in subjects_to_print:
                     val = s.get(subj, 'NA')
                     if val != "NA":
                        print(f"{subj}: {val}", flush=True)
                
                print(f"SGPA: {s.get('SGPA')}", flush=True)
                print(f"TOTAL CREDITS: {s.get('TOTAL CREDITS')}", flush=True)
                print("---------------------------------", flush=True)
                print(f"[INFO] Pages processed: {page_num} / {len(doc)}")
                print(f"[INFO] Students extracted so far: {total_students_found}")
                print("-" * 30)
            else:
                print(f"[INFO] Page {page_num}: No student data found.")

            # Save progress after every page
            try:
                temp_df = pd.DataFrame(all_students)
                temp_df = temp_df.reindex(columns=REQUIRED_COLUMNS)
                temp_df.to_excel(OUTPUT_XLSX, index=False)
                print(f"[INFO] 💾 Updated {OUTPUT_XLSX} ({len(all_students)} students)", flush=True)
            except PermissionError:
                print(f"[WARN] ⚠️ Could not update {OUTPUT_XLSX} because it is open. Please close it!", flush=True)
            except Exception as e:
                print(f"[WARN] ⚠️ Failed to save progress: {e}", flush=True)

        # Final Save
        if all_students:
            df = pd.DataFrame(all_students)
            # Ensure proper column ordering
            df = df.reindex(columns=REQUIRED_COLUMNS)
            df.to_excel(OUTPUT_XLSX, index=False)
            print(f"\n✅ Processing Complete! Processed {len(doc)} pages.")
            print(f"📁 Results saved to: {OUTPUT_XLSX}")
        else:
            print("\n⚠️ No student data extracted.")

    except Exception as e:
        print(f"\n❌ Fatal Error: {e}")

if __name__ == "__main__":
    main()


