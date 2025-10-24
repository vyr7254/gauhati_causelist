import os
import time
import logging
import tempfile
import re
from datetime import datetime, timedelta
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import pdfplumber
import pandas as pd

# === CONFIGURATION ===
OUTPUT_FOLDER = r"C:\Users\Dell\OneDrive\Desktop\gauhatistate_hc\gauhati_causelists"
LOG_FILE = os.path.join(OUTPUT_FOLDER, "gauhati_download_log.txt")
EXCEL_FILE = os.path.join(OUTPUT_FOLDER, "gauhati_causelists_data.xlsx")
CAUSELIST_URL = "https://hcservices.ecourts.gov.in/ecourtindiaHC/cases/highcourt_causelist.php?state_cd=6&dist_cd=1&court_code=1&stateNm=Assam"

# Date range configuration
START_DATE = datetime(2025, 9, 1)
END_DATE = datetime(2025, 10, 30)

# === LOGGING SETUP ===
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# === CHROME DRIVER SETUP ===
def setup_driver():
    """Configure Chrome driver with download preferences."""
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    temp_dir = tempfile.mkdtemp()
    chrome_options.add_argument(f"--user-data-dir={temp_dir}")
    
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    prefs = {
        "download.default_directory": OUTPUT_FOLDER,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "plugins.plugins_disabled": ["Chrome PDF Viewer"],
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )
    return driver


def wait_for_download(download_folder, timeout=60):
    """Wait for download to complete."""
    seconds = 0
    while seconds < timeout:
        time.sleep(1)
        files = os.listdir(download_folder)
        if not any(f.endswith('.crdownload') or f.endswith('.tmp') for f in files):
            time.sleep(2)
            return True
        seconds += 1
    logging.warning(f"Download timeout after {timeout} seconds")
    return False


def get_latest_pdf(folder):
    """Get the most recently downloaded PDF."""
    pdf_files = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
    if not pdf_files:
        return None
    pdf_files.sort(key=lambda x: os.path.getmtime(os.path.join(folder, x)), reverse=True)
    return pdf_files[0]


# === DATE PICKER INTERACTION ===
def select_date_in_picker(driver, target_date):
    """Select a specific date in the date picker."""
    try:
        wait = WebDriverWait(driver, 15)
        date_str = target_date.strftime("%d-%m-%Y")
        
        selectors = [
            (By.ID, "date_in_01"),
            (By.NAME, "date_in_01"),
            (By.XPATH, "//input[@type='text' and contains(@placeholder, 'date')]"),
            (By.XPATH, "//input[@type='text' and @id='date_in_01']"),
            (By.XPATH, "//input[contains(@class, 'date')]"),
            (By.CSS_SELECTOR, "input[type='text'][id='date_in_01']")
        ]
        
        date_input = None
        for by_type, selector in selectors:
            try:
                date_input = wait.until(EC.element_to_be_clickable((by_type, selector)))
                logging.info(f"Found date input using: {by_type} = {selector}")
                break
            except:
                continue
        
        if not date_input:
            logging.error("Could not find date input field with any selector")
            return False
        
        date_input.click()
        time.sleep(0.5)
        date_input.clear()
        time.sleep(0.5)
        
        driver.execute_script("arguments[0].value = arguments[1];", date_input, date_str)
        time.sleep(0.5)
        date_input.send_keys(date_str)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", date_input)
        
        logging.info(f"✅ Selected date: {date_str}")
        time.sleep(1)
        return True
        
    except Exception as e:
        logging.error(f"Error selecting date: {e}")
        try:
            screenshot_path = os.path.join(OUTPUT_FOLDER, f"error_date_picker_{int(time.time())}.png")
            driver.save_screenshot(screenshot_path)
            logging.info(f"Screenshot saved: {screenshot_path}")
        except:
            pass
        return False


def click_go_button(driver):
    """Click the GO button to load cause lists."""
    try:
        wait = WebDriverWait(driver, 10)
        
        selectors = [
            (By.XPATH, "//input[@value='GO']"),
            (By.XPATH, "//input[@value='Go']"),
            (By.XPATH, "//button[contains(text(), 'GO')]"),
            (By.XPATH, "//input[@type='submit' and contains(@value, 'GO')]"),
            (By.CSS_SELECTOR, "input[value='GO']")
        ]
        
        go_button = None
        for by_type, selector in selectors:
            try:
                go_button = wait.until(EC.element_to_be_clickable((by_type, selector)))
                logging.info(f"Found GO button using: {by_type} = {selector}")
                break
            except:
                continue
        
        if not go_button:
            logging.error("Could not find GO button with any selector")
            return False
        
        go_button.click()
        logging.info("✅ Clicked GO button")
        time.sleep(3)
        return True
        
    except Exception as e:
        logging.error(f"Error clicking GO button: {e}")
        return False


# === CAUSELIST TABLE PROCESSING ===
def get_causelist_table_rows(driver):
    """Extract all rows from the causelist table."""
    try:
        wait = WebDriverWait(driver, 10)
        
        table = wait.until(
            EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'table') or .//th[contains(text(), 'Bench')] or .//th[contains(text(), 'Sr No')]]"))
        )
        
        try:
            tbody = table.find_element(By.TAG_NAME, "tbody")
            rows = tbody.find_elements(By.TAG_NAME, "tr")
        except:
            all_rows = table.find_elements(By.TAG_NAME, "tr")
            rows = all_rows[1:] if len(all_rows) > 1 else all_rows
        
        logging.info(f"Found {len(rows)} causelist entries in table")
        
        # Extract bench names for each row
        causelist_data = []
        for idx, row in enumerate(rows, start=1):
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 3:
                sr_no = cells[0].text.strip()
                bench_name = cells[1].text.strip() if len(cells) > 1 else "N/A"
                causelist_type = cells[2].text.strip() if len(cells) > 2 else "N/A"
                
                causelist_data.append({
                    'sr_no': sr_no,
                    'bench_name': bench_name,
                    'causelist_type': causelist_type,
                    'row': row
                })
        
        return causelist_data
        
    except TimeoutException:
        logging.warning("No causelist table found for this date")
        return []
    except Exception as e:
        logging.error(f"Error getting table rows: {e}")
        return []


def download_causelist_pdf(driver, row_data, current_date):
    """Download PDF for a specific causelist row."""
    try:
        row = row_data['row']
        sr_no = row_data['sr_no']
        bench_name = row_data['bench_name']
        
        cells = row.find_elements(By.TAG_NAME, "td")
        
        if len(cells) < 3:
            logging.warning(f"  Sr No {sr_no}: Row has insufficient columns ({len(cells)})")
            return None, bench_name
        
        sr_no_text = cells[0].text.strip()
        bench_text = cells[1].text.strip() if len(cells) > 1 else "Unknown"
        causelist_type = cells[2].text.strip() if len(cells) > 2 else "Unknown"
        
        logging.info(f"  Sr No {sr_no_text}: Bench - {bench_text}, Type - {causelist_type}")
        
        view_link = None
        try:
            view_link = cells[-1].find_element(By.LINK_TEXT, "View")
        except:
            try:
                view_link = cells[-1].find_element(By.PARTIAL_LINK_TEXT, "View")
            except:
                try:
                    view_link = cells[-1].find_element(By.TAG_NAME, "a")
                except:
                    logging.warning(f"    ⚠️ Could not find View link for Sr No {sr_no_text}")
                    return None, bench_name
        
        if not view_link:
            logging.warning(f"    ⚠️ No View link found for Sr No {sr_no_text}")
            return None, bench_name
        
        main_window = driver.current_window_handle
        view_link.click()
        time.sleep(3)
        
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(2)
            
            try:
                download_selectors = [
                    (By.XPATH, "//button[contains(@title, 'Download')]"),
                    (By.XPATH, "//button[contains(@class, 'download')]"),
                    (By.XPATH, "//a[contains(@title, 'Download')]"),
                    (By.XPATH, "//button[contains(text(), 'Download')]"),
                    (By.ID, "download"),
                    (By.CSS_SELECTOR, "button[title*='Download']")
                ]
                
                download_btn = None
                for by_type, selector in download_selectors:
                    try:
                        download_btn = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((by_type, selector))
                        )
                        break
                    except:
                        continue
                
                if download_btn:
                    download_btn.click()
                    logging.info(f"    ✅ Clicked download button for Sr No {sr_no_text}")
                else:
                    logging.info(f"    📄 PDF opened (auto-download expected) for Sr No {sr_no_text}")
                
            except TimeoutException:
                logging.info(f"    📄 PDF auto-downloading for Sr No {sr_no_text}")
            
            if wait_for_download(OUTPUT_FOLDER, timeout=40):
                latest_pdf = get_latest_pdf(OUTPUT_FOLDER)
                if latest_pdf:
                    date_str = current_date.strftime("%d-%m-%Y")
                    new_name = f"causelist_{date_str}_{sr_no_text}.pdf"
                    
                    old_path = os.path.join(OUTPUT_FOLDER, latest_pdf)
                    new_path = os.path.join(OUTPUT_FOLDER, new_name)
                    
                    if os.path.exists(new_path):
                        logging.info(f"    ⚠️ PDF already exists: {new_name}")
                        try:
                            os.remove(old_path)
                        except:
                            pass
                    else:
                        try:
                            os.rename(old_path, new_path)
                            logging.info(f"    ✅ Downloaded: {new_name}")
                        except Exception as e:
                            logging.error(f"    ❌ Error renaming file: {e}")
                            new_name = latest_pdf
                    
                    driver.close()
                    driver.switch_to.window(main_window)
                    time.sleep(1)
                    return new_name, bench_name
            
            driver.close()
            driver.switch_to.window(main_window)
            time.sleep(1)
        else:
            if wait_for_download(OUTPUT_FOLDER, timeout=30):
                latest_pdf = get_latest_pdf(OUTPUT_FOLDER)
                if latest_pdf:
                    date_str = current_date.strftime("%d-%m-%Y")
                    new_name = f"causelist_{date_str}_{sr_no_text}.pdf"
                    
                    old_path = os.path.join(OUTPUT_FOLDER, latest_pdf)
                    new_path = os.path.join(OUTPUT_FOLDER, new_name)
                    
                    if not os.path.exists(new_path):
                        os.rename(old_path, new_path)
                        logging.info(f"    ✅ Downloaded: {new_name}")
                        return new_name, bench_name
        
        return None, bench_name
        
    except Exception as e:
        logging.error(f"    ❌ Error downloading Sr No {sr_no}: {e}")
        try:
            if len(driver.window_handles) > 1:
                driver.close()
            driver.switch_to.window(driver.window_handles[0])
        except:
            pass
        return None, bench_name


# === PDF EXTRACTION FUNCTIONS ===
def extract_header_info(pdf_text):
    """Extract court hall number and time from PDF header."""
    court_no = "N/A"
    time_info = "N/A"
    
    try:
        # Extract court number
        court_pattern = r'COURT\s*NO\s*:?\s*(\d+)'
        court_match = re.search(court_pattern, pdf_text, re.IGNORECASE)
        if court_match:
            court_no = court_match.group(1)
        
        # Extract time
        time_pattern = r'(\d{1,2}:\d{2}\s*[AP]M\s*to\s*\d{1,2}:\d{2}\s*[AP]M)'
        time_match = re.search(time_pattern, pdf_text, re.IGNORECASE)
        if time_match:
            time_info = time_match.group(1)
        
        # Check for multiple time slots
        time_pattern_multi = r'(\d{1,2}:\d{2}\s*[AP]M\s*to\s*\d{1,2}:\d{2}\s*[AP]M)\s*(\d{1,2}:\d{2}\s*[AP]M\s*to\s*\d{1,2}:\d{2}\s*[AP]M)'
        time_match_multi = re.search(time_pattern_multi, pdf_text, re.IGNORECASE)
        if time_match_multi:
            time_info = f"{time_match_multi.group(1)} {time_match_multi.group(2)}"
            
    except Exception as e:
        logging.error(f"Error extracting header info: {e}")
    
    return court_no, time_info


def extract_date_from_filename(filename):
    """Extract date from filename format: causelist_DD-MM-YYYY_X.pdf"""
    try:
        date_pattern = r'causelist_(\d{2}-\d{2}-\d{4})_\d+\.pdf'
        match = re.search(date_pattern, filename)
        if match:
            return match.group(1)
        return "N/A"
    except:
        return "N/A"


def parse_gauhati_causelist(pdf_path, bench_name_from_table):
    """Parse Gauhati High Court causelist PDF using positional text parsing."""
    cases = []
    
    try:
        pdf_filename = os.path.basename(pdf_path)
        logging.info(f"\n{'='*60}")
        logging.info(f"📄 Extracting: {pdf_filename}")
        logging.info(f"{'='*60}")
        
        # Extract date from filename
        cause_date = extract_date_from_filename(pdf_filename)
        
        # Extract text using pdfplumber with layout preservation
        all_text = ""
        
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                # Use extract_text with layout=True to preserve column positions
                page_text = page.extract_text(layout=True)
                if page_text:
                    all_text += page_text + "\n"
        
        if not all_text:
            logging.warning(f"No text extracted from {pdf_filename}")
            return cases
        
        # Extract header information
        court_no, time_info = extract_header_info(all_text)
        bench_name = bench_name_from_table if bench_name_from_table != "N/A" else "N/A"
        
        logging.info(f"Court No: {court_no}")
        logging.info(f"Time: {time_info}")
        logging.info(f"Bench: {bench_name}")
        logging.info(f"Date: {cause_date}")
        
        # Split text into lines
        lines = all_text.split('\n')
        
        # Find the header line with column positions
        header_idx = -1
        header_positions = {}
        
        for i, line in enumerate(lines):
            if 'Sr.No' in line and 'Case Number' in line and 'Main Parties' in line:
                header_idx = i
                logging.info(f"Found header at line {i}")
                
                # Find column positions
                if 'Sr.No' in line:
                    header_positions['sr_no'] = line.find('Sr.No')
                if 'Case Number' in line:
                    header_positions['case_number'] = line.find('Case Number')
                if 'Main Parties' in line:
                    header_positions['main_parties'] = line.find('Main Parties')
                if 'Petitioner Advocate' in line:
                    header_positions['pet_advocate'] = line.find('Petitioner Advocate')
                if 'Respondent Advocate' in line:
                    header_positions['resp_advocate'] = line.find('Respondent Advocate')
                
                logging.info(f"Column positions: {header_positions}")
                break
        
        if header_idx == -1:
            logging.warning(f"No header found in {pdf_filename}")
            return cases
        
        # Start parsing from the line after header and separator
        i = header_idx + 1
        
        # Skip separator lines (dashes)
        while i < len(lines) and '---' in lines[i]:
            i += 1
        
        logging.info(f"Starting case extraction from line {i}")
        
        # Parse cases
        case_count = 0
        while i < len(lines):
            line = lines[i]
            
            # Check if line starts with a number (Sr.No) at the beginning
            stripped = line.lstrip()
            sr_match = re.match(r'^(\d+)\s', stripped)
            
            if sr_match:
                case_count += 1
                sr_no = sr_match.group(1)
                
                logging.info(f"\n--- Case {case_count}: Sr.No {sr_no} ---")
                logging.info(f"Line {i}: {line[:100]}")
                
                # Collect all lines for this case until next Sr.No or section break
                case_lines = [line]
                j = i + 1
                
                while j < len(lines):
                    next_line = lines[j]
                    next_stripped = next_line.lstrip()
                    
                    # Stop if we hit next case number
                    if re.match(r'^\d+\s', next_stripped):
                        break
                    
                    # Stop if we hit a major section break
                    if '===' in next_line or 'LEAVE NOTE' in next_line:
                        break
                    
                    # Skip pure separator lines
                    if next_line.strip() and not next_line.strip().replace('-', '').strip():
                        j += 1
                        continue
                    
                    if next_line.strip():
                        case_lines.append(next_line)
                    
                    j += 1
                
                # Join all case lines
                full_case_text = '\n'.join(case_lines)
                
                # Extract case number from first line
                case_pattern = r'([A-Z]+(?:\([A-Z]\))?(?:\.[A-Z]+)?(?:\([A-Za-z]+\))?)/(\d+)/(\d{4})'
                case_match = re.search(case_pattern, full_case_text)
                
                case_type = "N/A"
                case_number = "N/A"
                case_year = "N/A"
                
                if case_match:
                    case_type = case_match.group(1)
                    case_number = case_match.group(2)
                    case_year = case_match.group(3)
                    logging.info(f"Case: {case_type}/{case_number}/{case_year}")
                else:
                    # Alternative pattern for complex types
                    alt_pattern = r'([A-Z\.\(\)]+)/(\d+)/(\d{4})'
                    alt_match = re.search(alt_pattern, full_case_text)
                    if alt_match:
                        case_type = alt_match.group(1)
                        case_number = alt_match.group(2)
                        case_year = alt_match.group(3)
                        logging.info(f"Case (alt): {case_type}/{case_number}/{case_year}")
                
                # Extract parties by splitting on "Versus"
                petitioner = "N/A"
                respondent = "N/A"
                petitioner_advocate = "N/A"
                respondent_advocate = "N/A"
                
                if 'Versus' in full_case_text:
                    # Split the entire text by Versus
                    parts = full_case_text.split('Versus', 1)
                    
                    # Part 1: Contains Sr.No, Case Number, and Petitioner
                    before_versus = parts[0]
                    
                    # Part 2: Contains Respondent and Advocates
                    after_versus = parts[1] if len(parts) > 1 else ""
                    
                    # Extract petitioner - remove sr no and case number
                    pet_text = before_versus
                    pet_text = re.sub(r'^\d+\s+', '', pet_text)  # Remove sr no
                    if case_match:
                        pet_text = pet_text.replace(case_match.group(0), '')  # Remove case number
                    
                    # Clean up petitioner
                    pet_lines = [l.strip() for l in pet_text.split('\n') if l.strip()]
                    # Filter out any WITH or other keywords
                    pet_lines = [l for l in pet_lines if not l.startswith('WITH') and not l.startswith('in ')]
                    petitioner = ' '.join(pet_lines).strip()
                    
                    # Process after_versus section
                    after_lines = [l.strip() for l in after_versus.split('\n') if l.strip()]
                    
                    # Separate respondent from advocates
                    resp_lines = []
                    pet_adv_lines = []
                    resp_adv_lines = []
                    
                    found_advocate = False
                    
                    for line_text in after_lines:
                        # Check if this line contains advocate keywords
                        has_advocate = any(kw in line_text.upper() for kw in ['MR.', 'MRS.', 'MS.', 'DR.', 'ADVOCATE', 'SC,', 'GA,', 'PP,'])
                        
                        if not found_advocate and not has_advocate:
                            # This is respondent
                            resp_lines.append(line_text)
                        elif has_advocate:
                            found_advocate = True
                            # This is advocate line
                            # Respondent advocates have (R- or (r- pattern
                            if '(R-' in line_text or '(r-' in line_text or '(R1' in line_text or '(R2' in line_text:
                                resp_adv_lines.append(line_text)
                            else:
                                # Petitioner advocate (no R- marker)
                                pet_adv_lines.append(line_text)
                    
                    respondent = ' '.join(resp_lines).strip() if resp_lines else "N/A"
                    petitioner_advocate = ' '.join(pet_adv_lines).strip() if pet_adv_lines else "N/A"
                    respondent_advocate = ' '.join(resp_adv_lines).strip() if resp_adv_lines else "N/A"
                    
                else:
                    # No Versus found - might be a WITH clause or other format
                    petitioner = full_case_text
                    if case_match:
                        petitioner = petitioner.replace(case_match.group(0), '').strip()
                    petitioner = re.sub(r'^\d+\s+', '', petitioner).strip()
                
                logging.info(f"Petitioner: {petitioner[:70]}")
                logging.info(f"Respondent: {respondent[:70]}")
                logging.info(f"Pet Advocate: {petitioner_advocate[:70]}")
                logging.info(f"Resp Advocate: {respondent_advocate[:70]}")
                
                # Create case entry
                case_data = {
                    "id": None,
                    "causelist_slno": sr_no,
                    "court_hall_number": court_no,
                    "Case_number": case_number,
                    "Case_type": case_type,
                    "case_year": case_year,
                    "bench_name": bench_name,
                    "cause_date": cause_date,
                    "time": time_info,
                    "petitioner": petitioner,
                    "respondent": respondent,
                    "petitioner_advocate": petitioner_advocate,
                    "respondent_advocate": respondent_advocate,
                    "particulars": "List Downloaded",
                    "Pdf_name": pdf_filename,
                    "case_status": "N/A"
                }
                
                cases.append(case_data)
                
                # Move to next case
                i = j
            else:
                i += 1
        
        logging.info(f"\n{'='*60}")
        logging.info(f"✅ Extracted {len(cases)} cases from {pdf_filename}")
        logging.info(f"{'='*60}\n")
        
    except Exception as e:
        logging.error(f"❌ Error processing {pdf_path}: {e}", exc_info=True)
    
    return cases


def save_to_excel(cases_data, excel_path):
    """Save or append case data to Excel file."""
    try:
        if not cases_data:
            logging.warning("No case data to save")
            return False
        
        columns = [
            "id", "causelist_slno", "court_hall_number", "Case_number", "Case_type",
            "case_year", "bench_name", "cause_date", "time",
            "petitioner", "respondent", "petitioner_advocate", "respondent_advocate",
            "particulars", "Pdf_name", "case_status"
        ]
        
        df_new = pd.DataFrame(cases_data)
        
        for col in columns:
            if col not in df_new.columns:
                df_new[col] = "N/A"
        
        df_new = df_new[columns]
        
        if os.path.exists(excel_path):
            df_existing = pd.read_excel(excel_path)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_combined["id"] = range(1, len(df_combined) + 1)
            df_combined.to_excel(excel_path, index=False)
            logging.info(f"✅ Appended {len(df_new)} cases → Total: {len(df_combined)}")
        else:
            df_new["id"] = range(1, len(df_new) + 1)
            df_new.to_excel(excel_path, index=False)
            logging.info(f"✅ Created Excel file with {len(df_new)} cases")
        
        return True
        
    except Exception as e:
        logging.error(f"Error saving to Excel: {e}")
        return False


# === MAIN EXECUTION ===
def main():
    logging.info("=" * 80)
    logging.info("GAUHATI HIGH COURT CAUSELIST PDF DOWNLOADER & EXTRACTOR")
    logging.info("=" * 80)
    
    driver = setup_driver()
    total_pdfs_downloaded = 0
    total_cases_extracted = 0
    failed_downloads = []
    
    try:
        driver.get(CAUSELIST_URL)
        time.sleep(3)
        logging.info(f"Opened URL: {CAUSELIST_URL}")
        
        current_date = START_DATE
        
        while current_date <= END_DATE:
            logging.info("\n" + "=" * 80)
            logging.info(f"PROCESSING DATE: {current_date.strftime('%d-%m-%Y')}")
            logging.info("=" * 80)
            
            if not select_date_in_picker(driver, current_date):
                logging.error(f"Failed to select date: {current_date}")
                failed_downloads.append(f"{current_date.strftime('%d-%m-%Y')} - Date selection failed")
                current_date += timedelta(days=1)
                continue
            
            if not click_go_button(driver):
                logging.error("Failed to click GO button")
                failed_downloads.append(f"{current_date.strftime('%d-%m-%Y')} - GO button click failed")
                current_date += timedelta(days=1)
                continue
            
            # Get causelist rows with bench names
            causelist_data = get_causelist_table_rows(driver)
            
            if not causelist_data:
                logging.warning(f"No cause lists found for {current_date.strftime('%d-%m-%Y')}")
                current_date += timedelta(days=1)
                continue
            
            # Process each row
            date_pdfs = 0
            for row_data in causelist_data:
                sr_no = row_data['sr_no']
                bench_name = row_data['bench_name']
                
                pdf_filename, bench = download_causelist_pdf(driver, row_data, current_date)
                
                if pdf_filename:
                    total_pdfs_downloaded += 1
                    date_pdfs += 1
                    
                    # Extract data from PDF immediately
                    pdf_path = os.path.join(OUTPUT_FOLDER, pdf_filename)
                    if os.path.exists(pdf_path):
                        cases = parse_gauhati_causelist(pdf_path, bench_name)
                        
                        if cases:
                            if save_to_excel(cases, EXCEL_FILE):
                                total_cases_extracted += len(cases)
                        else:
                            logging.warning(f"⚠️ No cases extracted from {pdf_filename}")
                    else:
                        logging.error(f"❌ PDF file not found: {pdf_path}")
                else:
                    failed_downloads.append(f"{current_date.strftime('%d-%m-%Y')} - Sr No {sr_no}")
                
                time.sleep(2)
            
            logging.info(f"Downloaded {date_pdfs} PDFs for {current_date.strftime('%d-%m-%Y')}")
            
            # Move to next date
            current_date += timedelta(days=1)
            time.sleep(3)
        
        # Final summary
        logging.info("\n" + "=" * 80)
        logging.info("PDF DOWNLOAD & EXTRACTION COMPLETED")
        logging.info("=" * 80)
        logging.info(f"Total PDFs Downloaded: {total_pdfs_downloaded}")
        logging.info(f"Total Cases Extracted: {total_cases_extracted}")
        logging.info(f"Failed Downloads: {len(failed_downloads)}")
        
        if failed_downloads:
            logging.info("\nFailed Download Details:")
            for fail in failed_downloads:
                logging.info(f"  ❌ {fail}")
        
        logging.info(f"\nPDFs saved to: {OUTPUT_FOLDER}")
        logging.info(f"Excel file saved to: {EXCEL_FILE}")
        logging.info(f"Log file saved to: {LOG_FILE}")
        
    except Exception as e:
        logging.error(f"Critical error in main execution: {e}", exc_info=True)
        
    finally:
        driver.quit()
        logging.info("\nBrowser closed. Process finished.")


if __name__ == "__main__":
    main()