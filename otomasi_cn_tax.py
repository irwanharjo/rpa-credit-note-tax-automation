import pandas as pd
import os
import glob
import datetime
import shutil
import time
import base64
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
import win32com.client as win32
from dotenv import load_dotenv

# Official Google API libraries
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.encoders import encode_base64

# Selenium for headless web automation
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

load_dotenv()

SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]

# --- Configuration (all client-specific values come from .env — see .env.example) ---
SPREADSHEET_ID = os.environ.get('TAX_TRACKING_SPREADSHEET_ID', '')
SHEET_NAME = os.environ.get('TAX_TRACKING_SHEET_NAME', 'Tax Invoice Requests')

PORTAL_URL = os.environ.get('PORTAL_BASE_URL', 'https://your-claims-portal.example.com')
PORTAL_USERNAME = os.environ.get('PORTAL_USERNAME', '')
PORTAL_PASSWORD = os.environ.get('PORTAL_PASSWORD', '')

CLIENT_CONTACT_EMAIL = os.environ.get('CLIENT_CONTACT_EMAIL', 'billing-contact@client.example.com')
CLIENT_TEAM_NAME = os.environ.get('CLIENT_TEAM_NAME', 'Client Billing Team')

SENDER_NAME = os.environ.get('SENDER_NAME', 'Your Name')
SENDER_SIGNATURE = os.environ.get(
    'SENDER_SIGNATURE',
    f"{SENDER_NAME}\nAR Department\nYour Company\nYour City"
)

# Illustrative batch-code -> client entity map. Replace with your own client list.
ESELON_MAP = {
    "BATCH_A": "CLIENT A ENTITY",
    "BATCH_B": "CLIENT B ENTITY",
    "BATCH_C": "CLIENT B ENTITY",
    "BATCH_D": "CLIENT C ENTITY",
    "BATCH_E": "CLIENT B ENTITY",
    "BATCH_F": "CLIENT D ENTITY",
    "BATCH_G": "CLIENT E ENTITY",
    "BATCH_H": "CLIENT F ENTITY",
    "BATCH_I": "CLIENT G ENTITY",
    "BATCH_J": "CLIENT H ENTITY",
}


# --- GMAIL API HELPERS ---
def find_email_thread(service, query):
    if not query or query == "-": return None, None, None
    try:
        gmail_q = f"from:{CLIENT_CONTACT_EMAIL} {query}"
        result = service.users().messages().list(userId='me', q=gmail_q, maxResults=1).execute()
        messages = result.get('messages', [])
        if messages:
            msg_id = messages[0]['id']
            thread_id = messages[0]['threadId']
            msg_details = service.users().messages().get(userId='me', id=msg_id, format='metadata', metadataHeaders=['Subject', 'Message-ID']).execute()
            headers = msg_details.get('payload', {}).get('headers', [])
            subject = "Recalculated Tax (PPN) Value"
            msg_mime_id = None
            for h in headers:
                if h['name'].lower() == 'subject':
                    subject = h['value']
                    if not subject.lower().startswith('re:'): subject = f"Re: {subject}"
                if h['name'].lower() == 'message-id':
                    msg_mime_id = h['value']
            return thread_id, msg_mime_id, subject
    except Exception: pass
    return None, None, None

def send_gmail_message(service, to_email, subject, body_text, attachment_path=None, thread_id=None, message_id=None, is_html=False, inline_images=None):
    message = MIMEMultipart('related') if is_html else MIMEMultipart()
    message['to'] = to_email
    message['subject'] = subject
    if thread_id and message_id:
        message['In-Reply-To'] = message_id
        message['References'] = message_id

    if is_html:
        msg_alt = MIMEMultipart('alternative')
        message.attach(msg_alt)
        msg_alt.attach(MIMEText(body_text, 'html'))
        if inline_images:
            for idx, img_path in enumerate(inline_images):
                with open(img_path, 'rb') as f:
                    img = MIMEImage(f.read())
                    img.add_header('Content-ID', f'<img_{idx}>')
                    img.add_header('Content-Disposition', 'inline')
                    message.attach(img)
    else:
        message.attach(MIMEText(body_text, 'plain'))

    if attachment_path and os.path.exists(attachment_path):
        filename = os.path.basename(attachment_path)
        with open(attachment_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            message.attach(part)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    send_body = {'raw': raw_message}
    if thread_id: send_body['threadId'] = thread_id
    service.users().messages().send(userId='me', body=send_body).execute()

# --- SELENIUM BOT HELPERS ---
def init_selenium_driver():
    print("   [BOT] Starting headless Chrome engine...")
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--window-size=1920,1080")

    prefs = {
        "profile.default_content_setting_values.notifications": 1,
        "profile.default_content_setting_values.geolocation": 1,
        "profile.managed_default_content_settings.popups": 1
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.set_capability("unhandledPromptBehavior", "accept")

    try:
        driver = webdriver.Chrome(options=chrome_options)
        return driver
    except Exception as e:
        print(f"   [BOT ERROR] Failed to launch Chrome. Error: {e}")
        return None

def login_portal(driver, username, password):
    print("   [BOT] Logging into claims portal...")
    try:
        driver.get(f"{PORTAL_URL}/auth/sign-in")
        time.sleep(3)
        user_field = driver.find_element(By.XPATH, "//input[@type='text' or contains(@name, 'user') or contains(@name, 'email')]")
        pass_field = driver.find_element(By.XPATH, "//input[@type='password']")
        user_field.send_keys(username)
        pass_field.send_keys(password)
        login_btn = driver.find_element(By.XPATH, "//button[@type='submit' or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'login') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'masuk')]")
        login_btn.click()

        try:
            WebDriverWait(driver, 3).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            alert.accept()
            print("   [BOT] Auto-accepted a permission popup.")
        except Exception:
            pass

        time.sleep(3)
        return True
    except Exception as e:
        print(f"   [BOT ERROR] Portal login failed: {e}")
        return False

# --- MAIN PROCESS ---
def process_batch_credit_notes():
    base_dir = os.getcwd()
    input_folder = os.path.join(base_dir, '1_Input_Tagihan_Masuk')
    output_folder = os.path.join(base_dir, '2_Hasil_Perhitungan_Selesai')
    pdf_folder = os.path.join(base_dir, '3_File_PDF_Siap_Kirim')
    recap_folder = os.path.join(base_dir, '4_Rekap_Harian')
    archive_folder = os.path.join(base_dir, '5_Arsip_Tagihan_Asli')
    bukti_folder = os.path.join(base_dir, '6_Bukti_Sanggahan_VVIP')
    dispute_folder = os.path.join(base_dir, '7_Arsip_Tagihan_Dispute')

    for folder in [input_folder, output_folder, pdf_folder, recap_folder, archive_folder, bukti_folder, dispute_folder]:
        if not os.path.exists(folder): os.makedirs(folder)

    excel_files = glob.glob(os.path.join(input_folder, '*.xlsx'))
    if not excel_files:
        print("Input folder is empty! Please add new billing files.")
        return

    google_service, sheets_service = None, None
    if os.path.exists('credentials.json'):
        try:
            creds = None
            if os.path.exists('token.json'): creds = Credentials.from_authorized_user_file('token.json', SCOPES)
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token: creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                with open('token.json', 'w') as token: token.write(creds.to_json())
            google_service = build('gmail', 'v1', credentials=creds)
            sheets_service = build('sheets', 'v4', credentials=creds)
            print("Gmail & Sheets API connection successful!")
        except Exception as e: print(f"API connection failed: {e}")

    excel_app = None
    try:
        excel_app = win32.GetActiveObject("Excel.Application")
        print("   [INFO] Found an active Excel instance! Reusing the one already open.")
    except Exception:
        try:
            excel_app = win32.Dispatch("Excel.Application")
            print("   [INFO] Launching a new Excel instance...")
        except Exception as e:
            print(f"   [ERROR] Failed to launch Excel via win32com: {e}")
            return

    try:
        excel_app.DisplayAlerts = False
    except Exception:
        pass

    resume_list = []
    sheets_batch_data = []

    driver = None
    is_logged_in = False
    vvip_cache = {}

    print(f"Processing {len(excel_files)} file(s)...\n")

    for idx, file_path in enumerate(excel_files, start=1):
        file_name = os.path.basename(file_path)
        print(f"[{idx}/{len(excel_files)}] File: {file_name}")

        nama_baru_base = file_name.replace(".xlsx", "_PERHITUNGAN_PPN")
        excel_output = os.path.abspath(os.path.join(output_folder, nama_baru_base + ".xlsx"))
        pdf_output = os.path.abspath(os.path.join(pdf_folder, nama_baru_base + ".pdf"))

        try:
            df = pd.read_excel(file_path, engine='calamine', header=None)
            no_batch = str(df.iloc[2, 2]).strip() if pd.notna(df.iloc[2, 2]) else "-"
            no_invoice = str(df.iloc[4, 2]).strip() if pd.notna(df.iloc[4, 2]) else "-"

            eselon_val = "OTHER"
            for prefix, entity_name in ESELON_MAP.items():
                if prefix in no_batch:
                    eselon_val = entity_name
                    break

            header_idx = -1
            for i in range(min(30, len(df))):
                baris = df.iloc[i].astype(str).str.lower().tolist()
                if 'kategori layanan' in baris and 'nilai akhir tagihan' in baris:
                    header_idx = i; break

            if header_idx == -1: continue
            baris_header = df.iloc[header_idx].astype(str).str.lower().tolist()

            col_kategori = baris_header.index('kategori layanan')
            col_nilai_awal = baris_header.index('nilai awal tagihan ') if 'nilai awal tagihan ' in baris_header else 12
            col_cn = baris_header.index('kredit nota') if 'kredit nota' in baris_header else 14
            col_nilai_akhir = baris_header.index('nilai akhir tagihan')
            col_status = baris_header.index('status') if 'status' in baris_header else 16
            col_dpp = baris_header.index('dpp') if 'dpp' in baris_header else len(df.columns)
            col_ppn = baris_header.index('ppn') if 'ppn' in baris_header else len(df.columns) + 1

            col_nama, col_nopes, col_tgl, col_catatan = 1, 2, 7, 10
            for idx_h, h_val in enumerate(baris_header):
                if 'nama pasien' in h_val: col_nama = idx_h
                if 'no. peserta' in h_val: col_nopes = idx_h
                if 'tgl item layan' in h_val: col_tgl = idx_h
                if 'catatan tim verif' in h_val: col_catatan = idx_h

            farmasi_count = farmasi_verified_count = 0
            has_reinvoice = has_item_cn = False

            current_pasien = ""
            current_nopes = ""
            current_tgl_picker = ""
            current_tgl_table = ""

            kandidat_dispute = {}
            pasien_tidak_preot = set()

            for i in range(header_idx + 1, len(df)):
                val_nama = str(df.iloc[i, col_nama]).strip()
                val_nopes = str(df.iloc[i, col_nopes]).strip()
                val_tgl = df.iloc[i, col_tgl]

                if val_nama and val_nama.lower() != 'nan': current_pasien = val_nama
                if val_nopes and val_nopes.lower() != 'nan': current_nopes = val_nopes

                if pd.notna(val_tgl) and str(val_tgl).strip() != "":
                    if isinstance(val_tgl, datetime.datetime):
                        current_tgl_picker = val_tgl.strftime("%d-%m-%Y")
                        current_tgl_table = val_tgl.strftime("%d/%m/%Y")
                    else:
                        try:
                            dt_obj = datetime.datetime.strptime(str(val_tgl).strip().split(" ")[0], "%d/%m/%y")
                            current_tgl_picker = dt_obj.strftime("%d-%m-%Y")
                            current_tgl_table = dt_obj.strftime("%d/%m/%Y")
                        except Exception:
                            current_tgl_picker = str(val_tgl).strip()
                            current_tgl_table = str(val_tgl).strip()

                val_catatan = str(df.iloc[i, col_catatan]).strip().lower()
                if any(keyword in val_catatan for keyword in ['tidak preotorisasi', 'tidak preot', 'tidak ada preot']):
                    if current_pasien and "total" not in current_pasien.lower():
                        pasien_tidak_preot.add(current_pasien)
                        if current_nopes:
                            kandidat_dispute[current_nopes] = {
                                'nama': current_pasien,
                                'tgl_picker': current_tgl_picker,
                                'tgl_table': current_tgl_table
                            }

                kategori = str(df.iloc[i, col_kategori]).strip().lower()
                status_val = str(df.iloc[i, col_status]).strip().lower()
                if 'reinvoice' in status_val: has_reinvoice = True
                if 'farmasi' in kategori:
                    farmasi_count += 1
                    if 'verified' in status_val: farmasi_verified_count += 1
                if 'item cn' in kategori: has_item_cn = True

            file_is_disputed = False
            vvip_found_in_file = []

            if kandidat_dispute and PORTAL_USERNAME:
                if driver is None:
                    driver = init_selenium_driver()
                if driver and not is_logged_in:
                    is_logged_in = login_portal(driver, PORTAL_USERNAME, PORTAL_PASSWORD)

                if is_logged_in:
                    print(f"   -> Found {len(kandidat_dispute)} 'not pre-authorized' notes. Checking portal...")
                    for nopes, info in kandidat_dispute.items():
                        if nopes in vvip_cache:
                            if vvip_cache[nopes] == True:
                                file_is_disputed = True
                                existing_ss = glob.glob(os.path.join(bukti_folder, f"*{nopes}.png"))
                                if existing_ss:
                                    vvip_found_in_file.append({"nama": info['nama'], "img": existing_ss[0]})
                            continue

                        try:
                            driver.get(f"{PORTAL_URL}/claim-participants")
                            time.sleep(3)

                            tgl_input = f"{info['tgl_picker']} To {info['tgl_picker']}"
                            date_xpath = '//*[@id="root"]/div/div/div[4]/div[2]/div/div[2]/div[1]/div/div/input'
                            try:
                                driver.execute_script(f"""
                                    var d = document.evaluate('{date_xpath}', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                                    if(d) {{
                                        d.removeAttribute('readonly');
                                        d.value = '{tgl_input}';
                                        d.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                        d.dispatchEvent(new Event('change', {{ bubbles: true }}));
                                    }}
                                """)
                            except Exception: pass
                            time.sleep(1)

                            search_xpath = '//*[@id="root"]/div/div/div[4]/div[2]/div/div[2]/div[2]/input'
                            search_box = driver.find_element(By.XPATH, search_xpath)
                            search_box.clear()
                            search_box.send_keys(nopes)
                            time.sleep(3)

                            try:
                                row_xpath = f"//tr[td[contains(., '{nopes}')] and td[contains(., '{info['tgl_table']}')]]//i[contains(@class, 'eye')]/.."
                                view_btns = driver.find_elements(By.XPATH, row_xpath)

                                if not view_btns:
                                    fallback_xpath = f"//tr[td[contains(., '{nopes}')]]//i[contains(@class, 'eye')]/.."
                                    view_btns = driver.find_elements(By.XPATH, fallback_xpath)

                                if view_btns:
                                    view_btns[0].click()
                                    time.sleep(3)

                                    body_text = driver.find_element(By.TAG_NAME, "body").text
                                    if "Peserta VVIP tidak memerlukan preotorisasi" in body_text:
                                        print("         [!] VALID: VVIP status detected!")
                                        vvip_cache[nopes] = True
                                        file_is_disputed = True

                                        ss_path = os.path.join(bukti_folder, f"{no_invoice}_{nopes}.png")
                                        driver.save_screenshot(ss_path)
                                        vvip_found_in_file.append({"nama": info['nama'], "img": ss_path})
                                    else:
                                        vvip_cache[nopes] = False
                                else:
                                    vvip_cache[nopes] = False
                            except Exception:
                                vvip_cache[nopes] = False

                        except Exception:
                            print(f"      [ERROR] Failed to check portal for {nopes}. Continuing normally.")

            if file_is_disputed:
                print("   [!] PPN RECALCULATION CANCELLED — VVIP CASE FOUND. SENDING DISPUTE EMAIL!")

                archive_path = os.path.join(dispute_folder, file_name)
                if os.path.exists(archive_path):
                    base, ext = os.path.splitext(file_name)
                    archive_path = os.path.join(dispute_folder, f"{base}_{int(time.time())}{ext}")
                shutil.move(file_path, archive_path)

                str_cat_verif = ", ".join(sorted(list(pasien_tidak_preot))) if pasien_tidak_preot else "-"
                resume_list.append({
                    'Nomor Batch': no_batch, 'No. DN': no_invoice, 'Nilai Awal Tagihan': "-", 'Nilai CN': "-",
                    'Nilai Akhir Tagihan': "-", 'Nilai PPN': "-",
                    'Cat. Verifikator': str_cat_verif, 'Keterangan': "HOLD - VVIP Dispute (awaiting client)"
                })

                if google_service:
                    thread_id, msg_mime_id, reply_subject = find_email_thread(google_service, no_invoice)
                    if not thread_id: thread_id, msg_mime_id, reply_subject = find_email_thread(google_service, no_batch)

                    list_nama_html = "".join([f"<li><b>{item['nama']}</b></li>" for item in vvip_found_in_file])
                    img_tags_html = "".join([f'<br><img src="cid:img_{idx}" style="max-width:900px; border:1px solid #ccc;"><br><br>' for idx in range(len(vvip_found_in_file))])

                    body_html = f"""
                    <p>Dear {CLIENT_TEAM_NAME}</p>
                    <p>Please disregard our previous email — on further review, a credit note was applied to the following patient(s) with reason <i>"not covered because the medication is not registered in the 2025 formulary and no pre-authorization was recorded"</i>:</p>
                    <ol>
                        {list_nama_html}
                    </ol>
                    <p>FYI, the patient(s) above are <b>VVIP</b> members who do not require pre-authorization for medication or procedure coverage. (Evidence attached below.)</p>
                    {img_tags_html}
                    <p>Kindly revise the verification result and resend the spreadsheet if a credit note was applied.</p>
                    <p>Thank you for your attention to this matter.</p>
                    <p>Best regards,<br>{SENDER_SIGNATURE}</p>
                    """

                    img_paths = [item['img'] for item in vvip_found_in_file]

                    if thread_id:
                        print("   -> SENDING DISPUTE EMAIL (HTML reply)...")
                        send_gmail_message(google_service, CLIENT_CONTACT_EMAIL, reply_subject, body_html, thread_id=thread_id, message_id=msg_mime_id, is_html=True, inline_images=img_paths)
                    else:
                        print("   -> SENDING NEW DISPUTE EMAIL (HTML)...")
                        new_subject = f"VVIP Verification Dispute - Invoice: {no_invoice}"
                        send_gmail_message(google_service, CLIENT_CONTACT_EMAIL, new_subject, body_html, is_html=True, inline_images=img_paths)
                continue

            butuh_hitung_ulang = (farmasi_count > 0) and (farmasi_verified_count < farmasi_count or has_item_cn)

            if butuh_hitung_ulang:
                keterangan_resume = "Ada Reinvoice" if has_reinvoice else "-"
            else:
                if farmasi_count > 0 and farmasi_verified_count == farmasi_count:
                    keterangan_resume = "Email tidak dikirim karena semua nilai farmasi terverifikasi"
                    if has_reinvoice:
                        keterangan_resume = "Ada Reinvoice, " + keterangan_resume
                else:
                    keterangan_resume = "Ada Reinvoice" if has_reinvoice else "Tidak Ada Perubahan Nilai PPN"

            total_sum_dpp = total_sum_ppn = 0

            if butuh_hitung_ulang:
                if 'dpp' not in baris_header: df[col_dpp] = None; df.iloc[header_idx, col_dpp] = 'DPP'
                if 'ppn' not in baris_header: df[col_ppn] = None; df.iloc[header_idx, col_ppn] = 'PPN'

                blocks = []; current_block = []
                for i in range(header_idx + 1, len(df)):
                    val_a, val_b = str(df.iloc[i, 0]).strip().lower(), str(df.iloc[i, 1]).strip().lower()
                    if "total" in val_b or "total" in val_a:
                        if current_block: blocks.append(current_block)
                        break
                    if pd.notna(df.iloc[i, 0]) and str(df.iloc[i, 0]).strip() != "":
                        if current_block: blocks.append(current_block)
                        current_block = [i]
                    else:
                        current_block.append(i) if current_block else [i]
                if current_block: blocks.append(current_block)

                for block in blocks:
                    farmasi_rows = []; item_cn_rows = []; item_cn_total = 0.0
                    for i in block:
                        kategori = str(df.iloc[i, col_kategori]).strip().lower()
                        if 'farmasi' in kategori: farmasi_rows.append(i)
                        elif 'item cn' in kategori:
                            item_cn_rows.append(i)
                            try: item_cn_total += float(df.iloc[i, col_nilai_akhir])
                            except Exception: pass

                    max_f_idx, max_f_val = -1, -float('inf')
                    if farmasi_rows and item_cn_total < 0:
                        for idx in farmasi_rows:
                            try: val = float(df.iloc[idx, col_nilai_akhir])
                            except Exception: val = 0.0
                            if val > max_f_val: max_f_val, max_f_idx = val, idx

                    for idx in farmasi_rows:
                        try: val = float(df.iloc[idx, col_nilai_akhir])
                        except Exception: val = 0.0
                        net_val = val + item_cn_total if (idx == max_f_idx and item_cn_total < 0) else val
                        dpp = round(net_val / 1.11, 0); ppn = round(net_val - dpp, 0)
                        df.iloc[idx, col_dpp] = dpp; df.iloc[idx, col_ppn] = ppn
                        total_sum_dpp += dpp; total_sum_ppn += ppn

                    for idx in item_cn_rows: df.iloc[idx, col_dpp] = ""; df.iloc[idx, col_ppn] = ""

            grand_awal = grand_cn = grand_akhir = 0
            grand_ppn = "-"

            for r in range(len(df) - 1, header_idx, -1):
                v_akhir = df.iloc[r, col_nilai_akhir]
                if pd.notna(v_akhir) and isinstance(v_akhir, (int, float)) and v_akhir > 0:
                    if pd.isna(df.iloc[r, 0]) or str(df.iloc[r, 0]).strip() == "" or "total" in str(df.iloc[r, 1]).lower():
                        if butuh_hitung_ulang:
                            df.iloc[r, col_dpp], df.iloc[r, col_ppn] = total_sum_dpp, total_sum_ppn
                            grand_ppn = total_sum_ppn
                        try: grand_awal = float(df.iloc[r, col_nilai_awal])
                        except Exception: grand_awal = 0.0
                        try: grand_cn = float(df.iloc[r, col_cn])
                        except Exception: grand_cn = 0.0
                        try: grand_akhir = float(v_akhir)
                        except Exception: grand_akhir = 0.0
                        break

            if isinstance(grand_ppn, (int, float)):
                if grand_ppn < 1000000:
                    row_data = [eselon_val, no_invoice, int(grand_ppn), int(grand_akhir)]
                    sheets_batch_data.append(row_data)
                else:
                    keterangan_resume += " - PPN blm direquest"

            str_cat_verif = ", ".join(sorted(list(pasien_tidak_preot))) if pasien_tidak_preot else "-"
            resume_list.append({
                'Nomor Batch': no_batch, 'No. DN': no_invoice, 'Nilai Awal Tagihan': grand_awal, 'Nilai CN': grand_cn,
                'Nilai Akhir Tagihan': grand_akhir, 'Nilai PPN': grand_ppn,
                'Cat. Verifikator': str_cat_verif, 'Keterangan': keterangan_resume
            })

            pdf_terbentuk = False
            if butuh_hitung_ulang:
                df.to_excel(excel_output, index=False, header=False)
                wb = load_workbook(excel_output)
                ws = wb.active
                gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                thin = Side(border_style="thin", color="000000")
                f_border = Border(left=thin, right=thin, top=thin, bottom=thin)
                c_widths = {1:5, 2:25, 3:20, 9:15, 10:35, 13:22, 14:22, 16:22, 19:18, 20:18}
                for c_idx, w in c_widths.items(): ws.column_dimensions[ws.cell(row=1, column=c_idx).column_letter].width = w

                for row in ws.iter_rows(min_row=header_idx+1):
                    is_header = (row[0].row == header_idx + 1)
                    for cell in row:
                        cell.border = f_border; cell.alignment = Alignment(wrap_text=True, vertical="center")
                        if is_header:
                            cell.font = Font(size=12, bold=True)
                            val_header = str(cell.value).strip().lower() if cell.value else ""
                            cell.fill = yellow_fill if val_header in ['dpp', 'ppn'] else gray_fill
                        else:
                            cell.font = Font(size=12)
                        if isinstance(cell.value, (int, float)) and cell.row > header_idx+1:
                            cell.number_format = '#,##0'

                ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1
                ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
                wb.save(excel_output)
                try: wb.close()
                except Exception: pass

                time.sleep(1)

                wb_pdf = None
                error_pdf_detail = "Unknown"

                for attempt in range(1, 4):
                    try:
                        wb_pdf = excel_app.Workbooks.Open(excel_output, False, True)
                        time.sleep(1)
                        wb_pdf.ExportAsFixedFormat(0, pdf_output)
                        wb_pdf.Close(False)
                        wb_pdf = None
                        if os.path.exists(pdf_output):
                            pdf_terbentuk = True
                            break
                    except Exception as pdf_err:
                        error_pdf_detail = str(pdf_err)
                        if wb_pdf:
                            try: wb_pdf.Close(False)
                            except Exception: pass
                            wb_pdf = None
                        time.sleep(2)

                if google_service and pdf_terbentuk:
                    thread_id, msg_mime_id, reply_subject = find_email_thread(google_service, no_invoice)
                    if not thread_id: thread_id, msg_mime_id, reply_subject = find_email_thread(google_service, no_batch)

                    str_ppn = f"{int(grand_ppn):,}".replace(",", ".")
                    body_email = (
                        f"Dear {CLIENT_TEAM_NAME}\n\n"
                        f"Based on the verification result attached in the previous email, we have reviewed and "
                        f"recalculated the tax (PPN) value on the total payable amount. The recalculated PPN is Rp{str_ppn},- "
                        f"(file attached). Please confirm this value. Thank you for your attention to this matter.\n\n"
                        f"Best regards,\n{SENDER_SIGNATURE}"
                    )

                    if thread_id:
                        print("   -> SENDING NORMAL EMAIL (plain-text reply)...")
                        send_gmail_message(google_service, CLIENT_CONTACT_EMAIL, reply_subject, body_email, pdf_output, thread_id, msg_mime_id)
                    else:
                        print("   -> SENDING NEW NORMAL EMAIL...")
                        new_subject = f"Recalculated PPN Value - Invoice: {no_invoice}"
                        send_gmail_message(google_service, CLIENT_CONTACT_EMAIL, new_subject, body_email, pdf_output)
                elif google_service and not pdf_terbentuk:
                    print(f"   -> WARNING: PDF export failed. Original Windows error: {error_pdf_detail}")

            archive_path = os.path.join(archive_folder, file_name)
            if os.path.exists(archive_path):
                base, ext = os.path.splitext(file_name)
                archive_path = os.path.join(archive_folder, f"{base}_{int(time.time())}{ext}")
            shutil.move(file_path, archive_path)

        except Exception as e:
            print(f"   -> FAILED ON FILE: {e}")

    if driver:
        print("   [BOT] Shutting down headless Chrome engine...")
        driver.quit()

    if resume_list:
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        existing = glob.glob(os.path.join(recap_folder, f"Rekap_Harian_{today}_Ke-*.xlsx"))
        resume_name = f"Rekap_Harian_{today}_Ke-{len(existing) + 1}.xlsx"
        resume_path = os.path.join(recap_folder, resume_name)

        res_df = pd.DataFrame(resume_list)
        res_df.index = res_df.index + 1
        res_df.index.name = 'No'
        res_df.reset_index(inplace=True)
        res_df.to_excel(resume_path, index=False)

        wb_res = load_workbook(resume_path)
        ws_res = wb_res.active
        for row in ws_res.iter_rows(min_row=1, max_row=ws_res.max_row, min_col=1, max_col=ws_res.max_column):
            for cell in row:
                thin_s = Side(border_style="thin", color="000000")
                cell.border = Border(left=thin_s, right=thin_s, top=thin_s, bottom=thin_s)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if cell.row == 1: cell.font = Font(bold=True)
                if isinstance(cell.value, (int, float)) and cell.column in [4, 5, 6, 7]: cell.number_format = '#,##0'

        c_widths_res = {1: 5, 2: 20, 3: 25, 4: 18, 5: 18, 6: 18, 7: 18, 8: 30, 9: 35}
        for c_idx, w in c_widths_res.items(): ws_res.column_dimensions[ws_res.cell(row=1, column=c_idx).column_letter].width = w
        wb_res.save(resume_path)
        print(f"\nDAILY RECAP CREATED: {resume_name}")

    if sheets_batch_data and sheets_service and SPREADSHEET_ID:
        print("\nSending tax-invoice requests to Google Sheets tracker...")
        try:
            response = sheets_service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID, range=f'{SHEET_NAME}!I:I'
            ).execute()

            kolom_i_data = response.get('values', [])
            next_row = len(kolom_i_data) + 1

            col_d_data = [[row[0]] for row in sheets_batch_data]
            col_i_data = [[row[1]] for row in sheets_batch_data]
            col_mn_data = [[row[2], row[3]] for row in sheets_batch_data]

            data = [
                {'range': f'{SHEET_NAME}!D{next_row}', 'values': col_d_data},
                {'range': f'{SHEET_NAME}!I{next_row}', 'values': col_i_data},
                {'range': f'{SHEET_NAME}!M{next_row}', 'values': col_mn_data}
            ]

            body = {
                'valueInputOption': 'USER_ENTERED',
                'data': data
            }

            sheets_service.spreadsheets().values().batchUpdate(
                spreadsheetId=SPREADSHEET_ID, body=body
            ).execute()

            print(f"SUCCESS! {len(sheets_batch_data)} row(s) appended without overwriting existing formulas.")
        except Exception as e:
            print(f"FAILED to write to Google Sheets: {e}")

    print("\nAll processing complete! Disputed VVIP invoices are archived in '7_Arsip_Tagihan_Dispute'.")

if __name__ == "__main__":
    process_batch_credit_notes()
