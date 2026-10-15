import os
import pickle
import json
import time
import pandas as pd
from datetime import datetime
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google import genai
from dotenv import load_dotenv

# Memuat variabel rahasia dari file .env
load_dotenv()

# ==========================================
# KONFIGURASI
# ==========================================
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY") 
client = genai.Client(api_key=GEMINI_API_KEY)
EXCEL_FILE = "Hasil_Klasifikasi_Lamaran.xlsx"

# ==========================================
# FUNGSI-FUNGSI
# ==========================================
def authenticate_gmail():
    creds = None
    if os.path.exists('token.pickle'):
        try:
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        except Exception:
            creds = None
            
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Token lama kedaluwarsa. Membuka browser untuk login ulang...")
                creds = None

        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('gmail', 'v1', credentials=creds)

def get_recent_emails(service, max_results=200):
    request = service.users().messages().list(userId='me', labelIds=['INBOX'], maxResults=max_results)
    results = request.execute()
    messages = results.get('messages', [])
    
    emails = []
    if not messages:
        return emails
        
    total_msg = len(messages)
    for i, message in enumerate(messages, 1):
        if i % 10 == 0:
            print(f"Mendownload email {i}/{total_msg}...")
        try:
            msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
            time.sleep(0.1) # Jeda aman untuk Gmail API
        except Exception as e:
            time.sleep(1)
            continue
        
        payload = msg.get('payload', {})
        headers = payload.get('headers', [])
        
        subject = ''
        sender = ''
        date_str = ''
        
        for header in headers:
            if header['name'] == 'Subject':
                subject = header['value']
            elif header['name'] == 'From':
                sender = header['value']
            elif header['name'] == 'Date':
                date_str = header['value']
                
        # Mengekstrak isi email secara lengkap (Body)
        import base64
        body = ""
        def get_body(payload_part):
            b = ""
            if 'parts' in payload_part:
                for part in payload_part['parts']:
                    b += get_body(part)
            else:
                data = payload_part.get('body', {}).get('data')
                if data:
                    try:
                        b = base64.urlsafe_b64decode(data).decode('utf-8', 'ignore')
                    except:
                        pass
            return b
            
        body = get_body(payload)
        snippet = msg.get('snippet', '')
        
        # Batasi body agar tidak terlalu panjang (misal max 3000 karakter)
        if len(body) > 3000:
            body = body[:3000] + "..."
            
        if not body:
            body = snippet
        
        emails.append({
            'id': message['id'],
            'date': date_str,
            'sender': sender,
            'subject': subject,
            'body': body,
            'snippet': snippet
        })
        
    return emails

def classify_email_with_ai(email_data):
    prompt = f"""
    Anda adalah sistem ekstraksi informasi lamaran kerja.
    Baca informasi email berikut:
    Pengirim: {email_data['sender']}
    Subjek: {email_data['subject']}
    Cuplikan: {email_data['snippet']}
    Isi Pesan: {email_data['body']}
    
    Tugas Anda adalah mengekstrak data dari email tersebut.
    Tentukan apakah email ini terkait dengan proses lamaran kerja.
    Jika ya, tebak nama perusahaan dan posisi (role) yang dilamar.
    Tentukan juga status terbarunya:
    - DITERIMA (offering letter, diterima kerja)
    - ON PROGRESS (panggilan tes, interview, lamaran diterima/sedang diproses)
    - DITOLAK (tidak lolos, posisi sudah terisi)
    - LAINNYA (newsletter, job alert, promosi, dll)
    
    PENTING: Balas HANYA dengan format JSON valid tanpa awalan atau akhiran apapun (tanpa markdown ```json).
    Contoh output yang diharapkan:
    {{
        "is_lamaran": true,
        "perusahaan": "Tokopedia",
        "posisi": "Data Analyst",
        "status": "ON PROGRESS",
        "ringkasan": "Diundang untuk interview tahap 1"
    }}
    """
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model='gemini-flash-lite-latest',
                contents=prompt
            )
            text = response.text.strip()
            if text.startswith("```json"):
                text = text[7:-3]
            elif text.startswith("```"):
                text = text[3:-3]
            
            data = json.loads(text.strip())
            time.sleep(4)
            return data
        except Exception as e:
            if attempt < 2:
                time.sleep(15)
            else:
                time.sleep(4)
                return {"is_lamaran": False}

def load_existing_data():
    lamaran_dict = {}
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name='Tracking')
            # Memasukkan data lama ke dalam dictionary
            for _, row in df.iterrows():
                perusahaan = str(row.get('Perusahaan', ''))
                posisi = str(row.get('Posisi', ''))
                if perusahaan and posisi:
                    key = f"{perusahaan.upper()} - {posisi.upper()}"
                    ringkasan_val = row.get('Ringkasan', '')
                    if pd.isna(ringkasan_val) or ringkasan_val == 'nan':
                        ringkasan_val = row.get('Ringkasan AI', '')
                    if pd.isna(ringkasan_val) or ringkasan_val == 'nan':
                        ringkasan_val = ''
                        
                    lamaran_dict[key] = {
                        "Tanggal Update": row.get('Tanggal Update', ''),
                        "Perusahaan": perusahaan,
                        "Posisi": posisi,
                        "Status": row.get('Status', ''),
                        "Ringkasan": ringkasan_val,
                        "Tindakan Lanjutan": row.get('Tindakan Lanjutan', '')
                    }
        except Exception as e:
            print(f"Gagal memuat Excel lama: {e}")
    return lamaran_dict

def main():
    print("Mendapatkan layanan Gmail API...")
    try:
        service = authenticate_gmail()
    except Exception as e:
        print(f"Error saat autentikasi Gmail: {e}")
        return

    print("Mengambil 30 email terbaru untuk pengecekan harian...")
    emails = get_recent_emails(service, max_results=30) 
    
    # Membalik urutan email agar diproses dari yang TERLAMA ke TERBARU
    # Ini memastikan email terbaru yang menimpa status lamaran
    emails.reverse()
    
    print("Memuat database Excel lama (jika ada)...")
    lamaran_dict = load_existing_data()

    print("Mulai menganalisis email...")
    for idx, email in enumerate(emails, start=1):
        safe_subject = email['subject'].encode('ascii', 'ignore').decode('ascii')
        data = classify_email_with_ai(email)
        
        if not data.get("is_lamaran") or data.get("status") in ["LAINNYA", None, ""]:
            print(f"[-] Email diabaikan (Bukan lamaran): {safe_subject}")
            continue

        perusahaan = data.get("perusahaan") or "Tidak Diketahui"
        posisi = data.get("posisi") or "Tidak Diketahui"
        status = data.get("status") or "ON PROGRESS"
        ringkasan = data.get("ringkasan") or email['snippet']
        
        try:
            parsed_date = email['date'][:16] 
        except:
            parsed_date = email['date']

        key = f"{perusahaan.upper()} - {posisi.upper()}"
        
        # Logika menimpa data lama
        if key in lamaran_dict:
            status_lama = lamaran_dict[key]["Status"]
            if status_lama != status:
                print(f"[*] UPDATE STATUS: {perusahaan} ({posisi}) | {status_lama} -> {status}")
            else:
                print(f"[+] Info terbaru: {perusahaan} ({posisi}) | Tetap {status}")
        else:
            print(f"[+] LAMARAN BARU: {perusahaan} ({posisi}) | {status}")
            
        tindakan = "Follow-up" if status == "ON PROGRESS" else ("Review kontrak" if status == "DITERIMA" else "Arsipkan")

        # Update atau Buat entri baru
        lamaran_dict[key] = {
            "Tanggal Update": parsed_date,
            "Perusahaan": perusahaan,
            "Posisi": posisi,
            "Status": status,
            "Ringkasan": ringkasan,
            "Tindakan Lanjutan": tindakan
        }

    # Menyimpan kembali ke Excel (Replace sheet sepenuhnya)
    if lamaran_dict:
        df = pd.DataFrame(list(lamaran_dict.values()))
        print(f"Menyimpan {len(df)} total riwayat lamaran ke Excel...")
        
        try:
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Tracking', index=False)
            print("Berhasil!")
        except Exception as e:
            df.to_excel("Backup_Lamaran_Baru.xlsx", index=False)
            print(f"File utama terkunci, disimpan ke Backup_Lamaran_Baru.xlsx")
    else:
        print("Belum ada data lamaran kerja untuk disimpan.")

if __name__ == "__main__":
    main()
