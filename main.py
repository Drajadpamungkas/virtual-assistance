import speech_recognition as srec
from gtts import gTTS
import os
import subprocess
import urllib.parse

def perintah():
    mendengar = srec.Recognizer()
    with srec.Microphone() as source:
        print('Mendengarkan....')
        suara = mendengar.listen(source, phrase_time_limit=5)
        try:
            print('Diterima...')
            dengar = mendengar.recognize_google(suara, language='id-ID')
            print(dengar)
        except Exception as e:
            print(f"Error: {e}")
            dengar = ""
        return dengar

def ngomong(teks):
    bahasa = 'id'
    namafile = 'audio.mp3'
    def reading():
        suara = gTTS(text=teks, lang=bahasa, slow=False)
        suara.save(namafile)
        os.system(f'start {namafile}')
    reading()

def jalankan_perintah(perintah):
    if "buka file manager" in perintah.lower():
        subprocess.Popen('explorer')
        ngomong("Membuka file manager")

    elif "buka notepad" in perintah.lower():
        subprocess.Popen('notepad')
        ngomong("Membuka Notepad")

    elif "buka google chrome" in perintah.lower():
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        subprocess.Popen(chrome_path)
        ngomong("Membuka Google Chrome")

    elif "buka microsoft word" in perintah.lower():
        word_path = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
        subprocess.Popen(word_path)
        ngomong("Membuka Microsoft Word")

    elif "buka microsoft excel" in perintah.lower():
        excel_path = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
        subprocess.Popen(excel_path)
        ngomong("Membuka Microsoft Excel")

    elif "buka microsoft powerpoint" in perintah.lower():
        powerpoint_path = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
        subprocess.Popen(powerpoint_path)
        ngomong("Membuka Microsoft PowerPoint")

    elif "buka microsoft edge" in perintah.lower():
        edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        subprocess.Popen(edge_path)
        ngomong("Membuka Microsoft Edge")

    elif "buka kalkulator" in perintah.lower():
        subprocess.Popen('calc')
        ngomong("Membuka Kalkulator")



    elif "cari di youtube" in perintah.lower():
        cari_keyword = perintah.lower().replace("cari di youtube", "").strip()
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        query = urllib.parse.quote(cari_keyword)
        url = f"https://www.youtube.com/results?search_query={query}"
        subprocess.Popen([chrome_path, url])
        ngomong(f"Mencari {cari_keyword} di YouTube menggunakan Google Chrome")

    elif "cari di chrome" in perintah.lower():
        cari_keyword = perintah.lower().replace("cari di chrome", "").strip()
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        query = urllib.parse.quote(cari_keyword)
        url = f"https://www.google.com/search?q={query}"
        subprocess.Popen([chrome_path, url])
        ngomong(f"Mencari {cari_keyword} di Google Chrome")

    else:
        ngomong("Perintah tidak dikenali")

def run_michelle():
    layanan = perintah()
    if layanan:
        jalankan_perintah(layanan)

run_michelle()
