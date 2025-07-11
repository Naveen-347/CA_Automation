import os
import time
import re
import pandas as pd
import requests
from threading import Thread, Lock
from concurrent.futures import ThreadPoolExecutor
from flask import Flask, render_template, request, jsonify, send_file
from bs4 import BeautifulSoup

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

progress_data = {
    "total": 0,
    "current": 0,
    "current_name": "",
    "output_file": ""
}

lock = Lock()

def generate_company_url(company_name, cin):
    name_slug = company_name.strip().upper().replace(" ", "-")
    cin = cin.strip().upper()
    return f"https://mcamasterdata.com/company/{name_slug}/{cin}"

def fetch_company_details(company_name, cin):
    url = generate_company_url(company_name, cin)
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        if resp.status_code != 200:
            return dict(url=url, email="Error", activity="Error", pan="Error", gst="Error")
        soup = BeautifulSoup(resp.text, 'html.parser')
        email = activity = pan = gst = "Not mentioned"
        for dt in soup.find_all("dt"):
            lbl = dt.get_text(strip=True).lower()
            dd = dt.find_next_sibling("dd")
            val = dd.get_text(strip=True) if dd else "Not mentioned"
            if "e-mail" in lbl:     email = val
            elif "activity" in lbl: activity = val
            elif "pan" in lbl:      pan = val
            elif "gst" in lbl:      gst = val
        if gst == "Not mentioned":
            m = re.search(r"\b\d{2}[A-Z]{5}\d{4}[A-Z]\wZ\w\b", soup.get_text())
            if m: gst = m.group()
        return dict(url=url, email=email, activity=activity, pan=pan, gst=gst)
    except:
        return dict(url=url, email="Exception", activity="Exception", pan="Exception", gst="Exception")

def background_process(input_path, output_path):
    df = pd.read_excel(input_path)
    progress_data.update(total=len(df), current=0, current_name="", output_file="")
    results = [None] * len(df)

    def process_row(i, row):
        name = str(row['Company Name']).strip()
        cin = str(row['CIN']).strip()
        details = fetch_company_details(name, cin)
        results[i] = {
            "Company Name": name,
            "CIN": cin,
            "URL": details['url'],
            "Email": details['email'],
            "Activity": details['activity'],
            "PAN": details['pan'],
            "GST": details['gst']
        }
        with lock:
            progress_data["current"] += 1
            progress_data["current_name"] = name

    with ThreadPoolExecutor(max_workers=5) as executor:
        for i, row in df.iterrows():
            executor.submit(process_row, i, row)
        executor.shutdown(wait=True)

    pd.DataFrame(results).to_excel(output_path, index=False)
    progress_data.update(current_name="done", output_file=output_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not (f and f.filename.lower().endswith('.xlsx')):
        return jsonify({"error": "Only .xlsx allowed"}), 400

    inp = os.path.join(UPLOAD_FOLDER, f.filename)
    f.save(inp)
    out_fname = f"output_{int(time.time())}.xlsx"
    outp = os.path.join(OUTPUT_FOLDER, out_fname)

    progress_data.update(total=0, current=0, current_name="", output_file="")

    Thread(target=background_process, args=(inp, outp), daemon=True).start()
    return jsonify({"status": "started"})

@app.route('/progress')
def progress():
    return jsonify(progress_data)

@app.route('/download')
def download():
    path = progress_data.get("output_file")
    if path and os.path.exists(path):
        return send_file(path, as_attachment=True)
    return "", 404

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000, debug=True)

