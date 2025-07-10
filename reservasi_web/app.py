from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)
excel_file = 'data.xlsx'

if not os.path.exists(excel_file):
    df = pd.DataFrame(columns=["Nama Pengunjung", "No HP", "Nama Perusahaan", "Waktu"])
    df.to_excel(excel_file, index=False)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        nama = request.form.get('nama', '')
        nohp = request.form.get('nohp', '')
        perusahaan = request.form.get('perusahaan', '')
        waktu = request.form.get('waktu', '')

        new_data = {
            "Nama Pengunjung": nama,
            "No HP": nohp,
            "Nama Perusahaan": perusahaan,
            "Waktu": waktu
        }

        df = pd.read_excel(excel_file)
        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
        df.to_excel(excel_file, index=False)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)