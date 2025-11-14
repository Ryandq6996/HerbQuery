from flask import Flask, render_template, request, send_file
import pandas as pd
import io
import requests
from bs4 import BeautifulSoup

app = Flask(__name__)

def query_medicine(name):
    url = f"https://www.yibian.com.tw/search?keyword={name}"
    res = requests.get(url)
    soup = BeautifulSoup(res.text, "html.parser")
    data = {}
    data['名稱'] = name
    data['類型'] = soup.select_one(".type").text.strip() if soup.select_one(".type") else ""
    data['功效'] = soup.select_one(".effect").text.strip() if soup.select_one(".effect") else ""
    data['主治'] = soup.select_one(".indication").text.strip() if soup.select_one(".indication") else ""
    data['組成'] = soup.select_one(".composition").text.strip() if soup.select_one(".composition") else ""
    data['禁忌'] = soup.select_one(".taboo").text.strip() if soup.select_one(".taboo") else ""
    data['主調臟腑'] = soup.select_one(".organ").text.strip() if soup.select_one(".organ") else ""
    data['來源摘要'] = soup.select_one(".summary").text.strip() if soup.select_one(".summary") else ""
    return data

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/search', methods=['POST'])
def search():
    name = request.form['name']
    result = query_medicine(name)
    return render_template('index.html', result=[result])

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    file = request.files['file']
    df = pd.read_excel(file)
    results = []
    for name in df.iloc[:,0]:
        results.append(query_medicine(name))
    df_result = pd.DataFrame(results)
    output = io.BytesIO()
    df_result.to_excel(output, index=False)
    output.seek(0)
    return send_file(output, download_name="查詢結果.xlsx", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
