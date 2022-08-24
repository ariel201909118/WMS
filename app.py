import pandas as pd
from flask import Flask, render_template, url_for, request, redirect, send_file
from read_file import read
import zipfile 

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    reference = ''
    
    if request.method == 'POST':
        file = request.files.get('file')
        container_no = request.form.get('container_no').upper()
        reference_no = request.form.get('reference_no')
        reference = read(file, container_no, reference_no)
        return render_template('form.html', reference=reference, doc='', doc1='', doc2='')
    return render_template('form.html', reference='')
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')