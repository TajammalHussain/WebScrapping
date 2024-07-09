from flask import Flask, request, redirect, url_for, render_template_string
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
df = pd.DataFrame()  # Global DataFrame to store the uploaded data
# the app is just working with some Jquery scripts and then provide some results, after uploading the excel files. 
# the formula to concatinate the values while searching the product
@app.route('/')
def index():
    return render_template_string('''
    <!doctype html>
    <title>Upload Excel File</title>
    <h1>Upload Excel File for Processing</h1>
    <form action="/upload" method="post" enctype="multipart/form-data">
      <input type="file" name="file">
      <input type="submit" value="Upload">
    </form>
    ''')

@app.route('/upload', methods=['POST'])
def upload_file():
    global df
    if 'file' not in request.files:
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    
    if file:
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        df = pd.read_excel(filepath)
        return redirect(url_for('display_file'))

@app.route('/display')
def display_file():
    global df
    if df.empty:
        return redirect(url_for('index'))
    
    table_html = df.to_html(classes='table table-striped', index=False)
    return render_template_string('''
    <!doctype html>
    <title>File Contents</title>
    <h1>File Contents</h1>
    <form>
      <input type="text" id="search" placeholder="Search across all columns">
    </form>
    <a href="{{ url_for('index') }}">Upload Another File</a>
    <div id="table">{{ table_html | safe }}</div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script>
      $(document).ready(function() {
        $('#search').on('input', function() {
          var query = $(this).val();
          $.get('/search', { query: query }, function(data) {
            $('#table').html(data);
          });
        });
      });
    </script>
    ''', table_html=table_html)

@app.route('/search', methods=['GET'])
def search():
    global df
    if df.empty:
        return ''
    
    query = request.args.get('query', '').strip().lower()
    if not query:
        return df.to_html(classes='table table-striped', index=False)
    
    filtered_df = df[df.apply(lambda row: row.astype(str).str.lower().str.contains(query).any(), axis=1)]
    table_html = filtered_df.to_html(classes='table table-striped', index=False)
    return table_html

if __name__ == '__main__':
    app.run(debug=True)
