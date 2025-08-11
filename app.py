from flask import Flask, request, render_template_string
import pandas as pd
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

HTML_TEMPLATE = """
<!doctype html>
<title>Part Number Processor</title>
<h2>Upload a .xlsx or .csv file with Part Numbers</h2>
<form method=post enctype=multipart/form-data>
  <input type=file name=file>
  <input type=submit value=Upload>
</form>
{% if processed %}
<h3>Processed Part Numbers:</h3>
<textarea rows="20" cols="80">{{ processed }}</textarea>
{% endif %}
"""

def process_part_number(part):
    digits = ''.join(filter(str.isdigit, part))
    prefix = 'A'
    if len(digits) == 10:
        return f"{prefix} {digits[:3]} {digits[3:6]} {digits[6:8]} {digits[8:]} *"
    elif len(digits) == 12:
        return f"{prefix} {digits[:3]} {digits[3:6]} {digits[6:8]} {digits[8:10]} *"
    elif len(digits) == 14:
        return f"{prefix} {digits[:3]} {digits[3:6]} {digits[6:8]} {digits[8:10]} ** {digits[12:]}"
    elif len(digits) == 16:
        return f"{prefix} {digits[:3]} {digits[3:6]} {digits[6:8]} {digits[8:10]} ** {digits[14:]}"
    else:
        return part

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    processed = None
    if request.method == 'POST':
        file = request.files['file']
        if file:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            if file.filename.endswith('.xlsx'):
                df = pd.read_excel(filepath, engine='openpyxl')
            elif file.filename.endswith('.csv'):
                df = pd.read_csv(filepath)
            else:
                return render_template_string(HTML_TEMPLATE, processed="Unsupported file format.")
            if 'Part Number' not in df.columns:
                return render_template_string(HTML_TEMPLATE, processed="Column 'Part Number' not found.")
            processed_parts = [process_part_number(str(p)) for p in df['Part Number']]
            processed = '\n'.join(processed_parts)
    return render_template_string(HTML_TEMPLATE, processed=processed)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

