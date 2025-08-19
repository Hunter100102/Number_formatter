from flask import Flask, request, render_template_string, send_from_directory
from werkzeug.utils import secure_filename
import pandas as pd
import os
import re
from collections import OrderedDict

app = Flask(__name__, static_url_path='', static_folder='.')

app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Serve favicon from project root (ensure ./favicon.ico exists)
@app.route('/favicon.ico')
def favicon():
    return send_from_directory(app.static_folder, 'favicon.ico', mimetype='image/vnd.microsoft.icon')

HTML_TEMPLATE = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Part Number Formatter</title>
  <link rel="icon" href="/favicon.ico" type="image/x-icon">
  <style>
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; padding: 24px; }
    form { margin-bottom: 16px; }
    textarea { width: 100%; height: 300px; }
    .msg { color: #555; margin: 8px 0; }
  </style>
</head>
<body>
  <h1>Part Number Formatter</h1>

  <form method="POST" enctype="multipart/form-data">
    <input type="file" name="file" accept=".xlsx,.xls,.csv" required>
    <button type="submit">Upload & Format</button>
  </form>

  {% if message %}
    <div class="msg">{{ message }}</div>
  {% endif %}

  {% if processed %}
    <h2>Processed Results</h2>
    <textarea readonly>{{ processed }}</textarea>
  {% endif %}

  <footer style="margin-top:40px; font-size:0.9em; color:#666;">
    <hr>
    <p style="text-align:center;">Powered by AutomateIT</p>
  </footer>
</body>
</html>
"""

def process_part_number(part: str) -> str:
    cleaned = re.sub(r'[^A-Za-z0-9]', '', part or '').upper()
    if not cleaned:
        return ''
    if not cleaned.startswith('A'):
        return cleaned

    digits = cleaned[1:]
    if len(digits) == 10:
        return f"A{digits}*"
    elif len(digits) == 12:
        return f"A{digits[:10]}*"
    elif len(digits) == 14:
        base = digits[:10]
        suffix = digits[-4:]
        return f"A{base}**{suffix}"
    elif len(digits) == 16:
        base = digits[:10]
        suffix = digits[12:]
        return f"A{base}**{suffix}"
    elif len(digits) >= 17:
        return f"A{digits}"
    else:
        return cleaned

def read_table(filepath: str) -> pd.DataFrame:
    _, ext = os.path.splitext(filepath.lower())
    if ext in ('.xlsx', '.xls'):
        return pd.read_excel(filepath)
    elif ext == '.csv':
        return pd.read_csv(filepath)
    # fallback: try Excel then CSV
    try:
        return pd.read_excel(filepath)
    except Exception:
        return pd.read_csv(filepath)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    processed = None
    message = None

    if request.method == 'POST':
        if 'file' not in request.files:
            message = "No file part in request."
            return render_template_string(HTML_TEMPLATE, processed=processed, message=message)

        file = request.files['file']
        if not file or file.filename.strip() == '':
            message = "No file selected."
            return render_template_string(HTML_TEMPLATE, processed=processed, message=message)

        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        try:
            df = read_table(filepath)
        except Exception as e:
            message = f"Error reading file: {e}"
            return render_template_string(HTML_TEMPLATE, processed=None, message=message)

        # Try to find the 'Part Number' column case-insensitively
        col = None
        for c in df.columns:
            if str(c).strip().lower() == 'part number':
                col = c
                break

        if not col:
            message = "Column 'Part Number' not found."
            return render_template_string(HTML_TEMPLATE, processed=None, message=message)

        series = df[col].dropna().astype(str)

        # Deduplicate while preserving order
        out_ordered = OrderedDict()
        for p in series:
            value = process_part_number(p)
            if value:  # skip empty results
                out_ordered.setdefault(value, None)

        processed = "\n".join(out_ordered.keys())
        if not processed:
            message = "No valid part numbers found."

    return render_template_string(HTML_TEMPLATE, processed=processed, message=message)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
