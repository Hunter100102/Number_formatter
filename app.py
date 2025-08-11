from flask import Flask, request, render_template_string
import pandas as pd
import os
import re

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

HTML_TEMPLATE = """
<!doctype html>
<title>Part Number Formatter</title>
<h2>Upload a .xlsx file with Part Numbers</h2>
<form method=post enctype=multipart/form-data>
  <input type=file name=file>
  <input type=submit value=Upload>
</form>
{% if processed %}
<h3>Formatted Part Numbers (no spaces):</h3>
<textarea rows="20" cols="80">{{ processed }}</textarea>
{% endif %}
"""




def process_part_number(part):
    cleaned = re.sub(r'[^A-Za-z0-9]', '', part).upper()

    if not cleaned.startswith('A'):
        return cleaned

    digits = cleaned[1:]

    if len(digits) == 10:
        return f"A{digits}*"
    elif len(digits) == 12:
        return f"A{digits[:10]}*"
    elif len(digits) >= 13:
        # Replace 11th and 12th characters with '**', preserve the rest
        return f"A{digits[:10]}**{digits[12:]}"
    else:
        return cleaned



@app.route('/', methods=['GET', 'POST'])
def upload_file():
    processed = None
    if request.method == 'POST':
        file = request.files['file']
        if file:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            try:
                df = pd.read_excel(filepath, engine='openpyxl')
            except Exception as e:
                return render_template_string(HTML_TEMPLATE, processed=f"Error reading file: {e}")
            if 'Part Number' not in df.columns:
                return render_template_string(HTML_TEMPLATE, processed="Column 'Part Number' not found.")
            processed_parts = [process_part_number(str(p)) for p in df['Part Number']]
            processed = '\n'.join(processed_parts)
    return render_template_string(HTML_TEMPLATE, processed=processed)

if __name__ == '__main__':
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
