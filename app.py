from flask import Flask, request, send_file
import pandas as pd
from io import BytesIO
import os

app = Flask(__name__)

# ✅ Homepage (fixes "Not Found")
@app.route('/')
def home():
    return '''
    <h2>Excel Converter</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="file" required>
        <button type="submit">Upload</button>
    </form>
    '''

# ✅ Upload + processing route
@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    
    # Use row 5 as header
    df = pd.read_excel(file, header=4)

    # Remove columns
    df = df.drop(columns=["   Department Name   ", "   Convertion Name   ","   Scale 1   ","   Operator   ","   Scale 2   ","   Low Stock   "], errors='ignore')

    # Add new columns
    df["benar"] = ""
    df["real value"] = ""

    # Save to memory (important for deployment)
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        download_name="output.xlsx",
        as_attachment=True
    )

# ✅ Required for Render (port handling)
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))