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

    # ✅ Row 5 is header
    df = pd.read_excel(file, header=4)
    
    df.columns = (
        df.columns
        .astype(str)
        .str.replace("\n", " ")
        .str.strip()
    )

    print("🔥 RAW COLUMNS:", list(df.columns))

    # ✅ Drop completely empty columns (important for messy Excel)
    df = df.dropna(axis=1, how='all')

    # Debug safety (prevents crash if structure changes)
    print(df.columns)

    # Remove columns by NAME (NOT index)
    # Based on your sample file:
    remove_cols = [
        "Scale 1",
        "Scale 2",
        "Department Name",
        "Convertion Name",
        "Operator",
        "Low Stock"
    ]

    print("🧨 TRYING TO REMOVE:", remove_cols)

    df = df.drop(columns=remove_cols, errors='ignore')



    # Add new columns
    df["benar"] = ""
    df["real value"] = ""


    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)


    return send_file(output, download_name="output.xlsx", as_attachment=True)




# Required for Render (port handling)
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))





