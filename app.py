from flask import Flask, request, send_file
import pandas as pd
from io import BytesIO
import os
from openpyxl.utils import get_column_letter

app = Flask(__name__)

#  Homepage 
@app.route('/')
def home():
    return '''
    <h2>Excel Converter</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="file" required>
        <button type="submit">Upload</button>
    </form>
    '''

# Upload + processing route
@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        file = request.files['file']

        # Row 5 is header
        df = pd.read_excel(file, header=4)

        df.columns = (
            df.columns
            .astype(str)
            .str.replace("\n", " ")
            .str.strip()
        )


        #  Drop completely empty columns 
        df = df.dropna(axis=1, how='all')

        # Debug safety (prevents crash if structure changes)
        print(df.columns)

        # Remove columns by NAME (NOT index)
        # Based on your sample file:
        remove_cols = [
            "Scale 1",
            "Department Name",
            "Convertion Name",
            "Operator",
            "Low Stock"
        ]


        df = df.drop(columns=remove_cols, errors='ignore')



        # Add new columns
        df["benar"] = ""
        df["real value"] = ""

        # fit output
        output = BytesIO()

        # create excel writer
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')

            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # auto-fit columns
            for i, col in enumerate(df.columns):
                max_len = max(
                    df[col].astype(str).map(len).max(),
                    len(str(col))
                ) + 2

                worksheet.set_column(i, i, min(max_len, 50))

        output.seek(0)

        return send_file(
            output,
            download_name="output.xlsx",
            as_attachment=True
        )
    
    except Exception as e:
        import traceback
        print("ERROR OCCURRED:")
        traceback.print_exc()
        return str(e), 500





# Required for Render (port handling)
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))





