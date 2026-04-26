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
        df.to_excel(output, index=False)

        # reopen workbook to adjust column width
        output.seek(0)
        from openpyxl import load_workbook

        wb = load_workbook(output)
        ws = wb.active

        # auto-fit columns
        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)

            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass

            ws.column_dimensions[col_letter].width = max_length + 2

        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)


        return send_file(final_output, download_name="output.xlsx", as_attachment=True)
    
    except Exception as e:
        import traceback
        print("ERROR OCCURRED:")
        traceback.print_exc()
        return str(e), 500





# Required for Render (port handling)
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))





