from flask import Flask, request, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)  

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    
    df = pd.read_excel(file, header=4)

    cols_to_remove = df.columns[[2, 3, 4, 5, 7, 8]]
    df = df.drop(columns=cols_to_remove, errors='ignore')

    df["benar"] = ""
    df["real value"] = ""

    # Use memory instead of saving file
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        download_name="output.xlsx",
        as_attachment=True
    )


if __name__ == "__main__":
    app.run()