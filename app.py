from flask import Flask, render_template, request, send_file, make_response
import zipfile, os, tempfile
from extracter import extract_xls_data

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_zip():
    file = request.files["zipfile"]
    token = request.form.get("downloadToken")

    if not file:
        return "No file uploaded"

    temp_dir = tempfile.mkdtemp()
    zip_path = os.path.join(temp_dir, "input.zip")
    file.save(zip_path)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(temp_dir)

    output_path = os.path.join(temp_dir, "extracted_output.xlsx")
    extract_xls_data(temp_dir, output_path)

    response = make_response(send_file(output_path, as_attachment=True))
    response.set_cookie("downloadToken", token)

    return response

if __name__ == "__main__":
    app.run(debug=True)
