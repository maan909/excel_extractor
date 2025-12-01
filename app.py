from flask import Flask, render_template, request, send_file, make_response
import os, tempfile
from extracter import extract_xls_data

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")


@app.route("/upload-folder", methods=["POST"])
def upload_folder():
    uploaded_files = request.files.getlist("folder")
    token = request.form.get("downloadToken")

    if not uploaded_files:
        return "No folder uploaded"

    # Create a temp directory to store XLS files
    temp_dir = tempfile.mkdtemp()

    # Save only .xls files
    for file in uploaded_files:
        if file.filename.lower().endswith(".xls"):
            save_path = os.path.join(temp_dir, os.path.basename(file.filename))
            file.save(save_path)

    # Run extractor
    output_path = os.path.join(temp_dir, "extracted_output.xlsx")
    extract_xls_data(temp_dir, output_path)

    response = make_response(send_file(output_path, as_attachment=True))
    response.set_cookie("downloadToken", token)

    return response


if __name__ == "__main__":
    app.run(debug=True)
