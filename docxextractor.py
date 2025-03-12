from flask import Flask, request, send_file, render_template
import zipfile
import os
import shutil
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
EXTRACT_FOLDER = "extracted"
RESULT_FOLDER = "results"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['EXTRACT_FOLDER'] = EXTRACT_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

# Zorg ervoor dat de folders bestaan
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXTRACT_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return "Geen bestand geüpload", 400
        file = request.files['file']
        if file.filename == '':
            return "Geen bestand geselecteerd", 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Extract images from Word file
        extracted_images_path = extract_images(filepath)
        
        if not extracted_images_path:
            return "Geen afbeeldingen gevonden", 400
        
        # Maak een ZIP van de afbeeldingen
        zip_filename = filename.replace(".docx", "_afbeeldingen.zip")
        zip_filepath = os.path.join(app.config['RESULT_FOLDER'], zip_filename)
        with zipfile.ZipFile(zip_filepath, 'w') as img_zip:
            for root, _, files in os.walk(extracted_images_path):
                for img in files:
                    img_path = os.path.join(root, img)
                    img_zip.write(img_path, os.path.basename(img))
        
        # Verwijder tijdelijke bestanden
        shutil.rmtree(extracted_images_path)
        os.remove(filepath)
        
        return send_file(zip_filepath, as_attachment=True)
    
    return render_template('upload.html')

def extract_images(docx_path):
    temp_extract_dir = os.path.join(app.config['EXTRACT_FOLDER'], os.path.basename(docx_path).replace(".docx", ""))
    os.makedirs(temp_extract_dir, exist_ok=True)
    
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(temp_extract_dir)
    
    media_path = os.path.join(temp_extract_dir, "word/media")
    if not os.path.exists(media_path):
        return None
    
    result_dir = os.path.join(app.config['RESULT_FOLDER'], os.path.basename(docx_path).replace(".docx", ""))
    os.makedirs(result_dir, exist_ok=True)
    
    for img in os.listdir(media_path):
        shutil.move(os.path.join(media_path, img), os.path.join(result_dir, img))
    
    return result_dir

if __name__ == '__main__':
    app.run(debug=True)
