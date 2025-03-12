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
        extracted_images_path, warning = extract_images(filepath)
        
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
            
            # Voeg een waarschuwing toe als er .emf of .tmp bestanden waren
            if warning:
                warning_path = os.path.join(app.config['RESULT_FOLDER'], "!LET OP! Niet alle figuren zijn succesvol geëxtraheerd.txt")
                with open(warning_path, "w") as warning_file:
                    warning_file.write("Waarschuwing: Je Word-bestand bevat hoogstwaarschijnlijk figuren die in Word zijn samengesteld. Dit herken je aan bestanden in deze map met een .emf or .tmp bestandsextensie. Deze worden niet correct geconverteerd naar losse afbeeldingen. Controleer het Word-document, kijk of de klant deze specifieke figuren los kan aanleveren, of exporteer ze eventueel vanuit Word als een .pdf-bestand.\n")
                img_zip.write(warning_path, "!LET OP! Niet alle figuren zijn succesvol geëxtraheerd.txt")
        
        # Verwijder tijdelijke bestanden
        shutil.rmtree(extracted_images_path)
        os.remove(filepath)
        if warning:
            os.remove(warning_path)
        
        return send_file(zip_filepath, as_attachment=True)
    
    return render_template('upload.html')

def extract_images(docx_path):
    temp_extract_dir = os.path.join(app.config['EXTRACT_FOLDER'], os.path.basename(docx_path).replace(".docx", ""))
    os.makedirs(temp_extract_dir, exist_ok=True)
    
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(temp_extract_dir)
    
    media_path = os.path.join(temp_extract_dir, "word/media")
    if not os.path.exists(media_path):
        return None, False
    
    result_dir = os.path.join(app.config['RESULT_FOLDER'], os.path.basename(docx_path).replace(".docx", ""))
    os.makedirs(result_dir, exist_ok=True)
    
    warning = False
    for img in os.listdir(media_path):
        if img.endswith(".emf") or img.endswith(".tmp"):
            warning = True
        shutil.move(os.path.join(media_path, img), os.path.join(result_dir, img))
    
    return result_dir, warning

if __name__ == '__main__':
    app.run(debug=True)
