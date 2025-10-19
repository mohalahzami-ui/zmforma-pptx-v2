# app.py
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import os
from datetime import datetime
from utils.extract import extract_text_pdf, extract_text_pptx
from utils.slide_builder import PresentationBuilder

app = Flask(__name__)
CORS(app)

TEMPLATE_URL = os.environ.get("TEMPLATE_URL")  # à configurer via Render ou .env

# === AJOUT : constantes pour servir le PPTX local ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PPTX_DIR = os.path.join(BASE_DIR, "utils")   # dossier où se trouve ton fichier
PPTX_FILE = "Formation.pptx"                 # nom exact du fichier à télécharger

@app.route('/health', methods=['GET'])
def health():
    return jsonify({
        "service": "ZMForma PowerPoint Generator v2",
        "status": "healthy",
        "template_url_set": bool(TEMPLATE_URL),
        "version": "2.1.0",
        "timestamp": datetime.now().isoformat()
    })

@app.route('/extract', methods=['POST'])
def extract():
    if 'file' not in request.files:
        return jsonify({"error": "Missing file"}), 400
    f = request.files['file']
    filename = f.filename or ""
    ext = filename.rsplit('.',1)[-1].lower()
    data = f.read()
    if ext == 'pdf':
        out = extract_text_pdf(data)
    elif ext == 'pptx':
        out = extract_text_pptx(data)
    else:
        return jsonify({"error": "Unsupported filetype"}), 400
    out['filename'] = filename
    return jsonify(out)

@app.route('/generate', methods=['POST'])
def generate():
    payload = request.get_json()
    if not payload or not isinstance(payload.get('slides'), list):
        return jsonify({"error":"Invalid payload"}), 400
    builder = PresentationBuilder(payload, template_url=TEMPLATE_URL)
    pptx_path = builder.build()
    filename = payload.get('filename', f"Formation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx")
    return send_file(
        pptx_path,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        as_attachment=True,
        download_name=filename
    )

# === AJOUT : route /download pour servir utils/Formation.pptx ===
@app.route('/download', methods=['GET'])
def download_pptx():
    try:
        file_path = os.path.join(PPTX_DIR, PPTX_FILE)
        if not os.path.exists(file_path):
            return jsonify({"error": f"Le fichier {PPTX_FILE} est introuvable dans utils/."}), 404
        return send_from_directory(
            PPTX_DIR,
            PPTX_FILE,
            as_attachment=True
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)
