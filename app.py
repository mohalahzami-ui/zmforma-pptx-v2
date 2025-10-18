from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import tempfile
import traceback
from datetime import datetime
from utils.slide_builder import PresentationBuilder

app = Flask(__name__)
CORS(app)

TEMPLATE_URL = os.environ.get("TEMPLATE_URL")  # <- lit l’URL du template

@app.route('/health', methods=['GET'])
def health():
    """Endpoint de santé"""
    return jsonify({
        "status": "healthy",
        "service": "ZMForma PowerPoint Generator v2",
        "version": "2.1.0",
        "template_url_set": bool(TEMPLATE_URL),
        "timestamp": datetime.now().isoformat()
    })

@app.route('/generate', methods=['POST'])
def generate_pptx():
    """Génère un PowerPoint à partir du JSON"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No JSON data provided"}), 400
        if 'slides' not in data or not isinstance(data['slides'], list):
            return jsonify({"error": "Invalid format: 'slides' array required"}), 400
        if len(data['slides']) == 0:
            return jsonify({"error": "No slides to generate"}), 400

        print(f"📊 Génération de {len(data['slides'])} slides...")
        print(f"🎨 Template actif: {bool(TEMPLATE_URL)}")

        builder = PresentationBuilder(data, template_url=TEMPLATE_URL)
        pptx_path = builder.build()

        filename = data.get('filename', f'Formation_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pptx')
        return send_file(
            pptx_path,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        print(f"❌ Erreur: {str(e)}")
        print(traceback.format_exc())
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500

@app.route('/test', methods=['GET'])
def test():
    """Slide de test minimale (utilise aussi le template si dispo)"""
    test_data = {
        "slides": [
            {
                "type": "qcm",
                "background": None,  # laissez None pour garder le fond du template
                "layout": {
                    "kicker": {"text": "MODULE • Exercice 1", "x": 0.6, "y": 0.6, "w": 4.8, "h": 0.3, "fontSize": 12, "color": "64748B"},
                    "question": {"text": "Titre de test", "x": 0.6, "y": 1.0, "w": 4.8, "h": 1.0, "fontSize": 28, "bold": True},
                    "choices": {"items": ["A) Choix A","B) Choix B","C) Choix C"], "x": 0.6, "y": 2.2, "w": 4.8, "h": 1.2, "fontSize": 16}
                }
            }
        ],
        "theme": {"font": "Arial"}
    }
    builder = PresentationBuilder(test_data, template_url=TEMPLATE_URL)
    pptx_path = builder.build()
    return send_file(
        pptx_path,
        as_attachment=True,
        download_name="test_template.pptx",
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
