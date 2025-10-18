pythonfrom flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import tempfile
import traceback
from datetime import datetime
from utils.slide_builder import PresentationBuilder

app = Flask(__name__)
CORS(app)

@app.route('/health', methods=['GET'])
def health():
    """Endpoint de santé"""
    return jsonify({
        "status": "healthy",
        "service": "ZMForma PowerPoint Generator v2",
        "version": "2.0.0",
        "timestamp": datetime.now().isoformat()
    })

@app.route('/generate', methods=['POST'])
def generate_pptx():
    """
    Génère un PowerPoint professionnel à partir du JSON
    """
    try:
        # Récupération des données
        data = request.get_json()
        
        if not data:
            return jsonify({"error": "No JSON data provided"}), 400
        
        if 'slides' not in data or not isinstance(data['slides'], list):
            return jsonify({"error": "Invalid format: 'slides' array required"}), 400
        
        if len(data['slides']) == 0:
            return jsonify({"error": "No slides to generate"}), 400
        
        print(f"📊 Génération de {len(data['slides'])} slides...")
        
        # Construction de la présentation
        builder = PresentationBuilder(data)
        pptx_path = builder.build()
        
        # Lecture du fichier
        with open(pptx_path, 'rb') as f:
            pptx_content = f.read()
        
        # Nettoyage
        try:
            os.remove(pptx_path)
        except:
            pass
        
        # Envoi du fichier
        filename = data.get('filename', f'Formation_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pptx')
        
        # Sauvegarde temporaire pour envoi
        temp_path = os.path.join(tempfile.gettempdir(), filename)
        with open(temp_path, 'wb') as f:
            f.write(pptx_content)
        
        print(f"✅ Présentation générée: {filename}")
        
        response = send_file(
            temp_path,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=filename
        )
        
        # Nettoyage après envoi
        @response.call_on_close
        def cleanup():
            try:
                os.remove(temp_path)
            except:
                pass
        
        return response
        
    except Exception as e:
        print(f"❌ Erreur: {str(e)}")
        print(traceback.format_exc())
        return jsonify({
            "error": str(e),
            "traceback": traceback.format_exc()
        }), 500

@app.route('/test', methods=['GET'])
def test():
    """Endpoint de test avec exemple minimal"""
    test_data = {
        "slides": [
            {
                "type": "cover",
                "background": "2563EB",
                "layout": {
                    "title": {
                        "text": "Test Présentation ZMForma",
                        "x": 0.5, "y": 2.0, "w": 5.0, "h": 1.0,
                        "fontSize": 44,
                        "bold": True,
                        "color": "FFFFFF",
                        "align": "left"
                    },
                    "subtitle": {
                        "text": "API v2.0 - Système de génération automatique",
                        "x": 0.5, "y": 3.2, "w": 5.0, "h": 0.6,
                        "fontSize": 20,
                        "color": "FFFFFF",
                        "align": "left"
                    }
                }
            }
        ],
        "theme": {
            "font": "Arial",
            "background": "FFFFFF"
        },
        "filename": "test_api_v2.pptx"
    }
    
    try:
        builder = PresentationBuilder(test_data)
        pptx_path = builder.build()
        return send_file(
            pptx_path, 
            as_attachment=True, 
            download_name="test_zmforma_v2.pptx",
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    except Exception as e:
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
