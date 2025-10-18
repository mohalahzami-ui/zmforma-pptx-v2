@app.route('/test', methods=['GET'])
def test():
    """Génère 3 slides de test avec le template"""
    test_data = {
        "slides": [
            # SLIDE 1 : QCM
            {
                "type": "qcm",
                "background": None,
                "layout": {
                    "kicker": {
                        "text": "ACCUEIL ET COMMUNICATION • Exercice 1",
                        "x": 0.6, "y": 0.6, "w": 8.8, "h": 0.35,
                        "fontSize": 11, "color": "64748B", "bold": True
                    },
                    "question": {
                        "text": "Quelle est la formule d'accueil téléphonique professionnelle et complète ?",
                        "x": 0.6, "y": 1.1, "w": 5.5, "h": 1.2,
                        "fontSize": 26, "bold": True, "color": "1E293B"
                    },
                    "context": {
                        "text": "📋 Contexte : Vous êtes secrétaire dans un cabinet d'avocats réputé. C'est lundi matin 9h. Le téléphone sonne, c'est le premier appel de la journée.",
                        "x": 0.6, "y": 2.4, "w": 5.5, "h": 0.9,
                        "fontSize": 13, "color": "475569"
                    },
                    "choices": {
                        "items": [
                            "A) Oui, allô, c'est pour quoi ?",
                            "B) Bonjour, Cabinet Martin & Associés, bonjour !",
                            "C) Bonjour, Cabinet Martin & Associés, Sophie à votre écoute, comment puis-je vous aider ?",
                            "D) Cabinet Martin, que puis-je faire pour vous ?"
                        ],
                        "x": 0.6, "y": 3.4, "w": 5.5, "h": 1.6,
                        "fontSize": 15, "color": "0F172A"
                    }
                }
            },
            
            # SLIDE 2 : CORRECTION
            {
                "type": "correction",
                "background": None,
                "layout": {
                    "label": {
                        "text": "✅ CORRECTION",
                        "x": 0.6, "y": 0.6, "w": 8.8, "h": 0.35,
                        "fontSize": 12, "bold": True, "color": "059669"
                    },
                    "answer": {
                        "text": "Réponse correcte : C",
                        "x": 0.6, "y": 1.1, "w": 8.8, "h": 0.65,
                        "fontSize": 22, "bold": True, "color": "0F172A"
                    },
                    "answer_text": {
                        "text": "Bonjour, Cabinet Martin & Associés, Sophie à votre écoute, comment puis-je vous aider ?",
                        "x": 0.6, "y": 1.85, "w": 8.8, "h": 0.55,
                        "fontSize": 15, "color": "475569"
                    },
                    "explanation": {
                        "text": "💡 Explication :\n\nUne formule d'accueil téléphonique professionnelle complète doit contenir 4 éléments essentiels :\n\n1. La salutation (Bonjour)\n2. Le nom de l'entreprise/service\n3. Votre prénom\n4. Une proposition d'aide\n\nCette structure rassure l'interlocuteur, le situe immédiatement et montre votre disponibilité. Les autres réponses sont soit trop familières (A), incomplètes (B et D), ou manquent de professionnalisme.",
                        "x": 0.6, "y": 2.5, "w": 8.8, "h": 2.5,
                        "fontSize": 14, "color": "1E293B"
                    }
                }
            },
            
            # SLIDE 3 : VRAI/FAUX
            {
                "type": "vrai_faux",
                "background": None,
                "layout": {
                    "kicker": {
                        "text": "MODULE 1 : RÉDACTION • Exercice 2",
                        "x": 0.6, "y": 0.6, "w": 8.8, "h": 0.35,
                        "fontSize": 11, "color": "64748B", "bold": True
                    },
                    "title": {
                        "text": "Vrai ou Faux : Les Formules de Politesse",
                        "x": 0.6, "y": 1.1, "w": 8.8, "h": 0.8,
                        "fontSize": 24, "bold": True, "color": "1E293B"
                    },
                    "consigne": {
                        "text": "Indiquez si les affirmations suivantes concernant les formules de politesse professionnelles sont vraies ou fausses.",
                        "x": 0.6, "y": 2.0, "w": 8.8, "h": 0.5,
                        "fontSize": 12, "color": "64748B"
                    },
                    "items": {
                        "items": [
                            "1. On peut terminer un email professionnel par 'Bisous' si on connaît bien son interlocuteur",
                            "2. 'Je vous prie d'agréer, Madame, Monsieur' est une formule adaptée uniquement aux courriers papier",
                            "3. Dans un email, 'Cordialement' est une formule de politesse professionnelle appropriée",
                            "4. Il faut toujours utiliser la même formule de politesse quel que soit le destinataire"
                        ],
                        "x": 0.6, "y": 2.6, "w": 8.8, "h": 2.7,
                        "fontSize": 14, "bullet": True, "color": "0F172A"
                    }
                }
            }
        ],
        "theme": {
            "font": "Arial"
        },
        "filename": "test_3_slides_zmforma.pptx"
    }
    
    try:
        builder = PresentationBuilder(test_data, template_url=TEMPLATE_URL)
        pptx_path = builder.build()
        return send_file(
            pptx_path,
            as_attachment=True,
            download_name="test_3_slides_template.pptx",
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    except Exception as e:
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500
