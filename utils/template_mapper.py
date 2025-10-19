# utils/template_mapper.py
from pptx import Presentation
from pptx.util import Inches, Pt
from copy import deepcopy

class TemplateMapper:
    """Mappe les exercices aux slides du template et remplit les placeholders."""
    
    def __init__(self, template_path):
        self.template = Presentation(template_path)
        
        # Mapping type d'exercice → index de slide template
        self.slide_mapping = {
            'qcm': 2,           # Slide 3 (index 2)
            'vrai_faux': 4,     # Slide 5 (index 4)
            'cas_pratique': 6,  # Slide 7 (index 6)
            'mise_en_situation': 8,  # Slide 9 (index 8)
            'exercice_pratique': 6   # Slide 7 aussi
        }
    
    def find_text_shapes(self, slide):
        """Trouve toutes les zones de texte dans une slide."""
        text_shapes = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                text_shapes.append(shape)
        return text_shapes
    
    def replace_text_in_shape(self, shape, old_text, new_text):
        """Remplace du texte dans un shape."""
        if not shape.has_text_frame:
            return False
        
        text_frame = shape.text_frame
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                if old_text.lower() in run.text.lower():
                    run.text = run.text.replace(old_text, new_text)
                    return True
        return False
    
    def fill_exercise_slide(self, slide, exercise_data):
        """Remplit une slide d'exercice avec les données."""
        text_shapes = self.find_text_shapes(slide)
        
        # Remplacer le titre de l'exercice
        titre = exercise_data.get('titre', 'Exercice')
        for shape in text_shapes:
            if 'Rédiger' in shape.text or 'Créer' in shape.text or 'Exercice' in shape.text:
                self.replace_text_in_shape(shape, shape.text, titre)
                break
        
        # Remplacer la consigne
        consigne = exercise_data.get('consigne', '')
        for shape in text_shapes:
            if 'Consigne' in shape.text:
                self.replace_text_in_shape(shape, 'Consigne :', f"Consigne : {consigne}")
                break
        
        # Gérer les cas spécifiques
        if exercise_data.get('type') == 'qcm':
            self._fill_qcm(slide, exercise_data)
        elif exercise_data.get('type') == 'vrai_faux':
            self._fill_vrai_faux(slide, exercise_data)
    
    def _fill_qcm(self, slide, data):
        """Remplit les choix d'un QCM."""
        if not data.get('choix'):
            return
        
        for shape in slide.shapes:
            if shape.has_text_frame and any(word in shape.text for word in ['Respecter', 'Maîtriser', 'Utiliser']):
                tf = shape.text_frame
                tf.clear()
                
                # Ajouter la question
                p = tf.paragraphs[0] if tf.paragraphs else tf.add_paragraph()
                p.text = data.get('question', '')
                p.level = 0
                
                # Ajouter les choix
                for i, choix in enumerate(data.get('choix', [])):
                    p = tf.add_paragraph()
                    p.text = f"{chr(65+i)}. {choix}"
                    p.level = 0
                    p.space_before = Pt(6)
                break
    
    def _fill_vrai_faux(self, slide, data):
        """Remplit les affirmations Vrai/Faux."""
        if not data.get('affirmations'):
            return
        
        for shape in slide.shapes:
            if shape.has_text_frame and any(word in shape.text for word in ['Respecter', 'Maîtriser', 'Extraire']):
                tf = shape.text_frame
                tf.clear()
                
                for aff in data.get('affirmations', []):
                    p = tf.paragraphs[0] if len(tf.paragraphs) == 1 else tf.add_paragraph()
                    p.text = f"• {aff.get('affirmation', '')}"
                    p.level = 0
                break