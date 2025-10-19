# utils/slide_builder.py
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import tempfile, os

class PresentationBuilder:
    def __init__(self, data, template_url=None):
        self.data = data
        self.slides_data = data.get('slides', [])
        self.default_font = data.get('theme', {}).get('font', 'Calibri')
        
        # Charger le template
        template_path = self._get_template_path(template_url)
        if template_path and os.path.exists(template_path):
            self.prs = Presentation(template_path)
            print(f"✅ Template chargé : {template_path}")
        else:
            print("⚠️ Template non trouvé")
            self.prs = Presentation()
        
        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(5.625)
    
    def _get_template_path(self, template_url):
        """Retourne le chemin du template local."""
        local_path = os.path.join(
            os.path.dirname(__file__),
            "Formation.pptx"
        )
        return local_path if os.path.exists(local_path) else None
    
    def build(self):
        """Génère le PowerPoint complet."""
        # Supprimer toutes les slides sauf la première (titre)
        while len(self.prs.slides) > 1:
            rId = self.prs.slides._sldIdLst[1].rId
            self.prs.part.drop_rel(rId)
            del self.prs.slides._sldIdLst[1]
        
        print(f"📊 Génération de {len(self.slides_data)} slides...")
        
        # Générer les slides d'exercices
        for i, slide_data in enumerate(self.slides_data):
            self._add_slide_from_scratch(slide_data)
            print(f"✅ Slide {i+2}/{len(self.slides_data)+1} : {slide_data.get('type')}")
        
        # Sauvegarder
        tmp_out = tempfile.mkstemp(suffix='.pptx')[1]
        self.prs.save(tmp_out)
        print(f"✅ PPTX généré avec {len(self.prs.slides)} slides")
        return tmp_out
    
    def _add_slide_from_scratch(self, slide_data):
        """Crée une nouvelle slide avec le layout du template."""
        # Utiliser un layout vide du template (blank ou title only)
        # Le layout 6 est généralement "Blank"
        layout_index = 6 if len(self.prs.slide_layouts) > 6 else 1
        layout = self.prs.slide_layouts[layout_index]
        
        slide = self.prs.slides.add_slide(layout)
        
        # Récupérer les données
        titre = slide_data.get('titre', 'Exercice')
        ex_type = slide_data.get('type', 'exercice_pratique')
        
        # TITRE en haut (grand, centré)
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.5), Inches(9), Inches(0.8)
        )
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        p = title_frame.paragraphs[0]
        p.text = titre
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.name = self.default_font
        
        # CONTENU selon le type
        if ex_type == 'qcm':
            self._add_qcm_content(slide, slide_data)
        elif ex_type == 'vrai_faux':
            self._add_vrai_faux_content(slide, slide_data)
        else:
            self._add_generic_content(slide, slide_data)
    
    def _add_qcm_content(self, slide, data):
        """Ajoute le contenu d'un QCM."""
        # Question
        question = data.get('question', '')
        if question:
            q_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(1.5), Inches(9), Inches(1)
            )
            tf = q_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = question
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.name = self.default_font
        
        # Choix
        choix = data.get('choix', [])
        if choix:
            choices_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(2.8), Inches(9), Inches(2.5)
            )
            tf = choices_box.text_frame
            tf.word_wrap = True
            
            for i, choix_text in enumerate(choix):
                if i > 0:
                    p = tf.add_paragraph()
                else:
                    p = tf.paragraphs[0]
                
                p.text = f"{chr(65+i)}. {choix_text}"
                p.font.size = Pt(16)
                p.font.name = self.default_font
                p.space_before = Pt(6)
                p.space_after = Pt(6)
    
    def _add_vrai_faux_content(self, slide, data):
        """Ajoute le contenu Vrai/Faux."""
        affirmations = data.get('affirmations', [])
        if affirmations:
            aff_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(1.5), Inches(9), Inches(3.5)
            )
            tf = aff_box.text_frame
            tf.word_wrap = True
            
            for i, aff in enumerate(affirmations):
                if i > 0:
                    p = tf.add_paragraph()
                else:
                    p = tf.paragraphs[0]
                
                p.text = f"• {aff.get('affirmation', '')}"
                p.font.size = Pt(16)
                p.font.name = self.default_font
                p.space_before = Pt(8)
                p.space_after = Pt(8)
    
    def _add_generic_content(self, slide, data):
        """Ajoute le contenu générique (exercice pratique, mise en situation)."""
        # Contexte
        contexte = data.get('contexte', '')
        if contexte:
            ctx_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(1.5), Inches(9), Inches(1.5)
            )
            tf = ctx_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = contexte
            p.font.size = Pt(14)
            p.font.name = self.default_font
        
        # Consigne
        consigne = data.get('consigne', '')
        if consigne:
            cons_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(3.2), Inches(9), Inches(1.5)
            )
            tf = cons_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = f"Consigne : {consigne}"
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.name = self.default_font