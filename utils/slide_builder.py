# utils/slide_builder.py
from pptx import Presentation
from pptx.util import Inches
import tempfile, os, requests
from .template_mapper import TemplateMapper

class PresentationBuilder:
    def __init__(self, data, template_url=None):
        self.data = data
        self.slides_data = data.get('slides', [])
        self.default_font = data.get('theme', {}).get('font', 'Arial')
        
        # Charger le template
        template_path = self._get_template_path(template_url)
        if template_path and os.path.exists(template_path):
            self.mapper = TemplateMapper(template_path)
            self.prs = Presentation(template_path)
            print(f"✅ Template chargé : {template_path}")
        else:
            print("⚠️ Template non trouvé, création d'une présentation vide")
            self.prs = Presentation()
            self.mapper = None
    
    def _get_template_path(self, template_url):
        """Retourne le chemin du template local."""
        local_path = os.path.join(
            os.path.dirname(__file__),
            "Formation.pptx"
        )
        return local_path if os.path.exists(local_path) else None
    
    def build(self):
        """Génère le PowerPoint complet."""
        if not self.mapper:
            print("❌ Impossible de générer sans template")
            return None
        
        # Supprimer toutes les slides sauf la première (page de titre)
        while len(self.prs.slides) > 1:
            rId = self.prs.slides._sldIdLst[1].rId
            self.prs.part.drop_rel(rId)
            del self.prs.slides._sldIdLst[1]
        
        # Générer les slides d'exercices
        for i, slide_data in enumerate(self.slides_data):
            self._add_exercise_slide(slide_data)
            print(f"✅ Slide {i+1}/{len(self.slides_data)} générée")
        
        # Sauvegarder
        tmp_out = tempfile.mkstemp(suffix='.pptx')[1]
        self.prs.save(tmp_out)
        print(f"✅ PPTX généré : {tmp_out}")
        return tmp_out
    
    def _add_exercise_slide(self, exercise_data):
        """Ajoute une slide d'exercice en clonant le template."""
        ex_type = exercise_data.get('type', 'exercice_pratique')
        
        # Obtenir l'index de la slide template
        template_index = self.mapper.slide_mapping.get(ex_type, 7)
        
        # Cloner la slide du template
        source_slide = self.prs.slides[template_index]
        
        # Dupliquer la slide
        slide_layout = source_slide.slide_layout
        new_slide = self.prs.slides.add_slide(slide_layout)
        
        # Copier tous les shapes
        for shape in source_slide.shapes:
            if not shape.is_placeholder:
                el = shape.element
                new_slide.shapes._spTree.insert_element_before(
                    el, 'p:extLst'
                )
        
        # Remplir avec les données de l'exercice
        self.mapper.fill_exercise_slide(new_slide, exercise_data)
