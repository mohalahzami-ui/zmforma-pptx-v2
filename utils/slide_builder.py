# utils/slide_builder.py
from pptx import Presentation
from pptx.util import Inches
import tempfile, os
from copy import deepcopy
from .template_mapper import TemplateMapper

class PresentationBuilder:
    def __init__(self, data, template_url=None):
        self.data = data
        self.slides_data = data.get('slides', [])
        self.default_font = data.get('theme', {}).get('font', 'Arial')
        
        # Charger le template
        template_path = self._get_template_path(template_url)
        if template_path and os.path.exists(template_path):
            self.prs = Presentation(template_path)
            self.mapper = TemplateMapper(template_path)
            print(f"✅ Template chargé : {template_path}")
        else:
            print("⚠️ Template non trouvé")
            self.prs = Presentation()
            self.mapper = None
        
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
        if not self.mapper:
            print("❌ Impossible de générer sans template")
            return self._build_without_template()
        
        # Supprimer toutes les slides sauf la première (titre)
        while len(self.prs.slides) > 1:
            rId = self.prs.slides._sldIdLst[1].rId
            self.prs.part.drop_rel(rId)
            del self.prs.slides._sldIdLst[1]
        
        print(f"📊 Génération de {len(self.slides_data)} slides d'exercices...")
        
        # Générer les slides d'exercices
        for i, slide_data in enumerate(self.slides_data):
            try:
                self._add_exercise_slide(slide_data)
                print(f"✅ Slide {i+2}/{len(self.slides_data)+1} générée : {slide_data.get('type')}")
            except Exception as e:
                print(f"❌ Erreur slide {i+2} : {e}")
        
        # Sauvegarder
        tmp_out = tempfile.mkstemp(suffix='.pptx')[1]
        self.prs.save(tmp_out)
        print(f"✅ PPTX généré : {tmp_out}")
        return tmp_out
    
    def _add_exercise_slide(self, exercise_data):
        """Ajoute une slide d'exercice en dupliquant une slide du template."""
        ex_type = exercise_data.get('type', 'exercice_pratique')
        
        # Obtenir l'index de la slide template
        template_index = self.mapper.slide_mapping.get(ex_type, 6)
        
        # IMPORTANT : Utiliser self.mapper.template au lieu de self.prs
        if template_index >= len(self.mapper.template.slides):
            template_index = 2  # Fallback
            print(f"⚠️ Index {template_index} hors limites, fallback sur slide 3")
        
        # Récupérer la slide SOURCE depuis le template original
        source_slide = self.mapper.template.slides[template_index]
        
        # Créer une nouvelle slide dans la présentation finale
        slide_layout = source_slide.slide_layout
        new_slide = self.prs.slides.add_slide(slide_layout)
        
        # Copier tous les shapes de la source vers la nouvelle slide
        for shape in source_slide.shapes:
            try:
                el = shape.element
                new_el = deepcopy(el)
                new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
            except Exception as e:
                print(f"⚠️ Impossible de copier shape : {e}")
        
        # Remplir avec les données de l'exercice
        self.mapper.fill_exercise_slide(new_slide, exercise_data)
    
    def _build_without_template(self):
        """Fallback : génération basique sans template."""
        for slide_data in self.slides_data:
            layout = self.prs.slide_layouts[1]
            slide = self.prs.slides.add_slide(layout)
            title = slide.shapes.title
            title.text = slide_data.get('titre', 'Exercice')
        
        tmp_out = tempfile.mkstemp(suffix='.pptx')[1]
        self.prs.save(tmp_out)
        return tmp_out