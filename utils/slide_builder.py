from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE
import tempfile
import os
import requests
from io import BytesIO
from PIL import Image
from .styles import Colors, Formatter

class PresentationBuilder:
    """
    BUILDER OPTIMISÉ POUR TEMPLATE GAMMA
    - Utilise les LAYOUTS NATIFS du template .pptx
    - NE repeint JAMAIS le fond (garde la charte graphique)
    - Positionne le contenu EXACTEMENT selon le template
    """

    def __init__(self, data, template_url=None):
        self.data = data
        self.slides_data = data.get('slides', [])
        self.theme = data.get('theme', {})
        self.default_font = self.theme.get('font', 'Arial')
        self._tmp_template_path = None

        # Charger template distant
        if template_url:
            try:
                print(f"⬇️  Téléchargement template: {template_url}")
                r = requests.get(template_url, timeout=25)
                r.raise_for_status()
                fd, path = tempfile.mkstemp(suffix=".pptx")
                os.close(fd)
                with open(path, "wb") as f:
                    f.write(r.content)
                self._tmp_template_path = path
                self.prs = Presentation(self._tmp_template_path)
                print("✅ Template chargé")
            except Exception as e:
                print(f"⚠️  Erreur template: {e}")
                self.prs = Presentation()
        else:
            self.prs = Presentation()

        # Forcer 16:9
        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(5.625)

        # Lister layouts disponibles
        self._layout_names = []
        try:
            for i, l in enumerate(self.prs.slide_layouts):
                name = getattr(l, "name", f"Layout {i}")
                self._layout_names.append(name)
            print("📐 Layouts template:", self._layout_names)
        except Exception:
            pass

    def _pick_layout(self, slide_type):
        """
        Sélection INTELLIGENTE du layout selon le type de slide
        Priorité : layouts du template Gamma
        """
        layouts = self.prs.slide_layouts
        
        # Mapping type -> nom de layout (ajuste selon ton template)
        type_map = {
            'cover': ['title', 'cover', 'titre'],
            'section': ['section', 'separator', 'chapitre'],
            'qcm': ['content', 'title and content', 'titre et contenu'],
            'vrai_faux': ['content', 'title and content'],
            'cas_pratique': ['content', 'title and content'],
            'mise_en_situation': ['content', 'title and content'],
            'objectifs': ['content', 'title and content'],
            'correction': ['content', 'title and content']
        }
        
        preferred = type_map.get(slide_type, ['content', 'title and content'])
        
        # Chercher un layout correspondant
        for keyword in preferred:
            for i, layout in enumerate(layouts):
                try:
                    if keyword.lower() in layout.name.lower():
                        return layout
                except Exception:
                    pass
        
        # Fallback : "Title and Content" (index 1) ou Blank (index 6)
        try:
            if len(layouts) > 1:
                return layouts[1]
        except Exception:
            pass
        
        return layouts[0]

    def build(self):
        """Construction de la présentation"""
        print(f"🔨 Construction de {len(self.slides_data)} slides...")

        for i, slide_data in enumerate(self.slides_data):
            try:
                slide_type = slide_data.get('type', 'generic')
                layout = self._pick_layout(slide_type)
                layout_name = getattr(layout, 'name', '?')
                print(f"  • Slide {i+1}: {slide_type} -> {layout_name}")

                # Créer slide AVEC le layout du template
                slide = self.prs.slides.add_slide(layout)

                # ⚠️ NE JAMAIS repeindre le fond (sauf si explicitement demandé)
                bg = slide_data.get('background')
                if bg and bg not in (None, '', 'None', 'null'):
                    background = slide.background
                    fill = background.fill
                    fill.solid()
                    fill.fore_color.rgb = Colors.hex_to_rgb(bg)

                # Remplir contenu
                self._fill_slide(slide, slide_data)

            except Exception as e:
                print(f"❌ Erreur slide {i+1}: {str(e)}")
                import traceback
                print(traceback.format_exc())
                self._build_error_slide(i+1, str(e))

        # Sauvegarder
        temp_path = os.path.join(tempfile.gettempdir(), 'presentation_zmforma.pptx')
        self.prs.save(temp_path)
        print(f"✅ Présentation sauvegardée: {temp_path}")

        # Nettoyage
        if self._tmp_template_path and os.path.exists(self._tmp_template_path):
            try:
                os.remove(self._tmp_template_path)
            except Exception:
                pass

        return temp_path

    def _fill_slide(self, slide, data):
        """Remplit une slide selon son layout"""
        layout_cfg = data.get('layout', {})
        
        # Ajouter chaque élément
        for key, element in layout_cfg.items():
            if not element:
                continue
            
            if isinstance(element, dict):
                if 'items' in element:
                    self._add_bullets(slide, element)
                elif 'text' in element:
                    self._add_textbox(slide, element)
                elif 'url' in element:
                    self._add_image(slide, element)
                elif 'fill' in element:
                    self._add_shape(slide, element)

    def _add_textbox(self, slide, config):
        """Ajoute une textbox"""
        if not config or 'text' not in config:
            return
        
        x = Inches(config.get('x', 0.5))
        y = Inches(config.get('y', 1.0))
        w = Inches(config.get('w', 5.0))
        h = Inches(config.get('h', 1.0))

        textbox = slide.shapes.add_textbox(x, y, w, h)
        textbox.text = str(config['text'])
        
        try:
            textbox.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        except Exception:
            pass
        
        Formatter.format_textbox(textbox, config, self.default_font)

    def _add_bullets(self, slide, config):
        """Ajoute une liste à puces"""
        if not config or 'items' not in config:
            return
        
        items = config['items']
        if not items:
            return

        x = Inches(config.get('x', 0.5))
        y = Inches(config.get('y', 2.0))
        w = Inches(config.get('w', 5.0))
        h = Inches(config.get('h', 2.0))

        textbox = slide.shapes.add_textbox(x, y, w, h)
        
        try:
            textbox.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        except Exception:
            pass
        
        Formatter.add_bullet_points(textbox.text_frame, items, config, self.default_font)

    def _add_shape(self, slide, config):
        """Ajoute une forme (rectangle)"""
        if not config:
            return
        
        x = Inches(config.get('x', 0))
        y = Inches(config.get('y', 0))
        w = Inches(config.get('w', 1))
        h = Inches(config.get('h', 1))
        
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
        
        if 'fill' in config:
            shape.fill.solid()
            shape.fill.fore_color.rgb = Colors.hex_to_rgb(config['fill'])
        
        shape.line.fill.background()

    def _add_image(self, slide, config):
        """Ajoute une image"""
        if not config or 'url' not in config:
            return
        
        url = config['url']
        x = Inches(config.get('x', 5.0))
        y = Inches(config.get('y', 1.0))
        w = Inches(config.get('w', 3.0))
        h = Inches(config.get('h', 2.0))

        try:
            if url.startswith('http'):
                response = requests.get(url, timeout=12)
                response.raise_for_status()
                image_stream = BytesIO(response.content)
                img = Image.open(image_stream)
                img.verify()
                image_stream.seek(0)
                slide.shapes.add_picture(image_stream, x, y, width=w, height=h)
            else:
                if os.path.exists(url):
                    slide.shapes.add_picture(url, x, y, width=w, height=h)
        except Exception as e:
            print(f"⚠️  Image impossible: {url} - {str(e)}")

    def _build_error_slide(self, slide_num, error_msg):
        """Crée une slide d'erreur en cas de problème"""
        try:
            layout = self._pick_layout('generic')
            slide = self.prs.slides.add_slide(layout)
            
            # Fond rouge pâle
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(255, 240, 240)
            
            # Titre erreur
            title_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(1.5), Inches(9.0), Inches(0.8)
            )
            title_box.text = f"❌ Erreur - Slide {slide_num}"
            Formatter.format_textbox(title_box, {
                'fontSize': 32,
                'bold': True,
                'color': 'D32F2F',
                'align': 'center'
            })
            
            # Message d'erreur
            error_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(2.5), Inches(9.0), Inches(2.0)
            )
            error_box.text = str(error_msg)[:700]
            Formatter.format_textbox(error_box, {
                'fontSize': 14,
                'color': '666666',
                'align': 'left'
            })
        except Exception as e:
            print(f"⚠️ Impossible de créer slide d'erreur: {str(e)}")
