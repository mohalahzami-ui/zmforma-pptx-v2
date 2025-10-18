from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
import tempfile
import os
import requests
from io import BytesIO
from PIL import Image
from .styles import Colors, Formatter

class PresentationBuilder:
    """
    BUILDER OPTIMISÉ - UTILISE LES PLACEHOLDERS DU TEMPLATE
    """

    def __init__(self, data, template_url=None):
        self.data = data
        self.slides_data = data.get('slides', [])
        self.theme = data.get('theme', {})
        self.default_font = self.theme.get('font', 'Arial')
        self._tmp_template_path = None

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
                
                # Suppression des slides existantes
                nb_slides = len(self.prs.slides)
                if nb_slides > 0:
                    print(f"🗑️  Suppression de {nb_slides} slides du template...")
                    while len(self.prs.slides) > 0:
                        rId = self.prs.slides._sldIdLst[0].rId
                        self.prs.part.drop_rel(rId)
                        del self.prs.slides._sldIdLst[0]
                    print("✅ Template nettoyé")
                
            except Exception as e:
                print(f"⚠️  Erreur template: {e}")
                self.prs = Presentation()
        else:
            self.prs = Presentation()

        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(5.625)

        # Liste des layouts
        self._layout_names = []
        try:
            for i, l in enumerate(self.prs.slide_layouts):
                name = getattr(l, "name", f"Layout {i}")
                self._layout_names.append(name)
                print(f"  Layout {i}: {name}")
        except Exception:
            pass

    def _pick_layout(self, slide_type):
        """Sélection layout - PRIORITÉ AU BLANK pour contrôle total"""
        layouts = self.prs.slide_layouts
        
        # Pour TOUS les types, on veut "Blank" pour avoir le contrôle total
        for layout in layouts:
            try:
                if 'blank' in layout.name.lower():
                    return layout
            except Exception:
                pass
        
        # Fallback
        return layouts[min(6, len(layouts)-1)] if len(layouts) > 6 else layouts[0]

    def build(self):
        """Construction"""
        print(f"🔨 Construction de {len(self.slides_data)} slides...")

        for i, slide_data in enumerate(self.slides_data):
            try:
                slide_type = slide_data.get('type', 'generic')
                layout = self._pick_layout(slide_type)
                layout_name = getattr(layout, 'name', '?')
                print(f"  • Slide {i+1}: {slide_type} -> {layout_name}")

                slide = self.prs.slides.add_slide(layout)

                # NE JAMAIS repeindre le fond
                self._fill_slide(slide, slide_data)

            except Exception as e:
                print(f"❌ Erreur slide {i+1}: {str(e)}")
                import traceback
                print(traceback.format_exc())

        temp_path = os.path.join(tempfile.gettempdir(), 'presentation_zmforma.pptx')
        self.prs.save(temp_path)
        print(f"✅ Présentation sauvegardée: {temp_path}")

        if self._tmp_template_path and os.path.exists(self._tmp_template_path):
            try:
                os.remove(self._tmp_template_path)
            except Exception:
                pass

        return temp_path

    def _fill_slide(self, slide, data):
        """Remplit la slide - SANS BORDURES"""
        layout_cfg = data.get('layout', {})
        
        for key, element in layout_cfg.items():
            if not element or not isinstance(element, dict):
                continue
            
            if 'items' in element:
                self._add_bullets_clean(slide, element)
            elif 'text' in element:
                self._add_textbox_clean(slide, element)
            elif 'url' in element:
                self._add_image(slide, element)
            elif 'fill' in element:
                self._add_shape(slide, element)

    def _add_textbox_clean(self, slide, config):
        """Ajoute textbox SANS bordures ni cadres"""
        if not config or 'text' not in config:
            return
        
        x = Inches(config.get('x', 0.5))
        y = Inches(config.get('y', 1.0))
        w = Inches(config.get('w', 5.0))
        h = Inches(config.get('h', 1.0))

        textbox = slide.shapes.add_textbox(x, y, w, h)
        
        # CRITIQUE : Supprimer toutes les bordures
        textbox.line.fill.background()  # Pas de bordure
        
        text_frame = textbox.text_frame
        text_frame.clear()
        text_frame.word_wrap = True
        text_frame.margin_left = 0
        text_frame.margin_right = 0
        text_frame.margin_top = 0
        text_frame.margin_bottom = 0
        
        # Ajouter le texte
        p = text_frame.paragraphs[0]
        p.text = str(config['text'])
        p.alignment = PP_ALIGN.LEFT
        
        # Formatage
        for run in p.runs:
            run.font.name = config.get('font', self.default_font)
            run.font.size = Pt(config.get('fontSize', 16))
            run.font.bold = config.get('bold', False)
            
            color = config.get('color')
            if color:
                run.font.color.rgb = Colors.hex_to_rgb(color)

    def _add_bullets_clean(self, slide, config):
        """Ajoute bullets SANS bordures"""
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
        
        # CRITIQUE : Supprimer bordures
        textbox.line.fill.background()
        
        text_frame = textbox.text_frame
        text_frame.clear()
        text_frame.word_wrap = True
        text_frame.margin_left = 0
        text_frame.margin_right = 0
        text_frame.margin_top = 0
        text_frame.margin_bottom = 0
        
        for i, item in enumerate(items):
            if not item:
                continue
            
            p = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
            p.text = str(item)
            p.level = 0
            p.alignment = PP_ALIGN.LEFT
            
            if config.get('bullet', True):
                p.bullet = True
            
            for run in p.runs:
                run.font.name = config.get('font', self.default_font)
                run.font.size = Pt(config.get('fontSize', 16))
                run.font.bold = config.get('bold', False)
                
                color = config.get('color')
                if color:
                    run.font.color.rgb = Colors.hex_to_rgb(color)

    def _add_shape(self, slide, config):
        """Ajoute forme"""
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
        """Ajoute image"""
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
