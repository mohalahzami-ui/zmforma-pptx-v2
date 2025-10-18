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
    BUILDER POUR TEMPLATE GAMMA - SUPPRESSION AUTO DES SLIDES
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
                
                # 🔥 SUPPRESSION AUTOMATIQUE DES SLIDES EXISTANTES
                nb_slides = len(self.prs.slides)
                if nb_slides > 0:
                    print(f"🗑️  Suppression de {nb_slides} slides du template...")
                    # Méthode propre pour supprimer les slides
                    while len(self.prs.slides) > 0:
                        rId = self.prs.slides._sldIdLst[0].rId
                        self.prs.part.drop_rel(rId)
                        del self.prs.slides._sldIdLst[0]
                    print("✅ Template nettoyé (slides supprimées, design conservé)")
                
            except Exception as e:
                print(f"⚠️  Erreur template: {e}")
                self.prs = Presentation()
        else:
            self.prs = Presentation()

        # Forcer 16:9
        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(5.625)

        # Lister layouts
        self._layout_names = []
        try:
            for i, l in enumerate(self.prs.slide_layouts):
                name = getattr(l, "name", f"Layout {i}")
                self._layout_names.append(name)
            print("📐 Layouts disponibles:", self._layout_names)
        except Exception:
            pass

    def _pick_layout(self, slide_type):
        """Sélection du layout selon le type"""
        layouts = self.prs.slide_layouts
        
        type_map = {
            'cover': ['title', 'cover', 'titre'],
            'section': ['section', 'separator'],
            'qcm': ['content', 'title and content', 'blank'],
            'vrai_faux': ['content', 'title and content', 'blank'],
            'cas_pratique': ['content', 'title and content', 'blank'],
            'mise_en_situation': ['content', 'title and content', 'blank'],
            'correction': ['content', 'title and content', 'blank']
        }
        
        preferred = type_map.get(slide_type, ['blank', 'content'])
        
        for keyword in preferred:
            for layout in layouts:
                try:
                    if keyword.lower() in layout.name.lower():
                        return layout
                except Exception:
                    pass
        
        # Fallback : Blank ou premier layout
        for layout in layouts:
            try:
                if 'blank' in layout.name.lower():
                    return layout
            except Exception:
                pass
        
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

                # NE JAMAIS repeindre le fond (garde le template)
                bg = slide_data.get('background')
                if bg and bg not in (None, '', 'None', 'null'):
                    background = slide.background
                    fill = background.fill
                    fill.solid()
                    fill.fore_color.rgb = Colors.hex_to_rgb(bg)

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
        """Remplit la slide"""
        layout_cfg = data.get('layout', {})
        
        for key, element in layout_cfg.items():
            if not element or not isinstance(element, dict):
                continue
            
            if 'items' in element:
                self._add_bullets(slide, element)
            elif 'text' in element:
                self._add_textbox(slide, element)
            elif 'url' in element:
                self._add_image(slide, element)
            elif 'fill' in element:
                self._add_shape(slide, element)

    def _add_textbox(self, slide, config):
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
