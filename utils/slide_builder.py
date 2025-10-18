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
    Constructeur de présentation PowerPoint
    - Utilise un template .pptx si template_url est fourni
    - Format 16:9 (10" x 5.625")
    """

    def __init__(self, data, template_url=None):
        self.data = data
        self.slides_data = data.get('slides', [])
        self.theme = data.get('theme', {})
        self.default_font = self.theme.get('font', 'Arial')

        self._tmp_template_path = None

        # Charger la présentation depuis template si dispo
        if template_url:
            try:
                print(f"⬇️ Téléchargement du template: {template_url}")
                r = requests.get(template_url, timeout=25)
                r.raise_for_status()
                fd, path = tempfile.mkstemp(suffix=".pptx")
                os.close(fd)
                with open(path, "wb") as f:
                    f.write(r.content)
                self._tmp_template_path = path
                self.prs = Presentation(self._tmp_template_path)
                print("🎨 Template chargé avec succès.")
            except Exception as e:
                print(f"⚠️ Impossible de charger le template: {e}. Utilisation d'un PPTX vierge.")
                self.prs = Presentation()
        else:
            self.prs = Presentation()

        # Forcer le 16:9
        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(5.625)

    def build(self):
        """Construit la présentation"""
        print(f"🔨 Construction de {len(self.slides_data)} slides...")

        for i, slide_data in enumerate(self.slides_data):
            try:
                slide_type = slide_data.get('type', 'generic')
                print(f"  • Slide {i+1}/{len(self.slides_data)}: {slide_type}")

                if slide_type == 'cover':
                    # souvent ignoré — mais on peut l’ajouter si demandé
                    self._build_generic(slide_data)
                elif slide_type == 'section':
                    self._build_generic(slide_data)
                elif slide_type == 'qcm':
                    self._build_qcm(slide_data)
                elif slide_type == 'correction':
                    self._build_correction(slide_data)
                elif slide_type == 'vrai_faux':
                    self._build_qcm(slide_data)  # même structure
                elif slide_type == 'cas_pratique':
                    self._build_qcm(slide_data)  # même grille base
                elif slide_type == 'mise_en_situation':
                    self._build_qcm(slide_data)
                elif slide_type == 'objectifs':
                    self._build_correction(slide_data)
                else:
                    self._build_generic(slide_data)

            except Exception as e:
                print(f"❌ Erreur slide {i+1}: {str(e)}")
                import traceback
                print(traceback.format_exc())
                self._build_error_slide(i+1, str(e))

        temp_path = os.path.join(tempfile.gettempdir(), 'presentation_zmforma.pptx')
        self.prs.save(temp_path)
        print(f"✅ Présentation sauvegardée: {temp_path}")

        # Nettoyage template temporaire
        if self._tmp_template_path and os.path.exists(self._tmp_template_path):
            try:
                os.remove(self._tmp_template_path)
            except:
                pass

        return temp_path

    # ---------- Builders ----------

    def _build_qcm(self, data):
        layout = data.get('layout', {})
        bg_color = data.get('background', None)  # None -> garder le fond du template

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

        if bg_color:  # seulement si on veut par-dessus le template
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = Colors.hex_to_rgb(bg_color)

        if 'accent_bar' in layout:
            self._add_shape(slide, layout['accent_bar'])
        if 'kicker' in layout:
            self._add_textbox(slide, layout['kicker'])
        if 'title' in layout:
            self._add_textbox(slide, layout['title'])
        if 'question' in layout:
            self._add_textbox(slide, layout['question'])
        if 'context' in layout and layout['context']:
            self._add_textbox(slide, layout['context'])
        if 'choices' in layout:
            self._add_bullets(slide, layout['choices'])
        if 'image' in layout and layout['image']:
            self._add_image(slide, layout['image'])

    def _build_correction(self, data):
        layout = data.get('layout', {})
        bg_color = data.get('background', None)

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

        if bg_color:
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = Colors.hex_to_rgb(bg_color)

        for key in ['accent_bar', 'label', 'answer', 'answer_text', 'title', 'explanation', 'corrections', 'elements']:
            if key in layout and layout[key]:
                if isinstance(layout[key], dict) and 'items' in layout[key]:
                    self._add_bullets(slide, layout[key])
                elif isinstance(layout[key], dict):
                    self._add_textbox(slide, layout[key])

    def _build_generic(self, data):
        layout = data.get('layout', {})
        bg_color = data.get('background', None)

        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

        if bg_color:
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = Colors.hex_to_rgb(bg_color)

        for key, element in layout.items():
            if element and isinstance(element, dict):
                if 'items' in element:
                    self._add_bullets(slide, element)
                elif 'text' in element:
                    self._add_textbox(slide, element)
                elif 'url' in element:
                    self._add_image(slide, element)
                elif 'fill' in element:
                    self._add_shape(slide, element)

    def _build_error_slide(self, slide_num, error_msg):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 240, 240)

        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9.0), Inches(0.8))
        title_box.text = f"❌ Erreur - Slide {slide_num}"
        Formatter.format_textbox(title_box, {'fontSize': 32, 'bold': True, 'color': 'D32F2F', 'align': 'center'})

        error_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9.0), Inches(2.0))
        error_box.text = str(error_msg)[:700]
        Formatter.format_textbox(error_box, {'fontSize': 14, 'color': '666666', 'align': 'left'})

    # ---------- Utilitaires ----------

    def _add_textbox(self, slide, config):
        if not config or 'text' not in config:
            return
        x = Inches(config.get('x', 0.6))
        y = Inches(config.get('y', 0.9))
        w = Inches(config.get('w', 4.8))
        h = Inches(config.get('h', 1.0))

        textbox = slide.shapes.add_textbox(x, y, w, h)
        textbox.text = str(config['text'])

        # Auto-fit pour éviter le chevauchement
        try:
            textbox.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        except Exception:
            pass

        Formatter.format_textbox(textbox, config, self.default_font)

    def _add_bullets(self, slide, config):
        if not config or 'items' not in config:
            return
        items = config['items']
        if not items or len(items) == 0:
            return

        x = Inches(config.get('x', 0.6))
        y = Inches(config.get('y', 2.0))
        w = Inches(config.get('w', 4.8))
        h = Inches(config.get('h', 1.8))

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
        x = Inches(config.get('x', 5.6))
        y = Inches(config.get('y', 1.0))
        w = Inches(config.get('w', 3.6))
        h = Inches(config.get('h', 2.25))

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
            print(f"⚠️ Impossible d'ajouter l'image {url}: {str(e)}")
