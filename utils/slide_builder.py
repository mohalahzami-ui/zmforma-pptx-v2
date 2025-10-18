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
    - Utilise un template .pptx distant si template_url est fourni
    - Sélectionne une MISE EN PAGE (layout) du template par nom ou index
    - Format 16:9 (10" x 5.625")
    """

    def __init__(self, data, template_url=None):
        self.data = data
        self.slides_data = data.get('slides', [])
        self.theme = data.get('theme', {})
        self.default_font = self.theme.get('font', 'Arial')

        self._tmp_template_path = None

        # Charger la présentation depuis un template distant si fourni
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

        # Forcer 16:9 pour cohérence (garde quand même le master)
        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(5.625)

        # Préparer les noms de layouts dispos (debug utile)
        try:
            self._layout_names = [getattr(l, "name", f"Layout {i}") for i, l in enumerate(self.prs.slide_layouts)]
            print("📐 Layouts disponibles:", self._layout_names)
        except Exception:
            self._layout_names = []

    # ---------- Sélection de layout ----------

    def _pick_layout(self, slide_type, hint=None):
        """
        Sélectionne une mise en page du template.
        - hint peut être:
          * int -> index du layout
          * str -> nom (contient)
          * dict -> {"name": "..."} ou {"index": 3}
        - fallback: mappe par type (cover/section/qcm/correction…)
        - dernier recours: Blank (index 6) ou index 1 (Title and Content)
        """
        layouts = self.prs.slide_layouts

        # 1) Si hint dict précis
        if isinstance(hint, dict):
            if "index" in hint and isinstance(hint["index"], int):
                idx = hint["index"]
                if 0 <= idx < len(layouts):
                    return layouts[idx]
            if "name" in hint and isinstance(hint["name"], str):
                name_part = hint["name"].strip().lower()
                for i in range(len(layouts)):
                    try:
                        if name_part in layouts[i].name.lower():
                            return layouts[i]
                    except Exception:
                        pass

        # 2) Si hint simple (str ou int)
        if isinstance(hint, int) and 0 <= hint < len(layouts):
            return layouts[hint]
        if isinstance(hint, str) and hint.strip():
            name_part = hint.strip().lower()
            for i in range(len(layouts)):
                try:
                    if name_part in layouts[i].name.lower():
                        return layouts[i]
                except Exception:
                    pass

        # 3) Mapping par type (essayons d'attraper tes masters)
        type_key = (slide_type or "").lower()
        preferred_names = []
        if type_key == "cover":
            preferred_names = ["cover", "couverture", "title", "titre"]
        elif type_key == "section":
            preferred_names = ["section", "separator", "chapitre"]
        elif type_key in ("qcm", "vrai_faux", "cas_pratique", "mise_en_situation", "objectifs"):
            preferred_names = ["content", "titre et contenu", "title and content", "contenu"]

        for part in preferred_names:
            for i in range(len(layouts)):
                try:
                    if part in layouts[i].name.lower():
                        return layouts[i]
                except Exception:
                    pass

        # 4) Fallbacks sûrs
        # Essayons "Title and Content" (souvent index 1)
        try:
            if len(layouts) > 1 and "title" in layouts[1].name.lower():
                return layouts[1]
        except Exception:
            pass

        # Sinon Blank (souvent index 6)
        try:
            if len(layouts) > 6 and "blank" in layouts[6].name.lower():
                return layouts[6]
        except Exception:
            pass

        # Dernier recours: premier layout
        return layouts[0]

    # ---------- Build principal ----------

    def build(self):
        print(f"🔨 Construction de {len(self.slides_data)} slides…")

        for i, slide_data in enumerate(self.slides_data):
            try:
                slide_type = slide_data.get('type', 'generic')
                layout_hint = slide_data.get('ppt_layout')  # NEW: string|int|dict
                layout = self._pick_layout(slide_type, layout_hint)
                print(f"  • Slide {i+1}/{len(self.slides_data)}: {slide_type} -> layout '{getattr(layout, 'name', '?')}'")

                # Ajout avec la mise en page choisie (=> héritage du master/template)
                slide = self.prs.slides.add_slide(layout)

                # Respect du template: on ne repeint PAS le fond si background est None/absent
                bg_color = self._safe_bg(slide_data)
                if bg_color is not None:  # si chaîne hex, on force la couleur
                    background = slide.background
                    fill = background.fill
                    fill.solid()
                    fill.fore_color.rgb = Colors.hex_to_rgb(bg_color)

                # Dispatcher contenu
                builder = {
                    'cover': self._fill_generic,
                    'section': self._fill_generic,
                    'qcm': self._fill_qcm,
                    'vrai_faux': self._fill_qcm,
                    'cas_pratique': self._fill_qcm,
                    'mise_en_situation': self._fill_qcm,
                    'objectifs': self._fill_correction,
                    'correction': self._fill_correction,
                }.get(slide_type, self._fill_generic)

                builder(slide, slide_data)

            except Exception as e:
                print(f"❌ Erreur slide {i+1}: {str(e)}")
                import traceback
                print(traceback.format_exc())
                self._build_error_slide_fallback(str(e))

        temp_path = os.path.join(tempfile.gettempdir(), 'presentation_zmforma.pptx')
        self.prs.save(temp_path)
        print(f"✅ Présentation sauvegardée: {temp_path}")

        if self._tmp_template_path and os.path.exists(self._tmp_template_path):
            try:
                os.remove(self._tmp_template_path)
            except:
                pass

        return temp_path

    def _safe_bg(self, data):
        """
        Retourne None si on veut garder le fond du template.
        Retourne une chaîne hex si on veut forcer un fond solide.
        """
        if 'background' not in data or data.get('background') in (None, "", "None", "null"):
            return None
        return data.get('background')

    # ---------- Fillers par type ----------

    def _fill_qcm(self, slide, data):
        layout_cfg = data.get('layout', {})

        if 'accent_bar' in layout_cfg:
            self._add_shape(slide, layout_cfg['accent_bar'])
        if 'kicker' in layout_cfg:
            self._add_textbox(slide, layout_cfg['kicker'])
        if 'title' in layout_cfg:
            self._add_textbox(slide, layout_cfg['title'])
        if 'question' in layout_cfg:
            self._add_textbox(slide, layout_cfg['question'])
        if 'context' in layout_cfg and layout_cfg['context']:
            self._add_textbox(slide, layout_cfg['context'])
        if 'choices' in layout_cfg:
            self._add_bullets(slide, layout_cfg['choices'])
        if 'image' in layout_cfg and layout_cfg['image']:
            self._add_image(slide, layout_cfg['image'])

    def _fill_correction(self, slide, data):
        layout_cfg = data.get('layout', {})
        for key in ['accent_bar', 'label', 'answer', 'answer_text', 'title', 'explanation', 'corrections', 'elements']:
            if key in layout_cfg and layout_cfg[key]:
                if isinstance(layout_cfg[key], dict) and 'items' in layout_cfg[key]:
                    self._add_bullets(slide, layout_cfg[key])
                elif isinstance(layout_cfg[key], dict):
                    self._add_textbox(slide, layout_cfg[key])

    def _fill_generic(self, slide, data):
        layout_cfg = data.get('layout', {})
        for key, element in layout_cfg.items():
            if element and isinstance(element, dict):
                if 'items' in element:
                    self._add_bullets(slide, element)
                elif 'text' in element:
                    self._add_textbox(slide, element)
                elif 'url' in element:
                    self._add_image(slide, element)
                elif 'fill' in element:
                    self._add_shape(slide, element)

    def _build_error_slide_fallback(self, error_msg):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 240, 240)

        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9.0), Inches(0.8))
        title_box.text = "❌ Erreur - Slide"
        Formatter.format_textbox(title_box, {'fontSize': 32, 'bold': True, 'color': 'D32F2F', 'align': 'center'})

        error_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9.0), Inches(2.0))
        error_box.text = str(error_msg)[:700]
        Formatter.format_textbox(error_box, {'fontSize': 14, 'color': '666666', 'align': 'left'})

    # ---------- Utilitaires d’ajout ----------

    def _add_textbox(self, slide, config):
        if not config or 'text' not in config:
            return
        x = Inches(config.get('x', 0.6))
        y = Inches(config.get('y', 0.9))
        w = Inches(config.get('w', 4.8))
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
