from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import tempfile
import os
import requests
import hashlib
from io import BytesIO
from PIL import Image
from .styles import Colors, Formatter

# Nom local sous lequel le template est mis en cache côté serveur
TEMPLATE_NAME = "template.pptx"


def _download_if_needed(url: str, dest_path: str, sha256: str = ""):
    """
    Télécharge le template si absent. Optionnellement vérifie l'empreinte SHA256.
    """
    if not url:
        return
    if os.path.exists(dest_path):
        return
    try:
        print(f"⬇️  Téléchargement du template depuis {url} ...")
        r = requests.get(url, timeout=60)
        r.raise_for_status()
        data = r.content
        if sha256:
            h = hashlib.sha256(data).hexdigest()
            if h.lower() != sha256.lower():
                raise RuntimeError(f"SHA256 mismatch (got {h}, expected {sha256})")
        with open(dest_path, "wb") as f:
            f.write(data)
        print(f"✅ Template téléchargé: {dest_path} ({len(data)//1024} Ko)")
    except Exception as e:
        print(f"⚠️ Impossible de télécharger le template: {e}")


class PresentationBuilder:
    """
    Constructeur de présentation PowerPoint
    - Charge un template téléchargé si TEMPLATE_URL est défini
    - Sinon, fallback sur une présentation vierge 1920×1080
    - Corrige l'ajout d'images (respect du ratio)
    - Active l'autofit texte via Formatter
    - Ignore les slides 'cover' & 'section' (comme demandé)
    """

    def __init__(self, data):
        self.data = data
        self.slides_data = data.get('slides', [])
        self.theme = data.get('theme', {})
        self.default_font = self.theme.get('font', 'Arial')

        # 1) Télécharger le template si besoin
        tpl_url = os.environ.get("TEMPLATE_URL", "").strip()
        tpl_sha = os.environ.get("TEMPLATE_SHA256", "").strip()  # optionnel
        tpl_path = os.path.join(os.getcwd(), TEMPLATE_NAME)

        force = os.environ.get("TEMPLATE_FORCE", "") == "1"
        if tpl_url and (force or not os.path.exists(tpl_path)):
            _download_if_needed(tpl_url, tpl_path, sha256=tpl_sha)

        # 2) Charger le template si disponible, sinon vierge 16:9
        if os.path.exists(tpl_path):
            self.prs = Presentation(tpl_path)
        else:
            self.prs = Presentation()
            self.prs.slide_width = Inches(10)      # 1920px
            self.prs.slide_height = Inches(5.625)  # 1080px

    def build(self):
        """Construit la présentation complète."""
        print(f"🔨 Construction de {len(self.slides_data)} slides (brutes)...")

        # 🔥 Comme demandé : ignorer cover & section
        filtered = []
        for sd in self.slides_data:
            t = (sd.get("type") or "").lower()
            if t in ("cover", "section"):
                continue
            filtered.append(sd)

        print(f"📉 Slides réellement générées (sans cover/section) : {len(filtered)}")

        for i, slide_data in enumerate(filtered):
            try:
                slide_type = (slide_data.get('type') or 'generic').lower()
                print(f"  • Slide {i+1}/{len(filtered)}: {slide_type}")

                if slide_type == 'qcm':
                    self._build_qcm(slide_data)
                elif slide_type == 'correction':
                    self._build_correction(slide_data)
                elif slide_type in ('vrai_faux', 'cas_pratique', 'mise_en_situation'):
                    # même structure que QCM
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

        # Sauvegarde
        temp_path = os.path.join(tempfile.gettempdir(), 'presentation_zmforma.pptx')
        self.prs.save(temp_path)
        print(f"✅ Présentation sauvegardée: {temp_path}")
        return temp_path

    # ----------------- BUILDERS -----------------

    def _build_qcm(self, data):
        """Slide QCM (structure moderne)."""
        layout = data.get('layout', {})
        bg_color = data.get('background', 'FFFFFF')

        slide = self._blank_with_bg(bg_color)

        # Barre d'accent
        if 'accent_bar' in layout:
            self._add_shape(slide, layout['accent_bar'])

        # Kicker
        if 'kicker' in layout:
            self._add_textbox(slide, layout['kicker'])

        # Question
        if 'question' in layout:
            self._add_textbox(slide, layout['question'])

        # Contexte
        if 'context' in layout and layout['context']:
            self._add_textbox(slide, layout['context'])

        # Choix (bullets)
        if 'choices' in layout:
            self._add_bullets(slide, layout['choices'])

        # Image
        if 'image' in layout and layout['image']:
            self._add_image(slide, layout['image'])

    def _build_correction(self, data):
        """Slide de correction / objectifs (même logique)."""
        layout = data.get('layout', {})
        bg_color = data.get('background', 'F0FDF4')

        slide = self._blank_with_bg(bg_color)

        if 'accent_bar' in layout:
            self._add_shape(slide, layout['accent_bar'])

        for key in ('label', 'answer', 'answer_text', 'title', 'explanation'):
            if key in layout and layout[key]:
                self._add_textbox(slide, layout[key])

        for key in ('corrections', 'elements', 'objectifs', 'conseils'):
            if key in layout and layout[key]:
                self._add_bullets(slide, layout[key])

    def _build_generic(self, data):
        """Slide générique (fallback)."""
        layout = data.get('layout', {})
        bg_color = data.get('background', 'FFFFFF')

        slide = self._blank_with_bg(bg_color)

        for _, element in layout.items():
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

    def _build_error_slide(self, slide_num, error_msg):
        slide = self._blank_with_bg_rgb(RGBColor(255, 240, 240))
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9.0), Inches(0.8))
        title_box.text = f"❌ Erreur - Slide {slide_num}"
        Formatter.format_textbox(title_box, {
            'fontSize': 32, 'bold': True, 'color': 'D32F2F', 'align': 'center'
        })
        error_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9.0), Inches(2.0))
        error_box.text = str(error_msg)[:500]
        Formatter.format_textbox(error_box, {'fontSize': 14, 'color': '666666', 'align': 'left'})

    # ----------------- UTILITAIRES -----------------

    def _blank_with_bg(self, hex_color):
        """Slide vierge avec fond couleur hex."""
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        try:
            bg = slide.background.fill
            bg.solid()
            bg.fore_color.rgb = Colors.hex_to_rgb(hex_color)
        except Exception:
            pass
        return slide

    def _blank_with_bg_rgb(self, rgb):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        try:
            bg = slide.background.fill
            bg.solid()
            bg.fore_color.rgb = rgb
        except Exception:
            pass
        return slide

    def _add_textbox(self, slide, config):
        if not config or 'text' not in config:
            return
        x = Inches(config.get('x', 0.5))
        y = Inches(config.get('y', 1.0))
        w = Inches(config.get('w', 5.0))
        h = Inches(config.get('h', 1.0))

        textbox = slide.shapes.add_textbox(x, y, w, h)
        textbox.text = str(config['text'])
        Formatter.format_textbox(textbox, config, self.default_font)

    def _add_bullets(self, slide, config):
        if not config or 'items' not in config:
            return
        items = config['items'] or []
        if not items:
            return

        x = Inches(config.get('x', 0.5))
        y = Inches(config.get('y', 1.0))
        w = Inches(config.get('w', 5.0))
        h = Inches(config.get('h', 2.0))

        textbox = slide.shapes.add_textbox(x, y, w, h)
        Formatter.add_bullet_points(textbox.text_frame, items, config, self.default_font)

    def _add_shape(self, slide, config):
        if not config:
            return
        x = Inches(config.get('x', 0))
        y = Inches(config.get('y', 0))
        w = Inches(config.get('w', 1))
        h = Inches(config.get('h', 0.1))
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
        if 'fill' in config:
            shape.fill.solid()
            shape.fill.fore_color.rgb = Colors.hex_to_rgb(config['fill'])
        shape.line.fill.background()

    def _add_image(self, slide, config):
        """
        Ajout d'image en respectant le ratio.
        - Si w & h sont fournis → on respecte le cadre mais on calcule width OU height
          pour conserver le ratio.
        - Si une seule dimension est fournie → l'autre est calculée automatiquement.
        """
        url = (config or {}).get('url')
        if not url:
            return

        x = Inches(config.get('x', 5.0))
        y = Inches(config.get('y', 1.0))
        w_in = config.get('w', 3.0)
        h_in = config.get('h', 2.0)

        try:
            img_stream = None
            if str(url).startswith('http'):
                resp = requests.get(url, timeout=15)
                resp.raise_for_status()
                img_stream = BytesIO(resp.content)
            elif os.path.exists(url):
                with open(url, 'rb') as f:
                    img_stream = BytesIO(f.read())

            if not img_stream:
                return

            # Ouvrir pour récupérer les dimensions réelles
            img_stream.seek(0)
            with Image.open(img_stream) as im:
                width_px, height_px = im.size
                ratio = width_px / float(height_px if height_px else 1)

            # Calcul ratio dans le cadre demandé
            target_w = float(w_in)
            target_h = float(h_in)
            if target_w <= 0 and target_h <= 0:
                target_w, target_h = 3.0, 2.0

            # Ajustement pour conserver le ratio sans déformation
            if (target_w / target_h) > ratio:
                # trop large → on limite sur la hauteur
                final_h = Inches(target_h)
                final_w = Inches(target_h * ratio)
            else:
                # trop haut → on limite sur la largeur
                final_w = Inches(target_w)
                final_h = Inches(target_w / ratio)

            img_stream.seek(0)
            slide.shapes.add_picture(img_stream, x, y, width=final_w, height=final_h)

        except Exception as e:
            print(f"⚠️ Impossible d'ajouter l'image {url}: {e}")
