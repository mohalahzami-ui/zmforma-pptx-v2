# utils/slide_builder.py
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import MSO_AUTO_SIZE
import tempfile, os, requests
from io import BytesIO
from PIL import Image
from .styles import Colors, Formatter

class PresentationBuilder:
    def __init__(self, data, template_url=None):
        self.data = data
        self.slides_data = data.get('slides', [])
        self.default_font = data.get('theme', {}).get('font', 'Arial')
        self.prs = self._load_template(template_url)
        self.prs.slide_width = Inches(10)
        self.prs.slide_height = Inches(5.625)

    def _load_template(self, template_url):
        """Charge le template local Formation.pptx ou crée un blanc."""
        # 1. Essayer de charger le template LOCAL en priorité
        local_template_path = os.path.join(
            os.path.dirname(__file__), 
            "Formation.pptx"
        )
        
        if os.path.exists(local_template_path):
            print(f"✅ Template trouvé : {local_template_path}")
            try:
                return Presentation(local_template_path)
            except Exception as e:
                print(f"⚠️ Erreur chargement template local : {e}")
        
        # 2. Sinon essayer l'URL (si fournie)
        if template_url:
            print(f"📥 Téléchargement template depuis : {template_url}")
            try:
                resp = requests.get(template_url, timeout=30)
                resp.raise_for_status()
                tmp = tempfile.mkstemp(suffix='.pptx')[1]
                with open(tmp, 'wb') as f:
                    f.write(resp.content)
                return Presentation(tmp)
            except Exception as e:
                print(f"❌ Erreur téléchargement template : {e}")
        
        # 3. Fallback : présentation vide
        print("⚠️ Aucun template disponible, création d'une présentation vierge")
        return Presentation()

    def build(self):
        for slide_data in self.slides_data:
            self._add_slide(slide_data)
        tmp_out = tempfile.mkstemp(suffix='.pptx')[1]
        self.prs.save(tmp_out)
        print(f"✅ PPTX généré : {tmp_out}")
        return tmp_out

    def _add_slide(self, slide_data):
        slide_type = slide_data.get('type','generic')
        hint = slide_data.get('ppt_layout')
        layout = self._pick_layout(hint, slide_type)
        slide = self.prs.slides.add_slide(layout)
        
        # Garder le fond du template sauf si background forcé
        bg = slide_data.get('background')
        if bg:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = Colors.hex_to_rgb(bg)
        
        # Remplissage
        for key, element in slide_data.get('layout', {}).items():
            if not element: continue
            if 'items' in element:
                self._add_bullets(slide, element)
            elif 'text' in element:
                self._add_text(slide, element)
            elif 'url' in element:
                self._add_image(slide, element)

    def _pick_layout(self, hint, slide_type):
        layouts = self.prs.slide_layouts
        
        # Chercher par nom partiel
        if isinstance(hint, str):
            key = hint.lower()
            for lay in layouts:
                if key in lay.name.lower():
                    return lay
        
        # Sinon "Title and Content" par défaut
        for lay in layouts:
            if "title" in lay.name.lower() and "content" in lay.name.lower():
                return lay
        
        # Fallback : premier layout
        return layouts[0] if len(layouts) > 0 else layouts[6]

    def _add_text(self, slide, cfg):
        x,y,w,h = Inches(cfg.get('x',0.5)), Inches(cfg.get('y',1.0)), Inches(cfg.get('w',5.0)), Inches(cfg.get('h',1.0))
        box = slide.shapes.add_textbox(x,y,w,h)
        box.text = str(cfg['text'])
        try:
            box.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        except:
            pass
        Formatter.format_textbox(box, cfg, self.default_font)

    def _add_bullets(self, slide, cfg):
        x,y,w,h = Inches(cfg.get('x',0.5)), Inches(cfg.get('y',1.0)), Inches(cfg.get('w',5.0)), Inches(cfg.get('h',2.0))
        tb = slide.shapes.add_textbox(x,y,w,h)
        try:
            tb.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        except:
            pass
        Formatter.add_bullet_points(tb.text_frame, cfg['items'], cfg, self.default_font)

    def _add_image(self, slide, cfg):
        url = cfg.get('url')
        x,y,w,h = cfg.get('x',5.6), cfg.get('y',1.0), cfg.get('w',3.6), cfg.get('h',2.25)
        if not url: 
            return
        
        try:
            if url.startswith('http'):
                resp = requests.get(url, timeout=15)
                resp.raise_for_status()
                raw = BytesIO(resp.content)
            else:
                raw = open(url, 'rb')
            
            img = Image.open(raw)
            W,H = img.size
            ratio = W/H if H else 1
            target_ratio = 16/9
            
            # Recadrage 16:9
            if abs(ratio - target_ratio) > 0.01:
                if ratio > target_ratio:
                    new_w = int(H * target_ratio)
                    x0 = (W - new_w) // 2
                    img = img.crop((x0, 0, x0 + new_w, H))
                else:
                    new_h = int(W / target_ratio)
                    y0 = (H - new_h) // 2
                    img = img.crop((0, y0, W, y0 + new_h))
            
            img = img.resize((1920, 1080))
            out = BytesIO()
            img.save(out, format='PNG')
            out.seek(0)
            slide.shapes.add_picture(out, Inches(x), Inches(y), width=Inches(w), height=Inches(h))
            print(f"✅ Image ajoutée : {url}")
            
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                print(f"⚠️ Image non disponible (404) : {url}")
            else:
                print(f"❌ Erreur HTTP image : {e}")
        except Exception as e:
            print(f"❌ Image KO : {e}")
