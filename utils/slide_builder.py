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
            /**
 * BUILDER DE SLIDES CALIBRÉ POUR TEMPLATE GAMMA
 * Positions optimisées pour le template (mesures en pouces, 16:9)
 */

const input = $input.first().json || {};
const exercices = Array.isArray(input.exercices) ? input.exercices : [];
const slides = [];
const theme = { font: "Arial" };

if (!exercices.length) {
  return [{ json: { error: true, message: "Aucun exercice", slides: [], theme } }];
}

const clip = (s, n = 700) => (typeof s === "string" && s.length > n ? s.slice(0, n-1) + "…" : (s || ""));
const A = (i) => String.fromCharCode(65 + i);

// ============== POSITIONS CALIBRÉES TEMPLATE GAMMA ==============
// Format 16:9 : 10" x 5.625"
// Zone de contenu : x=0.5" à 9.5", y=0.8" à 5.2"

function buildQCM(ex){
  const choices = Array.isArray(ex.choix) ? ex.choix.map((c,i)=> `${A(i)}) ${c}`) : [];

  // Slide question
  slides.push({
    type: "qcm",
    background: null, // Garde le fond du template
    layout: {
      kicker:   { text: (ex.module || "EXERCICE").toUpperCase(), x: 0.6, y: 0.6, w: 8.8, h: 0.35, fontSize: 11, color: "64748B", bold: true },
      question: { text: clip(ex.question || ex.titre || "Question"), x: 0.6, y: 1.1, w: 5.5, h: 1.2, fontSize: 26, bold: true, color: "1E293B" },
      context:  ex.contexte ? { text: clip(ex.contexte, 350), x: 0.6, y: 2.4, w: 5.5, h: 0.9, fontSize: 13, color: "475569" } : null,
      choices:  { items: choices, x: 0.6, y: ex.contexte ? 3.4 : 2.5, w: 5.5, h: 1.6, fontSize: 15, color: "0F172A" },
      image:    ex.image_url && ex.image_ready ? { url: ex.image_url, x: 6.4, y: 1.1, w: 2.9, h: 2.2 } : null
    }
  });

  // Slide correction
  if (typeof ex.reponse_correcte === "number") {
    const lettre = A(ex.reponse_correcte);
    const reponseTxt = Array.isArray(ex.choix) ? ex.choix[ex.reponse_correcte] : "";
    const explication = ex.consigne || ex.description || "";
    
    slides.push({
      type: "correction",
      background: null,
      layout: {
        label:       { text: "✅ CORRECTION", x: 0.6, y: 0.6, w: 8.8, h: 0.35, fontSize: 12, bold: true, color: "059669" },
        answer:      { text: `Réponse correcte : ${lettre}`, x: 0.6, y: 1.1, w: 8.8, h: 0.65, fontSize: 22, bold: true, color: "0F172A" },
        answer_text: { text: clip(reponseTxt, 150), x: 0.6, y: 1.85, w: 8.8, h: 0.55, fontSize: 15, color: "475569" },
        explanation: explication ? { text: clip(explication, 650), x: 0.6, y: 2.5, w: 8.8, h: 2.5, fontSize: 14, color: "1E293B" } : null
      }
    });
  }
}

function buildVraiFaux(ex){
  const items = Array.isArray(ex.affirmations)
    ? ex.affirmations.map((a,i)=> `${i+1}. ${clip(a.affirmation || a.texte || '')}`)
    : [];

  slides.push({
    type: "vrai_faux",
    background: null,
    layout: {
      kicker: { text: (ex.module || "VRAI/FAUX").toUpperCase(), x: 0.6, y: 0.6, w: 8.8, h: 0.35, fontSize: 11, color: "64748B", bold: true },
      title:  { text: clip(ex.titre || "Vrai ou Faux ?"), x: 0.6, y: 1.1, w: 8.8, h: 0.8, fontSize: 24, bold: true, color: "1E293B" },
      consigne: ex.consigne ? { text: clip(ex.consigne, 220), x: 0.6, y: 2.0, w: 8.8, h: 0.5, fontSize: 12, color: "64748B" } : null,
      items:  { items, x: 0.6, y: ex.consigne ? 2.6 : 2.1, w: 8.8, h: 2.7, fontSize: 14, bullet: true, color: "0F172A" }
    }
  });

  // Correction
  const corrBullets = Array.isArray(ex.affirmations)
    ? ex.affirmations.map((a, i) => {
        const tag = a.reponse === "VRAI" || a.correct ? "✅ VRAI" : "❌ FAUX";
        return `${i+1}. ${tag} — ${clip(a.justification || '', 240)}`;
      })
    : [];

  if (corrBullets.length) {
    slides.push({
      type: "correction",
      background: null,
      layout: {
        label:  { text: "✅ CORRECTION", x: 0.6, y: 0.6, w: 8.8, h: 0.35, fontSize: 12, bold: true, color: "059669" },
        title:  { text: clip(ex.titre || "Correction Vrai/Faux"), x: 0.6, y: 1.1, w: 8.8, h: 0.65, fontSize: 22, bold: true, color: "0F172A" },
        corrections: { items: corrBullets, x: 0.6, y: 1.9, w: 8.8, h: 3.2, fontSize: 13, bullet: true, color: "1E293B" }
      }
    });
  }
}

function buildCasPratique(ex){
  slides.push({
    type: "cas_pratique",
    background: null,
    layout: {
      kicker: { text: (ex.module || "CAS PRATIQUE").toUpperCase(), x: 0.6, y: 0.6, w: 8.8, h: 0.35, fontSize: 11, color: "64748B", bold: true },
      title:  { text: clip(ex.titre || "Cas pratique"), x: 0.6, y: 1.1, w: 8.8, h: 0.8, fontSize: 24, bold: true, color: "1E293B" },
      context:{ text: ex.contexte ? clip("📋 " + ex.contexte, 580) : "", x: 0.6, y: 2.0, w: 8.8, h: 1.3, fontSize: 13, color: "475569" },
      mission:{ text: ex.consigne ? clip("🎯 " + ex.consigne, 260) : "", x: 0.6, y: 3.4, w: 8.8, h: 0.8, fontSize: 14, bold: true, color: "0F172A" }
    }
  });
}

function buildMiseEnSituation(ex){
  slides.push({
    type: "mise_en_situation",
    background: null,
    layout: {
      kicker: { text: (ex.module || "MISE EN SITUATION").toUpperCase(), x: 0.6, y: 0.6, w: 8.8, h: 0.35, fontSize: 11, color: "64748B", bold: true },
      title:  { text: clip(ex.titre || "Mise en situation"), x: 0.6, y: 1.1, w: 8.8, h: 0.8, fontSize: 24, bold: true, color: "1E293B" },
      scenario: ex.description || ex.contexte ? { text: clip(ex.description || ex.contexte, 680), x: 0.6, y: 2.0, w: 8.8, h: 2.9, fontSize: 14, color: "0F172A" } : null
    }
  });
}

// ============== CONSTRUCTION ==============
for (const ex of exercices) {
  const t = (ex.type || "qcm").toLowerCase();
  if (t === "qcm") { buildQCM(ex); }
  else if (t === "vrai_faux") { buildVraiFaux(ex); }
  else if (t === "cas_pratique" || t === "cas_pratique_court") { buildCasPratique(ex); }
  else if (t === "mise_en_situation") { buildMiseEnSituation(ex); }
  else { buildQCM(ex); } // fallback
}

// ============== SORTIE ==============
const today = new Date().toISOString().split('T')[0];
const filename = `Formation_${input.code_rncp || 'RNCP'}_${today}.pptx`;

return [{
  json: {
    slides,
    theme,
    filename
  }
}];
