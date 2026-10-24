#!/usr/bin/env python3
"""
Genera el Excel maestro de MCR Agency a partir de las fuentes disponibles.

Jerarquía de resolución ante datos contradictorios:
  Nivel 1 · Interno  (archivos internos de MCR)
  Nivel 2 · Oficial  (tema web / sitio mcr-agency.com / comunicados propios)
  Nivel 3 · Externo  (prensa, directorios, referencias generales)
Gana el nivel más bajo con dato; dentro del mismo nivel gana la fuente más reciente.
Todo valor perdedor queda registrado en la hoja Conflictos.

Uso:
  python3 build_master.py --theme <tema_wp> --interno <carpeta_xlsx> --out <archivo.xlsx>

Este script no contiene datos de la agencia: todo se lee de las fuentes en tiempo de ejecución,
salvo la capa de investigación pública (research.py), que sólo cita información publicada.
"""
import argparse
import datetime as dt
import glob
import json
import os
import re
import subprocess
import tempfile
import unicodedata
from collections import Counter, defaultdict

import openpyxl
from bs4 import BeautifulSoup
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

import layout
import org
import research
import staff
import textfmt
from brands import BRAND_PARENTS, BRAND_RULES

VERSION = "1.0.0"
TODAY = dt.date.today()

# ─────────────────────────── utilidades ───────────────────────────

def strip_accents(s):
    s = unicodedata.normalize("NFKC", s)  # letras "estilizadas" (𝙋𝙀𝙎𝘼…) → ASCII
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def key(s):
    """Clave de matching: minúsculas, sin acentos ni símbolos."""
    if s is None:
        return ""
    return re.sub(r"[^a-z0-9]", "", strip_accents(str(s)).lower())


def slugify(s):
    return re.sub(r"[^a-z0-9]+", "", strip_accents(str(s)).lower()) or "sin-nombre"


def blank(v):
    return v is None or (isinstance(v, str) and v.strip().upper() in ("", "N/A", "NA", "-", "NONE"))


def parse_count(v):
    """Convierte conteos escritos a mano ("1.200.000", "80.6 mil", "23.8K", "1,100.000") a int.
    Devuelve (valor, nota) — nota explica conversiones ambiguas."""
    if blank(v):
        return None, None
    if isinstance(v, (int, float)):
        return int(round(v)), None
    s = str(v).strip().lower().replace(" ", "")
    mult = 1
    for suf, m in (("mil", 1_000), ("k", 1_000), ("m", 1_000_000)):
        if s.endswith(suf):
            s, mult = s[: -len(suf)], m
            break
    if mult > 1:
        try:
            return int(round(float(s.replace(",", ".")) * mult)), None
        except ValueError:
            return None, f"no interpretable: {v}"
    if re.fullmatch(r"\d+", s):
        return int(s), None
    if re.fullmatch(r"\d{1,3}([.,]\d{3})+", s):
        return int(re.sub(r"[.,]", "", s)), None
    if re.fullmatch(r"\d+[.,]\d+([.,]\d+)*", s):
        # mezcla de separadores o grupos raros ("2.900000", "1,100.000"): se asume entero sin decimales
        return int(re.sub(r"[.,]", "", s)), f"formato ambiguo '{v}' → interpretado como entero"
    return None, f"no interpretable: {v}"


def clean_url(v):
    """Primer URL del campo (algunos traen dos pegadas). Devuelve (url, nota)."""
    if blank(v):
        return None, None
    found = re.findall(r"(?:https?://)?(?:www\.|m\.)?[a-z0-9.-]+\.[a-z]{2,}/[^\s,]*", str(v).strip(), re.I)
    if not found:
        return str(v).strip(), None
    first = re.split(r"(?<=.)https?://", found[0])[0]
    url = first if first.lower().startswith("http") else "https://" + first
    note = "el campo trae más de una URL; se tomó la primera" if len(found) > 1 or first != found[0] else None
    return url, note


def tidy_url(u):
    """Quita parámetros de rastreo (?_t=, igsh=, utm_…, hl=, lang=) y fuerza https."""
    if not u:
        return u
    u = re.sub(r"^http://", "https://", str(u).strip())
    u = u.split("#")[0].split("?")[0]
    return u if u.lower().startswith("http") else "https://" + u


def extra_urls(v):
    """URLs adicionales distintas de la primera (p. ej. una segunda cuenta en el mismo campo)."""
    if blank(v):
        return []
    found = re.findall(r"(?:https?://)?(?:www\.|m\.)?[a-z0-9.-]+\.[a-z]{2,}/[^\s,]*", str(v), re.I)
    parts = [x for f in found for x in re.split(r"(?<=.)(?=https?://)", f)]
    out, seen = [], set()
    for u in parts:
        k = url_key(u).split("?")[0]
        if k and k not in seen:
            seen.add(k)
            out.append(u if u.lower().startswith("http") else "https://" + u)
    return out[1:]


def fix_email(v):
    """Corrige variantes obvias del dominio corporativo (mcr-agencycom, mcr.agency.com…)."""
    if blank(v):
        return None
    e = str(v).strip().lower().replace(" ", "")
    m = re.fullmatch(r"([^@]+)@(.+)", e)
    if not m:
        return e
    user, dom = m.groups()
    if re.fullmatch(r"mcr[-.]?agency[.]?(com|co|cm)?\.?", dom) or re.fullmatch(r"mcr[-.]agency\.com", dom):
        dom = "mcr-agency.com"
    return f"{user}@{dom}"


def url_key(v):
    s = str(v or "").lower().strip()
    s = re.sub(r"^https?://", "", s)
    s = re.sub(r"^(www\.|m\.)", "", s)
    s = s.split("?")[0].rstrip("/")
    return s


PHONE_CC = {"ECU": ("593", 9), "MEX": ("52", 10), "COL": ("57", 10), "PER": ("51", 9), "ESP": ("34", 9),
            "ARG": ("54", 10), "CHI": ("56", 9), "PAN": ("507", 8), "DOM": ("1", 10), "NIC": ("505", 8)}


def parse_phone(v, pais=None):
    """+código país cuando se puede deducir; si no, se deja como venía (sólo dígitos)."""
    if blank(v):
        return None, None
    if isinstance(v, float):
        v = str(int(v))
    digits = re.sub(r"\D", "", str(v))
    if not digits:
        return None, None
    cc, local_len = PHONE_CC.get(str(pais or "").strip().upper(), (None, None))
    if len(digits) >= 10 and not (cc and len(digits) == local_len and not digits.startswith(cc)):
        return "+" + digits, None
    if cc and len(digits) == local_len:
        return f"+{cc}{digits}", None
    return digits, None


COUNTRY = {
    "ESP": ("España", "ES"), "MEX": ("México", "MX"), "PER": ("Perú", "PE"), "URU": ("Uruguay", "UY"),
    "DOM": ("Rep. Dominicana", "DO"), "COL": ("Colombia", "CO"), "ECU": ("Ecuador", "EC"),
    "CHI": ("Chile", "CL"), "ARG": ("Argentina", "AR"), "PAN": ("Panamá", "PA"), "VEN": ("Venezuela", "VE"),
    "CAN": ("Canadá", "CA"), "NIC": ("Nicaragua", "NI"), "USA": ("EE. UU.", "US"), "CUBA": ("Cuba", "CU"),
    "BOL": ("Bolivia", "BO"), "SLV": ("El Salvador", "SV"), "BRA": ("Brasil", "BR"),
}
COUNTRY_BY_NAME = {key(v[0]): k for k, v in COUNTRY.items()}


def country_code(v):
    if blank(v):
        return None
    s = str(v).strip().upper()
    if s in COUNTRY:
        return s
    return COUNTRY_BY_NAME.get(key(v))


CATEGORY = {
    "GAMINGMINECRAFT": "Gaming · Minecraft", "GAMINGSHOOTING": "Gaming · Shooter", "GAMINGVARIETY": "Gaming · Variety",
    "GAMINGROBLOX": "Gaming · Roblox", "BAILE": "Baile", "BEAUTY": "Beauty", "VLOG": "Vlog",
    "LIFESTYLE&VLOG": "Lifestyle & Vlog", "COMEDY": "Comedia", "COMICS": "Cómics", "ANIME": "Anime",
    "ANIMACIÓN": "Animación", "ANIMACION": "Animación", "ENTRETENIMIENTO": "Entretenimiento",
    "COSPLAY": "Cosplay", "MUSIC REACTION": "Music reaction",
}


def norm_categories(v):
    if blank(v):
        return []
    out = []
    for part in re.split(r"[,/]", str(v)):
        p = part.strip().upper()
        if p:
            out.append(CATEGORY.get(p, part.strip().title()))
    return list(dict.fromkeys(out))


MONTHS = {m: i + 1 for i, m in enumerate(
    "ENERO FEBRERO MARZO ABRIL MAYO JUNIO JULIO AGOSTO SEPTIEMBRE OCTUBRE NOVIEMBRE DICIEMBRE".split())}
MONTHS["SETIEMBRE"] = 9
MONTH_NAMES = {v: k.capitalize() for k, v in MONTHS.items() if k != "SETIEMBRE"}


def parse_since_date(s):
    """Extrae la fecha de 'Sin campaña desde 07/23/2025' / '8 de mayo 2025' / '2023'."""
    if blank(s):
        return None
    s = str(s)
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        mm, dd, yy = map(int, m.groups())
        try:
            return dt.date(yy, mm, dd)
        except ValueError:
            return None
    m = re.search(r"(\d{1,2})\s+de\s+([a-záéíóú]+)(?:\s+(?:del?\s+)?(\d{4}))?", s, re.I)
    if m and strip_accents(m.group(2)).upper() in MONTHS:
        yy = int(m.group(3)) if m.group(3) else None
        if yy:
            return dt.date(yy, MONTHS[strip_accents(m.group(2)).upper()], int(m.group(1)))
    m = re.search(r"\b(20\d{2})\b", s)
    if m:
        return dt.date(int(m.group(1)), 1, 1)
    return None


# ─────────────────────────── fuentes ───────────────────────────

class Sources:
    """Registro de fuentes con su nivel jerárquico."""

    def __init__(self):
        self.rows = {}

    def add(self, sid, nombre, nivel, tipo, ubicacion, fecha=None, notas=None):
        self.rows[sid] = dict(id=sid, nombre=nombre, nivel=nivel, tipo=tipo, ubicacion=ubicacion,
                              fecha=fecha, notas=notas)
        return sid

    def level(self, sid):
        return self.rows[sid]["nivel"]

    def date(self, sid):
        return self.rows[sid]["fecha"] or dt.date(1900, 1, 1)


NIVEL_LABEL = {1: "1 · Interno", 2: "2 · Oficial", 3: "3 · Externo"}


class Resolver:
    """Elige el valor ganador por jerarquía y registra conflictos."""

    def __init__(self, sources):
        self.sources = sources
        self.conflicts = []

    def pick(self, entity_type, entity_id, field, candidates, numeric=False):
        cands = [(v, s) for v, s in candidates if not blank(v)]
        if not cands:
            return None, None
        cands.sort(key=lambda c: (self.sources.level(c[1]), -self.sources.date(c[1]).toordinal()))
        win_v, win_s = cands[0]
        for v, s in cands[1:]:
            if numeric:
                if not win_v and not v:
                    continue
                diff = abs(win_v - v) / max(abs(win_v), abs(v), 1)
                if diff < 0.005:
                    continue
                sev = "Alta" if diff >= 0.25 else "Media" if diff >= 0.10 else "Baja"
                dpct = round(diff * 100, 1)
            else:
                same = url_key(v) == url_key(win_v) if field.endswith("url") else key(v) == key(win_v)
                if same:
                    continue
                sev, dpct = "Revisar", None
            self.conflicts.append(dict(
                entidad=entity_type, entidad_id=entity_id, campo=field,
                valor_ganador=win_v, fuente_ganadora=win_s, nivel_ganador=NIVEL_LABEL[self.sources.level(win_s)],
                valor_descartado=v, fuente_descartada=s, nivel_descartado=NIVEL_LABEL[self.sources.level(s)],
                diferencia_pct=dpct, severidad=sev))
        return win_v, win_s


# ─────────────────────────── carga: tema web ───────────────────────────

def load_theme(theme_dir):
    exporter = os.path.join(os.path.dirname(os.path.abspath(__file__)), "export_theme_data.php")
    with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as tmp:
        out = tmp.name
    subprocess.run(["php", exporter, theme_dir, out], check=True, capture_output=True)
    with open(out, encoding="utf-8") as fh:
        data = json.load(fh)
    os.unlink(out)
    return data


def read_html(path):
    with open(path, encoding="utf-8") as fh:
        return BeautifulSoup(fh.read(), "html.parser")


def text(el):
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)) if el else None


def load_team(theme_dir):
    soup = read_html(os.path.join(theme_dir, "page-templates", "page-about.php"))
    team = []
    for section in soup.select(".about-team__section"):
        level = text(section.select_one(".about-team__subtitle"))
        for card in section.select("article.about-member"):
            m = re.search(r"TEAM_[A-Z0-9_]+\.webp", str(card))
            team.append(dict(nombre=text(card.select_one(".about-member__name")),
                             cargo=text(card.select_one(".about-member__role")), nivel=level,
                             foto=m.group(0) if m else None))
    return team


def load_portfolio(theme_dir):
    soup = read_html(os.path.join(theme_dir, "page-templates", "page-portfolio.php"))
    cases = []
    for art in soup.select("article.case"):
        tags = [text(t) for t in art.select(".case__tag")]
        imgs = sorted(set(re.findall(r"PORT_[A-Z0-9_]+\.webp", str(art))))
        highlights = [f"{text(h.select_one('.case__highlight-num'))} — {text(h.select_one('.case__highlight-label'))}"
                      for h in art.select(".case__highlight")]
        cases.append(dict(
            slug=art.get("id"), indice=text(art.select_one(".case__index")), titulo=text(art.select_one(".case__title")),
            subtitulo=text(art.select_one(".case__subtitle")), cliente=tags[0] if tags else None,
            servicios=tags[1] if len(tags) > 1 else None, region=tags[2] if len(tags) > 2 else None,
            descripcion=text(art.select_one(".case__desc")), kpis=" | ".join(highlights), assets=", ".join(imgs)))
    return cases


def load_press(theme_dir):
    soup = read_html(os.path.join(theme_dir, "page-templates", "page-press.php"))
    awards = []
    for a in soup.select("article.press-award"):
        awards.append(dict(tipo="Premio / reconocimiento", anio=text(a.select_one(".press-award__year")),
                           titulo=text(a.select_one(".press-award__title")), medio=text(a.select_one(".press-award__org")),
                           descripcion=text(a.select_one(".press-award__desc")), url=None))
    for a in soup.select("article.press-recognition"):
        kind = text(a.select_one(".press-recognition__type"))
        awards.append(dict(tipo=f"Presencia institucional · {kind}" if kind else "Presencia institucional",
                           anio=text(a.select_one(".press-recognition__year")),
                           titulo=text(a.select_one(".press-recognition__title")), medio=None,
                           descripcion=text(a.select_one(".press-recognition__desc")), url=None))
    articles = []
    for a in soup.select("a.press-article"):
        articles.append(dict(tipo="Prensa", anio=None, fecha=text(a.select_one(".press-article__date")),
                             titulo=text(a.select_one(".press-article__title")),
                             medio=text(a.select_one(".press-article__source")),
                             descripcion=text(a.select_one(".press-article__byline")), url=a.get("href")))
    return awards, articles


def load_legal(theme_dir):
    path = os.path.join(theme_dir, "template-parts", "legal", "aviso-legal.php")
    if not os.path.exists(path):
        return {}
    t = text(read_html(path))
    out = {}
    m = re.search(r"C\.?I\.?F\.?\s*([A-Z]\d{8})", t)
    if m:
        out["cif"] = m.group(1)
    m = re.search(r"domicilio social en (.+?), C\.?I\.?F", t)
    if m:
        out["domicilio"] = m.group(1).strip()
    m = re.search(r"inscrita en el (Registro Mercantil de [A-ZÁÉÍÓÚa-záéíóú ]+)", t)
    if m:
        out["registro"] = m.group(1).strip()
    return out


def load_contact(theme_dir):
    soup = read_html(os.path.join(theme_dir, "page-templates", "page-contact.php"))
    offices = []
    for o in soup.select(".contact-office"):
        parts = [x for x in (text(e) for e in o.find_all(["span", "p", "div", "h3", "h4"])) if x]
        parts = list(dict.fromkeys(p for p in parts if not re.fullmatch(r"\d{2}", p)))
        city = parts[0] if parts else None
        addr = next((p for p in parts[1:] if p != city), None)
        offices.append(dict(oficina=city, direccion=addr))
    page = text(soup)
    horario = re.search(r"Horario\s+(.+?\d{1,2}:\d{2}\s*[–-]\s*\d{1,2}:\d{2})", page)
    idiomas = re.search(r"Idiomas\s+([A-Z]{2}(?:\s*·\s*[A-Z]{2})+)", page)
    return dict(oficinas=offices, horario=horario.group(1) if horario else None,
                idiomas=idiomas.group(1) if idiomas else None)


def load_services(theme_dir):
    soup = read_html(os.path.join(theme_dir, "page-templates", "page-services.php"))
    rows = []
    for art in soup.select("article.service"):
        blocks = {}
        for b in art.select(".service__block"):
            label = text(b.select_one(".service__block-label"))
            items = [text(i) for i in b.select("li, .service__tag")]
            blocks[label] = ", ".join(items)
        rows.append(dict(num=text(art.select_one(".service__num")), servicio=text(art.select_one(".service__title")),
                         descripcion=text(art.select_one(".service__desc")), bloques=blocks,
                         ejemplo=re.sub(r"^Ejemplo ejecutado:\s*", "", text(art.select_one(".service__note")) or "") or None))
    for r in rows:
        r["b0"] = "; ".join(v for l, v in r["bloques"].items() if l.lower().startswith("alcance") and v) or None
        r["b1"] = "; ".join(v for l, v in r["bloques"].items() if not l.lower().startswith("alcance") and v) or None
    return rows, ["Alcance", "Formatos y tipos"]


def inventory_assets(theme_dir):
    try:
        from PIL import Image
    except ImportError:  # pragma: no cover
        Image = None
    code_text = ""
    for f in glob.glob(os.path.join(theme_dir, "**", "*.*"), recursive=True):
        if f.endswith((".php", ".js", ".css")):
            with open(f, encoding="utf-8", errors="ignore") as fh:
                code_text += f"\n@@FILE {os.path.relpath(f, theme_dir)}\n" + fh.read()
    files_by_usage = defaultdict(set)
    current = None
    for line in code_text.splitlines():
        if line.startswith("@@FILE "):
            current = line[7:]
            continue
        for m in re.findall(r"[A-Za-z0-9_\-]+\.(?:webp|svg|mp4|png|jpg|otf|ttf)", line):
            files_by_usage[m].add(current)
    rows = []
    for f in sorted(glob.glob(os.path.join(theme_dir, "**", "*"), recursive=True)):
        if os.path.isdir(f):
            continue
        rel = os.path.relpath(f, theme_dir)
        name = os.path.basename(f)
        ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        cat = asset_category(rel, name, ext)
        dims = None
        if Image and ext in ("webp", "png", "jpg", "jpeg"):
            try:
                with Image.open(f) as im:
                    dims = f"{im.width}×{im.height}"
            except Exception:
                dims = None
        used = sorted(u for u in files_by_usage.get(name, set()) if u and u != rel)
        rows.append(dict(archivo=name, ruta=rel, tipo=ext, categoria=cat, kb=round(os.path.getsize(f) / 1024, 1),
                         dimensiones=dims, usado_en=", ".join(used)))
    return rows


def asset_category(rel, name, ext):
    if rel.startswith("data/private-control"):
        return "Control privado (no exponer)"
    if ext in ("php",):
        return "Código · plantilla PHP"
    if ext in ("js", "css"):
        return f"Código · {ext.upper()}"
    if ext in ("otf", "ttf"):
        return "Tipografía"
    if re.match(r"(AAA|MID|RISING)_", name):
        return "Retrato creador"
    if name.startswith("TEAM_"):
        return "Foto equipo"
    if name.startswith("PRESS_LOGO"):
        return "Logo prensa"
    if name.endswith("LOGO.svg"):
        return "Logo marca"
    if name.startswith("PORT_"):
        return "Portafolio"
    if name.startswith("INST_"):
        return "Institucional"
    if name.startswith("DEMO_REEL"):
        return "Demo reel"
    if name.startswith(("logo-", "MCR_")) or "garabato" in name:
        return "Identidad MCR"
    if ext in ("webp", "png", "jpg", "svg"):
        return "Imagen"
    return "Otro"


# ─────────────────────────── carga: internos ───────────────────────────

INTERNAL_HEADERS = {
    "CREADOR": "handle", "CATEGORÍAS": "categorias", "INSTAGRAM": "instagram_url", "SEGUIDORES IG": "instagram_n",
    "TIKTOK": "tiktok_url", "SEGUIDORES TIKTOK": "tiktok_n", "YOUTUBE": "youtube_url", "SUSCRIPTORES YT": "youtube_n",
    "TWITCH": "twitch_url", "SUSCRIPTORES TWITCH": "twitch_n", "TWITTER / X": "x_url", "SEGUIDORES X": "x_n",
    "CORREO CORPORATIVO": "email", "TELÉFONO": "telefono", "NOMBRE DEL CREADOR": "nombre_real", "PAÍS": "pais",
    "ESTADO": "estado", "CAMPAÑAS HISTÓRICAS": "campanas_historicas", "CAMPAÑA VIGENTE": "campana_vigente",
    "CORREO CORPORATIVO EN REDES": "correo_en_redes", "ESTADO DE CAMPAÑA": "estado_campana",
}


def load_internal_creators(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    header_row, cols = None, {}
    for r in range(1, 15):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if "CREADOR" in [str(v).strip() if v else "" for v in vals]:
            header_row = r
            for i, v in enumerate(vals, 1):
                if v and str(v).strip() in INTERNAL_HEADERS:
                    cols[INTERNAL_HEADERS[str(v).strip()]] = i
            break
    if header_row is None:
        raise ValueError(f"No encontré encabezados en {path}")
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        handle = ws.cell(r, cols["handle"]).value
        if blank(handle):
            continue
        row = {k: ws.cell(r, c).value for k, c in cols.items()}
        row["handle"] = str(handle).strip()
        row["_row"] = r
        rows.append(row)
    return rows


def load_birthdays(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    month, out = None, []
    for row in ws.iter_rows(values_only=True):
        a, b = row[0], row[1] if len(row) > 1 else None
        c = row[2] if len(row) > 2 else None
        if a and strip_accents(str(a)).upper().strip() in MONTHS and not b:
            month = MONTHS[strip_accents(str(a)).upper().strip()]
            continue
        if not month or blank(b) or str(a).strip().lower() == "día":
            continue
        try:
            day = int(str(a).strip())
        except ValueError:
            continue
        out.append(dict(nombre=str(b).strip(), dia=day, mes=month, pais=c))
    return out


# ─────────────────────────── matching de creadores ───────────────────────────

# Alias conocidos entre fuentes (nombre en otra fuente → creator_id del tema web).
# Se revisan en la hoja Mapeo_IDs; "probable" significa que conviene confirmarlo.
ALIASES = {
    "gatocapitan": ("capitngato", "confirmado por país y cumpleaños"),
    "capitangato": ("capitngato", "confirmado"),
    "starcita": ("starvtuber", "probable: mismo país y correo star@"),
    "themagmarec": ("themagmaboi", "probable: variante de nombre en calendario"),
    "johanzuu": ("johaanzu", "typo en calendario"),
}


def pick_account(num_c, url_c, src):
    """Por plataforma: gana la fuente con más seguidores; si no hay cifras o empatan, la de mayor jerarquía.
    La URL se toma de la misma fuente ganadora."""
    rows = []
    for (n, s), (u, s2) in zip(num_c, url_c):
        rows.append((n, u, s))
    rows = [r for r in rows if r[0] is not None or r[1]]
    if not rows:
        return None, None, None, None
    rows.sort(key=lambda r: (-(r[0] or -1), src.level(r[2]), -src.date(r[2]).toordinal()))
    n, u, s = rows[0]
    if not u:
        u = next((r[1] for r in rows if r[1]), None)
    return n, s, u, s


def build_creator_index(web_creators):
    idx = {}
    for c in web_creators:
        cid = c["creator_id"]
        idx[key(c["identity"]["display_name"])] = cid
        idx[key(c["identity"]["slug"])] = cid
        idx[key(cid)] = cid
        email = (c.get("contact") or {}).get("corporate_email")
        if email:
            idx["@" + email.lower()] = cid
    return idx


def match_creator(name, email, idx):
    k = key(name)
    if email and "@" + str(email).strip().lower() in idx:
        return idx["@" + str(email).strip().lower()], "correo corporativo"
    if k in idx:
        return idx[k], "nombre normalizado"
    if k in ALIASES and ALIASES[k][0] in idx.values():
        return ALIASES[k][0], "alias: " + ALIASES[k][1]
    return None, None


# ─────────────────────────── campañas por creador ───────────────────────────

def classify_campaign(raw):
    up = strip_accents(raw).upper()
    for pattern, brand in BRAND_RULES:
        if re.search(pattern, up):
            return brand
    return None


def split_campaigns(handle, raw):
    if blank(raw):
        return []
    items = []
    for part in re.split(r",|\|", str(raw)):
        p = re.sub(r"\s+", " ", part).strip()
        if not p:
            continue
        up = strip_accents(p).upper()
        month = next((MONTHS[m] for m in MONTHS if re.search(rf"\b{m}\b", up)), None)
        year = re.search(r"\b(20\d{2})\b", up)
        brand = classify_campaign(p)
        items.append(dict(texto=p, marca=brand, mes=month, anio=int(year.group(1)) if year else None))
    return items


# ─────────────────────────── consolidación ───────────────────────────

PLATFORMS = ("youtube", "tiktok", "instagram", "twitch", "x", "facebook", "kick", "other")


def consolidate(theme, internals, birthdays, src, resolver, research_rows):
    web = theme["private"]["creators"]
    web_by_id = {c["creator_id"]: c for c in web}
    idx = build_creator_index(web)
    S_WEB = "S03"
    creators, platforms_rows, audience_rows, camp_rows, mapping, gaps = {}, [], [], [], [], []

    # 1 · creadores internos (nivel 1)
    for path, sid, rows in internals:
        for r in rows:
            cid, how = match_creator(r["handle"], r.get("email"), idx)
            if not cid:
                cid, how = slugify(r["handle"]), "nuevo (sólo interno)"
                if cid in creators:
                    cid = f"{cid}-{r['_row']}"
            creators.setdefault(cid, dict(creator_id=cid, internal=[], web=None, bday=None))
            creators[cid]["internal"].append((sid, r))
            mapping.append(dict(creator_id=cid, fuente=sid, nombre_en_fuente=r["handle"], metodo=how,
                                fila=r["_row"]))

    # 2 · creadores web (nivel 2)
    for c in web:
        cid = c["creator_id"]
        creators.setdefault(cid, dict(creator_id=cid, internal=[], web=None, bday=None))
        creators[cid]["web"] = c
        mapping.append(dict(creator_id=cid, fuente=S_WEB, nombre_en_fuente=c["identity"]["display_name"],
                            metodo="id del tema", fila=None))

    # 3 · cumpleaños (nivel 1)
    name_idx = {}
    for cid, c in creators.items():
        for _, r in c["internal"]:
            name_idx[key(r["handle"])] = cid
        if c["web"]:
            name_idx[key(c["web"]["identity"]["display_name"])] = cid
            name_idx[key(cid)] = cid
    for b in birthdays:
        k = key(b["nombre"])
        cid = name_idx.get(k)
        how = "nombre normalizado"
        if not cid and k in ALIASES:
            cid, how = ALIASES[k][0], "alias: " + ALIASES[k][1]
            cid = cid if cid in creators else None
        if not cid:
            gaps.append(dict(area="Cumpleaños", entidad=b["nombre"], problema="Cumpleaños sin creador asociado",
                             detalle=f"{b['dia']}/{b['mes']} · {b['pais']}", accion="Confirmar a quién corresponde"))
            continue
        if creators[cid]["bday"]:
            prev = creators[cid]["bday"]
            if (prev["dia"], prev["mes"]) != (b["dia"], b["mes"]):
                resolver.conflicts.append(dict(entidad="Creador", entidad_id=cid, campo="cumpleaños",
                                               valor_ganador=f"{prev['dia']}/{prev['mes']}", fuente_ganadora="S02",
                                               nivel_ganador=NIVEL_LABEL[1], valor_descartado=f"{b['dia']}/{b['mes']}",
                                               fuente_descartada="S02", nivel_descartado=NIVEL_LABEL[1],
                                               diferencia_pct=None, severidad="Revisar"))
            continue
        creators[cid]["bday"] = dict(b, fuente="S02")
        mapping.append(dict(creator_id=cid, fuente="S02", nombre_en_fuente=b["nombre"], metodo=how, fila=None))

    # 4 · resolver campos
    master = []
    for cid, c in sorted(creators.items(), key=lambda kv: kv[0]):
        w = c["web"]
        ints = c["internal"]
        wid = (w or {}).get("identity", {})
        wplat = {p["platform"]: p for p in (w or {}).get("platforms", [])}
        R = lambda field, cands, numeric=False: resolver.pick("Creador", cid, field, cands, numeric)

        name, name_src = R("nombre público", [(r["handle"], s) for s, r in ints] + [(wid.get("display_name"), S_WEB)])
        pais_code, pais_src = R("país", [(country_code(r.get("pais")), s) for s, r in ints]
                                + [(country_code(wid.get("country")), S_WEB)])
        cats_int = [", ".join(norm_categories(r.get("categorias"))) for s, r in ints]
        cats, cats_src = (cats_int[0], ints[0][0]) if ints and cats_int[0] else (
            ", ".join(wid.get("categories", [])) or None, S_WEB if wid.get("categories") else None)
        email, email_src = R("correo corporativo", [(fix_email(r.get("email")), s)
                                                     for s, r in ints]
                             + [((((w or {}).get("contact") or {}).get("corporate_email") or "").lower() or None, S_WEB)])
        tel, tel_note, tel_src = None, None, None
        for s, r in ints:
            tel, tel_note = parse_phone(r.get("telefono"), r.get("pais"))
            if tel:
                tel_src = s
                break
        real_name = next((str(r["nombre_real"]).strip() for s, r in ints if not blank(r.get("nombre_real"))), None)

        plat_out, extras = {}, []
        for p in ("youtube", "tiktok", "instagram", "twitch", "x"):
            num_c, url_c = [], []
            for s, r in ints:
                n, note = parse_count(r.get(f"{p}_n"))
                if note:
                    gaps.append(dict(area="Creadores", entidad=cid, problema=f"Conteo {p} con formato dudoso",
                                     detalle=note, accion="Corregir en CREADORES_INTERNOS"))
                num_c.append((n, s))
                u, unote = clean_url(r.get(f"{p}_url"))
                if unote:
                    gaps.append(dict(area="Creadores", entidad=cid, problema=f"URL {p} malformada",
                                     detalle=str(r.get(f"{p}_url"))[:120], accion="Corregir en CREADORES_INTERNOS"))
                url_c.append((u, s))
            wp = wplat.get(p)
            if wp:
                num_c.append((wp.get("followers"), S_WEB))
                url_c.append((wp.get("url"), S_WEB))
            n, n_src, u, u_src = pick_account(num_c, url_c, src)
            u = tidy_url(u)
            extras += [tidy_url(x) for s_, r in ints for x in extra_urls(r.get(f"{p}_url"))]
            avg = wp.get("avg_views") if wp else None
            plat_out[p] = dict(n=n, n_src=n_src, url=u, url_src=u_src, avg=avg)
            if n is not None or u:
                platforms_rows.append(dict(creator_id=cid, creador=name, plataforma=p, url=u, seguidores=n,
                                           fuente_seguidores=n_src, promedio_views=avg,
                                           fuente_views=S_WEB if avg is not None else None,
                                           principal=bool(wp and wp.get("primary")), fuente_url=u_src))
        for p, wp in wplat.items():  # plataformas que sólo existen en el tema (facebook, kick, other)
            if p in plat_out:
                continue
            platforms_rows.append(dict(creator_id=cid, creador=name, plataforma=p, url=wp.get("url"),
                                       seguidores=wp.get("followers"), fuente_seguidores=S_WEB,
                                       promedio_views=wp.get("avg_views"), fuente_views=S_WEB,
                                       principal=bool(wp.get("primary")), fuente_url=S_WEB))

        reach = sum(v["n"] or 0 for v in plat_out.values()) or None
        yt = plat_out["youtube"]
        yt_rate = round(yt["avg"] / yt["n"] * 100, 3) if yt["avg"] and yt["n"] else None
        primary = max(plat_out.items(), key=lambda kv: kv[1]["n"] or 0)[0] if reach else None

        # campañas
        hist_items = []
        for s, r in ints:
            for it in split_campaigns(r["handle"], r.get("campanas_historicas")):
                hist_items.append(it)
                camp_rows.append(dict(creator_id=cid, creador=name, campana=it["texto"], marca=it["marca"],
                                      empresa_matriz=BRAND_PARENTS.get(it["marca"]) if it["marca"] else None,
                                      mes=MONTH_NAMES.get(it["mes"]), anio=it["anio"], fuente=s))
        vigente = next((r.get("campana_vigente") for s, r in ints if not blank(r.get("campana_vigente"))), None)
        estado_c = next((r.get("estado_campana") for s, r in ints if not blank(r.get("estado_campana"))), None)
        ult = parse_since_date(vigente) if vigente and "sin" in str(vigente).lower() else None
        top_brands = Counter(i["marca"] for i in hist_items if i["marca"]).most_common(5)

        if ints and w:
            roster = "Interno + web"
        elif ints:
            roster = "Sólo interno (no está en la web)"
        else:
            roster = "Sólo web (no está en el roster interno · verificar baja)"

        bd = c["bday"]
        dq = (w or {}).get("data_quality", {})
        names = [r["handle"] for _, r in ints] + ([wid.get("display_name")] if w else []) + ([bd["nombre"]] if bd else [])
        variants = []
        for nm in names:
            if nm and key(nm) != key(name) and key(nm) not in {key(v) for v in variants}:
                variants.append(nm)
        row = dict(
            creator_id=cid, nombre_publico=name, variaciones=", ".join(variants) or None, fuente_nombre=name_src,
            nombre_real=real_name,
            estado_roster=roster, tier=wid.get("tier") or "Sin tier", fuente_tier=S_WEB if wid.get("tier") else None,
            pais=COUNTRY.get(pais_code, (None,))[0] if pais_code else None, iso=COUNTRY.get(pais_code, (None, None))[1] if pais_code else None,
            fuente_pais=pais_src, categorias=cats, fuente_categorias=cats_src,
            categorias_web=", ".join(wid.get("categories", [])) or None,
            cumple_dia=bd["dia"] if bd else None, cumple_mes=bd["mes"] if bd else None,
            correo=email, fuente_correo=email_src, telefono=tel, fuente_telefono=tel_src,
            correo_en_redes=next((r.get("correo_en_redes") for s, r in ints if not blank(r.get("correo_en_redes"))), None),
            **{f"{p}_seguidores": plat_out[p]["n"] for p in ("youtube", "tiktok", "instagram", "twitch", "x")},
            **{f"{p}_url": plat_out[p]["url"] for p in ("youtube", "tiktok", "instagram", "twitch", "x")},
            youtube_promedio_views=yt["avg"], youtube_view_rate_pct=yt_rate, otras_cuentas=", ".join(dict.fromkeys(extras)) or None,
            alcance_total=reach, plataforma_principal=primary,
            estado_campana=estado_c, campana_vigente=vigente, sin_campana_desde=ult,
            campanas_historicas_n=len(hist_items), marcas_frecuentes=", ".join(f"{b} ({n})" for b, n in top_brands) or None,
            bio=wid.get("bio"), retrato=wid.get("portrait_asset"),
            audiencia_estado=dq.get("audience_status"),
            media_kit=(w or {}).get("commercial_context", {}).get("media_kit_available"),
            stream=", ".join((w or {}).get("commercial_context", {}).get("stream_platforms", [])) or None,
            datos_web_generados=(w or {}).get("metadata", {}).get("profile_generated_at"),
        )
        master.append(row)

        # audiencia (sólo la trae el tema)
        for a in (w or {}).get("audience", []):
            base = dict(creator_id=cid, creador=name, plataforma=a.get("platform"), estado=a.get("quality_status"),
                        origen=a.get("source_type"), fuente=S_WEB)
            for dim, key_name in (("País", "countries"), ("Ciudad", "cities"), ("Idioma", "languages"),
                                  ("Interés", "interests")):
                for e in a.get(key_name) or []:
                    audience_rows.append(dict(base, dimension=dim, valor=e.get("name"), porcentaje=e.get("percentage"),
                                              calidad=(a.get("field_quality") or {}).get(key_name)))
            for e in a.get("gender_distribution") or []:
                audience_rows.append(dict(base, dimension="Género", valor=e.get("name") or e.get("gender"),
                                          porcentaje=e.get("percentage"),
                                          calidad=(a.get("field_quality") or {}).get("gender_distribution")))
            for e in a.get("age_distribution") or []:
                audience_rows.append(dict(base, dimension="Edad", valor=e.get("name") or e.get("range"),
                                          porcentaje=e.get("percentage"), calidad=None))

        # huecos
        if not row["retrato"] and w:
            gaps.append(dict(area="Assets", entidad=cid, problema="Sin retrato en la web", detalle=None,
                             accion="Producir o pedir foto"))
        if roster.startswith("Sólo web"):
            gaps.append(dict(area="Roster", entidad=cid, problema="Está en la web pero no en CREADORES_INTERNOS",
                             detalle=name, accion="Confirmar si sigue en MCR; si salió, retirarlo del sitio"))
        if roster.startswith("Sólo interno"):
            gaps.append(dict(area="Roster", entidad=cid, problema="Está en el roster interno pero no en la web",
                             detalle=name, accion="Decidir si se publica (tier, bio, retrato)"))
        if not bd:
            gaps.append(dict(area="Info personal", entidad=cid, problema="Sin cumpleaños", detalle=name,
                             accion="Pedir fecha al creador / account"))
        if email and not email.endswith("@mcr-agency.com"):
            gaps.append(dict(area="Contacto", entidad=cid, problema="Correo fuera de dominio @mcr-agency.com",
                             detalle=email, accion="Verificar / migrar a correo corporativo"))
        if tel_note:
            gaps.append(dict(area="Contacto", entidad=cid, problema="Teléfono", detalle=f"{tel} · {tel_note}",
                             accion="Completar con código de país"))
        if ints and not real_name:
            gaps.append(dict(area="Creadores", entidad=cid, problema="Sin nombre real", detalle=name,
                             accion="Completar en CREADORES_INTERNOS"))

    # nombres reales duplicados entre creadores distintos
    by_real = defaultdict(list)
    for m in master:
        if m["nombre_real"]:
            by_real[key(m["nombre_real"])].append(m["creator_id"])
    for k, ids in by_real.items():
        if len(ids) > 1:
            gaps.append(dict(area="Creadores", entidad=", ".join(ids), problema="Mismo nombre real en varios creadores",
                             detalle=next(m["nombre_real"] for m in master if m["creator_id"] == ids[0]),
                             accion="Probable error de captura: verificar"))
    return master, platforms_rows, audience_rows, camp_rows, mapping, gaps


# ─────────────────────────── escritura ───────────────────────────

SH = dict(idx="00_Índice", res="01_Resumen", agencia="02_Agencia", cifras="03_Cifras", crono="04_Cronología",
          servicios="05_Servicios", oficinas="06_Oficinas", equipo="07_Equipo", organigrama="08_Organigrama",
          p_staff="09_Personal_Staff", p_creadores="10_Personal_Creadores", creadores="11_Creadores",
          contratos="12_Contratos", rates="13_Rate_Card", plataformas="14_Plataformas", audiencia="15_Audiencia",
          campanas="16_Campañas_Creador", clientes="17_Clientes", registro="18_Campañas", marcas="19_Marcas",
          portafolio="20_Portafolio", prensa="21_Prensa_Premios")

DESCRIPTIONS = dict(res="Conteos generales", agencia="Ficha corporativa", cifras="Cifras actuales de la agencia",
                    crono="Hitos de 2019 a hoy", servicios="Servicios, alcance y formatos", oficinas="Oficinas",
                    equipo="Equipo actual, cuentas y contacto", organigrama="Estructura por cuenta y vertical",
                    p_staff="Cumpleaños, antigüedad y aniversarios del equipo",
                    p_creadores="Cumpleaños y antigüedad en MCR de los creadores", creadores="Roster completo",
                    contratos="Contratos de creadores", rates="Tarifas por creador y formato",
                    plataformas="Cuentas por plataforma", audiencia="Audiencia por país, idioma y género",
                    campanas="Historial de campañas por creador", clientes="Clientes y contactos",
                    registro="Registro de campañas", marcas="Marcas, clientes y empresa matriz",
                    portafolio="Casos publicados", prensa="Prensa, premios y presencia institucional")

def cols(*pairs):
    return list(pairs)


def shape(ctx):
    """Formato de textos y columnas finales de cada hoja."""
    handles = [c["nombre_publico"] for c in ctx["creators"]]
    brand_names = sorted({b for _, b in BRAND_RULES} | {"REDMAGIC", "Bloodstrike", "Gang Star", "GangStar"})
    keep = handles + brand_names
    for c in ctx["creators"]:
        c["nombre_real"] = textfmt.proper_name(c["nombre_real"])
        c["estado_campana"] = textfmt.sentence(c["estado_campana"])
        c["campana_vigente"] = textfmt.us_dates_to_es(textfmt.sentence(c["campana_vigente"], keep))
        c["plataforma_principal"] = textfmt.platform(c["plataforma_principal"])
        c["media_kit"] = textfmt.yes_no(c["media_kit"])
        c["correo_en_redes"] = textfmt.yes_no(c["correo_en_redes"])
        c["stream"] = ", ".join(textfmt.platform(x.strip()) for x in (c["stream"] or "").split(",") if x.strip()) or None
        if c.get("notas_internas"):
            c["notas_internas"] = textfmt.sentence(c["notas_internas"], keep)
    for p in ctx["platforms"]:
        p["plataforma"] = textfmt.platform(p["plataforma"])
        p["principal"] = textfmt.yes_no(p["principal"])
    for a in ctx["audience"]:
        a["plataforma"] = textfmt.platform(a["plataforma"])
    for c in ctx["campaigns_creator"]:
        own = [w.capitalize() if w.islower() else w for w in re.split(r"[_\W\d]+", c["creador"] or "") if len(w) > 2]
        c["campana"] = textfmt.sentence(c["campana"], own + keep)
    for t in ctx["team"]:
        t["nombre"] = textfmt.proper_name(t["nombre"])
    for r in ctx["org"]:
        r["persona"] = textfmt.proper_name(r["persona"])
    for r in ctx["staff_personal"]:
        r["nombre"] = textfmt.proper_name(r["nombre"])
    for r in ctx["creator_personal"]:
        r["nombre_real"] = textfmt.proper_name(r["nombre_real"])
    for p in ctx["press"]:
        p["fecha_txt"] = textfmt.es_date(p.get("fecha"))
        p["tipo"] = re.sub(r"^Presencia institucional · Presencia institucional$", "Presencia institucional", p["tipo"])
        p["tipo"] = p["tipo"].replace("Presencia institucional · ", "")
    ctx["press"].sort(key=lambda p: textfmt.date_sort_key(p["fecha_txt"]), reverse=True)
    for t in ctx["timeline"]:
        t["fecha_txt"] = textfmt.es_date(t["fecha"]) if t["fecha"] else str(t["anio"])
    ctx["timeline"].sort(key=lambda t: textfmt.date_sort_key(t["fecha_txt"]))


def current_figures(ctx):
    """Una cifra por métrica: la más actual (las que vive en el documento van como fórmula)."""
    cr, eq = f"'{SH['creadores']}'", f"'{SH['equipo']}'"
    countries = len({c["pais"] for c in ctx["creators"] if c["pais"]})
    return [
        dict(metrica="Años en el sector", valor="=YEAR(TODAY())-2019"),
        dict(metrica="Equipo", valor=f"=COUNTA({eq}!$A$5:$A$500)"),
        dict(metrica="Creadores en roster", valor=f"=COUNTA({cr}!$A$5:$A$2000)"),
        dict(metrica="Países con talento", valor=countries),
        dict(metrica="Red de creadores", valor="2K+"),
        dict(metrica="Marcas globales", valor="30+"),
        dict(metrica="Campañas ejecutadas", valor="1K+"),
        dict(metrica="Views en YouTube", valor="73B+"),
        dict(metrica="Facturación anual", valor="+5M USD"),
        dict(metrica="Oficinas", valor="4 (Barcelona, Madrid, Ciudad de México, São Paulo)"),
        dict(metrica="Stumble Guys · views acumuladas", valor="+1B"),
    ]


def write_book(out, ctx):
    shape(ctx)
    wb = layout.new_workbook()
    idx_ws = wb.create_sheet(SH["idx"])
    res = wb.create_sheet(SH["res"])
    entries = []

    def add(key_, title, columns, rows, **kw):
        ws = layout.table_sheet(wb, SH[key_], title, columns, rows, **kw)
        entries.append((SH[key_], DESCRIPTIONS[key_], len(rows)))
        return ws

    agency = list(ctx["agency"])
    legal, contact = ctx["legal"], ctx["contact"]
    pos = next((i for i, r in enumerate(agency) if r["campo"] == "Entidad legal"), 0) + 1
    extra = [dict(campo="CIF", valor=legal.get("cif")), dict(campo="Domicilio social", valor=legal.get("domicilio")),
             dict(campo="Registro", valor=legal.get("registro"))]
    agency[pos:pos] = [e for e in extra if e["valor"]]
    for campo, valor in (("Horario", contact.get("horario")), ("Idiomas", contact.get("idiomas"))):
        if valor:
            agency.append(dict(campo=campo, valor=valor))
    agency = [r for r in agency if r["campo"] != "Dirección Barcelona"]
    add("agencia", "Agencia", cols(("campo", "Campo"), ("valor", "Valor")), agency,
        widths=dict(campo=32, valor=100), wrap_cols=("valor",), freeze_cols=0)
    add("cifras", "Cifras", cols(("metrica", "Métrica"), ("valor", "Valor")), current_figures(ctx),
        widths=dict(metrica=36, valor=52), int_cols=("valor",), freeze_cols=0)
    add("crono", "Cronología", cols(("fecha_txt", "Fecha"), ("hito", "Hito"), ("categoria", "Categoría")),
        ctx["timeline"], widths=dict(fecha_txt=14, hito=96, categoria=16), wrap_cols=("hito",), freeze_cols=0)
    srv_rows, srv_labels = ctx["services"]
    add("servicios", "Servicios",
        cols(("num", "N.º"), ("servicio", "Servicio"), ("descripcion", "Descripción"),
             *[(f"b{i}", textfmt.sentence(l)) for i, l in enumerate(srv_labels)], ("ejemplo", "Ejemplo ejecutado")),
        srv_rows, widths=dict(num=6, servicio=30, descripcion=80, ejemplo=50,
                              **{f"b{i}": 44 for i in range(len(srv_labels))}),
        wrap_cols=("descripcion", "ejemplo") + tuple(f"b{i}" for i in range(len(srv_labels))), freeze_cols=0)
    office_rows = []
    for o in ctx["contact"]["oficinas"]:
        country = {"Barcelona": "España", "Madrid": "España", "Ciudad de México": "México", "São Paulo": "Brasil"}.get(o["oficina"])
        addr = o["direccion"]
        if addr and addr.replace(",", "").strip() in (f"{o['oficina']} {country}", o["oficina"]):
            addr = None
        office_rows.append(dict(oficina=o["oficina"], pais=country, direccion=addr))
    add("oficinas", "Oficinas",
        cols(("oficina", "Oficina"), ("pais", "País"), ("direccion", "Dirección"), ("responsable", "Responsable"),
             ("telefono", "Teléfono"), ("correo", "Correo")),
        office_rows, widths=dict(oficina=20, pais=14, direccion=56, responsable=26, telefono=18, correo=28), freeze_cols=0)
    add("equipo", "Equipo",
        cols(("person_id", "ID"), ("nombre", "Nombre"), ("cargo", "Cargo"), ("area", "Área"), ("cuentas", "Cuentas"),
             ("roles_cuenta", "Rol en la cuenta"), ("lidera", "Lidera"), ("modalidad", "Modalidad"),
             ("junta", "Junta directiva"), ("correo", "Correo"), ("telefono", "Teléfono"), ("oficina", "Oficina"),
             ("pais", "País"), ("reporta_a", "Reporta a")),
        ctx["team"], widths=dict(person_id=8, nombre=24, cargo=34, area=22, cuentas=40, roles_cuenta=36, lidera=16,
                                 modalidad=16, junta=16, correo=28, telefono=16, oficina=18, pais=12, reporta_a=22),
        lists=dict(oficina=["Barcelona", "Madrid", "Ciudad de México", "São Paulo", "Remoto"]))
    add("organigrama", "Organigrama",
        cols(("bloque", "Bloque"), ("cuenta", "Cuenta / área"), ("vertical", "Vertical"), ("person_id", "ID"),
             ("persona", "Persona"), ("rol", "Rol"), ("lidera", "Lidera"), ("modalidad", "Modalidad"),
             ("detalle", "Detalle")),
        ctx["org"], widths=dict(bloque=20, cuenta=20, vertical=30, person_id=8, persona=24, rol=34, lidera=10,
                                modalidad=14, detalle=40), center_cols=("lidera",), freeze_cols=0)
    for key_, title, columns, rows in (("p_staff", "Información personal · staff", staff.STAFF_PERSONAL_COLS,
                                        ctx["staff_personal"]),
                                       ("p_creadores", "Información personal · creadores", staff.CREATOR_PERSONAL_COLS,
                                        ctx["creator_personal"])):
        ws = add(key_, title, [(f, lab) for f, lab, _ in columns], rows,
                 widths=dict(person_id=14, nombre=26, nombre_real=28, cargo=32, notas=36, antiguedad=16,
                             precision=16, proximo_cumple=14, proximo_aniversario=14, fecha_ingreso=14))
        staff.finish_personal_sheet(ws, columns, len(rows),
                                    signing_lookup=(SH["contratos"], "A", "D") if key_ == "p_creadores" else None)
    add("creadores", "Creadores", ctx["creator_cols"], ctx["creators"],
        widths=dict(creator_id=16, nombre_publico=20, variaciones=22, nombre_real=28, categorias=26, bio=70,
                    campana_vigente=36, marcas_frecuentes=40, correo=30, otras_cuentas=30, notas_internas=36,
                    youtube_url=30, tiktok_url=30, instagram_url=30, twitch_url=26, x_url=24, kick_url=24),
        link_cols=("youtube_url", "tiktok_url", "instagram_url", "twitch_url", "x_url", "kick_url", "otras_cuentas"),
        int_cols=tuple(f"{p}_seguidores" for p in ("youtube", "tiktok", "instagram", "twitch", "x", "kick"))
        + ("alcance_total", "youtube_promedio_views"),
        pct_cols=("youtube_view_rate_pct",), center_cols=("tier", "en_web", "cumple_dia", "cumple_mes", "media_kit"))
    base = [dict(creator_id=c["creator_id"], creador=c["nombre_publico"], tier=c["tier"],
                 plataforma_principal=c["plataforma_principal"]) for c in ctx["creators"]]
    add("contratos", "Contratos",
        cols(("creator_id", "ID"), ("creador", "Creador"), ("tier", "Tier"), ("firma", "Fecha de firma"),
             ("exclusividad", "Exclusividad"), ("vencimiento", "Vencimiento"), ("comision", "Comisión %"),
             ("territorio", "Territorio"), ("manager", "Influencer manager"), ("estatus", "Estatus"), ("notas", "Notas")),
        base, widths=dict(creator_id=16, creador=20, firma=14, vencimiento=14, territorio=20, manager=24, notas=36),
        date_cols=("firma", "vencimiento"), center_cols=("tier", "exclusividad"),
        lists=dict(exclusividad=["Sí", "No"], estatus=["Activo", "En renovación", "Finalizado"]))
    add("rates", "Rate card",
        cols(("creator_id", "ID"), ("creador", "Creador"), ("tier", "Tier"), ("plataforma_principal", "Plataforma principal"),
             ("yt_integracion", "YouTube integración"), ("yt_dedicado", "YouTube dedicado"), ("yt_short", "YouTube Short"),
             ("tiktok", "TikTok"), ("ig_reel", "Instagram Reel"), ("ig_story", "Instagram Story"),
             ("ig_post", "Instagram post"), ("stream_hora", "Stream (por hora)"), ("moneda", "Moneda"),
             ("vigencia", "Vigencia"), ("notas", "Notas")),
        base, widths=dict(creator_id=16, creador=20, notas=36), date_cols=("vigencia",), center_cols=("tier", "moneda"),
        lists=dict(moneda=["USD", "MXN", "EUR", "BRL", "COP"]))
    add("plataformas", "Plataformas",
        cols(("creator_id", "ID"), ("creador", "Creador"), ("plataforma", "Plataforma"), ("url", "Cuenta"),
             ("seguidores", "Seguidores"), ("promedio_views", "Promedio views"), ("principal", "Principal")),
        ctx["platforms"], link_cols=("url",), int_cols=("seguidores", "promedio_views"),
        widths=dict(creator_id=16, creador=20, url=40), center_cols=("principal",))
    add("audiencia", "Audiencia",
        cols(("creator_id", "ID"), ("creador", "Creador"), ("plataforma", "Plataforma"), ("dimension", "Dimensión"),
             ("valor", "Valor"), ("porcentaje", "%")),
        ctx["audience"], pct_cols=("porcentaje",), widths=dict(creator_id=16, creador=20, valor=24))
    add("campanas", "Campañas por creador",
        cols(("creator_id", "ID"), ("creador", "Creador"), ("campana", "Campaña"), ("marca", "Marca / título"),
             ("empresa_matriz", "Empresa matriz"), ("mes", "Mes"), ("anio", "Año")),
        ctx["campaigns_creator"], widths=dict(creator_id=16, creador=20, campana=56, marca=28, empresa_matriz=30))
    add("clientes", "Clientes",
        cols(("client_id", "ID"), ("cliente", "Cliente"), ("empresa_matriz", "Empresa matriz"), ("cuenta", "Cuenta MCR"),
             ("lead", "Lead MCR"), ("exclusiva", "Agencia exclusiva"), ("estatus", "Estatus"), ("contacto", "Contacto"),
             ("cargo", "Cargo"), ("correo", "Correo"), ("telefono", "Teléfono"), ("pais", "País"), ("desde", "Desde"),
             ("notas", "Notas")),
        ctx["clients"], widths=dict(client_id=8, cliente=22, empresa_matriz=28, cuenta=18, lead=24, contacto=26,
                                    cargo=22, correo=28, notas=36),
        date_cols=("desde",), center_cols=("exclusiva",), empty_rows=15, freeze_cols=0,
        lists=dict(estatus=["Activo", "Pitch", "Pausado", "Perdido"], exclusiva=["Sí", "No"]))
    add("registro", "Campañas",
        cols(("campaign_id", "ID"), ("campana", "Campaña"), ("cliente", "Cliente"), ("marca", "Marca / título"),
             ("cuenta", "Cuenta MCR"), ("inicio", "Inicio"), ("fin", "Fin"), ("region", "Región"), ("creadores", "Creadores"),
             ("entregables", "Entregables"), ("presupuesto", "Presupuesto"), ("fee", "Fee MCR"), ("moneda", "Moneda"),
             ("resultados", "Resultados / KPIs"), ("estatus", "Estatus"), ("responsable", "Responsable"), ("notas", "Notas")),
        [], widths=dict(campana=34, cliente=20, marca=24, creadores=36, entregables=36, resultados=40, responsable=22,
                        notas=36), date_cols=("inicio", "fin"), empty_rows=40, freeze_cols=1,
        lists=dict(estatus=["Propuesta", "Aprobada", "En curso", "Finalizada", "Cancelada"],
                   moneda=["USD", "MXN", "EUR", "BRL", "COP"]))
    add("marcas", "Marcas",
        cols(("brand_id", "ID"), ("marca", "Marca / título"), ("empresa_matriz", "Empresa matriz"),
             ("en_web", "Logo en web"), ("exclusiva", "Agencia exclusiva"), ("campanas_n", "Campañas"),
             ("creadores_n", "Creadores"), ("casos_portafolio", "Casos en portafolio")),
        ctx["brands"], widths=dict(brand_id=8, marca=30, empresa_matriz=32, casos_portafolio=44),
        int_cols=("campanas_n", "creadores_n"), center_cols=("en_web", "exclusiva"), freeze_cols=0)
    add("portafolio", "Portafolio",
        cols(("case_id", "ID"), ("titulo", "Caso"), ("subtitulo", "Subtítulo"), ("cliente", "Cliente"),
             ("servicios", "Servicios"), ("region", "Región"), ("kpis", "KPIs"), ("descripcion", "Descripción")),
        ctx["cases"], widths=dict(case_id=8, titulo=32, subtitulo=30, servicios=28, kpis=54, descripcion=80),
        wrap_cols=("kpis", "descripcion", "subtitulo"))
    add("prensa", "Prensa y premios",
        cols(("tipo", "Tipo"), ("fecha_txt", "Fecha"), ("titulo", "Título"), ("medio", "Medio / organización"),
             ("descripcion", "Descripción"), ("url", "Enlace")),
        ctx["press"], widths=dict(tipo=24, fecha_txt=16, titulo=70, medio=30, descripcion=60, url=12),
        link_cols=("url",), link_text="Abrir", wrap_cols=("titulo", "descripcion"), freeze_cols=0)

    write_summary(res, ctx)
    staff.protect(res)
    layout.index_sheet(idx_ws, [(SH["res"], DESCRIPTIONS["res"], None)] + entries, TODAY.strftime("%d/%m/%Y"))
    wb.save(out)


def write_summary(ws, ctx):
    cr = f"'{SH['creadores']}'"
    ncol = {f: get_column_letter(i) for i, (f, _) in enumerate(ctx["creator_cols"], 1)}
    rng = lambda f: f"{cr}!${ncol[f]}$5:${ncol[f]}$2000"
    ps, pc = f"'{SH['p_staff']}'", f"'{SH['p_creadores']}'"
    blocks = [
        ("Creadores", [
            ("Total en roster", f"=COUNTA({rng('creator_id')})"),
            ("Publicados en la web", f'=COUNTIF({rng("en_web")},"Sí")'),
            ("Sin publicar en la web", f'=COUNTIF({rng("en_web")},"No")'),
            ("Con campaña activa", f'=COUNTIF({rng("estado_campana")},"Con campaña")'),
            ("Alcance total sumado", f"=SUM({rng('alcance_total')})"),
        ]),
        ("Equipo", [
            ("Personas", f"=COUNTA('{SH['equipo']}'!$A$5:$A$500)"),
            ("Cumpleaños en 30 días · staff", f'=COUNTIFS({ps}!$J$5:$J$500,">=0",{ps}!$J$5:$J$500,"<=30")'),
            ("Cumpleaños en 30 días · creadores", f'=COUNTIFS({pc}!$L$5:$L$500,">=0",{pc}!$L$5:$L$500,"<=30")'),
        ]),
        ("Por tier", [(t, f'=COUNTIF({rng("tier")},"{t}")') for t in ("AAA", "Mid", "Rising")]),
        ("Negocio", [
            ("Marcas", f"=COUNTA('{SH['marcas']}'!$A$5:$A$1000)"),
            ("Casos de portafolio", f"=COUNTA('{SH['portafolio']}'!$A$5:$A$500)"),
            ("Campañas registradas", f"=COUNTA('{SH['campanas']}'!$A$5:$A$10000)"),
        ]),
        ("Por país", [(p, f'=COUNTIF({rng("pais")},"{p}")')
                      for p, _ in Counter(c["pais"] for c in ctx["creators"] if c["pais"]).most_common()]),
    ]
    layout.summary_sheet(ws, blocks)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--theme", required=True, help="carpeta del tema WP descomprimido")
    ap.add_argument("--interno", required=True, help="carpeta con los .xlsx internos")
    ap.add_argument("--out", required=True)
    ap.add_argument("--previo", help="maestro anterior: conserva lo capturado a mano en las celdas amarillas")
    args = ap.parse_args()

    src = Sources()
    resolver = Resolver(src)

    # internos
    internals = []
    files = sorted(glob.glob(os.path.join(args.interno, "*.xlsx")))
    roster_files = [f for f in files if re.search(r"creadores|roster", os.path.basename(f), re.I)]
    bday_files = [f for f in files if re.search(r"cumplea", os.path.basename(f), re.I)]
    for n, f in enumerate(roster_files, 1):
        sid = src.add(f"S01{'' if n == 1 else chr(96 + n)}", "Roster interno · CREADORES_INTERNOS_CONSOLIDADO", 1,
                      "Excel interno", os.path.basename(f), dt.date.fromtimestamp(os.path.getmtime(f)),
                      "Sobreescribe a la web en todo campo que tenga.")
        internals.append((f, sid, load_internal_creators(f)))
    birthdays = []
    for f in bday_files:
        src.add("S02", "Calendario de cumpleaños de creadores", 1, "Excel interno", os.path.basename(f),
                dt.date.fromtimestamp(os.path.getmtime(f)), "Día y mes; sin año de nacimiento.")
        birthdays += load_birthdays(f)

    theme = load_theme(args.theme)
    gen = theme["private"].get("generated_at")
    src.add("S03", "Tema web MCR v2.5.0 · dataset de creadores", 2, "Tema WordPress (zip)",
            "data/creators-private.php", dt.date.fromisoformat(gen) if gen else None,
            "Generado desde MCR_CREATOR_DATA_v3_0_0.xlsx. Tokens de perfiles privados excluidos a propósito.")
    src.add("S04", "Tema web MCR v2.5.0 · páginas (about, portfolio, press)", 2, "Tema WordPress (zip)",
            "page-templates/", dt.date.fromisoformat(gen) if gen else None, None)
    research.register_sources(src)

    master, plats, aud, camps, mapping, gaps = consolidate(theme, internals, birthdays, src, resolver, None)

    # ajustes internos (hojas Staff / Ajustes_Creadores) y maestro previo
    adj_files = [f for f in files if f not in roster_files and f not in bday_files]
    staff_raw, adjustments, nicknames = staff.load_internal_adjustments(adj_files)
    staff_sources = []
    for n, (f, rows) in enumerate(staff_raw, 1):
        sid = src.add(f"S3{n}", f"Staff interno · {os.path.basename(f)}", 1, "Excel interno", os.path.basename(f),
                      dt.date.fromtimestamp(os.path.getmtime(f)), "Altas, bajas, cumpleaños e ingreso del equipo.")
        staff_sources.append((sid, rows))
    for f in sorted({a["fuente"] for a in adjustments}):
        src.add("S39", f"Ajustes internos de creadores · {f}", 1, "Excel interno", f,
                dt.date.fromtimestamp(os.path.getmtime(os.path.join(args.interno, f))), None)
    prev_staff, prev_creators = staff.load_previous_master(args.previo)
    if args.previo:
        src.add("S40", "Captura manual en el maestro previo", 1, "Excel maestro", os.path.basename(args.previo),
                dt.date.fromtimestamp(os.path.getmtime(args.previo)), "Celdas amarillas de las hojas de información personal.")

    # creadores que sólo están en la web: se asume que ya no están en MCR
    gone = {m["creator_id"] for m in master if m["estado_roster"].startswith("Sólo web")}
    master = [m for m in master if m["creator_id"] not in gone]
    plats = [x for x in plats if x["creator_id"] not in gone]
    aud = [x for x in aud if x["creator_id"] not in gone]
    camps = [x for x in camps if x["creator_id"] not in gone]
    for m in master:
        m["en_web"] = "Sí" if m["estado_roster"] == "Interno + web" else "No"
        if m["tier"] == "Sin tier":
            m["tier"] = tier_for_reach(m["alcance_total"])
    apply_creator_adjustments(master, adjustments, resolver, gaps)
    fix_platform_labels(master, plats, gaps)
    flag_suspicious_counts(master, gaps)

    # equipo
    team_raw = load_team(args.theme)
    team = []
    used_photos = set()
    for i, t in enumerate(team_raw, 1):
        used_photos.add(t["foto"])
        team.append(dict(person_id=f"P{i:03d}", nombre=t["nombre"], cargo=t["cargo"], area=role_area(t["cargo"]),
                         nivel_org=t["nivel"], estado="Activo", foto=t["foto"], fuente="S04",
                         notas=None if t["foto"] else "Sin foto en la web"))
    all_team_photos = sorted(os.path.basename(f) for f in glob.glob(os.path.join(args.theme, "assets/img/TEAM_*")))
    for ph in all_team_photos:
        if ph not in used_photos:
            nm = ph.replace("TEAM_", "").rsplit(".", 1)[0].title()
            team.append(dict(person_id=f"P{len(team) + 1:03d}", nombre=nm, cargo=None, area=None, nivel_org=None,
                             estado=None, foto=ph, fuente="S04",
                             notas="Foto existe en assets pero no aparece en la página Nosotros"))
    team = staff.merge_staff(team, staff_sources, prev_staff, src, resolver, role_area)
    for t in team:
        if t["estado"] == "Por confirmar":
            gaps.append(dict(area="Equipo", entidad=t["nombre"], problema="Estado por confirmar", detalle=t.get("foto"),
                             accion="Confirmar si sigue en MCR"))
        if t["estado"] == "Ex-equipo" and t.get("foto"):
            gaps.append(dict(area="Web", entidad=t["nombre"], problema="Foto de ex-equipo sigue en los assets del tema",
                             detalle=t["foto"], accion="Borrarla en la próxima versión del tema"))
        if t["estado"] == "Ex-equipo":
            continue
        if not t["foto"]:
            gaps.append(dict(area="Equipo", entidad=t["nombre"], problema="Sin foto de equipo",
                             detalle=t["cargo"], accion="Producir foto (Kevin Moss)"))
        if t["fuente"] != "S04":
            gaps.append(dict(area="Web", entidad=t["nombre"], problema="Está en el equipo interno pero no en la web",
                             detalle=t["cargo"], accion="Agregarlo a la página Nosotros"))
        missing = [lab for f, lab in (("cumple_dia", "cumpleaños"), ("ing_anio", "fecha de ingreso")) if not t.get(f)]
        if missing:
            gaps.append(dict(area="Info personal", entidad=t["nombre"], problema="Staff sin " + " ni ".join(missing),
                             detalle=t["cargo"], accion="Plantilla RH"))
        elif not t.get("ing_mes"):
            gaps.append(dict(area="Info personal", entidad=t["nombre"], problema="Ingreso sólo con año",
                             detalle=f"{t['ing_anio']}", accion="Plantilla RH: mes y día para calcular el aniversario"))

    team = [t for t in team if t["estado"] != "Ex-equipo"]
    for i, t in enumerate(team, 1):
        t["notas"] = None
        t["person_id"] = f"P{i:03d}"
    org_files = sorted(glob.glob(os.path.join(args.interno, "*rganigrama*.html")))
    ex_keys = {staff.key(r["nombre"]) for _, rows in staff_sources for r in rows
               if staff.normalize_estado(r.get("estado")) == "Ex-equipo"}
    ex_keys |= {staff.key(str(r["nombre"]).split()[0]) for _, rows in staff_sources for r in rows
                if staff.normalize_estado(r.get("estado")) == "Ex-equipo" and len(str(r["nombre"]).split()) > 1}
    org_rows = merge_org(team, org_files[-1], nicknames, ex_keys) if org_files else []
    staff_rows = staff.staff_personal_rows(team)
    creator_personal = staff.creator_personal_rows(master, prev_creators)

    # marcas
    cases = load_portfolio(args.theme)
    brands = build_brands(args.theme, camps, cases)
    case_rows = [dict(c, case_id=f"C{i:02d}", fuente="S04") for i, c in enumerate(cases, 1)]

    # prensa
    awards, articles = load_press(args.theme)
    press = []
    for it in awards + articles:
        press.append(dict(tipo=it["tipo"], fecha=it.get("fecha") or it.get("anio"), titulo=it["titulo"], medio=it["medio"],
                          caracter=None, descripcion=it["descripcion"], url=it["url"], fuente="S04", nivel=NIVEL_LABEL[2]))
    press += research.press_items(src)
    press = dedupe_press(press)
    for i, p in enumerate(press, 1):
        p["item_id"] = f"PR{i:03d}"
        p["caracter"] = p["caracter"] or research.classify_outlet(p["medio"], p["titulo"], p["url"])

    # assets
    assets = inventory_assets(args.theme)
    creator_by_asset = {c["retrato"]: c["creator_id"] for c in master if c["retrato"]}
    team_by_asset = {t["foto"]: t["person_id"] for t in team if t.get("foto")}
    brand_by_asset = {b["logo"]: b["brand_id"] for b in brands if b["logo"]}
    case_by_asset = {}
    for c in case_rows:
        for a in (c["assets"] or "").split(", "):
            if a:
                case_by_asset.setdefault(a, c["case_id"])
    keep_photos = {t["foto"] for t in team if t.get("foto")}
    gone_portraits = {c["identity"].get("portrait_asset") for c in theme["private"]["creators"] if c["creator_id"] in gone}
    assets = [a for a in assets if not ((a["archivo"].startswith("TEAM_") and a["archivo"] not in keep_photos)
                                        or a["archivo"] in gone_portraits)]
    for i, a in enumerate(assets, 1):
        a["asset_id"] = f"A{i:03d}"
        a["entidad"] = (creator_by_asset.get(a["archivo"]) or team_by_asset.get(a["archivo"])
                        or brand_by_asset.get(a["archivo"]) or case_by_asset.get(a["archivo"]))

    add_conflict_gaps(resolver.conflicts, gaps)
    unclassified = sum(1 for c in camps if not c["marca"])
    if unclassified:
        gaps.append(dict(area="Campañas", entidad=f"{unclassified} menciones", problema="Campañas sin marca clasificada",
                         detalle="Ver 11_Campañas_Creador con filtro Marca vacía", accion="Asignar marca o ampliar reglas"))
    gaps = dedupe_gaps(gaps)
    for g in gaps:
        g["prioridad"] = gap_priority(g)
    gaps.sort(key=lambda g: ({"Alta": 0, "Media": 1, "Baja": 2}[g["prioridad"]], g["area"], str(g["entidad"])))
    ctx = dict(
        src=src, agency=research.agency_rows(src), figures=research.figures(src), timeline=research.timeline(src),
        team=team, org=org_rows, clients=build_clients(org_rows, team, master), legal=load_legal(args.theme), contact=load_contact(args.theme),
        services=load_services(args.theme), staff_personal=staff_rows, creator_personal=creator_personal, creators=master, creator_cols=CREATOR_COLS,
        web_fixes=web_corrections(resolver.conflicts, master),
        platforms=plats, audience=aud, campaigns_creator=camps, brands=brands, cases=case_rows, press=press,
        assets=assets, conflicts=resolver.conflicts, gaps=gaps, mapping=mapping,
        notes=[
            "Documento de trabajo del equipo: aquí se captura y se consulta todo. Celda vacía = dato por rellenar.",
            "Información personal: amarillo = por rellenar, morado = fórmula bloqueada (hoja protegida sin contraseña).",
            "Uso interno: contiene nombres reales, teléfonos y cumpleaños. No compartir fuera de MCR.",
            "Seguidores: por plataforma se toma la cuenta con más seguidores; si no hay cifra, la del roster interno.",
            "Tier de creadores sin tier en la web: por alcance total (AAA ≥ 12M · Mid ≥ 4M · Rising < 4M).",
            "Empresa matriz de cada marca es referencia general: verificar antes de usarla en un pitch.",
        ])
    write_book(args.out, ctx)
    print(f"OK · {args.out}")
    print(f"creadores={len(master)} plataformas={len(plats)} audiencia={len(aud)} campañas={len(camps)} "
          f"marcas={len(brands)} casos={len(case_rows)} prensa={len(press)} assets={len(assets)} "
          f"equipo={len(team)} conflictos={len(resolver.conflicts)} huecos={len(gaps)}")


def role_area(cargo):
    if not cargo:
        return None
    c = cargo.lower()
    if any(w in c for w in ("ceo", "cofundador", "cofundadora", "directora general", "director general", "cfo", "coo")):
        return "Dirección"
    if "account" in c or "project" in c:
        return "Cuentas"
    if any(w in c for w in ("influencer", "contract", "creator program", "talent")):
        return "Talento"
    if any(w in c for w in ("pr ", "community", "campaign", "social media")) or c.startswith("pr"):
        return "Campañas, comunidad y PR"
    if any(w in c for w in ("design", "diseñ", "fotógraf", "videógraf", "production")):
        return "Diseño y producción"
    if "people" in c or "rr. hh" in c:
        return "Personas"
    if c.startswith("it ") or " it " in c or "it manager" in c:
        return "Tecnología"
    return "Operación"


CREATOR_COLS = [
    ("creator_id", "creator_id"), ("nombre_publico", "Creador"), ("variaciones", "Variaciones de nombre"),
    ("nombre_real", "Nombre real"), ("nombre_compartido", "Nombre real compartido"), ("tier", "Tier"),
    ("en_web", "Publicado en web"), ("pais", "País"), ("iso", "ISO"), ("categorias", "Categorías"),
    ("cumple_dia", "Cumple día"), ("cumple_mes", "Cumple mes"),
    ("correo", "Correo corporativo"), ("telefono", "Teléfono"), ("correo_en_redes", "Correo en redes"),
    ("youtube_seguidores", "YouTube"), ("tiktok_seguidores", "TikTok"), ("instagram_seguidores", "Instagram"),
    ("twitch_seguidores", "Twitch"), ("x_seguidores", "X"), ("kick_seguidores", "Kick"), ("alcance_total", "Alcance total"),
    ("plataforma_principal", "Plataforma principal"), ("youtube_promedio_views", "YT prom. views"),
    ("youtube_view_rate_pct", "YT view rate %"), ("estado_campana", "Estado campaña"),
    ("campana_vigente", "Campaña vigente / última"), ("sin_campana_desde", "Sin campaña desde"),
    ("campanas_historicas_n", "Campañas históricas"), ("marcas_frecuentes", "Marcas frecuentes"),
    ("youtube_url", "URL YouTube"), ("tiktok_url", "URL TikTok"), ("instagram_url", "URL Instagram"),
    ("twitch_url", "URL Twitch"), ("x_url", "URL X"), ("kick_url", "URL Kick"), ("otras_cuentas", "Otras cuentas"),
    ("bio", "Bio"), ("media_kit", "Media kit"), ("stream", "Stream en"), ("notas_internas", "Notas"),
]


def apply_creator_adjustments(master, adjustments, resolver, gaps):
    """Ajustes internos (hoja Ajustes_Creadores): nivel 1, se aplican al final."""
    by_id = {m["creator_id"]: m for m in master}
    for a in adjustments:
        m = by_id.get(a["creator_id"])
        campo = str(a["campo"] or "").strip()
        if not m:
            gaps.append(dict(area="Ajustes", entidad=a["creator_id"], problema="Ajuste para un creator_id que no existe",
                             detalle=f"{campo} = {a['valor']}", accion="Revisar el ID en Ajustes_Creadores"))
            continue
        if campo == "nombre_real_compartido":
            m["nombre_compartido"] = "Sí"
        elif campo == "nota":
            m["notas_internas"] = "; ".join(filter(None, [m.get("notas_internas"), str(a["valor"])]))
            continue
        else:
            old = m.get(campo)
            if old not in (None, "") and str(old) != str(a["valor"]):
                resolver.conflicts.append(dict(entidad="Creador", entidad_id=m["creator_id"], campo=campo,
                                               valor_ganador=a["valor"], fuente_ganadora="S39", nivel_ganador=NIVEL_LABEL[1],
                                               valor_descartado=old, fuente_descartada="(consolidado)",
                                               nivel_descartado="—", diferencia_pct=None, severidad="Revisar"))
            m[campo] = a["valor"]
        if a.get("motivo"):
            m["notas_internas"] = "; ".join(filter(None, [m.get("notas_internas"), str(a["motivo"])]))
    shared = {m["creator_id"] for m in master if m.get("nombre_compartido") == "Sí"}
    gaps[:] = [g for g in gaps if not (g["problema"] == "Mismo nombre real en varios creadores"
                                       and set(str(g["entidad"]).split(", ")) <= shared)]


def fix_platform_labels(master, plats, gaps):
    """URLs de Kick capturadas en la columna de Twitch."""
    for m in master:
        url = m.get("twitch_url") or ""
        if "kick.com" in url.lower():
            m["kick_url"], m["kick_seguidores"] = url, m.get("twitch_seguidores")
            m["plataforma_principal"] = "kick" if m.get("plataforma_principal") == "twitch" else m.get("plataforma_principal")
            m["twitch_url"], m["twitch_seguidores"] = None, None
            for p in plats:
                if p["creator_id"] == m["creator_id"] and p["plataforma"] == "twitch":
                    p["plataforma"] = "kick"
            gaps.append(dict(area="Creadores", entidad=m["creator_id"], problema="URL de Kick en la columna Twitch",
                             detalle=url, accion="Reclasificado como Kick; mover a su columna en CREADORES_INTERNOS"))


ORG_AREA = {"Dirección": "Dirección", "Soporte transversal": "Soporte transversal", "Operaciones": "Operaciones"}


def merge_org(team, path, nicknames, ex_keys):
    """Cruza el organigrama con el equipo: cuentas, verticales, rol en cada cuenta, modalidad y junta directiva."""
    rows, board = org.load_org(path)
    out = []
    for t in team:
        t["area"] = None  # el área sale del organigrama; quien no aparece queda por rellenar
    for r in rows:
        person = org.resolve(r["apodo"], team, nicknames)
        if not person:
            if org.key(r["apodo"]) in ex_keys:
                continue
            person = dict(person_id=f"P{len(team) + 1:03d}", nombre=r["apodo"],
                          cargo=(r["rol"] or "").split(" · ")[0] or None, nivel_org="Equipo")
            team.append(person)
        r["person_id"], r["persona"] = person["person_id"], person["nombre"]
        out.append(r)
    for t in team:
        mine = [r for r in out if r["person_id"] == t["person_id"]]
        if not mine:
            continue
        first = mine[0]
        t["area"] = first["cuenta"] if first["bloque"] == "Áreas internas" else (
            "Finanzas" if first["cuenta"] == "Finanzas" else ORG_AREA.get(first["bloque"]))
        cuentas = []
        for r in mine:
            if r["bloque"] in ("Operaciones",):
                cuentas.append(r["cuenta"] + (f" · {r['vertical'].replace('Vertical ', '')}" if r["vertical"] else ""))
        t["cuentas"] = ", ".join(dict.fromkeys(cuentas)) or None
        t["roles_cuenta"] = ", ".join(dict.fromkeys(r["rol"] for r in mine if r["rol"])) or None
        t["lidera"] = ", ".join(dict.fromkeys(r["cuenta"] for r in mine if r["lidera"] and r["bloque"] != "Dirección")) or None
        mods = {r["modalidad"] for r in mine if r["modalidad"]}
        if len({r["cuenta"] for r in mine if r["bloque"] == "Operaciones"}) > 1:
            mods.add("Multi-cuenta")
        t["modalidad"] = ", ".join(sorted(mods)) or "Fijo"
    for b in board:
        person = org.resolve(b["apodo"], team, nicknames)
        if person:
            person["junta"] = b["rol"]
    for r in out:
        r["lidera"] = "Sí" if r["lidera"] else None
        r["vertical"] = ", ".join(x for x in (r["vertical"], r["subvertical"]) if x) or None
    return out


ACCOUNT_CLIENT = {"Bloodstrike": "Blood Strike"}


def build_clients(org_rows, team, creators):
    """Clientes actuales: cuentas del organigrama + agencias exclusivas del sitio. Contactos por rellenar."""
    from brands import EXCLUSIVE
    by_id = {t["person_id"]: t for t in team}
    rows, seen = [], set()
    accounts = [r["cuenta"] for r in org_rows if r["bloque"] == "Operaciones"]
    for acc in dict.fromkeys(accounts):
        if acc in ("Bailes", "Nuevas Cuentas"):
            continue
        name = ACCOUNT_CLIENT.get(acc, acc)
        leads = [r["person_id"] for r in org_rows if r["cuenta"] == acc and r["lidera"]]
        if not leads:
            leads = [t["person_id"] for t in team if t.get("junta") and key(t["junta"]) == key(acc)]
        rows.append(dict(cliente=name, cuenta=acc, lead=", ".join(by_id[l]["nombre"] for l in leads) or None,
                         estatus="Activo"))
        seen.add(key(name))
    for ex in sorted(EXCLUSIVE):
        if key(ex) not in seen:
            active = any(c.get("estado_campana") == "Con campaña" and key(ex) in key(c.get("campana_vigente"))
                         for c in creators)
            rows.append(dict(cliente=ex, estatus="Activo" if active else None))
    for i, r in enumerate(rows, 1):
        r["client_id"] = f"CL{i:02d}"
        r["empresa_matriz"] = BRAND_PARENTS.get(r["cliente"])
        r["exclusiva"] = "Sí" if r["cliente"] in EXCLUSIVE else None
    return rows


def tier_for_reach(reach):
    """Umbrales derivados de los tiers que ya asignó la web (alcance sumado, sin deduplicar)."""
    if not reach:
        return None
    if reach >= 12_000_000:
        return "AAA"
    if reach >= 4_000_000:
        return "Mid"
    return "Rising"


def flag_suspicious_counts(master, gaps):
    for m in master:
        vals = {p: m.get(f"{p}_seguidores") for p in ("youtube", "tiktok", "instagram", "twitch", "x", "kick")}
        top = max([v for v in vals.values() if v] or [0])
        for p, v in vals.items():
            if v is not None and v < 1000 and top >= 100_000:
                gaps.append(dict(area="Seguidores", entidad=m["creator_id"], problema=f"Cifra sospechosa en {p}",
                                 detalle=f"{v:,} vs {top:,} en su plataforma principal",
                                 accion="Confirmar: ¿cuenta correcta o dato mal capturado?"))


def web_corrections(conflicts, master):
    names = {m["creator_id"]: m["nombre_publico"] for m in master}
    rows = []
    for c in conflicts:
        if c["entidad"] != "Creador" or c["fuente_descartada"] != "S03" or not str(c["fuente_ganadora"]).startswith("S01"):
            continue
        rows.append(dict(creator_id=c["entidad_id"], creador=names.get(c["entidad_id"]), campo=c["campo"],
                         en_web=c["valor_descartado"], correcto=c["valor_ganador"], fuente=c["fuente_ganadora"],
                         diferencia_pct=c["diferencia_pct"], severidad=c["severidad"]))
    order = {"Alta": 0, "Revisar": 1, "Media": 2, "Baja": 3}
    return sorted(rows, key=lambda r: (order.get(r["severidad"], 9), r["creator_id"]))


def add_conflict_gaps(conflicts, gaps):
    for c in conflicts:
        if c["entidad"] != "Creador":
            continue
        if c["campo"].endswith("url"):
            gaps.append(dict(area="Cuentas", entidad=c["entidad_id"], problema=f"Cuenta distinta en {c['campo'].split(' ')[0]}",
                             detalle=f"interno: {c['valor_ganador']} · web: {c['valor_descartado']}",
                             accion="Confirmar cuál es la cuenta vigente"))
        elif c["campo"] == "nombre público":
            gaps.append(dict(area="Creadores", entidad=c["entidad_id"], problema="Nombre público distinto entre fuentes",
                             detalle=f"interno: {c['valor_ganador']} · web: {c['valor_descartado']}",
                             accion="Confirmar el nombre oficial (gana el interno por jerarquía)"))


def gap_priority(g):
    p = g["problema"]
    if g["area"] in ("Roster", "Cuentas") or p.startswith("Cifra sospechosa") or "Ajuste para" in p:
        return "Alta"
    if g["area"] in ("Assets", "Web", "Campañas") or p in ("Sin foto de equipo",):
        return "Baja"
    return "Media"


def build_brands(theme_dir, camps, cases):
    logos = {}
    for f in glob.glob(os.path.join(theme_dir, "assets/img/*LOGO.svg")):
        name = os.path.basename(f)
        if name.startswith("PRESS_"):
            continue
        logos[name] = name.replace("LOGO.svg", "")
    from brands import LOGO_BRAND, EXCLUSIVE
    by_brand = defaultdict(lambda: dict(menciones=0, creadores=set()))
    for c in camps:
        if c["marca"]:
            by_brand[c["marca"]]["menciones"] += 1
            by_brand[c["marca"]]["creadores"].add(c["creator_id"])
    names = set(by_brand) | {LOGO_BRAND.get(v, v.title()) for v in logos.values()}
    logo_for = {LOGO_BRAND.get(v, v.title()): k for k, v in logos.items()}
    rows = []
    for i, n in enumerate(sorted(names, key=lambda x: (-by_brand[x]["menciones"], x)), 1):
        matriz = BRAND_PARENTS.get(n)
        related = [c["titulo"] for c in cases if key(n) in key(c["titulo"]) or key(n) == key(c["cliente"] or "")]
        rows.append(dict(brand_id=f"B{i:03d}", marca=n, empresa_matriz=matriz, en_web="Sí" if n in logo_for else "No",
                         exclusiva="Sí" if n in EXCLUSIVE else None, campanas_n=by_brand[n]["menciones"] or None,
                         creadores_n=len(by_brand[n]["creadores"]) or None, casos_portafolio=", ".join(related) or None,
                         logo=logo_for.get(n), fuente_matriz="Referencia general (nivel 3)" if matriz else None))
    return rows


def dedupe_press(items):
    seen, out = set(), []
    for p in sorted(items, key=lambda x: (x["nivel"], str(x["titulo"]))):
        k = key(p["titulo"])[:60]
        if k in seen:
            continue
        seen.add(k)
        out.append(p)
    return sorted(out, key=lambda x: (x["tipo"], str(x["fecha"] or "")), reverse=False)


def dedupe_gaps(gaps):
    seen, out = set(), []
    for g in gaps:
        k = (g["area"], str(g["entidad"]), g["problema"], str(g["detalle"]))
        if k not in seen:
            seen.add(k)
            out.append(g)
    return sorted(out, key=lambda g: (g["area"], str(g["entidad"])))


if __name__ == "__main__":
    main()
