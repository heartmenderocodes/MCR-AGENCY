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

import research
import staff
from brands import BRAND_RULES, BRAND_PARENTS

VERSION = "0.2.0"
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


def url_key(v):
    s = str(v or "").lower().strip()
    s = re.sub(r"^https?://", "", s)
    s = re.sub(r"^(www\.|m\.)", "", s)
    s = s.split("?")[0].rstrip("/")
    return s


def parse_phone(v):
    if blank(v):
        return None, None
    if isinstance(v, float):
        v = str(int(v))
    digits = re.sub(r"\D", "", str(v))
    if not digits:
        return None, f"teléfono no interpretable: {v}"
    note = "sin código de país aparente" if len(digits) < 10 else None
    return "+" + digits, note


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
        email, email_src = R("correo corporativo", [(str(r.get("email")).strip().lower() if not blank(r.get("email")) else None, s)
                                                     for s, r in ints]
                             + [((((w or {}).get("contact") or {}).get("corporate_email") or "").lower() or None, S_WEB)])
        tel, tel_note, tel_src = None, None, None
        for s, r in ints:
            tel, tel_note = parse_phone(r.get("telefono"))
            if tel:
                tel_src = s
                break
        real_name = next((str(r["nombre_real"]).strip() for s, r in ints if not blank(r.get("nombre_real"))), None)

        plat_out = {}
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
            n, n_src = R(f"{p} · seguidores", num_c, numeric=True)
            u, u_src = R(f"{p} · url", url_c)
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
        row = dict(
            creator_id=cid, nombre_publico=name, fuente_nombre=name_src, nombre_real=real_name,
            estado_roster=roster, tier=wid.get("tier") or "Sin tier", fuente_tier=S_WEB if wid.get("tier") else None,
            pais=COUNTRY.get(pais_code, (None,))[0] if pais_code else None, iso=COUNTRY.get(pais_code, (None, None))[1] if pais_code else None,
            fuente_pais=pais_src, categorias=cats, fuente_categorias=cats_src,
            categorias_web=", ".join(wid.get("categories", [])) or None,
            cumple_dia=bd["dia"] if bd else None, cumple_mes=bd["mes"] if bd else None,
            correo=email, fuente_correo=email_src, telefono=tel, fuente_telefono=tel_src,
            correo_en_redes=next((r.get("correo_en_redes") for s, r in ints if not blank(r.get("correo_en_redes"))), None),
            **{f"{p}_seguidores": plat_out[p]["n"] for p in ("youtube", "tiktok", "instagram", "twitch", "x")},
            **{f"{p}_url": plat_out[p]["url"] for p in ("youtube", "tiktok", "instagram", "twitch", "x")},
            youtube_promedio_views=yt["avg"], youtube_view_rate_pct=yt_rate,
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

C_BG = "1B1033"
C_ACCENT = "7B3FE4"
C_HEADER_FONT = "FFFFFF"
C_ZEBRA = "F5F1FC"
TIER_FILL = {"AAA": "FFE08A", "Mid": "CFE3FF", "Rising": "D9F2D9"}
LEVEL_FILL = {"1 · Interno": "D9F2D9", "2 · Oficial": "CFE3FF", "3 · Externo": "F3E2C7"}
THIN = Side(style="thin", color="DDD6EE")


class Book:
    def __init__(self):
        self.wb = openpyxl.Workbook()
        self.wb.remove(self.wb.active)
        self.tables = 0
        self.index = []

    def sheet(self, name, title, description, columns, rows, widths=None, link_cols=(), date_cols=(),
              int_cols=(), pct_cols=(), wrap_cols=()):
        labels = [label for _, label in columns]
        dupes = {l for l in labels if labels.count(l) > 1}
        if dupes:  # Excel rechaza tablas con encabezados repetidos
            raise ValueError(f"{name}: encabezados repetidos {dupes}")
        ws = self.wb.create_sheet(name)
        self.index.append((name, description, len(rows)))
        ws["A1"] = title
        ws["A1"].font = Font(bold=True, size=16, color=C_BG)
        ws["A2"] = description
        ws["A2"].font = Font(italic=True, size=10, color="6B6480")
        hdr = 4
        for i, (_, label) in enumerate(columns, 1):
            c = ws.cell(hdr, i, label)
            c.font = Font(bold=True, color=C_HEADER_FONT)
            c.fill = PatternFill("solid", fgColor=C_BG)
            c.alignment = Alignment(vertical="center", wrap_text=True)
        for r_i, row in enumerate(rows, hdr + 1):
            for c_i, (field, _) in enumerate(columns, 1):
                v = row.get(field) if isinstance(row, dict) else row[c_i - 1]
                if isinstance(v, (list, tuple)):
                    v = ", ".join(map(str, v))
                cell = ws.cell(r_i, c_i, v)
                cell.border = Border(bottom=THIN)
                if field in link_cols and isinstance(v, str) and v.startswith(("http", "www", "instagram", "tiktok", "twitch")):
                    url = v if v.startswith("http") else "https://" + v
                    cell.hyperlink = url.split()[0]
                    cell.font = Font(color=C_ACCENT, underline="single")
                if field in date_cols and isinstance(v, (dt.date, dt.datetime)):
                    cell.number_format = "DD/MM/YYYY"
                if field in int_cols and isinstance(v, (int, float)):
                    cell.number_format = "#,##0"
                if field in pct_cols and isinstance(v, (int, float)):
                    cell.number_format = "0.00"
                if field in wrap_cols:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
        last_row = hdr + max(len(rows), 1)
        last_col = get_column_letter(len(columns))
        if rows:
            self.tables += 1
            t = Table(displayName=f"T_{re.sub(r'[^A-Za-z0-9]', '_', name)}", ref=f"A{hdr}:{last_col}{last_row}")
            t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
            ws.add_table(t)
        ws.freeze_panes = ws.cell(hdr + 1, 3 if len(columns) > 4 else 1)
        widths = widths or {}
        for i, (field, label) in enumerate(columns, 1):
            ws.column_dimensions[get_column_letter(i)].width = widths.get(field, max(12, min(40, len(label) + 4)))
        ws.row_dimensions[hdr].height = 32
        ws.sheet_view.showGridLines = False
        self._level_format(ws, columns, hdr, last_row)
        return ws

    def _level_format(self, ws, columns, hdr, last_row):
        for i, (field, _) in enumerate(columns, 1):
            col = get_column_letter(i)
            rng = f"{col}{hdr + 1}:{col}{last_row}"
            if field in ("nivel", "nivel_ganador", "nivel_descartado"):
                for lab, color in LEVEL_FILL.items():
                    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{lab}"'],
                                                                  fill=PatternFill("solid", fgColor=color)))
            if field == "tier":
                for lab, color in TIER_FILL.items():
                    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{lab}"'],
                                                                  fill=PatternFill("solid", fgColor=color)))
            if field == "severidad":
                for lab, color in (("Alta", "F8C9C9"), ("Media", "FBE3B5"), ("Baja", "EEF0F2"), ("Revisar", "E6DAFB")):
                    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{lab}"'],
                                                                  fill=PatternFill("solid", fgColor=color)))
            if field == "estado_roster":
                ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({col}{hdr + 1},4)="Sólo"'],
                                                               fill=PatternFill("solid", fgColor="FBE3B5")))

    def save(self, path):
        self.wb.save(path)


def cols(*pairs):
    return list(pairs)


SH = dict(idx="00_Índice", res="01_Resumen", agencia="02_Agencia", cifras="03_Cifras_declaradas",
          crono="04_Cronología", equipo="05_Equipo", p_staff="06_Personal_Staff", p_creadores="07_Personal_Creadores",
          creadores="08_Creadores", plataformas="09_Plataformas", audiencia="10_Audiencia",
          campanas="11_Campañas_Creador", marcas="12_Marcas", portafolio="13_Portafolio", prensa="14_Prensa_Premios",
          assets="15_Assets", web="16_Correcciones_Web", conflictos="17_Conflictos", pendientes="18_Pendientes",
          mapeo="19_Mapeo_IDs", fuentes="20_Fuentes")


def write_book(out, ctx):
    b = Book()
    src = ctx["src"]
    idx_ws = b.wb.create_sheet(SH["idx"])
    res = b.wb.create_sheet(SH["res"])

    b.sheet(SH["agencia"], "Ficha de la agencia", "Datos corporativos. Las cifras declaradas van por fuente y fecha.",
            cols(("campo", "Campo"), ("valor", "Valor"), ("fuente", "Fuente"), ("nivel", "Nivel"), ("fecha", "Fecha"),
                 ("notas", "Notas")),
            ctx["agency"], widths=dict(campo=30, valor=60, notas=50), date_cols=("fecha",), wrap_cols=("valor", "notas"))

    b.sheet(SH["cifras"], "Cifras declaradas por fuente",
            "Cada cifra pública de la agencia con su fecha. Sirve para ver evolución y contradicciones.",
            cols(("metrica", "Métrica"), ("valor", "Valor"), ("fecha", "Fecha"), ("fuente", "Fuente"),
                 ("nivel", "Nivel"), ("notas", "Notas")),
            ctx["figures"], widths=dict(metrica=32, valor=22, notas=50), date_cols=("fecha",))

    b.sheet(SH["crono"], "Cronología", "Hitos de la agencia 2019 → hoy.",
            cols(("fecha", "Fecha"), ("anio", "Año"), ("hito", "Hito"), ("categoria", "Categoría"),
                 ("fuente", "Fuente"), ("nivel", "Nivel")),
            ctx["timeline"], widths=dict(hito=80), date_cols=("fecha",), wrap_cols=("hito",))

    b.sheet(SH["equipo"], "Equipo", "Web (nivel 2) corregida con el staff interno (nivel 1): altas, bajas y cargos.",
            cols(("person_id", "ID"), ("nombre", "Nombre"), ("cargo", "Cargo"), ("area", "Área"), ("nivel_org", "Nivel"),
                 ("estado", "Estado"), ("salida", "Fecha salida"), ("foto", "Foto"), ("fuente", "Fuente alta"),
                 ("fuente_cargo", "Fuente cargo"), ("fuente_estado", "Fuente estado"), ("notas", "Notas")),
            sorted(ctx["team"], key=lambda t: ({"Activo": 0, "Por confirmar": 1, "Ex-equipo": 2}[t["estado"]], t["person_id"])),
            widths=dict(nombre=24, cargo=34, notas=50), date_cols=("salida",))

    for key_, title, desc, columns, rows in (
        ("p_staff", "Información personal · staff",
         "Cumpleaños, edad, antigüedad y aniversario del equipo. Activos primero, ex-equipo al final (sin aniversario).",
         staff.STAFF_PERSONAL_COLS, ctx["staff_personal"]),
        ("p_creadores", "Información personal · creadores",
         "Cumpleaños y antigüedad en MCR (desde la firma). La firma se captura en las celdas amarillas.",
         staff.CREATOR_PERSONAL_COLS, ctx["creator_personal"]),
    ):
        ws = b.sheet(SH[key_], title, desc, [(f, lab) for f, lab, _ in columns], rows,
                     widths=dict(nombre=24, nombre_real=26, cargo=30, estado=22, notas=36, antiguedad=14))
        staff.finish_personal_sheet(ws, columns, len(rows))

    b.sheet(SH["creadores"], "Creadores · maestra",
            "Una fila por creador. Los valores siguen la jerarquía Interno > Oficial > Externo; la columna fuente_* dice quién ganó.",
            ctx["creator_cols"], ctx["creators"],
            widths=dict(nombre_publico=20, nombre_real=26, estado_roster=26, categorias=24, bio=60, campana_vigente=34,
                        marcas_frecuentes=40, correo=30, notas_internas=40),
            link_cols=("youtube_url", "tiktok_url", "instagram_url", "twitch_url", "x_url", "kick_url"),
            int_cols=tuple(f"{p}_seguidores" for p in ("youtube", "tiktok", "instagram", "twitch", "x", "kick"))
            + ("alcance_total", "youtube_promedio_views"),
            pct_cols=("youtube_view_rate_pct",), date_cols=("sin_campana_desde",))

    b.sheet(SH["plataformas"], "Creadores × plataforma", "Formato largo para tablas dinámicas.",
            cols(("creator_id", "creator_id"), ("creador", "Creador"), ("plataforma", "Plataforma"), ("url", "URL"),
                 ("seguidores", "Seguidores"), ("fuente_seguidores", "Fuente seguidores"),
                 ("promedio_views", "Promedio views"), ("fuente_views", "Fuente views"), ("principal", "Principal"),
                 ("fuente_url", "Fuente URL")),
            ctx["platforms"], link_cols=("url",), int_cols=("seguidores", "promedio_views"), widths=dict(url=44))

    b.sheet(SH["audiencia"], "Audiencia", "Distribuciones de audiencia disponibles (media kits procesados en la web).",
            cols(("creator_id", "creator_id"), ("creador", "Creador"), ("plataforma", "Plataforma"),
                 ("dimension", "Dimensión"), ("valor", "Valor"), ("porcentaje", "%"), ("calidad", "Calidad campo"),
                 ("estado", "Estado registro"), ("origen", "Origen"), ("fuente", "Fuente")),
            ctx["audience"], pct_cols=("porcentaje",))

    b.sheet(SH["campanas"], "Historial de campañas por creador",
            "Cada mención de CAMPAÑAS HISTÓRICAS. 'marca' se clasifica por reglas; vacía = sin clasificar.",
            cols(("creator_id", "creator_id"), ("creador", "Creador"), ("campana", "Campaña (texto original)"),
                 ("marca", "Marca / título"), ("empresa_matriz", "Empresa matriz"), ("mes", "Mes"), ("anio", "Año"),
                 ("fuente", "Fuente")),
            ctx["campaigns_creator"], widths=dict(campana=50, marca=26, empresa_matriz=26))

    b.sheet(SH["marcas"], "Marcas y clientes", "Marcas del sitio + marcas detectadas en el historial de campañas.",
            cols(("brand_id", "ID"), ("marca", "Marca / título"), ("empresa_matriz", "Empresa matriz"),
                 ("en_web", "Logo en web"), ("exclusiva", "Agencia exclusiva"), ("campanas_n", "Menciones en historial"),
                 ("creadores_n", "Creadores distintos"), ("casos_portafolio", "Casos en portafolio"), ("logo", "Asset logo"),
                 ("fuente_matriz", "Fuente matriz")),
            ctx["brands"], widths=dict(marca=28, empresa_matriz=30, casos_portafolio=36))

    b.sheet(SH["portafolio"], "Casos de portafolio", "Casos publicados en la web.",
            cols(("case_id", "ID"), ("titulo", "Caso"), ("subtitulo", "Subtítulo"), ("cliente", "Cliente"),
                 ("servicios", "Servicios"), ("region", "Región"), ("kpis", "KPIs"), ("descripcion", "Descripción"),
                 ("assets", "Assets"), ("fuente", "Fuente")),
            ctx["cases"], widths=dict(titulo=30, kpis=50, descripcion=70, assets=40), wrap_cols=("kpis", "descripcion"))

    b.sheet(SH["prensa"], "Prensa, premios y presencia",
            "Lo que publica la web (nivel 2) + lo encontrado en investigación (nivel 3).",
            cols(("item_id", "ID"), ("tipo", "Tipo"), ("fecha", "Fecha"), ("titulo", "Título"), ("medio", "Medio / organización"),
                 ("caracter", "Carácter"), ("descripcion", "Descripción"), ("url", "URL"), ("fuente", "Fuente"),
                 ("nivel", "Nivel")),
            ctx["press"], widths=dict(titulo=56, descripcion=46, url=40, caracter=24), link_cols=("url",),
            wrap_cols=("titulo",))

    b.sheet(SH["assets"], "Inventario de assets", "Todos los archivos del tema web v2.5.0.",
            cols(("asset_id", "ID"), ("archivo", "Archivo"), ("categoria", "Categoría"), ("entidad", "Entidad vinculada"),
                 ("tipo", "Tipo"), ("kb", "KB"), ("dimensiones", "Dimensiones"), ("usado_en", "Usado en"),
                 ("ruta", "Ruta")),
            ctx["assets"], widths=dict(archivo=34, usado_en=40, ruta=44))

    b.sheet(SH["web"], "Correcciones para la web",
            "Datos del sitio que el roster interno contradice. Valor correcto = interno (nivel 1).",
            cols(("creator_id", "creator_id"), ("creador", "Creador"), ("campo", "Campo"), ("en_web", "Hoy en la web"),
                 ("correcto", "Valor correcto"), ("fuente", "Fuente"), ("diferencia_pct", "Dif. %"), ("severidad", "Severidad")),
            ctx["web_fixes"], widths=dict(en_web=40, correcto=40), int_cols=("en_web", "correcto"), pct_cols=("diferencia_pct",))

    b.sheet(SH["conflictos"], "Conflictos entre fuentes",
            "Cada valor descartado por la jerarquía. Filtra por severidad para revisar lo importante.",
            cols(("entidad", "Entidad"), ("entidad_id", "ID"), ("campo", "Campo"), ("valor_ganador", "Valor ganador"),
                 ("fuente_ganadora", "Fuente ganadora"), ("nivel_ganador", "Nivel ganador"),
                 ("valor_descartado", "Valor descartado"), ("fuente_descartada", "Fuente descartada"),
                 ("nivel_descartado", "Nivel descartado"), ("diferencia_pct", "Dif. %"), ("severidad", "Severidad")),
            ctx["conflicts"], widths=dict(valor_ganador=36, valor_descartado=36), pct_cols=("diferencia_pct",))

    b.sheet(SH["pendientes"], "Pendientes", "Lo que falta o no cuadra, por prioridad, con la acción sugerida.",
            cols(("prioridad", "Prioridad"), ("area", "Área"), ("entidad", "Entidad"), ("problema", "Problema"),
                 ("detalle", "Detalle"), ("accion", "Acción")),
            ctx["gaps"], widths=dict(problema=44, detalle=50, accion=44))

    b.sheet(SH["mapeo"], "Mapeo de IDs", "Cómo se llama cada creador en cada fuente y cómo se emparejó.",
            cols(("creator_id", "creator_id"), ("fuente", "Fuente"), ("nombre_en_fuente", "Nombre en la fuente"),
                 ("metodo", "Método de match"), ("fila", "Fila origen")),
            ctx["mapping"], widths=dict(metodo=44))

    b.sheet(SH["fuentes"], "Fuentes", "Registro de fuentes y su nivel en la jerarquía.",
            cols(("id", "ID"), ("nombre", "Fuente"), ("nivel", "Nivel"), ("tipo", "Tipo"), ("ubicacion", "Ubicación"),
                 ("fecha", "Fecha"), ("notas", "Notas")),
            [dict(r, nivel=NIVEL_LABEL[r["nivel"]]) for r in src.rows.values()],
            widths=dict(nombre=44, ubicacion=60, notas=50), link_cols=("ubicacion",), date_cols=("fecha",))

    write_summary(res, ctx)
    staff.protect(res)
    write_index(idx_ws, b, ctx)
    b.save(out)


def write_summary(ws, ctx):
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Resumen"
    ws["A1"].font = Font(bold=True, size=16, color=C_BG)
    ws["A2"] = "Conteos vivos (fórmulas sobre las hojas). Si editas las hojas, esto se actualiza."
    ws["A2"].font = Font(italic=True, size=10, color="6B6480")
    cr = f"'{SH['creadores']}'"
    ncol = {f: get_column_letter(i) for i, (f, _) in enumerate(ctx["creator_cols"], 1)}
    rng = lambda f: f"{cr}!${ncol[f]}$5:${ncol[f]}$2000"
    blocks = [
        ("Creadores", [
            ("Total en la base", f"=COUNTA({rng('creator_id')})"),
            ("Interno + web", f'=COUNTIF({rng("estado_roster")},"Interno + web")'),
            ("Sólo interno", f'=COUNTIF({rng("estado_roster")},"Sólo interno*")'),
            ("Sólo web (verificar baja)", f'=COUNTIF({rng("estado_roster")},"Sólo web*")'),
            ("Con campaña", f'=COUNTIF({rng("estado_campana")},"Con campaña")'),
            ("Con cumpleaños", f"=COUNT({rng('cumple_dia')})"),
            ("Alcance total sumado", f"=SUM({rng('alcance_total')})"),
        ]),
        ("Por tier", [(t, f'=COUNTIF({rng("tier")},"{t}")') for t in ("AAA", "Mid", "Rising", "Sin tier")]),
        ("Por país", [(p, f'=COUNTIF({rng("pais")},"{p}")')
                      for p, _ in Counter(c["pais"] for c in ctx["creators"] if c["pais"]).most_common()]),
        ("Calidad", [
            ("Conflictos registrados", f"=COUNTA('{SH['conflictos']}'!$A$5:$A$5000)"),
            ("Conflictos severidad alta", f"=COUNTIF('{SH['conflictos']}'!$K$5:$K$5000,\"Alta\")"),
            ("Correcciones para la web", f"=COUNTA('{SH['web']}'!$A$5:$A$5000)"),
            ("Pendientes · prioridad alta", f"=COUNTIF('{SH['pendientes']}'!$A$5:$A$5000,\"Alta\")"),
            ("Pendientes · total", f"=COUNTA('{SH['pendientes']}'!$A$5:$A$5000)"),
        ]),
        ("Equipo y otros", [
            ("Equipo activo", f"=COUNTIF('{SH['equipo']}'!$F$5:$F$500,\"Activo\")"),
            ("Ex-equipo", f"=COUNTIF('{SH['equipo']}'!$F$5:$F$500,\"Ex-equipo\")"),
            ("Cumpleaños en los próximos 30 días (staff)",
             f"=COUNTIFS('{SH['p_staff']}'!$K$5:$K$500,\">=0\",'{SH['p_staff']}'!$K$5:$K$500,\"<=30\")"),
            ("Cumpleaños en los próximos 30 días (creadores)",
             f"=COUNTIFS('{SH['p_creadores']}'!$M$5:$M$500,\">=0\",'{SH['p_creadores']}'!$M$5:$M$500,\"<=30\")"),
            ("Marcas", f"=COUNTA('{SH['marcas']}'!$A$5:$A$1000)"),
            ("Casos de portafolio", f"=COUNTA('{SH['portafolio']}'!$A$5:$A$500)"),
            ("Menciones de campañas", f"=COUNTA('{SH['campanas']}'!$A$5:$A$10000)"),
            ("Assets", f"=COUNTA('{SH['assets']}'!$A$5:$A$2000)"),
        ]),
    ]
    r = 4
    for title, items in blocks:
        ws.cell(r, 1, title).font = Font(bold=True, color=C_HEADER_FONT)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=C_BG)
        ws.cell(r, 2).fill = PatternFill("solid", fgColor=C_BG)
        r += 1
        for label, formula in items:
            ws.cell(r, 1, label)
            c = ws.cell(r, 2, formula)
            c.number_format = "#,##0"
            c.font = Font(bold=True, color=C_ACCENT)
            r += 1
        r += 1
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 18


def write_index(ws, b, ctx):
    ws.sheet_view.showGridLines = False
    ws["A1"] = "MCR Agency · Base maestra"
    ws["A1"].font = Font(bold=True, size=20, color=C_BG)
    ws["A2"] = f"Versión {VERSION} · generada el {TODAY.strftime('%d/%m/%Y')} · uso interno, contiene datos personales"
    ws["A2"].font = Font(italic=True, color="6B6480")
    ws["A4"] = "Jerarquía ante datos contradictorios"
    ws["A4"].font = Font(bold=True, size=12)
    rules = [
        ("1 · Interno", "Archivos internos de MCR (roster interno, calendarios, nóminas…). Siempre gana."),
        ("2 · Oficial", "Tema web / mcr-agency.com / comunicados propios. Complementa lo que el interno no tiene."),
        ("3 · Externo", "Prensa, directorios, referencias generales. Sólo llena huecos."),
        ("Empate", "Mismo nivel: gana la fuente más reciente. Todo lo descartado queda en 15_Conflictos."),
    ]
    for i, (k, v) in enumerate(rules, 5):
        c = ws.cell(i, 1, k)
        c.font = Font(bold=True)
        if k in LEVEL_FILL:
            c.fill = PatternFill("solid", fgColor=LEVEL_FILL[k])
        ws.cell(i, 2, v)
    r = 10
    ws.cell(r, 1, "Hoja").font = Font(bold=True, color=C_HEADER_FONT)
    ws.cell(r, 2, "Contenido").font = Font(bold=True, color=C_HEADER_FONT)
    ws.cell(r, 3, "Filas").font = Font(bold=True, color=C_HEADER_FONT)
    for c in range(1, 4):
        ws.cell(r, c).fill = PatternFill("solid", fgColor=C_BG)
    entries = [(SH["res"], "Conteos vivos", None)] + b.index
    for name, desc, n in entries:
        r += 1
        c = ws.cell(r, 1, name)
        c.hyperlink = f"#'{name}'!A1"
        c.font = Font(color=C_ACCENT, underline="single")
        ws.cell(r, 2, desc)
        ws.cell(r, 3, n)
    r += 2
    ws.cell(r, 1, "Notas").font = Font(bold=True, size=12)
    for note in ctx["notes"]:
        r += 1
        ws.cell(r, 1, "•")
        ws.cell(r, 2, note).alignment = Alignment(wrap_text=True)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 100
    ws.column_dimensions["C"].width = 8


# ─────────────────────────── main ───────────────────────────

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
    staff_raw, adjustments = staff.load_internal_adjustments(adj_files)
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
    team_by_asset = {t["foto"]: t["person_id"] for t in team if t["foto"]}
    brand_by_asset = {b["logo"]: b["brand_id"] for b in brands if b["logo"]}
    case_by_asset = {}
    for c in case_rows:
        for a in (c["assets"] or "").split(", "):
            if a:
                case_by_asset.setdefault(a, c["case_id"])
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
    staff.write_rh_template(os.path.join(os.path.dirname(os.path.abspath(args.out)), "PLANTILLA_RH_STAFF.xlsx"),
                            team, VERSION)
    ctx = dict(
        src=src, agency=research.agency_rows(src), figures=research.figures(src), timeline=research.timeline(src),
        team=team, staff_personal=staff_rows, creator_personal=creator_personal, creators=master, creator_cols=CREATOR_COLS,
        web_fixes=web_corrections(resolver.conflicts, master),
        platforms=plats, audience=aud, campaigns_creator=camps, brands=brands, cases=case_rows, press=press,
        assets=assets, conflicts=resolver.conflicts, gaps=gaps, mapping=mapping,
        notes=[
            "Datos personales (nombres reales, teléfonos, cumpleaños): no subir este archivo a repositorios ni carpetas públicas.",
            "Los tokens de acceso de los perfiles privados del tema NO se incluyen: son la llave de páginas noindex.",
            "Conteos de seguidores del roster interno vienen escritos a mano; se normalizan y los dudosos van a 18_Pendientes.",
            "Hojas de información personal: amarillo = captura, morado = fórmula bloqueada. Protegidas sin contraseña.",
            "Lo capturado a mano en celdas amarillas se conserva al regenerar (opción --previo).",
            "Staff: PLANTILLA_RH_STAFF.xlsx va a RR. HH.; al regresar llena se deja en la carpeta interna y entra como nivel 1.",
            "Promedio de views, bio, tier y audiencia sólo existen en la web (nivel 2), así que se toman de ahí.",
            "Empresa matriz de cada marca es referencia general (nivel 3): verificar antes de usar en un pitch.",
            "Para actualizar: correr tools/build_master.py con las fuentes nuevas; fuera de las celdas amarillas no se edita a mano.",
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
    if any(w in c for w in ("ceo", "cofundador", "cofundadora", "directora general", "cfo", "coo")):
        return "Dirección"
    if "account" in c:
        return "Cuentas"
    if "influencer" in c or "contract" in c:
        return "Talento"
    if "pr " in c or c.startswith("pr") or "community" in c or "campaign" in c:
        return "Campañas / Comunidad / PR"
    if "design" in c:
        return "Diseño y producción"
    return "Operación"


CREATOR_COLS = [
    ("creator_id", "creator_id"), ("nombre_publico", "Creador"), ("nombre_real", "Nombre real"),
    ("estado_roster", "Estado roster"), ("tier", "Tier"), ("pais", "País"), ("iso", "ISO"),
    ("categorias", "Categorías"), ("cumple_dia", "Cumple día"), ("cumple_mes", "Cumple mes"),
    ("correo", "Correo corporativo"), ("telefono", "Teléfono"), ("correo_en_redes", "Correo en redes"),
    ("youtube_seguidores", "YouTube"), ("tiktok_seguidores", "TikTok"), ("instagram_seguidores", "Instagram"),
    ("twitch_seguidores", "Twitch"), ("x_seguidores", "X"), ("alcance_total", "Alcance total"),
    ("plataforma_principal", "Plataforma principal"), ("youtube_promedio_views", "YT prom. views"),
    ("youtube_view_rate_pct", "YT view rate %"), ("estado_campana", "Estado campaña"),
    ("campana_vigente", "Campaña vigente / última"), ("sin_campana_desde", "Sin campaña desde"),
    ("campanas_historicas_n", "Campañas históricas"), ("marcas_frecuentes", "Marcas frecuentes"),
    ("youtube_url", "URL YouTube"), ("tiktok_url", "URL TikTok"), ("instagram_url", "URL Instagram"),
    ("twitch_url", "URL Twitch"), ("x_url", "URL X"), ("kick_url", "URL Kick"), ("kick_seguidores", "Kick"),
    ("nombre_compartido", "Nombre real compartido"), ("notas_internas", "Notas internas"), ("bio", "Bio (web)"), ("retrato", "Retrato"),
    ("audiencia_estado", "Audiencia"), ("media_kit", "Media kit"), ("stream", "Stream en"),
    ("categorias_web", "Categorías (web)"), ("fuente_nombre", "Fuente nombre"), ("fuente_pais", "Fuente país"),
    ("fuente_categorias", "Fuente categorías"), ("fuente_correo", "Fuente correo"),
    ("fuente_telefono", "Fuente teléfono"), ("fuente_tier", "Fuente tier"), ("datos_web_generados", "Datos web al"),
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
            m["twitch_url"], m["twitch_seguidores"] = None, None
            for p in plats:
                if p["creator_id"] == m["creator_id"] and p["plataforma"] == "twitch":
                    p["plataforma"] = "kick"
            gaps.append(dict(area="Creadores", entidad=m["creator_id"], problema="URL de Kick en la columna Twitch",
                             detalle=url, accion="Reclasificado como Kick; mover a su columna en CREADORES_INTERNOS"))


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
