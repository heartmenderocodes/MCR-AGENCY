"""Formato de textos: nombres propios, oraciones, plataformas y fechas en español."""
import datetime as dt
import re
import unicodedata

PARTICLES = {"de", "del", "la", "las", "los", "y", "da", "dos", "do", "van", "von"}

# Nombres y apellidos frecuentes que se escriben con tilde (clave: sin tilde, minúsculas)
ACCENTS = {w: w for w in []}
for w in ("José Jesús María Sofía Lucía Verónica Mónica Belén Valentín Fabián Iván Adrián Julián Sebastián Andrés "
          "Nicolás Tomás Martín Joaquín Germán Rubén Simón Ángel Elías Saúl Raúl Moisés Josué Inés Rocío Matías "
          "Ramón Héctor Víctor Ivón Noé Efraín Darío Mario Alán Yahír Ximena Íñigo Óscar "
          "Suárez Martínez García Hernández González Pérez Rodríguez Sánchez Ramírez Gómez Díaz López Fernández "
          "Jiménez Álvarez Gutiérrez Muñoz Domínguez Vázquez Núñez Ordóñez Alcántara Benítez Chávez Cortés Durán "
          "Galván Guzmán Ibáñez Juárez Márquez Méndez Velázquez Ríos Báez Beltrán Calderón Castañeda Colón León "
          "Monzón Peñaloza Narváez Mejía Mejías Añorve Espinosa Tejeda Covarrubias Pastrana Saldaña Patiño Rangel "
          "Ortíz Zúñiga Bermúdez Estévez Gálvez Míguez Sáenz Jácome Pesántez Layana Peña Muñiz Avilés Cárdenas "
          "Robles Ávila Cázares Gil Leal Pineda").split():
    k = "".join(c for c in unicodedata.normalize("NFD", w.lower()) if unicodedata.category(c) != "Mn")
    if k != w.lower() and k not in ("mario", "ortiz", "alan", "oscar"):
        ACCENTS[k] = w
ACCENTS.pop("ortiz", None)


def _plain(s):
    return "".join(c for c in unicodedata.normalize("NFD", s.lower()) if unicodedata.category(c) != "Mn")


def proper_name(s):
    """'ALBERTO SALUDES MORA' → 'Alberto Saludes Mora'; respeta mayúsculas internas (KiongSu) y añade tildes
    comunes (Martinez → Martínez)."""
    if not s:
        return s
    words = re.split(r"(\s+|-)", str(s).strip())
    out = []
    first = True
    for w in words:
        if not w or w.isspace() or w == "-":
            out.append(w)
            continue
        mixed = any(c.isupper() for c in w[1:]) and not w.isupper()
        low = w.lower()
        if mixed:
            word = w
        elif low in PARTICLES and not first:
            word = low
        else:
            word = low[:1].upper() + low[1:]
        word = ACCENTS.get(_plain(word), word) if not mixed else word
        out.append(word)
        first = False
    return "".join(out)


MONTHS = "enero febrero marzo abril mayo junio julio agosto septiembre setiembre octubre noviembre diciembre".split()
ACRONYMS = {"HSR", "MLBB", "PUBGM", "PUBG", "COD", "CODM", "NTE", "R6M", "R6", "MW3", "MWIII", "BS", "WT", "HP", "OPPO",
            "ESIM", "TT", "IG", "YT", "FF", "BO6", "BO7", "EGS", "GX", "VW", "PK", "AFA", "TDR", "TDRS", "BJFS", "EPAL",
            "HOK", "FNAF", "DJ", "MC", "CDMX", "LATAM", "OOH", "AAA", "UGC", "VR", "IA", "AI", "PR", "CEO", "CFO",
            "COO", "IT", "MX", "ES", "BR", "GT", "OBOC", "S2", "S4", "S5", "S6", "UID", "GU", "NBA", "NFL"}


COMMON = set("""sin con campaña campañas activa activas activo desde de del y e en para por la las el los un una mes mensual
stream baile launch tablet laptop short shorts tiktok reel reels general update evento event long term varias varios
campaign oficial nuevo nueva temporada season reaccion reacción""".split())
COMMON = {_plain(w) for w in COMMON}


def sentence(s, keep=()):
    """Oración: primera letra en mayúscula y el resto en minúscula, salvo siglas, cifras con letras (BO6, 340GU),
    nombres propios recibidos en `keep` (marcas, handles) y palabras que ya venían como nombre propio."""
    if s is None:
        return s
    s = re.sub(r"\s+", " ", str(s)).strip()
    if not s:
        return s
    keep_map = {}
    for k in keep:
        if k:
            for token in re.findall(r"[\wÀ-ÿ&'.]+", str(k)):
                if _plain(token) not in COMMON:
                    keep_map.setdefault(_plain(token), token)

    def fix(m):
        w = m.group(0)
        p = _plain(w)
        if p in keep_map:
            return keep_map[p]
        if w.upper() in ACRONYMS or (re.search(r"\d", w) and re.search(r"[A-Za-z]", w)):
            return w.upper()
        if len(w) > 1 and w.isupper():
            return w.lower()
        if w[:1].isupper() and (p in COMMON or p in MONTHS):
            return w.lower()
        return w

    out = re.sub(r"[\wÀ-ÿ]+", fix, s)
    i = next((i for i, c in enumerate(out) if c.isalnum()), None)
    if i is not None and out[i].isalpha() and not out[i].isupper():
        out = out[:i] + out[i].upper() + out[i + 1:]
    return out.replace("actiiva", "activa")


def us_dates_to_es(s):
    """'07/23/2025' (mm/dd/aaaa, como viene del roster) → '23/07/2025'."""
    if not s:
        return s
    return re.sub(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", lambda m: f"{int(m.group(2)):02d}/{int(m.group(1)):02d}/{m.group(3)}", s)


PLATFORM_LABEL = {"youtube": "YouTube", "tiktok": "TikTok", "instagram": "Instagram", "twitch": "Twitch", "x": "X",
                  "kick": "Kick", "facebook": "Facebook", "other": "Otra"}


def platform(p):
    return PLATFORM_LABEL.get(str(p or "").lower(), p)


def yes_no(v):
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return "Sí" if v else "No"
    s = str(v).strip().lower()
    return {"sí": "Sí", "si": "Sí", "true": "Sí", "no": "No", "false": "No"}.get(s, v)


MES = {m[:3]: i + 1 for i, m in enumerate("enero febrero marzo abril mayo junio julio agosto septiembre octubre "
                                          "noviembre diciembre".split())}
MES["set"] = 9


def es_date(v):
    """Normaliza fechas sueltas ('2 de julio de 2024', '25 nov 2024', '2025-07-25', '2024') a texto dd/mm/aaaa o aaaa."""
    if v is None:
        return None
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%d/%m/%Y")
    s = str(v).strip().lower()
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    m = re.fullmatch(r"(\d{1,2})\s*(?:de\s+)?([a-záéíóú]+)\.?\s*(?:de\s+|del\s+)?(\d{4})", s)
    if m and _plain(m.group(2))[:3] in MES:
        return f"{int(m.group(1)):02d}/{MES[_plain(m.group(2))[:3]]:02d}/{m.group(3)}"
    m = re.fullmatch(r"(\d{4})-(\d{2})\s*→\s*(\d{4})-(\d{2})", s)
    if m:
        return f"{m.group(2)}/{m.group(1)} – {m.group(4)}/{m.group(3)}"
    if re.fullmatch(r"\d{4}", s):
        return s
    return None


def date_sort_key(text):
    if not text:
        return "0000"
    parts = text.split("/")
    if len(parts) == 3:
        return f"{parts[2]}{parts[1]}{parts[0]}"
    if len(parts) == 2:
        return f"{parts[1][:4]}{parts[0]}00"
    return f"{text[:4]}0000"
