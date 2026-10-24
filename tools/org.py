"""Lectura del organigrama interno (HTML) y cruce con el equipo."""
import re

from bs4 import BeautifulSoup


def _t(el):
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)) if el else None


def _node(el, bloque, cuenta, vertical=None, sub=None, lidera=False):
    tags = [_t(t) for t in el.select(".tag")]
    cls = el.get("class", [])
    modalidad = "Temporal" if "temporal" in cls else "Flota" if any("Flota" in t for t in tags) else (
        "Multi-cuenta" if "multi" in cls else None)
    return dict(bloque=bloque, cuenta=cuenta, vertical=vertical, subvertical=sub, apodo=_t(el.select_one(".name")),
                rol=_t(el.select_one(".role")), lidera=lidera, modalidad=modalidad, detalle="; ".join(tags) or None)


def load_org(path):
    with open(path, encoding="utf-8") as fh:
        soup = BeautifulSoup(fh.read(), "html.parser")
    rows = []
    lead = soup.select_one(".leadership")
    if lead:
        for el in lead.select(".lead-node"):
            rows.append(_node(el, "Dirección", "Dirección", lidera=True))
        for el in lead.select(".finance-unit .node"):
            rows.append(_node(el, "Dirección", "Finanzas"))
    for el in soup.select(".support-band .node"):
        rows.append(_node(el, "Soporte transversal", "Agencia"))
    for ops in soup.select(".ops"):
        bloque = "Operaciones" if "ops-main" in ops.get("class", []) else "Áreas internas"
        for team in ops.select(".team"):
            cuenta = _t(team.select_one(".team-name"))
            vertical = sub = None
            for el in team.find_all(class_=re.compile(r"^(node|subteam-title|nest-label)$")):
                cls = el.get("class", [])
                if "subteam-title" in cls:
                    vertical, sub = _t(el), None
                elif "nest-label" in cls:
                    sub = _t(el)
                elif "node" in cls:
                    rows.append(_node(el, bloque, cuenta, vertical, sub, lidera="lead-of-team" in cls))
    board = [dict(apodo=_t(c.select_one(".name")), rol=_t(c.select_one(".role")))
             for c in soup.select(".board-card") if "add" not in c.get("class", [])]
    return rows, board


def key(s):
    import unicodedata
    s = "".join(c for c in unicodedata.normalize("NFD", str(s or "")) if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", s.lower())


def resolve(apodo, team, nicknames):
    """Apodo del organigrama → persona del equipo (nombre completo, apodo registrado o primer nombre único)."""
    k = key(apodo)
    if k in nicknames:
        target = key(nicknames[k])
        return next((t for t in team if key(t["nombre"]) == target), None)
    exact = [t for t in team if key(t["nombre"]) == k]
    if exact:
        return exact[0]
    first = [t for t in team if key(str(t["nombre"]).split()[0]) == k]
    return first[0] if len(first) == 1 else None
