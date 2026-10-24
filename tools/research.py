"""
Capa de investigación pública sobre MCR Agency (sitio oficial, comunicados y prensa).
Sólo contiene información publicada; nada interno ni personal.
"""
import datetime as dt

D = dt.date

EXTERNAL_SOURCES = [
    # id, nombre, nivel, tipo, url, fecha
    ("S05", "mcr-agency.com (sitio en vivo)", 2, "Sitio oficial", "https://www.mcr-agency.com/", D(2026, 10, 24)),
    ("S06", "Comunicado MCR · Europa Press / Infobae / Merca2", 2, "Comunicado propio",
     "https://www.merca2.es/2024/11/25/de-agencia-emergente-a-referente-en-el-gaming-en-hispanoamerica-la-historia-de-exito-de-mcr-agency-2077106/",
     D(2024, 11, 25)),
    ("S07", "Comunicado MCR · PUBG MOBILE LatAm (Motor16 / Infobae)", 2, "Comunicado propio",
     "https://www.motor16.com/agencia/mcr-agency-gestionara-las-rrpp-e-influencer-marketing-de-pubg-mobile-en-america-latina/",
     D(2025, 7, 25)),
    ("S10", "Rock Culture · entrevista", 3, "Prensa",
     "https://www.rockculture.es/entrevista-mcr-agency-es-una-agencia-de-creadores-de-contenido-online-que-tiene-especialistas-en-todas-las-plataformas/",
     D(2021, 9, 26)),
    ("S11", "Que.es / Merca2 · notas 2022-2024", 3, "Prensa patrocinada",
     "https://www.que.es/2024/02/19/la-agencia-de-representacion-de-creadores-de-contenido-de-barcelona-mcr-agency/",
     D(2024, 2, 19)),
    ("S12", "El Publicista · Especial Influencer Marketing 2025", 3, "Especial patrocinado",
     "https://www.elpublicista.es/negocio-visto-por-las-agencias/sergi-cerrato-mcr-agency-creadores-gaming-han-pasado-ser-nicho",
     D(2025, 4, 22)),
    ("S13", "Líderes Mexicanos · entrevista a Sergi Cerrato", 3, "Prensa",
     "https://lideresmexicanos.com/entrevistas/sergi-cerrato-empoderando-a-los-creadores-de-gaming-en-mcr-agency",
     D(2025, 8, 8)),
    ("S14", "Excélsior · inversión de Hugo Pibernat", 3, "Prensa",
     "https://www.excelsior.com.mx/nacional/mcr-agency-anuncia-inversion-hugo-pibernat-para-expandirse-hacia-nuevos-mercados-verticales",
     D(2026, 2, 13)),
    ("S15", "Revista Mercado · CorporateLiveWire 2026", 3, "Prensa",
     "https://mercado.com.ar/marketing/mcr-agency-recibe-el-premio-content-creation-agency-of-the-year-spain-2026",
     D(2026, 8, 13)),
    ("S16", "Revista Mercado · GIMA 2026", 3, "Prensa",
     "https://mercado.com.ar/marketing/mcr-agency-volvio-a-quedar-finalista-en-los-global-influencer-marketing-awards-2026",
     D(2026, 9, 7)),
    ("S17", "TV Azteca Deportes · entrevista", 3, "Prensa",
     "https://www.tvazteca.com/aztecadeportes/entrevista-sergi-cerrato-ceo-mcr-agency/", D(2026, 8, 14)),
    ("S18", "Uno Editorial · The keys to influencer marketing", 3, "Editorial",
     "https://www.unoeditorial.com/libros/the-keys-to-influencer-marketing/", None),
    ("S19", "Soy.Marketing · columnas de Sergi Cerrato", 3, "Prensa",
     "https://soy.marketing/columnista/sergi-cerrato/", D(2025, 7, 14)),
    ("S20", "Merca2 · nominación GIMA 2023", 3, "Prensa patrocinada",
     "https://www.merca2.es/2023/07/18/mcr-agency-nominada-a-los-global-influencer-marketing-awards-1377319/",
     D(2023, 7, 18)),
    ("S21", "Que.es · MCR Creators Exclusive Event México", 3, "Prensa patrocinada",
     "https://www.que.es/2023/08/18/mcr-creators-exclusive-event-el-evento-que-va-a-efectuar-mcr-agency-en-mexico-con-sus-creadores-de-contenido/",
     D(2023, 8, 18)),
    ("S22", "Dealroom · ficha MCR Agency", 3, "Directorio", "https://app.beta.dealroom.co/companies/mcr_agency/team",
     None),
]


def register_sources(src):
    for sid, nombre, nivel, tipo, url, fecha in EXTERNAL_SOURCES:
        src.add(sid, nombre, nivel, tipo, url, fecha)


def _lv(src, sid):
    return {1: "1 · Interno", 2: "2 · Oficial", 3: "3 · Externo"}[src.level(sid)]


def agency_rows(src):
    rows = [
        ("Nombre comercial", "MCR Agency", "S05", None, None),
        ("Entidad legal", "SCIA Developers S.L.", "S05", None, "Según el pie del sitio. No localizada en directorios mercantiles públicos."),
        ("Fundación", "2019", "S05", None, "El Publicista (abr 2025) habla de '7 años en activo', que apuntaría a 2018."),
        ("Fundadores", "Sergi Cerrato (CEO) · Inés Alexandre (cofundadora)", "S04", None,
         "Ambos con formación en Derecho y Comercio Internacional según prensa 2022-2023."),
        ("Dirección", "Sergi Cerrato (CEO) · Betsy Constantino (Directora General) · Yeray Torres (CFO) · Carlos Naranjo (COO)",
         "S04", None, None),
        ("Socio inversor", "Hugo Pibernat (ex King, ex Scopely)", "S14", D(2026, 2, 13),
         "Monto y participación no revelados. Plan: nuevos mercados y verticales en 24 meses."),
        ("Tagline", "We make gaming happen.", "S05", None, None),
        ("Tagline secundario", "We manage, create, represent.", "S04", None, None),
        ("Posicionamiento", "Agencia de referencia en gaming para mercados de habla hispana y Brasil", "S05", None, None),
        ("Servicios", "Influencer Marketing · Talent Management · Strategy & Creativity · Content Production · Events · OOH",
         "S05", None, None),
        ("Agencia exclusiva de", "Tencent · Activision · Duolingo · Infinix · Bambu Lab", "S05", None, None),
        ("Oficinas", "Barcelona · Madrid · Ciudad de México · São Paulo", "S05", None,
         "Dealroom ubica la sede en CDMX; la prensa española la presenta como barcelonesa."),
        ("Dirección Barcelona", "Carrer Diputació 338 E03, 08009 Barcelona", "S05", None, None),
        ("Operación", "Trilingüe ES / PT / EN · equipos dedicados por cliente", "S05", None, None),
        ("Correo general", "hello@mcr-agency.com", "S05", None, None),
        ("Teléfono", "+34 931 59 77 45", "S05", None, None),
        ("Instagram", "https://instagram.com/mcr.agency", "S05", None, None),
        ("LinkedIn", "https://linkedin.com/company/mcragency", "S05", None, None),
        ("Dominios", "mcr-agency.com (principal) · mcr-agency.online (espejo)", "S03", None, None),
        ("Membresías", "Forbes Agency Council (desde mayo 2024) · IAB México", "S04", None,
         "Forbes Agency Council es una membresía de pago por invitación."),
        ("Sergi Cerrato · nombre completo", "Sergi Cerrato Recasens", "S19", None, None),
        ("Sergi Cerrato · perfil", "Abogado, nacido en Sant Andreu de Llavaneres (Barcelona). Máster en Comercio Internacional y "
         "Supply Chain; Máster en Derecho español.", "S18", None, None),
        ("Sergi Cerrato · origen de MCR", "Venía de un proyecto con creadores; vio que tenían influencia pero no conocimiento "
         "de contratos o facturas, y fundó la agencia para profesionalizarlos.", "S13", D(2025, 8, 8), None),
        ("Sergi Cerrato · libro", "The keys to influencer marketing · Uno Editorial · 176 págs. · ISBN 978-84-10319-26-4",
         "S18", None, None),
        ("Sergi Cerrato · columnas", "Soy.Marketing (6 columnas 2024-2025) · Forbes Agency Council (2024)", "S19", None, None),
        ("Sergi Cerrato · discurso", "Profesionalizar la industria; comunidad > seguidores; colaboraciones de largo plazo; "
         "'gamificar' marcas no endémicas.", "S12", D(2025, 4, 22), None),
    ]
    return [dict(campo=c, valor=v, fuente=s, nivel=_lv(src, s), fecha=f or src.rows[s]["fecha"], notas=n)
            for c, v, s, f, n in rows]


def figures(src):
    rows = [
        ("Empleados", "6", D(2021, 9, 26), "S10", None),
        ("Empleados", "27 (desde 4)", D(2024, 11, 25), "S06", None),
        ("Empleados", "~30", D(2025, 8, 8), "S13", "Meta declarada: 40"),
        ("Creadores representados", "100+ (red no exclusiva)", D(2021, 9, 26), "S10", None),
        ("Creadores representados", "60+", D(2024, 2, 19), "S11", None),
        ("Creadores exclusivos", "60+ (LatAm y España)", D(2025, 4, 22), "S12", None),
        ("Creadores exclusivos", "70+ (65 de LatAm)", D(2025, 8, 8), "S13", "Meta: 80-90 exclusivos"),
        ("Creadores exclusivos", "80", D(2026, 10, 24), "S05", None),
        ("Red de creadores", "2K+", D(2026, 10, 24), "S05", None),
        ("Marcas globales", "30+", D(2026, 10, 24), "S05", None),
        ("Campañas ejecutadas", "1K+", D(2026, 10, 24), "S05", None),
        ("Views en YouTube", "73B+", D(2026, 10, 24), "S05", None),
        ("Facturación", "> 5M USD (año previo a la entrevista)", D(2025, 8, 8), "S13", "Única fuente que la menciona"),
        ("Años en activo", "7", D(2025, 4, 22), "S12", "Choca con fundación 2019 del sitio"),
        ("Stumble Guys", "> 1B views · 50+ creadores", D(2024, 11, 25), "S06", None),
    ]
    return [dict(metrica=m, valor=v, fecha=f, fuente=s, nivel=_lv(src, s), notas=n) for m, v, f, s, n in rows]


def timeline(src):
    rows = [
        (None, 2019, "Fundación de MCR en Barcelona por Sergi Cerrato e Inés Alexandre.", "Corporativo", "S05"),
        (None, 2020, "La pandemia frena casi por completo las campañas con marcas.", "Contexto", "S10"),
        (D(2021, 9, 26), 2021, "6 personas y 100+ creadores. Clientes: TikTok, Epic Games, Binomo, Frag Games, Druni y Five Guys.",
         "Escala", "S10"),
        (D(2022, 1, 14), 2022, "Presencia en medios digitales; se posiciona entre las agencias más fuertes de TikTok gaming.", "Prensa", "S11"),
        (D(2023, 7, 18), 2023, "Nominada a Best Boutique Influencer Marketing Agency · GIMA.", "Premios", "S20"),
        (D(2023, 8, 18), 2023, "MCR Creators Exclusive Event en Ciudad de México.", "Evento", "S21"),
        (D(2024, 5, 1), 2024, "Entra al Forbes Agency Council.", "Membresía", "S04"),
        (D(2024, 7, 2), 2024, "Columna de Sergi Cerrato en Forbes: impacto cultural de los creadores en LatAm.", "Prensa", "S04"),
        (D(2024, 11, 25), 2024, "Crece de 4 a 27 empleados; Stumble Guys supera 1B de views; finalista en los SABRE Awards LatAm.", "Escala", "S06"),
        (D(2025, 4, 22), 2025, "Entrevista en El Publicista: 60+ creadores exclusivos y presencia en México, Argentina, Chile y Brasil.", "Prensa", "S12"),
        (D(2025, 7, 23), 2025, "PUBG MOBILE nombra a MCR su agencia de RR. PP. e influencer marketing en LatAm.", "Cliente", "S07"),
        (D(2025, 8, 8), 2025, "Entrevista en Líderes Mexicanos: ~30 personas, 70+ talentos y facturación superior a 5M USD.", "Escala", "S13"),
        (None, 2025, "Finalista GIMA · Boutique Talent Agency (Londres).", "Premios", "S04"),
        (D(2026, 2, 13), 2026, "Hugo Pibernat entra como socio inversor estratégico.", "Corporativo", "S14"),
        (D(2026, 8, 13), 2026, "Premio Content Creation Agency of the Year – Spain · CorporateLiveWire.", "Premios", "S15"),
        (D(2026, 8, 14), 2026, "Entrevista de Sergi Cerrato en TV Azteca Deportes.", "Prensa", "S17"),
        (D(2026, 9, 3), 2026, "Finalista GIMA por segundo año consecutivo (The Londoner, Londres).", "Premios", "S16"),
        (None, 2026, "Media Partner de Gamescom Latam 2026 · presencia en GDC 2026.", "Eventos", "S04"),
    ]
    return [dict(fecha=f, anio=a, hito=h, categoria=c, fuente=s, nivel=_lv(src, s)) for f, a, h, c, s in rows]


def press_items(src):
    rows = [
        ("Prensa", "2026-09-07", "MCR Agency volvió a quedar finalista en los Global Influencer Marketing Awards 2026",
         "Revista Mercado", "S16"),
        ("Prensa", "2026-08-14", "Las marcas también quieren jugar · Entrevista con Sergi Cerrato", "TV Azteca Deportes", "S17"),
        ("Prensa", "2026-08-13", "MCR Agency recibe el premio Content Creation Agency of the Year – Spain 2026",
         "Revista Mercado", "S15"),
        ("Prensa", "2026-02-13", "MCR Agency anuncia inversión de Hugo Pibernat para expandirse hacia nuevos mercados y verticales",
         "Excélsior", "S14"),
        ("Prensa", "2025-08-08", "Sergi Cerrato empoderando a los creadores de gaming en MCR Agency", "Líderes Mexicanos", "S13"),
        ("Prensa", "2025-07-25", "MCR Agency gestionará las RRPP e influencer marketing de PUBG MOBILE en América Latina",
         "Motor16", "S07"),
        ("Prensa", "2025-04-22", "Sergi Cerrato (MCR Agency): \"Los creadores de gaming han pasado de ser un nicho a dominar "
         "el entretenimiento digital\"", "El Publicista", "S12"),
        ("Prensa", "2024-11-25", "De agencia emergente a referente en el gaming en hispanoamérica, la historia de éxito de MCR Agency",
         "Merca2", "S06"),
        ("Prensa", "2023-07-18", "MCR-Agency nominada a los Global Influencer Marketing Awards", "Merca2", "S20"),
        ("Prensa", "2021-09-26", "MCR Agency es una agencia de creadores de contenido online que tiene especialistas en todas "
         "las plataformas", "Rock Culture", "S10"),
        ("Libro", None, "The keys to influencer marketing (Sergi Cerrato Recasens)", "Uno Editorial", "S18"),
        ("Columnas", "2024-06 → 2025-07", "6 columnas de Sergi Cerrato", "Soy.Marketing", "S19"),
        ("Premio / reconocimiento", "2024", "Finalista SABRE Awards Latin America · Media, Arts & Entertainment (Stumble Guys)",
         "SABRE Awards", "S06"),
        ("Premio / reconocimiento", "2026", "Content Creation Agency of the Year – Spain", "CorporateLiveWire", "S15"),
        ("Premio / reconocimiento", "2026", "GIMA Awards — Finalistas (2.º año)", "Global Influencer Marketing Awards", "S16"),
    ]
    return [dict(tipo=t, fecha=f, titulo=ti, medio=m, caracter=None, descripcion=None, url=src.rows[s]["ubicacion"],
                 fuente=s, nivel=_lv(src, s)) for t, f, ti, m, s in rows]


SPONSORED = ("merca2", "que.es", "moncloa", "cantabria", "siglo xxi", "majadahonda", "corporate")
EDITORIAL = ("tv azteca", "excélsior", "excelsior", "mercado", "líderes", "lideres", "teragames", "motor16",
             "rock culture", "control publicidad")


def classify_outlet(medio, titulo, url):
    m = (medio or "").lower()
    t = (titulo or "").lower()
    u = (url or "").lower()
    if "comunicado" in u or "europa press" in m or "infobae" in m:
        return "Comunicado de MCR (pagado)"
    if "forbes" in m:
        return "Membresía pagada / columna de miembro"
    if "publicista" in m:
        return "Especial patrocinado"
    if "corporatelivewire" in m.replace(" ", "") or "eu business" in m or "best gaming influencer" in t:
        return "Programa de reconocimiento comercial (probable)"
    if "gima" in t or "global influencer" in m or "sabre" in m:
        return "Premio con jurado de industria"
    if any(s in m for s in SPONSORED):
        return "Contenido patrocinado (probable)"
    if "líderes" in t or "lideres" in t or "top 300" in t:
        return "Ranking editorial"
    if any(s in m for s in EDITORIAL):
        return "Editorial"
    if "uno editorial" in m:
        return "Libro"
    return None
