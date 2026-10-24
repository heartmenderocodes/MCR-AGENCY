# Base maestra MCR

Genera el Excel maestro de la agencia a partir de las fuentes internas y del tema web.

```
python3 tools/build_master.py \
  --theme   <carpeta del tema WP descomprimido> \
  --interno <carpeta con los .xlsx internos> \
  --out     <salida>/MCR_MAESTRO_vX.Y.xlsx
```

Requiere `php` (para leer los datasets del tema), `openpyxl`, `beautifulsoup4` y `Pillow`.

## Jerarquía ante datos contradictorios

1. **Interno**: archivos internos de MCR. Siempre gana.
2. **Oficial**: tema web, mcr-agency.com y comunicados propios.
3. **Externo**: prensa, directorios y referencias generales.

En el mismo nivel gana la fuente más reciente. Cada valor descartado queda en la hoja `15_Conflictos`.

## Archivos

- `build_master.py`: carga, consolida y escribe el libro.
- `export_theme_data.php`: exporta los datasets PHP del tema a JSON.
- `brands.py`: reglas para clasificar campañas por marca y el mapa de empresas matriz.
- `research.py`: investigación pública (sitio, comunicados, prensa).

**Los datos no van al repo.** El repo es público y el Excel trae nombres reales, teléfonos y cumpleaños; `.gitignore` bloquea `*.xlsx`, `*.zip` y las carpetas de datos.
