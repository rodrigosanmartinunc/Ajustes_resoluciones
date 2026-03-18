# Pipeline ERSeP — Boletín Oficial CBA

Scraper automático que descarga PDFs del Boletín Oficial de Córdoba,
filtra resoluciones de ERSeP y extrae los campos clave a un Excel.

---

## Archivos del proyecto

```
├── scraper.py                          ← Script principal
├── ERSeP_Analisis_Colab.ipynb          ← Notebook para Colab
├── .github/
│   └── workflows/
│       └── scraper_ersep.yml           ← Automatización en GitHub Actions
└── pdfs_descargados/                   ← PDFs descargados (se crea automáticamente)
```

---

## Cómo usar

### Opción A — GitHub Actions (automático)

1. Subí estos archivos a un repo en GitHub (puede ser privado)
2. Ir a **Actions → Scraper ERSeP → Run workflow**
3. Elegí cuántos días buscar (default: 30)
4. Al terminar, descargá el Excel desde **Artifacts**

El workflow también corre automáticamente lunes a viernes a las 10am (hora Argentina).

### Opción B — Google Colab (manual)

1. Abrí `ERSeP_Analisis_Colab.ipynb` en Colab
2. Seguí las celdas en orden
3. Al final descargás el Excel directo a tu PC

### Opción C — Local (si tenés Python instalado)

```bash
pip install requests beautifulsoup4 pdfplumber openpyxl
python scraper.py 30    # busca los últimos 30 días
```

---

## Campos extraídos al Excel

| Campo            | Descripción                                   |
|------------------|-----------------------------------------------|
| Fecha Boletín    | Fecha de publicación del boletín              |
| Año              | Año de la resolución                          |
| Resolución       | Número de resolución General ERSeP            |
| Período Costos   | Período de costos considerado                 |
| Período Inicio   | Fecha inicio del período                      |
| Período Cierre   | Fecha cierre del período                      |
| % Otorgado       | Porcentaje de aumento tarifario aprobado      |
| Rige Desde       | Fecha desde la que rige la resolución         |
| Prestadora       | Nombre de la cooperativa / prestadora         |
| CUIT             | CUIT de la prestadora                         |
| Expediente       | Número de expediente                          |
| Tarifa Agua      | Tarifa promedio agua                          |
| Tarifa Cloacas   | Tarifa promedio cloacas                       |
| Página PDF       | Página del PDF donde se encontró              |
| URL PDF          | Link directo al PDF del boletín               |

---

## Notas técnicas

- El sitio usa **AWS CloudFront** que bloquea IPs de Google Colab.
  GitHub Actions usa IPs de Microsoft Azure que no están bloqueadas.
- Los PDFs siguen el patrón:
  `https://boletinoficial.cba.gov.ar/wp-content/4p96humuzp/YYYY/MM/SECCION_DDMMYY.pdf`
- La extracción de campos usa **regex** sobre el texto del PDF.
  Si el PDF está escaneado (imagen), los campos pueden no extraerse correctamente.
- Para ajustar las palabras clave de filtrado, editá la variable `PALABRAS_CLAVE` en `scraper.py`.
