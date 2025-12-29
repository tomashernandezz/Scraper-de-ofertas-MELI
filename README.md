# Scraper de Ofertas Mercado Libre (MELI AR) + Ranking de Relevancia

Scraper en Python que extrae productos desde la página de **ofertas de Mercado Libre Argentina** y genera un **Excel (.xlsx)** con las ofertas **ordenadas por un score de relevancia configurable** (descuento, ahorro, “cheapness”, y match con keywords/marcas).  
Incluye **deduplicación** para evitar repetidos y **descarga opcional de imágenes**.

---

## ✨ Features

- Scrapea: `https://www.mercadolibre.com.ar/ofertas`
- Extrae (por producto):
  - Nombre
  - Precio antes
  - Precio actual
  - Descuento
  - Link de compra
  - Score de relevancia
- Ordena por score (mayor primero)
- Deduplica (por ID/URL y nombre normalizado)
- Exporta a **Excel** con:
  - Encabezados formateados
  - Filtros (AutoFilter / Table)
  - Freeze panes (fila 1)
  - Tabla estilo Excel (banded rows)
- (Opcional) descarga imágenes a `./imagenes/`

---

## 📦 Requisitos

- Python **3.10+**
- Dependencias:
  - `requests`
  - `beautifulsoup4`
  - `openpyxl`

Instalación:
```bash
pip install -r requirements.txt
