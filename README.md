# KPI — Vegetación Costera Uruguay

**Indicador de Performance (KPI) de vegetación costera** para la costa del Uruguay, basado en el índice NDVI y el cociente DVR (Dune Vegetation Ratio) derivados de imágenes Sentinel-2 procesadas en Google Earth Engine.

🌐 **Miniweb:** https://gaitapi.github.io/kpi-ndvi-costera/

---

## Descripción

El indicador cubre **~357–360 tramos costeros** definidos sobre la línea de costa del Uruguay (Colonia → Rocha), con un buffer de análisis de **45 metros** a cada lado. Se calcula el NDVI mediano anual para el período **2017–2026** en temporada estival (diciembre–marzo), usando imágenes Sentinel-2 SR con máscara de nubes < 20 %.

La tendencia se calcula mediante la prueba de **Mann-Kendall** y la pendiente de **Sen (Theil-Sen estimator)** para cada tramo.

---

## Estructura del repositorio

```
kpi-ndvi-costera/
├── index.html                         # Visor interactivo NDVI por tramo (autocontenido)
├── kpi_dvr_visor.html                 # Visor DVR — carga CSV exportado de GEE
├── kpi_ndvi_senslope.py               # Script Python: cálculo tendencias Sen + Mann-Kendall
├── data/
│   ├── KPI_NDVI_CostaUruguay_buf45_2017_2026.csv  # NDVI anual por tramo (buf 45 m)
│   ├── KPI_NDVI_tendencia_45m.csv                 # Tendencias Sen/MK por tramo
│   └── KPI_DVR_CostaUruguay_2017_2024.csv         # DVR + NDVI + NDSSI por tramo (buf 90 m)
└── gee/
    ├── kpi_ndvi_tramos_v3_45m.js      # Script GEE — buffer 45 m (versión actual)
    ├── kpi_ndvi_tramos_v3_90m.js      # Script GEE — buffer 90 m
    └── kpi_ndvi_tramos_v1.js          # Script GEE — versión inicial
```

---

## Índices utilizados

| Índice | Fórmula | Uso |
|--------|---------|-----|
| **NDVI** | (NIR − RED) / (NIR + RED) | Cobertura y vigor de vegetación |
| **DVR** | NDVI > 0.2 | Fracción de píxeles con vegetación activa |
| **NDSSI** | (GREEN − NIR) / (GREEN + NIR) | Suelo desnudo / arena |

---

## Metodología

- **Satélite:** Sentinel-2 Level-2A (Surface Reflectance)
- **Período:** diciembre–marzo (verano austral), 2017–2026
- **Buffer:** 45 m a cada lado de la línea de costa (IMFIA Q80)
- **Filtro de nubes:** < 20 % cobertura
- **Reducción:** mediana anual por tramo
- **Tendencia:** Mann-Kendall + pendiente de Sen (α = 0.05)

---

## Clasificación de tramos

Cada tramo tiene atributos de **departamento** (d), **vulnerabilidad** (v: 1=alta, 2=media, 3=baja) y **acciones registradas** (a: cantidad de acciones de manejo registradas en la capa base), heredados de la zonificación de la DINABISE/DGCM.

---

## Fuentes y créditos

- Línea de costa: IMFIA (cuantil 80)
- Zonificación y tramos: DINABISE / DGCM — Ministerio de Ambiente, Uruguay
- Procesamiento: Google Earth Engine + Python
- Autor: gustavo_pineiro / gaitapi · 2025–2026
