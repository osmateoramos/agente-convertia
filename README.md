# Agente de descarga y consolidación de reportes Convertia

Agente que automatiza el flujo operativo de reportes de ventas/gestión de Convertia
(inConcert) para varias cuentas de call center: descarga, limpieza y consolidación en
los libros de seguimiento de cada cuenta.

## Arquitectura

El agente es **Claude Code + un conjunto de skills** (`.claude/skills/`), no un script
único. Cada skill le da a Claude las instrucciones, el contexto operativo y las
salvedades aprendidas en producción para invocar la herramienta correcta y decidir qué
hacer con el resultado — el trabajo mecánico (miles de filas, fórmulas, XML) lo hacen
scripts Python deterministas; el criterio de qué pegar, cómo resolver un desajuste de
columnas, o si un resultado es correcto, lo aporta el agente.

```
.claude/skills/
  descargar-reportes-convertia/   -- login + descarga vía Playwright
  limpiar-reportes-convertia/     -- normaliza fechas y detecta encabezados
  comentario-operacional/         -- redacta el comentario analítico de cierre

descargar_reportes_convertia.py   -- descarga los reportes crudos (Playwright)
limpiar_reportes_convertia.py     -- limpieza determinista (openpyxl)
pegar_en_cdm.py / pegar_en_claro.py -- consolida los reportes limpios en los libros
                                       de seguimiento de cada cuenta (edición XML
                                       quirúrgica, sin recargar el libro completo)
reportes_config.json              -- sitios, cuentas y reportes (declarativo)
```

## Setup

```powershell
pip install -r requirements.txt
playwright install chrome
```

Copiá `.env.example` a `convertia.env` (junto al script) y completá tus credenciales
de Convertia. `convertia.env` nunca se versiona.

## Uso

```powershell
# Descargar con el rango de fechas por defecto
python descargar_reportes_convertia.py

# Rango explícito
python descargar_reportes_convertia.py --fecha-inicio 2026-07-01 --fecha-fin 2026-08-23

# Limpiar lo descargado (detecta encabezado, normaliza fechas, no toca otras columnas)
python limpiar_reportes_convertia.py
```

También podés simplemente pedirle a Claude Code "descargá los reportes de Convertia"
o "limpiá los reportes" — las skills en `.claude/skills/` cubren el flujo completo,
incluyendo qué hacer ante errores conocidos (reporte ocupado, columnas movidas, cuenta
con variantes en el sitio, etc.).

## Qué no está en este repositorio

Los `.xlsx`/`.xlsm` descargados o procesados y `convertia.env` quedan fuera
deliberadamente (`.gitignore`): contienen datos comerciales reales de clientes y
credenciales. El repositorio versiona el agente (código + skills), no los datos que
produce.
