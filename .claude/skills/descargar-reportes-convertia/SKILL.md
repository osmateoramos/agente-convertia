---
name: descargar-reportes-convertia
description: Descarga los reportes de Convertia/inConcert (marketing.convertia.com y convertia.infunnel.inconcert.cloud) definidos en reportes_config.json, corriendo el script de Playwright. Usar cuando el usuario pida "descargar los reportes de Convertia", "bajar las bases de brutas/gestionventas", o similar.
---

Ejecuta la descarga automática de reportes de Convertia para el usuario.

## Archivos involucrados

Ubicación esperada: `C:\Users\dy2059\Mis Archivos\CLAUDE\Codigo\` (el código y la config viven en la subcarpeta `Codigo\`; `Procesados\`, `Libros de Cuenta\`, `Legado\` y `Comentarios\` quedan al mismo nivel que `Codigo\`, no adentro).
- Script: `descargar_reportes_convertia.py`
- Config (cuentas, sitios y reportes): `reportes_config.json`
- Credenciales: `convertia.env` (nunca leer ni mostrar su contenido)
- Los Excel descargados se guardan un nivel arriba de `Codigo\` (el script usa `DOWNLOAD_DIR = RAIZ`, la carpeta padre, no la carpeta del script), para no mezclar datos con código.

**Si no encuentras alguno de estos archivos en esa ruta, no asumas ni busques por tu cuenta en múltiples carpetas — pregúntale directamente al usuario dónde los movió.** Ya ha reorganizado estos archivos varias veces (Descargas → `Documents\CLAUDE` → `OneDrive\Documentos\CLAUDE` → brevemente `OneDrive - Universidad Santo Tomás\Documentos\CLAUDE` → `C:\Users\dy2059\CLAUDE` → `C:\Users\dy2059\Mis Archivos\CLAUDE` → el 25/08/2026 se separó el código en `Codigo\` dentro de esa misma carpeta. **Pendiente**: el usuario quiere renombrar la carpeta raíz de `CLAUDE` a `Agente Convertia` (no se pudo hacer en el momento porque VS Code la tenía abierta como workspace) — si la encuentras con ese nombre nuevo, es la misma carpeta, solo actualiza esta nota. Así que la ruta de arriba es solo la última conocida, no una garantía.

## Pasos

1. Verifica que `convertia.env` exista junto al script (`Test-Path`, sin leer su contenido). Si no aparece ahí, pregúntale al usuario dónde está antes de continuar — sin eso el login falla.
2. No hace falta cerrar Chrome ni tocar procesos: el script abre su propia ventana de Chrome con un perfil temporal, independiente del Chrome normal del usuario. No mates procesos de Chrome — podría cerrarle pestañas que tenga abiertas.
3. Pregunta al usuario qué rango de fechas quiere (si no lo dio ya), y si alguna cuenta necesita un rango distinto al resto (p.ej. "todas del 1 de junio al 16 de agosto, pero Claro_Cr del 1 de julio al 16 de agosto"). Si no da fechas, el script usa el default automático (2 meses atrás + domingo anterior a hoy) para todas las cuentas.
4. Corre el script con salida en vivo y en segundo plano (puede tardar varios minutos: cada reporte puede tomar hasta 5 minutos en generarse, y hay hasta 3 reportes x N cuentas). Ejemplos:
   ```powershell
   # Fechas por defecto para todas las cuentas
   python -u "C:\Users\dy2059\Mis Archivos\CLAUDE\Codigo\descargar_reportes_convertia.py"

   # Fechas explícitas para todas las cuentas
   python -u "C:\Users\dy2059\Mis Archivos\CLAUDE\Codigo\descargar_reportes_convertia.py" --fecha-inicio 2026-06-01 --fecha-fin 2026-08-16

   # Fechas explícitas + una cuenta con rango distinto
   python -u "C:\Users\dy2059\Mis Archivos\CLAUDE\Codigo\descargar_reportes_convertia.py" --fecha-inicio 2026-06-01 --fecha-fin 2026-08-16 --fechas-cuenta "{\"Claro_Cr\": [\"2026-07-01\", \"2026-08-16\"]}"

   # Prueba rápida (mes presente, una cuenta por sitio)
   python -u "C:\Users\dy2059\Mis Archivos\CLAUDE\Codigo\descargar_reportes_convertia.py" --modo-prueba
   ```
   También existe `run_convertia.bat` en la misma carpeta (`run_convertia.bat [fecha-inicio] [fecha-fin] [forzar]`) para que el usuario lo corra por su cuenta sin depender de ti; hace lo mismo que el comando de arriba y además loguea a `run_log_<timestamp>.txt`.
5. El script se puede correr de nuevo sin miedo a duplicar trabajo: si ya existe el `.xlsx` de hoy para una cuenta/alias, lo salta automáticamente (así una corrida que falló a mitad de camino se completa sola en el segundo intento, sin re-descargar lo que ya estaba bien). Si el usuario quiere forzar la re-descarga de algo que ya existe, agrega `--forzar`.
6. Un login fallido en un sitio ya no aborta la corrida completa: ese sitio se salta (con su propia captura `_debug_error_login_<sitio_id>.png`) y el script sigue con el resto.
7. Cuando termine, lee el output y repórtale al usuario, por cada cuenta/reporte, si se descargó bien (línea "Guardado: ..."), si se saltó por ya existir (línea "Ya existe ...") o si falló (línea "ERROR descargando ...").
8. Si algo falló, revisa la captura de diagnóstico que el script guarda automáticamente (`_debug_error_<cuenta>_<alias>.png`, en la misma carpeta) para entender qué pantalla se encontró, y decide si es un problema puntual (reintentar) o un cambio en la interfaz que requiere ajustar un selector en el script (en ese caso, avisa al usuario del hallazgo antes de tocar el código). Para un problema puntual, lo más simple es correr el mismo comando de nuevo: el paso 5 hace que solo se reintente lo que faltó.

## Notas importantes (aprendidas a pulso, no las repitas por las malas)

- El botón "Seleccionar columnas" debe recibir **un solo clic** en "Seleccionar todo" tras abrirse — un doble clic desmarca todas las columnas y el Excel sale incompleto.
- Los botones de exportar/columnas no tienen texto visible ni aria-label; se ubican por el atributo `title` (`button[title="Seleccionar columnas"]`, `button[title="Exportar datos"]`).
- El campo "Periodo" a veces no viene en "Personalizado" por defecto (varía por sitio/reporte) — el script ya lo fuerza explícitamente.
- Si aparece el aviso "Existe otro reporte ejecutándose", el script ya reintenta solo cada 15s hasta 15 veces.
- Cada sitio tiene su propia sesión (dominios distintos) pero usan las mismas credenciales de `convertia.env`.
- Si un reporte trae "Resultados totales: 1" (o muy pocos) puede salir solo con la portada, sin datos — es un comportamiento del sistema BI con datasets mínimos, no un bug del script. Repórtalo al usuario en vez de asumir que falló silenciosamente.
- El nombre exacto de un reporte puede variar por cuenta aunque el alias sea el mismo (p.ej. "gestionventas" es `ReportDetalleGestionVentas` para la mayoría de cuentas, pero `ReportMultiventa` para Mundo_Pacifico) — respeta lo que diga `reportes_config.json`, no asumas que todas las cuentas del mismo alias usan el mismo nombre de reporte.
- Una cuenta "comercial" puede tener varias variantes en el desplegable de cuenta del sitio (p.ej. `wowperu`, `wowperu_in_out`, `wowperu_out`). Cada variante que haya que descargar va como su **propia entrada de nivel superior** en `cuentas` del JSON (mismo `sitio`, con solo los reportes que aplican a esa variante) — no hace falta tocar el script, el nombre de la entrada es literalmente el texto exacto a click-ear en el desplegable de cuenta.
- El PC debe quedar encendido para que esto se pueda disparar desde el celular del usuario (no hay ejecución en la nube).
- El nombre del archivo guardado es `{ALIAS}_{cuenta}_{fecha_inicio}_{fecha_fin}.xlsx` — lleva el **rango consultado**, no la fecha de descarga. Antes solo llevaba la fecha de hoy, y dos descargas el mismo día con rangos distintos (p.ej. modo-prueba y después un rango custom) pisaban silenciosamente el mismo nombre — el limpiador borraba el crudo nuevo pensando que ya estaba procesado. Se corrigió el 25/08/2026; no vuelvas a usar solo la fecha de hoy en el nombre.
