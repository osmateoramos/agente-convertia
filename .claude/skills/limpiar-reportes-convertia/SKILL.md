---
name: limpiar-reportes-convertia
description: Limpia los reportes de Convertia ya descargados (ver descargar-reportes-convertia) corriendo limpiar_reportes_convertia.py -- detecta la fila real de encabezados y normaliza las columnas de fecha quitandoles la hora. No borra ni toca ninguna otra columna (las tablas se pegan despues en otro archivo que espera la misma estructura). Usar cuando el usuario pida "limpiar los reportes", "limpiarle las fechas a los reportes de Convertia", "quitar la hora de las fechas", o similar.
---

Corre la limpieza automática de los reportes de Convertia ya descargados.

## Archivos involucrados

Misma carpeta que `descargar-reportes-convertia`: `C:\Users\dy2059\Mis Archivos\CLAUDE\`
- Script: `limpiar_reportes_convertia.py`
- Config (de dónde saca los alias válidos para reconocer qué archivos son reportes): `reportes_config.json` — el mismo que usa el descargador, no hay config propia.
- Entrada: los `.xlsx` sueltos en esa carpeta cuyo nombre empiece con un alias conocido (`BRUTAS_...`, `GESTIONVENTAS_...`, los que estén definidos en `reportes_config.json`).
- Salida: una copia limpia de cada uno en la subcarpeta `Procesados\`, con el mismo nombre de archivo. **El crudo original se borra** una vez generado el procesado (con éxito, o si ya existía de antes) — la carpeta principal debe quedar solo con los scripts y `Procesados\`, sin `.xlsx` crudos sueltos. Si un archivo da ERROR al limpiarlo, el crudo correspondiente NO se borra, para poder diagnosticarlo.

**Si no encuentras el script en esa ruta, no asumas ni busques por tu cuenta — pregúntale al usuario dónde está.** Esta carpeta ya se movió varias veces (ver la nota equivalente en el skill de descarga).

## Qué hace exactamente (para que puedas explicarle al usuario qué encontró)

1. Detecta la fila real de encabezados de forma dinámica, no fija: busca la primera fila completamente en blanco y toma la siguiente como encabezado. Así no importa si la guía de arriba tiene 7, 8 o cualquier otro número de filas.
2. De la guía extrae "Cuenta: X" y "Resultados totales: N" y los compara contra el nombre del archivo y la cantidad de filas realmente escritas — si no coinciden, lo reporta como aviso (no rompe la corrida, pero hay que mirarlo).
3. Cualquier columna cuyo nombre empiece con `FCH` o `Fecha` se trata como columna de fecha: se convierte a fecha sin hora (`yyyy-mm-dd`). Si algún valor no se puede interpretar como fecha, se deja tal cual y se avisa en vez de inventar un valor.
4. **No se borra ni se reordena ninguna otra columna** — ni siquiera las `NULL01`...`NULL30` que trae el export. El usuario pega estas tablas en otro archivo que espera exactamente esa estructura de columnas, así que el único cambio permitido es normalizar las fechas. (Hubo una versión anterior de este script que sí las borraba — se corrigió el 25/08/2026 porque rompía el pegado. Si alguna vez alguien pide "quitar las columnas NULL" de nuevo, confirmar bien el motivo antes de tocar esto.)
5. No depende de la IA para decidir archivo por archivo: es el mismo script para todos, determinista — mismo archivo de entrada siempre da el mismo resultado.

Importante: esto **no tiene relación con `Plantilla Unificacion.xlsm`** (el Excel con Power Query que existía antes). Ese Excel busca archivos con el nombre que pone el sitio por defecto (`ConvertiaVentasBrutasCall_...`), que no es el nombre que usa nuestro descargador (`BRUTAS_<cuenta>_<fecha_inicio>_<fecha_fin>.xlsx`) — por eso este script es una ruta nueva e independiente, no un reemplazo que necesite ser compatible con ese Excel.

## Pasos

1. Verifica que `limpiar_reportes_convertia.py` y `reportes_config.json` existan en la carpeta esperada.
2. Corré el script:
   ```powershell
   python "C:\Users\dy2059\Mis Archivos\CLAUDE\limpiar_reportes_convertia.py"
   ```
   Si el usuario quiere rehacer archivos que ya se limpiaron antes (por ejemplo, volvió a descargar el mismo día y quiere refrescar el limpio también):
   ```powershell
   python "C:\Users\dy2059\Mis Archivos\CLAUDE\limpiar_reportes_convertia.py" --forzar
   ```
3. Si ya existe el limpio correspondiente (mismo nombre en `Procesados\`) y no se pasó `--forzar`, el script lo salta solo — repórtaselo al usuario como "ya estaba limpio", no como error.
4. Lee el output y repórtale al usuario, por archivo: fila donde se detectó el encabezado, cantidad de filas de datos, columnas de fecha limpiadas, y cualquier **AVISO** (cuenta no coincide, total de resultados no coincide, valores de fecha no interpretables). Los avisos no son errores fatales, pero hay que señalarlos — pueden indicar un reporte con un formato distinto al esperado.
5. Si un archivo da **ERROR** (no "AVISO"), es porque el formato no coincidió con lo esperado (no se encontró la fila en blanco separadora, o la fila siguiente no parece encabezado real) — no se generó archivo de salida para ese caso. Revisa las primeras filas del archivo a mano (`openpyxl`, `read_only=True`, `values_only=True`) para ver qué cambió antes de tocar el script.
6. Solo procesa archivos cuyo nombre empiece con un alias que esté en `reportes_config.json` — si el usuario agrega un alias nuevo ahí (como se hizo con `wowperu_in_out`/`wowperu_out`), este script lo reconoce automáticamente sin tocar código.
7. Como el crudo se borra tras limpiarlo, si el usuario corre el descargador de nuevo el mismo día DESPUÉS de haber limpiado, el descargador ya no va a encontrar el archivo de hoy como "ya existe" y lo va a volver a descargar — es el costo de no dejar crudos sueltos. Avísale esto si te pregunta por qué se re-descargó algo que ya había bajado ese día.
