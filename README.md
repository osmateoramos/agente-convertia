# Agente de reportes operativos Convertia

Agente que automatiza el flujo operativo de varias cuentas de call center que usan
Convertia/inConcert: descarga los reportes de ventas y gestión, los limpia, los
consolida en el libro de seguimiento de cada cuenta, y redacta el comentario
analítico de cierre que acompaña esos números.

Cuentas cubiertas: **WOW Perú CPA, Mundo Pacífico (Chile), DirecTV Perú y Claro Costa
Rica**.

## Por qué es un agente y no un script

Cada paso de este flujo tiene una parte mecánica (miles de filas, fórmulas, XML) y una
parte que requiere criterio (qué mapeo de columnas usar si el sitio cambió un campo,
si un desvío de margen es de volumen o de valor por unidad, si un resultado hay que
reportarlo o es normal). Lo mecánico lo hacen scripts Python deterministas. El
criterio lo aporta el agente: **Claude Code, orquestando esos scripts a través de
skills** (`.claude/skills/`) que le dan el contexto operativo, las salvedades
aprendidas en producción, y cuándo preguntar en vez de asumir.

No hay ejecución en la nube: el agente necesita Chrome real para loguearse en
Convertia y edita directamente los libros de Excel que ya están en la máquina, así
que corre en local (ver `descargar-reportes-convertia/SKILL.md` para el detalle).

## Los tres skills

| Skill | Qué hace |
|---|---|
| `descargar-reportes-convertia` | Login y descarga vía Playwright de los reportes de brutas/gestión de ventas, por cuenta y rango de fechas |
| `limpiar-reportes-convertia` | Detecta la fila real de encabezados y normaliza las columnas de fecha (sin hora), sin tocar nada más |
| `comentario-operacional` | Lee el libro de seguimiento de una cuenta (hoja `INF` y las de detalle) y redacta el comentario analítico de cierre, con el contexto y las tablas de escalón propias de cada cuenta |

Se invocan pidiéndoselo a Claude Code en lenguaje natural, por ejemplo:

- "Descargá los reportes de Claro CR del 1 de julio al 23 de agosto"
- "Limpiá los reportes que acabo de bajar"
- "Hacé el comentario de cierre de Mundo Pacífico"

Las fechas y la cuenta las elegís vos en el momento; el agente no las tiene fijas.

## Estructura de carpetas

```
descargar_reportes_convertia.py   -- descarga (Playwright)
limpiar_reportes_convertia.py     -- limpieza determinista (openpyxl)
pegar_en_cdm.py / pegar_en_claro.py
                                   -- consolidan lo limpio en el libro de seguimiento
                                      de WOW Perú/Mundo Pacífico y de Claro CR
                                      (edición XML quirúrgica, sin recargar el libro
                                      completo -- ver nota abajo)
reportes_config.json              -- sitios, cuentas y reportes (declarativo)
.claude/skills/                   -- las instrucciones de los 3 skills

Procesados/                       -- reportes ya descargados y limpios (no versionado)
Libros de Cuenta/                 -- el libro de seguimiento real de cada cuenta,
                                      con datos comerciales (no versionado)
Comentarios/                      -- comentarios operacionales en Word (no versionado)
Legado/                           -- Plantilla Unificacion.xlsm, ya no se usa
```

### Sobre `pegar_en_cdm.py` / `pegar_en_claro.py`

Cada libro de seguimiento tiene su propia estructura de hojas, columnas y fórmulas,
así que estos dos scripts están calibrados a mano para el libro de WOW Perú/Mundo
Pacífico y el de Claro CR respectivamente (qué columnas del reporte crudo coinciden
con las del libro, en qué fila hay que desplegar la fórmula, qué pivot cache hay que
marcar para refrescar). Si aparece un libro nuevo o cambia su estructura, hace falta
que el agente lo audite de nuevo antes de escribir el script correspondiente -- no es
un mapeo genérico automático, y no debería serlo: un mapeo mal supuesto deja datos en
la columna equivocada sin que se note.

## Setup

```powershell
pip install -r requirements.txt
playwright install chrome
```

Copiá `.env.example` a `convertia.env` (junto al script) con tus credenciales de
Convertia. `convertia.env` nunca se versiona.

## Uso directo (sin pedírselo a Claude)

```powershell
# Descargar con el rango de fechas por defecto
python descargar_reportes_convertia.py

# Rango explícito
python descargar_reportes_convertia.py --fecha-inicio 2026-07-01 --fecha-fin 2026-08-23

# Limpiar lo descargado
python limpiar_reportes_convertia.py
```

## Qué no está en este repositorio

`convertia.env`, los `.xlsx`/`.xlsm` de `Procesados/`, `Libros de Cuenta/` y
`Legado/`, y los comentarios en `Comentarios/` quedan fuera a propósito
(`.gitignore`): son credenciales y datos comerciales reales de clientes. El
repositorio versiona el agente -- código y skills -- no los datos que produce.
