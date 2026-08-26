---
name: comentario-operacional
description: Redacta el comentario operacional quincenal o de cierre de una cuenta de call center de Convertia a partir de su archivo Excel de indicadores. Úsala siempre que el usuario pida un comentario de corte, un comentario quincenal, un precierre, un comentario de cierre de mes, o el análisis de desempeño de una cuenta, aunque no diga la palabra "comentario". Cubre WOW Perú CPA, Mundo Pacífico Chile, DirecTV Perú y Claro Costa Rica. Úsala también cuando el usuario simplemente adjunte el archivo de una de esas cuentas y diga "vamos con X" o "ahora X".
---

# Comentario operacional

## Qué es esto

El comentario acompaña un cuadro de indicadores que el lector ya tiene delante. No es un
resumen del cuadro: es la lectura analítica de lo que el cuadro no muestra. Quien lo lee
conoce la operación.

## Insumo

Un archivo Excel por cuenta. Las hojas de interés, cuando existan:

| Hoja | Qué esperar |
|---|---|
| `INF` | Cuadro principal: TOTAL, ESTIMACIÓN CIERRE, OBJETIVO, DESVIACIÓN por indicador |
| `Tablas` | Cosecha diaria, ratio diario, tipificación de última gestión, campañas de MKT |
| `Tablas asesor` | Detalle por agente: liquidadas y escalón de pago |
| `Tabla Moviles` | Segmentación fijo contra móvil |

**Los nombres de hoja son una pista, no una garantía.** Si alguna no está, o si están con
otro nombre, recorre todas las hojas del archivo y ubica los bloques por su contenido. El
mismo criterio aplica dentro de la hoja: las tablas no siempre arrancan en A1 ni tienen el
encabezado en la primera fila.

Procedimiento de navegación:

1. Lista las hojas del archivo.
2. Abre primero las cuatro de la tabla de arriba, en ese orden.
3. Para las que falten, recorre el resto buscando por contenido: filas con
   `VENTAS BRUTAS`, `RATIO CONVERSIÓN`, `TOTAL MARGEN`, `CPL`, `CPA`; bloques con
   fechas en columnas; listados de campañas con LEAD y CPL; conteos de
   `RESULT_ULTIM_GESTION`.
4. Si un bloque no aparece en ninguna hoja, sigue sin él. No lo menciones.

## Datos faltantes

Trabaja con lo que haya. Si falta un indicador, **omítelo en silencio**: no escribas que
falta, no dejes un hueco, no lo estimes. El comentario debe leerse completo aunque se haya
construido con la mitad de los bloques.

Excepción única: si el usuario ya te advirtió que una cifra está mal calculada o
incompleta, y aun así quiere comentarla, menciónala una sola vez al inicio como estimación
preliminar y sigue. Nunca hagas énfasis en la falta de datos.

Nunca inventes una cifra ni la deduzcas de un supuesto no verificable.

## Procedimiento analítico

1. Ubica la fecha de corte y el mes en curso.
2. Trabaja siempre sobre **ESTIMACIÓN CIERRE contra OBJETIVO**. Nunca comentes el
   acumulado parcial.
3. Compara contra el corte anterior si el usuario lo dio o si está en el archivo.
4. Determina el escalón alcanzado con la tabla de la cuenta y calcula la brecha en
   unidades contra el escalón siguiente. Traduce esa brecha a ventas por día con los
   días que faltan, y avisa si hay festivos de por medio.
5. Identifica **un** driver principal de la desviación de margen. Descompón entre
   volumen de leads, conversión y valor por unidad. No atribuyas a un solo factor sin
   descartar los otros dos.
6. Busca en las hojas de detalle lo que el cuadro no muestra: cosecha diaria por
   tramos, tipificación de la entrada, eficiencia por campaña, mix de producto, RGU
   por liquidada.
7. Redacta según la estructura de abajo.

## Diagnóstico: cómo separar volumen, conversión y valor

Esta es la parte que aporta el comentario. Antes de escribir, resuelve estas tres:

**¿Falta entrada?** Compara leads estimados contra objetivo. Si los leads cumplen y las
brutas no, el problema no es de captación.

**¿Falta conversión?** Si el ratio está en meta pero faltan brutas, la causa es la entrada.
Si el ratio cae con leads sobre objetivo, revisa la tipificación antes de atribuirlo a
gestión: leads de atención al cliente, fuera de cobertura, buzón o no califica inflan el
denominador sin ser oportunidad comercial. Recalcula el ratio depurando esos bloques y
compara contra la meta. Si depurado cumple, el problema es de composición de leads, no de
cierre, y así debe decirse.

**¿Falta valor por unidad?** Si el volumen cumple y el ingreso no, mira ticket, RGU por
liquidada y mix de producto. Un producto nuevo con ticket menor baja el promedio aunque
suban las unidades.

Cuidado con dos trampas recurrentes:

- Conversión a liquidada por encima de lo normal en cortes tempranos significa ventas de
  periodos anteriores aún cerrando, o ventas recientes sin madurar. En ambos casos el
  indicador se normaliza después: trátalo como techo, no como base.
- Un CPL bajo no es eficiencia si está sostenido por volumen que no convierte. Mide costo
  por unidad liquidada, no por lead.

## Escalones del factor de pago

**WOW Perú CPA.** Sobre ventas brutas proyectadas.
Cortes: 1.000 / 1.100 / 1.250 / 1.350 / 1.550 / 1.700 y más. El esquema subió 5% en julio 2026.

**Mundo Pacífico.** Sobre brutas proyectadas, separado por producto.
Fijo: hasta 1.500 factor 2 · 1.501 a 2.000 factor 2,3 · 2.001 a 2.500 factor 2,5 · 2.501 y más factor 2,6.
Móvil: hasta 500 factor 2 · 501 a 1.000 factor 2,3 · 1.001 a 1.500 factor 2,5 · 1.501 y más factor 2,6.

**DirecTV Perú.** El escalón se define sobre **liquidadas proyectadas**, no sobre brutas.
Tabla de comisión por volumen en soles por unidad, con cortes cada 50 unidades.

**Claro Costa Rica.** Cada liquidada se clasifica en escalón de pago 1, 2 o 3.

El ingreso final resulta del ticket promedio aplicado sobre las liquidadas, no sobre las
brutas. Un CPA bajo no implica rentabilidad: el margen depende del ingreso por unidad, que
lo define el escalón.

## Responsables

Cada comentario cierra con dos bloques, operación primero y MKT después.

| Cuenta | Operación | MKT |
|---|---|---|
| WOW Perú CPA | +Andrea | +Sergio |
| Mundo Pacífico | +Andrea | +Sergio |
| DirecTV Perú | +Fernanda | +Andreina |
| Claro Costa Rica | +Andrea | +Andreina |

Cada bloque nombra **una palanca concreta**, no un resumen de lo ya dicho.

## Estructura del comentario

Tres o cuatro párrafos breves, de dos o tres frases cada uno, más los bloques de cierre.

1. Margen e ingreso estimado contra objetivo, y el driver de la desviación.
2. Escalón y brecha en unidades, o el frente donde se está perdiendo (conversión,
   liquidación o valor por unidad).
3. Lo que no se ve en el cuadro: cosecha diaria, tipificación de la entrada, mix.
4. Marketing: CPL, eficiencia por campaña y costo por unidad liquidada.

Si el corte es temprano, agrega una nota de cautela de una o dos frases antes de los
bloques de responsable.

## Reglas de estilo

- Prosa analítica. **Sin guiones largos y sin paréntesis explicativos**, el usuario los
  considera lenguaje de IA.
- Párrafos cortos y concisos. Media página o poco más en total.
- No releas cifras que ya están en el cuadro. Aporta lectura y causa.
- No expliques mecánicas básicas ni definas indicadores.
- No compares cuentas distintas dentro de un mismo comentario.
- No detalles los costes operativos, a veces no son cifras firmes. Sí habla de margen y CPL.
- Cifras en formato local: coma decimal y punto de miles.
- Si el usuario aporta contexto de una reunión, intégralo como causa dentro del análisis,
  no como una sección aparte, y contrástalo con los datos antes de darlo por bueno.

## Contexto por cuenta

Léelo como antecedente, no lo repitas si el corte no lo confirma.

**WOW Perú CPA.** Los canales de WhatsApp captan barato pero convierten bajo 9%. M_EX,
M_05_GEN, PLANES y TELEFONO son los de mejor cierre. Desde mitad de julio 2026 el cliente
puja en Google por su cuenta y desde agosto se puja alto para competir con Movistar, lo que
abarata el CPL pero trae entrada menos calificada. Fuera de cobertura y consultas de
atención concentran buena parte de los leads.

**Mundo Pacífico.** EX concentra cerca de la mitad de los leads y entrega la liquidada al
doble del costo de TELEFONO, que tiene el CPL más bajo pese a peor conversión. TD es
orgánico y no se escala con inversión. Por acuerdo con el cliente no pueden pujar alto ni
pasar del 40% de cuota superior absoluta, así que el volumen de EX tiene techo y la palanca
disponible es el costo por unidad facturada. El cuadro principal puede subestimar las
liquidaciones: contrasta contra el CDM Financiero.

**DirecTV Perú.** Facebook aporta más del 60% de los leads pagados a CPL bajísimo y no
convierte. Google sostiene la generación de bruta. La columna de ventas por campaña del
archivo de MKT no cuadra con las liquidadas reales: usa leads, costo y CPL, no conteos de
venta por campaña. Las etiquetas de producto están invertidas, "fijo" es televisión. Desde
agosto 2026 la cuenta opera con dotación mínima y tiempos de respuesta largos.

**Claro Costa Rica.** Cerca de la mitad de los leads cierra como atención al cliente y
diluye el ratio. M_INTERNET es el motor de la cuenta. PMAX genera tráfico que no llega a
registrarse como lead. Desde agosto 2026 se vende móvil, con ticket bastante menor que
fijo, así que suben las unidades y baja el ticket promedio. No comentes desempeño por
agente: no se tiene fecha de ingreso de cada uno.

## Verificación antes de entregar

- Toda cifra sale del archivo o de un cálculo sobre él.
- El comentario habla de estimación de cierre, no de acumulado parcial.
- Hay un driver identificado, no una lista de variaciones.
- Ningún guion largo ni paréntesis explicativo.
- No se menciona ningún dato faltante.
- Están los dos bloques de responsable, con la palanca concreta en cada uno.
