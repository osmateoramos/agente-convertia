---
name: comentario-operacional
description: Redacta el comentario operacional de cierre de una cuenta de call center a partir de su CSV de indicadores mensuales. Úsala cuando el usuario pida un comentario de cierre, un comentario quincenal o mensual, o el análisis de desempeño de una cuenta con archivos del tipo *_indicadores.csv. Cubre WOW Perú CPA y Mundo Pacífico.
---

# Comentario operacional de cierre

## Qué es esto

El comentario acompaña un cuadro de indicadores que el lector ya tiene delante. No es un
resumen del cuadro: es la lectura analítica de lo que el cuadro no muestra.

## Insumo

Un CSV con formato `BLOQUE, INDICADOR, <mes 1>, ... <mes n>`. Cada fila es un indicador,
cada columna un mes. Las últimas filas, bloque `CALENDARIO`, traen el día de cierre y los
días laborables del mes.

Bloques disponibles: LEADS, CONTACTACIÓN, VENTAS, TICKET, CALL CENTER, AGENTES,
MARGENES, CPA, CPL, MARGEN OPERATIVO.

## Procedimiento

1. Carga el CSV y ubica el último mes con dato. Ese es el mes en curso.
2. Si el mes en curso está parcial, proyecta el cierre:
   `estimación = acumulado / días transcurridos * días laborables del mes`.
   Proyecta leads, brutas y liquidadas. Nunca comentes sobre el acumulado parcial.
3. Calcula la variación contra el mes anterior y contra el promedio de los tres
   meses previos, para cada indicador del bloque VENTAS, CPL, CPA y MARGENES.
4. Determina el escalón alcanzado con la tabla de la cuenta y la brecha en unidades
   contra el escalón siguiente.
5. Identifica el driver principal de la desviación de margen. Descompón siempre entre
   volumen de leads, conversión y costo unitario. No atribuyas a un solo factor sin
   verificar los otros dos.
6. Redacta siguiendo la estructura y el estilo de abajo.
## Escalones del factor de pago

**WOW Perú CPA.** El escalón se define sobre VENTAS BRUTAS proyectadas.
Cortes: 1.000 / 1.100 / 1.250 / 1.350 / 1.550 / 1.700 y más.

**Mundo Pacífico.** Tabla sobre brutas proyectadas, separada por producto.
Fijo: hasta 1.500 factor 2 · 1.501 a 2.000 factor 2,3 · 2.001 a 2.500 factor 2,5 · 2.501 y más factor 2,6.
Móvil: hasta 500 factor 2 · 501 a 1.000 factor 2,3 · 1.001 a 1.500 factor 2,5 · 1.501 y más factor 2,6.

El ingreso final resulta del ticket promedio aplicado sobre las liquidadas, no sobre
las brutas. Un CPA bajo no implica rentabilidad: el margen depende del ingreso por
unidad, que lo define el escalón.

## Estructura del comentario

Tres párrafos breves, en este orden, más los bloques de cierre.

1. Margen e ingreso proyectado contra objetivo. Driver principal de la desviación.
2. Escalón alcanzado y brecha en unidades. Conversión a bruta y conversión a liquidada.
3. Bloque de marketing: CPL y comportamiento de la inversión.
Cierra con los bloques de responsable, separando operación y MKT:

```
+<responsable operación>
+<responsable MKT>
```

## Reglas de estilo

- Prosa analítica. Sin guiones largos y sin paréntesis explicativos.
- Media página o poco más. Alrededor de tres párrafos breves.
- No repitas cifras que ya están en el cuadro. Aporta lo que no se ve en él.
- No expliques mecánicas básicas. Quien lee conoce la operación.
- No compares cuentas distintas dentro de un mismo comentario.
- No detalles los costes operativos. Sí habla de margen y de CPL.
- Si el corte es temprano y la proyección tiene alta incertidumbre, incluye una nota
  de cautela explícita.
- Nunca inventes una cifra. Si un indicador no está en el CSV, dilo y sigue.
## Verificación antes de entregar

- Toda cifra citada sale del CSV o de un cálculo sobre el CSV.
- El comentario habla de estimación de cierre, no de acumulado parcial.
- Hay un driver identificado, no una lista de variaciones.
- Están los dos bloques de responsable al final.
