# Metodología de Riesgo GSM (Probabilidad × Impacto)

## Escalas
| Punt. | Probabilidad | Descripción |
|---|---|---|
| 5 | Frecuente | Ocurre en la mayoría de las circunstancias |
| 4 | Probable | Controles o procesos no suficientemente efectivos |
| 3 | Ocasional | Controles no del todo efectivos |
| 2 | Posible | Posibilidad limitada con los controles actuales |
| 1 | Improbable | Controles altamente efectivos |

| Punt. | Impacto | Descripción |
|---|---|---|
| 5 | Catastrófico | Sanciones mayores, pérdidas masivas, daño reputacional crítico |
| 4 | Mayor | Sanciones significativas; intervención de la Alta Gerencia y la Junta |
| 3 | Moderado | Afecta operaciones o el cumplimiento de forma parcial |
| 2 | Menor | Correcciones administrativas menores |
| 1 | Insignificante | Se resuelve con los controles normales |

## Niveles (puntaje = P × I)
| Nivel | Puntaje | Acción requerida |
|---|---|---|
| CRÍTICO | 15–25 | Acción inmediata (0-30 días); reporte a la Junta y al Oficial de Cumplimiento |
| MEDIO-ALTO | 8–12 | Plan de acción formal (30-60 días); monitoreo mensual |
| MEDIO | 4–6 | Medidas correctivas (60-90 días); monitoreo trimestral |
| BAJO | 1–3 | Controles mínimos; monitoreo semestral |
| OPT. DE MEJORA | N/A | No es incumplimiento; mejora preventiva o prospectiva |

Una función en el código es suficiente: `s>=15 CRÍTICO; s>=8 MEDIO-ALTO; s>=4 MEDIO; si no, BAJO`. Coincide con la tabla de la hoja *Variables* (la clave "P:I" va en VLOOKUP).

## Regla de calificación por componente
- **Nivel 1 – CRÍTICO (No Cumple):** dos o más brechas CRÍTICAS.
- **Nivel 2 – ALTO (Brecha Mayor):** una brecha CRÍTICA, o tres o más MEDIO-ALTO.
- **Nivel 3 – MEDIO (Brecha Parcial):** una o dos brechas MEDIO-ALTO.
- **Nivel 4 – BAJO (Cumple):** solo brechas MEDIO o BAJO.

La calificación global se determina con juicio experto sobre la distribución; en Ancón quedó en Nivel 2 con dos componentes en Nivel 1. Calcula los niveles en el código para que nunca se desalineen con la tabla.

## Cómo asignar P e I (criterios usados)
- **Condición confirmada en pruebas y sistémica** (por ejemplo, 88 % de expedientes con observaciones; 96.7 % de colaboradores con riesgo desactualizado): P5.
- **Condición presente y confirmada** (doble función del Oficial, ausencia de una matriz exigida): P4.
- **Observación de la SSRP que persiste sin cronograma de cierre:** I4 por el riesgo de desobediencia; anótalo como "recalificado por persistencia".
- **Omisión potencial de reporte a la UAF:** I5 (posible gravedad máxima).
- **Brecha solo documental, cuando el control opera en la práctica** (por ejemplo, el Manual dice "solo alto riesgo" pero hay barridos de toda la cartera): P2, I3 (MEDIO).
- **Documentación aportada que solo requiere ajustes de forma** (versiones, homologación): P3, I2 (MEDIO).
- **Hechos posteriores al período** (por ejemplo, restricción del efectivo en 2026): P1, I2 (BAJO).
- **"Conviene verificar…" sin evidencia de incumplimiento:** OM, no brecha.
- **Oportunidades de la hoja de ruta de la norma nueva (RM):** OM. En el Excel se les asigna P e I referenciales según la fase: Fase 1 Ocasional/Moderado, Fase 2 Posible/Moderado, Fase 3 Posible/Menor. La columna de calificación muestra "OM".

## Estructura de la matriz Excel (plantilla)
- Hoja **Variables**: escalas (A6:E10 probabilidad, A14:E18 impacto), tabla de clave P:I a nivel (A22:B47) y descripción de niveles (A51:B55).
- Hoja **Riesgos** (fila 3 = encabezado, datos desde la fila 4):
  - **A** Descripción: "ID · Título. Texto".
  - **B** Tipo (Brecha / Oportunidad de Mejora).
  - **C** Documentos.
  - **D** Observación / nota de validación.
  - **E** Probabilidad (lista).
  - **F** = VLOOKUP sobre Variables.
  - **G** Impacto (lista).
  - **H** = VLOOKUP.
  - **I** = P:I, o "OM".
  - **J** = Nivel.
  - **K** Componente (agregada).
  - **L** Base normativa (agregada).
  - **M** = Puntaje (agregada).
  - **N** Plan o Fase (agregada).
  - **O** = Acción requerida según el nivel (agregada).
- Hoja **Matriz de Riesgos**: mapa 5×5 con COUNTIF sobre Riesgos!I, resumen por nivel con COUNTIF sobre J y B, y resumen por componente con COUNTIFS sobre K y J (lo agrega el script).
- Al ampliar filas, sustituye en **todas** las fórmulas tanto `$50)` como `$50,` (los COUNTIF del mapa terminan en `$50,"5:1")`). Recalcula y comprueba que el total de brechas coincida con el del informe.
