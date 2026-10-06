---
name: evaluacion-independiente-pbc-seguros
description: Elabora el Informe Final de Evaluación Independiente del sistema de prevención BC/FT/FPADM (AML/CFT/PBC) de una aseguradora o sujeto obligado en Panamá, en Word, con matriz de riesgo GSM en Excel (Probabilidad × Impacto, mapa de calor) y hoja de ruta hacia el Acuerdo SSRP 07-2025. Integra informes parciales, cuadros de control de entregables, entrevistas, pruebas de expedientes, actas de Comité y normativa (Ley 23 de 2015, D.E. 35 de 2022, Acuerdos SSRP 03-2022 y 07-2025). Úsala siempre que el usuario pida consolidar, redactar, actualizar o "humanizar" un informe de evaluación independiente de cumplimiento, una matriz de riesgo de hallazgos AML, un plan de acción de remediación regulatoria o un roadmap de adecuación normativa, aunque no diga "evaluación independiente" (p. ej., "actualiza el informe con el nuevo cuadro de control", "genera la matriz de riesgo con base en el Excel", "que el informe no golpee al consultor").
---

# Evaluación Independiente PBC/FT/FPADM – Sector Seguros (Panamá)

Esta skill reproduce el método con el que se construyó el informe final de la Evaluación Independiente 2024-2025 de Aseguradora Ancón, S.A. (consultor: GSM Consulting). Sirve para nuevas evaluaciones del mismo tipo y para actualizar un informe existente cuando llega nueva evidencia.

El entregable típico son dos archivos coherentes entre sí:
1. **Informe Final Consolidado (.docx)**: índice, objetivos, alcance, marco normativo, metodología, desarrollo por componente, matriz de riesgo, conclusiones, recomendaciones estratégicas, plan de acción y hoja de ruta de valor agregado.
2. **Matriz de Riesgo (.xlsx)**: hojas *Variables*, *Riesgos* y *Matriz de Riesgos*, con fórmulas, mapa de calor y resumen por componente.

## Principios que no se negocian

- **Alcance temporal estricto ("2 por 1").** La Parte I califica el período evaluado solo con la norma vigente en ese período, por ejemplo el Acuerdo 03-2022 para jul-2024 a jun-2025. Una norma publicada después, como el Acuerdo 07-2025, queda **fuera del alcance** y se declara así en un recuadro de la Sección 3.1. Esa norma se aborda en la Parte II como *valor agregado*, mediante un catálogo de mejoras y una hoja de ruta por fases. Nunca mezcles criterios: ninguna brecha de la Parte I cita la norma posterior.
- **Proteger al consultor.** El informe es para el cliente y el supervisor. No puede exponer al consultor. Estas son las reglas mínimas; el detalle está en `references/redaccion-y-proteccion.md`:
  - No mencionar borradores previos, errores corregidos, versiones internas, avance interno de revisión, "consulta TS" ni "pendiente de revisar".
  - No mencionar trabajos previos del consultor con el cliente, salvo que el usuario lo pida.
  - Nunca afirmar que un documento "no se recibió" si el cliente lo entregó y solo está pendiente de análisis. En ese caso, no opinar sobre su contenido.
  - Evitar la primera persona negativa ("no recibimos", "no pudimos"). Usar "pendiente de entrega", "no se cuenta con evidencia" o "Ancón aportó".
- **Evidencia antes que conclusión.** Cada brecha indica la fuente (prueba, entrevista, documento) y la base normativa exacta. Valida cada artículo citado; los errores frecuentes están en `references/marco-normativo.md`.
- **Tono humanizado y constructivo.** Usa primera persona institucional ("observamos", "recomendamos"). Reconoce las fortalezas antes de las brechas y explica por qué importa cada punto. Al cerrar el riesgo de cada componente, indica qué solicitaría previsiblemente la SSRP.

## Flujo de trabajo

### 1. Inventario y lectura de insumos
Lee todo antes de redactar. Para leer archivos, consulta primero las skills de lectura (file-reading, docx, xlsx, pdf). Los insumos típicos son:
- Informes parciales: gobierno corporativo, políticas, entrevistas, expedientes de clientes y de colaboradores.
- Borrador de presentación con calificaciones por componente.
- Cuadro de control de entregables (Excel): estatus de cada requerimiento.
- Aclaraciones del cliente, por ejemplo la trazabilidad de aprobaciones en actas de Comité.
- Informe de inspección de la SSRP y matriz de seguimiento de su cierre.
- Plantilla de matriz de riesgo GSM (`assets/plantilla_matriz_riesgo_GSM.xlsx`).
- Normativa en PDF. Si está escaneada, aplica OCR para validar el articulado.

### 2. Conciliar evidencia y fechas
Las reglas completas están en `references/lecciones-aprendidas.md`. Las más importantes:
- Compara la fecha de cada acta, informe o alerta con el período. Una aprobación anterior al período es base normativa, no evidencia de supervisión dentro del período.
- Distingue entre aprobación del **Comité** y ratificación de la **Junta Directiva**: una aclaración del Comité no sustituye a las actas de Junta.
- Clasifica cada entregable, siempre desde el lado del cliente: *aportado*, *pendiente de entrega* o *no existe según el cliente*. Cuenta por ítem; los subítems no se cuentan.
- Detecta inconsistencias numéricas entre informes parciales y corrígelas en silencio usando el dato verificable.

### 3. Identificar brechas por componente
Los componentes van en este orden fijo:
1. Gobierno Corporativo (GC)
2. Políticas y Procedimientos (PP)
3. Programa Anual de Cumplimiento (PAC)
4. Monitoreo Transaccional (MT)
5. Debida Diligencia (DD)
6. Evaluación de Riesgo AML (ER)
7. Reportes Regulatorios (RR)
8. Conservación de Documentos (CD)

Cada brecha lleva: ID, título corto, texto humanizado, base normativa, P, I, documentos, nota de validación y referencia al plan de acción. Lo que solo "conviene verificar" se registra como **Oportunidad de Mejora (OM)**, no como brecha.

### 4. Valorar en la matriz de riesgo GSM
Aplica la metodología de `references/metodologia-riesgo.md`. El nivel se calcula como P×I: 1–3 BAJO, 4–6 MEDIO, 8–12 MEDIO-ALTO y 15–25 CRÍTICO. La calificación del componente (Nivel 1–4) se deriva con la regla documentada. Recalifica cuando cambia la evidencia: una validación completada sube o baja el puntaje y se anota en la columna de validación.

### 5. Redactar el informe
Sigue la estructura de `references/estructura-informe.md` (11 secciones y anexos). Construye la Parte II a partir de `references/catalogo-roadmap-07-2025.md` y adapta la columna "situación observada" a los hallazgos reales.

### 6. Generar los archivos
Los scripts están en `scripts/`. El generador de referencia contiene el caso Ancón completo; adapta sus bloques de datos (`C`, `PA`, `RM`, `FASES`, `OMX`) y los textos de las secciones 1, 3, 8 y 9.

```bash
# requisitos: node + docx (npm i docx), python3 + openpyxl, LibreOffice
node scripts/generar_informe.js                      # escribe el .docx y matrix_data.json
python3 scripts/paginas_indice.py Informe.pdf        # tras convertir a PDF; escribe pages.json
node scripts/generar_informe.js                      # 2.ª pasada con números de página
python3 scripts/construir_matriz.py assets/plantilla_matriz_riesgo_GSM.xlsx ANCON_Matriz.xlsx
# recalcula el Excel (LibreOffice) y verifica que tenga 0 errores
```
- Índice: es estático, con números de página calculados en dos pasadas. Repite hasta que `pages.json` quede estable.
- Matriz: conserva las hojas, fórmulas y formatos condicionales de la plantilla. Amplía los rangos a la fila 120 y agrega columnas K–O y un resumen por componente.

### 7. Control de calidad antes de entregar
- Las cifras del informe (resumen ejecutivo, mapa de calor, sección 7) coinciden con el Excel recalculado.
- Ninguna brecha de la Parte I cita la norma fuera de alcance.
- Ninguna frase prohibida aparece en el docx ni en el xlsx. Usa la búsqueda con grep de `references/redaccion-y-proteccion.md`.
- Las afirmaciones sobre el contenido de documentos se limitan a los que ya fueron analizados.
- Renderiza a PDF y revisa visualmente: encabezados de tabla sin cortes, colores de nivel e índice.
- El .docx pasa la validación de la skill docx.

### 8. Entrega
Presenta ambos archivos. En la respuesta, resume en pocas líneas qué cambió y qué debe confirmar el usuario: documentos en consulta, actas pendientes y recalificaciones posibles.

## Referencias (leer según necesidad)
- `references/estructura-informe.md`: contenido de cada sección y tablas de cada componente.
- `references/metodologia-riesgo.md`: escalas, niveles, regla por componente y guía para asignar P e I.
- `references/marco-normativo.md`: mapa de artículos validados, régimen sancionatorio, transición 03-2022 → 07-2025 y errores comunes.
- `references/redaccion-y-proteccion.md`: estilo, tabla de frases a evitar y su reemplazo, y reglas de protección del consultor.
- `references/catalogo-roadmap-07-2025.md`: las 24 mejoras (RM) y su asignación por fases.
- `references/lecciones-aprendidas.md`: trampas de evidencia y problemas técnicos de docx y xlsx ya resueltos.
