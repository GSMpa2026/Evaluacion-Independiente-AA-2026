# Lecciones aprendidas (evidencia y técnica)

## Evidencia y análisis
- **Fechas frente al período.** Un acta, informe o registro de alertas anterior al inicio del período es contexto o base normativa, no evidencia de operación en el período. En Ancón, el acta del 17-04-2024 y las alertas de oct-2023 a abr-2024 eran anteriores al período (jul-2024 a jun-2025).
- **Evidencia posterior al período.** Las evaluaciones fechadas después del cierre (por ejemplo, el 21-07-2025) no sustentan la gestión del período. Se mencionan como avance.
- **Comité frente a Junta.** Las aprobaciones del Comité no sustituyen la ratificación de la Junta, que la norma exige para el Manual, las políticas y la matriz. Si las actas de Junta están "en consulta", no las des por confirmadas.
- **Informes combinados.** Un informe que cubre dos trimestres excede el máximo de tres meses (Art. 10, num. 14, del Acuerdo 03-2022).
- **Ausencia de ROS.** No es un incumplimiento por sí misma. La brecha es no poder demostrar que responde a un análisis: falta un registro de inusuales y de las decisiones de no reportar.
- **Brecha documental frente a operativa.** Si el texto del Manual dice una cosa y la práctica evidenciada es mejor (barridos de toda la cartera), califica como documental (MEDIO), no como crítica.
- **"Conviene verificar".** No es una brecha; regístralo como OM.
- **Validaciones completadas.** Cuando se completa una prueba antes pendiente (por ejemplo, la muestra de expedientes), actualiza el P e I y la nota de validación. El nivel puede subir.
- **Conteo de entregables.** Cuenta por ítem, no por subítem. Un ítem padre con subítems entregados cuenta como aportado. Si el cliente responde "no se tiene", se clasifica como "no existe según el cliente".
- **Normativa escaneada.** Aplica OCR (`pdftoppm` + `tesseract`) para validar los artículos; los borradores suelen citar mal.

## Técnica docx (docx-js)
- **Tabuladores.** `PositionalTab` no se renderiza en LibreOffice. Usa `tabStops: [{type: TabStopType.RIGHT, position: ancho, leader: LeaderType.DOT}]` con `new Tab()` en encabezado, pie e índice. El ancho útil en vertical es 9360 DXA y en horizontal 12960 DXA.
- **Índice estático en dos pasadas.** Genera el docx, conviértelo a PDF y busca cada título de forma secuencial a partir de la página 3, sin retroceder, porque los títulos también aparecen en las tablas resumen. Escribe `pages.json`, regenera y repite hasta que quede estable.
- **Secciones horizontales.** Usa un `section` separado con `orientation: LANDSCAPE` y su propio encabezado y pie.
- **Cabeceras de tabla.** Las que tienen 3 o 4 letras ("MA", "Puntaje") se parten si la columna es estrecha. Ensancha la columna a 460 DXA o más para "MA" y a 820 DXA para "Puntaje".
- **Contenido de celdas.** Pasa un arreglo de strings con `**negrita**` parseado; el título de la brecha va en negrita dentro de la misma celda.
- **Validación.** Usa el script de validación de la skill docx; debe reportar "All validations PASSED".

## Técnica xlsx (openpyxl)
- Edita **sobre la plantilla original**: conserva estilos copiando `_style` de las filas 4 (blanca) y 5 (gris), los formatos condicionales y las fórmulas.
- Los formatos condicionales se reconstruyen con un rango ampliado: guarda las reglas, reinicia `ws.conditional_formatting` y vuelve a agregarlas.
- Las validaciones de datos en extensiones x14 se pierden. Recrea las listas para B, E y G.
- Al desagrupar o agrupar los títulos (A1:J1 → A1:O1), copia el estilo a las columnas nuevas.
- **Ampliación de rangos.** Reemplaza `$50)` y `$50,`. El primer intento dejó el mapa de calor contando solo 47 de 58 brechas porque faltaba el segundo patrón.
- **Recálculo.** Usa `recalc.py` de la skill xlsx; debe dar `status: success` y `total_errors: 0`. Después lee con `data_only=True` y compara el mapa y los totales con el informe.
