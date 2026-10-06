# Redacción humanizada y protección del consultor

## Voz y tono
- Usa la primera persona institucional para lo que el consultor hizo o recomienda: "revisamos", "observamos", "recomendamos", "valoramos".
- Para describir carencias, el sujeto es la entidad o la evidencia, nunca el consultor: "no se cuenta con evidencia de…", "está pendiente de entrega", "Ancón indicó que no existe".
- Antes de las brechas, reconoce las fortalezas concretas: "Queremos destacar…", "Hay una base valiosa sobre la cual construir…".
- Explica por qué importa cada punto y qué pediría el supervisor. Evita el tono alarmista ("prácticamente cierta"); prefiere "es muy probable que…".
- Cierra con una salida positiva cuando exista: la subsanación voluntaria reduce la sanción hasta en 40 %; "Completar ese expediente es una tarea relativamente sencilla y de alto impacto".
- Escribe frases cortas, en párrafos y no en viñetas dentro del texto. Las tablas se reservan para hallazgos, recomendaciones y planes.

## Reglas de protección del consultor
1. **No hay historia interna en el informe.** No se mencionan borradores, versiones internas (v3, "informe JS", "Hallazgo #3 corregido"), correcciones de errores propios, anexos de "ajustes respecto a borradores" ni inconsistencias entre informes parciales. Las inconsistencias se corrigen en silencio con el dato verificable.
2. **No hay avance interno visible.** No aparecen "pendiente de revisar", "consulta TS", % revisado, responsables internos ni cobertura del informe. El Anexo C muestra solo el estado de entrega **del cliente**.
3. **No se afirma la ausencia de algo que el cliente entregó.** Si un documento fue aportado pero no se analizó:
   - No digas "no se recibió".
   - No opines sobre su contenido.
   - Menciónalo como aportado (Anexo A) y formula la brecha sobre la evidencia operativa que falta. Ejemplo: "la política existe, pero no se aportaron el inventario de repositorios ni pruebas de recuperación".
4. **No se sostienen afirmaciones no verificadas.** Si un dato está en consulta con el cliente, atribúyelo: "según la información aportada por Ancón…" o "sujeto a confirmación con la Oficial de Cumplimiento".
5. **No se mencionan trabajos previos del consultor con el cliente** (por ejemplo, haber revisado el Manual años atrás), salvo que el usuario lo pida. La declaración de independencia es neutral (Sección 3.4). Si el usuario elimina esa mención, sugiérele en el chat, no en el informe, que tenga preparada una respuesta por si el supervisor pregunta.
6. **Delimita el alcance para que el consultor no quede expuesto.** La norma publicada después del período se declara fuera del alcance y se ofrece como valor agregado; así nadie puede reprochar que no se evaluó con ella.
7. **No uses notas del autor en los entregables.** En la matriz Excel, la columna D se llama "Observación / Nota GSM (actualizada [mes-año])" y contiene el estado de validación, no comentarios internos (como "criterio revisado por GSM en 2023" o "coordinar con TS").

## Tabla de reemplazos
| Evitar | Usar |
|---|---|
| No recibimos X | X está pendiente de entrega |
| X no nos fue entregado para revisión | No se aportó copia de X |
| No pudimos validar | No puede verificarse con la evidencia disponible |
| No encontramos X | No se cuenta con evidencia de X |
| No verificamos que… | No se cuenta con evidencia de que… |
| Confirmamos que (dato en consulta) | Según la información aportada por [entidad]… |
| [Documento] no recibido (cuando sí se recibió) | [Documento] aportado; falta evidencia de su aplicación |
| La Junta ratificó (sin confirmar) | Se aportaron actas de Junta cuyo alcance está sujeto a confirmación |
| Recibido / recibida (en notas) | Aportado / aportada |
| "El borrador decía…", "corregimos…" | (omitir) |

## Verificación final (grep)
Sobre el texto extraído del .docx (`extract-text` o `pandoc`) y sobre las columnas A y D del Excel:
```bash
grep -n -i -E "no recibimos|recibimos|no nos |pudimos|no encontramos|no verificamos|consulta TS|pendiente de revisar|borrador del informe|informe JS|v3|autorrevisi|participó en 2023" informe.txt
```
Revisa también "no se recibió": solo es aceptable si el cuadro de control confirma que el cliente no lo entregó, y aun así es preferible "pendiente de entrega".

## Fórmulas de redacción útiles
- **Alcance:** "El Acuerdo SSRP 07-2025 está fuera del alcance de esta evaluación. Se aprobó el … y se publicó en Gaceta Oficial el …, después del cierre del período. Por eso no lo utilizamos como criterio para calificar los componentes ni emitimos opinión sobre el cumplimiento de [entidad] respecto de él. Dado que hoy es la norma aplicable, en la Sección 11 presentamos, como valor agregado, un catálogo de mejoras y una hoja de ruta para la transición."
- **Dos en uno:** "Lo organizamos para que [entidad] obtenga dos productos en uno: la evaluación del período (Parte I)… y una hoja de ruta hacia el Acuerdo SSRP 07-2025 (Parte II)…"
- **Aclaraciones del cliente:** "Según la aclaración de Ancón, en los trimestres II a IV el Comité sesionó para dar seguimiento, sin nuevas aprobaciones; las copias de esas actas, sin embargo, no han sido aportadas, de modo que su celebración no puede verificarse documentalmente."
- **Recalificable:** "Al aportarse, este componente puede recalificarse."
- **Causa de fondo:** "la distancia entre lo que está aprobado y lo que se puede demostrar que se ejecuta".
