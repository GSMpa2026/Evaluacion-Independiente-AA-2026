/*
 * generar_informe.js — Generador de referencia del Informe Final de Evaluación Independiente
 * (caso real: Aseguradora Ancón, S.A., período jul-2024 / jun-2025; versión final v4).
 *
 * Uso:   OUT=Informe_Final.docx node generar_informe.js
 *        (requiere: npm i docx). Lee pages.json si existe (2.ª pasada del índice) y escribe
 *        toc.json (títulos para paginas_indice.py) y matrix_data.json (para construir_matriz.py).
 *
 * Cómo adaptarlo a otra entidad o período:
 *   1. Helpers (paleta, tablas, encabezado/pie): primeras ~110 líneas. Normalmente no se tocan.
 *   2. ENT: conteo de entregables por sección (Anexo C).
 *   3. C[]: un objeto por componente {num, title, situacion[], tablaExtra?, hallazgos[], riesgo, recs[], ruta{}, rm}.
 *      hallazgo = [id, título corto, texto, base normativa, P(1-5), I(1-5), documentos, nota de validación, plan].
 *      El nivel (P×I) y la calificación del componente se calculan automáticamente.
 *   4. PA[] (plan de acción), RM[] (catálogo de la norma nueva), FASES[], OMX[] (oportunidades del período).
 *   5. Textos de las secciones 1, 3, 8 y 9 y del Anexo A dentro de "CONSTRUCCIÓN DEL DOCUMENTO".
 *   Respeta las reglas de references/redaccion-y-proteccion.md al editar textos.
 */
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType,
  HeadingLevel, WidthType, ShadingType, BorderStyle, Header, Footer, PageNumber,
  PageBreak, LevelFormat, PageOrientation, PositionalTab, PositionalTabAlignment,
  PositionalTabRelativeTo, PositionalTabLeader, VerticalAlign, TabStopType, LeaderType, Tab
} = require('docx');

const PAGES = fs.existsSync('pages.json') ? JSON.parse(fs.readFileSync('pages.json')) : {};

// ---------- Paleta ----------
const TEAL = '2F4F5F', ORANGE = 'D9822B', LIGHT = 'E3EAED', GREY = '666666';
const LVL = {
  'Crítico': 'E06666', 'Alto': 'F0A04B', 'Medio': 'FFD966', 'Bajo': '93C47D',
  'Crítica': 'E06666', 'Alta': 'F0A04B', 'Media': 'FFD966', 'Baja': '93C47D'
};
const NIVEL = {
  1: { t: 'Nivel 1 – CRÍTICO (No Cumple)', c: 'E06666' },
  2: { t: 'Nivel 2 – ALTO (Brecha Mayor)', c: 'F0A04B' },
  3: { t: 'Nivel 3 – MEDIO (Brecha Parcial)', c: 'FFD966' },
  4: { t: 'Nivel 4 – BAJO (Cumple)', c: '93C47D' },
};
const FONT = 'Arial';
const PW = 9360;   // ancho útil vertical (Carta, márgenes 1")
const LW = 12960;  // ancho útil horizontal

// ---------- Helpers de texto ----------
function runs(text, o = {}) {
  const parts = String(text).split('**');
  return parts.map((s, i) => new TextRun({
    text: s, bold: o.bold || i % 2 === 1, italics: o.italics, color: o.color,
    size: o.size || 21, font: FONT
  }));
}
const P = (text, o = {}) => new Paragraph({
  children: runs(text, o), alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { after: o.after ?? 120, before: o.before ?? 0, line: 276 },
  keepNext: o.keepNext
});
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT })], pageBreakBefore: true });
const H1nb = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT })] });
const H2 = (t, pb) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: FONT })], pageBreakBefore: !!pb });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: t, font: FONT })], keepNext: true });
const NUM = (text, ref = 'num') => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(text), alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, line: 276 } });
const BUL = (text) => new Paragraph({ numbering: { reference: 'bul', level: 0 }, children: runs(text), alignment: AlignmentType.JUSTIFIED, spacing: { after: 60, line: 276 } });
const SP = () => new Paragraph({ children: [], spacing: { after: 60 } });

// ---------- Helpers de tabla ----------
const border = { style: BorderStyle.SINGLE, size: 4, color: 'B7C4CA' };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(content, w, o = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map(t =>
    t instanceof Paragraph ? t : new Paragraph({
      children: runs(t, { size: o.size || 17, bold: o.bold, color: o.color }),
      alignment: o.align || AlignmentType.LEFT, spacing: { after: 40, line: 252 }
    }));
  return new TableCell({
    children: paras, width: { size: w, type: WidthType.DXA }, borders,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
    verticalAlign: o.valign || VerticalAlign.TOP, columnSpan: o.span
  });
}
function table(headers, rows, widths, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const trs = [];
  if (headers) trs.push(new TableRow({
    tableHeader: true, cantSplit: true,
    children: headers.map((h, i) => cell(h, widths[i], { fill: TEAL, color: 'FFFFFF', bold: true, size: o.hsize || 17 }))
  }));
  rows.forEach((r, ri) => trs.push(new TableRow({
    cantSplit: o.cantSplit !== false,
    children: r.map((c, i) => {
      if (c && typeof c === 'object' && !(c instanceof Paragraph) && !Array.isArray(c)) {
        return cell(c.t, widths[i], { fill: c.fill, bold: c.bold, align: c.align, size: c.size || o.size, span: c.span, color: c.color });
      }
      const fill = (o.zebra && ri % 2 === 1) ? 'F4F7F8' : (o.firstColFill && i === 0 ? LIGHT : undefined);
      return cell(c, widths[i], { fill, size: o.size, bold: o.firstColBold && i === 0 });
    })
  })));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: trs });
}
const lvlCell = (l) => ({ t: l, fill: LVL[l], bold: true, align: AlignmentType.CENTER });

// ---------- Encabezado y pie ----------
function hdr(w) {
  return new Header({
    children: [new Paragraph({
      children: [
        new TextRun({ text: 'Evaluación Independiente PBC/FT/FPADM  |  Aseguradora Ancón, S.A.', font: FONT, size: 16, color: GREY }),
        new TextRun({ children: [new Tab(), 'Informe Final Consolidado'], font: FONT, size: 16, color: GREY })
      ], tabStops: [{ type: TabStopType.RIGHT, position: w }],
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ORANGE, space: 4 } }
    })]
  });
}
function ftr(w) {
  return new Footer({
    children: [new Paragraph({
      children: [
        new TextRun({ text: 'CONFIDENCIAL – Uso exclusivo de Aseguradora Ancón, S.A.  |  GSM Consulting', font: FONT, size: 16, color: GREY }),
        new TextRun({ children: [new Tab(), 'Página ', PageNumber.CURRENT, ' de ', PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: GREY })
      ], tabStops: [{ type: TabStopType.RIGHT, position: w }]
    })]
  });
}
const portrait = { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1300, left: 1440 } } };
const landscape = { page: { size: { width: 12240, height: 15840, orientation: PageOrientation.LANDSCAPE }, margin: { top: 1200, right: 1440, bottom: 1200, left: 1440 } } };

// ---------- Metodología GSM (Probabilidad x Impacto) ----------
Object.assign(LVL, { 'CRÍTICO': 'E06666', 'MEDIO-ALTO': 'F0A04B', 'MEDIO': 'FFD966', 'BAJO': '93C47D', 'OPT. DE MEJORA': 'BDD7EE' });
const PNAME = { 5: 'Frecuente', 4: 'Probable', 3: 'Ocasional', 2: 'Posible', 1: 'Improbable' };
const INAME = { 5: 'Catastrófico', 4: 'Mayor', 3: 'Moderado', 2: 'Menor', 1: 'Insignificante' };
const nivelPI = (p, i) => { const s = p * i; return s >= 15 ? 'CRÍTICO' : s >= 8 ? 'MEDIO-ALTO' : s >= 4 ? 'MEDIO' : 'BAJO'; };
const ENT = { A: { tot: 74, rec: 54, nr: 17, ne: 3 }, B: { tot: 13, rec: 9, nr: 0, ne: 4 }, C: { tot: 9, rec: 9, nr: 0, ne: 0 } };

// =====================================================================
// PARTE I – COMPONENTES
// hallazgo: [id, título corto, texto, base normativa, P, I, documentos, nota de validación, plan]
// =====================================================================
const C = [];
C.push({
  num: '6.1', title: 'Gobierno Corporativo',
  situacion: [
    'Ancón cuenta con una Junta Directiva activa y con un Comité de Riesgo y Cumplimiento que, durante el período, se rigió por el Reglamento AA_GOBCORP03 (V03, junio de 2019). Queremos destacar que la Oficial de Cumplimiento reporta directamente a la Junta Directiva, lo que le da la posición jerárquica que la función requiere, y que el Departamento incorporó una asistente de apoyo.',
    'El Cuadro de Trazabilidad presentado por Ancón ordena con claridad el ciclo de decisiones: en la sesión del 17 de abril de 2024 el Comité aprobó por unanimidad el Manual, las políticas, la matriz de productos, el manual de monitoreo, los formularios y la reprogramación del Plan 2024; y en 2025 ratificó la actualización de matrices, redefinió excepciones (bloqueos en renovaciones y pagos a terceros por cesión de derechos), aprobó la reprogramación del Plan 2025 y designó al proveedor de esta evaluación. Dentro del período se cuenta, además, con el acta del 28 de mayo de 2025 y con los informes de seguimiento presentados al Comité.',
    'Después del período, la entidad actualizó su marco de gobierno corporativo conforme al Acuerdo SSRP No. 3 de 19 de junio de 2025, separó el Comité de Cumplimiento del de Riesgo y realizó su primera autoevaluación de comités. Valoramos estos avances. La brecha principal está en la evidencia documental de 2024: según la aclaración de Ancón, en los trimestres II a IV el Comité sesionó para dar seguimiento, sin nuevas aprobaciones, pero las copias de esas actas no han sido aportadas. La aclaración, además, se refiere al Comité; las ratificaciones de la Junta Directiva requieren su propia evidencia.'
  ],
  hallazgos: [
    ['GC-01', 'Oficial de Cumplimiento con doble función y sin apoyo profesional', 'La Oficial de Cumplimiento también dirige el área Legal, rol en el que reporta al Gerente General. El Departamento cuenta con una asistente de apoyo administrativo, pero no con personal profesional para monitoreo, debida diligencia ni gestión de alertas. La SSRP ya había señalado esta situación en la inspección DSR-0924-2023 y no se ha corregido.', 'Acuerdo 03-2022, Arts. 5, 7 (num. 11), 8, 10 (párr. final) y 62 (num. 9, lit. b)', 4, 4, 'Organigramas sep-2025; Resolución OAL-362; hojas de vida del Departamento; Informe SSRP DSR-0924-2023', 'Validado: la asistente administrativa no resuelve la dedicación exclusiva ni la falta de apoyo profesional.', 'PA-01'],
    ['GC-02', 'Sin estudio de capacidad del Departamento de Cumplimiento', 'No existe un estudio que demuestre que los recursos humanos y tecnológicos de Cumplimiento son suficientes para el volumen de clientes, alertas y debidas diligencias de la aseguradora.', 'Acuerdo 03-2022, Arts. 5 y 62 (num. 9, lit. d)', 3, 3, 'Organigrama del Departamento de Cumplimiento', 'Descripción de puesto pendiente de entrega; no se cuenta con análisis de capacidad.', 'PA-25'],
    ['GC-03', 'Comité sin calendario ni temario y sin actas entre abril 2024 y mayo 2025', 'El Reglamento permitía sesionar con dos de seis integrantes, incluso sin directores. La convocatoria y la agenda dependían de la Oficial de Cumplimiento, a quien el Comité supervisa, y no existían calendario anual ni temario mínimo. Ancón indica que entre la sesión del 17 de abril de 2024 y la del 28 de mayo de 2025 hubo sesiones de seguimiento sin nuevas aprobaciones, pero sus actas no han sido aportadas, de modo que el mínimo de dos sesiones anuales que el Reglamento fijaba para 2024 no puede verificarse documentalmente.', 'Reglamento AA_GOBCORP03; Acuerdo 03-2022, Arts. 10 (num. 14) y 11', 4, 3, 'AA_GOBCORP03; actas 17-04-2024 y 28-05-2025; Cuadro de Trazabilidad', 'Actas II a IV trimestre 2024 pendientes de entrega (sesiones de seguimiento según la aclaración). Ancón indicó que no existen calendario ni temario.', 'PA-12'],
    ['GC-04', 'Sin informes del Comité a la Junta y autoevaluación sin soporte', 'El Reglamento no definía periodicidad, formato ni contenido del reporte del Comité a la Junta; no se aportaron informes a la Junta, y Ancón indicó que no existe un informe anual de gestión del Comité. La autoevaluación de octubre de 2025 es un buen comienzo, pero sus 19 respuestas afirmativas no tienen soporte, la hoja de resumen marca 0 % por un error de fórmulas y no está firmada ni fue presentada a la Junta.', 'Acuerdo 03-2022, Art. 10 (nums. 14 y 15)', 3, 3, 'AA_GOBCORP03; autoevaluación del Comité 01-10-2025', 'Informes a Junta pendientes de entrega; memorandos a Gerencia: Ancón indicó que no existen.', 'PA-12, PA-26'],
    ['GC-05', 'Código de Ética incompleto y sin política de partes relacionadas', 'El Código de Ética (AIN-28-CODETI, noviembre de 2023) existe, pero no incluye un procedimiento para declarar y gestionar conflictos de interés, reglas para operaciones con partes vinculadas, un canal de denuncia confidencial con protección frente a represalias, un responsable de seguimiento ni consecuencias por incumplimiento; además, lo elaboró Auditoría Interna, que luego debe evaluarlo. Ancón indicó que no tiene una política de transacciones con partes relacionadas, y en las entrevistas el personal no identificó un canal de denuncia anónimo.', 'Manual del Sistema de Control Interno, secc. 6.1; Ley 23 de 2015, Art. 57; Acuerdo 03-2022, Art. 15 (num. 8); Recomendación 18 del GAFI', 3, 3, 'AIN-28-CODETI; entrevistas', 'Política de partes relacionadas: Ancón indicó que no existe.', 'PA-18'],
    ['GC-06', 'Debilidades de trazabilidad en documentos de gobierno', 'El acta de Junta del 20-01-2021 aprueba un reglamento sin identificar código ni versión, no muestra firmas y cita un "Acuerdo 22" que no corresponde; la lista de integrantes del Comité no está al día; el Manual del Sistema de Control Interno V2.0 aún remite a normas de la Superintendencia Financiera de Colombia; y el Manual de RR.HH. no indica quién lo aprobó.', 'Manual del Sistema de Control Interno, secc. 6.1; Acuerdo 03-2022, Arts. 11 y 14', 3, 2, 'Acta JD 20-01-2021; AA_GOBCORP03; AIN-28-MSCI-1; GESREHU-1', 'Revisado.', 'PA-22'],
    ['GC-07', 'Planificación de Auditoría Interna sin enfoque de riesgo explícito', 'El Plan de Auditoría Interna 2024 no muestra una planificación basada en riesgos ni los elementos formales de aprobación, alcance, recursos y reprogramación de trabajos. Como aspecto positivo, Ancón aportó un informe anual de Auditoría Interna al Departamento de Cumplimiento; recomendamos que la revisión de Cumplimiento quede programada de forma expresa en cada plan anual.', 'Ley 23 de 2015, Art. 45; Acuerdo 03-2022, Art. 18', 2, 3, 'Planes de Auditoría Interna 2024 y 2025; Informe de Auditoría Interna al Dpto. de Cumplimiento', 'Informe de Auditoría Interna al Dpto. de Cumplimiento aportado.', 'PA-21'],
    ['GC-08', 'Sin suplente ni continuidad de la función de prevención', 'No hay un suplente designado ni un plan que asegure la continuidad de la función de prevención si la Oficial de Cumplimiento se ausenta.', 'Acuerdo 03-2022, Art. 9', 2, 3, 'Organigrama; Manual de Prevención', 'No evidenciado.', 'PA-27'],
    ['GC-09', 'Vacantes en funciones que sostienen el modelo de prevención', 'Están vacantes o concentradas las posiciones de Oficial de Protección de Datos, Gerencia de Recursos Humanos y una gerencia propia para Emisión.', 'Ley 81 de 2019; Acuerdo 03-2022, Art. 16', 3, 2, 'Organigrama general; entrevistas', 'Confirmado en entrevistas.', 'PA-27'],
    ['GC-10', 'Designación del Oficial sin acta de Junta ni descripción de puesto', 'Para la designación de la Oficial de Cumplimiento se cuenta con la Resolución OAL-362 de 30 de diciembre de 2022; el acta de Junta Directiva que la designa y la descripción del puesto están pendientes de entrega.', 'Acuerdo 03-2022, Arts. 5 y 8', 2, 3, 'Resolución OAL-362', 'Acta de Junta y descripción de puesto pendientes de entrega.', 'PA-01'],
    ['GC-11', 'Ratificación de la Junta Directiva sin evidencia concluyente', 'La aclaración de Ancón documenta las aprobaciones del Comité, pero la ratificación por la Junta Directiva del Manual de Prevención, de la matriz de riesgo y de la herramienta de listas no está evidenciada, y las actas de Junta aportadas sobre los Programas 2024 y 2025 y el Informe de Gestión 2024 están sujetas a confirmación con la Oficial de Cumplimiento.', 'Acuerdo 03-2022, Arts. 11, 12, 13, 14 y 15 (aprobación del Comité y ratificación de la Junta)', 3, 3, 'Aclaración de aprobaciones del Comité; actas de Junta aportadas', 'Actas de Junta del Manual, la matriz y la herramienta de listas pendientes de entrega; las demás, sujetas a confirmación.', 'PA-10'],
  ],
  riesgo: 'Bajo el Acuerdo 03-2022, que el Oficial de Cumplimiento no tenga dedicación exclusiva es una infracción de gravedad media (Art. 62, num. 9, lit. b), con multas de más de B/.15,000 y hasta B/.1,000,000. Como la SSRP ya lo había observado, mantener la situación puede considerarse desobediencia (Art. 59, num. 4), lo que lleva la multa al tope del rango. En la próxima inspección, lo más probable es que la SSRP pida la decisión de la Junta sobre la separación de funciones, la dotación aprobada y las actas de todas las sesiones del Comité.',
  recs: [
    ['R-GC-01', 'Separar las funciones de Cumplimiento y Legal para que el Oficial tenga dedicación exclusiva y a tiempo completo, formalizar su designación en acta de Junta con su descripción de puesto y cumplir el trámite ante la SSRP si cambia el titular (ver RM-02).', 'Crítica'],
    ['R-GC-02', 'Realizar un estudio de capacidad y dotar al Departamento de al menos un analista profesional dedicado a PBC/FT/FPADM, con aprobación de la Junta.', 'Alta'],
    ['R-GC-03', 'Confirmar que el reglamento del nuevo Comité de Cumplimiento fije quórum con mayoría de directores, presidencia y secretaría a cargo de directores, convocatoria por el presidente, calendario trimestral y temario mínimo, y que cada sesión quede en acta firmada.', 'Alta'],
    ['R-GC-04', 'Establecer un informe trimestral del Comité a la Junta y un informe anual de gestión; ajustar la autoevaluación (soporte por respuesta, fórmulas corregidas, firma) y presentarla a la Junta con un plan de acción.', 'Alta'],
    ['R-GC-05', 'Completar el Código de Ética y aprobar una política de transacciones con partes relacionadas; implementar un canal de denuncia confidencial con opción de anonimato; trasladar la titularidad del Código a Cumplimiento o Recursos Humanos.', 'Alta'],
    ['R-GC-06', 'Incluir en el Plan de Auditoría Interna, con metodología basada en riesgos, la revisión periódica de Cumplimiento y de las herramientas AML/CFT.', 'Alta'],
    ['R-GC-07', 'Llevar un registro maestro de documentos de gobierno y cumplimiento, con actas firmadas que identifiquen código, versión y fecha de lo aprobado.', 'Media'],
    ['R-GC-08', 'Designar un suplente del Oficial de Cumplimiento, documentar la continuidad de la función y cubrir las vacantes de Protección de Datos, Recursos Humanos y Emisión.', 'Media'],
  ],
  ruta: { '0–30 días': 'PA-01, PA-02, PA-10', '30–90 días': 'PA-12, PA-18, PA-21', '90–180 días': 'PA-25, PA-26, PA-27' },
  rm: 'RM-02, RM-03, RM-18'
});

C.push({
  num: '6.2', title: 'Políticas y Procedimientos',
  situacion: [
    'Revisamos el Manual de Prevención MPREMET-21 (v3.0, octubre de 2023) y las políticas que lo complementan: Conoce al Consumidor (CONCLI-15), Herramientas Tecnológicas (GESADHE-12), Perfil Financiero y Transaccional (MPRMOPTR-20), Capacitación (CAPAC-8 y CAPPCCASC-13), Reserva y Confidencialidad (RESYCON-7), Conoce a tu Colaborador (CONASEM-5 y DPRPACO-17), Sujetos Regulados del Grupo B (CONSUREG-10 y DPRPASO-19), Sujetos Obligados del Grupo A (CONCOMGF-11) y Proveedores (CONTUPROV-6 y DPRPAPR-18). Todos citan correctamente el Acuerdo 03-2022 y fueron aprobados por el Comité el 17 de abril de 2024.',
    'La base documental tiene fortalezas claras: el Manual cubre los tres niveles de debida diligencia, la clasificación de riesgo y el tratamiento de PEP; CONCLI-15 distingue bien entre persona natural y jurídica; CONASEM-5 incluye señales de alerta sobre la conducta del colaborador; las políticas de capacitación segmentan bien a sus destinatarios; RESYCON-7 incorpora la Ley 81 de 2019 y controles como DLP, cifrado y bloqueo de USB; y el catálogo de Señales de Alerta 2025 es un documento actualizado y útil.'
  ],
  hallazgos: [
    ['PP-01', 'Manual describe la consulta en listas solo para alto riesgo', 'El Manual describe la consulta en listas OFAC y ONU solo para clientes de alto riesgo, aspecto que la SSRP observó en 2022. En la práctica, Ancón aportó evidencia de la herramienta de consulta, de barridos sobre la cartera existente y de pólizas rechazadas o canceladas por coincidencias, por lo que la brecha es principalmente documental: el Manual debe reflejar que el control se aplica a todos.', 'Ley 23 de 2015, Art. 49; Acuerdo 03-2022, Arts. 21 (num. 2) y 23 (num. 9)', 2, 3, 'MPREMET-21; CONCLI-15; herramienta de listas; constancia de barridos', 'Validado con evidencia de barridos del período: el control opera; falta alinear el texto del Manual.', 'PA-06'],
    ['PP-02', 'Cuadros de control de versiones desactualizados', 'Los cuadros de control de versiones no reflejan la aprobación efectiva de CONSUREG-10, DPRPASO-19, CONCOMGF-11, CONTUPROV-6 y DPRPAPR-18; DPRPACO-17 muestra una fecha de modificación (2019-11) distinta de su vigencia (2023-05); CAPPCCASC-13 tiene una fila vacía; y la política de capacitación aparece con dos códigos.', 'Manual del Sistema de Control Interno, secc. 6.1; Acuerdo 03-2022, Arts. 14 y 15', 3, 2, 'CONSUREG-10; DPRPASO-19; CONCOMGF-11; CONTUPROV-6; DPRPAPR-18; DPRPACO-17; CAPPCCASC-13', 'Revisado.', 'PA-22'],
    ['PP-03', 'Políticas complementarias sin homologar', 'CONSUREG-10 y DPRPASO-19 difieren en contenido; los glosarios de CONASEM-5 y DPRPACO-17 no definen igual la debida diligencia simplificada y reforzada; y la prueba diagnóstica es obligatoria en una política de capacitación pero no en la otra.', 'Acuerdo 03-2022, Arts. 14 y 15 (párr. final)', 3, 2, 'CONSUREG-10; DPRPASO-19; CONASEM-5; DPRPACO-17; CAPAC-8; CAPPCCASC-13', 'Revisado.', 'PA-24'],
    ['PP-04', 'Actualización de colaboradores sin diferenciar por riesgo', 'La política fija una revisión anual para todos los colaboradores, sin importar su riesgo, por lo que la Matriz de Riesgo de Colaboradores no cambia la frecuencia del control. Además, la Política Conoce a tu Colaborador la emite Legal sin participación de Recursos Humanos, que la ejecuta.', 'Ley 23 de 2015, Arts. 26 y 42; Acuerdo 03-2022, Art. 16', 3, 3, 'CONASEM-5; DPRPACO-17', 'Revisado; confirmado en entrevistas.', 'PA-24'],
    ['PP-05', 'Política de Excepción limitada a la suscripción', 'La Política de Excepción del Manual solo cubre la suscripción. Valoramos que el Comité ratificara en 2025 la excepción de pagos a terceros por cesión de derechos, pero los criterios, los aprobadores y el registro de los casos aprobados en el pago de siniestros no están documentados en la política.', 'Acuerdo 03-2022, Arts. 20 y 62 (num. 1)', 2, 3, 'MPREMET-21; Cuadro de Trazabilidad; entrevistas', 'Ratificación 2025 de la excepción confirmada en el Cuadro de Trazabilidad.', 'PA-15'],
    ['PP-06', 'Política de Proveedores sin verificación de lo declarado', 'La Política Conozca a su Proveedor pide el formulario, pero no exige verificar lo declarado (listas, beneficiario final, documentos) según el riesgo del proveedor.', 'Acuerdo 03-2022, Arts. 10 (num. 13) y 15 (num. 9)', 3, 2, 'CONTUPROV-6; DPRPAPR-18', 'Catálogo de proveedores activos pendiente de entrega.', 'PA-24'],
    ['PP-07', 'Política de herramientas tecnológicas desactualizada', 'GESADHE-12 (v1.0, enero de 2023) es la política más antigua del conjunto; no describe el estado real de DEIVID ni los controles alternativos mientras la automatización no esté lista.', 'Acuerdo 03-2022, Arts. 10 (num. 8) y 15 (num. 11)', 3, 3, 'GESADHE-12', 'Revisado.', 'PA-13'],
    ['PP-08', 'Políticas de capacitación sin mínimos medibles', 'No fijan nota aprobatoria, horas ni contenidos mínimos, indicadores de efectividad ni consecuencias para agentes y canales que incumplan.', 'Ley 23 de 2015, Art. 47; Acuerdo 03-2022, Arts. 15 (num. 12) y 16', 2, 2, 'CAPAC-8; CAPPCCASC-13', 'Revisado.', 'PA-24'],
    ['PP-09', 'Restricción del efectivo no incorporada a la política', 'La restricción de pagos en efectivo en Casa Matriz y Transístmica, vigente desde enero de 2026 (hecho posterior), todavía no está en la política escrita.', 'Acuerdo 03-2022, Art. 14', 1, 2, 'Política de medios de pago; entrevista a Cobros', 'Confirmado en entrevistas.', 'PA-24'],
    ['PP-10', 'Sin política de dependencia de terceros', 'Ancón indicó que no cuenta con una política de dependencia de terceros y acuerdos de servicio entre empresas del grupo, que el Acuerdo 03-2022 incluye entre las políticas mínimas, con criterios para prevenir conflictos de interés.', 'Acuerdo 03-2022, Art. 15 (num. 8)', 2, 3, 'Cuadro de control de entregables', 'Ancón indicó que no existe.', 'PA-24'],
  ],
  riesgo: 'Un Manual que no se ajusta al tamaño y la complejidad de la entidad es infracción leve (Art. 61, num. 6, del Acuerdo 03-2022), y no aplicar un Manual basado en riesgo es de gravedad media (Art. 62, num. 6). En este componente la exposición es moderada: la mayoría de las brechas son de forma y de homologación. La SSRP suele verificar que las políticas estén aprobadas, vigentes y alineadas con la práctica, por lo que conviene que el texto del Manual refleje la consulta universal en listas que ya se realiza.',
  recs: [
    ['R-PP-01', 'Ajustar el Manual para que la consulta en listas OFAC, ONU y locales se aplique expresamente a todos los consumidores del servicio de seguros, en la vinculación y en barridos periódicos, como ya ocurre en la práctica.', 'Alta'],
    ['R-PP-02', 'Adoptar un procedimiento de control de cambios normativos y corregir los cuadros de control de versiones.', 'Media'],
    ['R-PP-03', 'Homologar las políticas complementarias y unificar el glosario institucional.', 'Media'],
    ['R-PP-04', 'Diferenciar la frecuencia de actualización de colaboradores según su riesgo, añadir disparadores fuera de ciclo y revisar la política con Recursos Humanos.', 'Alta'],
    ['R-PP-05', 'Extender la Política de Excepción a los pagos de siniestros a terceros, incluir verificación en listas en la Política de Proveedores, formalizar la restricción del efectivo y aprobar una política de dependencia de terceros (aunque sea para declarar que no se utiliza).', 'Media'],
    ['R-PP-06', 'Actualizar GESADHE-12 con el estado real de DEIVID y los controles compensatorios.', 'Alta'],
    ['R-PP-07', 'Incluir en las políticas de capacitación mínimos medibles, indicadores y consecuencias.', 'Media'],
  ],
  ruta: { '0–30 días': 'PA-06', '30–90 días': 'PA-22', '90–180 días': 'PA-24' },
  rm: 'RM-01, RM-05, RM-22'
});

C.push({
  num: '6.3', title: 'Programa Anual de Cumplimiento',
  situacion: [
    'Ancón aportó los programas anuales y cronogramas de 2024 y 2025, el Informe de Gestión Anual 2024 con su resumen ejecutivo y el cronograma de cierre, el correo de remisión del Programa 2025 a la SSRP y el memorial de remisión del Informe de Gestión 2023. Según la aclaración de Ancón, la reprogramación del Plan 2024 se aprobó en el Comité el 17 de abril de 2024 y la del Plan 2025 en las sesiones de 2025. También se aportaron actas de Junta relacionadas con la ratificación de los programas y del Informe de Gestión, cuyo alcance está sujeto a confirmación con la Oficial de Cumplimiento.',
    'A esto se suman cinco informes de seguimiento presentados al Comité (II trimestre de 2024; III y IV trimestre de 2024, en un solo informe de 3 de enero de 2025; I trimestre de 2025; y II trimestre de 2025) y evidencia de capacitaciones introductorias, generales y especializadas, de la formación de la Oficial y de una jornada de sensibilización para la Junta. Es una base sólida; lo que falta es completar la evidencia formal de aprobación y mejorar la medición del avance.'
  ],
  hallazgos: [
    ['PAC-01', 'Evidencia incompleta de remisión y ratificación de los programas', 'La aprobación del Comité de los programas está documentada en la aclaración de Ancón. Sin embargo, el memorial de remisión del Programa 2024 a la SSRP está pendiente de entrega; para 2025 se cuenta con el correo de remisión, pero no con el sello de recibido; y las ratificaciones de la Junta están sujetas a confirmación.', 'Acuerdo 03-2022, Art. 12', 3, 2, 'Programas y cronogramas 2024-2025; aclaración de aprobaciones; correo de remisión 2025', 'Memorial 2024 y sello de recibido 2025 pendientes de entrega.', 'PA-10'],
    ['PAC-02', 'Informe combinado del III y IV trimestre de 2024', 'El seguimiento del III y IV trimestre de 2024 se presentó en un solo informe el 3 de enero de 2025, cubriendo seis meses, más que el plazo máximo de tres meses, y no hay actas que evidencien su conocimiento por el Comité en 2024.', 'Acuerdo 03-2022, Art. 10 (num. 14)', 3, 2, 'Informes de seguimiento al Comité 2024-2025', 'Cinco informes recibidos y revisados.', 'PA-12'],
    ['PAC-03', 'Informe de Gestión 2024 sin acta de aprobación del Comité', 'Ancón elaboró el Informe de Gestión Anual 2024 con su resumen ejecutivo y el cronograma de cierre, lo que es una buena práctica. El acta del Comité que lo aprueba está pendiente de entrega y su ratificación por la Junta está sujeta a confirmación.', 'Acuerdo 03-2022, Art. 13', 2, 2, 'Informe de Gestión Anual 2024', 'Acta de aprobación del Comité pendiente de entrega.', 'PA-10'],
    ['PAC-04', 'Programa sin indicadores de avance', 'El programa no tiene indicadores ni un registro de avance por actividad que ayude a la Junta a entender el nivel de exposición.', 'Acuerdo 03-2022, Arts. 10 (num. 14) y 12', 3, 2, 'Programas y cronogramas 2024-2025', 'Revisado.', 'PA-19'],
    ['PAC-05', 'Capacitación ejecutada sin constancia individual', 'Las capacitaciones se realizaron, pero el 81.8 % de los expedientes de colaboradores revisados no tiene constancia, y Fianzas, Tecnología y Atención al Cliente no reciben capacitación especializada para su función.', 'Ley 23 de 2015, Art. 47; Acuerdo 03-2022, Arts. 16 y 61 (num. 5)', 3, 3, 'Cuadro de capacitaciones 2024-2025; expedientes de colaboradores; entrevistas', 'Capacitaciones evidenciadas; constancias individuales incompletas.', 'PA-20'],
    ['PAC-06', 'Seguimiento de hallazgos SSRP no integrado al programa', 'Ancón preparó una matriz de seguimiento del cierre de hallazgos de la SSRP y de la evaluación independiente anterior, lo que valoramos. Sin embargo, ese seguimiento no está incorporado al Programa Anual como actividades con hitos y evidencia, y varias condiciones observadas en 2022 se mantienen.', 'Acuerdo 03-2022, Arts. 12 y 62 (num. 11)', 3, 3, 'Matriz de seguimiento de cierre de hallazgos (16-01-2026)', 'Matriz de seguimiento aportada.', 'PA-02'],
  ],
  riesgo: 'No ejecutar el programa de capacitación es infracción leve (Art. 61, num. 5) y no actuar a tiempo frente a observaciones de la SSRP es de gravedad media (Art. 62, num. 11). La buena noticia es que la mayoría de los documentos existen; con completar las actas de aprobación, las ratificaciones de la Junta y los acuses de remisión, y archivar las constancias individuales, la entidad quedaría en una posición sólida ante la SSRP.',
  recs: [
    ['R-PAC-01', 'Completar el expediente del Programa Anual de cada año con el acta de aprobación del Comité, el acta de ratificación de la Junta y el acuse de remisión a la SSRP con sello de recibido.', 'Alta'],
    ['R-PAC-02', 'Mantener informes de seguimiento estrictamente trimestrales y reorganizar el programa con indicadores, responsables, fechas y avance por actividad.', 'Media'],
    ['R-PAC-03', 'Aprobar en el Comité el Informe de Gestión Anual antes de su ratificación en Junta y su remisión a la SSRP.', 'Media'],
    ['R-PAC-04', 'Incorporar al programa la remediación de los hallazgos de la SSRP y de esta evaluación como actividades con hitos verificables.', 'Alta'],
    ['R-PAC-05', 'Archivar la constancia individual de capacitación en el expediente de cada colaborador e impartir módulos especializados por función.', 'Alta'],
  ],
  ruta: { '0–30 días': 'PA-10', '30–90 días': 'PA-19, PA-20', '90–180 días': 'Informe de Gestión 2026 aprobado por el Comité (enero de 2027)' },
  rm: 'RM-04, RM-20, RM-24'
});

C.push({
  num: '6.4', title: 'Monitoreo Transaccional', nota: 'Incluye el sistema de monitoreo y la gestión de alertas.',
  situacion: [
    'Hoy el monitoreo transaccional se hace de forma manual y al cierre de cada mes. El sistema DEIVID, cuyo módulo de monitoreo la SSRP encontró "100 % en revisión" en 2022, sigue sin una fecha comprometida de implementación, y el documento de requerimientos del sistema de alertas (DNR Alertas Monitoreos) tiene el plan de pruebas incompleto y las firmas de aprobación en blanco.',
    'Hay una base valiosa sobre la cual construir: trece tipologías parametrizadas con umbrales (primas superiores a USD 50,000, cancelaciones en los primeros 12 meses, pagos a terceros, cambios de beneficiario), un Informe de Tipologías de Lavado de Dinero propio del sector, un Informe de Fraude presentado a la Junta el 9 de junio de 2025, el catálogo de Señales de Alerta 2025, una muestra de alerta de monitoreo y registros de alertas de cancelaciones anticipadas, pagos de reclamos a personas distintas del asegurado y pérdidas totales.'
  ],
  hallazgos: [
    ['MT-01', 'Monitoreo transaccional manual sin cronograma de automatización', 'El monitoreo es manual, mensual y depende del criterio de cada persona. No existe monitoreo automatizado ni un cronograma comprometido para DEIVID, situación que la SSRP observó en 2022 y que se mantiene; por su persistencia, la calificamos con impacto mayor.', 'Acuerdo 03-2022, Arts. 10 (nums. 8 y 9) y 62 (nums. 10 y 11)', 4, 4, 'GESADHE-12; DNR Alertas Monitoreos; Informe SSRP 2022', 'Cronograma formal de DEIVID pendiente de entrega.', 'PA-13'],
    ['MT-02', 'Requerimientos del sistema de alertas sin aprobar', 'El documento de requerimientos del sistema de alertas no está aprobado, su plan de pruebas está incompleto y no tiene fecha de implementación.', 'Acuerdo 03-2022, Arts. 10 (num. 8) y 15 (num. 11)', 2, 2, 'DNR Alertas Monitoreos', 'Revisado.', 'PA-13'],
    ['MT-03', 'Sin listado completo de alertas del período', 'Ancón aportó una muestra de alerta de monitoreo y registros de alertas de octubre de 2023 a abril de 2024. No se cuenta, en cambio, con un listado completo de las alertas generadas y atendidas entre julio de 2024 y junio de 2025.', 'Acuerdo 03-2022, Art. 43; D.E. 35 de 2022, Art. 23', 3, 3, 'Muestra de alerta; registros de alertas oct-2023 a abr-2024', 'Listado completo del período pendiente de entrega.', 'PA-14'],
    ['MT-04', 'Cierre de alertas sin documentar y sin registro de falsos positivos', 'No está documentado cómo se analiza y cierra cada alerta (falso positivo, escalamiento o ROS) ni se registran fecha, área, analista, descripción, resultado, aprobador y fecha de cierre. Ancón indicó que no lleva un listado de alertas cerradas como falsos positivos, y no hay reporte consolidado de alertas al Comité.', 'Acuerdo 03-2022, Art. 43', 4, 3, 'Registros de alertas; cuadro de control de entregables', 'Listado de falsos positivos: Ancón indicó que no existe.', 'PA-14'],
    ['MT-05', 'Sin calibración de reglas y detección manual de PEP', 'No hay una metodología para calibrar reglas y umbrales ni para gestionar falsos positivos; las herramientas AML/CFT no tienen revisión independiente; y la detección de PEP no autodeclarados depende de búsquedas manuales.', 'Acuerdo 03-2022, Arts. 10 (num. 18) y 30 (nums. 1 y 5)', 3, 3, 'Entrevista a Tecnología; GESADHE-12', 'Confirmado en entrevistas.', 'PA-23'],
    ['MT-06', 'Monitoreo de noticias negativas discontinuo', 'El monitoreo de noticias negativas tuvo meses sin cobertura, sus reportes omiten datos mínimos y no se cuenta con un procedimiento formal de revisión de noticias.', 'Acuerdo 03-2022, Art. 10 (num. 9); inspección DSR-0924-2023', 4, 3, 'Reportes de monitoreo de noticias; Informe SSRP 2022', 'Procedimiento formal de noticias pendiente de entrega.', 'PA-14'],
    ['MT-07', 'Reclamos sin verificación del perfil de riesgo (R3 y R4)', 'Los riesgos R3 y R4 (tramitar sin verificar el perfil del cliente ni actualizar su debida diligencia) tienen frecuencia MUY ALTA; el riesgo del cliente no activa revisiones adicionales en el pago de siniestros.', 'Acuerdo 03-2022, Arts. 10 (num. 9), 20 y 30', 4, 3, 'Evaluación de Riesgo de Reclamos v0', 'Confirmado en entrevistas a Reclamos.', 'PA-15'],
    ['MT-08', 'Monitoreo de conducta de colaboradores sin indicadores', 'El monitoreo de conducta de los colaboradores no tiene indicadores sobre alertas generadas y su resultado.', 'Acuerdo 03-2022, Art. 16', 2, 1, 'CONASEM-5', 'Revisado.', 'PA-23'],
  ],
  riesgo: 'Un sistema de detección de operaciones inusuales que no se aplica o no se vigila adecuadamente es infracción de gravedad media (Art. 62, nums. 1 y 10). Al tratarse de un hallazgo de 2022, también podría considerarse desobediencia (Art. 59, num. 4). Esperamos que la SSRP pregunte por el cronograma aprobado de DEIVID, por el listado de alertas del período y por cómo se cerraron.',
  recs: [
    ['R-MT-01', 'Aprobar en Junta un cronograma de DEIVID con hitos mensuales, presupuesto y responsable, con reporte mensual al Comité; completar y firmar el documento de requerimientos.', 'Alta'],
    ['R-MT-02', 'Documentar y aprobar controles compensatorios mientras no haya automatización.', 'Alta'],
    ['R-MT-03', 'Formalizar el registro y cierre de alertas, incluidos los falsos positivos, con reporte trimestral al Comité; aplicarlo también al monitoreo de noticias.', 'Alta'],
    ['R-MT-04', 'Llevar las 13 tipologías y el catálogo de Señales de Alerta 2025 a reglas parametrizadas en DEIVID.', 'Alta'],
    ['R-MT-05', 'En reclamos, verificar perfil y listas del asegurado y del beneficiario antes del pago y exigir revisión de Cumplimiento en pagos significativos de alto riesgo o PEP.', 'Alta'],
    ['R-MT-06', 'Documentar la calibración de reglas, automatizar la detección de PEP e incluir las herramientas AML/CFT en el plan de Auditoría Interna.', 'Media'],
  ],
  ruta: { '0–30 días': 'PA-13 (plan y cronograma)', '30–90 días': 'PA-13, PA-14, PA-15', '90–180 días': 'PA-23' },
  rm: 'RM-16, RM-22'
});

C.push({
  num: '6.5', title: 'Debida Diligencia',
  situacion: [
    'El diseño es sólido: CONCLI-15 está bien estructurada; Sucursales y el área comercial conocen los requisitos por nivel de riesgo; Producción actúa como segunda revisión; y la entidad usa Inspektor y Visor Judicial. Ancón aportó además evidencia de barridos de listas sobre la cartera existente y de pólizas rechazadas o canceladas por coincidencias, lo que demuestra que el control de listas opera. El reto está en la calidad de la ejecución, como muestran las pruebas.',
    'Revisamos **85 expedientes de clientes** (57 personas naturales y 28 jurídicas): 75 (88.24 %) tenían una o más observaciones y 10 (11.76 %) estaban completos. También revisamos **33 de 181 expedientes de colaboradores** (18.2 %), con los resultados de la tabla siguiente.'
  ],
  tablaExtra: {
    headers: ['Atributo evaluado (Conoce a tu Colaborador)', '% de la muestra sin evidencia'],
    rows: [['Nivel de riesgo del colaborador actualizado en los últimos 12 meses', '96.7 %'], ['Constancia de capacitación en PBC/FT/FPADM', '81.8 %'], ['Evaluación de desempeño documentada', '72.7 %'], ['Constancia de entrega del Reglamento Interno de Trabajo', '69.7 %'], ['Contrato de trabajo firmado', '24.2 %']],
    widths: [6860, 2500]
  },
  hallazgos: [
    ['DD-01', 'Observaciones en el 88.24 % de los expedientes de clientes', 'Las más frecuentes: CTC ausentes, incompletos, de otro cliente o en versiones anteriores; formularios sin firma o firmados por terceros; perfil financiero, actividad e información tributaria incompletos; preguntas PEP sin responder; identificaciones ausentes, ilegibles, vencidas o de otra persona; consultas en Inspektor y Visor Judicial sin evidencia o hechas sobre otra persona; y un caso verificado después de emitir. En personas jurídicas, certificados de Registro Público vencidos, avisos de operación sin firma, documentos de otra sociedad y aprobaciones especiales faltantes.', 'Ley 23 de 2015, Arts. 26, 27, 28 y 36; D.E. 35 de 2022, Arts. 12, 13, 17 y 18; Acuerdo 03-2022, Arts. 19, 21 a 24 y 62 (nums. 1 y 4)', 5, 4, 'Matriz de revisión de 85 expedientes; CONCLI-15', 'Validado con la prueba de expedientes.', 'PA-07'],
    ['DD-02', 'Emisión sin perfil de riesgo y canal web con CTC incompleto', 'Hay clientes sin perfil de riesgo en el sistema (hallazgo crítico de la SSRP en 2022) y ningún bloqueo técnico impide emitir sin ese perfil; en el canal web se emiten pólizas de auto, incendio y multirriesgo con el CTC incompleto, sin control posterior.', 'Acuerdo 03-2022, Art. 19; D.E. 35 de 2022, Arts. 16 y 29 (num. 1, lit. f)', 4, 4, 'MPREMET-21; CONCLI-15; Informe SSRP 2022; entrevistas', 'Validado con la prueba de expedientes y entrevistas.', 'PA-05'],
    ['DD-03', 'Asegurado y beneficiario sin perfil de riesgo propio', 'El perfil se centra en el contratante o pagador. Al asegurado distinto solo se le recogen documentos, y al beneficiario de reclamos de vida se le pide cédula y CTC, pero no se le perfila ni se le consulta en listas antes de pagarle.', 'Acuerdo 03-2022, Arts. 20 y 30; Recomendaciones 10 y 12 del GAFI', 4, 3, 'MPREMET-21; entrevistas a Emisión y Reclamos', 'Confirmado en entrevistas.', 'PA-15'],
    ['DD-04', 'Fianzas perfila fuera del módulo institucional', 'Fianzas perfila a sus clientes por su cuenta, con un enfoque de capacidad financiera, fuera del Módulo de Perfil de Riesgo, lo que limita la trazabilidad y la aplicación uniforme de la metodología.', 'Acuerdo 03-2022, Arts. 11 y 25', 3, 3, 'Entrevista a Fianzas', 'Confirmado en entrevistas.', 'PA-16'],
    ['DD-05', 'Procedimiento ante coincidencias en listas no interiorizado', 'Atención al Cliente, punto de primer contacto, no conoce el procedimiento ante una coincidencia en listas, y en algunos expedientes no hay evidencia de la consulta o se hizo sobre otra persona. Aunque los barridos de cartera están evidenciados, esto expone el deber de congelamiento preventivo.', 'Ley 23 de 2015, Art. 49; D.E. 35 de 2022, Art. 29 (num. 1, lit. d); Acuerdo 03-2022, Art. 21 (num. 2)', 2, 4, 'Herramienta de listas; constancia de barridos; entrevistas', 'Barridos evidenciados; procedimiento ante coincidencias por formalizar.', 'PA-06'],
    ['DD-06', 'Aprobaciones especiales faltantes en expedientes', 'La SSRP observó en 2022 aprobaciones especiales vencidas o sin soporte, y en la prueba de expedientes de personas jurídicas encontramos aprobaciones especiales faltantes. Conviene que el listado de clientes con documentación pendiente registre el motivo de cada excepción, su responsable y su fecha límite de subsanación.', 'Acuerdo 03-2022, Art. 19', 3, 3, 'Matriz de expedientes; Informe SSRP 2022; listado de clientes con documentación pendiente', 'Validado con la prueba de expedientes.', 'PA-07'],
    ['DD-07', 'Enfoque basado en riesgo no interiorizado en el área comercial', 'En la Gerencia Comercial de Casa Matriz se aplica la misma diligencia a todos los clientes ("a todos les buscamos lo mismo").', 'Ley 23 de 2015, Art. 26; D.E. 35 de 2022, Art. 7', 3, 2, 'Entrevistas', 'Confirmado en entrevistas.', 'PA-20'],
    ['DD-08', 'Conoce a tu Colaborador desactualizado', 'El nivel de riesgo del colaborador está desactualizado en el 96.7 % de la muestra (mayoría cercana a 20 meses), con brechas en constancias de capacitación, evaluaciones de desempeño, contratos y Reglamento Interno (ver tabla).', 'Ley 23 de 2015, Art. 42; Acuerdo 03-2022, Arts. 16 y 62 (num. 14); D.E. 35, Art. 29 (num. 2, lit. b)', 5, 3, 'Análisis de 33 expedientes de colaboradores; CONASEM-5', 'Validado con la prueba de expedientes.', 'PA-08'],
  ],
  riesgo: 'Es el componente con mayor exposición. No contar con procedimientos de identificación y verificación, o no aplicarlos bien, y no analizar el riesgo del perfil del cliente son infracciones de gravedad media (Art. 62, nums. 1 y 4). El D.E. 35 califica como de gravedad máxima iniciar o mantener una relación con un cliente que no facilite la debida diligencia ampliada (Art. 29, num. 1, lit. f); el canal web y la verificación posterior a la emisión son escenarios donde ese riesgo puede concretarse. Como la SSRP revisa muestras de expedientes (D.E. 35, Art. 6), con el nivel de observaciones encontrado es muy probable que surjan hallazgos.',
  recs: [
    ['R-DD-01', 'Configurar en DEIVID y en el canal web un bloqueo que impida emitir sin perfil de riesgo aprobado y sin CTC completo, firmado y vigente.', 'Crítica'],
    ['R-DD-02', 'Formalizar el procedimiento ante coincidencias en listas (incluido el congelamiento preventivo y la comunicación a la UAF), capacitar a Atención al Cliente y conservar en cada expediente la evidencia de la consulta.', 'Alta'],
    ['R-DD-03', 'Aprobar un plan para subsanar los 75 expedientes observados, empezando por riesgo alto y medio y personas jurídicas, y completar el registro de excepciones con motivo, responsable, fecha límite y escalamiento.', 'Alta'],
    ['R-DD-04', 'Perfilar por separado a contratante, asegurado, beneficiario y tercero con interés legítimo, e integrar a Fianzas en el módulo institucional.', 'Alta'],
    ['R-DD-05', 'Actualizar de inmediato la Matriz de Riesgo de Colaboradores y regularizar contratos y constancias.', 'Alta'],
    ['R-DD-06', 'Revisar cada trimestre una muestra de expedientes (segunda línea), con indicadores de error por área.', 'Media'],
  ],
  ruta: { '0–30 días': 'PA-05, PA-06, PA-07, PA-08', '30–90 días': 'PA-15, PA-16', '90–180 días': 'PA-29, PA-30' },
  rm: 'RM-08, RM-09, RM-10, RM-11, RM-12, RM-13'
});

C.push({
  num: '6.6', title: 'Evaluación de Riesgo AML',
  situacion: [
    'En 2022 la SSRP determinó que la Matriz de Riesgo institucional solo cubría riesgos operativos. Según el Cuadro de Trazabilidad, el Comité aprobó en 2024 la matriz de productos y su metodología y ratificó en 2025 la actualización de las matrices de productos, clientes, proveedores, empleados y provincias. La entidad elaboró además evaluaciones por proceso para Emisión (v2, 21 de julio de 2025), con metodología adecuada, y para Reclamos (v0, 21 de julio de 2025), en borrador, y utiliza una Matriz de Riesgo de Colaboradores en Inspektor.',
    'Las copias de las matrices de clientes, colaboradores y proveedores y de la metodología de clasificación de clientes, así como el acta de Junta que aprueba la matriz, están pendientes de entrega, por lo que su contenido y su integración en una evaluación institucional no pueden verificarse. Al aportarse, este componente puede recalificarse.'
  ],
  hallazgos: [
    ['ER-01', 'Matrices BC/FT/FPADM sin copia aportada y sin evaluación institucional integrada', 'Las matrices de clientes, colaboradores, proveedores y zonas geográficas, su metodología y el acta de Junta que las aprueba están pendientes de entrega, y no se cuenta con una evaluación institucional que integre esos factores con riesgo inherente, controles y riesgo residual. Sin esa evidencia, el hallazgo crítico de la SSRP de 2022 no puede darse por cerrado.', 'Ley 23 de 2015, Art. 26; D.E. 35 de 2022, Art. 7; Acuerdo 03-2022, Arts. 10 (nums. 4 a 6), 11, 17 y 62 (num. 3)', 4, 4, 'Informe SSRP 2022; aclaración de aprobaciones del Comité', 'Aprobadas por el Comité según la aclaración; copias y acta de Junta pendientes de entrega. Recalificable al aportarse.', 'PA-09'],
    ['ER-02', 'Evaluaciones por proceso posteriores al período y en borrador', 'Las evaluaciones de Emisión y Reclamos son de julio de 2025, posteriores al cierre del período; la de Reclamos es un borrador y no consta la aprobación de la de Emisión.', 'Acuerdo 03-2022, Art. 17', 2, 2, 'Evaluación de Riesgo de Emisión v2; Evaluación de Riesgo de Reclamos v0', 'Estatus de aprobación por confirmar.', 'PA-17'],
    ['ER-03', 'Impacto subestimado en R3 y R4', 'Calificar con impacto MUY BAJO los riesgos R3 y R4, de frecuencia MUY ALTA, subestima el riesgo regulatorio y oculta que el control no está funcionando.', 'Acuerdo 03-2022, Arts. 4 (num. 14) y 11', 3, 2, 'Evaluación de Riesgo de Reclamos v0', 'Revisado.', 'PA-17'],
    ['ER-04', 'Canal web sin evaluación de riesgo previa', 'No se cuenta con evidencia de la evaluación de riesgo previa al lanzamiento del canal de autoservicio web.', 'Acuerdo 03-2022, Arts. 10 (num. 7) y 11', 3, 3, 'Entrevistas a Comercialización', 'No evidenciado.', 'PA-31'],
    ['ER-05', 'Evaluación de riesgo sin incorporar resultados de supervisión', 'No hay evidencia de que la evaluación recoja los resultados del monitoreo, de las auditorías, de evaluaciones independientes y de los hallazgos de la SSRP.', 'Acuerdo 03-2022, Art. 17 (párr. 2)', 2, 2, 'Evaluaciones de riesgo por proceso', 'No evidenciado.', 'PA-09'],
  ],
  riesgo: 'No tener una evaluación de riesgos acorde con la realidad de clientes, productos, zonas y canales es infracción leve (Art. 61, num. 7), y no diseñar controles con enfoque basado en riesgo es de gravedad media (Art. 62, num. 3). La matriz institucional es la pieza que justifica todo lo demás, por lo que suele ser lo primero que pide el supervisor. Recomendamos tenerla consolidada y aprobada antes de cualquier inspección.',
  recs: [
    ['R-ER-01', 'Consolidar las matrices aprobadas en una Evaluación Institucional de Riesgo BC/FT/FPADM con metodología documentada (inherente, controles y residual), aprobarla en Comité y Junta y ponerla a disposición de la SSRP.', 'Crítica'],
    ['R-ER-02', 'Formalizar la Evaluación de Reclamos recalibrando R3 y R4 y confirmar la aprobación de la de Emisión.', 'Alta'],
    ['R-ER-03', 'Hacer que la evaluación se alimente formalmente de los resultados del monitoreo, las auditorías y la supervisión.', 'Media'],
    ['R-ER-04', 'Documentar la evaluación de riesgo del canal web y exigir evaluación previa para nuevos productos, canales y tecnologías.', 'Media'],
  ],
  ruta: { '0–30 días': 'PA-09 (entrega y metodología)', '30–90 días': 'PA-09 (aprobación), PA-17', '90–180 días': 'PA-31' },
  rm: 'RM-06, RM-07, RM-14'
});

C.push({
  num: '6.7', title: 'Reportes Regulatorios',
  situacion: [
    'Según la información aportada por Ancón, durante el período no se presentaron ROS a la UAF. En 2022, la SSRP había observado que no se mantenía el registro de operaciones sospechosas, y en 2021 uno de los dos ROS presentados fue rechazado por la UAF por formato. No existe un procedimiento formal de gestión de ROS.',
    'Hay señales positivas: Cobros ha reportado casos a Cumplimiento, existen remisiones documentadas a la SSRP (memorial del Informe de Gestión 2023 y correo del Programa 2025) y Ancón preparó una matriz de seguimiento del cierre de los hallazgos de la SSRP. Sin embargo, en las entrevistas supimos que en 2024 Fianzas detectó una carta de referencia bancaria con indicios de alteración, declinó la operación y **no lo reportó a la Oficial de Cumplimiento**.'
  ],
  hallazgos: [
    ['RR-01', 'Sin ROS en el período ni registro de operaciones inusuales', 'La ausencia de ROS no es en sí misma un incumplimiento, pero no se cuenta con un registro de operaciones inusuales que deje constancia del análisis, del responsable, del sustento y de la decisión de no reportar; la entidad no puede demostrar que esa ausencia responde a un análisis y no a fallas de detección.', 'Ley 23 de 2015, Art. 54; D.E. 35 de 2022, Arts. 23 y 25; Acuerdo 03-2022, Arts. 43, 44 y 62 (nums. 5 y 10)', 4, 4, 'MPREMET-21; Informe SSRP 2022; listado de ROS del período', 'Según Ancón, no se generaron ROS en el período.', 'PA-04'],
    ['RR-02', 'Caso de Fianzas 2024 no escalado a Cumplimiento', 'Declinar una operación no reemplaza el reporte interno: la obligación incluye las tentativas, y sin ese reporte Cumplimiento no pudo evaluar si correspondía un ROS.', 'Ley 23 de 2015, Art. 54; D.E. 35 de 2022, Arts. 25 y 29 (num. 1, lit. e); Acuerdo 03-2022, Arts. 43, 44 y 63 (num. 1)', 3, 5, 'Entrevista a Fianzas', 'Pendiente análisis y decisión documentada por Cumplimiento.', 'PA-03'],
    ['RR-03', 'Sin procedimiento formal de ROS', 'No hay un procedimiento con criterios, plazos, formato, revisión de calidad antes del envío y registro de las decisiones de no reportar.', 'Acuerdo 03-2022, Art. 44', 3, 3, 'MPREMET-21', 'Revisado.', 'PA-04'],
    ['RR-04', 'Condiciones observadas por la SSRP en 2022 aún abiertas', 'Ancón aportó una matriz de seguimiento del cierre de hallazgos, lo que es un paso positivo. Las pruebas de esta evaluación, sin embargo, muestran que las condiciones de fondo observadas por la SSRP —registro de ROS, matriz BC/FT/FPADM, estructura de Cumplimiento, perfiles de riesgo y monitoreo— siguen total o parcialmente abiertas, y no hay evidencia de comunicaciones de seguimiento con la SSRP.', 'Acuerdo 03-2022, Arts. 59 (nums. 1 y 4) y 62 (num. 11)', 3, 4, 'Matriz de seguimiento de cierre de hallazgos; Informe SSRP 2022', 'Matriz de seguimiento aportada; las pruebas muestran condiciones aún abiertas.', 'PA-02'],
    ['RR-05', 'Reportes de efectivo del período no evidenciados', 'Los reportes de efectivo y cuasi efectivo enviados a la UAF durante el período, que debían presentarse dentro de los cinco primeros días hábiles de cada mes, están pendientes de entrega.', 'Ley 23 de 2015, Art. 53; D.E. 35 de 2022, Art. 24; Acuerdo 03-2022, Art. 42', 2, 3, 'Cuadro de control de entregables', 'Pendientes de entrega.', 'PA-11'],
    ['RR-06', 'Sin control centralizado de reportes a la SSRP', 'Hay remisiones puntuales documentadas, pero no se evidencia un control centralizado de reportes y cuestionarios a la SSRP con aprobación previa del Comité.', 'Acuerdo 03-2022, Arts. 10 (num. 21), 12 y 13', 2, 2, 'Memorial Informe de Gestión 2023; correo de remisión Programa 2025', 'Remisiones puntuales evidenciadas.', 'PA-10'],
  ],
  riesgo: 'No comunicar operaciones sospechosas es infracción de gravedad media (Art. 62, num. 5) y, si fuera intencional existiendo fuertes indicios, de gravedad máxima (Art. 63, num. 1). El D.E. 35 también califica como de gravedad máxima no reportar cuando un colaborador o directivo conoció internamente los indicios (Art. 29, num. 1, lit. e); por eso recomendamos analizar y documentar el caso de Fianzas cuanto antes. La norma premia la iniciativa: subsanar por cuenta propia antes de una resolución sancionatoria reduce la sanción hasta en 40 % (Art. 59, num. 1).',
  recs: [
    ['R-RR-01', 'Analizar de inmediato el caso de Fianzas de 2024, documentar el análisis y la decisión y, si procede, presentar el ROS explicando la oportunidad del reporte.', 'Crítica'],
    ['R-RR-02', 'Aprobar un procedimiento de operaciones inusuales y sospechosas con registro único y registro de decisiones de no reportar; enviar una circular recordando que toda señal de alerta se reporta, aunque la operación no se concrete.', 'Crítica'],
    ['R-RR-03', 'Actualizar la matriz de seguimiento del cierre de hallazgos con evidencia por hallazgo, aprobarla en Comité y Junta y valorar compartirla de forma proactiva con la SSRP.', 'Crítica'],
    ['R-RR-04', 'Conciliar los reportes de efectivo y cuasi efectivo del período y conservar sus acuses.', 'Alta'],
    ['R-RR-05', 'Llevar un calendario regulatorio de reportes con responsable, fecha límite, aprobación del Comité y acuse.', 'Alta'],
  ],
  ruta: { '0–30 días': 'PA-02, PA-03, PA-04, PA-11', '30–90 días': 'PA-12, PA-20', '90–180 días': 'Seguimiento permanente en el Programa Anual (PA-19)' },
  rm: 'RM-15, RM-16, RM-17'
});

C.push({
  num: '6.8', title: 'Conservación de Documentos',
  situacion: [
    'Ancón cuenta con una política de conservación de documentos, aprobada por el Comité el 17 de abril de 2024, y RESYCON-7 contempla controles tecnológicos (DLP de Office 365, cifrado y bloqueo de USB) y la Ley 81 de 2019. El reto está en la operación: las pruebas de expedientes muestran problemas de integridad que afectan la capacidad de reconstruir operaciones, y no se cuenta con evidencia de cómo se aplica la política en la práctica.'
  ],
  hallazgos: [
    ['CD-01', 'Aplicación de la política de conservación sin evidencia operativa', 'La política existe, pero no se aportaron el inventario de repositorios, los plazos aplicados por tipo de documento ni pruebas de recuperación, por lo que no puede confirmarse que los registros se conserven al menos cinco años desde que termina la relación ni que las operaciones puedan reconstruirse.', 'Ley 23 de 2015, Art. 29; D.E. 35 de 2022, Arts. 20 y 24; Acuerdo 03-2022, Arts. 15 (num. 13), 33 y 42', 2, 3, 'Política de conservación de documentos; aclaración de aprobaciones', 'Inventario y pruebas de recuperación pendientes de entrega.', 'PA-28'],
    ['CD-02', 'Integridad de expedientes comprometida', 'Los expedientes de clientes contienen documentos ilegibles, vencidos, de otras personas o sociedades y formularios en versiones anteriores; los de colaboradores carecen de contratos y constancias.', 'D.E. 35 de 2022, Art. 20; Acuerdo 03-2022, Art. 33', 4, 3, 'Matriz de expedientes de clientes; análisis de expedientes de colaboradores', 'Validado con las pruebas de expedientes.', 'PA-07'],
    ['CD-03', 'Registros de alertas en hojas de cálculo y sin respaldo evidenciado', 'Las alertas se registran en hojas de cálculo sin control de versiones, y no hay evidencia de copia de seguridad del registro de operaciones en efectivo.', 'Acuerdo 03-2022, Art. 62 (num. 2, lit. d)', 2, 3, 'Registros de alertas', 'No evidenciado.', 'PA-28'],
    ['CD-04', 'Expediente de gobierno incompleto', 'Hay actas sin firma y las actas del Comité de los trimestres II a IV de 2024 están pendientes de entrega.', 'Acuerdo 03-2022, Arts. 11 y 17', 3, 2, 'Actas de Junta y del Comité', 'Actas pendientes de entrega.', 'PA-10'],
    ['CD-05', 'Controles DLP sin reportes a Cumplimiento', 'No se cuenta con evidencia de que los controles DLP generen reportes a la Oficial, y la vacante del Oficial de Protección de Datos dificulta equilibrar conservación y protección de datos.', 'Acuerdo 03-2022, Art. 15 (num. 10); Ley 81 de 2019', 2, 2, 'RESYCON-7', 'No evidenciado.', 'PA-28'],
  ],
  riesgo: 'No conservar los registros por el plazo legal o no tener copia de seguridad del registro de efectivo son infracciones de gravedad media (Art. 62, num. 2). El archivo es, sobre todo, la forma en que la entidad demuestra que cumple: si un expediente no se puede reconstruir, para el supervisor es como si el control no se hubiera hecho.',
  recs: [
    ['R-CD-01', 'Complementar la política de conservación con el inventario de repositorios, plazos por tipo de documento, respaldo del registro de efectivo y pruebas anuales de recuperación.', 'Media'],
    ['R-CD-02', 'Revisar la calidad de los documentos al vincular y ejecutar la subsanación de expedientes prevista en Debida Diligencia.', 'Alta'],
    ['R-CD-03', 'Trasladar el registro de alertas y de operaciones inusuales a un repositorio con control de acceso, versiones y trazabilidad.', 'Media'],
    ['R-CD-04', 'Confirmar que los controles DLP reportan a Cumplimiento y designar al Oficial de Protección de Datos.', 'Media'],
  ],
  ruta: { '0–30 días': 'PA-10 (expediente de gobierno)', '30–90 días': 'PA-07 (registro de excepciones)', '90–180 días': 'PA-28' },
  rm: 'RM-23'
});

// ---------- Cálculos ----------
const LEV = ['CRÍTICO', 'MEDIO-ALTO', 'MEDIO', 'BAJO'];
C.forEach(c => {
  c.cnt = { 'CRÍTICO': 0, 'MEDIO-ALTO': 0, 'MEDIO': 0, 'BAJO': 0 };
  c.hallazgos.forEach(h => { h.nivel = nivelPI(h[4], h[5]); c.cnt[h.nivel]++; });
  c.nivel = c.cnt['CRÍTICO'] >= 2 ? 1 : (c.cnt['CRÍTICO'] >= 1 || c.cnt['MEDIO-ALTO'] >= 3) ? 2 : c.cnt['MEDIO-ALTO'] >= 1 ? 3 : 4;
});
const TOT = { 'CRÍTICO': 0, 'MEDIO-ALTO': 0, 'MEDIO': 0, 'BAJO': 0 };
C.forEach(c => LEV.forEach(k => TOT[k] += c.cnt[k]));
const TOTAL = LEV.reduce((a, k) => a + TOT[k], 0);
const ALLH = C.flatMap(c => c.hallazgos.map(h => ({ comp: c.title, num: c.num, h })));
const HEAT = {}; ALLH.forEach(({ h }) => { const k = `${h[4]}:${h[5]}`; HEAT[k] = (HEAT[k] || 0) + 1; });
const N1 = C.filter(c => c.nivel === 1).map(c => c.title);

const DETERM = {
  '6.1': 'Oficial con doble función; sin actas del Comité entre abr-2024 y may-2025; Código de Ética y partes relacionadas.',
  '6.2': 'Brechas mayormente documentales: Manual vs. práctica de listas, versiones, homologación.',
  '6.3': 'Programas e Informe de Gestión existen; faltan actas de aprobación del Comité y constancias.',
  '6.4': 'Monitoreo manual; DEIVID sin cronograma; sin listado de alertas ni falsos positivos.',
  '6.5': '88.24 % de expedientes con observaciones; emisión sin perfil; colaboradores desactualizados.',
  '6.6': 'Matrices referidas como aprobadas, no entregadas; sin evaluación institucional.',
  '6.7': 'Sin ROS ni registro de inusuales en el período; caso Fianzas 2024.',
  '6.8': 'Política no entregada; integridad de expedientes comprometida.',
};
const PA = [
  ['0–30 días', 'PA-01', 'Llevar a la Junta la decisión de separar las funciones de Cumplimiento y Legal, aprobar al menos un analista PBC dedicado y formalizar en acta la designación y la descripción de puesto del Oficial.', 'GC', 'Crítica', 'Junta Directiva / Gerencia General', '30 días', 'Acta de Junta; perfil y requisición'],
  ['0–30 días', 'PA-02', 'Actualizar la matriz de seguimiento del cierre de hallazgos de la inspección DSR-0924-2023 (entregada el 16-01-2026) con estado y evidencia por hallazgo; aprobarla en Comité y Junta y valorar su envío a la SSRP.', 'RR / GC', 'Crítica', 'Oficial de Cumplimiento', '30 días', 'Matriz actualizada y aprobada; actas'],
  ['0–30 días', 'PA-03', 'Analizar y documentar el caso de Fianzas 2024 y decidir sobre el ROS; enviar una circular sobre el reporte obligatorio de señales de alerta, incluidas tentativas.', 'RR', 'Crítica', 'Oficial de Cumplimiento / Fianzas', '15 días', 'Expediente del caso; decisión firmada; circular'],
  ['0–30 días', 'PA-04', 'Aprobar el procedimiento de operaciones inusuales y sospechosas con registro único, criterios de escalamiento, revisión de calidad del formato UAF y registro de decisiones de no reportar.', 'RR', 'Crítica', 'Oficial de Cumplimiento / Comité', '30 días', 'Procedimiento aprobado; registro en uso'],
  ['0–30 días', 'PA-05', 'Configurar el bloqueo de emisión sin perfil de riesgo aprobado y sin CTC completo, incluido el canal web, con aviso de perfiles por vencer.', 'DD', 'Crítica', 'Tecnología / Oficial de Cumplimiento / Producción', '30 días (diseño) / 60 días (producción)', 'Pruebas y paso a producción'],
  ['0–30 días', 'PA-06', 'Alinear el Manual con la consulta universal en listas que ya se practica; aprobar el procedimiento ante coincidencias y congelamiento preventivo; capacitar a Atención al Cliente; documentar los barridos periódicos.', 'DD / PP', 'Crítica', 'Oficial de Cumplimiento / Tecnología', '30 días', 'Manual ajustado; procedimiento; asistencia; constancia de barridos'],
  ['0–30 días', 'PA-07', 'Aprobar el plan de subsanación de los 75 expedientes observados y completar el registro de excepciones con motivo, responsable, fecha límite y escalamiento.', 'DD / CD', 'Alta', 'Producción / Comercial / Oficial de Cumplimiento', '30 días (plan) / 90 días (ejecución)', 'Plan aprobado; avance quincenal'],
  ['0–30 días', 'PA-08', 'Actualizar la Matriz de Riesgo de Colaboradores y regularizar contratos, constancias de capacitación y del Reglamento Interno.', 'DD', 'Alta', 'Recursos Humanos / Oficial de Cumplimiento', '60 días', 'Matriz actualizada; expedientes completos'],
  ['0–30 días', 'PA-09', 'Entregar las matrices aprobadas (productos, clientes, proveedores, empleados y provincias) y consolidarlas en una Evaluación Institucional de Riesgo BC/FT/FPADM aprobada en Comité y Junta.', 'ER', 'Crítica', 'Oficial de Cumplimiento', '30 días (metodología) / 90 días (aprobación)', 'Metodología; matriz aprobada'],
  ['0–30 días', 'PA-10', 'Completar el expediente de gobierno: actas del Comité de los trimestres II a IV de 2024; actas de ratificación de la Junta del Manual, la matriz de riesgo, la herramienta de listas, los Programas 2024 y 2025 y el Informe de Gestión 2024; memorial del Programa 2024 y sello de recibido del Programa 2025; e informes a la Junta.', 'PAC / CD', 'Alta', 'Secretaría del Comité / Oficial de Cumplimiento', '15 días', 'Expediente de gobierno completo'],
  ['0–30 días', 'PA-11', 'Conciliar los reportes de efectivo y cuasi efectivo enviados a la UAF durante el período y conservar los acuses.', 'RR', 'Alta', 'Oficial de Cumplimiento', '30 días', 'Conciliación y acuses'],
  ['30–90 días', 'PA-12', 'Aprobar el reglamento del Comité de Cumplimiento (quórum de directores, calendario trimestral, temas mínimos) e iniciar el informe trimestral a la Junta, con el consolidado de alertas y ROS.', 'GC', 'Alta', 'Junta Directiva / Secretaría', '60 días', 'Reglamento; calendario; primer informe'],
  ['30–90 días', 'PA-13', 'Aprobar el cronograma de DEIVID con hitos, presupuesto y responsable; completar y firmar el documento de requerimientos; documentar los controles compensatorios.', 'MT', 'Alta', 'Tecnología / Gerencia General / Oficial de Cumplimiento', '30 días (plan) / 90 días (hitos)', 'Cronograma aprobado; documento firmado'],
  ['30–90 días', 'PA-14', 'Formalizar el procedimiento de gestión y cierre de alertas (incluido el registro de falsos positivos) y del monitoreo de noticias, con formulario de datos mínimos.', 'MT', 'Alta', 'Oficial de Cumplimiento', '45 días', 'Procedimiento; formulario; reporte mensual'],
  ['30–90 días', 'PA-15', 'Perfilar al asegurado y al beneficiario; consultar listas antes de pagar reclamos; exigir revisión de Cumplimiento en pagos significativos de alto riesgo o PEP; extender la Política de Excepción a pagos a terceros.', 'DD / MT', 'Alta', 'Reclamos / Oficial de Cumplimiento / Tecnología', '90 días', 'Parametrización; política actualizada'],
  ['30–90 días', 'PA-16', 'Integrar el perfilamiento de Fianzas al Módulo de Perfil de Riesgo institucional.', 'DD', 'Alta', 'Fianzas / Tecnología', '60 días', 'Evidencia en el módulo'],
  ['30–90 días', 'PA-17', 'Aprobar la Evaluación de Reclamos (recalibrando R3 y R4) y confirmar la aprobación de la Evaluación de Emisión.', 'ER', 'Alta', 'Oficial de Cumplimiento', '90 días', 'Evaluaciones aprobadas'],
  ['30–90 días', 'PA-18', 'Completar el Código de Ética, aprobar la política de transacciones con partes relacionadas y poner en marcha el canal de denuncia con anonimato.', 'GC', 'Alta', 'Junta Directiva / Recursos Humanos / Oficial de Cumplimiento', '90 días', 'Código aprobado; canal operativo'],
  ['30–90 días', 'PA-19', 'Reorganizar el Programa Anual con indicadores y avance por actividad, mantener informes estrictamente trimestrales y calendarizar la aprobación del Informe de Gestión por el Comité.', 'PAC', 'Media', 'Oficial de Cumplimiento / Comité', '60 días', 'Programa ajustado; tablero de indicadores'],
  ['30–90 días', 'PA-20', 'Impartir capacitación por función (Fianzas, Tecnología, Atención al Cliente, Gerencia Comercial), incluidos los criterios y formatos de la UAF, con constancia individual.', 'PAC / RR', 'Alta', 'Oficial de Cumplimiento / Recursos Humanos', '90 días', 'Constancias; evaluaciones'],
  ['30–90 días', 'PA-21', 'Incluir en el Plan de Auditoría Interna 2027, basado en riesgos, la revisión de Cumplimiento y de las herramientas AML/CFT, tomando como insumo el informe de Auditoría Interna al Departamento de Cumplimiento.', 'GC', 'Alta', 'Auditoría Interna', '90 días', 'Plan aprobado'],
  ['30–90 días', 'PA-22', 'Corregir los cuadros de control de versiones y crear el registro maestro de documentos.', 'PP', 'Media', 'Oficial de Cumplimiento', '60 días', 'Registro maestro'],
  ['90–180 días', 'PA-23', 'Documentar la calibración de reglas y la gestión de falsos positivos, y automatizar la detección de PEP en toda la cartera.', 'MT', 'Media', 'Oficial de Cumplimiento / Tecnología', '180 días', 'Metodología; informe de calibración'],
  ['90–180 días', 'PA-24', 'Homologar y actualizar las políticas del Grupo B, colaboradores (por riesgo), proveedores, medios de pago, GESADHE-12 y capacitación, y aprobar la política de dependencia de terceros.', 'PP', 'Media', 'Oficial de Cumplimiento / Recursos Humanos / Compras', '120 días', 'Políticas aprobadas'],
  ['90–180 días', 'PA-25', 'Completar el estudio de capacidad de Cumplimiento y llevar la revisión anual de recursos al Comité.', 'GC', 'Alta', 'Gerencia General / Junta Directiva', '120 días', 'Estudio aprobado'],
  ['90–180 días', 'PA-26', 'Rediseñar la autoevaluación del Comité con soporte por respuesta y presentarla a la Junta con un plan de acción.', 'GC', 'Media', 'Comité de Cumplimiento', '120 días', 'Autoevaluación firmada'],
  ['90–180 días', 'PA-27', 'Designar suplente del Oficial de Cumplimiento y Oficial de Protección de Datos, documentar la continuidad de la función y actualizar el Manual de Control Interno al marco panameño.', 'GC', 'Media', 'Junta Directiva / Gerencia General', '180 días', 'Designaciones; documentos aprobados'],
  ['90–180 días', 'PA-28', 'Actualizar la política de conservación (inventario, respaldo del registro de efectivo, pruebas de recuperación y reportes DLP).', 'CD', 'Media', 'Oficial de Cumplimiento / Tecnología', '120 días', 'Política; prueba de recuperación'],
  ['90–180 días', 'PA-29', 'Revisar cada trimestre una muestra de expedientes (segunda línea), con indicadores por área.', 'DD', 'Media', 'Oficial de Cumplimiento', 'Desde el día 90, permanente', 'Informe trimestral de calidad'],
  ['90–180 días', 'PA-30', 'Revisar los expedientes de intermediarios del Grupo B, reaseguradoras y proveedores.', 'DD', 'Media', 'Oficial de Cumplimiento / Comercial', '150 días', 'Informe de revisión'],
  ['90–180 días', 'PA-31', 'Documentar la evaluación de riesgo del canal web y exigir evaluación previa para nuevos productos, canales y tecnologías.', 'ER', 'Media', 'Oficial de Cumplimiento', '180 días', 'Evaluación documentada'],
];


const RM = [
  ['RM-01', 'Adecuación del Manual y las políticas (Arts. 10, 11, 46 y 47)', 'El Acuerdo rige desde el 10-12-2025 y dio 60 días para adecuar manuales y controles; deroga parcialmente el Acuerdo 03-2022 para el Grupo A.', 'Los documentos revisados del período citan el Acuerdo 03-2022, como correspondía. Las entrevistas mencionan un Manual v4.0.', 'Confirmar que la versión vigente del Manual y de cada política ya incorpora el Acuerdo 07-2025, con aprobación del Comité y ratificación de la Junta. Si quedara algún documento pendiente, cerrarlo en una sesión extraordinaria apoyada en una matriz de brechas.', 'Fase 1', 'GC / PP'],
  ['RM-02', 'Nombramiento del Oficial de Cumplimiento (Arts. 4 y 6)', 'Notificación a la SSRP 15 días calendario antes del nombramiento para su aprobación; el candidato debe dejar cualquier función incompatible.', 'Oficial con doble función Cumplimiento/Legal (GC-01).', 'Diseñar la separación de funciones bajo este estándar y preparar el expediente del perfil para la SSRP si cambia el titular.', 'Fase 1', 'GC'],
  ['RM-03', 'Informes trimestrales a Comité y Junta (Art. 7, lit. h)', 'La periodicidad pasa a ser trimestral fija.', 'Informe combinado del III y IV trimestre de 2024 (PAC-02); sin informes del Comité a la Junta (GC-04).', 'Calendarizar cuatro informes al año con un contenido estándar: avance del programa, alertas y ROS, riesgos, capacitación y observaciones.', 'Fase 1', 'GC'],
  ['RM-04', 'Programa de Cumplimiento (Art. 9)', 'Contenido mínimo definido (literales a a i), alcance de grupo y remisión a la SSRP en los primeros 30 días del año.', 'Programa 2025 remitido a la SSRP; sin indicadores de avance (PAC-04).', 'Estructurar el Programa 2027 con los nueve contenidos mínimos y un tablero de indicadores, y aprobarlo antes del 30 de enero de 2027.', 'Fase 2', 'PAC'],
  ['RM-05', 'Políticas mínimas (Art. 11)', 'Cinco políticas mínimas, incluida una de debida diligencia sobre relaciones de reaseguro.', 'Ancón ya cuenta con un conjunto más amplio de políticas (fortaleza).', 'Mantener las políticas actuales como buena práctica y verificar que la de reaseguro cumpla el lit. b; aprobarlas en Comité y ratificarlas en Junta.', 'Fase 2', 'PP'],
  ['RM-06', 'Evaluación de riesgo y apetito (Arts. 8 y 12)', 'Las medidas de control deben responder al apetito de riesgo; la evaluación se hace con la periodicidad que fije el Manual y debe estar disponible para la SSRP.', 'Matrices por factor referidas como aprobadas, no entregadas; sin evaluación institucional ni apetito de riesgo (ER-01).', 'Construir la matriz institucional de una vez bajo este estándar y acompañarla de una declaración de apetito de riesgo aprobada por la Junta, con límites e indicadores.', 'Fase 2', 'ER'],
  ['RM-07', 'Variables de riesgo del cliente (Art. 23)', 'Incorpora la residencia fiscal y exige escala alto, medio y bajo descrita en la metodología.', 'Metodología de clasificación del cliente no entregada para revisión.', 'Incluir las once variables del Art. 23 en el Módulo de Perfil de Riesgo y documentarlas en la metodología.', 'Fase 2', 'ER'],
  ['RM-08', 'Debida diligencia simplificada (Art. 15)', 'No se permite en vida individual con ahorro, fianzas ni productos con devolución de prima.', 'Fianzas perfila fuera del módulo institucional (DD-04).', 'Parametrizar en el sistema que estos productos no puedan recibir debida diligencia simplificada y llevar Fianzas al módulo institucional.', 'Fase 1', 'DD'],
  ['RM-09', 'Personas jurídicas y beneficiario final (Arts. 16 y 18)', 'Lista de datos mínimos (incluidos estados financieros o declaración de renta), beneficiario final desde 10 % y abstención si persiste la duda.', 'Documentación societaria incompleta o vencida (DD-01).', 'Actualizar el CTC de persona jurídica con los campos del Art. 16 y un control de vigencia de certificados.', 'Fase 2', 'DD'],
  ['RM-10', 'Debida diligencia ampliada (Art. 17)', 'Supuestos obligatorios: efectivo de B/.10,000 o más sin justificación, vida con ahorro, PEP, jurisdicciones de riesgo y bienes de doble uso.', 'Sin activación automática por riesgo (MT-07).', 'Configurar disparadores automáticos de debida diligencia ampliada para cada supuesto del Art. 17.', 'Fase 2', 'DD'],
  ['RM-11', 'Personas expuestas políticamente (Arts. 19 y 20)', 'La aprobación de la relación corresponde al Gerente General; se prohíbe el trato discriminatorio.', 'Detección de PEP no autodeclarados manual (MT-05).', 'Ajustar el flujo de aprobación de PEP para que lo firme el Gerente General y automatizar la detección.', 'Fase 1', 'DD'],
  ['RM-12', 'Perfil financiero del contratante y pagador (Art. 21)', 'Debe considerar patrimonio e ingresos fijos y variables del contratante o de quien paga la prima.', 'Perfil financiero incompleto en expedientes (DD-01).', 'Hacer obligatorios estos campos en el CTC y validarlos antes de emitir.', 'Fase 2', 'DD'],
  ['RM-13', 'Jurisdicciones de riesgo (Art. 22)', 'Se consideran sanciones de la UE, Reino Unido, ONU y EE. UU., y las jurisdicciones señaladas por el GAFI.', 'Barridos evidenciados; procedimiento ante coincidencias por formalizar (DD-05).', 'Mantener una lista de países de riesgo actualizada trimestralmente e integrada al perfil del cliente.', 'Fase 2', 'DD'],
  ['RM-14', 'Nuevos productos y tecnologías (Art. 7, lit. c)', 'Evaluación de riesgo previa al lanzamiento.', 'Canal web sin evaluación previa documentada (ER-04).', 'Incorporar un formulario de evaluación de riesgo obligatorio en el proceso de aprobación de productos y canales.', 'Fase 3', 'ER'],
  ['RM-15', 'Reportes de efectivo (Art. 24)', 'Plazo de 10 días hábiles (antes 5); régimen semestral para operaciones ocasionales y declaración jurada si no hay operaciones.', 'Restricción del efectivo desde enero de 2026 (PP-09).', 'Ajustar el procedimiento al nuevo plazo y evaluar si la restricción del efectivo justifica pasar al régimen semestral o a la declaración jurada, comunicándolo a la UAF.', 'Fase 1', 'RR'],
  ['RM-16', 'Operaciones inusuales y sospechosas (Arts. 25 y 26)', 'Constancia de cada operación inusual con su responsable y decisión; notificación inmediata del ROS.', 'Sin ROS en el período y sin registro de inusuales (RR-01); caso Fianzas 2024 (RR-02).', 'Diseñar el registro único y el procedimiento de ROS directamente con este estándar.', 'Fase 1', 'RR / MT'],
  ['RM-17', 'Reportes a la SSRP (Art. 7, lit. l)', 'Todo reporte o cuestionario a la SSRP requiere aprobación previa del Comité.', 'Sin control centralizado (RR-06).', 'Incluir esta aprobación en el calendario regulatorio y en el temario del Comité.', 'Fase 2', 'RR'],
  ['RM-18', 'Auditoría interna (Art. 8)', 'Procedimientos continuos de Auditoría Interna sobre la efectividad del sistema de prevención.', 'Plan 2025 sin revisión de Cumplimiento (GC-07).', 'Incluir en el plan anual una revisión de Cumplimiento y de las herramientas AML/CFT.', 'Fase 2', 'GC'],
  ['RM-19', 'Régimen sancionatorio (Arts. 37 a 43)', 'Tipifica expresamente, entre otras, la falta de exclusividad del Oficial (gravedad media) y la emisión de una póliza a quien no facilite la debida diligencia ampliada (gravedad máxima).', 'Varias de estas conductas aparecen en los hallazgos del período.', 'Presentar a la Junta un mapa de tipificaciones del nuevo régimen y usarlo para priorizar el plan de acción.', 'Fase 2', 'Transversal'],
  ['RM-20', 'Capacitación y divulgación (Arts. 7, lits. b e i; 41, lit. d)', 'Programa anual para colaboradores, Oficial de Cumplimiento y Gerente General; divulgación de la normativa a todo el personal.', 'Capacitaciones ejecutadas; constancias individuales incompletas (PAC-05).', 'Incluir un módulo sobre el Acuerdo 07-2025 para Junta, Gerencia y áreas sensibles, con constancia individual.', 'Fase 2', 'PAC'],
  ['RM-21', 'Evaluación independiente (Ley 23, Art. 45)', 'El Acuerdo 07-2025 no la regula expresamente para el Grupo A; sigue rigiendo la Ley 23.', 'Esta evaluación cubre el período 2024-2025.', 'Mantener la evaluación independiente al menos cada dos años como buena práctica y programar la próxima con alcance sobre el Acuerdo 07-2025.', 'Fase 3', 'Transversal'],
  ['RM-22', 'Señales de alerta (Art. 47)', 'Se deroga el Acuerdo 02-2015, que contenía el catálogo oficial de señales de alerta.', 'Ancón tiene catálogo propio 2025 y en la evaluación se utilizó el catálogo de tipologías de la UAF.', 'Adoptar formalmente, por acuerdo del Comité, el catálogo de tipologías de la UAF como referencia y mantener actualizado cada año el catálogo propio por línea de negocio.', 'Fase 3', 'MT / PP'],
  ['RM-23', 'Conservación y respaldo (Arts. 24 y 42, lit. a)', 'Conservación de formularios y soportes por al menos cinco años y copia de seguridad del registro de efectivo.', 'Sin respaldo evidenciado (CD-03).', 'Incluir estos requisitos en la política de conservación y probar la recuperación una vez al año.', 'Fase 3', 'CD'],
  ['RM-24', 'Planes de actualización de expedientes (Art. 7, lits. f y g)', 'Planes de actualización de clientes con operaciones inconsistentes y de colaboradores, con informe a Comité y Junta.', 'Expedientes de colaboradores desactualizados (DD-08).', 'Formalizar ambos planes con metas trimestrales y reportarlos en el informe trimestral.', 'Fase 2', 'PAC / DD'],
];
const FASES = [
  ['Fase 1 – Asegurar lo esencial', '0–30 días', 'Confirmar la adecuación normativa y resolver los temas que el nuevo régimen tipifica de forma expresa.', 'RM-01, RM-02, RM-03, RM-08, RM-11, RM-15, RM-16'],
  ['Fase 2 – Alinear el modelo', '30–120 días', 'Llevar la metodología de riesgo, la debida diligencia y el programa anual al estándar del Acuerdo 07-2025.', 'RM-04, RM-05, RM-06, RM-07, RM-09, RM-10, RM-12, RM-13, RM-17, RM-18, RM-19, RM-20, RM-24'],
  ['Fase 3 – Optimizar y sostener', '90–180 días', 'Consolidar buenas prácticas que dan permanencia al sistema.', 'RM-14, RM-21, RM-22, RM-23'],
];

const OMX = [
  ['OM-01', 'Verificación integral de expedientes de intermediarios (Grupo B) y reaseguradoras', 'Confirmar que los expedientes de corredores, agencias y reaseguradoras estén al día conforme a los plazos bienal y anual de CONSUREG-10.', 'POL-CUM-27-CONSUREG-10', 'Revisión puntual antes del cierre de año (PA-30).'],
  ['OM-02', 'Procedimiento formal de control de cambios normativos', 'Asegurar que las referencias regulatorias de todos los documentos se actualicen cuando entre en vigor una nueva norma de la SSRP u otro regulador.', 'Marco normativo institucional', 'Diseñar alertas normativas con responsables y plazos (PA-22).'],
  ['OM-03', 'Validación de la metodología de clasificación de riesgo del cliente', 'Confirmar que la metodología incluya todas las variables del Acuerdo 03-2022 (transferencias, origen y fuente de recursos, canales) y la periodicidad de actualización por riesgo.', 'MPREMET-21; CONCLI-15', 'Entregar la metodología para su revisión (PA-09).'],
];

// ---------- Exportar datos para la matriz Excel ----------
fs.writeFileSync('matrix_data.json', JSON.stringify({
  hallazgos: ALLH.map(({ comp, h }) => ({ id: h[0], comp, short: h[1], text: h[2], base: h[3], p: PNAME[h[4]], i: INAME[h[5]], docs: h[6], nota: h[7], pa: h[8] })),
  rm: RM.map(r => ({ id: r[0], tema: r[1], cambia: r[2], sit: r[3], mejora: r[4], fase: r[5], comp: r[6] })),
  omx: OMX.map(o => ({ id: o[0], tema: o[1], desc: o[2], docs: o[3], accion: o[4] })),
}, null, 1));

// =====================================================================
// CONSTRUCCIÓN DEL DOCUMENTO
// =====================================================================
const TOC = [];
const T1 = (t) => { TOC.push([t, 1]); return H1(t); };
const T1nb = (t) => { TOC.push([t, 1]); return H1nb(t); };
const T2 = (t, pb) => { TOC.push([t, 2]); return H2(t, pb); };
const S1 = [];

S1.push(new Paragraph({ children: [], spacing: { before: 1800 } }));
S1.push(new Paragraph({ children: [new TextRun({ text: 'INFORME FINAL CONSOLIDADO', font: FONT, size: 26, bold: true, color: ORANGE })], spacing: { after: 120 } }));
S1.push(new Paragraph({ children: [new TextRun({ text: 'Evaluación Independiente del Sistema de Prevención de BC/FT/FPADM', font: FONT, size: 44, bold: true, color: TEAL })], spacing: { after: 200 } }));
S1.push(new Paragraph({ children: [new TextRun({ text: 'Aseguradora Ancón, S.A.', font: FONT, size: 32, color: TEAL })], spacing: { after: 160 } }));
S1.push(new Paragraph({ children: [new TextRun({ text: 'Incluye matriz de riesgo GSM y hoja de ruta de mejoras hacia el Acuerdo SSRP 07-2025', font: FONT, size: 22, italics: true, color: GREY })], spacing: { after: 600 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ORANGE, space: 8 } } }));
S1.push(table(null, [
  ['Entidad evaluada', 'Aseguradora Ancón, S.A., sujeto obligado del Grupo A supervisado por la Superintendencia de Seguros y Reaseguros de Panamá (SSRP)'],
  ['Período evaluado', '1 de julio de 2024 al 30 de junio de 2025'],
  ['Criterio de evaluación', 'Ley 23 de 2015 y sus modificaciones; Decreto Ejecutivo 35 de 2022; Acuerdo SSRP 03-2022'],
  ['Valor agregado', 'Catálogo de mejoras y hoja de ruta frente al Acuerdo SSRP 07-2025 (fuera del alcance de la evaluación)'],
  ['Corte de la evidencia', 'Cuadro de control de entregables (octubre de 2026) y aclaración de aprobaciones en actas del Comité de Cumplimiento'],
  ['Fecha del informe', 'Octubre de 2026'],
  ['Elaborado por', 'GSM Consulting – José Antonio Serrano, Melva Gutiérrez, Isaac Muñoz'],
  ['Clasificación', 'CONFIDENCIAL – Uso exclusivo de Aseguradora Ancón, S.A.'],
], [2800, 6560], { firstColFill: true, firstColBold: true, size: 19 }));
const TOC_POS = S1.length;

// ---- 1 ----
S1.push(T1('1. Resumen Ejecutivo'));
S1.push(P('Este informe recoge los resultados de la Evaluación Independiente del sistema de prevención de Blanqueo de Capitales, Financiamiento del Terrorismo y Financiamiento de la Proliferación de Armas de Destrucción Masiva (PBC/FT/FPADM) de Aseguradora Ancón, S.A., para el período del 1 de julio de 2024 al 30 de junio de 2025. Lo organizamos para que Ancón obtenga dos productos en uno:'));
S1.push(NUM('**La evaluación del período (Parte I)**, medida contra la norma vigente entonces: la Ley 23 de 2015, el Decreto Ejecutivo 35 de 2022 y el Acuerdo SSRP 03-2022, con su matriz de riesgo.', 'two'));
S1.push(NUM('**Una hoja de ruta hacia el Acuerdo SSRP 07-2025 (Parte II)**, que no formaba parte del alcance porque se publicó después del período, pero que hoy es la norma aplicable. La incluimos como valor agregado para que cada acción de remediación se diseñe una sola vez y con el nuevo estándar.', 'two'));
S1.push(H3('Resultado de la evaluación del período'));
S1.push(P(`Calificamos el sistema en **Nivel 2 – ALTO (Brecha Mayor)**. Dos componentes quedan en **Nivel 1 – CRÍTICO**: ${N1.join(' y ')}. En la matriz de riesgo documentamos ${TOTAL} brechas —${TOT['CRÍTICO']} críticas, ${TOT['MEDIO-ALTO']} de nivel medio-alto, ${TOT['MEDIO']} medias y ${TOT['BAJO']} bajas— y ${RM.length + OMX.length} oportunidades de mejora. De los 96 requerimientos de información, Ancón aportó 72; 17 están pendientes de entrega y en 7 casos la entidad indicó que el documento no existe.`));
S1.push(P('Ancón tiene una base documental amplia y avances que vale la pena reconocer: un ciclo de aprobaciones del Comité ordenado y trazable, informes de seguimiento al Comité, programas anuales y un Informe de Gestión Anual, barridos de listas sobre la cartera, el catálogo de Señales de Alerta 2025, evaluaciones de riesgo por proceso y una Oficial de Cumplimiento que reporta directamente a la Junta. El reto principal no está en el diseño, sino en la **ejecución y en poder demostrarla**: varias condiciones observadas por la SSRP en 2022 se mantienen, y las pruebas muestran fallas en la debida diligencia (88.24 % de los expedientes de clientes con observaciones y 96.7 % de los colaboradores con su nivel de riesgo desactualizado).'));
S1.push(H3('Calificación por componente'));
S1.push(table(['Componente', 'Calificación', 'C', 'MA', 'M', 'B', 'Aspectos determinantes'],
  C.map(c => [`${c.num} ${c.title}`, { t: NIVEL[c.nivel].t, fill: NIVEL[c.nivel].c, bold: true },
    ...LEV.map(k => ({ t: String(c.cnt[k]), align: AlignmentType.CENTER })), DETERM[c.num]]),
  [1900, 1700, 380, 460, 380, 380, 4160], { size: 16 }));
S1.push(P('C = Crítico, MA = Medio-Alto, M = Medio, B = Bajo (número de brechas por nivel según la matriz de riesgo GSM).', { size: 16, italics: true, color: GREY, before: 60 }));
S1.push(H3('Lo que recomendamos atender primero'));
[
  '**Completar el cierre de los hallazgos de la SSRP de 2022**, actualizando con evidencia la matriz de seguimiento que Ancón ya preparó. La subsanación por iniciativa propia reduce la sanción hasta en 40 %; su persistencia puede considerarse desobediencia.',
  '**Separar las funciones de Cumplimiento y Legal.** Es una decisión de la Junta Directiva y una infracción expresamente tipificada.',
  '**Asegurar que no se emita ninguna póliza sin debida diligencia completa**, empezando por el canal web, mediante un bloqueo en el sistema.',
  '**Analizar y documentar el caso de Fianzas de 2024** y crear el registro de operaciones inusuales, dado que en el período no se presentaron ROS.',
  '**Entregar y consolidar las matrices de riesgo aprobadas** en una evaluación institucional, y **diseñar todas las soluciones con el estándar del Acuerdo 07-2025** (Sección 11).',
].forEach(t => S1.push(NUM(t, 'msg')));

// ---- 2 ----
S1.push(T1('2. Objetivos'));
S1.push(H3('Objetivo general'));
S1.push(P('Evaluar de forma independiente el diseño y la efectividad del sistema de prevención de BC/FT/FPADM de Aseguradora Ancón durante el período evaluado, determinar su grado de cumplimiento con la normativa aplicable e identificar las brechas, los riesgos y las acciones necesarias para proteger la licencia y la reputación de la entidad. Como objetivo complementario, ofrecer a Ancón una hoja de ruta para alinear su sistema con el Acuerdo SSRP 07-2025.'));
S1.push(H3('Objetivos específicos'));
[
  'Evaluar el **gobierno corporativo** del programa: supervisión de la Junta y del Comité, e independencia y recursos del Oficial de Cumplimiento.',
  'Verificar que las **políticas y procedimientos** estén alineados con la normativa, vigentes, trazables y al alcance del personal.',
  'Evaluar cómo se planifica, aprueba, ejecuta y da seguimiento al **Programa Anual de Cumplimiento**.',
  'Evaluar la adecuación y efectividad del **monitoreo transaccional** y de la gestión de alertas.',
  'Verificar la efectividad de la **debida diligencia** de clientes, colaboradores, intermediarios y proveedores.',
  'Evaluar la existencia, la metodología y el uso de la **evaluación de riesgo BC/FT/FPADM**.',
  'Evaluar los controles para detectar y **reportar** operaciones inusuales y sospechosas, y los reportes a la UAF y la SSRP.',
  'Verificar la **conservación de documentos** y la capacidad de reconstruir operaciones.',
  'Valorar los riesgos identificados en una **matriz de probabilidad e impacto** e identificar las **mejoras necesarias frente al Acuerdo 07-2025**.',
].forEach(t => S1.push(NUM(t, 'obj')));

// ---- 3 ----
S1.push(T1('3. Alcance'));
S1.push(T2('3.1 Entidad, período y criterio de evaluación'));
S1.push(P('Evaluamos a Aseguradora Ancón, S.A., sujeto obligado del Grupo A ante la SSRP, durante el período del 1 de julio de 2024 al 30 de junio de 2025. Los criterios son la Ley 23 de 2015 y sus modificaciones, el Decreto Ejecutivo 35 de 2022 y el **Acuerdo SSRP 03-2022**. Siguiendo la aclaración de aprobaciones presentada por Ancón, tomamos en cuenta las decisiones del Comité de octubre de 2023 y abril de 2024 como base normativa vigente durante el período. Esa aclaración documenta las aprobaciones del Comité; las ratificaciones de la Junta Directiva y las remisiones a la SSRP se evalúan con la documentación propia de esos órganos. Cuando un hecho posterior es relevante, lo indicamos.'));
S1.push(P('Revisamos ocho componentes: Gobierno Corporativo; Políticas y Procedimientos; Programa Anual de Cumplimiento; Monitoreo Transaccional; Debida Diligencia; Evaluación de Riesgo AML; Reportes Regulatorios; y Conservación de Documentos.'));
S1.push(table(null, [[{ t: '**El Acuerdo SSRP 07-2025 está fuera del alcance de esta evaluación.** Se aprobó el 27 de noviembre de 2025 y se publicó en Gaceta Oficial el 10 de diciembre de 2025, después del cierre del período. Por eso no lo utilizamos como criterio para calificar los componentes ni emitimos opinión sobre el cumplimiento de Ancón respecto de él. Dado que hoy es la norma aplicable, en la Sección 11 presentamos, como valor agregado, un catálogo de mejoras y una hoja de ruta para la transición.', fill: 'FDF1E4', size: 19 }]], [9360]));
S1.push(T2('3.2 Procedimientos realizados'));
S1.push(table(['Procedimiento', 'Cobertura'], [
  ['Solicitud y revisión documental', '96 requerimientos de información sobre gobierno AML y administración de riesgo (74), gobierno corporativo (13) y capacitación y conocimiento del colaborador (9). Ver Anexo C.'],
  ['Revisión del ciclo de aprobaciones', 'Aclaración de aprobaciones en actas del Comité (IV trimestre 2023 – 2025), actas del 17-04-2024 y 28-05-2025 y cinco informes de seguimiento al Comité.'],
  ['Entrevistas', '13 entrevistas a responsables y personal de Comercial, Fianzas, Emisión, Cobros, Tecnología, Recursos Humanos, Sucursales, Reclamos de Auto, Atención al Cliente, Personas, Reclamos Patrimoniales y la sucursal de La Chorrera.'],
  ['Prueba de expedientes de clientes', '85 expedientes (57 personas naturales y 28 jurídicas).'],
  ['Prueba de expedientes de colaboradores', '33 de 181 colaboradores activos (18.2 %).'],
  ['Verificación de controles de listas', 'Herramienta de consulta, constancia de barridos sobre clientes existentes y listado de pólizas rechazadas o canceladas por coincidencias.'],
  ['Seguimiento de supervisión', 'Hallazgos de la inspección SSRP DSR-0924-2023 y matriz de seguimiento de cierre entregada por Ancón (16-01-2026).'],
  ['Valoración de riesgos', 'Matriz de riesgo GSM de probabilidad por impacto (Sección 7 y archivo Excel adjunto).'],
], [3000, 6360], { size: 17 }));
S1.push(T2('3.3 Limitaciones'));
S1.push(table(['Información', 'Situación y tratamiento en el informe'], [
  ['Actas del Comité de los trimestres II a IV de 2024.', 'Según la aclaración de Ancón, fueron sesiones de seguimiento sin nuevas aprobaciones; al no aportarse las copias, su celebración no puede verificarse documentalmente (GC-03).'],
  ['Actas de Junta Directiva de ratificación del Manual, la matriz de riesgo, la herramienta de listas, los Programas 2024 y 2025 y el Informe de Gestión 2024; acta de designación del Oficial.', 'Algunas se aportaron y su alcance está sujeto a confirmación con la Oficial de Cumplimiento; otras están pendientes de entrega (GC-10, GC-11, PAC-01, PAC-03).'],
  ['Acta del Comité que aprueba el Informe de Gestión 2024; memorial de remisión del Programa 2024; sello de recibido del Programa 2025.', 'Se reflejan en PAC-01 y PAC-03.'],
  ['Informes a la Junta Directiva y descripción de puesto del Oficial.', 'Se reflejan en GC-04 y GC-10.'],
  ['Copias de las matrices de riesgo de clientes, colaboradores y proveedores, la metodología de clasificación de clientes y el acta de Junta que aprueba la matriz.', 'Aprobadas por el Comité según la aclaración; ER-01 se califica con la evidencia disponible y puede recalificarse al aportarse.'],
  ['Catálogos de clientes activos, PEP, canales y proveedores; programa de visitas a sucursales.', 'Las pruebas se realizaron sobre las muestras aportadas.'],
  ['Reportes de efectivo y cuasi efectivo del período; listado completo de alertas del período.', 'Se reflejan en RR-05 y MT-03.'],
  ['Ancón indicó no contar con: política de dependencia de terceros, política de partes relacionadas, calendario y temario del Comité, informe anual del Comité, memorandos mensuales a Gerencia y listado de falsos positivos.', 'Se tratan como brechas (PP-10, GC-05, GC-03, GC-04, MT-04).'],
  ['Las entrevistas mencionan una versión 4.0 del Manual.', 'Para el período se toma como base la versión 3.0 (octubre de 2023).'],
], [5160, 4200], { size: 16 }));
S1.push(T2('3.4 Independencia y responsabilidades'));
S1.push(P('La evaluación fue realizada por especialistas de GSM Consulting con criterio independiente de las áreas evaluadas, en los términos del Art. 45 de la Ley 23 de 2015 y del Art. 18 del Acuerdo 03-2022. Nuestras conclusiones se apoyan en la documentación entregada, en las entrevistas y en las pruebas de expedientes. El diseño, la implementación y la operación del sistema de prevención son responsabilidad de la Junta Directiva y de la Alta Gerencia de Ancón.'));

// ---- 4 ----
S1.push(T1('4. Marco Normativo de Referencia'));
S1.push(H3('Normas aplicadas en la evaluación del período'));
S1.push(table(['Instrumento', 'Uso en la evaluación'], [
  ['Ley 23 de 27 de abril de 2015 y sus modificaciones (incluida la Ley 254 de 2021)', 'Marco legal: debida diligencia basada en riesgo (Art. 26), actualización y conservación (Art. 29), PEP (Art. 34), conocimiento del empleado (Art. 42), evaluación independiente (Art. 45), capacitación (Art. 47), congelamiento preventivo (Art. 49) y reportes de efectivo (Art. 53) y de operación sospechosa (Art. 54).'],
  ['Decreto Ejecutivo 35 de 6 de septiembre de 2022', 'Reglamento de la Ley 23: enfoque basado en riesgo, perfiles, debida diligencia, operaciones inusuales y sospechosas, conservación y criterios de gravedad (Art. 29).'],
  ['Acuerdo SSRP 03-2022 (24-11-2022; Gaceta de 03-01-2023)', 'Norma sectorial vigente durante el período y principal criterio, incluido su régimen sancionatorio (Arts. 55 a 63).'],
  ['Ley 81 de 2019', 'Protección de datos personales.'],
  ['Recomendaciones del GAFI', 'R.1 (enfoque de riesgo), R.10 (debida diligencia), R.12 (PEP), R.18 (controles internos) y R.20 (operaciones sospechosas).'],
], [3300, 6060], { size: 17, firstColBold: true }));
S1.push(H3('Norma de referencia para el valor agregado (fuera del alcance)'));
S1.push(table(['Instrumento', 'Uso en este informe'], [
  ['Acuerdo SSRP 07-2025 (27-11-2025; Gaceta de 10-12-2025)', 'Base del catálogo de mejoras y de la hoja de ruta de la Sección 11.'],
  ['Acuerdo SSRP 3-2025 (19-06-2025) – Gobierno Corporativo', 'Referido por Ancón como base de la actualización posterior de su marco de gobierno corporativo.'],
], [3300, 6060], { size: 17, firstColBold: true }));

// ---- 5 ----
S1.push(T1('5. Metodología y Escala de Calificación'));
S1.push(P('Trabajamos en cuatro etapas: identificar las brechas y su causa raíz; valorar cada brecha en la matriz de riesgo GSM según su probabilidad e impacto; definir estrategias de control y mitigación; y priorizarlas en un plan de acción con horizontes de 0–30, 30–90 y 90–180 días. Las mejoras que no constituyen incumplimiento del período se registran como **oportunidades de mejora**.'));
S1.push(H3('Escala de probabilidad'));
S1.push(table(['Puntaje', 'Nivel', 'Descripción'], [
  ['5', 'Frecuente', 'La probabilidad de ocurrencia es muy alta; sucede en la mayoría de las circunstancias.'],
  ['4', 'Probable', 'Es probable que ocurra; los controles o procesos no son suficientemente efectivos para prevenirlo.'],
  ['3', 'Ocasional', 'Existe la posibilidad de que ocurra; los controles de mitigación no son del todo efectivos.'],
  ['2', 'Posible', 'Posibilidad limitada; es poco probable con los controles actuales.'],
  ['1', 'Improbable', 'Probabilidad extremadamente baja; los controles son altamente efectivos.'],
], [900, 1600, 6860], { size: 16 }));
S1.push(H3('Escala de impacto'));
S1.push(table(['Puntaje', 'Nivel', 'Descripción'], [
  ['5', 'Catastrófico', 'Impacto severo e irreversible: sanciones regulatorias mayores, pérdidas masivas, daño reputacional crítico.'],
  ['4', 'Mayor', 'Pérdidas o sanciones significativas que requieren intervención inmediata de la Alta Gerencia y la Junta.'],
  ['3', 'Moderado', 'Impacto considerable; puede afectar operaciones o comprometer parcialmente el cumplimiento.'],
  ['2', 'Menor', 'Impacto leve; correcciones administrativas menores sin afectar operaciones críticas.'],
  ['1', 'Insignificante', 'Impacto mínimo; se resuelve con los controles normales.'],
], [900, 1600, 6860], { size: 16 }));
S1.push(H3('Niveles de riesgo (Probabilidad × Impacto)'));
S1.push(table(['Nivel', 'Puntaje', 'Acción requerida'], [
  [lvlCell('CRÍTICO'), '15 – 25', 'Riesgo inaceptable. Acción inmediata (0-30 días). Reporte obligatorio a la Junta Directiva y al Oficial de Cumplimiento.'],
  [lvlCell('MEDIO-ALTO'), '8 – 12', 'Riesgo significativo. Plan de acción formal (30-60 días). Monitoreo mensual.'],
  [lvlCell('MEDIO'), '4 – 6', 'Riesgo manejable. Medidas correctivas (60-90 días). Monitoreo trimestral.'],
  [lvlCell('BAJO'), '1 – 3', 'Riesgo aceptable con controles mínimos. Monitoreo semestral.'],
  [lvlCell('OPT. DE MEJORA'), 'N/A', 'No es hallazgo de incumplimiento. Mejora preventiva o de adecuación normativa prospectiva.'],
], [2000, 1200, 6160], { size: 16 }));
S1.push(H3('Calificación por componente'));
S1.push(table(['Nivel', 'Estado', 'Regla aplicada'], [
  [{ t: 'Nivel 1 – CRÍTICO', fill: NIVEL[1].c, bold: true }, 'No Cumple', 'Dos o más brechas de nivel CRÍTICO en el componente.'],
  [{ t: 'Nivel 2 – ALTO', fill: NIVEL[2].c, bold: true }, 'Brecha Mayor', 'Una brecha CRÍTICA o tres o más de nivel MEDIO-ALTO.'],
  [{ t: 'Nivel 3 – MEDIO', fill: NIVEL[3].c, bold: true }, 'Brecha Parcial', 'Una o dos brechas de nivel MEDIO-ALTO.'],
  [{ t: 'Nivel 4 – BAJO', fill: NIVEL[4].c, bold: true }, 'Cumple', 'Solo brechas de nivel MEDIO o BAJO.'],
], [2200, 1600, 5560], { size: 17 }));

// ---- 6 ----
S1.push(T1('6. Desarrollo por Componente'));
S1.push(P('Cada componente sigue la misma estructura: calificación, situación actual, brechas con su base normativa y valoración en la matriz (P:I = probabilidad e impacto), riesgo regulatorio y posible observación de la SSRP, recomendaciones y ruta de acción, incluida su conexión con la hoja de ruta del Acuerdo 07-2025.'));
C.forEach((c, idx) => {
  S1.push(T2(`${c.num} ${c.title}`, idx > 0));
  S1.push(table(null, [
    [{ t: 'Calificación', bold: true, fill: LIGHT }, { t: NIVEL[c.nivel].t, fill: NIVEL[c.nivel].c, bold: true }],
    [{ t: 'Brechas', bold: true, fill: LIGHT }, `${c.hallazgos.length} (Crítico: ${c.cnt['CRÍTICO']} · Medio-Alto: ${c.cnt['MEDIO-ALTO']} · Medio: ${c.cnt['MEDIO']} · Bajo: ${c.cnt['BAJO']})`],
    ...(c.nota ? [[{ t: 'Alcance', bold: true, fill: LIGHT }, c.nota]] : []),
  ], [2000, 7360], { size: 18 }));
  S1.push(H3('Situación actual'));
  c.situacion.forEach(t => S1.push(P(t)));
  if (c.tablaExtra) S1.push(table(c.tablaExtra.headers, c.tablaExtra.rows, c.tablaExtra.widths, { size: 17 }));
  S1.push(H3('Brechas identificadas'));
  S1.push(table(['ID', 'Brecha', 'Base normativa', 'P:I', 'Nivel'],
    c.hallazgos.map(h => [{ t: h[0], bold: true }, [`**${h[1]}.** ${h[2]}`], h[3], { t: `${h[4]}:${h[5]}`, align: AlignmentType.CENTER }, lvlCell(h.nivel)]),
    [760, 4840, 2400, 500, 860], { size: 16 }));
  S1.push(H3('Riesgo regulatorio y posible observación de la SSRP'));
  S1.push(P(c.riesgo));
  S1.push(H3('Recomendaciones'));
  S1.push(table(['ID', 'Recomendación', 'Prioridad'], c.recs.map(r => [{ t: r[0], bold: true }, r[1], lvlCell(r[2])]), [1000, 7360, 1000], { size: 16 }));
  S1.push(H3('Ruta de acción'));
  S1.push(table(['Horizonte', 'Acciones'],
    [...Object.entries(c.ruta).map(([k, v]) => [{ t: k, bold: true }, v]),
     [{ t: 'Hacia el Acuerdo 07-2025', bold: true, fill: 'FDF1E4' }, { t: `${c.rm} (ver Sección 11)`, fill: 'FDF1E4' }]],
    [2400, 6960], { size: 17 }));
});

// ---- 7 (landscape) ----
const S2 = [];
S2.push(T1nb('7. Matriz de Riesgo GSM'));
S2.push(P(`La matriz valora cada brecha del período por su probabilidad (P) y su impacto (I) con la metodología descrita en la Sección 5. El resultado es el mapa de calor siguiente, que concentra ${TOT['CRÍTICO']} brechas críticas y ${TOT['MEDIO-ALTO']} de nivel medio-alto. La matriz completa, con sus fórmulas, el mapa de calor y el registro de oportunidades de mejora, se entrega en el archivo Excel "ANCON_Matriz_Riesgo_GSM_2024-2025" adjunto a este informe.`));
S2.push(T2('7.1 Mapa de calor'));
const heatRows = [5, 4, 3, 2, 1].map(p => [{ t: `${p} – ${PNAME[p]}`, bold: true, fill: LIGHT }, ...[1, 2, 3, 4, 5].map(i => {
  const n = HEAT[`${p}:${i}`] || 0; return { t: n ? String(n) : '–', fill: LVL[nivelPI(p, i)], bold: !!n, align: AlignmentType.CENTER, size: n ? 22 : 16 };
})]);
S2.push(table(['Probabilidad / Impacto', ...[1, 2, 3, 4, 5].map(i => `${i} – ${INAME[i]}`)], heatRows, [2560, 2080, 2080, 2080, 2080, 2080], { size: 17 }));
S2.push(SP());
S2.push(table(['Resumen', 'Cantidad', 'Acción requerida'], [
  [lvlCell('CRÍTICO'), { t: String(TOT['CRÍTICO']), align: AlignmentType.CENTER, bold: true }, 'Acción inmediata (0-30 días); reporte a la Junta y al Oficial de Cumplimiento.'],
  [lvlCell('MEDIO-ALTO'), { t: String(TOT['MEDIO-ALTO']), align: AlignmentType.CENTER, bold: true }, 'Plan de acción formal (30-60 días); monitoreo mensual.'],
  [lvlCell('MEDIO'), { t: String(TOT['MEDIO']), align: AlignmentType.CENTER, bold: true }, 'Medidas correctivas (60-90 días); monitoreo trimestral.'],
  [lvlCell('BAJO'), { t: String(TOT['BAJO']), align: AlignmentType.CENTER, bold: true }, 'Controles mínimos; monitoreo semestral.'],
  [{ t: 'Total de brechas', bold: true, fill: LIGHT }, { t: String(TOTAL), align: AlignmentType.CENTER, bold: true, fill: LIGHT }, { t: '', fill: LIGHT }],
  [lvlCell('OPT. DE MEJORA'), { t: String(RM.length + OMX.length), align: AlignmentType.CENTER, bold: true }, `${RM.length} mejoras hacia el Acuerdo 07-2025 (Sección 11) y ${OMX.length} oportunidades del período.`],
], [2560, 1400, 9000], { size: 17 }));
S2.push(T2('7.2 Registro de brechas ordenado por nivel de riesgo', true));
const sorted = [...ALLH].sort((a, b) => (b.h[4] * b.h[5]) - (a.h[4] * a.h[5]) || a.h[0].localeCompare(b.h[0]));
S2.push(table(['ID', 'Brecha', 'Componente', 'Probabilidad', 'Impacto', 'P:I', 'Puntaje', 'Nivel', 'Plan'],
  sorted.map(({ comp, h }) => [{ t: h[0], bold: true }, h[1], comp, `${PNAME[h[4]]} (${h[4]})`, `${INAME[h[5]]} (${h[5]})`, { t: `${h[4]}:${h[5]}`, align: AlignmentType.CENTER }, { t: String(h[4] * h[5]), align: AlignmentType.CENTER, bold: true }, lvlCell(h.nivel), h[8]]),
  [760, 3580, 1700, 1300, 1300, 560, 820, 1240, 1700], { size: 15 }));
S2.push(P('Rangos de multa del Acuerdo 03-2022: infracciones leves de B/.5,000 a B/.15,000 (Art. 61); de gravedad media, más de B/.15,000 y hasta B/.1,000,000 (Art. 62); de gravedad máxima, más de B/.1,000,000 y hasta B/.5,000,000 (Art. 63). La Superintendencia puede imponer, además, la prohibición de operar ciertos ramos o la suspensión o cancelación de la licencia (Arts. 55 y 56).', { before: 120, size: 18 }));

// ---- 8, 9 ----
const S3 = [];
S3.push(T1nb('8. Conclusiones'));
[
  'Ancón ha hecho un esfuerzo sostenido por construir su programa de PBC/FT/FPADM. Tiene una arquitectura documental amplia, un ciclo de aprobaciones del Comité ordenado y explicado en su aclaración de aprobaciones, programas anuales y un Informe de Gestión Anual, una Oficial de Cumplimiento que reporta directamente a la Junta, barridos de listas sobre la cartera y una cultura de consulta valorada por las áreas.',
  'Aun así, el programa todavía no alcanza el nivel de efectividad que exige la normativa. La causa de fondo es la **distancia entre lo que está aprobado y lo que se puede demostrar que se ejecuta**, acentuada por un Departamento de Cumplimiento con una sola profesional que además dirige Legal y por la dependencia de controles manuales. Por eso varias condiciones observadas por la SSRP en 2022 —estructura de Cumplimiento, perfiles de riesgo, monitoreo, registro de operaciones sospechosas y matriz de riesgo— siguen total o parcialmente abiertas, pese a la matriz de seguimiento que Ancón preparó.',
  'Las pruebas confirman que el riesgo es real: el 88.24 % de los expedientes de clientes tiene observaciones, al beneficiario no se le perfila antes del pago de siniestros de vida, el canal web emite con información incompleta y el 96.7 % de los colaboradores tiene su nivel de riesgo desactualizado. En el período no se presentaron ROS y no existe un registro que documente por qué; el caso de Fianzas de 2024 muestra que no todas las señales de alerta llegan a Cumplimiento.',
  'La evidencia de gobierno del período es parcial: hay informes de seguimiento al Comité y un acta de mayo de 2025, y la aclaración de Ancón explica que en 2024 hubo sesiones de seguimiento; pero las copias de esas actas, las ratificaciones de la Junta y las copias de las matrices de riesgo aprobadas están pendientes de entrega. Completar ese expediente es una tarea relativamente sencilla y de alto impacto ante la SSRP.',
  `Nuestra conclusión es que el sistema de prevención se encuentra en **Nivel 2 – ALTO**, con ${TOT['CRÍTICO']} brechas críticas que requieren acción en los próximos 30 días. Si se ejecutan las acciones inmediatas —separación de funciones del Oficial, actualización con evidencia del cierre de hallazgos de la SSRP, bloqueo de emisión sin debida diligencia completa, registro de operaciones inusuales y entrega de las matrices de riesgo—, la entidad estará en buena posición para avanzar con rapidez.`,
  'Por último, el cambio normativo es una oportunidad: como la mayor parte de la remediación todavía está por diseñarse, Ancón puede construirla directamente con el estándar del Acuerdo 07-2025 y resolver con un solo esfuerzo las brechas del período y la transición al nuevo marco. La Sección 11 está pensada para eso.',
].forEach(t => S3.push(P(t)));
S3.push(T1('9. Recomendaciones Estratégicas y Decisiones Críticas'));
S3.push(P('Las recomendaciones operativas están detalladas por componente en la Sección 6. Aquí resumimos las decisiones que corresponden a la Junta Directiva y a la Alta Gerencia, porque sin ellas el resto del plan no puede avanzar.'));
S3.push(table(['#', 'Decisión requerida', 'Órgano', 'Oportunidad', 'Por qué es importante'], [
  ['1', 'Separar las funciones de Cumplimiento y Legal y aprobar personal profesional dedicado a PBC/FT/FPADM.', 'Junta Directiva', 'Próxima sesión', 'Es una infracción tipificada y una observación reiterada de la SSRP.'],
  ['2', 'Aprobar la matriz de seguimiento del cierre de hallazgos SSRP 2022 actualizada con evidencia y decidir si se comparte de forma proactiva con el supervisor.', 'Comité / Junta Directiva', '30 días', 'Permite acceder al atenuante por subsanación voluntaria (hasta 40 %).'],
  ['3', 'Asignar presupuesto y fecha comprometida a DEIVID, incluidos el bloqueo de emisión y el monitoreo automatizado.', 'Gerencia General / Junta Directiva', '30 días', 'De ello dependen la debida diligencia y el monitoreo.'],
  ['4', 'Consolidar y aprobar la Evaluación Institucional de Riesgo BC/FT/FPADM a partir de las matrices ya aprobadas.', 'Comité / Junta Directiva', '90 días', 'Es la base que justifica todo el enfoque basado en riesgo.'],
  ['5', 'Instruir el análisis del caso de Fianzas de 2024, la creación del registro de operaciones inusuales y la circular de reporte obligatorio.', 'Gerencia General / Oficial de Cumplimiento', '15 días', 'Evita una posible omisión de reporte.'],
  ['6', 'Adoptar la hoja de ruta hacia el Acuerdo 07-2025 (Sección 11) y confirmar que la adecuación de manuales y políticas esté formalizada.', 'Comité / Junta Directiva', '30 días', 'Evita duplicar esfuerzos y alinea la remediación con la norma vigente.'],
], [400, 3000, 1600, 1300, 3060], { size: 16 }));
S3.push(P('**Dependencias:** las acciones de debida diligencia, monitoreo y reportes dependen de la dotación de personal (decisión 1) y del avance de DEIVID (decisión 3). Recomendamos que el Comité de Cumplimiento revise el avance del plan cada mes hasta su cierre y que el Oficial de Cumplimiento informe a la Junta en cada sesión trimestral.', { before: 120 }));

// ---- 10, 11 (landscape) ----
const S4 = [];
S4.push(T1nb('10. Plan de Acción Propuesto'));
S4.push(P('Este plan atiende las brechas del período. Los plazos cuentan desde que la Junta Directiva apruebe este informe y son consistentes con la acción requerida por nivel de riesgo de la matriz. Cada acción se cierra con la evidencia indicada, verificada por el Oficial de Cumplimiento y reportada al Comité. Recomendamos ejecutarlo con los criterios del Acuerdo 07-2025 descritos en la Sección 11. Componentes: GC Gobierno Corporativo; PP Políticas y Procedimientos; PAC Programa Anual; MT Monitoreo Transaccional; DD Debida Diligencia; ER Evaluación de Riesgo; RR Reportes Regulatorios; CD Conservación de Documentos.'));
let lastH = null; const paRows = [];
PA.forEach(a => { if (a[0] !== lastH) { paRows.push([{ t: `Horizonte ${a[0]}`, bold: true, fill: LIGHT, span: 7 }]); lastH = a[0]; } paRows.push([{ t: a[1], bold: true }, a[2], a[3], lvlCell(a[4]), a[5], a[6], a[7]]); });
S4.push(table(['ID', 'Acción', 'Comp.', 'Prioridad', 'Responsable', 'Plazo', 'Evidencia de cierre'], paRows, [800, 4660, 800, 1000, 2100, 1500, 2100], { size: 16 }));
S4.push(new Paragraph({ children: [new PageBreak()] }));
S4.push(T1nb('11. Valor Agregado: Hoja de Ruta hacia el Acuerdo SSRP 07-2025'));
S4.push(T2('11.1 Contexto y propósito'));
S4.push(P('El Acuerdo SSRP 07-2025 se publicó en Gaceta Oficial el 10 de diciembre de 2025 y rige desde esa fecha (Art. 48). Otorgó 60 días calendario para adecuar manuales, políticas y controles (Art. 46), derogó parcialmente el Acuerdo 03-2022 en lo relativo al Grupo A y derogó el Acuerdo 02-2015 sobre señales de alerta (Art. 47). Como se publicó después del período evaluado, no formó parte de nuestro alcance ni de las calificaciones de la Parte I.'));
S4.push(P('Con base en lo que observamos, preparamos este catálogo para que Ancón tenga una visión clara de qué cambia, dónde está hoy y qué conviene hacer. Cada mejora se relaciona con las brechas y con el componente correspondiente, de modo que el Plan de Acción y esta hoja de ruta se ejecuten como un solo esfuerzo. En la matriz de riesgo, estas mejoras se registran como oportunidades de mejora. Recomendamos que Ancón confirme, como primer paso, el estado de adecuación de su Manual y sus políticas a este Acuerdo.'));
S4.push(T2('11.2 Hoja de ruta por fases'));
S4.push(table(['Fase', 'Horizonte', 'Propósito', 'Mejoras incluidas'], FASES, [2800, 1400, 5160, 3600], { size: 17, firstColBold: true }));
S4.push(T2('11.3 Catálogo de mejoras y optimizaciones', true));
S4.push(table(['ID', 'Tema (artículo)', 'Qué cambia', 'Situación observada en Ancón', 'Mejora recomendada por GSM', 'Fase', 'Comp.'],
  RM.map(r => [{ t: r[0], bold: true }, { t: r[1], bold: true }, r[2], r[3], r[4], { t: r[5], align: AlignmentType.CENTER, fill: r[5] === 'Fase 1' ? 'FBE0C3' : r[5] === 'Fase 2' ? 'FDF1E4' : 'F4F7F8' }, r[6]]),
  [700, 1900, 2500, 2300, 3660, 900, 1000], { size: 15 }));
S4.push(T2('11.4 Otras oportunidades de mejora del período'));
S4.push(table(['ID', 'Oportunidad', 'Descripción', 'Documento', 'Acción sugerida'], OMX.map(o => [{ t: o[0], bold: true }, { t: o[1], bold: true }, o[2], o[3], o[4]]), [800, 2700, 4560, 2100, 2800], { size: 16 }));
S4.push(P('La hoja de ruta tiene carácter de recomendación y se basa en la evidencia obtenida para el período evaluado. Su aplicación debe validarse con la asesoría legal de Ancón y, cuando corresponda, con la SSRP.', { italics: true, size: 17, color: GREY, before: 120 }));

// ---- Anexos ----
const S5 = [];
S5.push(T1nb('Anexo A. Documentos Revisados'));
S5.push(table(['#', 'Documento', 'Versión / Fecha', 'Componente'], [
  ['1', 'POL-CUM-27-MPREMET-21 – Manual de Prevención', 'v3.0 / oct-2023', 'PP / DD'],
  ['2', 'CONCLI-15 – Conoce al Consumidor; política de identificación y verificación; política del subyacente que cede el riesgo; política de actualización del consumidor', 'v4.0 / may-2023', 'PP / DD'],
  ['3', 'GESADHE-12 – Herramientas Tecnológicas; MPRMOPTR-20 – Perfil Financiero y Transaccional', 'v1.0 / ene-2023', 'PP / MT'],
  ['4', 'DNR Alertas Monitoreos; parametrización de 13 tipologías; Informe de Tipologías; muestra de alerta; registros de alertas', 'oct-2023 a 2026', 'MT'],
  ['5', 'Señales de Alerta de Cumplimiento 2025; catálogo de tipologías UAF', '2025', 'MT'],
  ['6', 'Informe de Fraude a la Junta Directiva', '09-06-2025', 'MT'],
  ['7', 'Evaluaciones de Riesgo de Emisión (v2) y Reclamos (v0); Hoja Modelo ERM', '21-07-2025', 'ER'],
  ['8', 'CAPAC-8 y CAPPCCASC-13 – Capacitación; cuadro de capacitaciones 2024-2025; certificados y sustentos', '2023-2026', 'PP / PAC'],
  ['9', 'RESYCON-7 – Reserva y Confidencialidad; política de conservación de documentos', 'may-2023; abr-2024', 'PP / CD'],
  ['10', 'CONASEM-5 y DPRPACO-17 – Conoce a tu Colaborador y Actualización; Política de Reclutamiento', 'may-2023; 2026', 'PP / DD'],
  ['11', 'CONSUREG-10, DPRPASO-19, CONCOMGF-11 – Sujetos Regulados Grupos A y B', '2023', 'PP / DD'],
  ['12', 'CONTUPROV-6 y DPRPAPR-18 – Proveedores; plan de actualización de proveedores', 'may-2023', 'PP / DD'],
  ['13', 'Manual de Gobierno Corporativo; Reglamento del Comité AA_GOBCORP03 y acta de Junta que lo aprueba', 'V03 / jun-2019', 'GC'],
  ['14', 'Organigramas general y del Departamento de Cumplimiento; hojas de vida; Resolución OAL-362', 'sep-2025; 30-12-2022', 'GC'],
  ['15', 'Código de Ética AIN-28-CODETI; Manual del Sistema de Control Interno; Manual de RR.HH.', '2023-2024', 'GC'],
  ['16', 'Actas del Comité de 17-04-2024 y 28-05-2025; acta de Junta de 20-01-2021; autoevaluación del Comité', '2021-2025', 'GC'],
  ['17', 'Aclaración de aprobaciones en actas del Comité de Cumplimiento (IV trim. 2023 – 2025)', 'sep-2026', 'GC / PAC'],
  ['18', 'Informes de seguimiento al Comité (I, II, III-IV trimestre 2024; I y II trimestre 2025)', '2024-2025', 'GC / PAC'],
  ['19', 'Programas y cronogramas 2024 y 2025; actas de Junta aportadas; correo de remisión 2025', '2024-2026', 'PAC'],
  ['20', 'Informe de Gestión Anual 2024, resumen ejecutivo y cronograma de cierre; memorial Informe de Gestión 2023', '2024-2025', 'PAC / RR'],
  ['21', 'Planes de Auditoría Interna 2024 y 2025; Informe de Auditoría Interna al Dpto. de Cumplimiento', '2024-2026', 'GC'],
  ['22', 'Herramienta de listas, constancia de barridos y listado de pólizas rechazadas por coincidencias', '2024-2026', 'DD'],
  ['23', 'Listado de clientes con documentación pendiente; formularios CTC', '2025-2026', 'DD'],
  ['24', 'Matriz de revisión de 85 expedientes de clientes; análisis de 33 expedientes de colaboradores', 'ago-2026', 'DD'],
  ['25', 'Entrevistas a 13 responsables y colaboradores', 'ene-2026', 'Transversal'],
  ['26', 'Informe SSRP DSR-0924-2023; matriz de seguimiento de cierre de hallazgos; última evaluación independiente', '2023-2026', 'Transversal'],
  ['27', 'Ley 23 de 2015; D.E. 35 de 2022; Acuerdos SSRP 03-2022 y 07-2025', 'Gaceta Oficial', 'Marco normativo'],
], [500, 5060, 2500, 1300], { size: 16, zebra: true }));
S5.push(T1('Anexo B. Glosario'));
S5.push(table(['Término', 'Significado'], [
  ['BC/FT/FPADM', 'Blanqueo de capitales, financiamiento del terrorismo y financiamiento de la proliferación de armas de destrucción masiva.'],
  ['CTC', 'Formulario Conozca a su Cliente.'],
  ['DEIVID', 'Sistema tecnológico institucional de Ancón utilizado para perfiles de riesgo, emisión y monitoreo.'],
  ['DLP', 'Prevención de pérdida de datos (Data Loss Prevention).'],
  ['GAFI', 'Grupo de Acción Financiera Internacional.'],
  ['P:I', 'Calificación de probabilidad e impacto en la matriz de riesgo GSM.'],
  ['PEP', 'Persona expuesta políticamente.'],
  ['ROI / ROS', 'Reporte de operación inusual (interno) / Reporte de operación sospechosa (a la UAF).'],
  ['SSRP', 'Superintendencia de Seguros y Reaseguros de Panamá.'],
  ['UAF', 'Unidad de Análisis Financiero para la Prevención del Delito de Blanqueo de Capitales y Financiamiento del Terrorismo.'],
], [2400, 6960], { size: 17, firstColBold: true }));
S5.push(T1('Anexo C. Estado de los Requerimientos de Información'));
S5.push(P('Resumen de los requerimientos de información realizados a Ancón durante la evaluación, según el cuadro de control de entregables con corte a octubre de 2026. Agradecemos la colaboración del equipo de Cumplimiento y de las áreas entrevistadas.'));
S5.push(table(['Sección', 'Requeridos', 'Aportados', 'Pendientes de entrega', 'No existen (según Ancón)'], [
  ['A. Gobierno AML y administración de riesgo', String(ENT.A.tot), String(ENT.A.rec), String(ENT.A.nr), String(ENT.A.ne)],
  ['B. Gobierno corporativo', String(ENT.B.tot), String(ENT.B.rec), String(ENT.B.nr), String(ENT.B.ne)],
  ['C. Capacitación y conozca a su colaborador', String(ENT.C.tot), String(ENT.C.rec), String(ENT.C.nr), String(ENT.C.ne)],
  [{ t: 'Total', bold: true, fill: LIGHT }, { t: '96', bold: true, fill: LIGHT }, { t: '72', bold: true, fill: LIGHT }, { t: '17', bold: true, fill: LIGHT }, { t: '7', bold: true, fill: LIGHT }],
  ['D. Entrevistas a áreas clave', '13', '13 realizadas', '–', '–'],
], [3960, 1300, 1300, 1300, 1500], { size: 17 }));
S5.push(P('Los documentos pendientes de entrega y los que Ancón indicó no tener se detallan en la Sección 3.3 y se reflejan en las brechas correspondientes. Su entrega posterior permitiría revisar la calificación de las brechas afectadas.', { before: 120 }));
S5.push(P('Este informe es un análisis técnico de cumplimiento y no sustituye la opinión legal de un abogado idóneo. Las interpretaciones normativas deben validarse con la asesoría legal de Ancón y, cuando corresponda, con la SSRP.', { italics: true, size: 17, color: GREY, before: 240 }));

// ---- Índice ----
const tocParas = [new Paragraph({ children: [new TextRun({ text: 'Índice', font: FONT, size: 32, bold: true, color: TEAL })], pageBreakBefore: true, spacing: { after: 240 } })];
TOC.forEach(([t, lvl]) => tocParas.push(new Paragraph({
  children: [new TextRun({ text: t, font: FONT, size: lvl === 1 ? 21 : 20, bold: lvl === 1 }), new TextRun({ children: [new Tab(), String(PAGES[t] || '00')], font: FONT, size: 20 })],
  tabStops: [{ type: TabStopType.RIGHT, position: PW, leader: LeaderType.DOT }],
  indent: { left: lvl === 2 ? 440 : 0 }, spacing: { after: lvl === 1 ? 80 : 40, before: lvl === 1 ? 40 : 0 }
})));
S1.splice(TOC_POS, 0, ...tocParas);
fs.writeFileSync('toc.json', JSON.stringify(TOC.map(x => x[0])));

const doc = new Document({
  creator: 'GSM Consulting', title: 'Informe Final Consolidado – Evaluación Independiente PBC/FT/FPADM – Aseguradora Ancón',
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 30, bold: true, font: FONT, color: TEAL }, paragraph: { spacing: { before: 120, after: 200 }, outlineLevel: 0, keepNext: true, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ORANGE, space: 4 } } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 26, bold: true, font: FONT, color: TEAL }, paragraph: { spacing: { before: 240, after: 140 }, outlineLevel: 1, keepNext: true } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 22, bold: true, font: FONT, color: ORANGE }, paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2, keepNext: true } },
    ]
  },
  numbering: { config: ['msg', 'obj', 'two'].map(ref => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 440, hanging: 360 } } } }] })) },
  sections: [
    { properties: portrait, headers: { default: hdr(PW) }, footers: { default: ftr(PW) }, children: S1 },
    { properties: landscape, headers: { default: hdr(LW) }, footers: { default: ftr(LW) }, children: S2 },
    { properties: portrait, headers: { default: hdr(PW) }, footers: { default: ftr(PW) }, children: S3 },
    { properties: landscape, headers: { default: hdr(LW) }, footers: { default: ftr(LW) }, children: S4 },
    { properties: portrait, headers: { default: hdr(PW) }, footers: { default: ftr(PW) }, children: S5 },
  ]
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.env.OUT || 'Informe_Final.docx', b); console.log('OK', TOTAL, JSON.stringify(TOT), 'N1:', N1, C.map(c => c.num + ':' + c.nivel).join(' '), JSON.stringify(HEAT)); });
