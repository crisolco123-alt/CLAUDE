'use strict';
// Documento "Curs Clínic Argos" de prueba, en sucio, con todos los casos que trata la skill.

const GRIS = '#656565', AZUL = '#2196F3';
const N = (text, font) => ({ text, style: 'Normal', font });
const ARG = (text, extra = {}) => ({ text, style: 'Normal', font: Object.assign({ name: 'Arial', size: 12 }, extra.font), list: extra.list, img: extra.img });
const TIMES = (text, extra = {}) => ({ text, style: 'Normal', font: { name: 'Times New Roman', size: 12 }, list: extra.list });
const B = () => N('');
const H = (n, text) => ({ text, style: 'Título ' + n });
const L = '\u000b';

// nota de Argos en sucio: [CC] fecha / hora / nombre azul / 2 líneas grises / cuerpo
const nota = (cc, fecha, hora, nombre, servei, carrec, cos, carrecMezclado) => [
  ...(cc ? [ARG('CC', { font: { color: '#FFFFFF' } })] : []),
  ARG(fecha), ARG(hora), ARG(nombre, { font: { color: AZUL } }),
  ARG(servei, { font: { color: GRIS } }),
  ARG(carrec, { font: { color: carrecMezclado ? null : GRIS } }),
  ...cos.map(t => ARG(t)),
];

function documento() {
  return [
    H(1, 'ROSA POUS MASO (D, 83a)'),
    { text: 'ÍNDEX', style: 'Título TDC' },
    { text: 'PROVES\t1', style: 'TDC 1' },
    { text: '29/09/2026 — Función renal. Estudio genético indeterminado. Calcio sérico\t1', style: 'TDC 3' },

    // ── PROVES ──
    H(2, 'PROVES PENDENTS'),
    // formato antiguo
    N('ANALÍTIQUES'),
    N('28/09/2026 (10:00) — HEMOGRAMA, PROTEÏNA C REACTIVA Sèrum' + L + 'Pend. programar', { bold: true }),
    N('Hemograma' + L + 'Proteïna C Reactiva Sèrum'),
    B(),
    N('DIAGNOSTIC IMATGE'),
    N('24/09/2026 (11:29) — Rx de tòrax >2 projeccions', { bold: true }),
    N('23/09/2026 (09:00) — Ecografia abdominal', { bold: true }),
    B(),
    // formato nuevo, grupo con recuento en castellano y línea en blanco detrás
    N('ANALÍTIQUES (3 realizadas, 1 no realizada)'),
    B(),
    N('29/09/2026' + L + 'Función renal. Estudio genético indeterminado. Calcio sérico' + L +
      'Lab. Hospital de Mataró – Laboratori' + L + 'Pendiente de programar', { bold: true }),
    TIMES('Función renal', { list: true }),
    TIMES('Estudio genético indeterminado', { list: true }),
    TIMES('Calcio sérico', { list: true }),
    B(),
    N('27/09/2026 (08:15)' + L + 'Microalbúmina. Cocient albúmina/creatinina en orina' + L +
      'Lab. Hospital de Mataró – Laboratori' + L + 'Realizada', { bold: true }),
    TIMES('Microalbúmina. Cociente albúmina/creatinina en orina esporádica', { list: true }),
    TIMES('Creatinina en orina esporádica', { list: true }),
    TIMES('Tiroxina libre (T4 libre) sérica', { list: true }),
    B(),
    N('26/09/2026' + L + 'Glucosa Sèrum' + L + 'Lab. Hospital de Mataró – Laboratori' + L + 'Realitzada', { bold: true }),
    TIMES('Glucosa Sèrum', { list: true }),
    B(),
    N('25/09/2026' + L + 'Función hepática. Proteína C reactiva', { bold: true }),
    TIMES('Función hepática', { list: true }),
    TIMES('Proteína C reactiva sérica', { list: true }),
    TIMES('Ácido úrico sérico', { list: true }),
    B(),
    N('24/09/2026 (12:00)' + L + 'Hemograma' + L + 'Lab. Hospital de Mataró – Laboratori' + L + 'Anulada', { bold: true }),
    TIMES('Hemograma', { list: true }),
    B(),
    N('Septiembre de 2026' + L + 'Prueba de la marcha de 6 minutos' + L + 'Gabinet de Pneumologia' + L + 'Pendiente de programar', { bold: true }),
    B(),
    B(),
    ARG('', { img: 1 }),
    N('Setembre de 2026' + L + 'Test de marxa de 6 minuts' + L + 'Gabinet de Pneumologia' + L + 'Programada', { bold: true }),
    B(),

    // ── INTERCONSULTES ──
    H(2, 'INTERCONSULTES PENDENTS'),
    N('NUTRI'),
    N('Realizadas (1)'),
    N('DIANA GONZALEZ GOMEZ (21/09/2026)'),
    N('Valoració nutricional, si us plau.'),
    B(),
    N('Pacient amb pèrdua de pes.'),
    N('Respuestas: '),
    N('22/09/2026 (12:15) — GLORIA SALAS FRANCO' + L + 'Veure CC'),
    N('No realizadas (0)'),
    N('Ninguna.'),
    B(),
    N('ANESTESIA'),
    N('Realizadas (1)'),
    N('NATALIA GIL ALIBERAS (17/09/2026)'),
    N('Preoperatori de pròtesi de maluc.'),
    N('Respuestas:'),
    N('17/09/2026 (16:17) — VIRGINIA RADUA GIMENEZ' + L + 'Apte per a la intervenció. ASA III.'),
    N('No realizadas (0)'),
    N('Ninguna.'),
    B(),
    ...nota(false, '22/09/2026', '12:40', 'gloria salas franco', 'NUTRICIÓ I DIETÈTICA / Unitat de Nutrició',
      'Dietista/Nutricionista NUTRICIÓ I DIETÈTICA', ['Valoració: desnutrició moderada.', 'Pla: suplements.']),
    ...nota(true, '24/09/2026', '10:00', 'gloria salas franco', 'NUTRICIÓ I DIETÈTICA / Unitat de Nutrició',
      'Dietista/Nutricionista NUTRICIÓ I DIETÈTICA', ['Seguiment: bona tolerància als suplements.']),
    ...nota(true, '23/09/2026', '11:00', 'marta fisio lopez', 'REHABILITACIÓ / Unitat',
      'Altres treballadors clínics FISIOTERAPEUTA', ['Sedestació assolida.']),
    ...nota(true, '14/09/2026', '09:51', 'eloi cañamero giro', 'HEMATOLOGIA I HEMOTERÀPIA / Unitat',
      'Metge HEMATOLOGIA I HEMOTERÀPIA', ['Anèmia ferropènica. Ferro ev.']),
    ...nota(true, '21/09/2026', '13:00', 'gloria salas franco', 'NUTRICIÓ I DIETÈTICA / Unitat de Nutrició',
      'Dietista/Nutricionista NUTRICIÓ I DIETÈTICA', ['Veure curs clínic']),
    ...nota(true, 'NOTA APARELL DIGESTIU: 15/09/2026', '11:30', 'pep digestiu roca', 'APARELL DIGESTIU / Unitat',
      'Metge APARELL DIGESTIU', ['Sense signes d\'hemorràgia.']),

    // ── MEDICINA INTERNA ──
    H(2, 'MEDICINA INTERNA'),
    ...nota(false, '18/09/2026', '11:05', 'natalia gil aliberas', 'MEDICINA INTERNA / Unitat Hospitalització 5 (MT)',
      'Metge resident MEDICINA INTERNA',
      ['Evolutiu MI: 16è dia d\'ingrés', 'Dona de 83 anys, SAMC.', 'AP: HTA, Dislipèmia', 'AVUI: afebril.', 'PLA: seguir igual.'], true),
    ...nota(true, '17/09/2026', '16:04', 'natalia gil aliberas', 'MEDICINA INTERNA / Unitat Hospitalització 5 (MT)',
      'Metge resident MEDICINA INTERNA', ['EVOLUCIÓ: estable.', 'EF: normal.']),
    // nota ya formateada (uso incremental) + nota antigua v1
    { text: '16/09/2026 (10:00) → NATALIA GIL ALIBERAS', style: 'Título 3', font: { size: 12 } },
    { text: '', style: 'MiniEspacio', font: { size: 1 } },
    N('Nota ja formatada.', { size: 10 }),
    B(),
    { text: '15/09/2026 - 09:30', style: 'Título 3' },
    B(),
    ARG('aleix serrallonga', { font: { color: AZUL } }),
    ARG('MEDICINA INTERNA / Unitat', { font: { color: GRIS } }),
    ARG('Metge MEDICINA INTERNA', { font: { color: GRIS } }),
    B(),
    ARG('NOTA D\'INGRÉS'),
    ARG('JC: ICC descompensada.'),

    // ── MFiC ──
    H(2, 'FAMILIAR I COMUNTÀRIA'),
    ...nota(false, '10/09/2026', '09:00', 'metge familia pérez', 'EAP MATARÓ / Consulta', 'Metge MEDICINA FAMILIAR',
      ['Control TA.', 'PLA: revisió en 1 mes.']),

    // ── INFERMERIA ──
    H(2, 'INFERMERIA'),
    ...nota(false, '14/09/2026', '08:00', 'infermera garcia', 'INFERMERIA / Unitat 5', 'Infermera INFERMERIA',
      ['PLA: cures de la ferida.']),

    // ── CONSTANTS · SILICON ──
    H(2, 'CONSTANTS'),
    ARG('TA 130/80 FC 72'),
    ARG('', { img: 1 }),
    H(2, 'SILICON'),
    ARG('Pauta'),
  ];
}

// lo que Cris pega al día siguiente sobre el documento ya formateado
function pegadoNuevo() {
  return {
    finProves: [
      N('30/09/2026' + L + 'Magnesio sérico. Fosfato sérico' + L + 'Lab. Hospital de Mataró – Laboratori' + L + 'Programada', { bold: true }),
      TIMES('Magnesio sérico', { list: true }),
      TIMES('Fosfato sérico', { list: true }),
      B(),
    ],
    inicioMI: nota(false, '19/09/2026', '12:00', 'natalia gil aliberas', 'MEDICINA INTERNA / Unitat Hospitalització 5 (MT)',
      'Metge resident MEDICINA INTERNA', ['Evolutiu MI: 17è dia d\'ingrés', 'AVUI: millor.']),
  };
}

module.exports = { documento, pegadoNuevo };
