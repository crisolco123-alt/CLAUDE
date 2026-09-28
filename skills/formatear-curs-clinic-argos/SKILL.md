---
name: formatear-curs-clinic-argos
description: Da formato al "Curs Clínic Argos" de un paciente (CSdM) en Word con UNA llamada execute_office_js, entero o solo lo pegado en sucio ese día, igual que el de referencia de Rosa Pous. Notas MI / MFiC / INFERMERIA → Título 3 "dd/mm/yyyy (hh:mm) → NOMBRE" + cuerpo Calibri Light 10 (quita barras "CC" y líneas grises de autor, migra el formato antiguo "dd/mm/yyyy - hh:mm"). INTERCONSULTES pegadas de la app → "SERVEI – Realitzades (n)", petición y "RESPOSTES:", sustituyendo los "veure curs clínic" por la nota del curs clínic. PROVES de Argos, formato antiguo ("fecha — PERFILS") y nuevo ("fecha⏎perfils⏎laboratori⏎estat", una determinación por párrafo) → Título 4 + estado en Énfasis sutil + viñetas, traducidas al catalán si llegan en castellano. Actualiza el índice. Úsala cuando Cris pida formatear un Curs Clínic, "ponerlo como siempre" o "como el de Rosa Pous", o tras pegar notas o pruebas nuevas.
---

## Cuándo usarla

Cuando Cris tenga abierto en Word un "Curs Clínic Argos - <paciente>" (secciones Título 2: PROVES, INTERCONSULTES, MEDICINA INTERNA, MEDICINA FAMILIAR I COMUNITÀRIA, INFERMERIA, CONSTANTS, SILICON) y pida formatearlo.

Es la versión única y definitiva: sustituye a `formatear-curs-clinic-argos` (v1/v2) y a `formatear-curs-clinic-argos-v3`. Si alguna de ellas sigue instalada, hay que borrarla: con la antigua, las PROVES del formato nuevo se quedan en sucio y el control no lo avisa.

**Uso diario incremental.** Cris pega cada día lo nuevo de Argos dentro del documento ya formateado (o dentro de la plantilla vacía de un paciente nuevo) y relanza la skill. Solo se toca lo que está en sucio: lo ya formateado no encaja en ningún patrón de búsqueda y se salta solo. Relanzarla sobre un documento ya formateado no cambia nada (comprobado sobre el de Rosa Pous).

## Regla nº 1: el documento no entra en el chat

Todo se hace **dentro de un único `execute_office_js`**. El código recibe `context` ya dentro de `Word.run` (no lo envuelvas en otro) y lo único que vuelve al chat es el `return` con una línea de log. Un Curs Clínic tiene 300-1.700 párrafos: volcarlos al chat multiplica el coste de cada llamada posterior y provoca errores de capacidad ("high demand"). Localizar "CC", un `dd/mm/yyyy`, un `hh:mm` o un `Heading2` es comparación de cadenas: lo hace el JavaScript. Nunca leas párrafos "para ver qué hay" antes de escribir; si hace falta saber algo, averígualo dentro del script y sácalo en el log.

La traducción al catalán también va dentro de esa llamada (diccionario `DIC`). Solo si el log trae **"CASTELLANO PENDIENTE"** (términos que aún no están en el diccionario) se hace una **segunda llamada** corta para esos términos (ver "Traducción al catalán").

---

## FORMATO DE REFERENCIA ("Curs Clínic Argos - Rosa Pous Maso", 26/09/2026)

Todo el documento: **Calibri Light, negro, justificado, espaciado 0/0, interlineado sencillo**, salvo lo indicado. Márgenes 1 cm.

### Estilos (vienen en la plantilla de Cris)

| Estilo | Fuente | Párrafo |
|---|---|---|
| Normal | Calibri Light 11 negro | justificado, 0/0, sencillo |
| Título 1 | 20 pt negrita azul `2E74B5` | interlineado 1,5 |
| Título 2 | 16 pt negro | borde `single` 4 pt los 4 lados · sombreado `D9D9D9` · salto de página antes · siguiente: MiniEspacio |
| Título 3 | 12 pt negro | borde `dotted` los 4 lados · sombreado `F2F2F2` · siguiente: MiniEspacio |
| Título 4 | 11 pt azul `2E74B5` | sombreado `DEEAF6` · sangría 1ª línea 7,1 pt |
| Título 5 | 11 pt azul `2E74B5` | espacio antes 2 pt |
| MiniEspacio | Calibri Light **1 pt** | espacio antes 6 pt (separador fino bajo un título) |
| Lista con viñetas | Calibri Light 10 | viñeta Symbol "" de 8 pt · viñeta a 0,4 cm, texto a 0,8 cm (izq. 22,7 pt, francesa 11,35 pt) |
| Énfasis sutil (carácter) | 10,5 cursiva naranja `ED7D31` | — |
| TDC 1 / 2 / 3 | 12 / 10,5 / 8 pt | TDC 1 sombreado `E7E6E6`; tabulador derecho con puntos |

Cabecera: `[Título 1] NOMBRE APELLIDOS (D, 83a)` · `[Título TDC] ÍNDEX` · campo TOC `\h \z \u \t "Título 2;1;Título 3;2;Título 4;3"`.

### PROVES

```
[Título 2]  PROVES
[Título 3]  ANÀLISIS CLÍNIQUES (8 realitzades, 1 no realitzada)   ← grupo (el recuento, si Argos lo da)
[MiniEspacio]
[Título 4]  29/09/2026 — FUNCIÓ RENAL, ESTUDI GENÈTIC INDETERMINAT, CALCI Sèrum   ← "fecha (hh:mm) — PERFILS"
[Normal + Énfasis sutil] PENDENT DE PROGRAMAR       ← solo si hay estado; "Realitzada" no se pone
[Lista con viñetas] Funció Renal                    ← una viñeta por determinación
[Lista con viñetas] Calci Sèrum
[Normal]    (vacío)                                 ← un único blanco tras cada analítica
[Título 4]  24/09/2026 (11:29) — Rx de tòrax >2 projeccions
[MiniEspacio]                                       ← tras prueba sin lista (imagen, ECG…) si sigue otra prueba
[Título 4]  Setembre de 2026 — Test de marxa de 6 minuts
[Normal + Énfasis sutil] PENDENT DE PROGRAMAR
[Normal]    (vacío)                                 ← antes del siguiente Título 3
```

Grupos: ANÀLISIS CLÍNIQUES, CARDIOLOGIA, DIAGNÒSTIC D’IMATGE… Llegan como texto en MAYÚSCULAS (con o sin recuento) o ya como Título 3. Se normaliza `ANALÍTIQUES` / `ANÁLISIS CLÍNICOS` → `ANÀLISIS CLÍNIQUES`, `DIAGNOSTIC IMATGE` / `DIAGNÓSTICO POR IMAGEN` → `DIAGNÒSTIC D’IMATGE`, y el recuento a catalán: `(3 realizadas, 1 no realizada)` → `(3 realitzades, 1 no realitzada)`.

Criterio del texto en catalán: en el Título 4, perfiles en MAYÚSCULAS separados por coma y la muestra en tipo título (`Sèrum`, `Orina esporàdica`, `Orina 24 h`); en las viñetas, tipo título (`Funció Renal`, `Paratirina (PTH Intacta) Sèrum`).

**Formato ANTIGUO de Argos:** cada prueba es un párrafo en negrita `dd/mm/yyyy (hh:mm) — PERFILS` + salto de línea + estado en cursiva (`Pend. programar`, `Programada`, `Resultats parcials`), y debajo otro párrafo con todas las determinaciones separadas por saltos de línea. Llega en catalán.

**Formato NUEVO de Argos (desde 09/2026):** un párrafo en negrita con saltos de línea (Mayús+Intro):

```
29/09/2026                                     ← fecha, fecha (hh:mm) o "Setembre de 2026"
Función renal. Estudio genético indeterminado. Calcio sérico   ← perfiles
Lab. Hospital de Mataró – Laboratori           ← lugar (se borra)
Pendiente de programar                         ← estado
```

y debajo **cada determinación en su propio párrafo** (Normal o numerado, Times 12), hasta la siguiente línea vacía. Suele llegar **en castellano**.

La skill convierte ambos al formato de arriba: Título 4 `fecha — perfiles` (en el nuevo: borra desde el 2º salto de línea hasta el final, cambia los perfiles por su traducción si hace falta y el 1er salto por " — "; nunca borra el párrafo entero), estado en Énfasis sutil (`Pendiente de programar`/`Pend. programar` → `PENDENT DE PROGRAMAR`, `Programada` → `PROGRAMADA`, `Resultados…` → `RESULTATS PARCIALS`, `Anulada` → `ANUL·LADA`, `Realizada` → nada), viñetas Calibri Light 10 (las listas numeradas de Argos se desvinculan antes) y separadores: un solo blanco tras cada analítica, fuera MiniEspacios y blancos duplicados, viñetas vacías sueltas → blanco Normal, blanco justo debajo de un grupo nuevo → MiniEspacio. Lo que llegue en castellano se traduce (ver "Traducción al catalán"), incluido el mes (`Septiembre de 2026` → `Setembre de 2026`).

### INTERCONSULTES (referencia: "Curs Clínic Argos - Enrique Rus Sanchez nou.docx")

```
[Título 2]  INTERCONSULTES
[Título 3]  ANESTESIOLOGIA I REANIMACIÓ – Realitzades (1)   ← servicio + nº de IC realizadas
[MiniEspacio]
[Título 4]  (17/09/2026) → NATALIA GIL ALIBERAS              ← PETICIÓN: fecha entre paréntesis (sin hora) + quien la pide
[MiniEspacio]
[Normal 11] texto de la petición
[MiniEspacio]
[Título 5]  RESPOSTES:
[MiniEspacio]
[Título 4]  17/09/2026 (16:17) → VIRGINIA RADUA GIMENEZ      ← RESPUESTA: fecha (hora) + quien responde (sin categoría)
[MiniEspacio]
[Normal 11] texto de la respuesta
[Normal]    (vacío)
[Título 4]  …respuesta anterior…                             ← de más reciente a más antigua
…
[Título 3]  CC HEMATOLOGIA I HEMOTERÀPIA                     ← notas del curs clínic sin servicio, AL FINAL
[MiniEspacio]
[Título 4]  14/09/2026 (09:51) → ELOI CAÑAMERO GIRO
```

**Cómo llega de la app del hospital** (se pega la lista de interconsultes):

```
NUTRI                                  ← nombre del servicio (a veces abreviado)
Realizadas (2)
DIANA GONZALEZ GOMEZ (21/09/2026)      ← petición
texto de la petición…
Respuestas: [Sense respostes]
22/09/2026 (12:15) — GLORIA SALAS FRANCO⏎Veure CC      ← respuesta (fecha, hora, nombre, salto de línea, texto)
No realizadas (0)
Ninguna.
```

y **debajo, las notas del curs clínic** de esos especialistas (en sucio o ya en Título 4), porque cuando la app dice "veure curs clínic" el texto de verdad está allí.

**Qué hace la skill (fase B):**

1. `SERVEI` + `Realizadas (n)` → Título 3 `SERVEI – Realitzades (n)`. Nombres abreviados → completos (tabla `NOMS`): ANESTESIA → ANESTESIOLOGIA I REANIMACIÓ, RHB → MEDICINA FÍSICA REHABILITACIÓ, NUTRI → NUTRICIÓ I DIETÈTICA, INFO ACOD → INF ACOD I CONTROL SINTROM (ampliar la tabla si aparecen más).
2. `NOMBRE (dd/mm/yyyy)` → Título 4 `(dd/mm/yyyy) → NOMBRE`; su texto en Normal 11.
3. `Respuestas:` / `Respostes:` → Título 5 **`RESPOSTES:`** (siempre en catalán). Si no hay ninguna respuesta: `Sense respostes.` debajo.
4. Cada respuesta de la app → Título 4 `dd/mm/yyyy (hh:mm) → NOMBRE` + texto.
5. `No realizadas (0)` + `Ninguna.` → **se borran**. Si hay no realizadas (>0) se deja el texto.
6. Notas del curs clínic que están debajo, en este orden:
   - Si solo dicen "veure curs clínic" / "Veure CC" → se borran (duplican la de la app).
   - Si hay una respuesta de la app del **mismo especialista y el mismo día** → el texto del curs clínic **sustituye** al de la app (con la hora del curs clínic). Así se resuelven los "veure curs clínic".
   - Si no, se busca el servicio (mismo especialista en otra respuesta, o la categoría del cargo parecida al servicio, tabla `SIN`: DIETISTA ↔ NUTRICIÓ, FISIOTERAPEUTA/LOGOPEDA ↔ REHABILITACIÓ, PSIQUIATRIA ↔ SALUT MENTAL/ADDICCIONS, TREBALLADOR/A SOCIAL ↔ TREBALL SOCIAL) y se **añade** como respuesta a la petición de ese servicio con la fecha anterior más cercana, en su sitio por fecha.
   - Si no encaja en ningún servicio → se queda **al final** de INTERCONSULTES con Título 3 `CC <ESPECIALIDAD>` (p. ej. `CC CARDIOLOGIA`).
   - Si Cris ha escrito una etiqueta delante de la fecha (`NOTA APARELL DIGESTIU: 15/09/2026`), esa etiqueta pasa a ser el Título 3 de la nota.
7. Los cuerpos que se mueven se copian con `getOoxml` / `insertOoxml` para conservar negritas, subrayados, colores e imágenes.

Lo ya formateado (Título 3 con "– Realitzades", respuestas bajo un `RESPOSTES:`, notas bajo un Título 3 "CC …") no se vuelve a tocar.

### MEDICINA INTERNA · MEDICINA FAMILIAR I COMUNITÀRIA · INFERMERIA

```
[Título 2]  MEDICINA INTERNA
[MiniEspacio]
[Título 3]  18/09/2026 (11:05) → NATALIA GIL ALIBERAS
[MiniEspacio]
[Normal 10] Evolutiu MI: 16è dia d'ingrés          ← negrita + subrayado (también EVOLUCIÓ…, NOTA D'INGRÉS…, GUARDIA…)
[Normal 10] Dona de 83 anys, SAMC. No fumadora.
[Normal 10] AP: HTA, Dislipèmia…                  ← etiqueta en negrita: AP, AC, SB, MA, AVUI, EF, OD, PLA, ECG, JC
[Normal 10] …
[Normal 11] (vacío)                               ← separa cada nota
[Título 3]  17/09/2026 (16:04) → NATALIA GIL ALIBERAS
```

- Cuerpo en **Calibri Light 10**, conservando negritas, subrayados, cursivas, colores y sombreados del origen.
- El nombre pasa al Título 3 en MAYÚSCULAS; las líneas grises de Argos (servicio / cargo) desaparecen.
- INFERMERIA igual pero sin negrita en etiquetas (texto libre).

### CONSTANTS · SILICON

Solo el Título 2 y las imágenes/tablas que pegue Cris. No se tocan.

---

## Cómo llega cada nota de Argos (en sucio)

```
CC                          ← barra azul (no siempre; la 1ª de cada sección no la lleva)
17/09/2026                  ← fecha sola
16:24                       ← hora sola
natalia gil aliberas        ← nombre, azul #2196F3 (a veces mezclado con #0073E6)
MEDICINA INTERNA / Unitat Hospitalització 5 (MT)   ← gris #656565
Metge resident MEDICINA INTERNA                    ← gris, a veces MEZCLADO con azul
Evolutiu MI: …              ← cuerpo (Arial/Times 12)
```

Ojo: la línea de cargo suele venir con dos colores mezclados y Word devuelve `font.color = null`. Por eso una línea de autor se reconoce por color gris **o** (color mezclado **y** texto tipo "X / Y" o que empieza por Metge, Infermer, Altres treballadors, Auxiliar…). Solo se miran como mucho 2 líneas tras el nombre. De la última se saca la categoría (sin "Metge", "Infermera"…), que en INTERCONSULTES sirve para asignar la nota a su servicio.

**Notas del formato antiguo** (Título 3 `17/09/2026 - 16:24` + nombre + líneas grises como párrafos, de la v1) se migran al formato nuevo automáticamente.

## Dónde acaba cada nota

Desde su cabecera hasta el párrafo anterior al siguiente **corte**: una barra "CC", cualquier Título 1-5 (sección nueva o nota ya formateada) o el inicio de otra nota en sucio. El corte por título es imprescindible en el uso diario: sin él, una nota nueva pegada encima de otras ya formateadas se las "comería". Si justo debajo hay una nota ya formateada, se inserta la línea en blanco de separación; si la nota se ha pegado entre el Título 2 y su MiniEspacio, ese MiniEspacio se convierte en el blanco de separación (no se queda como cuerpo).

## Títulos de sección

- `PROVES PENDENTS` → **`PROVES`**
- `INTERCONSULTES PENDENTS` → **`INTERCONSULTES`**
- `FAMILIAR I COMUNTÀRIA` (errata de Argos) / `FAMILIAR I COMUNITÀRIA` → **`MEDICINA FAMILIAR I COMUNITÀRIA`**

Se reemplaza solo el contenido (`getRange("Content")`), sin tocar el formato del párrafo.

## Procedimiento: UNA sola llamada execute_office_js

Varios `context.sync()` dentro de la misma llamada:

| Sync | Qué hace |
|---|---|
| 1 | Única lectura del documento (texto, estilo, color, lista) + estilos de la plantilla + nº de imágenes |
| 1b | Busca los saltos de línea de las cabeceras nuevas de PROVES, desvincula las listas numeradas y mira qué párrafos "vacíos" tienen imagen |
| 2 | **Fase A**: títulos de sección, PROVES (con traducción), notas nuevas y migradas |
| 3 | Negrita de etiquetas + borrados (hora, nombre, líneas grises, "CC", blancos sobrantes) |
| 4-6 | **Fase B** INTERCONSULTES: relee → pide el OOXML de lo que se mueve → escribe y borra |
| 7-8 | Actualiza el índice (TDC) |
| 9 | Verificación → log |

La fase B y la TDC van en `try/catch`: si fallan, lo anterior ya queda guardado y el log lo avisa. Los borrados van **después** del commit del formato y **por referencia de objeto** (`it[j]` sigue apuntando al mismo párrafo aunque se inserten o borren otros), así que el orden da igual. Un párrafo "vacío" que contiene una imagen nunca se borra ni se convierte en separador.

```javascript
const paras = context.document.body.paragraphs;
paras.load("items/text,items/style,items/styleBuiltIn,items/font/color,items/isListItem");
const api15 = typeof Office !== "undefined" && Office.context.requirements.isSetSupported("WordApi", "1.5");
const estLista = api15 ? ["Lista con viñetas", "List Bullet"].map(n => {
  const e = context.document.getStyles().getByNameOrNullObject(n); e.load("isNullObject"); return e; }) : [];
const estMini = api15 ? context.document.getStyles().getByNameOrNullObject("MiniEspacio") : null;
if (estMini) estMini.load("isNullObject");
const fotos0 = context.document.body.inlinePictures; fotos0.load("items/width");
await context.sync(); // SYNC 1

const it = paras.items, N = it.length;
const RAW = it.map(p => p.text || ""), T = RAW.map(x => x.replace(/[\u000b\r\n]+/g, " ").trim());
const raw = i => RAW[i];
const txt = i => T[i];
const sb = i => it[i].styleBuiltIn;
const esH = i => /^Heading[1-5]$/.test(sb(i));
const esMini = i => /MiniEspacio/i.test(it[i].style || "");
const esLista = i => sb(i) === "ListBullet" || /vi[ñn]etas|List Bullet/i.test(it[i].style || "");
const col = i => (it[i].font.color || "").replace("#", "").toUpperCase();
const hayMini = estMini ? !estMini.isNullObject : it.some((p, i) => esMini(i));
const hayListaCasa = (estLista.length && !estLista[0].isNullObject) || it.some(p => /Lista con vi/.test(p.style || ""));

// ── helpers ──
const plano = (p, size) => {
  p.font.name = "Calibri Light"; p.font.size = size;
  p.font.bold = false; p.font.italic = false; p.font.underline = "None";
  p.spaceBefore = 0; p.spaceAfter = 0; p.lineSpacing = 12; p.alignment = "Justified";
};
const blanco = p => { p.styleBuiltIn = "Normal"; plano(p, 11); p.font.color = "#000000"; p.leftIndent = 0; p.firstLineIndent = 0; };
const vineta = p => {
  if (hayListaCasa) p.style = "Lista con viñetas"; else p.styleBuiltIn = "ListBullet";
  plano(p, 10); p.font.color = "#000000";
  p.leftIndent = 22.7; p.firstLineIndent = -11.35;
};
for (const e of estLista) if (!e.isNullObject) {
  e.font.name = "Calibri Light"; e.font.size = 10;
  e.paragraphFormat.leftIndent = 22.7; e.paragraphFormat.firstLineIndent = -11.35;
  e.paragraphFormat.spaceBefore = 0; e.paragraphFormat.spaceAfter = 0;
}
const mini = p => {
  if (hayMini) p.style = "MiniEspacio";
  else { p.styleBuiltIn = "Normal"; plano(p, 1); p.spaceBefore = 6; }
};
const H = (p, n, t) => {
  if (t != null) p.getRange("Content").insertText(t, "Replace");
  p.styleBuiltIn = "Heading" + n; plano(p, n === 3 ? 12 : 11); p.font.color = n === 3 ? "#000000" : "#2E74B5";
};

// ── 0) secciones ──
const CORTOS = { "PROVES PENDENTS": "PROVES", "INTERCONSULTES PENDENTS": "INTERCONSULTES",
                 "FAMILIAR I COMUNTÀRIA": "MEDICINA FAMILIAR I COMUNITÀRIA",
                 "FAMILIAR I COMUNITÀRIA": "MEDICINA FAMILIAR I COMUNITÀRIA" };
const sec = []; let actual = "", nTit = 0;
for (let i = 0; i < N; i++) {
  if (sb(i) === "Heading2") {
    actual = txt(i).toUpperCase();
    if (CORTOS[actual]) { it[i].getRange("Content").insertText(CORTOS[actual], "Replace"); actual = CORTOS[actual]; nTit++; }
  }
  sec.push(actual);
}

// ── 1) plan: notas ──
const RE_F = /^\d{2}\/\d{2}\/\d{4}$/, RE_H = /^\d{2}:\d{2}$/;
const RE_V1 = /^(\d{2}\/\d{2}\/\d{4}) - (\d{2}:\d{2})$/;
const RE_ROL = /^(Metge|Metgessa|Infermer|Altres treballadors|Auxiliar|TCAI|Tècnic|Fisioterapeut|Llevador|Psicòl|Farmac|Dietist|Treballador)/i;
const RE_FP = /^(.{2,60}?)\s*:\s*(\d{2}\/\d{2}\/\d{4})$/;
const esFecha = j => RE_F.test(txt(j)) || (sec[j] === "INTERCONSULTES" && RE_FP.test(txt(j)));
const inicioRaw = j => j + 1 < N && !esH(j) && esFecha(j) && RE_H.test(txt(j + 1));
const esNombre = j => j < N && !esH(j) && txt(j) !== "" && txt(j).length < 70 && !/[\d:]/.test(txt(j));
const esAutor = j => j < N && !esH(j) && txt(j) !== "" &&
  (col(j) === "656565" || (col(j) === "" && (/ \/ /.test(txt(j)) || RE_ROL.test(txt(j)))));
const corte = j => txt(j) === "CC" || esH(j) || inicioRaw(j);

const notas = [];
for (let i = 0; i < N - 1; i++) {
  let t = -1, fecha, hora, prefijo = "";
  if (inicioRaw(i)) {
    t = i + 1; hora = txt(i + 1); fecha = txt(i);
    const mp = RE_F.test(fecha) ? null : fecha.match(RE_FP);
    if (mp) { prefijo = mp[1].trim().toUpperCase(); fecha = mp[2]; }
  }
  else if (sb(i) === "Heading3" && RE_V1.test(txt(i))) { [, fecha, hora] = txt(i).match(RE_V1); }
  else continue;
  const borrar = t >= 0 ? [t] : [];
  let k = (t >= 0 ? t : i) + 1;
  if (k < N && txt(k) === "" && !esH(k) && esNombre(k + 1)) { borrar.push(k); k++; }
  let nombre = "", cat = "", g = k - 1;
  if (esNombre(k)) {
    nombre = txt(k).toUpperCase(); g = k;
    while (g - k < 2 && esAutor(g + 1)) g++;
    for (let j = k; j <= g; j++) borrar.push(j);
    if (g > k) cat = txt(g).replace(/^(Altres treballadors clínics|Altres treballadors|Metge resident|Metgessa|Metge|Infermera|Infermer)\s+/i, "").toUpperCase();
    if (g + 1 < N && txt(g + 1) === "" && !esH(g + 1)) { g++; borrar.push(g); }
  } else if (t < 0) continue;
  let fin = N - 1;
  for (let j = g + 1; j < N; j++) if (corte(j)) { fin = j - 1; break; }
  const sep = fin > g && esMini(fin) && fin + 1 < N && esH(fin + 1) ? fin : -1;
  if (sep >= 0) fin--;
  notas.push({ d: i, fecha, hora, nombre, cat, prefijo, borrar, ini: g + 1, fin, sep,
               cc: (i > 0 && txt(i - 1) === "CC") ? i - 1 : -1, s: sec[i], v1: t < 0 });
}

// ── 2) plan: PROVES ──
const RE_PROVA = /^\d{1,2}\/\d{1,2}\/\d{2,4}(\s*\(\d{1,2}:\d{2}\))?\s*[—–-]\s*\S/;
const RE_ESTAT = /^(PENDENT|PEND\.|PROGRAMAD|RESULTATS)/i;
const RE_FNOU = /^(\d{1,2}\/\d{1,2}\/\d{4}(\s*\(\d{1,2}:\d{2}\))?|[A-Za-zÀ-ÿ]+ de \d{4})$/;
const RE_LAB = /(^Lab\.|Laboratori|Gabinet|Hospital de|Servei de|Unitat )/i;
const RE_EST2 = /^(Pendiente|Pendent|Pend\.|Programad|Realizad|Realitzad|Resultad|Resultats|Anul)/i;
const RE_RECOMPTE = /\s*\((\d+)\s+reali(?:tz|z)ad[ae]s?(?:\s*,\s*(\d+)\s+no\s+reali(?:tz|z)ad[ae]s?)?\)$/i;
const lineas = j => raw(j).split(/[\u000b\r\n]+/).map(s => s.trim()).filter(Boolean);
const esNou = j => !esH(j) && /\u000b/.test(raw(j)) && RE_FNOU.test(lineas(j)[0] || "") && lineas(j).length >= 2;
const sinAcentos = s => s.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[’']/g, "'").toUpperCase();
const nomGrup = t => {
  const m = t.match(RE_RECOMPTE), base = m ? t.slice(0, m.index).trim() : t, k = sinAcentos(base);
  const nb = /^ANALI(SIS CLINIQUES|TIQUES|SIS CLINICOS)$/.test(k) ? "ANÀLISIS CLÍNIQUES"
           : /^DIAGNOSTICO? (D'|DE |POR )?(LA )?IMA(TGE|GEN)$/.test(k) ? "DIAGNÒSTIC D’IMATGE" : base;
  if (!m) return nb;
  const pl = n => +n === 1 ? "a" : "es";
  return `${nb} (${m[1]} realitzad${pl(m[1])}` + (m[2] != null ? `, ${m[2]} no realitzad${pl(m[2])}` : "") + ")";
};
const sigNoVacio = j => { for (let x = j + 1; x < N && sec[x] === "PROVES"; x++) if (txt(x) !== "") return x; return -1; };
const esGrupCand = j => {
  const t = txt(j).replace(RE_RECOMPTE, "").trim();
  if (t === "" || /\d/.test(t) || t !== t.toUpperCase() || t.length > 60 || RE_ESTAT.test(t)) return false;
  const s = sigNoVacio(j);
  return s >= 0 && (sb(s) === "Heading4" || (!esH(s) && (RE_PROVA.test(txt(s)) || esNou(s))));
};

const proves = [], proves2 = [], grups = [];
for (let i = 0; i < N; i++) {
  if (sec[i] !== "PROVES" || sb(i) === "Heading2") continue;
  const t = txt(i);
  if (sb(i) === "Heading3") { if (nomGrup(t) !== t) grups.push({ i, nn: nomGrup(t), nou: false }); continue; }
  if (esH(i) || esLista(i) || esMini(i) || t === "") continue;
  if (esNou(i)) {
    const ls = lineas(i);
    const est = ls.slice(2).filter(x => !RE_LAB.test(x)).find(x => RE_EST2.test(x)) || "";
    const idx = []; let j = i + 1;
    for (; j < N && sec[j] === "PROVES" && !esH(j) && !esMini(j) && txt(j) !== "" && !RE_PROVA.test(txt(j)) &&
           !esNou(j) && !RE_RECOMPTE.test(txt(j)); j++) idx.push(j);
    proves2.push({ i, fecha: ls[0], perf: ls[1], est, idx, fi: j });
    i = j - 1; continue;
  }
  if (RE_PROVA.test(t)) {
    const parts = lineas(i);
    let llista = -1;
    const q = i + 1;
    if (q < N && sec[q] === "PROVES" && !esH(q) && !esLista(q) && txt(q) !== "" &&
        !RE_PROVA.test(txt(q)) && (!esGrupCand(q) || /[\u000b\r\n]/.test(raw(q)))) llista = q;
    proves.push({ i, cap: parts[0], estat: parts.slice(1).join(" "), llista, items: llista >= 0 ? lineas(llista) : [] });
    if (llista >= 0) i++;
    continue;
  }
  if (esGrupCand(i)) grups.push({ i, nn: nomGrup(t), nou: true });
}

// ── SYNC 1b ──
const yaMini = new Set(), desvincular = new Set();
for (const gr of grups) if (gr.nou) {
  const s = gr.i + 1;
  if (s < N && txt(s) === "" && sec[s] === "PROVES" && !esH(s) && !esMini(s)) {
    yaMini.add(s); if (it[s].isListItem) desvincular.add(s);
  }
}
for (const pr of proves2) {
  pr.brk = it[pr.i].search("^l"); pr.brk.load("items");
  for (const j of pr.idx) if (it[j].isListItem) desvincular.add(j);
  const pv = pr.i - 1;
  if (pv >= 0 && txt(pv) === "" && esLista(pv) && it[pv].isListItem) desvincular.add(pv);
}
for (const j of desvincular) it[j].detachFromList();
const vacios = [];
for (let i = 0; i < N; i++) if (T[i] === "" && !esH(i) && sec[i] !== "CONSTANTS" && !/^SILICON/.test(sec[i])) vacios.push(i);
const fotosV = vacios.map(i => it[i].inlinePictures.load("items/width"));
if (proves2.length || desvincular.size || vacios.length) await context.sync();
const conImg = new Set(vacios.filter((i, k) => fotosV[k].items.length));

// ── 3) escrituras ──
const borrarP = new Set();
for (const gr of grups) {
  const p = it[gr.i];
  const nuevoTexto = gr.nn !== txt(gr.i) ? gr.nn : null;
  if (gr.nou) {
    H(p, 3, nuevoTexto);
    const s = gr.i + 1;
    if (s < N && txt(s) !== "") mini(p.insertParagraph("", "After"));
    else if (yaMini.has(s) && !conImg.has(s)) mini(it[s]);
  } else if (nuevoTexto) p.getRange("Content").insertText(nuevoTexto, "Replace");
}
const ESTAT = e => /^pend/i.test(e) ? "PENDENT DE PROGRAMAR" : e.toUpperCase();
const estadoP = (p, e) => {
  const x = p.insertParagraph(e, "After");
  x.styleBuiltIn = "Normal"; plano(x, 10.5);
  x.getRange("Content").styleBuiltIn = "SubtleEmphasis";
  x.font.italic = true; x.font.color = "#ED7D31";
  return x;
};
for (const pr of proves) {
  const p = it[pr.i];
  H(p, 4, pr.cap); p.leftIndent = 0;
  let ultimo = p;
  if (pr.estat) ultimo = estadoP(p, ESTAT(pr.estat));
  if (pr.llista >= 0 && pr.items.length) {
    const q = it[pr.llista];
    q.getRange("Content").insertText(pr.items[0], "Replace");
    vineta(q); ultimo = q;
    for (const item of pr.items.slice(1)) { ultimo = ultimo.insertParagraph(item, "After"); vineta(ultimo); }
  }
  const sig = (pr.llista >= 0 ? pr.llista : pr.i) + 1;
  if (sig < N && sec[sig] === "PROVES" && txt(sig) !== "" && !esMini(sig) && sb(sig) !== "Heading2") {
    if (pr.items.length) blanco(ultimo.insertParagraph("", "After"));
    else mini(ultimo.insertParagraph("", "After"));
  }
}

// ── traducción (DIC: ver "Traducción al catalán") ──
const DIC = {
  "FUNCION RENAL": "Funció Renal", "FUNCION HEPATICA": "Funció Hepàtica",
  "ESTUDIO GENETICO INDETERMINADO": "Estudi Genètic Indeterminat",
  "CALCIO": "Calci", "FOSFATO": "Fosfat", "MAGNESIO": "Magnesi",
  "PARATHORMONA (PTH INTACTA)": "Paratirina (PTH Intacta)", "TIROTROPINA (TSH)": "Tirotropina (TSH)",
  "TIROXINA LIBRE (T4 LIBRE)": "Tiroxina Lliure (T4 Lliure)", "NT-PROBNP": "NT-proBNP",
  "MICROALBUMINA. COCIENTE ALBUMINA/CREATININA": "Microalbúmina. Cocient albumina/creatinina",
  "MICROALBUMINA. COCIENT ALBUMINA/CREATININA": "Microalbúmina. Cocient albumina/creatinina",
  "CREATININA": "Creatinina", "UREA": "Urea", "IONOGRAMA": "Ionograma",
  "PROTEINA C REACTIVA": "Proteïna C Reactiva Sèrum",
};
const MOSTRA = [[/\s+en orina esporádica$/i, "Orina esporàdica"], [/\s+en orina de 24 ?h(oras)?$/i, "Orina 24 h"],
                [/\s+en orina$/i, "Orina esporàdica"], [/\s+séric[oa]s?$/i, "Sèrum"]];
const MESOS = { ENERO: "Gener", FEBRERO: "Febrer", MARZO: "Març", ABRIL: "Abril", MAYO: "Maig", JUNIO: "Juny",
                JULIO: "Juliol", AGOSTO: "Agost", SEPTIEMBRE: "Setembre", SETIEMBRE: "Setembre",
                OCTUBRE: "Octubre", NOVIEMBRE: "Novembre", DICIEMBRE: "Desembre" };
const RE_ES = /[áñ]|ión\b|(io|ico|ato|ido|ados?)\b|\b(y|libre|cociente|hierro|sangre|prueba|minutos|horas)\b/i;
const clauT = s => sinAcentos(s).replace(/\s+/g, " ").replace(/[.\s]+$/, "").trim();
const tradueix = s => {
  if (DIC[clauT(s)]) return DIC[clauT(s)];
  for (const [re, m] of MOSTRA) {
    const x = s.match(re); if (!x) continue;
    const n = DIC[clauT(s.slice(0, x.index))];
    return n ? (n.endsWith(" " + m) ? n : n + " " + m) : null;
  }
  return null;
};
const majus = t => { const m = t.match(/\s(Sèrum|Orina esporàdica|Orina 24 h)$/); return m ? t.slice(0, m.index).toUpperCase() + m[0] : t.toUpperCase(); };
const trPerf = perf => {
  const seg = perf.split(/\.\s+/).map(x => x.replace(/\.$/, "").trim()).filter(Boolean), out = [];
  for (let a = 0; a < seg.length;) {
    let b = seg.length, t = null;
    for (; b > a && !(t = tradueix(seg.slice(a, b).join(". "))); b--);
    if (!t) { if (RE_ES.test(seg[a])) return null; t = seg[a]; b = a + 1; }
    out.push(majus(t)); a = b;
  }
  return out.join(", ");
};
const ESTAT2 = e => /^(realiz|realitz)/i.test(e) ? "" : /^pend/i.test(e) ? "PENDENT DE PROGRAMAR"
  : /^programad/i.test(e) ? "PROGRAMADA" : /^result/i.test(e) ? "RESULTATS PARCIALS" : /^anul/i.test(e) ? "ANUL·LADA" : e.toUpperCase();
let nTrad = 0; const pend = [];
for (const pr of proves2) {
  const p = it[pr.i], h = pr.brk.items, items = pr.idx.map(txt);
  const esp = [pr.perf, ...items].some(x => RE_ES.test(x));
  const mm = pr.fecha.match(/^([A-Za-zÀ-ÿ]+) de (\d{4})$/), mes = mm && MESOS[sinAcentos(mm[1])];
  const perfCa = esp && h.length ? trPerf(pr.perf) : null;
  if (h.length) {
    const fin = () => p.getRange("Content").getRange("End");
    if (h.length > 1) h[1].expandTo(fin()).delete();
    if (perfCa && perfCa !== pr.perf) h[0].getRange("End").expandTo(fin()).insertText(perfCa, "Replace");
    h[0].insertText(" — ", "Replace");
    if (mes && mes !== mm[1]) p.search(mm[1], { matchCase: true }).getFirst().insertText(mes, "Replace");
  }
  H(p, 4); p.leftIndent = 0;
  const e = ESTAT2(pr.est);
  if (e) estadoP(p, e);
  const falta = perfCa || !RE_ES.test(pr.perf) ? [] : [pr.perf];
  pr.idx.forEach((j, k) => {
    const c = esp ? tradueix(items[k]) : null;
    if (c && c !== items[k]) it[j].getRange("Content").insertText(c, "Replace");
    else if (!c && esp && RE_ES.test(items[k])) falta.push(items[k]);
    vineta(it[j]);
  });
  if (esp) {
    if (falta.length) pend.push({ cab: (mes ? mes + " de " + mm[2] : pr.fecha) + " — " + (perfCa || pr.perf), falta });
    else nTrad++;
  }
  const pv = pr.i - 1;
  if (pv >= 0 && txt(pv) === "" && esLista(pv) && !yaMini.has(pv) && !conImg.has(pv)) blanco(it[pv]);
  if (pv >= 1 && esMini(pv) && txt(pv - 1) === "" && !esH(pv - 1)) borrarP.add(pv);
  const s = pr.fi;
  let nx = s; while (nx < N && txt(nx) === "" && !esH(nx)) nx++;
  const antesDeH2 = nx < N && sb(nx) === "Heading2";
  const antesDeH3 = nx >= N || /^Heading[23]$/.test(sb(nx));
  if (s < N && txt(s) === "" && !esH(s)) {
    if (!antesDeH2 && !conImg.has(s)) {
      if (pr.idx.length || antesDeH3) blanco(it[s]); else mini(it[s]);
      for (let k = s + 1; k < nx; k++) borrarP.add(k);
    }
  } else if (s < N && sec[s] === "PROVES") {
    const ult = pr.idx.length ? it[pr.idx[pr.idx.length - 1]] : p;
    if (pr.idx.length || antesDeH3) blanco(ult.insertParagraph("", "After")); else mini(ult.insertParagraph("", "After"));
  }
}

const RE_ENC = /^(Evolutiu|Evolució|EVOLUCIÓ|NOTA D['’]INGR|Nota d['’]ingr|GUARDIA)/;
const RE_ETQ = /^(AP|AC|SB|MA|AVUI|EF|OD|PLA|ECG|JC)\s*:/;
const etiquetas = [];
for (const b of notas) {
  const ic = b.s === "INTERCONSULTES";
  const dp = it[b.d];
  const prev = (b.cc >= 0 ? b.cc : b.d) - 1;
  if (!b.v1 && prev >= 0) {
    if (sb(prev) === "Heading2") (ic ? blanco : mini)(dp.insertParagraph("", "Before"));
    else if (txt(prev) !== "") blanco(dp.insertParagraph("", "Before"));
  }
  if (ic && b.prefijo) {
    H(dp.insertParagraph(b.prefijo, "Before"), 3);
    mini(dp.insertParagraph("", "Before"));
  }
  const quien = ic && b.cat && b.nombre ? b.cat + " – " + b.nombre : b.nombre;
  H(dp, ic ? 4 : 3, b.fecha + " (" + b.hora + ")" + (quien ? " → " + quien : ""));
  mini(dp.insertParagraph("", "After"));
  if (b.ini <= b.fin) {
    const r = it[b.ini].getRange("Whole").expandTo(it[b.fin].getRange("Whole"));
    r.font.name = "Calibri Light"; r.font.size = ic ? 11 : 10;
    if (!ic) {
      for (let j = b.ini; j <= b.fin; j++) {
        if (txt(j) === "") continue;
        if (RE_ENC.test(txt(j))) { it[j].font.bold = true; it[j].font.underline = "Single"; }
        break;
      }
      if (b.s !== "INFERMERIA") for (let j = b.ini; j <= b.fin; j++) {
        const m = txt(j).match(RE_ETQ);
        if (m) { const rc = it[j].search(m[0], { matchCase: true }); rc.load("items"); etiquetas.push(rc); }
      }
    }
    const s = b.fin + 1;
    if (s < N && txt(b.fin) !== "" && (sb(s) === "Heading3" || (ic && sb(s) === "Heading4")))
      blanco(it[b.fin].insertParagraph("", "After"));
  }
  if (b.sep >= 0 && (sb(b.sep + 1) === "Heading3" || (ic && sb(b.sep + 1) === "Heading4"))) blanco(it[b.sep]);
}
await context.sync(); // SYNC 2

for (const rc of etiquetas) if (rc.items.length) rc.items[0].font.bold = true;
let nCC = 0;
for (const b of notas) {
  for (const j of b.borrar) if (!conImg.has(j)) it[j].delete();
  if (b.cc >= 0) { it[b.cc].delete(); nCC++; }
}
for (const j of borrarP) if (!conImg.has(j)) it[j].delete();
await context.sync(); // SYNC 3

// ═══ FASE B: INTERCONSULTES ═══
const cB = { serveis: 0, pets: 0, resps: 0, cc: 0, afegides: 0, soltes: 0, descartades: 0, noReal: 0 };
let avisB = "";
try {
  const P2 = context.document.body.paragraphs;
  P2.load("items/text,items/style,items/styleBuiltIn");
  await context.sync(); // SYNC 4
  const q = P2.items, M = q.length;
  const rw = i => q[i].text || "";
  const tx = i => rw(i).replace(/[\u000b\r\n]+/g, " ").trim();
  const hN = i => { const m = /^Heading([1-5])$/.exec(q[i].styleBuiltIn); return m ? +m[1] : 0; };
  const esM = i => /MiniEspacio/i.test(q[i].style || "");
  const norm = x => sinAcentos(x).replace(/\s+/g, " ").trim();
  const clau = (d, h) => d.split("/").reverse().join("") + (h || "00:00").replace(":", "");
  let a = -1, z = M;
  for (let i = 0; i < M; i++) if (hN(i) === 2) {
    if (a < 0) { if (norm(tx(i)) === "INTERCONSULTES") a = i; }
    else { z = i; break; }
  }
  if (a >= 0) {
    const RE_REAL = /^(Realizadas|Realitzades)\s*\((\d+)\)$/i;
    const RE_NOREAL = /^No (realizadas|realitzades)\s*\((\d+)\)$/i;
    const RE_CAP = /^(Ninguna|Ninguno|Cap)\.?$/i;
    const RE_PET = /^([^:()\d]{3,70}?)\s*\((\d{2}\/\d{2}\/\d{4})\)$/;
    const RE_RH = /^(Respuestas|Respuesta|Respostes|Resposta)\s*:\s*(.*)$/i;
    const RE_APP = /^(\d{2}\/\d{2}\/\d{4})\s*\((\d{2}:\d{2})\)\s*[—–-]\s*(.+)$/;
    const RE_RF = /^(\d{2}\/\d{2}\/\d{4}) \((\d{2}:\d{2})\) → (?:(.+?) – )?(.+)$/;
    const RE_VEURE = /^(veure|ver)\s+(el\s+)?(curs cl[ií]nic|curso cl[ií]nico|cc)\.?$/i;
    const RE_SENSE = /^(sense|sin) resp/i;
    const NOMS = [[/^ANESTES/, "ANESTESIOLOGIA I REANIMACIÓ"],
                  [/^(RHB|REHABILITACIO|MEDICINA FISICA)/, "MEDICINA FÍSICA REHABILITACIÓ"],
                  [/^NUTRI/, "NUTRICIÓ I DIETÈTICA"],
                  [/^INFO? ACOD/, "INF ACOD I CONTROL SINTROM"]];
    const nomServei = t => { const k = norm(t); for (const [re, n] of NOMS) if (re.test(k)) return n; return t.trim().toUpperCase(); };
    const GEN = new Set(["MEDIC", "INTER", "HOSPI", "CLINI", "METGE", "GENER", "UNITA", "SERVE"]);
    const SIN = { REHAB: ["FISIO", "LOGOP"], NUTRI: ["DIETI"], ADDIC: ["PSIQU", "PSICO"], SALUT: ["PSIQU", "PSICO"] };
    const arrels = x => {
      const r = new Set();
      for (const w of norm(x).split(/[^A-Z]+/)) if (w.length >= 4 && !GEN.has(w.slice(0, 5))) {
        const k = w.slice(0, 5); r.add(k); (SIN[k] || []).forEach(y => r.add(y));
      }
      return r;
    };
    const cos11 = p => { p.styleBuiltIn = "Normal"; p.font.name = "Calibri Light"; p.font.size = 11; p.spaceBefore = 0; p.spaceAfter = 0; p.lineSpacing = 12; };

    // ── B1) bloques de la app ──
    const blocs = [];
    for (let i = a + 1; i < z; i++) {
      const m = RE_REAL.exec(tx(i)); if (!m) continue;
      let s = i - 1; while (s > a && tx(s) === "") s--;
      if (s <= a || hN(s) === 2 || RE_CAP.test(tx(s))) continue;
      blocs.push({ s, r: i, n: +m[2], nom: nomServei(tx(s)), pets: [], buits: [], sobren: [], noReal: null, ult: i });
    }
    for (let k = 0; k < blocs.length; k++) {
      const b = blocs[k], lim = k + 1 < blocs.length ? blocs[k + 1].s : z;
      let pet = null, enResp = false, fiBloc = lim;
      for (let x = b.s + 1; x < b.r; x++) if (tx(x) === "") b.buits.push(x);
      for (let i = b.r + 1; i < lim; i++) {
        if (hN(i)) { fiBloc = i; break; }
        const t = tx(i); let m;
        if (t === "") { b.buits.push(i); continue; }
        b.ult = i;
        if ((m = RE_NOREAL.exec(t))) { b.noReal = { i, n: +m[2], cap: -1 }; continue; }
        if (b.noReal && RE_CAP.test(t)) { b.noReal.cap = i; continue; }
        const l = rw(i).split(/[\u000b\r\n]+/).map(x => x.trim()).filter(Boolean);
        if ((m = RE_PET.exec(t)) && !RE_APP.test(l[0])) {
          pet = { i, data: m[2], nom: m[1].trim().toUpperCase(), text: [], rh: -1, sense: [], resps: [], noves: [], fi: i };
          b.pets.push(pet); enResp = false; continue;
        }
        if (!pet) { b.sobren.push(i); continue; }
        pet.fi = i;
        if ((m = RE_RH.exec(l[0] || ""))) { pet.rh = i; enResp = true; continue; }
        if (enResp && (m = RE_APP.exec(l[0]))) {
          const cos = l.slice(1);
          pet.resps.push({ i, data: m[1], hora: m[2], nom: m[3].trim().toUpperCase(), cos, extra: [], cc: null,
                           veure: !cos.length || (cos.length === 1 && RE_VEURE.test(cos[0])) });
          continue;
        }
        if (enResp && RE_SENSE.test(t)) pet.sense.push(i);
        else if (enResp && pet.resps.length) pet.resps[pet.resps.length - 1].extra.push(i);
        else pet.text.push(i);
      }
      const deText = new Set(b.pets.flatMap(p => [...p.text, ...p.resps.flatMap(r => r.extra)]));
      const nb = (x, d) => { for (x += d; x > b.s && x < fiBloc; x += d) if (tx(x) !== "") return x; return -1; };
      b.buits = b.buits.filter(x => !(deText.has(nb(x, -1)) && deText.has(nb(x, 1))));
    }
    const cabeceraBloc = new Set(blocs.flatMap(b => [b.s, b.r]));

    // ── B2) notas del curs clínic sin clasificar ──
    const prefijosA = new Set(notas.filter(b => b.prefijo).map(b => b.prefijo));
    const soltes = [];
    for (let j = a + 1; j < z; j++) {
      if (hN(j) !== 4) continue;
      const m = RE_RF.exec(tx(j)); if (!m) continue;
      let estat = "solta", passat4 = false;
      for (let k = j - 1; k > a; k--) {
        const h = hN(k); if (!h) continue;
        if (h === 4) { passat4 = true; continue; }
        if (h === 5) estat = "dins";
        else if (h === 3) {
          if (cabeceraBloc.has(k)) estat = "solta";
          else if (prefijosA.has(tx(k).toUpperCase()) && passat4) continue;
          else estat = "classif";
        }
        break;
      }
      let f = j + 1; while (f < z && !hN(f)) f++;
      let ini = j + 1; while (ini < f && esM(ini)) ini++;
      let fin = f - 1; while (fin >= ini && tx(fin) === "") fin--;
      const cosTxt = []; for (let x = ini; x <= fin; x++) if (tx(x) !== "") cosTxt.push(tx(x));
      soltes.push({ j, data: m[1], hora: m[2], cat: (m[3] || "").trim(), nom: m[4].trim().toUpperCase(),
                    ini, fin, fi: f - 1, estat, dest: "", veure: cosTxt.length === 1 && RE_VEURE.test(cosTxt[0]) });
    }

    // ── B3) destino de cada nota ──
    const totes = [];
    for (const b of blocs) for (const p of b.pets) for (const r of p.resps) totes.push({ b, p, r });
    for (const L of soltes) {
      if (L.estat !== "solta") continue;
      if (L.veure) { L.dest = "descartar"; continue; }
      const mismo = totes.find(o => !o.r.cc && o.r.data === L.data && norm(o.r.nom) === norm(L.nom));
      if (mismo) { mismo.r.cc = L; L.dest = "cc"; continue; }
      let b = (totes.find(o => norm(o.r.nom) === norm(L.nom)) || {}).b;
      if (!b && L.cat) { const ac = arrels(L.cat); b = blocs.find(x => [...arrels(x.nom)].some(y => ac.has(y))); }
      if (b && b.pets.length) {
        const cands = b.pets.filter(p => clau(p.data) <= clau(L.data));
        const p = cands.length ? cands.reduce((x, y) => clau(y.data) > clau(x.data) ? y : x) : b.pets[0];
        p.noves.push(L); L.dest = "afegida"; continue;
      }
      L.dest = "solta";
    }

    for (const L of soltes) if ((L.dest === "cc" || L.dest === "afegida") && L.ini <= L.fin)
      L.xml = q[L.ini].getRange("Whole").expandTo(q[L.fin].getRange("Whole")).getOoxml();
    const fotosB = [...new Set(blocs.flatMap(b => b.buits))].map(x => [x, q[x].inlinePictures.load("items/width")]);
    await context.sync(); // SYNC 5
    const conImgB = new Set(fotosB.filter(([, c]) => c.items.length).map(([x]) => x));

    // ── B4) escrituras ──
    const esborra = [];
    const posaCos = (despres, L) => {
      blanco(despres.insertParagraph("", "After"));
      if (L.xml) despres.insertParagraph("", "After").insertOoxml(L.xml.value, "Replace");
    };
    const posaAbans = (ancla, L) => {
      H(ancla.insertParagraph("", "Before"), 4, `${L.data} (${L.hora}) → ${L.nom}`);
      const m = ancla.insertParagraph("", "Before"); mini(m);
      posaCos(m, L);
    };
    for (const b of blocs) {
      cB.serveis++;
      const sp = q[b.s];
      H(sp, 3, `${b.nom} – Realitzades (${b.n})`);
      mini(sp.insertParagraph("", "After"));
      esborra.push(b.r, ...b.buits.filter(x => !conImgB.has(x)));
      if (b.noReal) {
        if (b.noReal.n === 0) { esborra.push(b.noReal.i); if (b.noReal.cap >= 0) esborra.push(b.noReal.cap); cB.noReal++; }
        else cos11(q[b.noReal.i]);
      }
      for (const p of b.pets) {
        cB.pets++;
        H(q[p.i], 4, `(${p.data}) → ${p.nom}`);
        mini(q[p.i].insertParagraph("", "After"));
        for (const j of p.text) cos11(q[j]);
        esborra.push(...p.sense);
        const hayResp = p.resps.length + p.noves.length > 0;
        const fiPet = Math.max(p.fi, ...p.resps.flatMap(r => [r.i, ...r.extra]));
        const anclaFi = fiPet + 1 < M ? q[fiPet + 1] : null;
        if (p.rh >= 0) {
          const x = q[p.rh];
          mini(x.insertParagraph("", "Before"));
          H(x, 5, "RESPOSTES:");
          const m = x.insertParagraph("", "After"); mini(m);
          if (!hayResp) cos11(m.insertParagraph("Sense respostes.", "After"));
        } else if (hayResp && anclaFi) {
          const ancla = p.resps.length ? q[p.resps[0].i] : anclaFi;
          mini(ancla.insertParagraph("", "Before"));
          H(ancla.insertParagraph("RESPOSTES:", "Before"), 5);
          mini(ancla.insertParagraph("", "Before"));
        }
        for (const r of p.resps) {
          cB.resps++;
          const x = q[r.i], L = r.cc;
          H(x, 4, `${r.data} (${L ? L.hora : r.hora}) → ${r.nom}`);
          const m = x.insertParagraph("", "After"); mini(m);
          if (L) { cB.cc++; posaCos(m, L); esborra.push(...r.extra); }
          else {
            let c = m;
            for (const linia of r.cos) { c = c.insertParagraph(linia, "After"); cos11(c); }
            for (const j of r.extra) cos11(q[j]);
            blanco((r.extra.length ? q[r.extra[r.extra.length - 1]] : c).insertParagraph("", "After"));
          }
        }
        p.noves.sort((x, y) => clau(y.data, y.hora).localeCompare(clau(x.data, x.hora)));
        for (const L of p.noves) {
          const seg = p.resps.find(r => clau(r.data, r.hora) < clau(L.data, L.hora));
          const ancla = seg ? q[seg.i] : anclaFi;
          if (ancla) { posaAbans(ancla, L); cB.afegides++; }
        }
      }
    }
    let prevSolta = null;
    for (const L of soltes) {
      if (L.dest === "cc" || L.dest === "afegida" || L.dest === "descartar") {
        for (let x = L.j; x <= L.fi; x++) esborra.push(x);
        if (L.dest === "descartar") cB.descartades++;
        continue;
      }
      if (L.cat) H(q[L.j], 4, `${L.data} (${L.hora}) → ${L.nom}`);
      if (L.dest === "solta") {
        cB.soltes++;
        if (!(prevSolta && prevSolta.fi + 1 === L.j && prevSolta.cat === L.cat)) {
          H(q[L.j].insertParagraph("CC " + (L.cat || "ALTRES"), "Before"), 3);
          mini(q[L.j].insertParagraph("", "Before"));
        }
        prevSolta = L;
      }
    }
    for (const i of new Set(esborra)) q[i].delete();
    await context.sync(); // SYNC 6
  }
} catch (e) {
  avisB = " | ⚠ INTERCONSULTES sin reorganizar (" + (e.message || e) + "): el resto sí está hecho";
}

// ── SYNC 7-8: TDC ──
let tdc = "actualízala a mano";
if (api15) try {
  const f = context.document.body.fields.getByTypes(["TOC"]); f.load("items");
  await context.sync();
  f.items.forEach(x => x.updateResult());
  await context.sync();
  tdc = f.items.length ? "actualizada" : "no hay";
} catch (e) { tdc = "actualízala a mano (" + (e.message || e) + ")"; }

// ── SYNC 9: verificación ──
const v = context.document.body.paragraphs;
v.load("items/text,items/styleBuiltIn,items/style,items/font/name");
const fotos1 = context.document.body.inlinePictures; fotos1.load("items/width");
await context.sync();
let s = "", ccRest = 0, viejas = 0, h3n = 0, h3mal = 0, v1Rest = 0, icResp = 0, provSucias = 0, largos = 0,
    icCrudo = 0, icCat = 0, veureRest = 0, grupSucio = 0, castRest = 0;
for (const p of v.items) {
  const rt = p.text || "", t = rt.replace(/[\u000b\r\n]+/g, " ").trim(), st = p.styleBuiltIn;
  if (st === "Heading2") { s = t.toUpperCase(); if (CORTOS[s]) largos++; }
  if (t === "CC") ccRest++;
  if (t !== "" && /Times New Roman|Arial/i.test(p.font.name || "") && s !== "CONSTANTS" && !/^SILICON/.test(s)) viejas++;
  if (st === "Heading3" && RE_V1.test(t)) v1Rest++;
  if (st === "Heading3" && /^\d/.test(t) && s !== "PROVES") { h3n++; if (!/^\d{2}\/\d{2}\/\d{4} \(\d{2}:\d{2}\) → \S/.test(t)) h3mal++; }
  if (s === "INTERCONSULTES") {
    if (st === "Heading4" && /^\d{2}\/\d{2}\/\d{4} \(/.test(t)) { icResp++; if (/ → .+ – /.test(t)) icCat++; }
    if (/^(No )?(Realizadas|Realitzades) \(\d+\)$|^Respuestas:|^Ninguna\.?$/i.test(t) && !/^Heading/.test(st)) icCrudo++;
    if (/^(veure|ver) (curs cl[ií]nic|cc)/i.test(t)) veureRest++;
  }
  if (s === "PROVES") {
    if (!/^Heading/.test(st) && (RE_PROVA.test(t) || (/\u000b/.test(rt) && RE_FNOU.test(rt.split(/\u000b/)[0].trim())))) provSucias++;
    if (!/^Heading/.test(st) && RE_RECOMPTE.test(t)) grupSucio++;
    if ((st === "Heading4" || st === "ListBullet" || /vi[ñn]etas/i.test(p.style || "")) && RE_ES.test(t)) castRest++;
  }
}
const imgPerdidas = fotos0.items.length - fotos1.items.length;
return "Notas nuevas: " + notas.filter(b => !b.v1).length + " | v1 migradas: " + notas.filter(b => b.v1).length +
  " | pruebas antiguo: " + proves.length + " | pruebas nuevo: " + proves2.length + " | traducidas: " + nTrad +
  " | grupos: " + grups.length + " | CC borradas: " + nCC + " | títulos acortados: " + nTit + " | TDC: " + tdc +
  " || IC servicios: " + cB.serveis + " | peticiones: " + cB.pets + " | respuestas: " + cB.resps +
  " | sustituidas CC: " + cB.cc + " | añadidas: " + cB.afegides + " | sueltas: " + cB.soltes +
  " | veure borradas: " + cB.descartades + " | No realizadas(0): " + cB.noReal + avisB +
  " || CONTROL → H3: " + h3n + " (mal " + h3mal + ") | IC resp: " + icResp + " (cat " + icCat + ") | IC sucio: " + icCrudo +
  " | veure sin sustituir: " + veureRest + " | v1 sin migrar: " + v1Rest + " | PROVES sucio: " + provSucias +
  " | grupos sucios: " + grupSucio + " | castellano: " + castRest + " | CC rest: " + ccRest +
  " | títulos largos: " + largos + " | Times/Arial: " + viejas +
  (imgPerdidas ? " | ⚠ IMÁGENES PERDIDAS: " + imgPerdidas : "") +
  (pend.length ? " || CASTELLANO PENDIENTE: " + JSON.stringify(pend) : "");
```

### Guía del código (para modificarlo sin romperlo)

- **Helpers.** `col(i)` devuelve "" cuando el párrafo mezcla colores (línea de cargo de Argos). `H(p, n, texto)` pone Título n con su formato: 3 = notas, grupos y servicios (12 negro); 4 = pruebas, peticiones y respuestas (11 azul); 5 = `RESPOSTES:` (11 azul). `blanco`, `mini` y `vineta` crean los separadores y las viñetas de la plantilla.
- **Notas (plan).** Una nota empieza en un párrafo con la fecha sola seguido de otro con la hora sola (en INTERCONSULTES la fecha puede llevar delante una etiqueta, `RE_FP`), o en un Título 3 antiguo `dd/mm/yyyy - hh:mm`. Se borran: la hora, un blanco antes del nombre (v1), el nombre, como mucho 2 líneas de autor (especialidad + cargo) y el blanco que las sigue. Un Título 3 antiguo sin nombre reconocible no se toca. En IC, la cabecera lleva la categoría (`CATEGORÍA – NOMBRE`) solo hasta la fase B, que la usa para asignar la nota y luego la quita.
- **Notas (escritura).** Antes de la cabecera: MiniEspacio si va justo tras el Título 2 (blanco en IC), blanco si lo anterior tiene texto. Etiqueta de IC → Título 3 propio + MiniEspacio. Cuerpo: Calibri Light 10 (11 en IC) sobre el rango entero, 1ª línea tipo "Evolutiu…" en negrita + subrayado y etiquetas AP/AC/… en negrita (búsqueda dentro del párrafo, se aplica tras el SYNC 2).
- **PROVES nuevo.** Las ediciones de la cabecera se hacen **de derecha a izquierda** (borrar desde el 2º salto → traducir los perfiles → 1er salto por " — " → mes) para que ningún rango se desplace antes de usarlo; el mes se busca con `search`.
- **Fase B.** La sección es el tramo entre el Título 2 INTERCONSULTES y el siguiente Título 2. Un bloque de la app acaba en el primer título (ahí empiezan las notas del curs clínic). Los blancos del pegado se borran salvo los que separan párrafos de un mismo texto (petición o respuesta). Cada nota en Título 4 con hora se clasifica mirando hacia arriba: bajo un Título 5 `RESPOSTES:` → ya colocada; bajo el Título 3 de un bloque de la app → suelta (se le busca destino); bajo otro Título 3 (p. ej. "CC CARDIOLOGIA" que puso Cris) → ya clasificada, no se toca. Las sueltas que se quedan pierden la categoría de la cabecera. `posaCos` inserta tras un párrafo el cuerpo copiado (OOXML) y un blanco; `posaAbans` inserta la respuesta completa (Título 4 + MiniEspacio + cuerpo) antes de un ancla.

## Traducción al catalán

**Automática (en la llamada principal).** `DIC` es el diccionario: la clave va SIN acentos y en MAYÚSCULAS (castellano, o catalán a medias como "COCIENT"), de la determinación entera o solo del nombre sin la muestra; el valor es el catalán en tipo título, como va en la viñeta. Si una prueba del formato nuevo llega en castellano (`RE_ES`: `á`, `ñ`, `-ión`, `-io`, `-ico`, "libre", "hierro"… nada de eso existe en los nombres catalanes de Argos), cada determinación se busca en `DIC`: primero entera y, si no está, el nombre sin la muestra (`sérico/a` → `Sèrum`, `en orina esporádica` / `en orina` → `Orina esporàdica`, `en orina de 24 h` → `Orina 24 h`). La cabecera se traduce perfil a perfil con el mismo diccionario (en MAYÚSCULAS salvo la muestra) y el mes (`Septiembre de 2026` → `Setembre de 2026`) siempre. Ejemplos ya comprobados:

- `Función renal. Estudio genético indeterminado. Calcio sérico` → `FUNCIÓ RENAL, ESTUDI GENÈTIC INDETERMINAT, CALCI Sèrum`
- `Microalbúmina. Cocient albúmina/creatinina en orina` → `MICROALBÚMINA. COCIENT ALBUMINA/CREATININA Orina esporàdica`
- Viñetas: Función renal → Funció Renal · Estudio genético indeterminado → Estudi Genètic Indeterminat · Calcio sérico → Calci Sèrum · Fosfato sérico → Fosfat Sèrum · Magnesio sérico → Magnesi Sèrum · Parathormona (PTH intacta) sérica → Paratirina (PTH Intacta) Sèrum · Tirotropina (TSH) sérica → Tirotropina (TSH) Sèrum · Tiroxina libre (T4 libre) sérica → Tiroxina Lliure (T4 Lliure) Sèrum · NT-proBNP sérico → NT-proBNP Sèrum · Microalbúmina. Cociente albúmina/creatinina en orina esporádica → Microalbúmina. Cocient albumina/creatinina Orina esporàdica · Creatinina/Urea/Ionograma en orina esporádica → Creatinina/Urea/Ionograma Orina esporàdica · Función hepática → Funció Hepàtica · Proteína C reactiva → Proteïna C Reactiva Sèrum.

**Manual (segunda llamada, solo si el log trae "CASTELLANO PENDIENTE").** Cada entrada da la cabecera actual (`cab`) y los fragmentos que faltan (`falta`).

1. Localiza la cabecera filtrando por `styleBuiltIn === "Heading4"` y `text === cab` (o `startsWith(fecha + " — ")`). **La TDC (TDC 3) tiene el mismo texto y aparece antes: si no filtras por Heading4 cambiarás el índice en vez del cuerpo.**
2. Reemplaza **solo el fragmento**: en la cabecera, `paragraph.search(fragmento, {matchCase:true})` → `getFirst().insertText(catalán, "Replace")` (la búsqueda admite como mucho 255 caracteres); en cada viñeta siguiente (hasta el blanco), `getRange("Content").insertText(catalán, "Replace")`.
3. Criterio: el de los ejemplos de arriba; usa como referencia los nombres catalanes que ya haya en el documento.
4. Actualiza la TDC (`body.fields.getByTypes(["TOC"])` → `updateResult()`).
5. Dile a Cris qué términos nuevos has traducido para **añadirlos a `DIC`** en esta skill: la próxima vez saldrán solos.

## Reglas

- **Viñetas de PROVES**: no basta con aplicar "Lista con viñetas", porque un documento exportado de Argos trae la versión de Word por defecto de ese estilo (viñeta pegada al margen, texto a 0,63 cm, sin Calibri Light), no la de la plantilla. Por eso el script (a) corrige la definición del estilo en el documento cuando la API lo permite (WordApi 1.5), también al relanzarla sobre un documento ya formateado, y (b) pone siempre en cada viñeta sangría izquierda 22,7 pt y francesa 11,35 pt, Calibri Light 10. Las listas numeradas de Argos se desvinculan antes (`detachFromList`).
- **Nunca leas los párrafos al chat para "inspeccionar".** Si hace falta saber algo, averígualo dentro del script y sácalo en el log.
- Solo se tocan: notas en sucio (fecha + hora sueltas), Títulos 3 del formato antiguo, PROVES pegadas de Argos, bloques de la app en INTERCONSULTES y los Título 2 con nombre largo. Lo ya formateado, CONSTANTS, SILICON y los párrafos con imágenes no se tocan.
- Lo único que se borra: párrafo de hora, nombre, líneas grises de autor, el blanco que las rodea, las barras "CC", la línea de laboratorio/estado dentro de la cabecera nueva, los blancos/MiniEspacios duplicados, y en INTERCONSULTES "Realizadas (n)", "No realizadas (0)", "Ninguna.", "Sin respuestas" y las notas del curs clínic que se han movido o que solo decían "veure curs clínic". Nada más. Un párrafo vacío con imagen nunca se borra.
- Reemplazar solo el fragmento que cambia (en la cabecera nueva, el trozo entre saltos de línea; en una viñeta, su contenido), nunca borrar y recrear el párrafo.
- No duplicar líneas en blanco: cada inserción comprueba antes si el vecino ya está vacío o es un Título 2.
- No añadir nada en encabezados ni pies. No redefinir estilos (salvo la corrección de "Lista con viñetas"): los da la plantilla.
- Si el documento no tiene el estilo MiniEspacio (exportación recién bajada de Argos, sin plantilla), el script lo imita con un Normal de 1 pt y 6 pt antes. Lo ideal es pegar siempre dentro de la plantilla de Cris.
- Para igualar los estilos de varios documentos a uno de referencia: copiar `word/styles.xml` del de referencia dentro del `.docx` de los demás (es un zip) y comprobar después que nº de párrafos, texto e imágenes no cambian. Sin dejar temporales en la carpeta de Cris.

## Reporte al usuario

Resume el log en 3-4 líneas: qué se ha formateado (notas nuevas, antiguas migradas, pruebas antiguo/nuevo, traducidas, CC borradas, títulos acortados), qué se ha hecho en INTERCONSULTES (servicios, peticiones, respuestas sustituidas por el curs clínic, añadidas, sueltas al final, "No realizadas (0)" borradas) y si la TDC se ha actualizado. Si algún contador de CONTROL es distinto de 0 (**mal**, **cat**, **IC sucio**, **veure sin sustituir**, **v1 sin migrar**, **PROVES sucio**, **grupos sucios**, **castellano**, **CC rest**, **títulos largos**, **Times/Arial**) o aparece un aviso ⚠ (INTERCONSULTES, IMÁGENES PERDIDAS), dilo explícitamente en vez de dar el trabajo por bueno. Recuerda siempre:

1. Si la TDC no se ha podido actualizar sola: clic derecho → Actualizar campo → Actualizar toda la tabla.
2. Revisar las notas que han quedado **al final de INTERCONSULTES** con Título 3 "CC …": son las que la skill no ha sabido asignar a ningún servicio; si alguna sí era de una interconsulta, se mueve a mano.
3. Si en PROVES o INTERCONSULTES quedan **imágenes** de Argos, eso lo transcribe Cris a mano con la misma estructura (Título 3 "SERVEI – Realitzades (n)", petición en Título 4, "RESPOSTES:" en Título 5).

## Si algo falla

- **"Espere a que finalice la llamada anterior"**: hay otra llamada en vuelo. Espera y reintenta **solo lecturas**; no relances la escritura, duplicaría cambios.
- **Errores de capacidad ("high demand")**: la conversación se ha llenado de datos del documento — vuelve a la Regla nº 1.
- **La herramienta no da `context`** (entornos antiguos): envuelve el código en `await Word.run(async (context) => { … });` y cambia el `return` final por `console.log(…)`.
- **Documentos gigantes (>150 notas)**: si hay un límite real de payload, trocea el bucle de escrituras **dentro de la misma llamada** (varios `await context.sync()`), nunca en varias llamadas.
- **Aviso ⚠ "INTERCONSULTES sin reorganizar"**: la fase B (que mueve bloques con OOXML) ha fallado, pero el resto ya está guardado. Se puede relanzar: lo ya hecho no se repite.
- **⚠ IMÁGENES PERDIDAS**: no debería pasar (los vacíos con imagen se respetan y lo que se mueve va con OOXML). Avisa a Cris para que lo deshaga con Ctrl+Z / historial de versiones antes de seguir.
- **Una nota del curs clínic asignada al servicio equivocado**: la asignación va por especialista + día, luego por especialista y luego por parecido entre cargo y servicio (tabla `SIN`). Muévela a mano y, si se repite, amplía `SIN` o `NOMS`.
- **Un estado de PROVES o una línea de lista mal clasificados**: los grupos se detectan como texto en MAYÚSCULAS sin cifras (el recuento aparte) seguido de una prueba; una lista de una sola determinación en mayúsculas (p. ej. solo "VSG") podría confundirse con un grupo. Corrígelo a mano; no merece otra pasada.
- **Prueba del formato nuevo no reconocida**: la 1ª línea debe ser `dd/mm/aaaa`, `dd/mm/aaaa (hh:mm)` o `Mes de aaaa` y el párrafo debe tener saltos de línea (Mayús+Intro). Si Argos cambia la línea de lugar amplía `RE_LAB`; si cambia el estado, `RE_EST2`.
- **"castellano" ≠ 0 o CASTELLANO PENDIENTE**: faltan términos en `DIC` → traducción manual (arriba) y añadirlos al diccionario.
- **"grupos sucios" ≠ 0**: hay un título de grupo con recuento que no se ha reconocido (p. ej. no está en MAYÚSCULAS o no tiene una prueba debajo); pásalo a Título 3 a mano.

## Historial

- **v1**: notas como Título 3 `dd/mm/yyyy - hh:mm` con el nombre y las líneas grises como párrafos (se migran solas).
- **v2 / `formatear-curs-clinic-argos`**: notas `dd/mm/yyyy (hh:mm) → NOMBRE`, INTERCONSULTES de la app, PROVES en formato antiguo.
- **`formatear-curs-clinic-argos-v3`**: + PROVES en formato nuevo, traducción al catalán en segunda llamada, viñetas con el estilo de la plantilla, MiniEspacio bajo cada grupo.
- **Esta versión (definitiva)**: traducción dentro de la misma llamada (diccionario `DIC` + muestra + mes); sin falsas alarmas de castellano ("Glucosa", "en orina"); grupos con recuento `(n realitzades, m no realitzada)`; blanco bajo un grupo nuevo → MiniEspacio; TDC actualizada al final; los párrafos vacíos con imagen ya no se borran y se avisa si se pierde alguna imagen; estado `Anulada` → `ANUL·LADA`; nota nueva pegada justo bajo el Título 2 → el MiniEspacio de debajo pasa a blanco de separación; código sin repeticiones (`H`).
