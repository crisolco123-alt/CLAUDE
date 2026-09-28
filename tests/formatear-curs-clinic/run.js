'use strict';
// Uso: node run.js <SKILL.md> [--api14] [--sin-campos] [--sin-plantilla] [--incremental] [--volcar fichero]
// Ejecuta el script de la skill sobre el documento de prueba, lo vuelve a ejecutar sobre el
// resultado (debe quedar igual) y muestra el log y el documento final.
const fs = require('fs');
const { Doc, ejecutar, volcar } = require('./mock-word');
const { documento, pegadoNuevo } = require('./fixture');

const extraer = ruta => {
  const md = fs.readFileSync(ruta, 'utf8');
  const m = md.match(/```javascript\n([\s\S]*?)\n```/);
  if (!m) throw new Error('No hay bloque ```javascript en ' + ruta);
  // la skill antigua va envuelta en Word.run: se desenvuelve
  const w = m[1].match(/^await Word\.run\(async \(context\) => \{\n([\s\S]*)\n\}\);\s*$/);
  return w ? w[1] : m[1];
};

(async () => {
  const [ruta, ...flags] = process.argv.slice(2);
  const op = { api15: !flags.includes('--api14'), sinCampos: flags.includes('--sin-campos') };
  const codigo = extraer(ruta);
  let specs = documento();
  if (flags.includes('--sin-plantilla')) specs = specs.filter(x => x.style !== 'MiniEspacio');
  const doc = new Doc(specs);
  if (flags.includes('--sin-plantilla')) doc.estilos = {};
  const r1 = await ejecutar(codigo, doc, op);
  const v1 = volcar(doc);
  const r2 = await ejecutar(codigo, doc, op);
  const v2 = volcar(doc);
  console.log('── LOG 1ª pasada (' + r1.syncs + ' syncs) ──\n' + r1.ret.split(' || ').join('\n'));
  console.log('── LOG 2ª pasada ──\n' + r2.ret.split(' || ').join('\n'));
  console.log('── TDC actualizada: ' + doc.toc.actualizado + ' veces ──');
  console.log('── ESTILOS ──\n' + JSON.stringify(doc.estilos));
  console.log('── DOCUMENTO ──\n' + v1);
  if (v1 !== v2) {
    console.log('\n✗ LA 2ª PASADA CAMBIA EL DOCUMENTO:');
    const a = v1.split('\n'), b = v2.split('\n');
    for (let i = 0; i < Math.max(a.length, b.length); i++) if (a[i] !== b[i]) console.log(i + '\n- ' + a[i] + '\n+ ' + b[i]);
    process.exitCode = 1;
  } else console.log('\n✓ 2ª pasada idempotente');
  if (flags.includes('--incremental')) {
    // uso diario: se pega lo nuevo dentro del documento ya formateado y se relanza
    const { Nodo } = require('./mock-word');
    const nuevo = pegadoNuevo();
    const pos = t => doc.nodos.findIndex(n => n.style === 'Título 2' && n.text === t);
    doc.nodos.splice(pos('INTERCONSULTES'), 0, ...nuevo.finProves.map(x => new Nodo(x)));
    doc.nodos.splice(pos('MEDICINA INTERNA') + 1, 0, ...nuevo.inicioMI.map(x => new Nodo(x)));
    const r3 = await ejecutar(codigo, doc, op);
    const v3 = volcar(doc);
    const r4 = await ejecutar(codigo, doc, op);
    console.log('\n── LOG pasada incremental ──\n' + r3.ret.split(' || ').join('\n'));
    console.log('── DOCUMENTO tras el pegado nuevo ──\n' + v3);
    if (v3 !== volcar(doc)) { console.log('✗ la pasada siguiente al pegado cambia el documento'); process.exitCode = 1; }
    else console.log('✓ pasada siguiente al pegado idempotente');
  }
  if (flags.includes('--volcar')) fs.writeFileSync(flags[flags.indexOf('--volcar') + 1], v1 + '\n');
})().catch(e => { console.error(e); process.exit(2); });
