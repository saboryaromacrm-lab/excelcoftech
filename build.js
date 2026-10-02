// Arma Coftech Excel en un único archivo HTML autocontenido (funciona sin internet).
// Uso: node build.js            -> dist/coftech-excel.html
//      node build.js --fragment <ruta>   (además genera el contenido sin <html>/<head> para publicar)
'use strict';
const fs = require('fs');
const path = require('path');

const root = __dirname;
const read = p => fs.readFileSync(path.join(root, p), 'utf8');
// Evita que un "</script" dentro de una librería cierre la etiqueta antes de tiempo.
const safe = js => js.replace(/<\/script/gi, '<\\/script').replace(/<!--/g, '<\\!--');

const libs = [
  'node_modules/xlsx/dist/xlsx.full.min.js',
  'node_modules/exceljs/dist/exceljs.min.js',
  'node_modules/hyperformula/dist/hyperformula.full.min.js',
  'node_modules/hyperformula/dist/languages/esES.js'
].map(p => '/* ' + path.basename(p) + ' */\n' + read(p)).join('\n;\n');

const app = ['src/formulas.js', 'src/format.js', 'src/io.js', 'src/app.js'].map(read).join('\n;\n');
const css = read('src/styles.css');

const body = read('src/index.html')
  .replace('/*__CSS__*/', () => css)
  .replace('/*__LIBS__*/', () => safe(libs))
  .replace('/*__APP__*/', () => safe(app));

const full = '<!doctype html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n' +
  body.replace(/\n<div id="app">/, '\n</head>\n<body>\n<div id="app">') + '\n</body>\n</html>\n';

fs.mkdirSync(path.join(root, 'dist'), { recursive: true });
const out = path.join(root, 'dist', 'coftech-excel.html');
fs.writeFileSync(out, full);
console.log('Generado', path.relative(root, out), (full.length / 1024 / 1024).toFixed(2) + ' MB');

// --site: carpeta lista para publicar (Vercel): la app como página principal y como archivo descargable.
if (process.argv.includes('--site')) {
  const site = path.join(root, '_site');
  fs.mkdirSync(site, { recursive: true });
  fs.writeFileSync(path.join(site, 'index.html'), full);
  fs.writeFileSync(path.join(site, 'coftech-excel.html'), full);
  console.log('Generado _site/');
}

const fi = process.argv.indexOf('--fragment');
if (fi > 0 && process.argv[fi + 1]) {
  fs.writeFileSync(process.argv[fi + 1], body);
  console.log('Generado', process.argv[fi + 1]);
}
