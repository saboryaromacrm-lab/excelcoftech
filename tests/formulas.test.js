const assert = require('assert');
const { HyperFormula } = require('hyperformula');
const es = require('hyperformula/commonjs/i18n/languages/esES');
const F = require('../src/formulas.js');
F.setLanguages(HyperFormula.getLanguage('enGB').functions, (es.default || es).functions);

const cases = [
  ['SUM(A1:A10)', 'SUMA(A1:A10)'],
  ['ROUND(C2*1.21,2)', 'REDONDEAR(C2*1,21;2)'],
  ['VLOOKUP(A2,Hoja2!A:C,3,FALSE)', 'BUSCARV(A2;Hoja2!A:C;3;FALSO)'],
  ['IF(A1>100,"Caro, mucho","Barato")', 'SI(A1>100;"Caro, mucho";"Barato")'],
  ['COUNTIF(B:B,">500")', 'CONTAR.SI(B:B;">500")'],
  ["SUMIF('Lista de precios'!A:A,\"Bebidas\",C:C)", "SUMAR.SI('Lista de precios'!A:A;\"Bebidas\";C:C)"],
  ['_xlfn.IFERROR(1/0,#N/A)', 'SI.ERROR(1/0;#N/A)'],
  ['A1/0+#DIV/0!', 'A1/0+#¡DIV/0!'],
];
for (const [en, esf] of cases) {
  assert.strictEqual(F.enToEs(en), esf, en);
  if (!en.startsWith('_xlfn')) assert.strictEqual(F.esToEn(esf), en, esf);
}
assert.strictEqual(F.shift('A1*$B$1+B$2+$C3', 2, 1), 'B3*$B$1+C$2+$C5');
assert.strictEqual(F.shift('SUMA(A1:A10)', 1, 0), 'SUMA(A2:A11)');
assert.strictEqual(F.shift('SUMA(A:A)', 5, 1), 'SUMA(B:B)');
assert.strictEqual(F.shift('SUMA(1:3)', 2, 0), 'SUMA(3:5)');
assert.strictEqual(F.shift('SUMA($1:$3)', 2, 0), 'SUMA($1:$3)');
assert.strictEqual(F.shift('A1*1,21', -1, 0), '#¡REF!*1,21');
assert.strictEqual(F.shift('Hoja2!A1+"A1"', 1, 0), 'Hoja2!A2+"A1"');
assert.strictEqual(F.shift('REDONDEAR(C2*1,21;2)', 3, 0), 'REDONDEAR(C5*1,21;2)');
console.log('formulas: ok');
