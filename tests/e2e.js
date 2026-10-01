// Prueba de punta a punta en Chromium: abrir, editar, fórmulas y guardar.
const { chromium } = require('playwright');
const path = require('path');
const ExcelJS = require('exceljs');
const assert = require('assert');

const shot = process.env.SHOTS;
(async () => {
  const browser = await chromium.launch({ executablePath: process.env.CHROME || undefined });
  const page = await browser.newPage({ viewport: { width: 1366, height: 800 }, acceptDownloads: true });
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  await page.goto('file://' + path.join(__dirname, '..', 'dist', 'coftech-excel.html'));
  if (shot) await page.screenshot({ path: shot + '/1-inicio.png' });
  await page.setInputFiles('#fileInput', path.join(__dirname, 'fixtures', 'precios.xlsx'));
  await page.waitForSelector('#grid td');
  const text = async (r, c) => page.$eval(`td[data-r="${r}"][data-c="${c}"]`, el => el.textContent);
  assert.strictEqual(await text(1, 2), '$ 1.850,00');
  assert.strictEqual(await text(1, 3), '$ 2.238,50');
  assert.strictEqual(await text(7, 2), '$ 9.080,00');
  assert.strictEqual(await text(8, 2), '2800');
  assert.strictEqual(await text(9, 2), 'Encontrado');
  // Fórmula traducida en la barra
  await page.click('td[data-r="1"][data-c="3"]');
  assert.strictEqual(await page.inputValue('#fbar'), '=REDONDEAR(C2*1,21;2)');
  // Editar precio y ver recálculo
  await page.click('td[data-r="1"][data-c="2"]');
  await page.keyboard.type('2000');
  await page.keyboard.press('Enter');
  assert.strictEqual(await text(1, 3), '$ 2.420,00');
  assert.strictEqual(await text(7, 2), '$ 9.230,00');
  // Fórmula nueva en español
  await page.click('td[data-r="13"][data-c="2"]');
  await page.keyboard.type('=PROMEDIO(C2:C6)');
  await page.keyboard.press('Enter');
  assert.strictEqual(await text(13, 2), '1846');
  await page.click('td[data-r="14"][data-c="2"]');
  await page.keyboard.type('=C2/0');
  await page.keyboard.press('Enter');
  assert.strictEqual(await text(14, 2), '#¡DIV/0!');
  // Deshacer
  await page.keyboard.press('Control+z');
  assert.strictEqual(await text(14, 2), '');
  // Negrita con Ctrl+Alt+N
  await page.click('td[data-r="13"][data-c="2"]');
  await page.keyboard.press('Control+Alt+n');
  assert.ok((await page.$eval('td[data-r="13"][data-c="2"]', el => el.style.fontWeight)) === '700');
  // Autocompletar: 1, 2 -> 3, 4
  await page.click('td[data-r="20"][data-c="0"]'); await page.keyboard.type('1'); await page.keyboard.press('Enter');
  await page.keyboard.type('2'); await page.keyboard.press('Enter');
  await page.click('td[data-r="20"][data-c="0"]');
  await page.click('td[data-r="21"][data-c="0"]', { modifiers: ['Shift'] });
  const fh = await page.$('#fillHandle');
  const b = await fh.boundingBox();
  const target = await (await page.$('td[data-r="23"][data-c="0"]')).boundingBox();
  await page.mouse.move(b.x + 3, b.y + 3); await page.mouse.down();
  await page.mouse.move(target.x + 10, target.y + 8, { steps: 5 }); await page.mouse.up();
  assert.strictEqual(await text(22, 0), '3');
  assert.strictEqual(await text(23, 0), '4');
  // Autocompletar fórmula hacia abajo (D6 -> D7 ajusta referencias)
  await page.click('td[data-r="5"][data-c="3"]');
  const fh2 = await (await page.$('#fillHandle')).boundingBox();
  const t2 = await (await page.$('td[data-r="6"][data-c="3"]')).boundingBox();
  await page.mouse.move(fh2.x + 3, fh2.y + 3); await page.mouse.down();
  await page.mouse.move(t2.x + 10, t2.y + 8, { steps: 5 }); await page.mouse.up();
  await page.click('td[data-r="6"][data-c="3"]');
  assert.strictEqual(await page.inputValue('#fbar'), '=REDONDEAR(C7*1,21;2)');
  // Insertar fila arriba de la 3 y ver que la fórmula del total se ajusta
  await page.click('th[data-hr="2"]');
  await page.click('#btnInsert'); await page.click('[data-cmd="insertRows"]');
  await page.click('td[data-r="8"][data-c="2"]');
  assert.strictEqual(await page.inputValue('#fbar'), '=SUMA(C2:C7)');
  if (shot) await page.screenshot({ path: shot + '/2-editado.png' });
  // Guardar xlsx y releer
  await page.click('#btnSave');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('[data-save="xlsx"]')]);
  const out = path.join(require('os').tmpdir(), 'coftech-out.xlsx');
  await dl.saveAs(out);
  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(out);
  const ws = wb.getWorksheet('Precios');
  assert.strictEqual(ws.getCell('D2').value.formula, 'ROUND(C2*1.21,2)');
  assert.strictEqual(ws.getCell('C9').value.formula, 'SUM(C2:C7)');
  assert.ok(ws.getCell('A1').font.bold);
  assert.strictEqual(ws.getCell('A1').fill.fgColor.argb, 'FF17725C');
  assert.strictEqual(ws.getCell('C2').numFmt, '"$" #,##0.00');
  assert.ok(ws.model.merges.includes('A13:C13'), JSON.stringify(ws.model.merges));
  // Otros formatos
  for (const f of ['xls', 'ods', 'csv', 'tsv']) {
    await page.click('#btnSave');
    const [d] = await Promise.all([page.waitForEvent('download'), page.click(`[data-save="${f}"]`)]);
    assert.ok(d.suggestedFilename().endsWith('.' + f));
    const p = path.join(require('os').tmpdir(), 'coftech-out.' + f);
    await d.saveAs(p);
  }
  // Reabrir el .xls y .ods guardados
  for (const f of ['xls', 'ods', 'csv']) {
    await page.evaluate(() => { window.CoftechExcel.state.dirty = false; });
    await page.setInputFiles('#fileInput', path.join(require('os').tmpdir(), 'coftech-out.' + f));
    await page.waitForFunction(n => document.getElementById('fileName').textContent === n, 'coftech-out.' + f);
    // .ods y .csv no conservan el formato moneda: se verifica el valor y la fórmula.
    assert.ok(['$ 2.420,00', '2420'].includes(await text(1, 3)), f);
    await page.click('td[data-r="1"][data-c="3"]');
    if (f === 'ods') assert.strictEqual(await page.inputValue('#fbar'), '=REDONDEAR(C2*1,21;2)', f);
  }
  assert.deepStrictEqual(errors, []);
  console.log('e2e: ok');
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
