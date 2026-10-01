// Ordenar, filtrar, buscar/reemplazar, pegar, hojas, combinar e impresión.
const { chromium } = require('playwright');
const path = require('path');
const assert = require('assert');

(async () => {
  const browser = await chromium.launch({ executablePath: process.env.CHROME || undefined });
  const page = await browser.newPage({ viewport: { width: 1366, height: 800 } });
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  await page.goto('file://' + path.join(__dirname, '..', 'dist', 'coftech-excel.html'));
  await page.setInputFiles('#fileInput', path.join(__dirname, 'fixtures', 'precios.xlsx'));
  await page.waitForSelector('#grid td');
  const text = (r, c) => page.$eval(`td[data-r="${r}"][data-c="${c}"]`, el => el.textContent);
  const cell = (r, c) => page.click(`td[data-r="${r}"][data-c="${c}"]`);

  // Ordenar por precio (columna C) de mayor a menor; el encabezado queda arriba
  await cell(2, 2);
  await page.click('#btnSortDesc');
  assert.strictEqual(await text(0, 2), 'Precio');
  assert.strictEqual(await text(1, 1), 'Yerba mate 1kg');
  assert.strictEqual(await text(1, 3), '$ 5.082,00'); // la fórmula siguió a su fila
  await cell(1, 3);
  assert.strictEqual(await page.inputValue('#fbar'), '=REDONDEAR(C2*1,21;2)');
  await page.click('#btnSortAsc');
  assert.strictEqual(await text(1, 1), 'Lavandina 1L');

  // Filtro por categoría
  await cell(1, 4);
  await page.click('#btnFilter');
  await page.click('.fbtn[data-fc="4"]');
  await page.click('#filterMenu label:has-text("(Seleccionar todo)") input');
  await page.click('#filterMenu label:has-text("Bebidas") input');
  await page.click('#filterMenu button.primary');
  const visible = await page.$$eval('#grid tbody tr', trs => trs.slice(0, 8).filter(t => t.style.display !== 'none').length);
  assert.strictEqual(visible, 5); // encabezado, 2 bebidas y las 2 filas fuera de la tabla
  await page.click('#btnFilter');
  const visible2 = await page.$$eval('#grid tbody tr', trs => trs.slice(0, 8).filter(t => t.style.display !== 'none').length);
  assert.strictEqual(visible2, 8);

  // Buscar y reemplazar
  await page.keyboard.press('Control+b');
  await page.fill('#findText', '1kg');
  await page.fill('#replaceText', '1 kg');
  await page.click('#btnReplaceAll');
  assert.match(await page.textContent('#findMsg'), /2 reemplazos/);
  await page.click('#findClose');

  // Pegar datos copiados desde Excel (texto con tabulaciones)
  await cell(14, 0);
  await page.evaluate(() => {
    const dt = new DataTransfer();
    dt.setData('text/plain', 'D001\tFideos 500g\t$ 1.200,50\r\nD002\tArroz\t990\r\n');
    document.getElementById('kbd').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true }));
  });
  assert.strictEqual(await text(14, 2), '$ 1.200,50');
  assert.strictEqual(await text(15, 1), 'Arroz');

  // Copiar y pegar dentro de la app ajusta referencias
  await cell(1, 3);
  await page.evaluate(() => {
    const dt = new DataTransfer();
    document.getElementById('kbd').dispatchEvent(new ClipboardEvent('copy', { clipboardData: dt, bubbles: true }));
    window.__clip = dt.getData('text/plain');
  });
  await cell(14, 3);
  await page.evaluate(() => {
    const dt = new DataTransfer();
    dt.setData('text/plain', window.__clip);
    document.getElementById('kbd').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true }));
  });
  await cell(14, 3);
  assert.strictEqual(await page.inputValue('#fbar'), '=REDONDEAR(C15*1,21;2)');
  assert.strictEqual(await text(14, 3), '$ 1.452,61');

  // Combinar
  await cell(17, 0);
  await page.click(`td[data-r="17"][data-c="2"]`, { modifiers: ['Shift'] });
  await page.click('#btnMerge');
  assert.strictEqual(await page.$eval('td[data-r="17"][data-c="0"]', el => el.colSpan), 3);
  await page.click('#btnMerge');
  assert.strictEqual(await page.$eval('td[data-r="17"][data-c="0"]', el => el.colSpan), 1);

  // Hojas: agregar, renombrar (y las fórmulas que la usan se actualizan), eliminar
  await page.dblclick('.tab[data-idx="1"]');
  await page.fill('#renameInput', 'Catálogo');
  await page.keyboard.press('Enter');
  await page.click('.tab[data-idx="0"]');
  await cell(9, 2);
  assert.strictEqual(await page.inputValue('#fbar'), '=BUSCARV("B001";Catálogo!A:B;2;FALSO)');
  await page.click('#btnAddSheet');
  assert.strictEqual(await page.$$eval('.tab', t => t.length), 3);
  await page.click('.tab[data-idx="2"]', { button: 'right' });
  await page.click('#ctxMenu button:has-text("Eliminar hoja")');
  await page.click('.modal button.primary');
  assert.strictEqual(await page.$$eval('.tab', t => t.length), 2);
  await page.click('.tab[data-idx="0"]');
  assert.strictEqual(await text(9, 2), 'Encontrado');

  // Escribir texto con acentos y formato porcentaje
  await page.click('.tab[data-idx="0"]');
  await cell(19, 0);
  await page.keyboard.type('Aumento');
  await page.keyboard.press('Tab');
  await page.keyboard.type('15%');
  await page.keyboard.press('Enter');
  assert.strictEqual(await text(19, 1), '15%');
  await cell(19, 2);
  await page.keyboard.type('=C2*(1+B20)');
  await page.keyboard.press('Enter');
  assert.strictEqual(await text(19, 2), '897');

  // Inmovilizar primera columna
  await page.click('#btnFreeze');
  await page.click('[data-freeze="both"]');
  assert.ok(await page.$eval('td[data-r="3"][data-c="0"]', el => el.classList.contains('fc')));

  // Impresión: tabla limpia sin la cinta
  await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
  await page.emulateMedia({ media: 'print' });
  assert.ok(await page.$eval('#print table', t => t.rows.length > 10));
  assert.strictEqual(await page.$eval('#app', el => getComputedStyle(el).display), 'none');
  await page.screenshot({ path: (process.env.SHOTS || '/tmp') + '/3-impresion.png' });
  await page.emulateMedia({ media: 'screen' });

  assert.deepStrictEqual(errors, []);
  console.log('e2e-datos: ok');
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
