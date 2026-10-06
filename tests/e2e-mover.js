// Zoom y mover columnas y filas arrastrando el encabezado.
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
  assert.strictEqual(await page.$eval('#zoomIn', b => b.disabled), true, 'el zoom espera un archivo');
  await page.setInputFiles('#fileInput', path.join(__dirname, 'fixtures', 'precios.xlsx'));
  await page.waitForSelector('#grid td');
  const text = (r, c) => page.$eval(`td[data-r="${r}"][data-c="${c}"]`, el => el.textContent);
  const box = sel => page.$eval(sel, el => { const b = el.getBoundingClientRect(); return { x: b.x, y: b.y, w: b.width, h: b.height }; });
  const formula = async (r, c) => { await page.click(`td[data-r="${r}"][data-c="${c}"]`); return page.inputValue('#fbar'); };

  // ---- Zoom ----
  const w100 = (await box('th[data-hc="1"]')).w;
  await page.click('#zoomIn'); await page.click('#zoomIn');
  assert.strictEqual(await page.textContent('#zoomLbl'), '125%');
  const w125 = (await box('th[data-hc="1"]')).w;
  assert.ok(Math.abs(w125 / w100 - 1.25) < 0.03, `ancho ${w100} -> ${w125}`);
  const h125 = (await box('td[data-r="2"][data-c="1"]')).h;
  assert.ok(h125 > 25 && h125 < 28, 'alto de fila ' + h125);
  // la selección sigue alineada con la celda
  await page.click('td[data-r="3"][data-c="2"]');
  const a = await box('#actBox'), c = await box('td[data-r="3"][data-c="2"]');
  assert.ok(Math.abs(a.x - c.x) < 3 && Math.abs(a.y - c.y) < 3, JSON.stringify([a, c]));
  for (let i = 0; i < 7; i++) await page.click('#zoomOut');
  assert.strictEqual(await page.textContent('#zoomLbl'), '50%');
  assert.strictEqual(await page.$eval('#zoomOut', b => b.disabled), true);
  assert.ok((await box('th[data-hc="1"]')).w < w100 * 0.6);
  await page.click('#zoomLbl');
  assert.strictEqual(await page.textContent('#zoomLbl'), '100%');
  assert.ok(Math.abs((await box('th[data-hc="1"]')).w - w100) < 1);
  // cambiar el ancho de una columna con zoom guarda el ancho real (sin zoom)
  await page.click('#zoomIn'); await page.click('#zoomIn'); // 125%
  const rz = await box('th[data-hc="4"] .rz');
  await page.mouse.move(rz.x + rz.w / 2, rz.y + 5); await page.mouse.down();
  await page.mouse.move(rz.x + rz.w / 2 + 25, rz.y + 5, { steps: 4 }); await page.mouse.up();
  const saved = await page.evaluate(() => CoftechExcel.state.sheets[0].colW[4]);
  assert.ok(Math.abs(saved - (103 + 20)) <= 2, 'ancho guardado sin zoom: ' + saved);
  await page.keyboard.press('Control+z');
  await page.click('#zoomLbl');
  await page.screenshot({ path: (process.env.SHOTS || '/tmp') + '/5-zoom100.png' });

  const dragHeader = async (axis, from, to, opts = {}) => {
    const sel = i => (axis === 'col' ? `th[data-hc="${i}"]` : `th[data-hr="${i}"]`);
    await page.click(sel(from));
    if (opts.extra) await page.click(sel(opts.extra), { modifiers: ['Shift'] });
    const f = await box(sel(from)), t = await box(sel(to));
    const [fx, fy] = axis === 'col' ? [f.x + f.w / 2, f.y + f.h / 2] : [f.x + f.w / 2, f.y + f.h / 2];
    await page.mouse.move(fx, fy); await page.mouse.down();
    const [tx, ty] = axis === 'col'
      ? [opts.after ? t.x + t.w - 6 : t.x + 10, t.y + t.h / 2]
      : [t.x + t.w / 2, opts.after ? t.y + t.h - 3 : t.y + 3];
    await page.mouse.move(tx, ty, { steps: 8 });
    if (opts.shot) await page.screenshot({ path: (process.env.SHOTS || '/tmp') + '/' + opts.shot });
    await page.mouse.up();
  };

  // ---- Una combinada que quedaría partida impide el movimiento ----
  await dragHeader('col', 3, 1);
  assert.strictEqual(await text(0, 3), 'Con IVA', 'no se movió');
  assert.match(await page.textContent('#toasts'), /celdas combinadas/);
  // Se separa la combinada (A12:C12) y se sigue
  await page.click('td[data-r="11"][data-c="0"]');
  await page.click('#btnMerge');
  assert.strictEqual(await page.$eval('td[data-r="11"][data-c="0"]', el => el.colSpan), 1);

  // ---- Mover columnas: D (Con IVA) delante de B (Producto) ----
  assert.strictEqual(await text(0, 3), 'Con IVA');
  await dragHeader('col', 3, 1, { shot: '6-mover.png' });
  assert.strictEqual(await page.$eval('#moveLine', l => l.hidden), true);
  assert.deepStrictEqual([await text(0, 0), await text(0, 1), await text(0, 2), await text(0, 3), await text(0, 4)], ['Código', 'Con IVA', 'Producto', 'Precio', 'Categoría']);
  // la columna movida queda seleccionada en su nueva posición
  assert.strictEqual(await page.inputValue('#nameBox'), 'B1:B' + await page.evaluate(() => CoftechExcel.state.L.rows));
  assert.strictEqual(await text(1, 1), '$ 2.238,50');
  assert.strictEqual(await formula(1, 1), '=REDONDEAR(D2*1,21;2)', 'la fórmula sigue apuntando al precio');
  assert.strictEqual(await text(7, 3), '$ 9.080,00'); // el total sigue calculando sobre Precio
  assert.strictEqual(await formula(7, 3), '=SUMA(D2:D6)');
  // deshacer
  await page.keyboard.press('Control+z');
  assert.deepStrictEqual([await text(0, 1), await text(0, 3)], ['Producto', 'Con IVA']);
  assert.strictEqual(await formula(1, 3), '=REDONDEAR(C2*1,21;2)');

  // ---- Mover varias columnas: A:B después de D ----
  await dragHeader('col', 0, 3, { extra: 1, after: true });
  assert.deepStrictEqual([await text(0, 0), await text(0, 1), await text(0, 2), await text(0, 3)], ['Precio', 'Con IVA', 'Código', 'Producto']);
  assert.strictEqual(await formula(1, 1), '=REDONDEAR(A2*1,21;2)');
  await page.keyboard.press('Control+z');

  // Un clic en una columna ya seleccionada no la mueve
  await page.click('th[data-hc="2"]'); await page.click('th[data-hc="2"]');
  assert.strictEqual(await text(0, 2), 'Precio');
  assert.strictEqual(await page.inputValue('#nameBox').then(v => v.startsWith('C1:')), true);

  // ---- Mover filas: fila 2 (Coca Cola) debajo de la 4 ----
  await dragHeader('row', 1, 4, { shot: '7-mover-fila.png' });
  assert.deepStrictEqual([await text(1, 1), await text(2, 1), await text(3, 1), await text(4, 1)], ['Agua mineral 2L', 'Yerba mate 1kg', 'Coca Cola 1,5L', 'Azúcar 1kg']);
  assert.strictEqual(await formula(3, 3), '=REDONDEAR(C4*1,21;2)');
  assert.strictEqual(await text(3, 3), '$ 2.238,50');
  await page.keyboard.press('Control+z');
  assert.strictEqual(await text(1, 1), 'Coca Cola 1,5L');

  // La combinada vuelve a armarse y viaja con su fila
  await page.click('td[data-r="11"][data-c="0"]');
  await page.click('td[data-r="11"][data-c="2"]', { modifiers: ['Shift'] });
  await page.click('#btnMerge');
  assert.strictEqual(await page.$eval('td[data-r="11"][data-c="0"]', el => el.colSpan), 3);
  await dragHeader('row', 11, 2);
  assert.strictEqual(await page.$eval('td[data-r="2"][data-c="0"]', el => el.colSpan), 3);
  assert.strictEqual(await text(2, 0), 'Lista vigente');
  assert.strictEqual(await text(1, 1), 'Coca Cola 1,5L');

  assert.deepStrictEqual(errors, []);
  console.log('e2e-mover: ok');
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
