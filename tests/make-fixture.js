// Genera una lista de precios de ejemplo para las pruebas.
const ExcelJS = require('exceljs');
const path = require('path');
(async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Precios');
  ws.columns = [{ width: 12 }, { width: 30 }, { width: 14 }, { width: 14 }, { width: 14 }];
  ws.addRow(['Código', 'Producto', 'Precio', 'Con IVA', 'Categoría']);
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).eachCell(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF17725C' } }; });
  const items = [['A001', 'Coca Cola 1,5L', 1850, 'Bebidas'], ['A002', 'Agua mineral 2L', 950, 'Bebidas'], ['B001', 'Yerba mate 1kg', 4200, 'Almacén'], ['B002', 'Azúcar 1kg', 1300, 'Almacén'], ['C001', 'Lavandina 1L', 780, 'Limpieza']];
  items.forEach((it, i) => {
    const r = i + 2;
    ws.addRow([it[0], it[1], it[2], { formula: `ROUND(C${r}*1.21,2)` }, it[3]]);
    ws.getCell(`C${r}`).numFmt = '"$" #,##0.00';
    ws.getCell(`D${r}`).numFmt = '"$" #,##0.00';
  });
  ws.getCell('B8').value = 'Total';
  ws.getCell('C8').value = { formula: 'SUM(C2:C6)' };
  ws.getCell('C8').numFmt = '"$" #,##0.00';
  ws.getCell('B9').value = 'Bebidas';
  ws.getCell('C9').value = { formula: 'SUMIF(E2:E6,"Bebidas",C2:C6)' };
  ws.getCell('B10').value = 'Busca B001';
  ws.getCell('C10').value = { formula: "VLOOKUP(\"B001\",Lista!A:B,2,FALSE)" };
  ws.mergeCells('A12:C12');
  ws.getCell('A12').value = 'Lista vigente';
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  const l = wb.addWorksheet('Lista');
  l.addRow(['B001', 'Encontrado']);
  await wb.xlsx.writeFile(path.join(__dirname, 'fixtures', 'precios.xlsx'));
  console.log('ok');
})();
