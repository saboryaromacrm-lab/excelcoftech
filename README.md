# Coftech Excel

App web para abrir, editar y guardar planillas (listas de precios y planillas simples) desde el navegador de la PC. Toda la interfaz está en español.

## Cómo usarla

- **Con internet:** abrí https://saboryaromacrm-lab.github.io/excelcoftech/ en el navegador. Se actualiza sola con cada cambio que se sube al repositorio.
- **Sin internet:** descargá https://saboryaromacrm-lab.github.io/excelcoftech/coftech-excel.html (o `dist/coftech-excel.html` de este repositorio) y abrilo con doble clic. Es un único archivo que trae todo adentro y no necesita conexión.
- Abrí un archivo con el botón **Abrir** o arrastrándolo a la ventana: `.xlsx`, `.xls`, `.xlsm` (las macros no se ejecutan), `.csv`, `.tsv` y `.ods`.
- **Guardar como** descarga una copia nueva en tu carpeta de Descargas:
  - `.xlsx`: datos, fórmulas y formato. Es el formato recomendado.
  - `.ods`: datos y fórmulas, sin colores ni bordes.
  - `.xls`: datos y resultados de las fórmulas, sin colores ni bordes. La librería gratuita no puede escribir las fórmulas en el formato viejo de Excel.
  - `.csv` y `.tsv`: solo los valores de la hoja activa, tal como se ven. El CSV usa `;` como separador, igual que Excel en Argentina.
- **Imprimir / PDF** usa la impresión del navegador. Para obtener un PDF, elegí "Guardar como PDF".

### Fórmulas

Se escriben en español, con `;` entre argumentos y `,` para los decimales. Por ejemplo: `=SUMA(A1:A10)`, `=SI(A1>100;"Caro";"Barato")`, `=BUSCARV(A2;Hoja2!A:C;3;FALSO)`, `=REDONDEAR(C2*1,21;2)`.

Funcionan también `PROMEDIO`, `MIN`, `MAX`, `CONTAR`, `CONTAR.SI`, `SUMAR.SI` y el resto de las funciones habituales de Excel. Las fórmulas de los archivos se muestran traducidas al español y al guardar se vuelven a escribir en el formato de Excel.

### Atajos de teclado

| Atajo | Acción |
|---|---|
| Ctrl+C / Ctrl+X / Ctrl+V | Copiar / cortar / pegar (también desde Excel) |
| Ctrl+Z / Ctrl+Y | Deshacer / rehacer |
| Ctrl+B | Buscar y reemplazar |
| Ctrl+Alt+N | Negrita. El navegador reserva Ctrl+N para abrir una ventana nueva. |
| Ctrl+K / Ctrl+S | Cursiva / subrayado |
| F2 | Editar la celda |
| Supr | Borrar el contenido |
| Enter / Tab / flechas | Moverse entre celdas |
| Ctrl+flechas | Saltar al borde de los datos |

## Desarrollo

```bash
npm install
npm run build      # genera dist/coftech-excel.html
npm test           # pruebas de fórmulas y formatos
npm run test:e2e   # pruebas en Chromium (requiere Playwright)
```

El código está en `src/`:

- `formulas.js`: traducción inglés/español y ajuste de referencias.
- `format.js`: formatos de número argentinos.
- `io.js`: abrir y guardar archivos.
- `app.js`: la interfaz.

Librerías: [SheetJS](https://sheetjs.com) para leer y escribir los formatos, [ExcelJS](https://github.com/exceljs/exceljs) para el formato de los `.xlsx` y [HyperFormula](https://hyperformula.handsontable.com) (licencia GPLv3) para calcular las fórmulas. Por usar HyperFormula, el proyecto queda bajo GPLv3.
