// API Serverless para Vercel: /api/productos.js
// Lee un archivo Excel incluido en el proyecto y devuelve JSON
// Coloca tu archivo .xls o .xlsx en la carpeta raíz del proyecto

export default function handler(req, res) {
  const fs = require('fs');
  const path = require('path');
  const xlsx = require('xlsx');

  // Buscar primer archivo .xls/.xlsx en el directorio raíz
  const root = process.cwd();
  const files = fs.readdirSync(root);
  const excelFile = files.find(f => /(xls|xlsx)$/i.test(f));
  if (!excelFile) {
    res.status(400).json({ error: 'No se encontró archivo .xls/.xlsx en la raíz del proyecto.' });
    return;
  }

  const filePath = path.join(root, excelFile);
  try {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    const startRow = 9; // fila 10 Excel -> index 9
    const productos = [];

    for (let r = startRow; r <= range.e.r; r++) {
      const code = worksheet[xlsx.utils.encode_cell({ c: 0, r })];
      if (code && code.v) {
        const desc = worksheet[xlsx.utils.encode_cell({ c: 2, r })];
        const rub = worksheet[xlsx.utils.encode_cell({ c: 5, r })];
        productos.push({
          codigo: String(code.v).trim(),
          descripcion: desc && desc.v ? String(desc.v).trim() : '',
          rubro: rub && rub.v ? String(rub.v).trim() : ''
        });
      }
    }

    res.status(200).json(productos);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
