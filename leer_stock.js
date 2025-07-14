// Programa Node.js con Express para leer un archivo Excel y exponerlo como JSON
// Uso: node leer_stock.js [<nombre-archivo.xls>]
// Requiere instalar dependencias: npm install xlsx express

const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const express = require('express');

const app = express();
const PORT = process.env.PORT || 3000;

// Función para buscar primer archivo .xls/.xlsx en el directorio
function buscarArchivoExcel(dir) {
  const archivos = fs.readdirSync(dir);
  const encontrados = archivos.filter(f => /\.(xls|xlsx)$/i.test(f));
  return encontrados.length > 0 ? encontrados[0] : null;
}

// Función que lee el Excel y devuelve arreglo de productos
function leerProductos(fileName) {
  const filePath = path.join(__dirname, fileName);
  if (!fs.existsSync(filePath)) {
    throw new Error(`Archivo no encontrado: ${filePath}`);
  }

  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const range = xlsx.utils.decode_range(worksheet['!ref']);
  const startRow = 9;  // fila 10 en Excel (0-based index)
  const productos = [];

  for (let r = startRow; r <= range.e.r; r++) {
    const codigoCell = worksheet[xlsx.utils.encode_cell({ c: 0, r })];
    const descCell   = worksheet[xlsx.utils.encode_cell({ c: 2, r })];
    const rubroCell  = worksheet[xlsx.utils.encode_cell({ c: 5, r })];
    if (codigoCell && codigoCell.v) {
      productos.push({
        codigo: String(codigoCell.v).trim(),
        descripcion: descCell && descCell.v ? String(descCell.v).trim() : '',
        rubro: rubroCell && rubroCell.v ? String(rubroCell.v).trim() : ''
      });
    }
  }

  return productos;
}

// Ruta principal
app.get('/', (req, res) => {
  res.send('Servidor activo. Ir a /productos para obtener el JSON.');
});

// Endpoint para obtener productos en JSON
app.get('/productos', (req, res) => {
  let fileName = req.query.file;
  if (!fileName) {
    const auto = buscarArchivoExcel(__dirname);
    if (!auto) {
      return res.status(400).json({ error: 'No se encontró archivo .xls/.xlsx en la carpeta.' });
    }
    fileName = auto;
  }

  try {
    const productos = leerProductos(fileName);
    res.json(productos);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Servidor Express corriendo en http://localhost:${PORT}`));
