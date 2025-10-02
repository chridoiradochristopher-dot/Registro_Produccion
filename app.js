const COLUMNAS = [
  "Hoja de ruta",
  "Contrato",
  "Solicitante",
  "Teléfono",
  "Destinatario",
  "Tipo de servicio",
  "Fecha de solicitud",
  "Fecha de entrega",
  "Tipo de anillo",
  "Hora",
  "Sala",
  "Coordinador",
  "Tamaño",
  "Modelo de marco",
  "Modelo de fondo",
  "Repuesto",
  "Precio Bs."
];

let registros = [];

const patterns = {
  "Hoja de ruta": [/ho[ij]a?\s*de\s*ruta.*?:?\s*([0-9]+)/i, /hoa\s*de\s*ruta.*?:?\s*([0-9]+)/i],
  "Contrato": [/ctt?o.*?:\s*([0-9]+)/i],
  "Solicitante": [/solicitante.*?:\s*(.+)/i],
  "Teléfono": [/tel[eé]fono.*?:\s*([0-9 +\-()]+)/i],
  "Destinatario": [/destinatario.*?:\s*(.+)/i],
  "Tipo de servicio": [/tipo\s*de\s*servicio.*?:\s*(.+)/i],
  "Fecha de solicitud": [/fecha\s*de\s*solicitud.*?:\s*([0-9\-\/]+)/i],
  "Fecha de entrega": [/fecha\s*de\s*entrega.*?:\s*([0-9\-\/]+)/i],
  "Tipo de anillo": [/anillo.*?:\s*(.+)/i],
  "Hora": [/hora.*?:\s*([0-9]{1,2}[:.][0-9]{2})/i],
  "Sala": [/sala.*?:\s*(.+)/i],
  "Coordinador": [/coordinador.*?:\s*(.+)/i],
  "Tamaño": [/tama[nñ]o.*?:\s*([0-9]+x[0-9]+)/i],
  "Modelo de marco": [/modelo\s*de\s*marco.*?:\s*(.+)/i],
  "Modelo de fondo": [/modelo\s*de\s*fondo.*?:\s*(.+)/i],
  "Repuesto": [/repuesto.*?:\s*(.+)/i],
  "Precio Bs.": [/precio\s*bs.*?:\s*([0-9]+)/i]
};

async function procesar() {
  registros = [];
  const files = document.getElementById("fileInput").files;
  const salida = document.getElementById("output");
  salida.textContent = "";

  for (let file of files) {
    salida.textContent += `\n=== Procesando: ${file.name} ===\n`;
    const result = await Tesseract.recognize(file, "spa");
    const texto = result.data.text;
    salida.textContent += `\n--- TEXTO OCR ---\n${texto}\n`;

    let registro = {};
    for (let col of COLUMNAS) {
      let valor = "";
      for (let re of (patterns[col] || [])) {
        const m = texto.match(re);
        if (m) { valor = m[1].trim(); break; }
      }
      registro[col] = valor;
      salida.textContent += `${col}: ${valor || "❌ NO ENCONTRADO"}\n`;
    }
    registros.push(registro);
  }
}

function descargarExcel() {
  if (!registros.length) return;
  const ws = XLSX.utils.json_to_sheet(registros);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Planillas");
  XLSX.writeFile(wb, "planillas.xlsx");
}