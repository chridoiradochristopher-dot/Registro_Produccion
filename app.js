let registros = [];

// Preprocesar imagen (blanco y negro + contraste)
function preprocessImage(file) {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      const ctx = canvas.getContext("2d");
      canvas.width = img.width;
      canvas.height = img.height;
      ctx.drawImage(img, 0, 0);
      const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
      const data = imageData.data;
      for (let i = 0; i < data.length; i += 4) {
        const avg = (data[i] + data[i+1] + data[i+2]) / 3;
        const val = avg > 150 ? 255 : 0; // umbral
        data[i] = data[i+1] = data[i+2] = val;
      }
      ctx.putImageData(imageData, 0, 0);
      canvas.toBlob(resolve, "image/png");
    };
    img.src = URL.createObjectURL(file);
  });
}

// Patrones regex ajustados
const patterns = {
  "Hoja de ruta": [/hoja\s*de\s*ruta.*?:\s*([0-9]+)/i, /hr.*?:\s*([0-9]+)/i],
  "Contrato": [/ctt?o.*?:\s*([0-9]+)/i, /cto.*?:\s*([0-9]+)/i],
  "Solicitante": [/solicitante.*?:\s*(.+)/i],
  "Teléfono": [/tel[eé]fono.*?:\s*([0-9 +\-()]+)/i],
  "Destinatario": [/destinatario.*?:\s*(.+)/i, /facilitador.*?:\s*(.+)/i],
  "Tipo de servicio": [/tipo\s*de\s*servicio.*?:\s*(.+)/i],
  "Fecha de solicitud": [/fecha\s*solicitud.*?:\s*([0-9\-\/]+)/i],
  "Fecha de entrega": [/fecha\s*entrega.*?:\s*([0-9\-\/]+)/i],
  "Tipo de anillo": [/anillo.*?:\s*(.+)/i, /lugar\s*de\s*entrega.*?:\s*(.+)/i],
  "Hora": [/hora.*?:\s*([0-9]{1,2}[:.][0-9]{2})/i],
  "Sala": [/sala.*?:\s*(.+)/i],
  "Coordinador": [/coordinador.*?:\s*(.+)/i],
  "Tamaño": [/tama[nñ]o.*?:\s*([0-9]+["']?\s*x\s*[0-9]+["']?)/i],
  "Modelo de marco": [/modelo\s*de\s*marco.*?:\s*(.+)/i],
  "Modelo de fondo": [/modelo\s*de\s*fondo.*?:\s*(.+)/i, /modelo\s*de\s*foto.*?:\s*(.+)/i],
  "Repuesto": [/repuesto.*?:\s*(.+)/i, /retoques.*?:\s*(.+)/i],
  "Precio Bs.": [/precio\s*bs.*?:\s*([0-9]+)/i]
};

async function procesar() {
  registros = [];
  const files = document.getElementById("fileInput").files;
  const salida = document.getElementById("output");
  const progreso = document.getElementById("progress");

  salida.textContent = "";
  progreso.textContent = "";

  if (!files.length) {
    alert("Selecciona al menos una imagen");
    return;
  }

  for (let file of files) {
    salida.textContent += `\n=== Procesando: ${file.name} ===\n`;

    const preprocessed = await preprocessImage(file);

    const result = await Tesseract.recognize(preprocessed, "spa+eng", {
      logger: m => {
        if (m.status === "recognizing text") {
          progreso.textContent = `Procesando ${file.name}: ${Math.round(m.progress * 100)}%`;
        }
      }
    });

    const texto = result.data.text;
    salida.textContent += `\n--- TEXTO OCR ---\n${texto}\n`;

    // Extraer campos
    let registro = {};
    for (let campo in patterns) {
      let valor = "";
      for (let re of patterns[campo]) {
        const m = texto.match(re);
        if (m) { valor = m[1].trim(); break; }
      }
      registro[campo] = valor || "❌";
      salida.textContent += `${campo}: ${registro[campo]}\n`;
    }

    registros.push(registro);
  }

  progreso.textContent = "✅ Procesamiento terminado";
}

function descargarExcel() {
  if (!registros.length) {
    alert("Primero procesa al menos una imagen");
    return;
  }
  const ws = XLSX.utils.json_to_sheet(registros);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Planillas");
  XLSX.writeFile(wb, "planillas.xlsx");
}