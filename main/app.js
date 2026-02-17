document.addEventListener('DOMContentLoaded', () => {

  const fileInput = document.getElementById('fileInput');
  const redaccionContainer = document.getElementById('redaccionTecnica');
  const botonContainer = document.getElementById('botonRedaccionContainer');
  const panelAcciones = document.getElementById('panelAcciones');

  let pointsDataMap = {};
  let boundariesData = [];

  fileInput.addEventListener('change', async (e) => {

    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (event) => {

      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const res = {};
      workbook.SheetNames.forEach(n => {
        res[n.toUpperCase()] = XLSX.utils.sheet_to_json(workbook.Sheets[n], { header: 1 });
      });

      const sheet = res["PREDIO"];

      if (sheet) {
        processHojaPredio(sheet);
        panelAcciones.classList.remove('hidden');
        mostrarBotonGenerar();
        dibujarPlano();
      } else {
        alert("No se encontró la hoja 'PREDIO'");
      }
    };

    reader.readAsArrayBuffer(file);
  });

  function processHojaPredio(data) {

    const headers = data[0].map(h => String(h || "").trim());

    const idIdx = headers.indexOf("Id");
    const yIdx = headers.indexOf("Y");
    const xIdx = headers.indexOf("X");
    const puntosRangeIdx = headers.indexOf("PUNTOS");
    const colindanteIdx = headers.indexOf("COLINDANTES");
    const distIdx = headers.indexOf("DISTANCIA (m)");
    const propIdx = headers.indexOf("PROPIETARIO");
    const fmiIdx = headers.indexOf("FMI");
    const nupreIdx = headers.indexOf("NUPRE");

    pointsDataMap = {};
    boundariesData = [];

    data.slice(1).forEach(row => {

      if (row[idIdx] !== undefined && row[idIdx] !== "") {
        pointsDataMap[String(row[idIdx]).trim()] = {
          norte: parseFloat(row[yIdx]),
          este: parseFloat(row[xIdx])
        };
      }

      if (row[puntosRangeIdx]) {

        const num = String(row[puntosRangeIdx]).match(/\d+/g);

        if (num && num.length >= 2) {

          boundariesData.push({
            pInicio: parseInt(num[0]),
            pFin: parseInt(num[1]),
            colindante: row[colindanteIdx] || "SIN INFORMACION",
            distAcu: row[distIdx] || "",
            prop: row[propIdx] || "SIN INFORMACION",
            fmi: row[fmiIdx] || "SIN INFORMACION",
            nupre: row[nupreIdx] || "SIN INFORMACION"
          });
        }
      }
    });
  }

  function obtenerSentidoCartesiano(p1, p2) {

    const dN = p2.norte - p1.norte;
    const dE = p2.este - p1.este;

    if (dN > 0 && dE > 0) return "Noreste";
    if (dN > 0 && dE < 0) return "Noroeste";
    if (dN < 0 && dE > 0) return "Sureste";
    if (dN < 0 && dE < 0) return "Suroeste";

    if (dN > 0) return "Norte";
    if (dN < 0) return "Sur";
    if (dE > 0) return "Este";
    if (dE < 0) return "Oeste";

    return "Norte";
  }

  function obtenerOrientacionDominante(b) {
  const step = b.pInicio < b.pFin ? 1 : -1;
  let puntosLindero = [];

  // Recolectamos los puntos que forman este lindero específico
  for (let i = b.pInicio; i !== b.pFin + step; i += step) {
    if (pointsDataMap[i]) {
      puntosLindero.push(pointsDataMap[i]);
    }
  }

  // Si no hay puntos, por defecto Norte (evita errores)
  if (puntosLindero.length === 0) return "POR EL NORTE:";

  // Calculamos los límites (extremos) de todo el terreno
  const todosLosPuntos = Object.values(pointsDataMap);
  const maxN = Math.max(...todosLosPuntos.map(p => p.norte));
  const minN = Math.min(...todosLosPuntos.map(p => p.norte));
  const maxE = Math.max(...todosLosPuntos.map(p => p.este));
  const minE = Math.min(...todosLosPuntos.map(p => p.este));

  // Hallamos el centro exacto del predio completo
  const centroPredioN = (maxN + minN) / 2;
  const centroPredioE = (maxE + minE) / 2;

  // Hallamos el centro promedio de este lindero
  const centroLinderoN = puntosLindero.reduce((a, p) => a + p.norte, 0) / puntosLindero.length;
  const centroLinderoE = puntosLindero.reduce((a, p) => a + p.este, 0) / puntosLindero.length;

  // Calculamos la diferencia de posición respecto al centro
  const difN = centroLinderoN - centroPredioN;
  const difE = centroLinderoE - centroPredioE;

  // Comparamos qué eje domina (si está más lejos horizontal o verticalmente)
  if (Math.abs(difE) > Math.abs(difN)) {
    // Si la diferencia Este/Oeste es mayor:
    return difE > 0 ? "POR EL ESTE:" : "POR EL OESTE:";
  } else {
    // Si la diferencia Norte/Sur es mayor:
    return difN > 0 ? "POR EL NORTE:" : "POR EL SUR:";
  }
}

  /*
  function obtenerOrientacionDominante(b) {
    const step = b.pInicio < b.pFin ? 1 : -1;
    let puntos = [];

    for (let i = b.pInicio; i !== b.pFin + step; i += step) {
      if (pointsDataMap[i]) {
        puntos.push(pointsDataMap[i]);
      }
    }

    if (puntos.length === 0) return "POR EL NORTE:";

    const todos = Object.values(pointsDataMap);
    const maxN = Math.max(...todos.map(p => p.norte));
    const minN = Math.min(...todos.map(p => p.norte));
    const maxE = Math.max(...todos.map(p => p.este));
    const minE = Math.min(...todos.map(p => p.este));

    // Centro del lindero actual
    const centroN = puntos.reduce((a, p) => a + p.norte, 0) / puntos.length;
    const centroE = puntos.reduce((a, p) => a + p.este, 0) / puntos.length;
    
    // Centro geográfico del predio completo
    const midN = (maxN + minN) / 2;
    const midE = (maxE + minE) / 2;

    // Distancias a los ejes centrales
    const distN = centroN - midN;
    const distE = centroE - midE;

    // Determinamos cuál es el eje más alejado del centro para asignar orientación
    if (Math.abs(distN) > Math.abs(distE)) {
      return distN > 0 ? "POR EL NORTE:" : "POR EL SUR:";
    } else {
      return distE > 0 ? "POR EL ESTE:" : "POR EL OESTE:";
    }
  }
*/



  function generarRedaccion() {

    let t = "**LINDEROS TÉCNICOS**\n\n";

    let orientacionActual = "";

    boundariesData.forEach((b, idx) => {

      const step = b.pInicio < b.pFin ? 1 : -1;
      const orientacion = obtenerOrientacionDominante(b);

      if (orientacion !== orientacionActual) {
        orientacionActual = orientacion;
        t += `${orientacion}\n\n`;
      }

      let tramos = [];
      let tramoActual = { inicio: b.pInicio, fin: null, puntos: [], sentido: "" };

      for (let i = b.pInicio; i !== b.pFin; i += step) {

        const p1 = pointsDataMap[i];
        const p2 = pointsDataMap[i + step];
        if (!p1 || !p2) continue;

        const sentido = obtenerSentidoCartesiano(p1, p2);

        if (tramoActual.sentido === "") tramoActual.sentido = sentido;

        if (sentido !== tramoActual.sentido) {

          tramoActual.fin = i;
          tramos.push(tramoActual);

          tramoActual = {
            inicio: i,
            fin: null,
            puntos: [],
            sentido: sentido
          };
        }

        tramoActual.puntos.push(i + step);
      }

      tramoActual.fin = b.pFin;
      tramos.push(tramoActual);

      tramos.forEach((tr, index) => {

        const pI = pointsDataMap[tr.inicio];
        const pF = pointsDataMap[tr.fin];

        let parrafo = "";

        if (index === 0) {
          parrafo += `<strong>Lindero ${idx + 1}:</strong> Inicia en el punto número ${tr.inicio} de coordenadas planas N= ${pI.norte.toFixed(2)}m, E= ${pI.este.toFixed(2)}m, `;
        } else {
          parrafo += `Continúa en el punto número ${tr.inicio} de coordenadas planas N= ${pI.norte.toFixed(2)}m, E= ${pI.este.toFixed(2)}m, `;
        }

        const esQuebrada = tr.puntos.length > 1;

        parrafo += `en línea ${esQuebrada ? 'quebrada' : 'recta'} en sentido ${tr.sentido}, `;

        if (esQuebrada) {

          const intermedios = tr.puntos.slice(0, -1).map(p =>
            `el punto número ${p} de coordenadas planas N= ${pointsDataMap[p].norte.toFixed(2)}m, E= ${pointsDataMap[p].este.toFixed(2)}m`
          );

          if (intermedios.length > 0) {
            parrafo += `pasando por ${intermedios.join(', ')}, `;
          }
        }

        parrafo += `hasta encontrar el punto número ${tr.fin} de coordenadas planas N= ${pF.norte.toFixed(2)}m, E= ${pF.este.toFixed(2)}m`;

        if (index === tramos.length - 1 && b.distAcu) {
          parrafo += ` con una distancia total acumulada de ${Number(b.distAcu).toFixed(1)}m`;
        }

        if (index === tramos.length - 1) {
          parrafo += ` colindando con un predio ${b.colindante}, el NUPRE Código predial ${b.nupre}, Folio de matrícula inmobiliaria ${b.fmi} y de propietario ${b.prop}.`;
        }

        t += parrafo + "\n\n";
      });

    });

    redaccionContainer.innerHTML = t.replace(/\n/g, "<br>");
    mostrarBotonDescarga(t);
  }

  function mostrarBotonGenerar() {

    botonContainer.innerHTML = '';

    const btn = document.createElement('button');
    btn.className = "w-full py-3 bg-blue-600 text-white rounded-xl font-bold mb-4";
    btn.innerText = "Generar Redacción Técnica";
    btn.onclick = generarRedaccion;

    botonContainer.appendChild(btn);
  }

  /* =========================
     NO TOQUÉ EL PLANO
  ========================== */

  function dibujarPlano() {

    const canvas = document.getElementById("previewPlano");
    if (!canvas) return;

    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const puntosOrdenados = Object.entries(pointsDataMap)
      .sort((a, b) => parseInt(a[0]) - parseInt(b[0]));

    if (puntosOrdenados.length < 2) return;

    const puntos = puntosOrdenados.map(p => p[1]);

    const maxN = Math.max(...puntos.map(p => p.norte));
    const minN = Math.min(...puntos.map(p => p.norte));
    const maxE = Math.max(...puntos.map(p => p.este));
    const minE = Math.min(...puntos.map(p => p.este));

    const padding = 40;
    const scaleX = (canvas.width - padding * 2) / (maxE - minE);
    const scaleY = (canvas.height - padding * 2) / (maxN - minN);
    const scale = Math.min(scaleX, scaleY);

    ctx.beginPath();

    puntosOrdenados.forEach((p, i) => {

      const id = p[0];
      const punto = p[1];

      const x = padding + (punto.este - minE) * scale;
      const y = canvas.height - padding - (punto.norte - minN) * scale;

      if (i === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);

      ctx.beginPath();
      ctx.arc(x, y, 4, 0, Math.PI * 2);
      ctx.fillStyle = "red";
      ctx.fill();

      ctx.fillStyle = "black";
      ctx.font = "12px Arial";
      ctx.fillText(id, x + 6, y - 6);
    });

    ctx.closePath();
    ctx.strokeStyle = "#2563eb";
    ctx.lineWidth = 2;
    ctx.stroke();
  }

  function mostrarBotonDescarga(texto) {

    if (document.getElementById('btnWord')) return;

    const btn = document.createElement('button');
    btn.id = 'btnWord';
    btn.className = "w-full py-3 bg-green-600 text-white rounded-xl font-bold shadow-lg hover:bg-green-700";
    btn.innerHTML = 'Descargar Word (.docx)';
    btn.onclick = () => exportarAWord(texto);

    botonContainer.appendChild(btn);
  }

  async function exportarAWord(texto) {

    const { Document, Packer, Paragraph, TextRun } = docx;

    const lineas = texto.split('\n');

    const paragraphs = lineas.map(linea => {

      if (linea.trim() === "") {
        return new Paragraph({
          children: [new TextRun("")],
          spacing: { after: 300 }
        });
      }

      return new Paragraph({
        children: [
          new TextRun({
            text: linea,
            font: "Arial",
            size: 22
          })
        ],
        alignment: "both",
        spacing: { line: 360, after: 200 }
      });
    });

    const doc = new Document({
      sections: [{ children: paragraphs }]
    });

    const blob = await Packer.toBlob(doc);

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Redaccion_Linderos.docx";
    link.click();
  }


});
