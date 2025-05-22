document.addEventListener('DOMContentLoaded', () => {
  // --------------------------------------------------
  // 1) Referencias a elementos del DOM
  // --------------------------------------------------
  const fileInput          = document.getElementById('fileInput');
  const progressBar        = document.getElementById('progressBar');
  const info               = document.getElementById('info');
  const fileName           = document.getElementById('fileName');
  const botonContainer     = document.getElementById('botonRedaccionContainer');
  const redaccionContainer = document.getElementById('redaccionTecnica');

  // --------------------------------------------------
  // 2) Variables que se llenarán al leer el Excel
  // --------------------------------------------------
  let propertyData   = {};
  let pointsDataMap  = {};   // { número_punto: { norte, este } }
  let boundariesData = [];   // Array de colindancias

  // --------------------------------------------------
  // 3) Listener: cuando el usuario elige un archivo
  // --------------------------------------------------
  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    // 3.1) Mostrar el nombre de archivo
    fileName.textContent = file.name;

    // 3.2) Limpiar vistas previas
    info.innerHTML = '';
    botonContainer.innerHTML = '';
    redaccionContainer.textContent = '';

    // 3.3) Simular la barra de progreso (3s)
    simulateProgressBar(3000, async () => {
      // 3.4) Leer el Excel con SheetJS
      const sheets = await readExcel(file);

      // 3.5) Convertir array de hojas a { NOMBRE_HOJA: datos }
      const excelData = {};
      sheets.forEach(sheet => {
        excelData[sheet.name.toUpperCase()] = sheet.data;
      });

      // --------------------------------------------------
      // 4) Extraer datos de la hoja "PREDIO"
      // --------------------------------------------------
      extractPropertyData(excelData["PREDIO"]);

      // --------------------------------------------------
      // 5) Extraer datos de la hoja "PUNTOS"
      // --------------------------------------------------
      extractPointsData(excelData["PUNTO_CONTROL_PTO_DETALLE"]);

      // --------------------------------------------------
      // 6) Extraer datos de la hoja "PREDIO_DIST_COLINDANTE"
      // --------------------------------------------------
      extractBoundariesData(excelData["PREDIO_DIST_COLINDANTE"]);

      // --------------------------------------------------
      // 7) Mostrar info general y tablas de colindancias
      // --------------------------------------------------
      displayPropertyData(propertyData);
      displayBoundaries(boundariesData);

      // --------------------------------------------------
      // 8) Mostrar botones para previsualizar y descargar
      // --------------------------------------------------
      mostrarBotonesRedaccion(propertyData, boundariesData, pointsDataMap);
    });
  });

  // --------------------------------------------------
  // 9) Función: simula la barra de progreso
  // --------------------------------------------------
  function simulateProgressBar(duration, callback) {
    const start = Date.now();
    const interval = setInterval(() => {
      const percent = Math.min(100, ((Date.now() - start) / duration) * 100);
      progressBar.style.width = percent + '%';
      if (percent >= 100) {
        clearInterval(interval);
        callback();
      }
    }, 100);
  }

  // --------------------------------------------------
  // 10) Función: lee un archivo Excel y devuelve array de hojas
  // --------------------------------------------------
  async function readExcel(file) {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheets = workbook.SheetNames.map(name => ({
          name,
          data: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 }),
        }));
        resolve(sheets);
      };
      reader.readAsArrayBuffer(file);
    });
  }

  // --------------------------------------------------
  // 11) Función: extrae datos generales del predio
  // --------------------------------------------------
  function extractPropertyData(sheet) {
    if (!sheet || sheet.length < 2) return;
    const headers = sheet[0].map(h => h.trim().toUpperCase());
    const row     = sheet[1];
    propertyData = {
      nombrePredio:      row[headers.indexOf("NOMBRE PREDIO")]    || "SIN INFORMACIÓN",
      nombrePropietario: row[headers.indexOf("PROPIETARIO")]     || "SIN INFORMACIÓN",
      numeroPredial:     row[headers.indexOf("NUMERO PREDIAL")]   || "SIN INFORMACIÓN",
      folioMatricula:    row[headers.indexOf("FOLIO")]            || "SIN INFORMACIÓN",
      vereda:            row[headers.indexOf("VEREDA")]           || "SIN INFORMACIÓN",
      municipio:         row[headers.indexOf("MUNICIPIO")]        || "SIN INFORMACIÓN",
      departamento:      row[headers.indexOf("DEPARTAMENTO")]     || "SIN INFORMACIÓN",
      grupoEtnico:       row[headers.indexOf("GRUPO ETNICO")]     || "SIN INFORMACIÓN",
      comunidad:         row[headers.indexOf("COMUNIDAD")]        || "SIN INFORMACIÓN",
      proyeccion:        row[headers.indexOf("PROYECCION")]       || "SIRGAS",
      epsg:              row[headers.indexOf("EPSG")]             || "7856"
    };
  }

  // --------------------------------------------------
  // 12) Función: extrae coordenadas de la hoja "PUNTOS"
  // --------------------------------------------------
function extractPointsData(sheet) {
    pointsDataMap = {};
    if (!sheet || sheet.length < 2) return;
    
    const headers = sheet[0].map(h => String(h).trim().toUpperCase());

    console.log(headers,'headers')
    
    // Buscar índices de las columnas con más alternativas posibles
    const numIndex = headers.findIndex(h => h.includes("NUMERO") || h.includes("PUNTO") || h.includes("PTO"));
    const norteIndex = headers.findIndex(h => h.includes("NORTE"));
    const esteIndex = headers.findIndex(h => h.includes("ESTE"));
    
    // Verificar que encontramos las columnas necesarias
    if (norteIndex === -1 || esteIndex === -1 || numIndex === -1) {
        console.error("No se encontraron las columnas necesarias en la hoja PUNTOS");
        console.log("Encabezados encontrados:", headers);
        return;
    }

    for (let i = 1; i < sheet.length; i++) {
        const row = sheet[i];
        if (!row || row.length === 0) continue;
        
        const numero = String(row[numIndex]).trim();
        if (!numero) continue;  // Saltar si no hay número de punto
        
        // Manejar diferentes formatos numéricos
        const norteStr = String(row[norteIndex]).replace(/[^\d,.-]/g, '').replace(',', '.');
        const esteStr = String(row[esteIndex]).replace(/[^\d,.-]/g, '').replace(',', '.');
        
        const norte = parseFloat(norteStr) || null;
        const este = parseFloat(esteStr) || null;
        if (norte !== null && este !== null) {
            pointsDataMap[numero] = { norte, este };
        } else {
            console.warn(`Punto ${numero} tiene coordenadas inválidas: N=${row[norteIndex]}, E=${row[esteIndex]}`);
        }
    }
    
    console.log("Datos de puntos cargados:", pointsDataMap);
}

  // --------------------------------------------------
  // 13) Función: extrae colindancias de "PREDIO_DIST_COLINDANTE"
  // --------------------------------------------------
  function extractBoundariesData(sheet) {
    boundariesData = [];
    if (!sheet || sheet.length < 2) return;
    const headers = sheet[0].map(h => h.trim().toUpperCase());
    for (let i = 1; i < sheet.length; i++) {
      const row = sheet[i];
      if (!row || row.length === 0) continue;
      boundariesData.push({
        puntoInicio: row[headers.indexOf("PTO_INICIO")]        || "",
        puntoFin:    row[headers.indexOf("PTO_FIN")]           || "",
        colindante:  row[headers.indexOf("COLINDANTE")]        || "",
        distancia:   row[headers.indexOf("DIST_COLINDANTE")]   || "",
        tipoLinea:   (row[headers.indexOf("LINEA")] || "recta").toLowerCase(),
        orientacion: row[headers.indexOf("ORIENTACION")]       || "",
        predio:      row[headers.indexOf("NOMBRE DEL PREDIO")] || "",
        nupre:       row[headers.indexOf("NUMERO PREDIAL")]    || "",
        folio:       row[headers.indexOf("FOLIO")]             || "",
        propietario: row[headers.indexOf("PROPIETARIO")]      || ""
      });
    }
  }

  // --------------------------------------------------
  // 14) Función: muestra info general del predio en HTML
  // --------------------------------------------------
  function displayPropertyData(pd) {
    info.innerHTML = `
      <div class="border-t pt-4">
        <h2 class="text-lg font-semibold text-blue-700 mb-2">Información del Predio</h2>
        <ul class="grid grid-cols-1 md:grid-cols-2 gap-2">
          <li><strong>Nombre del Predio:</strong> ${pd.nombrePredio}</li>
          <li><strong>Propietario:</strong> ${pd.nombrePropietario}</li>
          <li><strong>NUPRE / Código Predial:</strong> ${pd.numeroPredial}</li>
          <li><strong>Folio Matrícula:</strong> ${pd.folioMatricula}</li>
          <li><strong>Vereda:</strong> ${pd.vereda}</li>
          <li><strong>Municipio:</strong> ${pd.municipio}</li>
          <li><strong>Departamento:</strong> ${pd.departamento}</li>
          <li><strong>Grupo Étnico:</strong> ${pd.grupoEtnico}</li>
          <li><strong>Comunidad:</strong> ${pd.comunidad}</li>
          <li><strong>Proyección:</strong> ${pd.proyeccion}</li>
          <li><strong>EPSG:</strong> ${pd.epsg}</li>
        </ul>
      </div>
    `;
  }

  // --------------------------------------------------
  // 15) Función: muestra colindancias agrupadas por orientación
  // --------------------------------------------------
  function displayBoundaries(boundaries) {
    if (!boundaries || boundaries.length === 0) {
      info.insertAdjacentHTML('beforeend', `<p class="mt-4 text-gray-700">No hay colindancias para mostrar.</p>`);
      return;
    }

    // Agrupar por orientación
    const agrupadas = boundaries.reduce((acc, b) => {
      const ori = (b.orientacion || "").toUpperCase() || "SIN ORIENTACIÓN";
      if (!acc[ori]) acc[ori] = [];
      acc[ori].push(b);
      return acc;
    }, {});

    // Recorrer cada orientación
    Object.entries(agrupadas).forEach(([ori, lista]) => {
      let html = `
        <div class="mt-4">
          <h3 class="font-medium text-blue-600 uppercase">POR EL ${ori}:</h3>
          <table class="w-full text-sm text-left border border-gray-300 mt-2">
            <thead>
              <tr class="bg-gray-200">
                <th class="px-2 py-1 border border-gray-300">Pto Inicio</th>
                <th class="px-2 py-1 border border-gray-300">Pto Fin</th>
                <th class="px-2 py-1 border border-gray-300">Colindante</th>
                <th class="px-2 py-1 border border-gray-300">Distancia (m)</th>
                <th class="px-2 py-1 border border-gray-300">Línea</th>
              </tr>
            </thead>
            <tbody>
      `;

      lista.forEach(b => {
        html += `
          <tr class="border border-gray-300">
            <td class="px-2 py-1 border border-gray-300">${b.puntoInicio}</td>
            <td class="px-2 py-1 border border-gray-300">${b.puntoFin}</td>
            <td class="px-2 py-1 border border-gray-300">${b.colindante}</td>
            <td class="px-2 py-1 border border-gray-300">${b.distancia}</td>
            <td class="px-2 py-1 border border-gray-300">${b.tipoLinea}</td>
          </tr>
        `;
      });

      html += `</tbody></table></div>`;
      info.insertAdjacentHTML('beforeend', html);
    });
  }

  // --------------------------------------------------
  // 16) Función: muestra dos botones (previsualizar y descargar)
  // --------------------------------------------------
  function mostrarBotonesRedaccion(pd, bd, puntos) {
    botonContainer.innerHTML = '';

    // Botón 1 → Previsualizar en pantalla
    const btnPreview = document.createElement('button');
    btnPreview.textContent = "Previsualizar redacción técnica";
    btnPreview.className = "mr-2 px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded";
    btnPreview.addEventListener('click', () => {
      generarRedaccionEnPantalla(pd, bd, puntos);
    });
    botonContainer.appendChild(btnPreview);

    // Botón 2 → Descargar como .docx
    const btnDownload = document.createElement('button');
    btnDownload.textContent = "Descargar redacción técnica (.docx)";
    btnDownload.className = "px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded";
    btnDownload.addEventListener('click', () => {
      generarDocumentoWord(pd, bd, puntos);
    });
    botonContainer.appendChild(btnDownload);
  }

  // --------------------------------------------------
  // 17) Función: genera la redacción y la muestra en <pre id="redaccionTecnica">
  // --------------------------------------------------
  function generarRedaccionEnPantalla(pd, bd, puntos) {
    if (!pd || !bd || bd.length === 0) return;

    // Agrupar colindancias por orientación
    const agrupadas = bd.reduce((acc, b) => {
      const ori = (b.orientacion || "").toUpperCase() || "SIN ORIENTACIÓN";
      if (!acc[ori]) acc[ori] = [];
      acc[ori].push(b);
      return acc;
    }, {});

    // 1) Encabezado y descripción
    let texto = '';
    texto += "DESCRIPCIÓN TÉCNICA\n\n";
    texto += `El bien inmueble identificado con nombre ${pd.nombrePredio} y catastralmente con NUPRE / Número predial ${pd.numeroPredial}, folio de matrícula inmobiliaria ${pd.folioMatricula}, ubicado en la vereda ${pd.vereda} el Municipio de ${pd.municipio} departamento de ${pd.departamento}; del grupo étnico ${pd.grupoEtnico}, pueblo / resguardo / comunidad ${pd.comunidad}; presenta los siguientes linderos referidos al sistema de referencia magna sirgas, con proyección ${pd.proyeccion} y EPSG ${pd.epsg}:\n\n`;

    // 2) Título LINDEROS TÉCNICOS
    texto += "LINDEROS TÉCNICOS\n\n";

    // 3) Por cada orientación en el orden NORTE, ESTE, SUR, OESTE
    const ordenOrientaciones = ["NORTE", "ESTE", "SUR", "OESTE"];
    ordenOrientaciones.forEach(ori => {
      const lista = agrupadas[ori];
      if (!lista || lista.length === 0) return;

      texto += `POR EL ${ori}:\n`;
      lista.forEach((b, idx) => {
        // extraer puntos intermedios
        const startNum = parseInt(b.puntoInicio, 10);
        const endNum   = parseInt(b.puntoFin, 10);
        let todosPuntos = [];

        if (!isNaN(startNum) && !isNaN(endNum)) {
          const step = endNum > startNum ? 1 : -1;
          for (let k = startNum; step > 0 ? k <= endNum : k >= endNum; k += step) {
            todosPuntos.push(k);
          }
        } else {
          todosPuntos = [b.puntoInicio, b.puntoFin];
        }

        // Obtener coordenadas de todos los puntos
        const puntosCoords = todosPuntos.map(num => {
          const puntoKey = num.toString();
          return {
            num: puntoKey,
            coords: puntos[puntoKey] || { norte: 0, este: 0 }
          };
        });

        // preparar texto del lindero (1, 2, 3, …)
        // let linea = `Lindero ${idx + 1}: Inicia en el punto ${b.puntoInicio} con coordenadas planas N= ${puntosCoords[0].coords.norte.toFixed(2)} m, E= ${puntosCoords[0].coords.este.toFixed(2)} m, en línea ${b.tipoLinea} en sentido ${b.orientacion}`;
         let linea = `Lindero ${idx + 1}: Inicia en el punto ${b.puntoInicio} con coordenadas planas N= ${puntosCoords[0].coords.norte.toLocaleString('es-ES', {minimumFractionDigits: 2, maximumFractionDigits: 2})} m, E= ${puntosCoords[0].coords.este.toLocaleString('es-ES', {minimumFractionDigits: 2, maximumFractionDigits: 2})} m, en línea ${b.tipoLinea} en sentido ${b.orientacion}`;
        
        // Agregar puntos intermedios si existen
        if (puntosCoords.length > 2) {
          const intermedios = puntosCoords.slice(1, -1).map(p => {
            return `punto ${p.num} N= ${p.coords.norte.toFixed(2)} m, E= ${p.coords.este.toFixed(2)} m`;
          }).join(', ');
          linea += `, pasando por los puntos de coordenadas ${intermedios}`;
        }
        
        linea += `, en una distancia de ${parseFloat(b.distancia).toFixed(2)} m, hasta encontrar el punto número ${b.puntoFin} de coordenadas planas N= ${puntosCoords[puntosCoords.length-1].coords.norte.toFixed(2)} m, E= ${puntosCoords[puntosCoords.length-1].coords.este.toFixed(2)} m, colindando con el predio identificado con nombre ${b.predio}, NUPRE/ Código predial ${b.nupre}, Folio de matrícula inmobiliaria ${b.folio} del(a) señor(a) ${b.propietario}.\n\n`;

        texto += linea;
      });
    });

    // Resto del código permanece igual...
    // [Mantener las secciones de RESULTADOS, OBSERVACIONES y NOTAS ACLARATORIAS]

    // Mostrar en pantalla
    redaccionContainer.textContent = texto;
}
  // --------------------------------------------------
  // 18) Función: genera el archivo .docx con la misma redacción
  // --------------------------------------------------
  async function generarDocumentoWord(pd, bd, puntos) {
    if (!pd || !bd || bd.length === 0) return;

    // Check if docx is available
    if (!window.docx || !window.docx.Document) {
        alert("Error: La librería para generar documentos Word no está disponible. Por favor recarga la página.");
        console.error("docx library not found", window.docx);
        return;
    }

    // 18.1) Importar clases de docx.js
        const { Document, Packer, Paragraph, TextRun } = window.docx;

    // Agrupar como en pantalla
    const agrupadas = bd.reduce((acc, b) => {
      const ori = (b.orientacion || "").toUpperCase() || "SIN ORIENTACIÓN";
      if (!acc[ori]) acc[ori] = [];
      acc[ori].push(b);
      return acc;
    }, {});

    // 18.2) Armar el array de Paragraphs
    const children = [];

    // Encabezado
    children.push(new Paragraph({
      children: [ new TextRun({ text: "DESCRIPCIÓN TÉCNICA", bold: true, size: 32 }) ],
      spacing: { after: 300 }
    }));

    // Párrafo descriptivo
    const textoDesc = 
      `El bien inmueble identificado con nombre ${pd.nombrePredio} y catastralmente con NUPRE / Número predial ${pd.numeroPredial}, folio de matrícula inmobiliaria ${pd.folioMatricula}, ubicado en la vereda ${pd.vereda} el Municipio de ${pd.municipio} departamento de ${pd.departamento}; del grupo étnico ${pd.grupoEtnico}, pueblo / resguardo / comunidad ${pd.comunidad}; presenta los siguientes linderos referidos al sistema de referencia magna sirgas, con proyección ${pd.proyeccion} y EPSG ${pd.epsg}:`;

    children.push(new Paragraph({
      children: [ new TextRun({ text: textoDesc }) ],
      spacing: { after: 300 }
    }));

    // Subtítulo LINDEROS TÉCNICOS
    children.push(new Paragraph({
      children: [ new TextRun({ text: "LINDEROS TÉCNICOS", bold: true, size: 28 }) ],
      spacing: { after: 300 }
    }));

    // Por cada orientación
    const ordenOrientaciones = ["NORTE", "ESTE", "SUR", "OESTE"];
    ordenOrientaciones.forEach(ori => {
      const lista = agrupadas[ori];
      if (!lista || lista.length === 0) return;

      // Título “POR EL …”
      children.push(new Paragraph({
        children: [ new TextRun({ text: `POR EL ${ori}:`, bold: true }) ],
        spacing: { before: 200, after: 200 }
      }));

      lista.forEach((b, idx) => {
        // reconstruir puntos intermedios
        const startNum = parseInt(b.puntoInicio, 10);
        const endNum   = parseInt(b.puntoFin, 10);
        let todosPuntos = [];

        if (!isNaN(startNum) && !isNaN(endNum)) {
          const step = endNum > startNum ? 1 : -1;
          for (let k = startNum; step > 0 ? k <= endNum : k >= endNum; k += step) {
            todosPuntos.push(k);
          }
        } else {
          todosPuntos = [b.puntoInicio, b.puntoFin];
        }

        const inicioCoords = puntos[b.puntoInicio] || { norte: 0, este: 0 };
        const finCoords    = puntos[b.puntoFin]    || { norte: 0, este: 0 };

        const intermedios = todosPuntos.slice(1, -1).map(num => {
          const pm = puntos[num] || { norte: 0, este: 0 };
          return `punto ${num}  N= ${pm.norte} m, E= ${pm.este}`;
        }).join(', ');

        // Armar texto
        let textoL = 
          `Lindero ${idx + 1}: Inicia en el punto ${b.puntoInicio} con coordenadas planas N= ${inicioCoords.norte} m, E= ${inicioCoords.este} m, en línea ${b.tipoLinea} en sentido ${b.orientacion}`;
        if (intermedios) {
          textoL += `, pasando por los puntos de coordenadas ${intermedios};`;
        }
        textoL += 
          ` en una distancia de ${parseFloat(b.distancia)} m, hasta encontrar el punto número ${b.puntoFin} de coordenadas planas N= ${finCoords.norte} m, E= ${finCoords.este} m, colindando con el predio identificado con nombre ${b.predio}, NUPRE/ Código predial ${b.nupre}, Folio de matrícula inmobiliaria ${b.folio} del(a) señor(a) ${b.propietario}.`;

        children.push(new Paragraph({
          children: [ new TextRun({ text: textoL }) ],
          spacing: { after: 300 }
        }));
      });
    });

    // RESULTADOS
    children.push(new Paragraph({
      children: [ new TextRun({ text: "RESULTADOS:", bold: true }) ],
      spacing: { before: 300, after: 200 }
    }));
    children.push(new Paragraph({
      children: [ new TextRun({ text: `De acuerdo con los anteriores linderos, el área del citado bien inmueble es de: ________ ha + ________ m².` }) ],
      spacing: { after: 300 }
    }));

    // OBSERVACIONES
    const textoObs = 
      `OBSERVACIONES:\n` +
      `• El predio fue levantado mediante método <Directo, indirecto, mixto o colaborativo> de acuerdo con el Decreto DANE 148 de 2020 y la Resolución IGAC 1040 de 2023 modificada parcialmente por la Resolución IGAC 746 de 2024.\n` +
      `• El nombre del predio indicado corresponde a información <del interesado, del FMI, de las bases catastrales>.\n` +
      `• La presente redacción de linderos se hizo con base al plano ID: <Descripción>, con fecha de <Mes> del <Año>, levantado por el profesional <Topógrafo, Tec. Topografía ó Ingeniero Topográfico> <Nombre Completo>. Con Matrícula Profesional No. <Descripción>.\n\n` +
      `FIRMA:\n\n` +
      `_______________________________\n` +
      `Profesión: ____________________\n` +
      `Nombre: ______________________\n` +
      `Apellido: _____________________\n` +
      `Matrícula Profesional No.: _________\n\n`;

    children.push(new Paragraph({
      children: [ new TextRun({ text: textoObs }) ],
      spacing: { after: 300 }
    }));

    // NOTAS ACLARATORIAS
    const notas =
      `NOTAS ACLARATORIAS\n\n` +
      `1. Se tomará como punto de inicio, el que se ubique en el costado más noroccidental del predio y se hará conforme a las manecillas del reloj.\n` +
      `2. Colindando con el predio identificado con Nombre predio, NUPRE/Código Predial y FMI, Propietario, o (elemento y nombre geográfico...), solo se diligenciará en caso de separación diferente: vía, calle, río, etc.\n` +
      `3. Nombrar vértices: El vértice se nombrará como “punto”. Ejemplo: punto XXX.\n` +
      `4. Se deben REDACTAR TODOS LOS PUNTOS con sus respectivas coordenadas en el formato indicado. Si en un lindero hay cambio de sentido, se inicia un nuevo párrafo sin nombrar un nuevo lindero.\n` +
      `5. Direcciones: norte, sur, este, oeste y derivaciones (noreste, suroeste, etc.).\n` +
      `6. Distancias: en metros (m), aproximación al decímetro, separador decimal punto (.), según resolución IGAC 1101 de 2020.\n` +
      `7. Coordenadas: expresadas en metros, aproximación al centímetro. Ejemplo: N= 882285.77 m y E= 1542549.99 m.\n` +
      `8. Área: en (ha + m²) para predios rurales. Para urbanos (m²) con aproximación al decímetro.\n` +
      `9. Se utiliza punto como separador decimal. Si se usa otro, aclarar.\n` +
      `10. Redacción de linderos por Topógrafo, Tec. topografía, Ing. Topográfico, Ing. Catastral, Geodesta o Geógrafo.\n` +
      `11. Si falta información (nombre predio, propietario, NUPRE, folio), colocar “SIN INFORMACIÓN”.\n`;

    children.push(new Paragraph({
      children: [ new TextRun({ text: notas }) ],
      spacing: { after: 300 }
    }));

    // 18.3) Crear documento y descargar
    const doc = new Document({ sections: [{ properties: {}, children }] });
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Redaccion_Tecnica.docx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }
});
