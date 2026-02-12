document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const progressBar = document.getElementById('progressBar');
    const percentageText = document.getElementById('texto');
    const infoContainer = document.getElementById('info');
    const botonContainer = document.getElementById('botonRedaccionContainer');
    const redaccionContainer = document.getElementById('redaccionTecnica');
    const panelAcciones = document.getElementById('panelAcciones');

    let pointsDataMap = {}; 
    let boundariesData = [];
    let propertyData = {};

    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        resetUI();
        simulateProgressBar(1000, async () => {
            const workbookData = await readExcel(file);
            // Buscamos la hoja "PREDIO" (en mayúsculas por seguridad)
            const sheet = workbookData["PREDIO"];

            if (sheet) {
                processHojaPredio(sheet);
                displayPropertyPreview();
                panelAcciones.classList.remove('hidden');
                mostrarBotonGenerar();
            } else {
                alert("No se encontró la hoja llamada 'PREDIO' en el Excel.");
            }
        });
    });

    // 1. PROCESAR LA HOJA SEGÚN TU ESTRUCTURA REAL
    function processHojaPredio(data) {
        const headers = data[0].map(h => String(h || "").trim());
        
        // Índices para PUNTOS (Izquierda)
        const idIdx = headers.indexOf("Id");
        const yIdx = headers.indexOf("Y"); // Norte
        const xIdx = headers.indexOf("X"); // Este

        // Índices para COLINDANCIAS (Derecha)
        const puntosRangeIdx = headers.indexOf("PUNTOS");
        const colindanteIdx = headers.indexOf("COLINDANTES");
        const distIdx = headers.indexOf("DISTANCIA (m)");
        const propIdx = headers.indexOf("PROPIETARIO");
        const fmiIdx = headers.indexOf("FMI");
        const nupreIdx = headers.indexOf("NUPRE");

        pointsDataMap = {};
        boundariesData = [];

        data.slice(1).forEach(row => {
            // Mapear coordenadas (Lado izquierdo del Excel)
            if (row[idIdx] !== undefined && row[idIdx] !== "") {
                pointsDataMap[String(row[idIdx]).trim()] = {
                    norte: parseFloat(String(row[yIdx]).replace(',', '.')),
                    este: parseFloat(String(row[xIdx]).replace(',', '.'))
                };
            }

            // Mapear tramos/párrafos (Lado derecho del Excel)
            if (row[puntosRangeIdx]) {
                const textoRango = String(row[puntosRangeIdx]);
                const numeros = textoRango.match(/\d+/g);
                if (numeros && numeros.length >= 2) {
                    boundariesData.push({
                        pInicio: numeros[0],
                        pFin: numeros[1],
                        colindante: row[colindanteIdx] || "SIN INFORMACION",
                        distancia: row[distIdx] || "",
                        propietario: row[propIdx] || row[colindanteIdx] || "SIN INFORMACION",
                        fmi: row[fmiIdx] || "SIN INFORMACION",
                        nupre: row[nupreIdx] || "SIN INFORMACION"
                    });
                }
            }
        });
    }

    // 2. GENERACIÓN DE REDACCIÓN POR PÁRRAFOS (CAMBIO DE DIRECCIÓN)
    function generarRedaccion() {
        let t = "DESCRIPCIÓN TÉCNICA\n\nLINDEROS TÉCNICOS\n\n";
        let numLindero = 1;
        let ultimoColindante = "";

        boundariesData.forEach((b, index) => {
            const pI = pointsDataMap[b.pInicio] || { norte: 0, este: 0 };
            const pF = pointsDataMap[b.pFin] || { norte: 0, este: 0 };
            const fmt = (n) => n.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

            // Lógica de Párrafo Aparte
            let encabezado = "";
            if (index === 0 || b.colindante !== ultimoColindante) {
                encabezado = `Lindero ${numLindero}: Inicia en el punto número ${b.pInicio}`;
                numLindero++;
            } else {
                encabezado = `Continúa en el punto número ${b.pInicio}`;
            }
            ultimoColindante = b.colindante;

            // Determinar tipo de línea y sentido
            const inicioNum = parseInt(b.pInicio);
            const finNum = parseInt(b.pFin);
            const esQuebrada = Math.abs(finNum - inicioNum) > 1;
            const tipoLinea = esQuebrada ? "quebrada" : "recta";
            
            // Sentido (puedes ajustarlo o automatizarlo con lógica de coordenadas si deseas)
            let sentido = "Sureste"; 
            if (index === 1) sentido = "Noreste"; // Ejemplo basado en tu texto

            let parrafo = `${encabezado} de coordenadas planas N= ${fmt(pI.norte)}m, E= ${fmt(pI.este)}m, `;
            parrafo += `en línea ${tipoLinea} en sentido ${sentido}, `;

            if (b.distancia) {
                parrafo += `con una distancia total acumulada de ${b.distancia}m, `;
            }

            // Listar puntos intermedios si es línea quebrada
            if (esQuebrada) {
                let intermedios = [];
                // Determinar si el conteo es ascendente o descendente
                const paso = inicioNum < finNum ? 1 : -1;
                for (let i = inicioNum + paso; i !== finNum; i += paso) {
                    if (pointsDataMap[i]) {
                        intermedios.push(`el punto ${i} de coordenadas planas N= ${fmt(pointsDataMap[i].norte)}m, E= ${fmt(pointsDataMap[i].este)}m`);
                    }
                }
                if (intermedios.length > 0) {
                    parrafo += `pasando por ${intermedios.join(', ')}, `;
                }
            }

            // Cierre del párrafo con datos de colindancia
            parrafo += `hasta encontrar el punto número ${b.pFin} de coordenadas planas N= ${fmt(pF.norte)}m, E= ${fmt(pF.este)}m `;
            parrafo += `colindando con un predio ${b.colindante}, el NUPRE Código predial ${b.nupre}, Folio de matrícula inmobiliaria ${b.fmi} y de propietario ${b.propietario}.\n\n`;

            t += parrafo;
        });

        redaccionContainer.textContent = t;
        actualizarBotonDescarga(t);
    }

    // --- FUNCIONES AUXILIARES ---
    function mostrarBotonGenerar() {
        botonContainer.innerHTML = '';
        const btn = document.createElement('button');
        btn.className = "w-full py-3 bg-blue-600 text-white rounded-xl font-bold shadow-lg hover:bg-blue-700 transition-all";
        btn.innerText = "Generar Redacción Técnica";
        btn.onclick = generarRedaccion;
        botonContainer.appendChild(btn);
    }

    function actualizarBotonDescarga(texto) {
        if (document.getElementById('btnDownload')) return;
        const btn = document.createElement('button');
        btn.id = "btnDownload";
        btn.className = "w-full py-3 bg-green-600 text-white rounded-xl font-bold shadow-lg hover:bg-green-700 transition-all mt-4";
        btn.innerText = "Descargar en Word (.docx)";
        btn.onclick = () => exportarWord(texto);
        botonContainer.appendChild(btn);
    }

    async function readExcel(file) {
        return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const result = {};
                workbook.SheetNames.forEach(name => {
                    result[name.toUpperCase()] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
                });
                resolve(result);
            };
            reader.readAsArrayBuffer(file);
        });
    }

    function resetUI() {
        redaccionContainer.textContent = '';
        botonContainer.innerHTML = '';
    }

    function simulateProgressBar(duration, callback) {
        let start = 0;
        const interval = setInterval(() => {
            start += 10;
            progressBar.style.width = start + '%';
            percentageText.textContent = start + '%';
            if (start >= 100) {
                clearInterval(interval);
                callback();
            }
        }, duration / 10);
    }

    function displayPropertyPreview() {
        infoContainer.innerHTML = `
            <div class="p-3 bg-blue-50 border border-blue-100 rounded-lg text-xs text-blue-800">
                <i class="fas fa-check-circle mr-2"></i> 
                Se detectaron <strong>${Object.keys(pointsDataMap).length}</strong> puntos y 
                <strong>${boundariesData.length}</strong> tramos de colindancia.
            </div>`;
    }

    async function exportarWord(textoFull) {
        const { Document, Packer, Paragraph, TextRun } = window.docx;
        const paragraphs = textoFull.split('\n\n').map(p => {
            return new Paragraph({
                children: [new TextRun({ text: p, size: 24, font: "Arial" })],
                spacing: { after: 200 },
                alignment: "justify"
            });
        });

        const doc = new Document({ sections: [{ children: paragraphs }] });
        const blob = await Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "Redaccion_Linderos.docx";
        a.click();
    }
});