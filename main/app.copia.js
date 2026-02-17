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
            try {

                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                const res = {};
                workbook.SheetNames.forEach(n => {
                    res[n.toUpperCase()] = XLSX.utils.sheet_to_json(workbook.Sheets[n], { header: 1 });
                });

                const sheet = res["PREDIO"];

                if (sheet) {
                    processHojaPredio(sheet);
                    if (panelAcciones) panelAcciones.classList.remove('hidden');
                    mostrarBotonGenerar();
                } else {
                    alert("No se encontró la hoja 'PREDIO'");
                }

            } catch (err) {
                alert("Error al procesar el Excel");
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

    // ✅ SENTIDO CORRECTO
    function obtenerSentidoCartesiano(p1, p2) {

        const dN = p2.norte - p1.norte;
        const dE = p2.este - p1.este;

        if (dN > 0 && dE > 0) return "Noreste";
        if (dN > 0 && dE < 0) return "Noroeste";
        if (dN < 0 && dE > 0) return "Sureste";
        if (dN < 0 && dE < 0) return "Suroeste";

        if (dN > 0 && dE === 0) return "Norte";
        if (dN < 0 && dE === 0) return "Sur";
        if (dN === 0 && dE > 0) return "Este";
        if (dN === 0 && dE < 0) return "Oeste";

        return "Indeterminado";
    }

    function generarRedaccion() {

        let t = "";
        const f = (n) => Number(n).toFixed(2);

        t += "**LINDEROS TÉCNICOS**\n\n";

        boundariesData.forEach((b, bIdx) => {

            const step = b.pInicio < b.pFin ? 1 : -1;

            const pA = pointsDataMap[b.pInicio];
            const pB = pointsDataMap[b.pInicio + step];

            let orientacionReal = "POR EL NORTE:";

            if (pA && pB) {
                const sentidoInicial = obtenerSentidoCartesiano(pA, pB);

                if (sentidoInicial.includes("Norte")) orientacionReal = "POR EL NORTE:";
                else if (sentidoInicial.includes("Sur")) orientacionReal = "POR EL SUR:";
                else if (sentidoInicial.includes("Este")) orientacionReal = "POR EL ESTE:";
                else if (sentidoInicial.includes("Oeste")) orientacionReal = "POR EL OESTE:";
            }

            t += `**${orientacionReal}**\n\n`;

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
                    parrafo += `**Lindero ${bIdx + 1}:** Inicia en el punto número ${tr.inicio} de coordenadas planas N= ${f(pI.norte)}m, E= ${f(pI.este)}m, `;
                } else {
                    parrafo += `Continúa en el punto número ${tr.inicio} de coordenadas planas N= ${f(pI.norte)}m, E= ${f(pI.este)}m, `;
                }

                const esQuebrada = tr.puntos.length > 1;

                parrafo += `en línea ${esQuebrada ? 'quebrada' : 'recta'} en sentido ${tr.sentido}, `;

                if (index === tramos.length - 1 && b.distAcu) {
                    parrafo += `con una distancia total acumulada de ${Number(b.distAcu).toFixed(1)}m, `;
                }

                if (esQuebrada) {

                    const intermedios = tr.puntos.slice(0, -1).map((p, idx) => {
                        const label = idx === 0 ? "pasando por el punto número" : "el punto";
                        return `${label} ${p} de coordenadas planas N= ${f(pointsDataMap[p].norte)}m, E= ${f(pointsDataMap[p].este)}m`;
                    });

                    if (intermedios.length > 0) {
                        parrafo += `${intermedios.join(', ')}, `;
                    }
                }

                parrafo += `hasta encontrar el punto número ${tr.fin} de coordenadas planas N= ${f(pF.norte)}m, E= ${f(pF.este)}m`;

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
        btn.className = "w-full py-3 bg-blue-600 text-white rounded-xl font-bold mb-4 shadow-lg hover:bg-blue-700";
        btn.innerText = "Generar Redacción Técnica";
        btn.onclick = generarRedaccion;

        botonContainer.appendChild(btn);
    }

    function mostrarBotonDescarga(txt) {

        if (document.getElementById('btnWord')) return;

        const btn = document.createElement('button');
        btn.id = 'btnWord';
        btn.className = "w-full py-3 bg-green-600 text-white rounded-xl font-bold shadow-lg hover:bg-green-700";
        btn.innerHTML = 'Descargar Word (.doc)';
        btn.onclick = () => exportarAWord(txt);

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

            const partes = linea.split(/(\*\*.*?\*\*)/g);

            const children = partes.map(parte => {

                if (parte.startsWith("**") && parte.endsWith("**")) {
                    return new TextRun({
                        text: parte.replace(/\*\*/g, ''),
                        bold: true,
                        font: "Arial",
                        size: 22
                    });
                }

                return new TextRun({
                    text: parte,
                    font: "Arial",
                    size: 22
                });
            });

            return new Paragraph({
                children: children,
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
