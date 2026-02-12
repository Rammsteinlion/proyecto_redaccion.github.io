document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const redaccionContainer = document.getElementById('redaccionTecnica');
    const botonContainer = document.getElementById('botonRedaccionContainer');
    const progressBar = document.getElementById('progressBar');
    const percentageText = document.getElementById('texto');

    let pointsDataMap = {}; 
    let boundariesData = [];

    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const data = await readExcel(file);
        const sheet = data["PREDIO"];

        if (sheet) {
            processHojaPredio(sheet);
            mostrarBotonGenerar();
        }
    });

    function processHojaPredio(data) {
        const headers = data[0].map(h => String(h || "").trim());
        const idIdx = headers.indexOf("Id"), yIdx = headers.indexOf("Y"), xIdx = headers.indexOf("X");
        const puntosRangeIdx = headers.indexOf("PUNTOS"), colindanteIdx = headers.indexOf("COLINDANTES");
        const distIdx = headers.indexOf("DISTANCIA (m)"), propIdx = headers.indexOf("PROPIETARIO");
        const fmiIdx = headers.indexOf("FMI"), nupreIdx = headers.indexOf("NUPRE");

        pointsDataMap = {};
        boundariesData = [];

        data.slice(1).forEach(row => {
            if (row[idIdx]) {
                pointsDataMap[String(row[idIdx]).trim()] = {
                    norte: parseFloat(String(row[yIdx]).replace(',', '.')),
                    este: parseFloat(String(row[xIdx]).replace(',', '.'))
                };
            }
            if (row[puntosRangeIdx]) {
                const numeros = String(row[puntosRangeIdx]).match(/\d+/g);
                if (numeros && numeros.length >= 2) {
                    boundariesData.push({
                        pInicio: numeros[0], pFin: numeros[1],
                        colindante: row[colindanteIdx], distancia: row[distIdx],
                        propietario: row[propIdx], fmi: row[fmiIdx], nupre: row[nupreIdx]
                    });
                }
            }
        });
    }

    function generarRedaccion() {
        let t = "LINDEROS TÉCNICOS\n\n";
        const orientaciones = ["POR EL NORTE:", "POR EL ESTE:", "POR EL SUR:", "POR EL OESTE:"];
        let ultimoColindante = "";
        let numLindero = 1;

        boundariesData.forEach((b, index) => {
            // Insertar encabezado de orientación (NORTE, ESTE, etc.)
            if (index < orientaciones.length) {
                t += `${orientaciones[index]}\n`;
            }

            const pI = pointsDataMap[b.pInicio], pF = pointsDataMap[b.pFin];
            const fmt = (n) => n.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

            // Lógica de Inicio de Lindero vs Continuación
            let esNuevoLindero = (b.colindante !== ultimoColindante);
            let inicioTexto = "";

            if (esNuevoLindero) {
                inicioTexto = `Lindero ${numLindero}: Inicia en el punto número ${b.pInicio}`;
                numLindero++;
            } else {
                inicioTexto = `Continúa en el punto número ${b.pInicio}`;
            }
            ultimoColindante = b.colindante;

            // Construcción del párrafo
            let parrafo = `${inicioTexto} de coordenadas planas N= ${fmt(pI.norte)}m, E= ${fmt(pI.este)}m, `;
            
            // Sentido automático simplificado (puedes editar esto)
            const sentido = index === 0 ? "Sureste" : (index === 1 ? "Sureste" : (index === 2 ? "Suroeste" : "Noroeste"));
            const esQuebrada = Math.abs(parseInt(b.pFin) - parseInt(b.pInicio)) > 1;
            
            parrafo += `en línea ${esQuebrada ? 'quebrada' : 'recta'} en sentido ${sentido}, `;

            if (b.distancia) parrafo += `con una distancia total acumulada de ${b.distancia}m, `;

            // Puntos intermedios
            if (esQuebrada) {
                let inters = [];
                const start = parseInt(b.pInicio), end = parseInt(b.pFin);
                const step = start < end ? 1 : -1;
                for (let i = start + step; i !== end; i += step) {
                    if (pointsDataMap[i]) {
                        inters.push(`el punto ${i} de coordenadas planas N= ${fmt(pointsDataMap[i].norte)}m, E= ${fmt(pointsDataMap[i].este)}m`);
                    }
                }
                if (inters.length > 0) parrafo += `pasando por ${inters.join(', ')}, `;
            }

            parrafo += `hasta encontrar el punto número ${b.pFin} de coordenadas planas N= ${fmt(pF.norte)}m, E= ${fmt(pF.este)}m`;

            // Colindancia (Solo se agrega al final del tramo del colindante o si el siguiente es distinto)
            const esUltimoTramoDeEsteColindante = (index === boundariesData.length - 1 || boundariesData[index + 1].colindante !== b.colindante);
            
            if (esUltimoTramoDeEsteColindante) {
                parrafo += ` colindando con un predio ${b.colindante}, el NUPRE Código predial ${b.nupre}, Folio de matrícula inmobiliaria ${b.fmi} y de propietario ${b.propietario}.`;
            } else {
                parrafo += ".";
            }

            t += parrafo + "\n\n";
        });

        redaccionContainer.textContent = t;
    }

    function mostrarBotonGenerar() {
        botonContainer.innerHTML = '';
        const btn = document.createElement('button');
        btn.className = "w-full py-3 bg-blue-600 text-white rounded-xl font-bold shadow-lg";
        btn.innerText = "Generar Linderos Organizados";
        btn.onclick = generarRedaccion;
        botonContainer.appendChild(btn);
    }

    async function readExcel(file) {
        return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const res = {};
                workbook.SheetNames.forEach(n => res[n.toUpperCase()] = XLSX.utils.sheet_to_json(workbook.Sheets[n], { header: 1 }));
                resolve(res);
            };
            reader.readAsArrayBuffer(file);
        });
    }
});