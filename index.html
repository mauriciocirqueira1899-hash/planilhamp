<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Conversor Oficial Aghuse - Ministério Público</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        body { font-family: 'Segoe UI', system-ui, sans-serif; background-color: #f3f4f6; padding: 40px 20px; display: flex; justify-content: center; }
        .container { max-width: 720px; width: 100%; background: #ffffff; padding: 35px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.05); }
        h1 { color: #1e3a8a; font-size: 24px; margin-top: 0; }
        p.desc { color: #4b5563; font-size: 14px; margin-bottom: 25px; line-height: 1.5; }
        .dropzone { border: 2px dashed #3b82f6; padding: 50px 20px; border-radius: 8px; background: #f8fafc; text-align: center; cursor: pointer; transition: background .15s; }
        .dropzone:hover { background: #eff6ff; }
        .dropzone.dragover { background: #dbeafe; border-color: #1e3a8a; }
        .status-card { display: none; background-color: #f0f4f8; padding: 20px; border-radius: 8px; border-left: 6px solid #1e3a8a; margin-top: 25px; }
        .error-card { display: none; background-color: #fef2f2; padding: 20px; border-radius: 8px; border-left: 6px solid #dc2626; margin-top: 25px; color: #7f1d1d; }
        .btn-download { display: block; width: 100%; background-color: #1e3a8a; color: white; padding: 14px; border-radius: 6px; text-align: center; font-weight: bold; font-size: 16px; margin-top: 20px; cursor: pointer; border: none; }
        .btn-download:hover { background-color: #1d4ed8; }
        #loading { display: none; color: #1e3a8a; font-weight: bold; margin-top: 25px; text-align: center; }
        .progress-wrap { width: 100%; background: #e5e7eb; border-radius: 6px; margin-top: 10px; overflow: hidden; height: 10px; }
        .progress-bar { height: 100%; width: 0%; background: #1e3a8a; transition: width .1s; }
        .progress-txt { font-size: 12px; color: #6b7280; margin-top: 6px; font-weight: normal; }
        ul.rules { font-size: 13px; color: #4b5563; margin: 0 0 20px 0; padding-left: 18px; line-height: 1.6; }
    </style>
</head>
<body>
<div class="container">
    <h1>📊 Conversor de Relatórios Aghuse — MP</h1>
    <p class="desc">Envie o PDF do Aghuse (relatório de atendimentos). O sistema lê o conteúdo real do PDF diretamente no navegador (sem servidor, sem dependência externa de dados) e gera a planilha Excel no formato do relatório do Ministério Público.</p>
    <ul class="rules">
        <li>Cabeçalho de página (título da unidade e período) é descartado automaticamente</li>
        <li>Imagem/logo do relatório é ignorada</li>
        <li>A segunda tabela (resumo "Classificação de Risco x Total" no fim do PDF) é descartada</li>
        <li>A coluna <b>Idade</b> (após a coluna Cor) é removida</li>
        <li>Linhas sem diagnóstico preenchido são descartadas</li>
        <li>Nome do paciente é reduzido à letra inicial (padrão do relatório MP)</li>
    </ul>

    <div class="dropzone" id="dropzone" onclick="document.getElementById('fileInput').click()">
        <p>📥 Clique aqui ou arraste o arquivo PDF do Aghuse</p>
        <input type="file" id="fileInput" accept=".pdf" style="display:none">
    </div>
    <div id="loading">
        ⏳ Lendo e estruturando o relatório...
        <div class="progress-wrap"><div class="progress-bar" id="progressBar"></div></div>
        <div class="progress-txt" id="progressTxt">Página 0 de 0</div>
    </div>
    <div id="errorCard" class="error-card">
        <b>⚠️ Não foi possível processar o arquivo.</b>
        <p id="errorMsg" style="margin-bottom:0;"></p>
    </div>
    <div id="statusCard" class="status-card">
        <h3 style="margin-top:0; color:#1e3a8a;">✅ Processamento Concluído com Sucesso!</h3>
        <p>• <b>Páginas lidas:</b> <span id="lblPages"></span></p>
        <p>• <b>Total de registros válidos extraídos:</b> <span id="lblTotal" style="font-weight:bold; color:#16a34a;"></span> pacientes.</p>
        <p>• <b>Linhas descartadas (diagnóstico vazio):</b> <span id="lblSkipped"></span></p>
        <p>• <b>Nomes dos pacientes:</b> mantida apenas a primeira letra inicial.</p>
        <button id="btnDownload" class="btn-download">📥 Baixar Planilha Excel (.XLSX)</button>
    </div>
</div>
<script>
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

    const fileInput = document.getElementById('fileInput');
    const dropzone = document.getElementById('dropzone');
    const loading = document.getElementById('loading');
    const progressBar = document.getElementById('progressBar');
    const progressTxt = document.getElementById('progressTxt');
    const statusCard = document.getElementById('statusCard');
    const errorCard = document.getElementById('errorCard');
    const errorMsg = document.getElementById('errorMsg');
    const btnDownload = document.getElementById('btnDownload');
    let blobExcel = null;

    // ---- Configuração de layout do relatório Aghuse (JasperReports) ----
    // Limites de coluna (em pontos, espaço de visualização da página já rotacionada)
    const COL_BOUNDS = [20, 87, 154, 221, 288, 355, 422, 489, 556, 623, 690, 756, 822];
    const NUM_COLS = 12; // Data, Nome, Gravidade, Cidade, Classificação, Municipio, Especialidade, Diagnostico, Prontuario, Sexo, Cor, Idade
    const DATE_RE = /^\d{2}\/\d{2}\/\d{4}$/;
    const HEADER_FINAL = ["Data Atendimento", "Nome Paciente", "Gravidade", "Cidade de Origem", "Classificação de Risco", "Municipio", "Especialidade", "Diagnostico", "Prontuario", "Sexo", "Cor"];

    function colIndex(x) {
        for (let i = 0; i < NUM_COLS; i++) {
            if (x >= COL_BOUNDS[i] && x < COL_BOUNDS[i + 1]) return i;
        }
        return null;
    }

    dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('dragover'); });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone.addEventListener('drop', e => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            fileInput.dispatchEvent(new Event('change'));
        }
    });

    fileInput.addEventListener('change', async () => {
        if (!fileInput.files.length) return;
        errorCard.style.display = 'none';
        statusCard.style.display = 'none';
        loading.style.display = 'block';
        progressBar.style.width = '0%';
        blobExcel = null;

        try {
            const file = fileInput.files[0];
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            const numPages = pdf.numPages;

            const allRows = [];
            let skippedEmptyDiag = 0;

            for (let pageNum = 1; pageNum <= numPages; pageNum++) {
                progressTxt.textContent = `Página ${pageNum} de ${numPages}`;
                progressBar.style.width = ((pageNum / numPages) * 100).toFixed(0) + '%';

                const page = await pdf.getPage(pageNum);
                const viewport = page.getViewport({ scale: 1 });
                const textContent = await page.getTextContent();

                // Posiciona cada fragmento de texto no espaço visual da página (já considerando rotação)
                const words = [];
                for (const item of textContent.items) {
                    const str = (item.str || '').trim();
                    if (!str) continue;
                    const tx = pdfjsLib.Util.transform(viewport.transform, item.transform);
                    words.push({ text: str, x: tx[4], y: tx[5] });
                }

                // Âncoras de linha: valores de data isolados na coluna 0 (Data Atendimento)
                const anchors = words
                    .filter(w => DATE_RE.test(w.text) && colIndex(w.x) === 0)
                    .map(w => w.y)
                    .sort((a, b) => a - b);

                if (anchors.length === 0) continue; // página sem tabela de dados (ex: tabela-resumo "Table 2")

                // Limite inferior: rodapé do relatório ou início do resumo "Total de Registros"
                const stopWords = words.filter(w =>
                    w.text.includes('EMERG_GERA_RELATORIOS') || w.text.includes('Registros:')
                );
                const cutoff = stopWords.length ? Math.min(...stopWords.map(w => w.y)) : Infinity;
                anchors.push(cutoff);

                for (let ri = 0; ri < anchors.length - 1; ri++) {
                    const top = anchors[ri] - 1;
                    const bot = anchors[ri + 1] - 1;
                    const rowWords = words.filter(w => w.y >= top && w.y < bot);

                    const cols = Array.from({ length: NUM_COLS }, () => []);
                    for (const w of rowWords) {
                        const ci = colIndex(w.x);
                        if (ci !== null) cols[ci].push(w);
                    }

                    const cellVals = cols.map(c => {
                        c.sort((a, b) => (Math.round(a.y) - Math.round(b.y)) || (a.x - b.x));
                        return c.map(w => w.text).join(' ').trim();
                    });

                    // Descarta linhas incompletas (sem diagnóstico preenchido)
                    if (!cellVals[7]) {
                        skippedEmptyDiag++;
                        continue;
                    }

                    allRows.push(cellVals);
                }
            }

            if (allRows.length === 0) {
                throw new Error('Nenhum registro de atendimento foi encontrado neste PDF. Verifique se é um relatório Aghuse no formato esperado.');
            }

            // ---- Monta dados finais: iniciais do nome + remove coluna Idade (a última) ----
            const dadosFinais = [HEADER_FINAL];
            for (const row of allRows) {
                const nome = row[1] || '';
                const inicial = nome.length > 0 ? nome.trim().charAt(0).toUpperCase() : '';
                const prontuarioRaw = row[8];
                const prontuario = /^\d+$/.test(prontuarioRaw) ? Number(prontuarioRaw) : prontuarioRaw;

                dadosFinais.push([
                    row[0],       // Data Atendimento
                    inicial,      // Nome Paciente (apenas inicial)
                    row[2],       // Gravidade
                    row[3],       // Cidade de Origem
                    row[4],       // Classificação de Risco
                    row[5],       // Municipio
                    row[6],       // Especialidade
                    row[7],       // Diagnostico
                    prontuario,   // Prontuario
                    row[9],       // Sexo
                    row[10]       // Cor  (coluna Idade / row[11] é descartada)
                ]);
            }

            // ---- Gera planilha Excel ----
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(dadosFinais);

            ws['!rows'] = [];
            ws['!rows'][0] = { hpt: 26 };
            for (let r = 1; r < dadosFinais.length; r++) {
                ws['!rows'][r] = { hpt: 15 };
            }

            ws['!cols'] = [
                { wch: 18 }, { wch: 14 }, { wch: 12 }, { wch: 18 },
                { wch: 28 }, { wch: 14 }, { wch: 22 }, { wch: 45 },
                { wch: 12 }, { wch: 11 }, { wch: 10 }
            ];

            // Estilo (compatível com builds que suportam escrita de estilo; ignorado silenciosamente caso contrário)
            const thinBorder = { style: 'thin', color: { rgb: '000000' } };
            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = range.s.r; R <= range.e.r; R++) {
                for (let C = range.s.c; C <= range.e.c; C++) {
                    const addr = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!ws[addr]) continue;
                    ws[addr].s = {
                        font: { name: 'Times New Roman', sz: 10, bold: R === 0 },
                        alignment: { horizontal: R === 0 ? 'center' : 'left', vertical: 'top', wrapText: true },
                        border: { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder }
                    };
                }
            }

            XLSX.utils.book_append_sheet(wb, ws, 'Dados_Aghuse');
            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
            blobExcel = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

            document.getElementById('lblPages').textContent = numPages;
            document.getElementById('lblTotal').textContent = allRows.length;
            document.getElementById('lblSkipped').textContent = skippedEmptyDiag;

            loading.style.display = 'none';
            statusCard.style.display = 'block';
        } catch (err) {
            console.error(err);
            loading.style.display = 'none';
            errorMsg.textContent = err.message || 'Erro desconhecido ao processar o PDF.';
            errorCard.style.display = 'block';
        }
    });

    btnDownload.addEventListener('click', () => {
        if (!blobExcel) return;
        const url = window.URL.createObjectURL(blobExcel);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Relatorio_Aghuse_MP.xlsx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
    });
</script>
</body>
</html>
