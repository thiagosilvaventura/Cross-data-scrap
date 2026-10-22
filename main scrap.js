(async function() {
    // 1. Carrega SheetJS
    if (typeof XLSX === 'undefined') {
        await new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

    if (document.getElementById('sql-analyzer-panel')) {
        document.getElementById('sql-analyzer-panel').remove();
    }

    // VARIÁVEIS DE ESTADO MOVIDAS PARA O TOPO (Correção do Bug)
    let datasets = {};
    let finalResult = [];
    let currentLang = 'pt';

    const i18n = {
        pt: { title: "🔗 SQL Joiner Pro", step1: "1. Importar Bases (Tabelas)", noFiles: "Nenhum arquivo carregado.", filesLoaded: "{n} tabelas carregadas.", step2: "2. Relacionamento (PK / FK)", tableA: "Tabela A (Principal)", pkA: "Chave Primária (PK)", tableB: "Tabela B (Estrangeira)", fkB: "Chave Estrangeira (FK)", step3: "3. Tipo de JOIN", leftJoin: "LEFT JOIN", leftDesc: "Mantém <b>todos</b> os registros da Tabela A.", innerJoin: "INNER JOIN", innerDesc: "Mantém <b>apenas</b> registros com chaves exatas em ambas as tabelas.", fullJoin: "FULL JOIN", fullDesc: "Mantém <b>todos</b> os registros de ambas as tabelas.", runBtn: "Executar JOIN", clearBtn: "Limpar Tudo", processing: "Processando...", resultTitle: "✅ Query Executada", downloadBtn: "Baixar Tabela Resultante" },
        en: { title: "🔗 SQL Joiner Pro", step1: "1. Import Datasets (Tables)", noFiles: "No files loaded.", filesLoaded: "{n} tables loaded.", step2: "2. Relationship (PK / FK)", tableA: "Table A (Main)", pkA: "Primary Key (PK)", tableB: "Table B (Foreign)", fkB: "Foreign Key (FK)", step3: "3. JOIN Type", leftJoin: "LEFT JOIN", leftDesc: "Keeps <b>all</b> records from Table A.", innerJoin: "INNER JOIN", innerDesc: "Keeps <b>only</b> records with exact key matches.", fullJoin: "FULL JOIN", fullDesc: "Keeps <b>all</b> records from both tables.", runBtn: "Execute JOIN", clearBtn: "Clear All", processing: "Processing...", resultTitle: "✅ Query Executed", downloadBtn: "Download Resulting Table" }
    };

    // 2. CSS Animado
    const style = document.createElement('style');
    style.innerHTML = `
        @keyframes gradientMove { 0% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } 100% { background-position: 0% 50%; } }
        #sql-analyzer-panel { position: fixed; top: 20px; right: 20px; z-index: 999999; background: linear-gradient(135deg, rgba(20,20,30,0.95), rgba(40,30,60,0.95), rgba(20,40,50,0.95)); background-size: 200% 200%; animation: gradientMove 10s ease infinite; backdrop-filter: blur(12px); color: #e0e0e0; border-radius: 12px; font-family: 'Segoe UI', Tahoma, sans-serif; box-shadow: 0 10px 40px rgba(0,0,0,0.7); width: 500px; min-width: 350px; min-height: 200px; max-height: 90vh; display: flex; flex-direction: column; border: 1px solid rgba(255,255,255,0.15); resize: both; overflow: hidden; transition: width 0.3s, height 0.3s, top 0.3s, left 0.3s; }
        #sql-header { padding: 15px; background: rgba(0,0,0,0.3); border-bottom: 1px solid rgba(255,255,255,0.1); display: flex; justify-content: space-between; align-items: center; cursor: grab; user-select: none; }
        #sql-header:active { cursor: grabbing; }
        .sql-controls button { background: none; border: none; color: #aaa; cursor: pointer; font-size: 16px; margin-left: 8px; transition: color 0.2s; }
        .sql-controls button:hover { color: #fff; }
        #sql-body { padding: 20px; overflow-y: auto; flex: 1; }
        .sql-section { background: rgba(0,0,0,0.2); border: 1px solid rgba(255,255,255,0.05); padding: 15px; border-radius: 8px; margin-bottom: 15px; }
        .sql-title { margin: 0 0 15px 0; color: #5bc0de; font-size: 15px; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 8px; }
        .sql-form-group { display: flex; flex-direction: column; gap: 5px; margin-bottom: 10px; }
        .sql-select { padding: 8px; border-radius: 6px; background: rgba(0,0,0,0.5); color: #fff; border: 1px solid #555; outline: none; transition: 0.2s; }
        .sql-select:focus { border-color: #5bc0de; }
        .sql-btn { background: #5bc0de; color: #111; border: none; padding: 10px; border-radius: 6px; cursor: pointer; font-weight: bold; width: 100%; transition: 0.2s; margin-top: 5px;}
        .sql-btn:hover { background: #46b8da; transform: translateY(-1px); }
        .sql-btn.clear { background: #d9534f; color: white; }
        .sql-join-card { display: flex; align-items: flex-start; gap: 10px; padding: 10px; background: rgba(0,0,0,0.3); border: 1px solid #444; border-radius: 6px; margin-bottom: 8px; cursor: pointer; transition: 0.2s; }
        .sql-join-card:hover { border-color: #5bc0de; background: rgba(0,0,0,0.5); }
        .join-desc { font-size: 11px; color: #bbb; margin-top: 4px; line-height: 1.4; }
        .join-title { font-size: 13px; font-weight: bold; color: #fff; }
    `;
    document.head.appendChild(style);

    // 3. Estrutura da Interface
    const panel = document.createElement('div');
    panel.id = 'sql-analyzer-panel';
    panel.innerHTML = `
        <div id="sql-header">
            <h2 style="margin:0; font-size:16px; font-weight:bold;" data-i18n="title">🔗 SQL Joiner Pro</h2>
            <div class="sql-controls">
                <button id="sql-lang" title="Language">🇧🇷 PT</button>
                <button id="sql-min" title="Minimize">_</button>
                <button id="sql-max" title="Maximize">□</button>
                <button id="sql-close" title="Close">✖</button>
            </div>
        </div>
        <div id="sql-body">
            <div class="sql-section">
                <h3 class="sql-title" data-i18n="step1">1. Importar Bases (Tabelas)</h3>
                <input type="file" id="sql-files" multiple accept=".csv, .xls, .xlsx" style="width:100%; margin-bottom:10px;">
                <div id="file-status" style="font-size:12px; color:#aaa;" data-i18n="noFiles">Nenhum arquivo carregado.</div>
            </div>

            <div id="sql-query-builder" style="display:none; pointer-events:none; opacity:0.5; transition: 0.3s;">
                <div class="sql-section">
                    <h3 class="sql-title" data-i18n="step2">2. Relacionamento (PK / FK)</h3>
                    <div style="display:flex; gap:10px;">
                        <div style="flex:1;">
                            <div class="sql-form-group"><label style="font-size:12px; color:#5bc0de;" data-i18n="tableA">Tabela A</label><select class="sql-select" id="table-a"></select></div>
                            <div class="sql-form-group"><label style="font-size:12px;" data-i18n="pkA">Chave (PK)</label><select class="sql-select" id="pk-a"></select></div>
                        </div>
                        <div style="flex:1;">
                            <div class="sql-form-group"><label style="font-size:12px; color:#f0ad4e;" data-i18n="tableB">Tabela B</label><select class="sql-select" id="table-b"></select></div>
                            <div class="sql-form-group"><label style="font-size:12px;" data-i18n="fkB">Chave (FK)</label><select class="sql-select" id="fk-b"></select></div>
                        </div>
                    </div>
                </div>

                <div class="sql-section">
                    <h3 class="sql-title" data-i18n="step3">3. Tipo de JOIN</h3>
                    <label class="sql-join-card"><input type="radio" name="join-type" value="LEFT" checked><div><div class="join-title" data-i18n="leftJoin">LEFT JOIN</div><div class="join-desc" data-i18n="leftDesc">...</div></div></label>
                    <label class="sql-join-card"><input type="radio" name="join-type" value="INNER"><div><div class="join-title" data-i18n="innerJoin">INNER JOIN</div><div class="join-desc" data-i18n="innerDesc">...</div></div></label>
                    <label class="sql-join-card"><input type="radio" name="join-type" value="FULL"><div><div class="join-title" data-i18n="fullJoin">FULL JOIN</div><div class="join-desc" data-i18n="fullDesc">...</div></div></label>
                </div>
                <div style="display:flex; gap:10px;">
                    <button id="run-sql" class="sql-btn" style="flex: 2;" data-i18n="runBtn">Executar JOIN</button>
                    <button id="clear-btn" class="sql-btn clear" style="flex: 1;" data-i18n="clearBtn">Limpar</button>
                </div>
                <div id="sql-output" style="margin-top:15px; font-size:13px; display:none;"></div>
            </div>
        </div>
    `;
    document.body.appendChild(panel);

    // 4. Lógica de UI (Drag, Min/Max)
    const header = document.getElementById('sql-header');
    const body = document.getElementById('sql-body');
    let isDragging = false, startX, startY, initialLeft, initialTop, isMinimized = false, isMaximized = false, cachedRect = {};

    header.onmousedown = (e) => { if (e.target.tagName === 'BUTTON' || isMaximized) return; isDragging = true; startX = e.clientX; startY = e.clientY; initialLeft = panel.offsetLeft; initialTop = panel.offsetTop; panel.style.right = 'auto'; };
    document.onmousemove = (e) => { if (!isDragging) return; panel.style.left = `${initialLeft + e.clientX - startX}px`; panel.style.top = `${initialTop + e.clientY - startY}px`; };
    document.onmouseup = () => isDragging = false;
    document.getElementById('sql-close').onclick = () => panel.remove();
    document.getElementById('sql-min').onclick = () => { if(isMinimized) { panel.style.height = cachedRect.height || 'auto'; body.style.display = 'block'; document.getElementById('sql-min').innerText = '_'; } else { cachedRect.height = panel.style.height; panel.style.height = '50px'; body.style.display = 'none'; document.getElementById('sql-min').innerText = '⎕'; } isMinimized = !isMinimized; };
    document.getElementById('sql-max').onclick = () => { if(isMaximized) { panel.style.width = cachedRect.width; panel.style.height = cachedRect.height; panel.style.top = cachedRect.top; panel.style.left = cachedRect.left; } else { cachedRect = { width: panel.style.width || panel.offsetWidth+'px', height: panel.style.height || panel.offsetHeight+'px', top: panel.style.top || panel.offsetTop+'px', left: panel.style.left || panel.offsetLeft+'px' }; panel.style.width = '100vw'; panel.style.height = '100vh'; panel.style.top = '0'; panel.style.left = '0'; } isMaximized = !isMaximized; };

    // 5. Lógica de Tradução e Estado
    const translate = () => {
        document.querySelectorAll('[data-i18n]').forEach(el => {
            // Evita sobrescrever o texto de sucesso se arquivos já foram carregados
            if (el.id === 'file-status' && Object.keys(datasets).length > 0) return;
            el.innerHTML = i18n[currentLang][el.getAttribute('data-i18n')];
        });
        document.getElementById('sql-lang').innerText = currentLang === 'pt' ? '🇧🇷 PT' : '🇺🇸 EN';
        if(Object.keys(datasets).length > 0) document.getElementById('file-status').innerText = i18n[currentLang].filesLoaded.replace('{n}', Object.keys(datasets).length);
    };
    
    document.getElementById('sql-lang').onclick = () => { currentLang = currentLang === 'pt' ? 'en' : 'pt'; translate(); };
    
    document.getElementById('clear-btn').onclick = () => {
        document.getElementById('sql-files').value = ''; datasets = {}; finalResult = [];
        document.getElementById('file-status').innerText = i18n[currentLang].noFiles;
        document.getElementById('sql-query-builder').style.opacity = '0.5'; document.getElementById('sql-query-builder').style.pointerEvents = 'none'; document.getElementById('sql-output').style.display = 'none';
    };

    translate();

    // 6. Processamento de Arquivos
    document.getElementById('sql-files').onchange = async (e) => {
        const files = e.target.files;
        if (files.length < 2) return alert(currentLang === 'pt' ? "Selecione pelo menos 2 tabelas" : "Select at least 2 tables");
        
        document.getElementById('file-status').innerText = i18n[currentLang].processing;
        datasets = {};

        try {
            for (let file of files) {
                const data = await file.arrayBuffer();
                const workbook = XLSX.read(data, { type: 'array' });
                const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
                datasets[file.name] = { name: file.name, data: json, columns: Object.keys(json[0] || {}) };
            }

            const options = Object.keys(datasets).map(n => `<option value="${n}">${n}</option>`).join('');
            document.getElementById('table-a').innerHTML = options; document.getElementById('table-b').innerHTML = options; document.getElementById('table-b').selectedIndex = 1;
            updateCols();

            document.getElementById('file-status').innerText = i18n[currentLang].filesLoaded.replace('{n}', files.length);
            const builder = document.getElementById('sql-query-builder');
            builder.style.pointerEvents = 'all'; builder.style.opacity = '1'; builder.style.display = 'block';
        } catch (err) {
            document.getElementById('file-status').innerText = "Erro ao processar arquivos.";
            console.error(err);
        }
    };

    function updateCols() {
        const tA = document.getElementById('table-a').value, tB = document.getElementById('table-b').value;
        if(!datasets[tA] || !datasets[tB]) return;
        document.getElementById('pk-a').innerHTML = datasets[tA].columns.map(c => `<option value="${c}">${c}</option>`).join('');
        document.getElementById('fk-b').innerHTML = datasets[tB].columns.map(c => `<option value="${c}">${c}</option>`).join('');
    }
    document.getElementById('table-a').onchange = updateCols; document.getElementById('table-b').onchange = updateCols;

    // 7. Motor Relacional
    document.getElementById('run-sql').onclick = () => {
        const tA = datasets[document.getElementById('table-a').value].data, tB = datasets[document.getElementById('table-b').value].data;
        const pk = document.getElementById('pk-a').value, fk = document.getElementById('fk-b').value;
        const joinType = document.querySelector('input[name="join-type"]:checked').value;
        const out = document.getElementById('sql-output');
        out.style.display = "block"; out.innerHTML = i18n[currentLang].processing;

        setTimeout(() => {
            finalResult = []; let mapB = {};
            tB.forEach(r => { let k = String(r[fk]).trim(); if(!mapB[k]) mapB[k]=[]; mapB[k].push(r); });
            let matchedBKeys = new Set();
            
            tA.forEach(rowA => {
                let key = String(rowA[pk]).trim();
                if (mapB[key]) {
                    mapB[key].forEach(rowB => { let merged = { ...rowA }; for(let col in rowB) if (col !== fk) merged[merged.hasOwnProperty(col) ? `B_${col}` : col] = rowB[col]; finalResult.push(merged); matchedBKeys.add(key); });
                } else if (joinType !== 'INNER') finalResult.push({ ...rowA });
            });

            if (joinType === 'FULL') {
                tB.forEach(rowB => {
                    let key = String(rowB[fk]).trim();
                    if (!matchedBKeys.has(key)) {
                        let standalone = {};
                        for(let col in rowB) standalone[col !== fk && datasets[document.getElementById('table-a').value].columns.includes(col) ? `B_${col}` : col] = rowB[col];
                        finalResult.push(standalone);
                    }
                });
            }

            out.innerHTML = `
                <div style="background:rgba(0,0,0,0.4); padding:15px; border-radius:8px; border: 1px solid rgba(255,255,255,0.1);">
                    <h4 style="margin:0 0 10px 0; color:#5bc0de;" data-i18n="resultTitle">${i18n[currentLang].resultTitle}</h4>
                    <ul style="margin:0; padding-left:20px; font-size:12px; color:#ccc;"><li><b>Join:</b> ${joinType}</li><li><b>A:</b> ${tA.length} rows</li><li><b>B:</b> ${tB.length} rows</li><li><b style="color:#5bc0de">Total:</b> ${finalResult.length} rows</li></ul>
                    <button id="download-sql" class="sql-btn" style="background:#5cb85c; margin-top:15px;" data-i18n="downloadBtn">${i18n[currentLang].downloadBtn}</button>
                </div>
            `;
            document.getElementById('download-sql').onclick = () => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(finalResult), "Result"); XLSX.writeFile(wb, `Query_${joinType}.xlsx`); };
        }, 100);
    };
})();
