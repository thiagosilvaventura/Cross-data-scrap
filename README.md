<div align="center">
  <h1>🚀 Cross data Scrap</h1>
  <p><b>Web Data Cross-Referencing Tool</b></p>
  <img src="https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black" alt="JavaScript" />
  <img src="https://img.shields.io/badge/SheetJS-00A65A?style=for-the-badge&logo=spreadsheet&logoColor=white" alt="SheetJS" />
</div>

<br>

Inject a dynamic, floating analysis interface directly into any web page[cite: 1]. Upload and cross-reference multiple datasets, map common variables, identify primary keys, and generate instant variable presence reports[cite: 1].

## ✨ Key Features

| 🌟 Feature | 🛠️ Description |
| :--- | :--- |
| 🪟 **Floating UI** | Injects a fixed, dark-mode panel (`#data-analyzer-panel`) into the DOM[cite: 1]. |
| ⚡ **Dynamic Dependencies** | Automatically loads the `SheetJS` library on the fly[cite: 1]. |
| 🔑 **Auto Key Detection** | Identifies common columns across uploaded files and sets the first as the primary key[cite: 1]. |
| 📈 **Variable Mapping** | Calculates the percentage of appearance for every column across all databases[cite: 1]. |
| 🧬 **Record Extraction** | Collects unique keys and logs the array directly to the console[cite: 1]. |

## 🚀 How to Use

1. Copy the source code below.
2. Open your browser's Developer Tools (`F12` or `Ctrl+Shift+I`) and navigate to the **Console**.
3. Paste the code and press `Enter`.
4. In the floating panel, select at least 2 (ideally 4 or more) `.csv, .xls, or .xlsx` files[cite: 1].
5. Click **"Processar Análise"** to generate your report[cite: 1].

## 💻 Source Code

```javascript
(async function() {
    // 1. Carrega dinamicamente a biblioteca SheetJS para leitura de planilhas
    if (typeof XLSX === 'undefined') {
        await new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = '[https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js](https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js)';
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

    // 2. Injeta uma interface visual amigável na página
    if (document.getElementById('data-analyzer-panel')) {
        document.getElementById('data-analyzer-panel').remove();
    }

    const panel = document.createElement('div');
    panel.id = 'data-analyzer-panel';
    panel.innerHTML = `
        <div style="position:fixed;top:20px;right:20px;z-index:999999;background:#1e1e1e;color:#fff;padding:20px;border-radius:8px;font-family:sans-serif;box-shadow:0 4px 20px rgba(0,0,0,0.6);width:420px;max-height:85vh;overflow-y:auto;">
            <h3 style="margin-top:0;color:#4af;font-size:16px;">Analisador de Bases de Dados</h3>
            <p style="font-size:12px;color:#ccc;">Selecione 4 ou mais arquivos (XLS, XLSX, CSV):</p>
            <input type="file" id="db-files" multiple accept=".csv, .xls, .xlsx" style="margin-bottom:12px;width:100%;font-size:12px;"><br>
            <button id="run-analysis" style="background:#007acc;color:white;border:none;padding:10px;border-radius:4px;cursor:pointer;width:100%;font-weight:bold;">Processar Análise</button>
            <div id="analysis-output" style="margin-top:15px;font-size:12px;white-space:pre-wrap;background:#2d2d2d;padding:10px;border-radius:4px;max-height:300px;overflow-y:auto;color:#d4d4d4;">Aguardando arquivos...</div>
        </div>
    `;
    document.body.appendChild(panel);

    // 3. Lógica de processamento ao clicar no botão
    document.getElementById('run-analysis').onclick = async () => {
        const files = document.getElementById('db-files').files;
        const output = document.getElementById('analysis-output');
        
        if (files.length < 2) {
            output.innerText = 'Por favor, selecione pelo menos 2 arquivos (idealmente 4 ou mais).';
            return;
        }

        output.innerText = 'Lendo e processando arquivos...';
        let datasets = [];

        try {
            for (let file of files) {
                const data = await file.arrayBuffer();
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const json = XLSX.utils.sheet_to_json(worksheet);
                datasets.push({ name: file.name, data: json });
            }

            // Identifica variáveis (colunas) em comum entre todas as bases
            let commonCols = Object.keys(datasets[0].data[0] || {});
            for (let i = 1; i < datasets.length; i++) {
                const cols = Object.keys(datasets[i].data[0] || {});
                commonCols = commonCols.filter(c => cols.includes(c));
            }

            if (commonCols.length === 0) {
                output.innerText = 'Erro: Nenhuma coluna/variável em comum encontrada entre todas as bases!';
                return;
            }

            // Seleciona automaticamente a primeira coluna comum como chave primária
            const keyCol = commonCols[0];

            // Coleta valores únicos da chave em todas as bases
            let uniqueKeys = new Set();
            datasets.forEach(ds => {
                ds.data.forEach(row => {
                    if (row[keyCol] !== undefined && row[keyCol] !== null) {
                        uniqueKeys.add(String(row[keyCol]).trim());
                    }
                });
            });

            // Mapeia todas as variáveis existentes em todas as bases combinadas
            let allColumns = new Set();
            datasets.forEach(ds => {
                ds.data.forEach(row => {
                    Object.keys(row).forEach(k => allColumns.add(k));
                });
            });

            // Monta o relatório de percentual de presença de cada variável entre as bases
            let report = `=== RELATÓRIO DE ANÁLISE ===\n`;
            report += `Bases Analisadas: ${datasets.length}\n`;
            report += `Chave Comum Utilizada: "${keyCol}"\n`;
            report += `Total de Registros Únicos (Chave): ${uniqueKeys.size}\n\n`;
            report += `--- % de Aparição das Variáveis por Base ---\n`;

            allColumns.forEach(col => {
                let countInDatasets = datasets.filter(ds => {
                    const cols = Object.keys(ds.data[0] || {});
                    return cols.includes(col);
                }).length;
                let percentage = (countInDatasets / datasets.length) * 100;
                report += `• ${col}: ${percentage.toFixed(1)}% (${countInDatasets}/${datasets.length} bases)\n`;
            });

            output.innerText = report;
            console.log("--- CHAVES ÚNICAS ENCONTRADAS ---");
            console.log(Array.from(uniqueKeys));
            console.log(report);

        } catch (error) {
            output.innerText = 'Erro ao processar os arquivos: ' + error.message;
            console.error(error);
        }
    };
})();# Cross-data-scrap
