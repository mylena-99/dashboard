// Estados Globais de Aplicação
let rawData = [];         // Dados puros extraídos da planilha
let filteredData = [];    // Dados após processamento de filtros
let columnsConfig = { date: '', category: '', product: '', sector: '', numeric: '' };

// Instâncias Globais dos Gráficos (ChartJS)
let charts = { linha: null, pizza: null, barras: null, area: null, doughnut: null };

// Controle de Paginação de Tabela
let currentPage = 1;
const recordsPerPage = 10;
let sortDirection = false;

// Inicializadores Iniciais do Escopo do Navegador
document.addEventListener("DOMContentLoaded", () => {
    initTheme();
    setupEventListeners();
});

// Listener de Controle de Componentes
function setupEventListeners() {
    document.getElementById('menu-toggle').addEventListener('click', () => {
        document.getElementById('wrapper').classList.toggle('toggled');
    });

    document.getElementById('excel-file').addEventListener('change', handleExcelUpload);
    document.getElementById('theme-toggle').addEventListener('click', toggleTheme);

    // Conectores dos inputs de filtros para atualizar dinamicamente em tempo de execução
    document.getElementById('filter-search').addEventListener('input', applyFilters);
    document.getElementById('filter-category').addEventListener('change', applyFilters);
    document.getElementById('filter-setor').addEventListener('change', applyFilters);
    document.getElementById('filter-produto').addEventListener('change', applyFilters);
    document.getElementById('filter-date-start').addEventListener('change', applyFilters);
    document.getElementById('filter-date-end').addEventListener('change', applyFilters);

    // Botões de Paginação
    document.getElementById('btn-prev-page').addEventListener('click', () => { if (currentPage > 1) { currentPage--; displayTable(); } });
    document.getElementById('btn-next-page').addEventListener('click', () => { if (currentPage * recordsPerPage < filteredData.length) { currentPage++; displayTable(); } });

    // Exportação em PDF usando jsPDF + AutoTable
    document.getElementById('btn-pdf').addEventListener('click', exportDashboardPDF);
}

// Handler de Input e Upload da Planilha .XLSX
function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById('sheet-title').innerText = "Processando arquivo corporativo...";
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        
        // Obtém dados da primeira aba da planilha
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Converte em JSON Estruturado de Objetos
        rawData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        if(rawData.length > 0) {
            document.getElementById('sheet-title').innerText = `Base: ${file.name}`;
            analyzeAndMapColumns(rawData[0]);
            enableDashboardControls();
            applyFilters();
        } else {
            alert("A planilha processada está vazia!");
        }
    };
    reader.readAsArrayBuffer(file);
}

// Inteligência Artificial / Algoritmo de Mapeamento Heurístico de Colunas
function analyzeAndMapColumns(sampleRow) {
    const keys = Object.keys(sampleRow);
    
    // Reset de configurações
    columnsConfig = { date: '', category: '', product: '', sector: '', numeric: '' };
    let numericFields = [];
    let textFields = [];

    keys.forEach(key => {
        const val = sampleRow[key];
        
        // Detecção de campos de data
        if(val instanceof Date || (!isNaN(Date.parse(val)) && isNaN(val) && String(val).includes('-'))) {
            if(!columnsConfig.date) columnsConfig.date = key;
        }
        // Detecção de campos puramente numéricos
        else if(typeof val === 'number' || (!isNaN(parseFloat(val)) && isFinite(val))) {
            numericFields.push(key);
        } else {
            textFields.push(key);
        }
    });

    // Mapeamento prioritário do campo de valor base numérico para os KPIs
    columnsConfig.numeric = numericFields.find(f => f.toLowerCase().includes('valor') || f.toLowerCase().includes('total') || f.toLowerCase().includes('preço') || f.toLowerCase().includes('faturamento')) || numericFields[0] || '';
    
    // Mapeamento categórico cruzado
    columnsConfig.category = textFields.find(f => f.toLowerCase().includes('cat') || f.toLowerCase().includes('grupo')) || textFields[0] || '';
    columnsConfig.product = textFields.find(f => f.toLowerCase().includes('prod') || f.toLowerCase().includes('item')) || textFields[1] || textFields[0] || '';
    columnsConfig.sector = textFields.find(f => f.toLowerCase().includes('setor') || f.toLowerCase().includes('regiao') || f.toLowerCase().includes('loja')) || textFields[2] || textFields[0] || '';
}

// Ativação dos Inputs Bloqueados da UI
function enableDashboardControls() {
    document.getElementById('filter-search').disabled = false;
    document.getElementById('filter-category').disabled = false;
    document.getElementById('filter-setor').disabled = false;
    document.getElementById('filter-produto').disabled = false;
    document.getElementById('filter-date-start').disabled = false;
    document.getElementById('filter-date-end').disabled = false;
    document.getElementById('btn-pdf').disabled = false;

    populateSelectFilters();
}

// Alimentação Dinâmica dos Filtros de Seleção Multi-Escolha baseados nos dados lidos
function populateSelectFilters() {
    populateSingleSelect('filter-category', columnsConfig.category);
    populateSingleSelect('filter-setor', columnsConfig.sector);
    populateSingleSelect('filter-produto', columnsConfig.product);
}

function populateSingleSelect(elementId, columnKey) {
    const select = document.getElementById(elementId);
    select.innerHTML = '<option value="">Todos</option>';
    if(!columnKey) return;

    const uniqueValues = [...new Set(rawData.map(item => item[columnKey]))].filter(Boolean);
    uniqueValues.sort().forEach(val => {
        const opt = document.createElement('option');
        opt.value = val;
        opt.innerText = val;
        select.appendChild(opt);
    });
}

// Aplicação de Malha de Filtros Multidirecional
function applyFilters() {
    const searchVal = document.getElementById('filter-search').value.toLowerCase();
    const catVal = document.getElementById('filter-category').value;
    const setorVal = document.getElementById('filter-setor').value;
    const prodVal = document.getElementById('filter-produto').value;
    const dateStart = document.getElementById('filter-date-start').value;
    const dateEnd = document.getElementById('filter-date-end').value;

    filteredData = rawData.filter(row => {
        // Filtro Global de Pesquisa Textual
        const matchesSearch = !searchVal || Object.values(row).some(v => String(v).toLowerCase().includes(searchVal));
        
        // Filtros estruturados das colunas inferidas
        const matchesCat = !catVal || String(row[columnsConfig.category]) === catVal;
        const matchesSetor = !setorVal || String(row[columnsConfig.sector]) === setorVal;
        const matchesProd = !prodVal || String(row[columnsConfig.product]) === prodVal;
        
        // Filtro por Linha Temporal (Data)
        let matchesDate = true;
        if(columnsConfig.date && row[columnsConfig.date]) {
            const rowDate = new Date(row[columnsConfig.date]).toISOString().split('T')[0];
            if(dateStart && rowDate < dateStart) matchesDate = false;
            if(dateEnd && rowDate > dateEnd) matchesDate = false;
        }

        return matchesSearch && matchesCat && matchesSetor && matchesProd && matchesDate;
    });

    currentPage = 1;
    calculateKPIs();
    renderCharts();
    setupTableHeader();
    displayTable();
}

// Processamento Analítico dos Elementos Estatísticos (KPI)
function calculateKPIs() {
    const totalRecords = filteredData.length;
    let sumTotal = 0;
    let sumQty = 0;
    let valuesArray = [];

    filteredData.forEach(row => {
        // Soma quantitativa genérica (procura coluna de quantidade ou assume 1 por registro)
        const qtyKey = Object.keys(row).find(k => k.toLowerCase().includes('qtd') || k.toLowerCase().includes('quant'));
        sumQty += qtyKey ? (Number(row[qtyKey]) || 0) : 1;

        // Análise de faturamento baseado na coluna numérica inferida
        if(columnsConfig.numeric) {
            const val = Number(row[columnsConfig.numeric]) || 0;
            sumTotal += val;
            valuesArray.push(val);
        }
    });

    const maxVal = valuesArray.length ? Math.max(...valuesArray) : 0;
    const minVal = valuesArray.length ? Math.min(...valuesArray) : 0;
    const media = totalRecords ? (sumTotal / totalRecords) : 0;

    // Atualização de nós HTML na View dos KPIs
    document.getElementById('kpi-qtd-total').innerText = sumQty.toLocaleString('pt-BR');
    document.getElementById('kpi-valor-total').innerText = formatCurrency(sumTotal);
    document.getElementById('kpi-media').innerText = formatCurrency(media);
    document.getElementById('kpi-maior').innerText = formatCurrency(maxVal);
    document.getElementById('kpi-menor').innerText = formatCurrency(minVal);
    document.getElementById('kpi-registros').innerText = totalRecords.toLocaleString('pt-BR');
}

// Renderização Inteligente e Dinâmica dos Gráficos com ChartJS
function renderCharts() {
    // Coleta as cores de layout conforme o tema ativo
    const isDark = document.body.classList.contains('dark-mode');
    const textThemeColor = isDark ? '#94a3b8' : '#64748b';

    // Agrupamento Categórico para os Gráficos
    const catGroup = aggregateData(columnsConfig.category, columnsConfig.numeric);
    const prodGroup = aggregateData(columnsConfig.product, columnsConfig.numeric);
    const dateGroup = aggregateData(columnsConfig.date, columnsConfig.numeric, true);

    const paletteColors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#14b8a6'];

    // Destruição recursiva de instâncias prévias para evitar vazamentos de memória na atualização
    Object.keys(charts).forEach(key => { if(charts[key]) charts[key].destroy(); });

    // Configuração do Eixo Base Universal
    const chartConfigBase = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { labels: { color: textThemeColor, font: { family: 'Inter' } } } }
    };

    // 1. Gráfico de Linha (Evolução por Data)
    charts.linha = new Chart(document.getElementById('chart-linha'), {
        type: 'line',
        data: {
            labels: Object.keys(dateGroup),
            datasets: [{ label: 'Faturamento Cronológico', data: Object.values(dateGroup), borderColor: '#3b82f6', tension: 0.2, fill: false, pointBackgroundColor: '#3b82f6' }]
        },
        options: chartConfigBase
    });

    // 2. Gráfico de Pizza (Faturamento de Categorias)
    charts.pizza = new Chart(document.getElementById('chart-pizza'), {
        type: 'pizza',
        data: {
            labels: Object.keys(catGroup).slice(0, 7),
            datasets: [{ data: Object.values(catGroup).slice(0, 7), backgroundColor: paletteColors }]
        },
        options: chartConfigBase
    });

    // 3. Gráfico de Barras (Top Produtos)
    charts.barras = new Chart(document.getElementById('chart-barras'), {
        type: 'bar',
        data: {
            labels: Object.keys(prodGroup).slice(0, 8),
            datasets: [{ label: 'Performance Interna por Volume Monetário', data: Object.values(prodGroup).slice(0, 8), backgroundColor: '#f59e0b' }]
        },
        options: chartConfigBase
    });

    // 4. Gráfico de Área (Faturamento Acumulado)
    charts.area = new Chart(document.getElementById('chart-area'), {
        type: 'line',
        data: {
            labels: Object.keys(dateGroup),
            datasets: [{ label: 'Faturamento', data: Object.values(dateGroup), borderColor: '#10b981', backgroundColor: 'rgba(16, 185, 129, 0.15)', fill: true }]
        },
        options: chartConfigBase
    });

    // 5. Gráfico Doughnut (Mix Distribuição Estrutural)
    charts.doughnut = new Chart(document.getElementById('chart-doughnut'), {
        type: 'doughnut',
        data: {
            labels: Object.keys(catGroup).slice(0, 5),
            datasets: [{ data: Object.values(catGroup).slice(0, 5), backgroundColor: paletteColors.slice().reverse() }]
        },
        options: chartConfigBase
    });
}

// Estruturação da Tabela de Registros
function setupTableHeader() {
    const tr = document.getElementById('table-header');
    tr.innerHTML = '';
    if(!filteredData.length) return;

    Object.keys(filteredData[0]).forEach(key => {
        const th = document.createElement('th');
        th.innerHTML = `${key} <i class="bi bi-arrow-down-up ms-1" style="font-size:0.7rem;"></i>`;
        th.addEventListener('click', () => sortTableByColumn(key));
        tr.appendChild(th);
    });
}

function displayTable() {
    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';

    if (!filteredData.length) {
        tbody.innerHTML = `<tr><td colspan="100%" class="text-center py-4 text-muted">Nenhum dado condizente com os filtros informados.</td></tr>`;
        updatePaginationIndicators();
        return;
    }

    const start = (currentPage - 1) * recordsPerPage;
    const end = start + recordsPerPage;
    const paginatedItems = filteredData.slice(start, end);

    paginatedItems.forEach(row => {
        const tr = document.createElement('tr');
        Object.keys(row).forEach(key => {
            const td = document.createElement('td');
            let val = row[key];
            
            // Formatador dinâmico de saída visual de células
            if(val instanceof Date) {
                td.innerText = val.toLocaleDateString('pt-BR');
            } else if (typeof val === 'number') {
                td.innerText = val % 1 === 0 ? val : val.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                td.classList.add('text-end');
            } else {
                td.innerText = val;
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    updatePaginationIndicators();
}

// Auxiliares de Ordenação e Agrupamentos de Massa de Dados
function sortTableByColumn(key) {
    sortDirection = !sortDirection;
    filteredData.sort((a, b) => {
        let valA = a[key];
        let valB = b[key];
        if (typeof valA === 'string') return sortDirection ? valA.localeCompare(valB) : valB.localeCompare(valA);
        return sortDirection ? valA - valB : valB - valA;
    });
    displayTable();
}

function aggregateData(groupKey, sumKey, isDate = false) {
    const groups = {};
    filteredData.forEach(row => {
        let label = row[groupKey] || 'N/A';
        if(isDate && label instanceof Date) label = label.toLocaleDateString('pt-BR');
        
        const value = sumKey ? (Number(row[sumKey]) || 0) : 1;
        groups[label] = (groups[label] || 0) + value;
    });
    return groups;
}

function updatePaginationIndicators() {
    const totalPages = Math.ceil(filteredData.length / recordsPerPage) || 1;
    document.getElementById('pagination-indicator').innerText = `Página ${currentPage} de ${totalPages}`;
    document.getElementById('btn-prev-page').disabled = currentPage === 1;
    document.getElementById('btn-next-page').disabled = currentPage === totalPages;
    document.getElementById('table-info').innerText = `Visualizando ${filteredData.length} registros`;
}

// Utilitários Utilitários de Interface
function formatCurrency(val) {
    return val.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

// Alternador Nativo do Modo Escuro corporativo
function toggleTheme() {
    const body = document.body;
    const icon = document.getElementById('theme-icon');
    
    if (body.classList.contains('dark-mode')) {
        body.classList.remove('dark-mode');
        icon.className = "bi bi-moon-stars";
        localStorage.setItem('theme', 'light');
    } else {
        body.classList.add('dark-mode');
        icon.className = "bi bi-sun";
        localStorage.setItem('theme', 'dark');
    }
    if (rawData.length > 0) renderCharts();
}

function initTheme() {
    if (localStorage.getItem('theme') === 'dark') {
        document.body.classList.add('dark-mode');
        document.getElementById('theme-icon').className = "bi bi-sun";
    }
}

// Geração de Exportação do PDF Estruturado do Dashboard via jsPDF + AutoTable
function exportDashboardPDF() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('l', 'pt', 'a4'); // Paisagem para visualização estendida corporativa

    doc.setFont("helvetica", "bold");
    doc.setFontSize(22);
    doc.text("Relatório Analítico Executivo", 40, 50);

    doc.setFontSize(10);
    doc.setFont("helvetica", "normal");
    doc.text(`Extração gerada automaticamente via Nexus Analytics em: ${new Date().toLocaleDateString('pt-BR')} ${new Date().toLocaleTimeString('pt-BR')}`, 40, 70);

    // Renderiza tabela inteira filtrada no PDF
    doc.autoTable({
        html: '#data-table',
        startY: 100,
        theme: 'striped',
        styles: { fontSize: 8, font: 'helvetica' },
        headStyles: { fillColor: [30, 41, 59] }
    });

    doc.save(`Relatorio_Executivo_Nexus_${new Date().toISOString().split('T')[0]}.pdf`);
}