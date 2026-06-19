// ============================================
// DASHBOARD DE ANÁLISE DE ESTOQUE - SANTA CRUZ
// Lê arquivos .xlsx nativamente via SheetJS (offline)
// ============================================

// CONFIGURAÇÕES DE FAIXAS DE DIAS DE ESTOQUE
const DIAS_MIN_SAUDAVEL = 45;   // abaixo disso = déficit
const DIAS_MAX_SAUDAVEL = 60;   // 60 ou mais = excesso (alinha com o filtro >= 60 dias)
const DIAS_ENCALHADO = 100;     // acima disso = encalhado (estoque parado), mesmo com giro lento

// Aba Ritmo Entrada × Saída: variação % do giro recente vs base abaixo da qual o
// movimento é "Estável" (ruído), e giro mínimo (un/mês) para confiar na tendência.
const BANDA_TENDENCIA = 0.12;
const MIN_GIRO_TENDENCIA = 5;

// Reserva considerada "estoque livre do CD". Só esta reserva entra nos totais de
// estoque (livre/qualidade/bloqueado/total/R$), nas pendências E nas vendas/giro.
// As demais (Colaborativa, Loja a Loja, OL Indústria) NÃO são estoque livre
// disponível e não devem inflar a velocidade: os dias de estoque livre são
// Estoque Livre (Regular) ÷ Venda (Regular), idêntico à coluna do Excel.
// Use '' (string vazia) para voltar a somar TODAS as reservas.
const RESERVA_LIVRE = 'Regular';

// Quantos dias de cobertura mirar nas sugestões de compra (PADRÃO para todas as linhas)
const META_DIAS_COBERTURA = 60;
const META_DIAS_MIN = 7;
const META_DIAS_MAX = 365;

// Tolerância do arredondamento por caixa de embarque (fração de 1 caixa).
// A compra é arredondada PARA CIMA ao múltiplo da caixa, MAS uma caixa extra só é
// contada se a sobra exceder esta fração. Ex. com 0.10: necessidade de 3,02 caixas
// (sobra 2%) → 3 caixas; 3,30 caixas (sobra 30%) → 4 caixas. Evita comprar 1 caixa
// inteira só para cobrir uma folga ínfima. Use 0 para voltar ao ceil puro.
const TOLERANCIA_CAIXA = 0.10;

// Meta de cobertura POR LINHA (CD + produto). Vazio = usa o padrão de 60 dias.
// O usuário pode sobrescrever individualmente cada linha nas abas Sugestões e Detalhes.
const metaDiasPorLinha = {};
function chaveLinha(p) { return p.cd + '||' + p.sku; }
function getMetaDias(p) {
    const v = metaDiasPorLinha[chaveLinha(p)];
    return (typeof v === 'number' && !isNaN(v)) ? v : META_DIAS_COBERTURA;
}
function setMetaDiasLinha(p, valor) {
    let v = parseInt(valor, 10);
    if (isNaN(v)) v = META_DIAS_COBERTURA;
    v = Math.max(META_DIAS_MIN, Math.min(META_DIAS_MAX, v));
    metaDiasPorLinha[chaveLinha(p)] = v;
    return v;
}

// Tabela de múltiplos da Caixa de Embarque por código SAP
// Fonte: ESTOQUE_CD-RS_GSC_17_06_2026 - SUPERA (de-para fornecedor)
// 212 produtos. Toda sugestão de compra é arredondada PARA CIMA a este múltiplo.
const MULTIPLOS_CAIXA = {
    "106184": 128, "107420": 100, "107399": 100, "107421": 135, "110010": 135, "722186": 44, "112642": 60, "112643": 60,
    "100188": 50, "100189": 50, "106320": 50, "109831": 50, "100190": 50, "109833": 100, "111481": 60, "109832": 100,
    "111572": 60, "112263": 100, "112262": 100, "100283": 160, "100280": 160, "111764": 60, "111773": 60, "106345": 60,
    "106346": 60, "106344": 60, "106156": 100, "106153": 100, "106126": 100, "106152": 50, "106151": 100, "106158": 40,
    "111200": 50, "114738": 50, "106976": 50, "116289": 24, "106977": 50, "106975": 50, "116340": 24, "114979": 100,
    "114764": 60, "106974": 100, "106003": 50, "105939": 40, "106002": 100, "114694": 100, "114813": 60, "725354": 20,
    "721338": 40, "722135": 40, "115156": 100, "100469": 20, "100537": 50, "100538": 22, "111958": 50, "111959": 50,
    "106853": 100, "114502": 60, "106852": 90, "112032": 135, "112742": 50, "115956": 50, "110853": 50, "116201": 50,
    "110849": 50, "116193": 50, "110866": 100, "110860": 100, "110863": 60, "106946": 100, "112784": 100, "112785": 100,
    "112786": 100, "114948": 50, "114810": 60, "114762": 30, "114985": 50, "107184": 100, "107183": 100, "115309": 40,
    "115295": 40, "115296": 40, "717174": 20, "101029": 50, "112280": 60, "107111": 100, "111702": 100, "101213": 100,
    "114978": 5, "101269": 56, "101442": 96, "101491": 50, "101492": 50, "101494": 40, "101493": 40, "105825": 30,
    "101561": 100, "101563": 100, "105770": 100, "106588": 90, "108080": 60, "101737": 50, "101738": 50, "109555": 50,
    "112062": 100, "116032": 100, "112063": 100, "115873": 100, "115677": 60, "101772": 30, "101771": 24, "111711": 70,
    "115324": 50, "111204": 40, "106613": 32, "107920": 105, "107909": 105, "107091": 100, "721375": 24, "107102": 100,
    "107101": 100, "107100": 100, "114963": 96, "112781": 96, "717638": 100, "717637": 60, "717636": 100, "722596": 100,
    "722597": 60, "722595": 100, "722497": 100, "724476": 48, "114697": 100, "114698": 80, "405472": 25, "405475": 25,
    "405473": 25, "405474": 25, "405758": 24, "405753": 24, "405762": 24, "405754": 24, "405633": 25, "405634": 25,
    "109854": 50, "109853": 50, "106945": 72, "106927": 72, "106610": 40, "106611": 40, "106612": 40, "111313": 50,
    "111312": 50, "110470": 50, "110467": 50, "110811": 32, "102725": 48, "102724": 56, "102734": 40, "401053": 30,
    "106449": 60, "107422": 100, "110940": 40, "106743": 60, "103042": 40, "103044": 40, "112758": 60, "112751": 72,
    "103091": 32, "103092": 50, "110494": 50, "105302": 50, "111701": 30, "103094": 50, "110495": 50, "103095": 32,
    "103096": 30, "110496": 50, "103131": 50, "111319": 100, "111318": 100, "106179": 50, "106192": 50, "401184": 24,
    "724775": 40, "719536": 24, "114470": 20, "111190": 100, "111191": 60, "115111": 60, "115073": 60, "112031": 135,
    "111774": 50, "111830": 100, "111823": 100, "112888": 100, "112889": 100, "114771": 96, "114735": 96, "105775": 30,
    "115378": 135, "115431": 60, "717175": 18, "721693": 60, "717176": 20, "722001": 24, "115793": 100, "115811": 60,
    "115770": 100, "115689": 100, "115783": 60, "111505": 96
};

// Função utilitária: arredonda ao múltiplo da caixa de embarque do produto (para cima,
// com tolerância TOLERANCIA_CAIXA para não comprar caixa extra por uma folga ínfima).
// Retorna { multiplo, caixas, ajustado, temMultiplo }.
function ajustarAoMultiplo(p, quantidade) {
    // Normaliza o SAP: remove espaços e eventual ".0" que o Excel adiciona ao ler como número
    let sap = String(p.codigoSAP || '').trim();
    if (sap.endsWith('.0')) sap = sap.slice(0, -2);
    const multiplo = MULTIPLOS_CAIXA[sap];
    if (!multiplo || multiplo <= 0) {
        // Sem múltiplo cadastrado: mantém a quantidade original
        return { multiplo: null, caixas: null, ajustado: quantidade, temMultiplo: false };
    }
    if (quantidade <= 0) {
        return { multiplo, caixas: 0, ajustado: 0, temMultiplo: true };
    }
    // Arredonda PARA CIMA ao múltiplo, mas só conta a caixa extra se a sobra exceder a
    // tolerância. Ex.: 3,02 caixas (sobra ≤ 10%) → 3 caixas; 3,30 caixas → 4 caixas.
    const caixasExatas = quantidade / multiplo;
    const sobra = caixasExatas - Math.floor(caixasExatas);
    const caixas = (sobra > 0 && sobra <= TOLERANCIA_CAIXA)
        ? Math.floor(caixasExatas)
        : Math.ceil(caixasExatas);
    return { multiplo, caixas, ajustado: caixas * multiplo, temMultiplo: true };
}

// ESTADO GLOBAL
let dashboardData = null;
let allProducts = [];        // produtos atualmente filtrados (aba Produtos)
let originalProducts = [];   // base completa (por CD+Material)
let skuProducts = [];        // base agregada por SKU (Brasil)
let currentView = 'cd';      // 'cd' = produto+CD | 'sku' = SKU Brasil

// DOM ELEMENTS
const fileInput = document.getElementById('fileInput');
const fileName = document.getElementById('fileName');
const clearBtn = document.getElementById('clearBtn');
const exportBtn = document.getElementById('exportBtn');
const tabButtons = document.querySelectorAll('.tab-btn');
const tabContents = document.querySelectorAll('.tab-content');

// FILTRO GLOBAL de CD, Curva e Status (multi-seleção): um único estado válido para TODAS
// as abas. Cada aba mostra um seletor que é apenas uma "janela" para o mesmo conjunto, e
// todos ficam sincronizados. Conjunto vazio = "todos".
const filtroCD = new Set();
const filtroCurva = new Set();
const filtroStatus = new Set();
const productSearch = document.getElementById('productSearch');
const sortBy = document.getElementById('sortBy');
const recFilter = document.getElementById('recFilter');
const ritmoFilter = document.getElementById('ritmoFilter');
const ritmoViewSel = document.getElementById('ritmoViewSel');
const ritmoSort = document.getElementById('ritmoSort');
let ritmoView = 'sc';   // padrão: Santa Cruz (soma de todos os CDs). 'cd' = por CD
let ritmoScExplicito = false;   // true = SC escolhido de propósito; aí a visão ignora o filtro de CD

// EVENT LISTENERS
fileInput.addEventListener('change', handleFileUpload);
clearBtn.addEventListener('click', clearDashboard);
tabButtons.forEach(btn => btn.addEventListener('click', handleTabClick));

productSearch.addEventListener('input', applyFilters);
sortBy.addEventListener('change', applyFilters);
recFilter.addEventListener('change', updateRecommendations);
if (ritmoFilter) ritmoFilter.addEventListener('change', updateRitmo);
if (ritmoViewSel) ritmoViewSel.addEventListener('change', () => {
    ritmoView = ritmoViewSel.value;
    ritmoScExplicito = (ritmoView === 'sc');   // escolha explícita de SC passa a ignorar o filtro de CD
    updateRitmo();
});
if (ritmoSort) ritmoSort.addEventListener('change', updateRitmo);
const cdSalesSelectEl = document.getElementById('cdSalesSelect');
if (cdSalesSelectEl) cdSalesSelectEl.addEventListener('change', renderVendasCDDetalhe);

// --- Multi-seleção global (espelhada em várias abas) ---
const MS_STATUS_LABEL = {
    deficit: 'Necessidade de compra (< 45 dias)',
    saudavel: 'Necessidade de compra (45-60 dias)',
    excesso: 'Atenção (60-100 dias)',
    'sem-giro': 'Problema (> 100 dias)'
};
// Várias chaves apontam para o MESMO conjunto global (ex.: cd, prodCd e recCd = filtroCD).
const MS_FILTROS = {
    cd:       { set: filtroCD,     toggle: 'msCdToggle',       drop: 'msCdDrop',       opts: 'msCdOpts',       vazio: 'Todos os CDs',    rotulo: v => 'CD: ' + v,              plural: n => n + ' CDs',    onChange: aplicarTudo },
    prodCd:   { set: filtroCD,     toggle: 'msProdCdToggle',   drop: 'msProdCdDrop',   opts: 'msProdCdOpts',   vazio: 'Todos os CDs',    rotulo: v => 'CD: ' + v,              plural: n => n + ' CDs',    onChange: aplicarTudo },
    recCd:    { set: filtroCD,     toggle: 'msRecCdToggle',    drop: 'msRecCdDrop',    opts: 'msRecCdOpts',    vazio: 'Todos os CDs',    rotulo: v => 'CD: ' + v,              plural: n => n + ' CDs',    onChange: aplicarTudo },
    ritmoCd:  { set: filtroCD,     toggle: 'msRitmoCdToggle',  drop: 'msRitmoCdDrop',  opts: 'msRitmoCdOpts',  vazio: 'Todos os CDs',    rotulo: v => 'CD: ' + v,              plural: n => n + ' CDs',    onChange: ritmoCdChange },
    curva:    { set: filtroCurva,  toggle: 'msCurvaToggle',    drop: 'msCurvaDrop',    opts: 'msCurvaOpts',    vazio: 'Todas as Curvas', rotulo: v => 'Curva ' + v,            plural: n => n + ' curvas', onChange: aplicarTudo },
    recCurva: { set: filtroCurva,  toggle: 'msRecCurvaToggle', drop: 'msRecCurvaDrop', opts: 'msRecCurvaOpts', vazio: 'Todas as Curvas', rotulo: v => 'Curva ' + v,            plural: n => n + ' curvas', onChange: aplicarTudo },
    ritmoCurva: { set: filtroCurva, toggle: 'msRitmoCurvaToggle', drop: 'msRitmoCurvaDrop', opts: 'msRitmoCurvaOpts', vazio: 'Todas as Curvas', rotulo: v => 'Curva ' + v,         plural: n => n + ' curvas', onChange: aplicarTudo },
    status:   { set: filtroStatus, toggle: 'msStatusToggle',   drop: 'msStatusDrop',   opts: 'msStatusOpts',   vazio: 'Todos',           rotulo: v => MS_STATUS_LABEL[v] || v, plural: n => n + ' status', onChange: aplicarTudo }
};
const MS_CD_KEYS = ['cd', 'prodCd', 'recCd', 'ritmoCd'];
const MS_CURVA_KEYS = ['curva', 'recCurva', 'ritmoCurva'];

// Aplica e sincroniza TUDO quando qualquer filtro global muda
function aplicarTudo() {
    Object.keys(MS_FILTROS).forEach(q => msSincroniza(q)); // espelha o conjunto em todos os seletores
    applyFilters();          // Visão Geral + Produtos
    updateRecommendations(); // Sugestões
    atualizarDetalhes();     // Detalhes
    renderVendasCD();        // Vendas do CD
    updateRitmo();           // Ritmo Entrada × Saída
}

// Marcar um CD específico no seletor do Ritmo descarta a intenção de "ver SC":
// o switch automático para "Por CD" volta a valer (a guarda no updateRitmo cuida disso).
function ritmoCdChange() {
    if (filtroCD.size) ritmoScExplicito = false;
    aplicarTudo();
}
function atualizarDetalhes() {
    if (!dashboardData) return;
    const s = document.getElementById('msProdSearch');
    buildDetailProductOptions(s ? s.value : '');
    renderDetailBlocks();
}
// Um produto está dentro do filtro global de CD/Curva/Status?
function dentroDoFiltroGlobal(p, { status = true } = {}) {
    if (filtroCD.size && !filtroCD.has(p.cd)) return false;
    if (filtroCurva.size && !filtroCurva.has(p.curva)) return false;
    if (status && filtroStatus.size && !filtroStatus.has(p.status)) return false;
    return true;
}
function msAtualizaToggle(qual) {
    const cfg = MS_FILTROS[qual];
    const btn = document.getElementById(cfg.toggle);
    if (!btn) return;
    const n = cfg.set.size;
    const txt = n === 0 ? cfg.vazio : (n === 1 ? cfg.rotulo([...cfg.set][0]) : cfg.plural(n));
    btn.innerHTML = `${escapeHtml(txt)} <span class="ms-arrow">▾</span>`;
}
function msMontaOpcoes(qual, items) {
    const cfg = MS_FILTROS[qual];
    const box = document.getElementById(cfg.opts);
    if (!box) return;
    if (!items.length) { box.innerHTML = '<div class="ms-empty">Sem opções</div>'; return; }
    // Linha "Todos" no topo: marcada quando nada está selecionado (conjunto vazio = todos)
    const todos = `<label class="ms-opt ms-opt-todos"><input type="checkbox" data-mstodos ${cfg.set.size === 0 ? 'checked' : ''}> <span>${escapeHtml(cfg.vazio)}</span></label>`;
    const linhas = items.map(it => `<label class="ms-opt"><input type="checkbox" value="${escapeHtml(it.value)}" ${cfg.set.has(it.value) ? 'checked' : ''}> <span>${escapeHtml(it.label)}</span></label>`).join('');
    box.innerHTML = todos + linhas;
    const opt = box.querySelector('input[data-mstodos]');
    if (opt) opt.addEventListener('change', () => { cfg.set.clear(); (cfg.onChange || aplicarTudo)(); });
    box.querySelectorAll('input[type="checkbox"]:not([data-mstodos])').forEach(chk => {
        chk.addEventListener('change', () => {
            if (chk.checked) cfg.set.add(chk.value); else cfg.set.delete(chk.value);
            (cfg.onChange || aplicarTudo)();
        });
    });
    msAtualizaToggle(qual);
}
function msSincroniza(qual) {
    const cfg = MS_FILTROS[qual];
    const box = document.getElementById(cfg.opts);
    if (box) {
        box.querySelectorAll('input[type="checkbox"]:not([data-mstodos])').forEach(chk => { chk.checked = cfg.set.has(chk.value); });
        const opt = box.querySelector('input[data-mstodos]');
        if (opt) opt.checked = cfg.set.size === 0; // "Todos" marcado quando não há filtro
    }
    msAtualizaToggle(qual);
}
function msFechaTodos() {
    Object.values(MS_FILTROS).forEach(cfg => {
        const d = document.getElementById(cfg.drop);
        const t = document.getElementById(cfg.toggle);
        if (d) d.classList.remove('open');
        if (t) t.classList.remove('active');
    });
}
function setupFiltrosMulti() {
    Object.keys(MS_FILTROS).forEach(qual => {
        const cfg = MS_FILTROS[qual];
        const btn = document.getElementById(cfg.toggle);
        const drop = document.getElementById(cfg.drop);
        if (!btn || !drop) return;
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const aberto = drop.classList.contains('open');
            msFechaTodos();
            if (!aberto) { drop.classList.add('open'); btn.classList.add('active'); }
        });
        drop.addEventListener('click', e => e.stopPropagation());
    });
    document.addEventListener('click', msFechaTodos);
    // Opções fixas de Curva (Visão Geral e Sugestões); CD entra em populateFilters
    const curvas = [{ value: 'A', label: 'Curva A' }, { value: 'B', label: 'Curva B' }, { value: 'C', label: 'Curva C' }];
    msMontaOpcoes('curva', curvas);
    msMontaOpcoes('recCurva', curvas);
    msMontaOpcoes('ritmoCurva', curvas);
    msMontaOpcoes('status', [
        { value: 'deficit', label: MS_STATUS_LABEL.deficit },
        { value: 'saudavel', label: MS_STATUS_LABEL.saudavel },
        { value: 'excesso', label: MS_STATUS_LABEL.excesso },
        { value: 'sem-giro', label: MS_STATUS_LABEL['sem-giro'] }
    ]);
}
setupFiltrosMulti();

// Exportar para Excel
exportBtn.addEventListener('click', exportarExcel);

// Edição da meta de dias POR LINHA (delegação de eventos — a tabela é redesenhada)
document.getElementById('recommendationsContainer').addEventListener('change', (e) => {
    const inp = e.target.closest('.row-meta-input');
    if (inp && inp.dataset.recIdx !== undefined) onRecMetaChange(parseInt(inp.dataset.recIdx, 10), inp.value);
});
// Ordenação por clique no cabeçalho da tabela de Sugestões
document.getElementById('recommendationsContainer').addEventListener('click', (e) => {
    const th = e.target.closest('th.sortable');
    if (th && th.dataset.sortKey) ordenarRecPor(th.dataset.sortKey);
});
document.getElementById('detailBlocks').addEventListener('change', (e) => {
    const inp = e.target.closest('.row-meta-input');
    if (inp && inp.dataset.detIdx !== undefined) onDetMetaChange(parseInt(inp.dataset.detIdx, 10), inp.value);
});
// Ordenação por clique no cabeçalho da tabela de Detalhes
document.getElementById('detailBlocks').addEventListener('click', (e) => {
    const th = e.target.closest('th.sortable');
    if (th && th.dataset.sortKey) ordenarDetalhePor(th.dataset.sortKey);
});

// Estado da aba Detalhes (multi-seleção)
const detailCDs = new Set();        // CDs selecionados (vazio = todos)
const detailProducts = new Set();   // materiais selecionados

setupDetailMultiselects();

// Retorna a base ativa conforme a visão selecionada
function baseAtual() {
    return currentView === 'sku' ? skuProducts : originalProducts;
}

// KPIs clicáveis: levam à lista/aba correspondente
document.querySelectorAll('.kpi-clickable').forEach(card => {
    card.addEventListener('click', () => {
        if (!dashboardData) return;
        const kpi = card.getAttribute('data-kpi');
        if (kpi === 'compra') {
            // Necessidade de compra (< 60 dias): leva à aba Sugestões. O filtro de CD/Curva
            // já é global, então as Sugestões herdam automaticamente o que está selecionado.
            const recF = document.getElementById('recFilter');
            if (recF) recF.value = 'deficit';
            updateRecommendations();
            ativarAba('recommendations');
        } else {
            // Vai para a aba Produtos marcando o status no filtro global.
            filtroStatus.clear();
            if (kpi === 'excesso') filtroStatus.add('excesso');       // atenção 60-100
            else if (kpi === 'sem-giro') filtroStatus.add('sem-giro'); // problema > 100
            // garante ordenação útil
            if (kpi === 'excesso' || kpi === 'sem-giro') sortBy.value = 'dias-desc';
            else if (kpi === 'total') sortBy.value = 'estoque-desc';
            aplicarTudo();
            ativarAba('products');
        }
    });
});

// Ativa uma aba programaticamente
function ativarAba(tabId) {
    tabButtons.forEach(b => b.classList.remove('active'));
    tabContents.forEach(c => c.classList.remove('active'));
    const btn = Array.from(tabButtons).find(b => b.getAttribute('data-tab') === tabId);
    if (btn) btn.classList.add('active');
    const content = document.getElementById(tabId);
    if (content) content.classList.add('active');
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ============================================
// UPLOAD E LEITURA DO EXCEL
// ============================================

// ============================================
// PERSISTÊNCIA LOCAL DA PLANILHA (IndexedDB)
// Guarda o arquivo no navegador para reabrir já carregado.
// Apagado somente em "Limpar Dados".
// ============================================
const IDB_DB = 'santacruz-dashboard';
const IDB_STORE = 'arquivo';
const IDB_KEY = 'planilha';
const LS_NOME = 'sc_nome_arquivo';

function idbOpen() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(IDB_DB, 1);
        req.onupgradeneeded = () => req.result.createObjectStore(IDB_STORE);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}
async function idbSalvar(buffer) {
    try {
        const db = await idbOpen();
        await new Promise((resolve, reject) => {
            const tx = db.transaction(IDB_STORE, 'readwrite');
            tx.objectStore(IDB_STORE).put(buffer, IDB_KEY);
            tx.oncomplete = resolve;
            tx.onerror = () => reject(tx.error);
        });
        db.close();
    } catch (e) { console.warn('Não foi possível salvar a planilha localmente:', e); }
}
async function idbCarregar() {
    try {
        const db = await idbOpen();
        const buffer = await new Promise((resolve, reject) => {
            const tx = db.transaction(IDB_STORE, 'readonly');
            const req = tx.objectStore(IDB_STORE).get(IDB_KEY);
            req.onsuccess = () => resolve(req.result || null);
            req.onerror = () => reject(req.error);
        });
        db.close();
        return buffer;
    } catch (e) { console.warn('Não foi possível ler a planilha salva:', e); return null; }
}
async function idbApagar() {
    try {
        const db = await idbOpen();
        await new Promise((resolve, reject) => {
            const tx = db.transaction(IDB_STORE, 'readwrite');
            tx.objectStore(IDB_STORE).delete(IDB_KEY);
            tx.oncomplete = resolve;
            tx.onerror = () => reject(tx.error);
        });
        db.close();
    } catch (e) { console.warn('Não foi possível apagar a planilha salva:', e); }
    try { localStorage.removeItem(LS_NOME); } catch (e) {}
}

// Processa um buffer de Excel e monta o dashboard (usado pelo upload e pela restauração)
function carregarPlanilha(buffer, nomeArquivo) {
    const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
    const processedData = processWorkbook(workbook, nomeArquivo);
    if (!processedData || processedData.products.length === 0) {
        throw new Error('Nenhum produto encontrado nas abas de curva (CURVA - A/B/C).');
    }
    dashboardData = processedData;
    originalProducts = processedData.products;
    computarAmbiguidade(originalProducts);
    skuProducts = agregarPorSKU(processedData.products, processedData.diasCorridos);
    allProducts = processedData.products.slice();
    currentView = 'cd';
    document.getElementById('fileName').textContent = nomeArquivo;
    document.getElementById('dataDate').textContent = processedData.dataReferencia;
    clearBtn.style.display = 'inline-block';
    exportBtn.style.display = 'inline-block';
    fileInput.value = '';
    populateFilters();
    updateAllVisualizations();
    console.log('Dashboard carregado:', originalProducts.length, 'produtos');
}

async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    if (typeof XLSX === 'undefined') {
        alert('A biblioteca de leitura ainda não carregou. Recarregue a página e tente de novo.');
        return;
    }

    try {
        console.log('Lendo arquivo:', file.name, '(' + (file.size / 1024 / 1024).toFixed(2) + ' MB)');
        const buffer = await file.arrayBuffer();
        carregarPlanilha(buffer, file.name);
        // Guarda para reabrir já carregado (persiste até "Limpar Dados")
        idbSalvar(buffer);
        try { localStorage.setItem(LS_NOME, file.name); } catch (e) {}
    } catch (error) {
        console.error('Erro ao processar:', error);
        alert('Erro ao processar o arquivo:\n' + error.message);
    }
}

function processWorkbook(workbook, nomeArquivo) {
    // Data de referência vem do nome do arquivo: dia = corte do mês vigente,
    // mês = mês vigente (parcial). A janela de análise é esse mês + os 4 anteriores.
    const dataArq = parseDataArquivo(nomeArquivo);
    const diasCorridos = dataArq.dia;
    JANELA = construirJanela(dataArq.mes0);
    console.log('Mês vigente (parcial):', MESES_CAP[dataArq.mes0], '| corte dia', diasCorridos,
        '| janela:', JANELA.map(j => j.abbr).join(', '));

    // A fonte correta é a aba "Mapa Estoque", que traz cada produto/CD/reserva
    // uma única vez e tem o campo "Curva CD" com a curva real. As abas
    // CURVA - A/B/C contêm dados sobrepostos (a mesma linha repetida nas 3),
    // o que inflava os números e classificava a curva errada.
    const abaMapa = workbook.SheetNames.find(n =>
        n.toLowerCase().includes('mapa') && n.toLowerCase().includes('estoque')
    );

    if (!abaMapa) {
        throw new Error('Aba "Mapa Estoque" não encontrada na planilha.');
    }

    console.log('Lendo aba:', abaMapa);

    const ws = workbook.Sheets[abaMapa];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

    // Consolida por CD+Material, somando as reservas (Regular, Colaborativa,
    // Loja a Loja...). O estoque físico costuma ficar na reserva "Regular".
    const mapa = {};
    rows.forEach(row => acumularLinha(mapa, row, null, diasCorridos));
    console.log(`${abaMapa}: ${rows.length} linhas lidas`);

    // Finalizar cálculos por produto
    let products = Object.values(mapa);
    products.forEach(finalizarProduto);

    // Excluir produtos totalmente inativos (sem estoque e sem venda no período)
    const antes = products.length;
    products = products.filter(p => !(p.estoqueTotal === 0 && p.vendaMedia === 0));
    console.log(`Consolidados: ${antes} | Ativos: ${products.length} | Inativos removidos: ${antes - products.length}`);

    return consolidate(products, nomeArquivo, diasCorridos);
}

const _num = (v) => {
    const n = parseFloat(v);
    return isNaN(n) ? 0 : n;
};

function acumularLinha(mapa, row, curvaAba, diasCorridos) {
    const cd = String(row['CD'] || '').trim();
    const material = String(row['Material'] || '').trim();
    if (!cd || !material) return;

    // A curva REAL está no campo "Curva CD" da linha.
    let curva = String(row['Curva CD'] || '').trim().toUpperCase();
    if (curva !== 'A' && curva !== 'B' && curva !== 'C') {
        curva = curvaAba || 'N/A'; // fallback se o campo vier vazio
    }

    const codigoSAP = String(row['Código SAP'] || '');
    const ean = String(row['EAN'] || '');
    // Identidade pelo SKU (SAP), não pela descrição: SKUs distintos ficam separados.
    const sku = skuKey(codigoSAP, ean, material);
    const key = cd + '||' + sku;

    if (!mapa[key]) {
        mapa[key] = {
            cd, material, curva, sku,
            reserva: '',
            fornecedor: String(row['Fornecedor'] || '').trim(),
            codigoSAP: codigoSAP,
            ean: ean,
            statusMaterial: String(row['Status Material'] || '').trim(),
            pf: 0,

            // Vendas por mês da janela (índice 0 = mais antigo ... 4 = mês vigente parcial)
            vendas: [0, 0, 0, 0, 0],
            vendasRS: [0, 0, 0, 0, 0],
            vendaMedia: 0, vendaMediaRS: 0,

            estoqueLivre: 0, estoqueQualidade: 0, estoqueBloqueado: 0, estoqueTotal: 0,
            estoqueLivreRS: 0, pendenciaTransito: 0, pendenciaEntrega: 0,

            dataUltimaEntrada: '',
            linhasReserva: 0,

            // Valor pré-calculado de "dias de estoque livre" direto da coluna AM
            // ("DIAS DE ESTOQUE LIVRE UNDS") da planilha, capturado da linha Regular.
            // amCount conta quantas linhas Regular trouxeram um AM válido (esperado: 1).
            amPlanilha: null,
            amCount: 0,

            diasCorridos
        };
    }

    const o = mapa[key];
    o.linhasReserva++;

    // Preço de fábrica: mantém o maior valor não-zero
    o.pf = Math.max(o.pf, _num(row['PF']));

    // Tudo (vendas E estoque) conta APENAS a reserva livre (Regular).
    // O estoque livre só existe na Regular, então a velocidade de giro precisa
    // vir da mesma reserva: os dias de estoque livre são Estoque Livre (Regular)
    // ÷ Venda (Regular), igual à coluna do Excel. Somar a venda de todas as
    // reservas com estoque só da Regular inflava a velocidade e derrubava os dias.
    const reserva = String(row['Reserva'] || '').trim();
    if (!RESERVA_LIVRE || reserva === RESERVA_LIVRE) {
        // Vendas por mês da janela (quantidade e R$)
        JANELA.forEach((mes, i) => {
            o.vendas[i] += _num(row[mes.colQtd]);
            o.vendasRS[i] += _num(row[mes.colRS]);
        });
        o.vendaMedia += _num(row['Venda Média Ponderada']);
        o.vendaMediaRS += _num(row['Venda Média Ponderada R$']);

        o.estoqueLivre += _num(row['Estoque Livre']);
        o.estoqueQualidade += _num(row['Estoque Qualidade']);
        o.estoqueBloqueado += _num(row['Estoque Bloqueado']);
        o.estoqueTotal += _num(row['Estoque Total']);
        o.estoqueLivreRS += _num(row['Estoque Livre R$']);
        o.pendenciaTransito += _num(row['Pendência Trânsito']);
        o.pendenciaEntrega += _num(row['Pendência Entrega']);

        // Dias de estoque livre JÁ vêm calculados na planilha, na coluna AM
        // ("DIAS DE ESTOQUE LIVRE UNDS" = Estoque Livre ÷ Venda Média Ponderada × 30).
        // Captura esse valor da própria linha Regular para usar como verdade.
        const amRaw = row['DIAS DE ESTOQUE LIVRE UNDS'];
        const amNum = parseFloat(amRaw);
        if (amRaw !== '' && amRaw != null && !isNaN(amNum)) {
            o.amPlanilha = amNum;
            o.amCount++;
        }
    }

    // Data da última entrada: mantém a mais recente
    const d = String(row['Data Última Entrada'] || '').trim();
    if (d && chaveData(d) > chaveData(o.dataUltimaEntrada)) o.dataUltimaEntrada = d;

    // Completa metadados que possam ter vindo vazios na 1ª linha
    if (!o.fornecedor) o.fornecedor = String(row['Fornecedor'] || '').trim();
    if (!o.codigoSAP) o.codigoSAP = String(row['Código SAP'] || '');
}

function finalizarProduto(p) {
    // DIAS DE ESTOQUE LIVRE: usa o valor JÁ calculado pela planilha na coluna AM
    // ("DIAS DE ESTOQUE LIVRE UNDS") da linha Regular — é a coluna oficial e vale para
    // todas as curvas (A, B e C). Só recalcula (Estoque Livre Regular ÷ Venda Média
    // Regular × 30, idêntico à fórmula da coluna AM) quando o valor da planilha não está
    // disponível: visões agregadas (SKU/Brasil, médias por CD) ou venda não-positiva.
    if (p.vendaMedia > 0 && p.amCount === 1 && typeof p.amPlanilha === 'number' && isFinite(p.amPlanilha)) {
        p.diasLivre = p.amPlanilha;
    } else {
        p.diasLivre = p.vendaMedia > 0 ? (p.estoqueLivre / p.vendaMedia) * 30 : 0;
    }
    p.diasTotal = p.vendaMedia > 0 ? (p.estoqueTotal / p.vendaMedia) * 30 : 0;

    p.semGiro = p.vendaMedia <= 0;   // mantido só como dado; NÃO entra na classificação

    // Venda do mês vigente (parcial) e projeção de mês cheio a partir do corte
    p.vendaParcial = p.vendas[4];
    p.vendaParcialRS = p.vendasRS[4];
    p.vendaParcialProjetada = p.diasCorridos > 0 ? (p.vendaParcial / p.diasCorridos) * 30 : p.vendaParcial;

    p.status = classificar(p.diasLivre);
}

// Agrega os pontos (que estão por CD+SKU) num único registro por SKU,
// somando estoque, vendas e giro de todos os CDs (visão Brasil / portfólio).
function agregarPorSKU(products, diasCorridos) {
    const mapa = {};
    products.forEach(p => {
        if (!mapa[p.sku]) {
            mapa[p.sku] = {
                cd: 'BRASIL',
                material: p.material,
                sku: p.sku,
                curva: p.curva,
                fornecedor: p.fornecedor,
                codigoSAP: p.codigoSAP,
                ean: p.ean,
                statusMaterial: p.statusMaterial,
                pf: p.pf,
                vendas: [0, 0, 0, 0, 0],
                vendasRS: [0, 0, 0, 0, 0],
                vendaMedia: 0, vendaMediaRS: 0,
                estoqueLivre: 0, estoqueQualidade: 0, estoqueBloqueado: 0, estoqueTotal: 0,
                estoqueLivreRS: 0, pendenciaTransito: 0, pendenciaEntrega: 0,
                dataUltimaEntrada: '',
                nCDs: 0,
                cdsDeficit: [],
                diasCorridos
            };
        }
        const o = mapa[p.sku];
        o.nCDs++;
        o.pf = Math.max(o.pf, p.pf);
        for (let i = 0; i < 5; i++) { o.vendas[i] += p.vendas[i]; o.vendasRS[i] += p.vendasRS[i]; }
        ['vendaMedia', 'vendaMediaRS', 'estoqueLivre', 'estoqueQualidade',
            'estoqueBloqueado', 'estoqueTotal', 'estoqueLivreRS',
            'pendenciaTransito', 'pendenciaEntrega'].forEach(k => { o[k] += p[k]; });
        if (p.dataUltimaEntrada && chaveData(p.dataUltimaEntrada) > chaveData(o.dataUltimaEntrada)) {
            o.dataUltimaEntrada = p.dataUltimaEntrada;
        }
        // guarda CDs em que esse SKU está deficitário (útil pra ação)
        if (p.status === 'deficit') o.cdsDeficit.push(p.cd);
    });

    const lista = Object.values(mapa);
    lista.forEach(finalizarProduto);
    return lista;
}

// Nomes de meses em português (índice 0 = janeiro)
const MESES_FULL = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
const MESES_ABBR = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
const MESES_CAP = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

// Janela de meses ativa (5 meses: 4 cheios + o vigente parcial), derivada do nome do arquivo
let JANELA = [];

// Extrai dia/mês/ano do nome do arquivo. Aceita ano de 2 ou 4 dígitos e separador
// ponto, underscore, hífen ou barra (ex.: "12.06.26", "12_06_2026", "12-06-26").
// Ex.: "...base 12.06.26.xlsx" -> {dia:12, mes0:5, ano:2026}
function parseDataArquivo(nomeArquivo) {
    const m = String(nomeArquivo || '').match(/(\d{1,2})[._\-\/](\d{1,2})[._\-\/](\d{2,4})/);
    if (!m) {
        const now = new Date();
        return { dia: now.getDate(), mes0: now.getMonth(), ano: now.getFullYear() };
    }
    const dia = parseInt(m[1], 10);
    const mes0 = Math.min(11, Math.max(0, parseInt(m[2], 10) - 1));
    let ano = parseInt(m[3], 10);
    if (m[3].length <= 2) ano = 2000 + ano;
    return { dia, mes0, ano };
}

// Mantida por compatibilidade: dias corridos = dia do mês do arquivo
function extrairDiasCorridos(nomeArquivo) {
    return parseDataArquivo(nomeArquivo).dia;
}

// Nome (capitalizado) do mês vigente/parcial da janela ativa
function mesVigenteNome() {
    const ultimo = JANELA[JANELA.length - 1];
    return ultimo ? ultimo.cap : '';
}

// Constrói a janela de 5 meses terminando no mês vigente (último = parcial).
// Para cada mês gera os nomes de coluna esperados na planilha.
function construirJanela(mesVigente0) {
    const arr = [];
    for (let k = 4; k >= 0; k--) {
        const mi = ((mesVigente0 - k) % 12 + 12) % 12;
        arr.push({
            idx: mi,
            abbr: MESES_ABBR[mi],
            full: MESES_FULL[mi],
            cap: MESES_CAP[mi],
            colQtd: 'Qtde. Vend. ' + MESES_FULL[mi],
            colRS: 'Vendas de ' + MESES_CAP[mi] + ' R$',
            parcial: k === 0   // o último (k=0) é o mês vigente/parcial
        });
    }
    return arr;
}

// Classifica 100% pelos DIAS DE ESTOQUE LIVRE (coluna AM), por CD + produto.
// Usa o AM ARREDONDADO (mesma leitura da planilha): 59,5 a 60,4 contam como 60 e
// entram em atenção ("com o 60"). Padrão único para todas as curvas.
// SEM cruzamento com venda: AM = 0 (média ponderada zerada) NÃO é "sem venda" — é
// cobertura baixa, ou seja, necessidade de compra. Faixas:
//   < 60 dias  -> necessidade de compra (déficit < 45 / saudável 45-59)
//   60 a 100   -> atenção (60 e 100 inclusos)
//   > 100      -> problema (estoque travado)
function classificar(dias) {
    const d = Math.round(dias);
    if (d < DIAS_MIN_SAUDAVEL) return 'deficit';  // < 45  -> necessidade de compra
    if (d < DIAS_MAX_SAUDAVEL) return 'saudavel'; // 45-59 -> necessidade de compra
    if (d <= DIAS_ENCALHADO) return 'excesso';    // 60-100 -> atenção
    return 'sem-giro';                            // > 100 -> problema
}

// Identidade do SKU: SAP (preferencial) → EAN → descrição.
// Garante que SKUs distintos (SAP diferente) de uma mesma marca não sejam somados juntos.
function skuKey(codigoSAP, ean, material) {
    let sap = String(codigoSAP || '').trim();
    if (sap.endsWith('.0')) sap = sap.slice(0, -2);
    if (sap) return 'SAP:' + sap;
    const e = String(ean || '').trim();
    if (e) return 'EAN:' + e;
    return 'MAT:' + String(material || '').trim();
}

// Conjunto de descrições que correspondem a mais de um SKU (para desambiguar na tela)
let materiaisAmbiguos = new Set();
function computarAmbiguidade(produtos) {
    const porMaterial = {};
    produtos.forEach(p => {
        if (!porMaterial[p.material]) porMaterial[p.material] = new Set();
        porMaterial[p.material].add(p.sku);
    });
    materiaisAmbiguos = new Set(
        Object.keys(porMaterial).filter(m => porMaterial[m].size > 1)
    );
}

// Rótulo de exibição do produto: descrição + SAP quando a descrição é ambígua
function rotuloProduto(p) {
    if (materiaisAmbiguos.has(p.material)) {
        const sap = String(p.codigoSAP || '').replace(/\.0$/, '').trim();
        return p.material + (sap ? ` · SAP ${sap}` : '');
    }
    return p.material;
}

function consolidate(products, nomeArquivo, diasCorridos) {
    // Consolidação por CD
    const cdMap = {};
    products.forEach(p => {
        if (!cdMap[p.cd]) {
            cdMap[p.cd] = {
                cd: p.cd,
                totalProdutos: 0,
                estoqueLivreTotal: 0,
                estoqueLivreRSTotal: 0,
                vendaMediaTotal: 0,
                deficit: 0, saudavel: 0, excesso: 0, semGiro: 0
            };
        }
        const c = cdMap[p.cd];
        c.totalProdutos++;
        c.estoqueLivreTotal += p.estoqueLivre;
        c.estoqueLivreRSTotal += p.estoqueLivreRS;
        c.vendaMediaTotal += p.vendaMedia;
        if (p.status === 'deficit') c.deficit++;
        else if (p.status === 'saudavel') c.saudavel++;
        else if (p.status === 'excesso') c.excesso++;
        else if (p.status === 'sem-giro') c.semGiro++;
    });

    // Dias médios PONDERADOS: estoque livre total / venda média total * 30
    // (mesmo método da aba oficial "DIAS DE ESTOQUE POR CD")
    Object.values(cdMap).forEach(c => {
        c.diasMedios = c.vendaMediaTotal > 0 ? (c.estoqueLivreTotal / c.vendaMediaTotal) * 30 : 0;
    });

    // Data de referência (corrige ano de 2 ou 4 dígitos)
    const da = parseDataArquivo(nomeArquivo);
    const dataReferencia = `${String(da.dia).padStart(2, '0')}/${String(da.mes0 + 1).padStart(2, '0')}/${da.ano}`;

    return {
        products,
        cdSummary: Object.values(cdMap).sort((a, b) => a.cd.localeCompare(b.cd)),
        dataReferencia,
        diasCorridos
    };
}

// ============================================
// LIMPAR
// ============================================

function clearDashboard() {
    idbApagar(); // remove a planilha guardada — só "Limpar Dados" zera o armazenamento
    dashboardData = null;
    allProducts = [];
    originalProducts = [];
    skuProducts = [];
    currentView = 'cd';
    filtroCD.clear();
    filtroCurva.clear();
    filtroStatus.clear();
    msMontaOpcoes('cd', []);
    msSincroniza('curva');
    msSincroniza('status');
    fileInput.value = '';
    fileName.textContent = 'Selecionar Planilha Excel';
    clearBtn.style.display = 'none';
    exportBtn.style.display = 'none';
    Object.keys(metaDiasPorLinha).forEach(k => delete metaDiasPorLinha[k]);
    recListaAtual = [];
    detailSort = { key: null, dir: 'desc' };
    recSort = { key: null, dir: 'desc' };
    document.getElementById('dataDate').textContent = '--';

    document.getElementById('productsTableBody').innerHTML =
        '<tr class="empty-state"><td colspan="8">Carregue uma planilha para visualizar produtos</td></tr>';
    document.getElementById('cdGrid').innerHTML =
        '<div class="empty-state">Carregue uma planilha para visualizar análise por CD</div>';
    document.getElementById('recommendationsContainer').innerHTML =
        '<div class="empty-state">Carregue uma planilha para gerar sugestões de compra</div>';
    const ritmoCont = document.getElementById('ritmoContainer');
    if (ritmoCont) ritmoCont.innerHTML = '<div class="empty-state">Carregue uma planilha para analisar o ritmo entrada × saída</div>';
    const ritmoKpisEl = document.getElementById('ritmoKpis');
    if (ritmoKpisEl) ritmoKpisEl.innerHTML = '';
    document.getElementById('cdAverageChart').innerHTML = '';

    ['kpiEstoqueTotal', 'kpiExcesso', 'kpiSemGiro', 'kpiCds'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.textContent = '--';
    });

    filtroCD.clear();
    filtroCurva.clear();
    filtroStatus.clear();
    ['cd', 'prodCd', 'recCd', 'ritmoCd'].forEach(q => msMontaOpcoes(q, []));
    msSincroniza('curva'); msSincroniza('recCurva'); msSincroniza('ritmoCurva'); msSincroniza('status');
    productSearch.value = '';
    // Detalhes
    detailCDs.clear();
    detailProducts.clear();
    document.getElementById('msCDDropdown').innerHTML = '';
    document.getElementById('msProdOptions').innerHTML = '';
    document.getElementById('msCDChips').innerHTML = '';
    document.getElementById('msProdChips').innerHTML = '';
    document.getElementById('detailBlocks').innerHTML = '';
    const msProdSearch = document.getElementById('msProdSearch');
    if (msProdSearch) msProdSearch.value = '';
    updateDetailLabels();
    document.getElementById('noDetailSelected').style.display = 'block';
    document.getElementById('noDetailSelected').textContent = 'Selecione ao menos um produto para ver detalhes';
    const af = document.getElementById('activeFilters');
    if (af) { af.style.display = 'none'; af.innerHTML = ''; }
}

// ============================================
// FILTROS
// ============================================

function populateFilters() {
    if (!dashboardData) return;
    const cds = [...new Set(dashboardData.products.map(p => p.cd))].sort();

    // CD global em multi-seleção, espelhado em Visão Geral, Produtos e Sugestões.
    // Descarta seleções de CDs que sumiram na nova base.
    [...filtroCD].forEach(v => { if (!cds.includes(v)) filtroCD.delete(v); });
    const opcoesCd = cds.map(cd => ({ value: cd, label: cd }));
    ['cd', 'prodCd', 'recCd', 'ritmoCd'].forEach(q => msMontaOpcoes(q, opcoesCd));

    // Detalhes: monta as opções de CD e Produto (multi-seleção)
    detailCDs.clear();
    detailProducts.clear();
    buildDetailCdOptions(cds);
    buildDetailProductOptions('');
    renderDetailChips();
    updateDetailLabels();
    renderDetailBlocks();
}

function applyFilters() {
    if (!dashboardData) return;

    const search = productSearch.value.toLowerCase();
    const sort = sortBy.value || 'dias-asc';

    let filtered = baseAtual().filter(p => {
        if (currentView === 'cd' && filtroCD.size && !filtroCD.has(p.cd)) return false;
        if (filtroCurva.size && !filtroCurva.has(p.curva)) return false;
        if (filtroStatus.size && !filtroStatus.has(p.status)) return false;
        if (search && !p.material.toLowerCase().includes(search) &&
            !p.fornecedor.toLowerCase().includes(search) &&
            !String(p.codigoSAP).includes(search)) return false;
        return true;
    });

    switch (sort) {
        case 'dias-asc': filtered.sort((a, b) => a.diasLivre - b.diasLivre); break;
        case 'dias-desc': filtered.sort((a, b) => b.diasLivre - a.diasLivre); break;
        case 'venda-desc': filtered.sort((a, b) => b.vendaMedia - a.vendaMedia); break;
        case 'estoque-desc': filtered.sort((a, b) => b.estoqueLivre - a.estoqueLivre); break;
    }

    allProducts = filtered;
    updateProductsTable();
    updateKPIs();
    updateCharts();        // redesenha donut + barra por CD com a base JÁ FILTRADA
    updateCDAnalysis();    // mantém os cards por CD coerentes com o filtro
    updateActiveFilters();
}

// Mostra chips dos filtros ativos no topo da aba Produtos, com botão de limpar
function updateActiveFilters() {
    const bar = document.getElementById('activeFilters');
    if (!bar) return;

    const chips = [];
    const statusLabels = { deficit: 'Necessidade de compra (<45 dias)', saudavel: 'Necessidade de compra (45-60)', excesso: 'Atenção (60-100 dias)', 'sem-giro': 'Problema (>100 dias)' };

    filtroCD.forEach(v => chips.push({ k: 'cd', v, txt: 'CD: ' + v }));
    filtroCurva.forEach(v => chips.push({ k: 'curva', v, txt: 'Curva ' + v }));
    filtroStatus.forEach(v => chips.push({ k: 'status', v, txt: statusLabels[v] || v }));
    if (productSearch.value) chips.push({ k: 'search', txt: 'Busca: "' + productSearch.value + '"' });

    if (chips.length === 0) {
        bar.style.display = 'none';
        bar.innerHTML = '';
        return;
    }

    bar.style.display = 'flex';
    bar.innerHTML = '<span class="af-label">Filtros ativos:</span>' +
        chips.map(c => `<span class="af-chip">${escapeHtml(c.txt)} <span class="af-x" data-clear="${c.k}" data-val="${escapeHtml(c.v || '')}">✕</span></span>`).join('') +
        '<button class="af-clear-all">Limpar tudo</button>';

    bar.querySelectorAll('.af-x').forEach(x => {
        x.addEventListener('click', () => {
            const k = x.getAttribute('data-clear');
            const val = x.getAttribute('data-val');
            if (k === 'cd') filtroCD.delete(val);
            else if (k === 'curva') filtroCurva.delete(val);
            else if (k === 'status') filtroStatus.delete(val);
            else if (k === 'search') productSearch.value = '';
            aplicarTudo();
        });
    });
    bar.querySelector('.af-clear-all').addEventListener('click', () => {
        filtroCD.clear(); filtroCurva.clear(); filtroStatus.clear();
        productSearch.value = '';
        aplicarTudo();
    });
}

// ============================================
// VISUALIZAÇÕES
// ============================================

function updateAllVisualizations() {
    updateViewHint();
    updateKPIs();
    updateProductsTable();
    updateCharts();
    updateCDAnalysis();
    updateRecommendations();
    renderVendasCD();
    updateRitmo();
}

function updateViewHint() {
    const hint = document.getElementById('viewHint');
    if (!hint || !dashboardData) return;
    const nSku = skuProducts.length;
    const nPontos = originalProducts.length;
    if (currentView === 'sku') {
        hint.textContent = `${nSku} SKUs consolidados (soma de todos os CDs)`;
    } else {
        hint.textContent = `${nSku} SKUs distribuídos em ${nPontos.toLocaleString('pt-BR')} pontos de estoque`;
    }
}

function updateKPIs() {
    if (!dashboardData) return;
    const prods = (allProducts.length > 0 || isFilterActive()) ? allProducts : baseAtual();

    // Valores em R$ por faixa de DIAS DE ESTOQUE LIVRE (coluna AM), 100% pela AM:
    //   necessidade de compra = < 60 dias (status deficit + saudavel)
    //   atenção               = 60 a 100 dias (status excesso)
    //   problema              = > 100 dias (status sem-giro)
    const estTotal = prods.reduce((s, p) => s + p.estoqueLivreRS, 0);
    const estCompra = prods.filter(p => p.status === 'deficit' || p.status === 'saudavel')
        .reduce((s, p) => s + p.estoqueLivreRS, 0);
    const estAtencao = prods.filter(p => p.status === 'excesso').reduce((s, p) => s + p.estoqueLivreRS, 0);
    const estProblema = prods.filter(p => p.status === 'sem-giro').reduce((s, p) => s + p.estoqueLivreRS, 0);

    const pctCompra = estTotal > 0 ? (estCompra / estTotal * 100) : 0;
    const pctAtencao = estTotal > 0 ? (estAtencao / estTotal * 100) : 0;
    const pctProblema = estTotal > 0 ? (estProblema / estTotal * 100) : 0;

    document.getElementById('kpiEstoqueTotal').textContent = formatBRLCheio(estTotal);
    document.getElementById('kpiExcesso').textContent = formatBRLCheio(estAtencao);
    document.getElementById('kpiExcessoLabel').innerHTML = `Atenção (60-100 dias) &middot; ${pctAtencao.toFixed(0)}% do capital`;
    document.getElementById('kpiSemGiro').textContent = formatBRLCheio(estProblema);
    document.getElementById('kpiSemGiroLabel').innerHTML = `Problema (&gt; 100 dias) &middot; ${pctProblema.toFixed(0)}% do capital`;

    // 4o card: Necessidade de compra (< 60 dias de cobertura)
    document.getElementById('kpiCds').textContent = formatBRLCheio(estCompra);
    document.getElementById('kpiCdsLabel').innerHTML = `Necessidade de compra (&lt; 60 dias) &middot; ${pctCompra.toFixed(0)}% do capital`;
}

function isFilterActive() {
    return filtroCD.size || filtroCurva.size || filtroStatus.size || productSearch.value;
}

function updateProductsTable() {
    if (!dashboardData) return;
    const tbody = document.getElementById('productsTableBody');
    const prods = (allProducts.length > 0 || isFilterActive()) ? allProducts : baseAtual();
    const isSku = currentView === 'sku';

    if (prods.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:24px;">Nenhum produto encontrado</td></tr>';
        return;
    }

    const limite = 500;
    tbody.innerHTML = prods.slice(0, limite).map(p => `
        <tr>
            <td><strong>${isSku ? '🌐 ' + p.nCDs + ' CDs' : p.cd}</strong></td>
            <td title="${escapeHtml(rotuloProduto(p))}">${truncate(rotuloProduto(p), 36)}</td>
            <td>${truncate(p.fornecedor, 24)}</td>
            <td><span class="curva-badge curva-${p.curva}">${p.curva}</span></td>
            <td>${p.vendaMedia.toFixed(1)}</td>
            <td>${p.estoqueLivre.toFixed(0)}</td>
            <td><strong>${Math.round(p.diasLivre)}</strong></td>
            <td><span class="status-badge ${statusGrupo(p.status)}">${formatStatus(p.status)}</span></td>
        </tr>
    `).join('');

    if (prods.length > limite) {
        tbody.innerHTML += `<tr><td colspan="8" style="text-align:center;padding:12px;color:#64748b;font-style:italic;">Mostrando ${limite} de ${prods.length.toLocaleString('pt-BR')} ${isSku ? 'SKUs' : 'pontos'}. Use os filtros para refinar.</td></tr>`;
    }
}

function updateCharts() {
    if (!dashboardData) return;
    const prods = (allProducts.length > 0 || isFilterActive()) ? allProducts : baseAtual();

    // (Gráfico de pizza "Distribuição de Dias de Estoque por Faixa" removido a pedido:
    //  a classificação por faixa usa a cobertura recalculada do dashboard, que difere
    //  da coluna pré-calculada "DIAS DE ESTOQUE LIVRE UNDS" da planilha. As KPIs e a
    //  aba Produtos continuam refletindo os filtros.)

    // O gráfico por CD só faz sentido na visão CD
    const cdChartCard = document.getElementById('cdAverageChart').closest('.chart-container');
    if (currentView === 'sku') {
        if (cdChartCard) cdChartCard.style.display = 'none';
    } else {
        if (cdChartCard) cdChartCard.style.display = '';
        // Recalcula a cobertura ponderada por CD a partir da base FILTRADA (prods),
        // para a barra acompanhar os filtros de CD/Curva/Status — igual ao donut e às KPIs.
        // (Antes usava dashboardData.cdSummary, que é sempre a base cheia e não filtrava.)
        const cdAgg = {};
        prods.forEach(p => {
            if (p.cd === 'BRASIL') return;
            if (!cdAgg[p.cd]) cdAgg[p.cd] = { est: 0, venda: 0 };
            cdAgg[p.cd].est += p.estoqueLivre;
            cdAgg[p.cd].venda += p.vendaMedia;
        });
        const cdData = Object.entries(cdAgg)
            .map(([cd, v]) => ({ cd, diasMedios: v.venda > 0 ? (v.est / v.venda) * 30 : 0 }))
            .sort((a, b) => a.cd.localeCompare(b.cd));
        // Média dos CDs COM giro (venda > 0). CDs sem venda mostram 0 e não entram na média,
        // pois "dias de estoque" não se define sem demanda.
        const comGiro = Object.values(cdAgg).filter(v => v.venda > 0);
        const mediaDias = comGiro.length
            ? comGiro.reduce((s, v) => s + (v.est / v.venda) * 30, 0) / comGiro.length
            : 0;
        document.getElementById('cdAverageChart').innerHTML = createBarChart(
            cdData.map(c => c.cd),
            cdData.map(c => c.diasMedios),
            '#2563eb',
            mediaDias
        );
    }
}

// ============================================
// GRÁFICOS SVG (sem dependência)
// ============================================

function createBarChart(labels, values, color, media) {
    if (labels.length === 0) return '<p style="padding:24px;color:#64748b;">Sem dados</p>';

    const maxV = Math.max(...values, media || 0, 1);
    const h = 320, padL = 40, padB = 50, padT = 20;
    const barSlot = Math.max(46, Math.min(90, 700 / labels.length));
    const w = Math.max(640, labels.length * barSlot + padL);
    const plotH = h - padB - padT;

    let bars = '';
    values.forEach((v, i) => {
        const bh = (v / maxV) * plotH;
        const x = padL + i * barSlot + (barSlot * 0.15);
        const y = padT + plotH - bh;
        const bw = barSlot * 0.6;
        bars += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bw.toFixed(1)}" height="${bh.toFixed(1)}" fill="${color}" opacity="0.9" rx="4"/>`;
        bars += `<text x="${(x + bw / 2).toFixed(1)}" y="${h - padB + 18}" text-anchor="middle" font-size="11" fill="#64748b">${labels[i]}</text>`;
        bars += `<text x="${(x + bw / 2).toFixed(1)}" y="${(y - 6).toFixed(1)}" text-anchor="middle" font-size="11" font-weight="bold" fill="#1e293b">${v.toFixed(0)}</text>`;
    });

    // Linha de média (desenhada POR CIMA das barras, com etiqueta opaca p/ leitura)
    let mediaLine = '';
    if (media && media > 0) {
        const yM = padT + plotH - (media / maxV) * plotH;
        const txt = `Média dos CDs: ${media.toFixed(0)} dias`;
        const lblW = txt.length * 6.2 + 16;
        const boxY = Math.max(2, yM - 19);
        mediaLine = `
        <line x1="${padL}" y1="${yM.toFixed(1)}" x2="${w}" y2="${yM.toFixed(1)}" stroke="#1d4ed8" stroke-width="1.5" stroke-dasharray="6 4"/>
        <rect x="${(padL + 2)}" y="${boxY.toFixed(1)}" width="${lblW.toFixed(0)}" height="16" rx="3" fill="#ffffff" stroke="#bfdbfe" stroke-width="1"/>
        <text x="${(padL + 9)}" y="${(boxY + 11.5).toFixed(1)}" text-anchor="start" font-size="11" font-weight="700" fill="#1d4ed8">${txt}</text>`;
    }

    return `<div style="overflow-x:auto;">
        <svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
            <line x1="${padL}" y1="${padT + plotH}" x2="${w}" y2="${padT + plotH}" stroke="#e2e8f0" stroke-width="1"/>
            ${bars}
            ${mediaLine}
        </svg>
    </div>`;
}

// ============================================
// ANÁLISE POR CD
// ============================================

function updateCDAnalysis() {
    if (!dashboardData) return;
    const cds = dashboardData.cdSummary || [];
    const grid = document.getElementById('cdGrid');

    if (cds.length === 0) {
        grid.innerHTML = '<div class="empty-state">Nenhum CD encontrado</div>';
        return;
    }

    grid.innerHTML = cds.map(c => {
        let saudeClass = 'saudavel';
        if (c.diasMedios < DIAS_MIN_SAUDAVEL) saudeClass = 'deficit';
        else if (c.diasMedios > DIAS_MAX_SAUDAVEL) saudeClass = 'excesso';

        return `
        <div class="cd-card" style="border-top-color:${saudeClass === 'deficit' ? '#ef4444' : saudeClass === 'excesso' ? '#f97316' : '#22c55e'};">
            <h3>${c.cd} <span class="status-badge ${saudeClass}" style="float:right;font-size:0.7rem;">${c.diasMedios.toFixed(0)} dias</span></h3>
            <ul>
                <li><span class="label">Produtos:</span> <span class="value">${c.totalProdutos.toLocaleString('pt-BR')}</span></li>
                <li><span class="label">Estoque Livre (un):</span> <span class="value">${c.estoqueLivreTotal.toLocaleString('pt-BR')}</span></li>
                <li><span class="label">Estoque Livre (R$):</span> <span class="value">${formatBRL(c.estoqueLivreRSTotal)}</span></li>
                <li><span class="label">Dias Médios (c/ giro):</span> <span class="value">${c.diasMedios.toFixed(1)}</span></li>
            </ul>
            <div class="cd-status-row">
                <span class="status-badge deficit" title="Necessidade de compra (<45 dias)">${c.deficit} compra</span>
                <span class="status-badge saudavel" title="Necessidade de compra (45-60 dias)">${c.saudavel} ok</span>
                <span class="status-badge excesso" title="Atenção (60-100 dias)">${c.excesso} atenção</span>
                ${c.semGiro > 0 ? `<span class="status-badge sem-giro" title="Problema (>100 dias)">${c.semGiro} problema</span>` : ''}
            </div>
        </div>`;
    }).join('');
}

// ============================================
// SUGESTÕES DE COMPRA
// ============================================

function updateRecommendations() {
    if (!dashboardData) return;
    const filter = recFilter.value || 'deficit';
    const container = document.getElementById('recommendationsContainer');

    // Sugestão de compra é sempre por ponto de estoque (CD+Material). Usa o filtro GLOBAL
    // de CD e Curva (Status não entra aqui — "Focar em" é o eixo de classificação da aba).
    let prods = originalProducts;
    if (filtroCD.size) prods = prods.filter(p => filtroCD.has(p.cd));
    if (filtroCurva.size) prods = prods.filter(p => filtroCurva.has(p.curva));

    let lista = [];
    let intro = '';
    let jaRepostos = 0;

    if (filter === 'deficit') {
        intro = `Produtos abaixo de ${DIAS_MAX_SAUDAVEL} dias de cobertura. Comprar para atingir a meta de cobertura de cada linha (padrão ${META_DIAS_COBERTURA} dias, ajustável na coluna <strong>Meta</strong>), já descontado o que está a caminho e arredondado ao múltiplo da caixa de embarque.`;
        const emDeficit = prods.filter(p => p.status === 'deficit' || p.status === 'saudavel');
        lista = emDeficit.filter(p => memoriaCompra(p).comprar > 0)
            .sort((a, b) => a.diasLivre - b.diasLivre);
        jaRepostos = emDeficit.length - lista.length;
    } else if (filter === 'urgent') {
        intro = 'Compra urgente: cobertura abaixo de 30 dias e ainda sem reposição suficiente a caminho. Comprar para atingir a meta de cada linha, já descontado o que está a caminho e arredondado ao múltiplo da caixa de embarque.';
        const criticos = prods.filter(p => Math.round(p.diasLivre) < 30);
        lista = criticos.filter(p => memoriaCompra(p).comprar > 0)
            .sort((a, b) => (b.vendaMedia / (b.diasLivre + 1)) - (a.vendaMedia / (a.diasLivre + 1)));
        jaRepostos = criticos.length - lista.length;
    } else if (filter === 'performance') {
        intro = 'Estoque parado: cobertura acima de 60 dias (atenção 60-100) e problema (>100 dias). Avaliar redução de compra ou remanejamento.';
        lista = prods.filter(p => p.status === 'excesso' || p.status === 'sem-giro')
            .sort((a, b) => {
                const da = a.diasLivre;
                const db = b.diasLivre;
                return db - da;
            });
    }

    if (lista.length === 0) {
        const msg = jaRepostos > 0
            ? `Nenhum produto precisa de nova compra neste critério. ${jaRepostos} item(ns) em déficit já têm reposição suficiente a caminho.`
            : 'Nenhum produto neste critério.';
        container.innerHTML = `<div class="empty-state">${msg}</div>`;
        return;
    }

    const escopo = filtroCD.size === 0 ? 'todos os CDs'
        : (filtroCD.size === 1 ? `CD ${[...filtroCD][0]}` : `${filtroCD.size} CDs`);
    const totalComprar = lista.reduce((s, p) => s + memoriaCompra(p).comprar, 0);
    const totalCaixas = lista.reduce((s, p) => { const m = memoriaCompra(p); return s + (m.temMultiplo ? (m.caixas || 0) : 0); }, 0);
    const semMultiplo = lista.filter(p => !memoriaCompra(p).temMultiplo).length;
    const semMultTxt = semMultiplo > 0 ? ` · ${semMultiplo} item(ns) sem múltiplo cadastrado` : '';
    const caixasTxt = totalCaixas > 0 ? ` (${totalCaixas.toLocaleString('pt-BR')} caixas)` : '';
    const repostoTxt = jaRepostos > 0 ? ` · ${jaRepostos} já com reposição a caminho (fora da lista)` : '';
    const introHtml = `<div class="rec-intro">${intro}<br><strong>${lista.length} itens</strong> a comprar em ${escopo}${filter !== 'performance' ? ` · sugestão total: <span class="rec-total"><strong>${totalComprar.toLocaleString('pt-BR')} un</strong>${caixasTxt}</span>${repostoTxt}${semMultTxt}` : ''}</div>`;

    // Lista compacta (tabela). Sem teto baixo: todos os pontos abaixo de 60 dias com
    // necessidade de compra aparecem (o corte antigo de 300 escondia a maioria).
    recListaAtual = aplicarOrdenacaoRec(lista).slice(0, 5000);
    container.innerHTML = introHtml + renderRecTable(recListaAtual, filter, filtroCD.size === 1);
}

// Lista de pontos atualmente exibida na aba Sugestões (para edição de meta por linha)
let recListaAtual = [];

// Ordenação por coluna na aba Sugestões (mesma mecânica da aba Detalhes)
let recSort = { key: null, dir: 'desc' };
function aplicarOrdenacaoRec(lista) {
    if (!recSort.key) return lista; // sem ordenação de coluna: mantém a ordem padrão do filtro
    const isTexto = COLS_TEXTO_DETALHE.includes(recSort.key);
    const fator = recSort.dir === 'asc' ? 1 : -1;
    return lista.slice().sort((a, b) => {
        const va = valorOrdenacao(a, recSort.key);
        const vb = valorOrdenacao(b, recSort.key);
        let cmp = isTexto ? String(va).localeCompare(String(vb), 'pt-BR') : (va - vb);
        if (cmp === 0) cmp = a.material.localeCompare(b.material) || a.cd.localeCompare(b.cd);
        return fator * cmp;
    });
}
function ordenarRecPor(key) {
    if (recSort.key === key) {
        recSort.dir = recSort.dir === 'asc' ? 'desc' : 'asc';
    } else {
        recSort.key = key;
        recSort.dir = COLS_TEXTO_DETALHE.includes(key) ? 'asc' : 'desc';
    }
    updateRecommendations();
}
function thOrdRec(key, label, align) {
    const ativo = recSort.key === key;
    const seta = ativo
        ? `<span class="sort-arrow active">${recSort.dir === 'asc' ? '▲' : '▼'}</span>`
        : '<span class="sort-arrow">⇅</span>';
    return `<th class="sortable${ativo ? ' sort-active' : ''}" data-sort-key="${key}" style="text-align:${align};">${label} ${seta}</th>`;
}

// Lista compacta (tabela) para visão geral
// Monta a memória de cálculo de um produto.
// IMPORTANTE: vendaMedia é a VENDA MÉDIA MENSAL (Venda Média Ponderada da planilha),
// não o giro diário. Cobrir N dias = N/30 meses de venda.
// A sugestão desconta o estoque livre E o que já está a caminho (pendências de
// trânsito e entrega), que são mercadoria comprada que vai entrar no CD.
function memoriaCompra(p, metaOverride) {
    const diasMeta = (typeof metaOverride === 'number' && metaOverride > 0) ? metaOverride : getMetaDias(p);
    const mesesCobertura = diasMeta / 30;
    const necessario = Math.ceil(p.vendaMedia * mesesCobertura);
    const aCaminho = (p.pendenciaTransito || 0) + (p.pendenciaEntrega || 0);
    const disponivelFuturo = p.estoqueLivre + aCaminho;
    const comprarBruto = Math.max(0, necessario - disponivelFuturo);
    // Arredonda PARA CIMA ao múltiplo da caixa de embarque (não se compra caixa fracionada)
    const mult = ajustarAoMultiplo(p, comprarBruto);
    let comprar = mult.ajustado;   // valor final usado em todo o dashboard
    let caixas = mult.caixas;
    let pisoMinimo = false;
    // PISO DE 1 CAIXA DE EMBARQUE: produto em déficit COM necessidade líquida real
    // (comprarBruto > 0) nunca sugere menos que 1 caixa de embarque.
    // SÓ se aplica quando ainda falta comprar algo. Se o estoque livre + o que já está
    // a caminho (pendência) cobrem a meta (comprarBruto = 0), a sugestão fica 0 — não
    // faz sentido mandar comprar 1 caixa de um item já coberto pela pendência.
    if (mult.temMultiplo && p.vendaMedia > 0 && p.status === 'deficit'
        && comprarBruto > 0 && comprar < mult.multiplo) {
        comprar = mult.multiplo;
        caixas = 1;
        pisoMinimo = true;
    }
    // Média simples dos 4 meses cheios da janela (exclui o mês vigente parcial)
    const mediaMesesCheios = (p.vendas[0] + p.vendas[1] + p.vendas[2] + p.vendas[3]) / 4;
    const giroDia = p.vendaMedia / 30;
    // cobertura futura (livre + a caminho), em dias
    const diasFuturo = p.vendaMedia > 0 ? (disponivelFuturo / p.vendaMedia) * 30 : 0;
    return {
        necessario, comprar, comprarBruto, diasMeta, pisoMinimo,
        multiplo: mult.multiplo, caixas, temMultiplo: mult.temMultiplo,
        mediaMesesCheios, giroDia, mesesCobertura, aCaminho, disponivelFuturo, diasFuturo
    };
}

// Bloco de vendas por mês (reutilizado nas duas visões)
function vendasPorMesHtml(p) {
    return JANELA.map((mes, i) =>
        `<span class="mes-chip"><span class="mes-nome">${mes.abbr}${mes.parcial ? '*' : ''}</span><span class="mes-qtd">${p.vendas[i].toLocaleString('pt-BR')}</span></span>`
    ).join('');
}

// Texto da memória de cálculo da sugestão
function explicacaoCompra(p) {
    const { necessario, comprar, comprarBruto, multiplo, caixas, temMultiplo, mediaMesesCheios, giroDia, mesesCobertura, aCaminho, disponivelFuturo, diasFuturo, diasMeta, pisoMinimo } = memoriaCompra(p);
    const mesesTxt = Number.isInteger(mesesCobertura) ? `${mesesCobertura} meses` : `${mesesCobertura.toFixed(1)} meses`;

    const blocoCaminho = aCaminho > 0
        ? ` Já a caminho (trânsito + entrega): <strong>${aCaminho.toLocaleString('pt-BR')} un</strong>, elevando o disponível para <strong>${disponivelFuturo.toLocaleString('pt-BR')} un</strong> (cobertura futura ≈ ${diasFuturo.toFixed(0)} dias).`
        : '';

    // Necessidade "crua" (pode ser negativa = já coberto pelo livre + a caminho)
    const necessidadeRaw = necessario - disponivelFuturo;
    const jaCoberto = comprarBruto <= 0 && !pisoMinimo;

    const formulaBruto = jaCoberto
        ? `Necessidade líquida = ${necessario.toLocaleString('pt-BR')} − ${disponivelFuturo.toLocaleString('pt-BR')} (${p.estoqueLivre.toLocaleString('pt-BR')} livre + ${aCaminho.toLocaleString('pt-BR')} a caminho) = <strong>${necessidadeRaw.toLocaleString('pt-BR')} un</strong> → <strong style="color:var(--saudavel);">0 un a comprar</strong>. O disponível futuro já cobre a meta de ${diasMeta} dias (cobertura futura ≈ ${diasFuturo.toFixed(0)} dias).`
        : (aCaminho > 0
            ? `Necessidade líquida = ${necessario.toLocaleString('pt-BR')} − (${p.estoqueLivre.toLocaleString('pt-BR')} livre + ${aCaminho.toLocaleString('pt-BR')} a caminho) = <strong>${comprarBruto.toLocaleString('pt-BR')} un</strong>.`
            : `Necessidade líquida = ${necessario.toLocaleString('pt-BR')} − ${p.estoqueLivre.toLocaleString('pt-BR')} = <strong>${comprarBruto.toLocaleString('pt-BR')} un</strong>.`);

    // Ajuste ao múltiplo da caixa de embarque
    let blocoCaixa;
    if (jaCoberto) {
        blocoCaixa = ''; // já coberto: nada a comprar, sem arredondamento de caixa
    } else if (pisoMinimo) {
        blocoCaixa = `Caixa de embarque = <strong>${multiplo.toLocaleString('pt-BR')} un</strong>. A necessidade líquida (${comprarBruto.toLocaleString('pt-BR')} un) é menor que 1 caixa, mas o item está em déficit — aplicado o <strong>mínimo de 1 caixa</strong> = <strong style="color:var(--primary);">${comprar.toLocaleString('pt-BR')} un</strong>.`;
    } else if (temMultiplo) {
        blocoCaixa = `Caixa de embarque = <strong>${multiplo.toLocaleString('pt-BR')} un</strong>. Arredondando para cima: ${comprarBruto.toLocaleString('pt-BR')} ÷ ${multiplo.toLocaleString('pt-BR')} → <strong>${caixas.toLocaleString('pt-BR')} ${caixas === 1 ? 'caixa' : 'caixas'}</strong> = <strong style="color:var(--primary);">${comprar.toLocaleString('pt-BR')} un</strong>.`;
    } else {
        blocoCaixa = `<span style="color:var(--excesso);">⚠ SAP ${escapeHtml(String(p.codigoSAP))} sem múltiplo de caixa cadastrado — sugestão sem arredondamento: <strong>${comprar.toLocaleString('pt-BR')} un</strong>.</span>`;
    }

    return `
        <div class="rec-memoria">
            <div class="rec-memoria-vendas">
                <span class="rec-memoria-label">Vendas por mês (un):</span>
                ${vendasPorMesHtml(p)}
            </div>
            <div class="rec-memoria-calc">
                <span class="rec-memoria-label">Por que ${comprar.toLocaleString('pt-BR')} un:</span>
                Venda média de <strong>${p.vendaMedia.toLocaleString('pt-BR')} un/mês</strong> (média ponderada; média simples Fev–Mai = ${Math.round(mediaMesesCheios).toLocaleString('pt-BR')} un/mês ≈ ${giroDia.toFixed(0)} un/dia).
                Para cobrir <strong>${diasMeta} dias (${mesesTxt})</strong>: ${p.vendaMedia.toLocaleString('pt-BR')} × ${mesesCobertura.toFixed(mesesCobertura % 1 ? 1 : 0)} = <strong>${necessario.toLocaleString('pt-BR')} un</strong> necessárias.
                Estoque livre atual <strong>${p.estoqueLivre.toLocaleString('pt-BR')} un</strong>.${blocoCaminho}
                ${formulaBruto}
                ${blocoCaixa}
                <span class="jun-nota">* ${mesVigenteNome()} parcial até dia ${dashboardData.diasCorridos} (projeção mês cheio ≈ ${Math.round(p.vendaParcialProjetada).toLocaleString('pt-BR')} un).</span>
            </div>
        </div>`;
}

function renderRecTable(lista, filter, hideCD) {
    const isExcesso = filter === 'performance';
    const baseCols = hideCD ? 8 : 9;
    const colspan = isExcesso ? baseCols : baseCols + 1; // +1 da coluna Meta (oculta em excesso)
    const rows = lista.map((p, idx) => {
        const dias = Math.round(p.diasLivre);
        const { comprar, aCaminho, caixas, multiplo, temMultiplo } = memoriaCompra(p);
        const caixasTxt = (temMultiplo && caixas > 0)
            ? `<div style="font-size:0.72rem;color:var(--text-secondary);">${caixas} ${caixas === 1 ? 'cx' : 'cx'} × ${multiplo}</div>`
            : (temMultiplo ? '' : `<div style="font-size:0.7rem;color:var(--excesso);">sem múltiplo</div>`);
        const acaoCol = isExcesso
            ? `<span style="color:var(--excesso);font-weight:600;">${p.estoqueLivre.toFixed(0)} un</span>`
            : `<strong style="color:var(--primary);">${comprar.toLocaleString('pt-BR')} un</strong>${caixasTxt}`;
        const caminhoCol = aCaminho > 0
            ? `<span style="color:#0891b2;font-weight:600;">${aCaminho.toLocaleString('pt-BR')}</span>`
            : '<span style="color:#cbd5e1;">—</span>';
        const metaCol = isExcesso ? '' : `<td style="text-align:center;" onclick="event.stopPropagation()">
                <input type="number" class="row-meta-input" min="${META_DIAS_MIN}" max="${META_DIAS_MAX}" step="1"
                    value="${getMetaDias(p)}" data-rec-idx="${idx}" title="Dias de cobertura desta linha">
            </td>`;
        return `
        <tr class="rec-row" onclick="toggleRecDetail(${idx})">
            <td>${idx + 1}</td>
            ${hideCD ? '' : `<td><strong>${p.cd}</strong></td>`}
            <td title="${escapeHtml(rotuloProduto(p))}">${truncate(rotuloProduto(p), 34)} <span class="rec-expand" id="recexp-${idx}">▸</span></td>
            <td><span class="curva-badge curva-${p.curva}">${p.curva}</span></td>
            <td style="text-align:right;">${p.vendaMedia.toLocaleString('pt-BR')}</td>
            <td style="text-align:right;">${p.estoqueLivre.toFixed(0)}</td>
            <td style="text-align:right;">${caminhoCol}</td>
            <td style="text-align:right;"><span class="status-badge ${statusGrupo(p.status)}" style="padding:0.2rem 0.5rem;">${dias}</span></td>
            ${metaCol}
            <td style="text-align:right;" id="recsug-${idx}">${acaoCol}</td>
        </tr>
        <tr class="rec-detail-row" id="recdet-${idx}" style="display:none;">
            <td colspan="${colspan}">${explicacaoCompra(p)}</td>
        </tr>`;
    }).join('');

    return `
    <div class="rec-table-hint">Clique em qualquer linha para ver a memória de cálculo da sugestão. "A caminho" = trânsito + entrega (já descontado da compra).${isExcesso ? '' : ' Edite a coluna <strong>Meta</strong> para recalcular a compra de uma linha específica.'}</div>
    <div class="table-container">
        <table class="products-table rec-table">
            <thead>
                <tr>
                    <th>#</th>
                    ${hideCD ? '' : thOrdRec('cd', 'CD', 'left')}
                    ${thOrdRec('produto', 'Produto', 'left')}
                    ${thOrdRec('curva', 'Curva', 'left')}
                    ${thOrdRec('venda', 'Venda/mês', 'right')}
                    ${thOrdRec('estoque', 'Estoque', 'right')}
                    ${thOrdRec('pendencia', 'A caminho', 'right')}
                    ${thOrdRec('dias', 'Dias', 'right')}
                    ${isExcesso ? '' : '<th style="text-align:center;">Meta<br><span style="font-weight:400;font-size:0.7rem;color:var(--text-secondary);">dias</span></th>'}
                    ${isExcesso ? '<th style="text-align:right;">Parado</th>' : thOrdRec('sugestao', 'Comprar', 'right')}
                </tr>
            </thead>
            <tbody>${rows}</tbody>
        </table>
    </div>`;
}

// Recalcula uma linha da aba Sugestões quando o usuário muda a meta de dias
function onRecMetaChange(idx, valor) {
    const p = recListaAtual[idx];
    if (!p) return;
    const v = setMetaDiasLinha(p, valor);
    const inp = document.querySelector(`.row-meta-input[data-rec-idx="${idx}"]`);
    if (inp) inp.value = v;
    const memo = memoriaCompra(p);
    const caixasTxt = (memo.temMultiplo && memo.caixas > 0)
        ? `<div style="font-size:0.72rem;color:var(--text-secondary);">${memo.caixas} cx × ${memo.multiplo}</div>`
        : (memo.temMultiplo ? '' : `<div style="font-size:0.7rem;color:var(--excesso);">sem múltiplo</div>`);
    const sugCell = document.getElementById('recsug-' + idx);
    if (sugCell) sugCell.innerHTML = `<strong style="color:var(--primary);">${memo.comprar.toLocaleString('pt-BR')} un</strong>${caixasTxt}`;
    // Se a memória de cálculo dessa linha estiver aberta, atualiza-a também
    const det = document.getElementById('recdet-' + idx);
    if (det && det.style.display !== 'none') {
        const cell = det.querySelector('td');
        if (cell) cell.innerHTML = explicacaoCompra(p);
    }
    atualizarTotaisRec();
}

// Atualiza o resumo (total un / caixas) no topo da aba Sugestões após editar metas
function atualizarTotaisRec() {
    const introEl = document.querySelector('#recommendationsContainer .rec-intro .rec-total');
    if (!introEl) return;
    const totalComprar = recListaAtual.reduce((s, p) => s + memoriaCompra(p).comprar, 0);
    const totalCaixas = recListaAtual.reduce((s, p) => { const m = memoriaCompra(p); return s + (m.temMultiplo ? (m.caixas || 0) : 0); }, 0);
    const caixasTxt = totalCaixas > 0 ? ` (${totalCaixas.toLocaleString('pt-BR')} caixas)` : '';
    introEl.innerHTML = `<strong>${totalComprar.toLocaleString('pt-BR')} un</strong>${caixasTxt}`;
}

// Alterna a linha de detalhe na tabela de sugestões
window.toggleRecDetail = function (idx) {
    const row = document.getElementById('recdet-' + idx);
    const arrow = document.getElementById('recexp-' + idx);
    if (!row) return;
    const aberto = row.style.display !== 'none';
    row.style.display = aberto ? 'none' : 'table-row';
    if (arrow) arrow.textContent = aberto ? '▸' : '▾';
};

// ============================================
// DETALHES DO PRODUTO
// ============================================

// ============================================
// DETALHES DO PRODUTO (multi-seleção de CDs e Produtos)
// ============================================

// Monta a lista de checkboxes de CD no dropdown
// ============================================================
// RITMO ENTRADA × SAÍDA
// ------------------------------------------------------------
// Entrada = estoque livre + reposição a caminho (trânsito + entrega) + recência
// da última entrada. Saída = venda média + tendência do giro (meses fechados).
// Reusa a cobertura oficial (estoque ÷ venda × 30 = coluna AM) e só troca o
// denominador conforme a pergunta. O gap entre as coberturas é o sinal de ritmo.
// Sem threshold novo "do além": usa as faixas que já existem (45 / 100).
// ============================================================

// Dias entre hoje e uma data 'dd/mm/aaaa' (ou ''); null se vazia/inválida.
function diasDesde(dataStr) {
    const m = String(dataStr || '').match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
    if (!m) return null;
    const ano = m[3].length <= 2 ? 2000 + (+m[3]) : +m[3];
    const dt = new Date(ano, (+m[2]) - 1, +m[1]);
    return isNaN(dt) ? null : Math.max(0, Math.round((Date.now() - dt.getTime()) / 86400000));
}

// dd/mm/aaaa (ou -, .) -> aaaammdd numérico, p/ comparar datas de verdade
// (comparar como texto engana: ordena pelo dia, não pela data). 0 se inválida.
function chaveData(s) {
    const m = String(s || '').match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
    if (!m) return 0;
    const ano = m[3].length <= 2 ? 2000 + (+m[3]) : +m[3];
    return ano * 10000 + (+m[2]) * 100 + (+m[1]);
}

// Racional de fluxo de um ponto (CD+produto, reserva Regular já consolidada).
function calcularRitmo(p) {
    const PISO = DIAS_MIN_SAUDAVEL;   // 45 — abaixo = cobertura magra
    const TETO = DIAS_ENCALHADO;      // 100 — acima = estoque travado
    const vm = p.vendaMedia || 0;
    const aCaminho = (p.pendenciaTransito || 0) + (p.pendenciaEntrega || 0);

    // Saída: nível + tendência pelos MESES FECHADOS da janela (a última posição é
    // o mês vigente PARCIAL e fica fora). Negativos (devolução/estorno) viram 0.
    const v = (p.vendas || [0, 0, 0, 0, 0]).map(x => Math.max(_num(x), 0));
    const sBase = (v[0] + v[1]) / 2;   // 2 meses mais antigos fechados
    const sRec = (v[2] + v[3]) / 2;    // 2 meses recentes fechados
    let tendencia, varTend;
    if (vm < MIN_GIRO_TENDENCIA || sBase <= 0) {
        tendencia = 'baixo'; varTend = null;   // giro baixo/instável: não classifica
    } else {
        varTend = sRec / sBase - 1;
        tendencia = varTend >= BANDA_TENDENCIA ? 'acelerando'
            : varTend <= -BANDA_TENDENCIA ? 'desacelerando'
                : 'estavel';
    }

    // Coberturas (mesma fórmula da AM; só muda o denominador)
    const cobAtual = p.diasLivre || 0;                                       // oficial (= coluna AM)
    const cobCaminho = vm > 0 ? (p.estoqueLivre + aCaminho) / vm * 30 : 0;    // com o que vem a caminho
    const cobRecente = sRec > 0 ? (p.estoqueLivre / sRec) * 30                // se o ritmo recente seguir
        : (p.estoqueLivre > 0 ? Infinity : 0);

    // Camada 1: estrutura do estoque
    let estrutura;
    if (p.estoqueLivre > 0) estrutura = 'com-estoque';
    else if (aCaminho > 0) estrutura = 'em-fluxo';   // zerado, mas reposição a caminho
    else if (vm > 0) estrutura = 'ruptura';          // zerado, nada vindo, mas vende
    else estrutura = 'sem-giro';                     // sem venda: cobertura indefinida

    // Camada 3: veredito (acima / no ritmo / abaixo)
    let veredito;
    if (estrutura === 'sem-giro') veredito = 'sem-giro';
    else if (estrutura === 'ruptura') veredito = 'acima';
    else if (estrutura === 'em-fluxo') veredito = cobCaminho < PISO ? 'acima' : 'ritmo';
    else veredito = cobAtual < PISO ? 'acima' : cobAtual <= TETO ? 'ritmo' : 'abaixo';

    const diasSemEntrada = diasDesde(p.dataUltimaEntrada);
    let prioridade = 0;
    if (veredito === 'acima')
        prioridade = estrutura === 'ruptura' ? 100
            : tendencia === 'acelerando' ? 90
                : tendencia === 'estavel' ? 80 : 70;
    else if (veredito === 'abaixo')
        prioridade = tendencia === 'desacelerando' ? 40 : 30;
    else if (veredito === 'ritmo') prioridade = 10;

    return {
        estrutura, tendencia, varTend, veredito, sBase, sRec, aCaminho,
        cobAtual, cobCaminho, cobRecente, diasSemEntrada, prioridade
    };
}

const RITMO_VEREDITO_LABEL = { acima: 'Demanda acima', ritmo: 'No ritmo', abaixo: 'Demanda abaixo', 'sem-giro': 'Sem giro' };
const RITMO_TEND_LABEL = { acelerando: 'Acelerando', estavel: 'Estável', desacelerando: 'Desacelerando', baixo: 'Giro baixo' };
const RITMO_TEND_SETA = { acelerando: '▲', estavel: '▬', desacelerando: '▼', baixo: '·' };

// Lista atualmente exibida na aba Ritmo (para a memória por linha)
let ritmoListaAtual = [];

function updateRitmo() {
    if (!dashboardData) return;
    const container = document.getElementById('ritmoContainer');
    if (!container) return;
    const kpisEl = document.getElementById('ritmoKpis');
    const foco = (ritmoFilter && ritmoFilter.value) || 'acima';
    // Um CD específico selecionado ⇒ cai para "Por CD", a menos que SC tenha sido escolhido de propósito.
    if (ritmoView === 'sc' && filtroCD.size && !ritmoScExplicito) {
        ritmoView = 'cd';
        if (ritmoViewSel) ritmoViewSel.value = 'cd';
    }
    const sc = ritmoView === 'sc';   // Santa Cruz: soma de todos os CDs por SKU

    // Base: por CD (CD+produto) ou Santa Cruz (skuProducts, já somado entre CDs).
    let prods = sc ? skuProducts : originalProducts;
    if (!sc && filtroCD.size) prods = prods.filter(p => filtroCD.has(p.cd));   // CD só vale por CD
    if (filtroCurva.size) prods = prods.filter(p => filtroCurva.has(p.curva));

    // Calcula o ritmo de todos com giro (sem venda não tem cobertura definida).
    // Já anexa a sugestão de compra (meta 45) para permitir ordenar por ela.
    const todos = prods.map(p => ({ p, r: calcularRitmo(p), compra: memoriaCompra(p, RITMO_META_DIAS).comprar || 0 }))
        .filter(o => o.r.veredito !== 'sem-giro');

    const acima = todos.filter(o => o.r.veredito === 'acima');
    const ritmo = todos.filter(o => o.r.veredito === 'ritmo');
    const abaixo = todos.filter(o => o.r.veredito === 'abaixo');
    const ruptura = acima.filter(o => o.r.estrutura === 'ruptura').length;
    const rsParado = abaixo.reduce((s, o) => s + (o.p.estoqueLivreRS || 0), 0);
    if (kpisEl) kpisEl.innerHTML = renderRitmoKpis({ acima: acima.length, ruptura, ritmo: ritmo.length, abaixo: abaixo.length, rsParado }, sc);

    let lista;
    if (foco === 'acima') lista = acima.slice();
    else if (foco === 'abaixo') lista = abaixo.slice();
    else if (foco === 'ritmo') lista = ritmo.slice();
    else lista = todos.slice();

    // Ordenação: 'auto' = recomendada por foco; senão a escolha explícita do usuário.
    const ordem = (ritmoSort && ritmoSort.value) || 'auto';
    if (ordem === 'compra') lista.sort((a, b) => (b.compra - a.compra) || (a.r.cobAtual - b.r.cobAtual));
    else if (ordem === 'nome') lista.sort((a, b) => rotuloProduto(a.p).localeCompare(rotuloProduto(b.p), 'pt-BR'));
    else if (ordem === 'reprimido') lista.sort((a, b) => (b.r.prioridade - a.r.prioridade) || (a.r.cobAtual - b.r.cobAtual));
    else if (foco === 'abaixo') lista.sort((a, b) => b.r.cobAtual - a.r.cobAtual);
    else if (foco === 'ritmo') lista.sort((a, b) => a.r.cobAtual - b.r.cobAtual);
    else lista.sort((a, b) => (b.r.prioridade - a.r.prioridade) || (a.r.cobAtual - b.r.cobAtual));

    ritmoListaAtual = lista;

    if (!lista.length) {
        container.innerHTML = '<div class="empty-state">Nenhum produto neste critério.</div>';
        return;
    }

    const noun = sc ? 'SKUs' : 'itens';
    const escopo = sc ? ('Santa Cruz (todos os CDs)' + (filtroCD.size ? ' · filtro de CD não se aplica aqui' : ''))
        : (filtroCD.size === 0 ? 'todos os CDs' : (filtroCD.size === 1 ? `CD ${[...filtroCD][0]}` : `${filtroCD.size} CDs`));
    const introMap = {
        acima: `<strong>Demanda reprimida</strong>: a venda supera o estoque livre, cobertura abaixo de ${DIAS_MIN_SAUDAVEL} dias. Risco de ruptura, prioridade de compra. A sugestão repõe até ${RITMO_META_DIAS} dias de cobertura, já descontando o que vem a caminho. Os mais reprimidos primeiro.`,
        abaixo: `<strong>Estoque sobrando</strong>: cobertura acima de ${DIAS_ENCALHADO} dias. Capital parado, avaliar frear reposição. Maior cobertura primeiro.`,
        ritmo: `<strong>No ritmo</strong>: cobertura entre ${DIAS_MIN_SAUDAVEL} e ${DIAS_ENCALHADO} dias, ou reposição a caminho recompondo.`,
        tudo: 'Todos os produtos com giro, do mais reprimido ao mais parado.'
    };
    const intro = `<div class="rec-intro">${introMap[foco] || introMap.acima}<br><strong>${lista.length} ${noun}</strong> em ${escopo}.</div>`;
    container.innerHTML = intro + renderRitmoTable(lista, sc || filtroCD.size === 1, sc);
}

function renderRitmoKpis(k, sc) {
    const totalReal = k.acima + k.ritmo + k.abaixo;
    const total = totalReal || 1;
    const noun = sc ? 'SKUs' : 'produtos';
    const seg = (n, cor, label) => {
        const w = n / total * 100;
        if (w <= 0) return '';
        const txt = w >= 12 ? `${label} · ${n.toLocaleString('pt-BR')}` : n.toLocaleString('pt-BR');
        return `<div class="ritmo-portfolio-seg" style="flex:0 0 ${w.toFixed(2)}%;background:${cor};" title="${label}: ${n.toLocaleString('pt-BR')} ${noun}">${txt}</div>`;
    };
    const leg = (cor, label, n) => `<span class="ritmo-pleg"><span class="ritmo-pleg-dot" style="background:${cor};"></span>${label}: <strong>${n.toLocaleString('pt-BR')}</strong> ${noun}</span>`;
    const cap = [];
    if (k.ruptura > 0) cap.push(`${k.ruptura.toLocaleString('pt-BR')} em ruptura (estoque zerado, nada a caminho)`);
    if (k.rsParado > 0) cap.push(`${formatBRLCheio(k.rsParado)} parado em estoque (estoque sobrando)`);
    return `<div class="ritmo-portfolio">
        <div class="ritmo-portfolio-top">${totalReal.toLocaleString('pt-BR')} ${noun} com giro, por situação</div>
        <div class="ritmo-portfolio-bar">
            ${seg(k.acima, 'var(--deficit)', 'Demanda reprimida')}
            ${seg(k.ritmo, 'var(--saudavel)', 'No ritmo')}
            ${seg(k.abaixo, 'var(--excesso)', 'Estoque sobrando')}
        </div>
        <div class="ritmo-portfolio-leg">
            ${leg('var(--deficit)', 'Demanda reprimida', k.acima)}
            ${leg('var(--saudavel)', 'No ritmo', k.ritmo)}
            ${leg('var(--excesso)', 'Estoque sobrando', k.abaixo)}
        </div>
        ${cap.length ? `<div class="ritmo-portfolio-cap">${cap.join(' · ')}</div>` : ''}
    </div>`;
}

function ritmoZona(d) { return d < DIAS_MIN_SAUDAVEL ? 'acima' : d <= DIAS_ENCALHADO ? 'ritmo' : 'abaixo'; }
const RITMO_COR = { acima: 'var(--deficit)', ritmo: 'var(--saudavel)', abaixo: 'var(--excesso)' };
const RITMO_DIAS_MAX = 150;   // teto da régua da barra (dias); acima disso a barra enche
const RITMO_TOL_REC = 5;      // só desenha a marca de ritmo recente se ela ficar a 5+ dias da cobertura atual (senão fica em cima e polui)
const RITMO_META_DIAS = 45;   // meta da sugestão de compra NESTA visão (alinhada à linha de alerta de 45 d). A aba de Sugestões segue em 60.

// Formatos compactos para a leitura direta
function fmtMult(m) {
    if (!isFinite(m)) return '∞';
    if (m >= 10) return Math.round(m).toLocaleString('pt-BR');
    if (m >= 0.1) return m.toFixed(1).replace('.', ',');
    return '<0,1';   // muito parado: evita mostrar "0,0"
}
function fmtMeses(dias) {
    const m = dias / 30;
    return m >= 10 ? Math.round(m).toLocaleString('pt-BR') : m.toFixed(1).replace('.', ',');
}

// Linha de leitura direta: bate o olho e lê o veredito + quantas vezes a venda supera o estoque.
function renderRitmoTable(lista, hideCD, sc) {
    const rows = lista.map((o, idx) => {
        const p = o.p, r = o.r;
        const cor = RITMO_COR[r.veredito];
        const cob = Math.round(r.cobAtual);
        const est = p.estoqueLivre || 0;
        const vm = p.vendaMedia || 0;

        // Número grande SEMPRE = venda ÷ estoque (quantas vezes a venda mensal supera o estoque).
        // Os dias de cobertura ficam sempre na linha do veredito, ao lado. Mesma unidade em toda linha.
        let big, verd;
        if (r.estrutura === 'ruptura') {
            big = 'sem estoque'; verd = 'ruptura, nada a caminho';
        } else if (r.estrutura === 'em-fluxo') {
            big = '0 livre';
            verd = `${r.veredito === 'ritmo' ? 'repondo' : 'reposição insuf.'} · ${Math.round(r.cobCaminho).toLocaleString('pt-BR')} d a caminho`;
        } else {
            big = `${fmtMult(vm / est)}×`;
            const verdWord = r.veredito === 'acima' ? 'demanda reprimida'
                : r.veredito === 'abaixo' ? 'estoque sobrando' : 'no ritmo';
            verd = `${verdWord} · ${cob} d`;
        }

        const cdTxt = sc
            ? `Santa Cruz · ${(p.nCDs || 1)} CD${(p.nCDs || 1) > 1 ? 's' : ''}`
            : (hideCD ? '' : escapeHtml(p.cd));
        const meta = [cdTxt, `curva ${p.curva}`].filter(Boolean).join(' · ');
        const aCam = r.aCaminho > 0 ? ` · ${r.aCaminho.toLocaleString('pt-BR')} a caminho` : '';
        const nums = `vende ${vm.toLocaleString('pt-BR')}/mês · ${est.toLocaleString('pt-BR')} em estoque${aCam}`;

        // Sugestão de compra (mesma conta da aba de Sugestões, mas com meta de 45 dias nesta visão),
        // JÁ DESCONTANDO estoque livre + o que está em trânsito e pendente de entrega.
        const mc = memoriaCompra(p, RITMO_META_DIAS);
        const compra = mc.comprar || 0;
        const compraTip = compra > 0
            ? `Comprar para ${mc.diasMeta} dias, já descontado ${est.toLocaleString('pt-BR')} livre${r.aCaminho > 0 ? ' + ' + r.aCaminho.toLocaleString('pt-BR') + ' a caminho' : ''}.`
            : (r.veredito === 'acima' && r.aCaminho > 0
                ? `O que vem a caminho (${r.aCaminho.toLocaleString('pt-BR')} un) já cobre a meta de ${mc.diasMeta} dias.`
                : `Cobertura já na meta de ${mc.diasMeta} dias.`);
        const compraHtml = `<div class="ritmo-compra${compra > 0 ? '' : ' zero'}" title="${compraTip}"><div class="ritmo-compra-n">${compra.toLocaleString('pt-BR')}</div><div class="ritmo-compra-l">Sugestão de compra</div></div>`;

        return `
        <div class="ritmo-item">
            <div class="ritmo-srow" onclick="toggleRitmoDetail(${idx})" style="border-left-color:${cor};">
                <div class="ritmo-sid">
                    <div class="ritmo-sname" title="${escapeHtml(rotuloProduto(p))}">${truncate(rotuloProduto(p), 44)}</div>
                    <div class="ritmo-smeta">${meta}</div>
                    <div class="ritmo-snums">${nums}</div>
                </div>
                <div class="ritmo-sverd">
                    <div class="ritmo-sbig" style="color:${cor};">${big}</div>
                    <div class="ritmo-sverd-l" style="color:${cor};">${verd}</div>
                </div>
                ${compraHtml}
                <span class="ritmo-exp" id="ritexp-${idx}">▸</span>
            </div>
            <div class="ritmo-detalhe" id="ritdet-${idx}" style="display:none;">${explicacaoRitmo(p, r)}</div>
        </div>`;
    }).join('');

    const dica = sc
        ? 'Somado entre todos os CDs da Santa Cruz. Clique numa linha para a memória de cálculo.'
        : 'Cada linha é um produto num CD. Clique para a memória de cálculo.';
    return `<div class="ritmo-dica">${dica}</div><div class="ritmo-list">${rows}</div>`;
}

// Memória do ritmo — compacta: o gap de cobertura lado a lado, o veredito em 1
// linha, a conta de compra inteira e onde o SKU vende. Sem repetir o que a linha
// já mostra (venda média, estoque livre e a caminho ficam no cabeçalho).
function explicacaoRitmo(p, r) {
    const fmt = n => isFinite(n) ? Math.round(n).toLocaleString('pt-BR') : '∞';
    const cor = RITMO_COR[r.veredito] || 'var(--text-secondary)';
    const dir = (r.tendencia === 'baixo' || r.varTend == null) ? null
        : r.tendencia === 'acelerando' ? 'Acelerando'
        : r.tendencia === 'desacelerando' ? 'Desacelerando'
        : 'Estável';

    // 1) Coberturas lado a lado: o gap entre elas é o sinal do ritmo.
    const cob = [];
    cob.push(`<div class="rit-cob-cell"><div class="rit-cob-n" style="color:${cor};">${fmt(r.cobAtual)}<span class="rit-cob-u">d</span></div><div class="rit-cob-l">livre hoje</div></div>`);
    if (r.aCaminho > 0)
        cob.push(`<div class="rit-cob-cell"><div class="rit-cob-n">${fmt(r.cobCaminho)}<span class="rit-cob-u">d</span></div><div class="rit-cob-l">com o que vem</div></div>`);
    if (isFinite(r.cobRecente) && r.sRec > 0 && Math.abs(r.cobRecente - r.cobAtual) >= 1)
        cob.push(`<div class="rit-cob-cell"><div class="rit-cob-n">${fmt(r.cobRecente)}<span class="rit-cob-u">d</span></div><div class="rit-cob-l" title="se o ritmo recente de ${fmt(r.sRec)} un/mês seguir">ritmo recente</div></div>`);
    if (r.diasSemEntrada != null)
        cob.push(`<div class="rit-cob-cell rit-cob-ctx"><div class="rit-cob-n">${r.diasSemEntrada}<span class="rit-cob-u">d</span></div><div class="rit-cob-l">desde a entrada</div></div>`);
    const cobHtml = `<div class="rit-cob">${cob.join('')}</div>`;

    // 2) Veredito em uma linha: palavra-chave colorida + leitura curta.
    let kw, txt;
    if (r.estrutura === 'ruptura') {
        kw = 'Ruptura'; txt = `livre zerado e nada a caminho; vende ${fmt(p.vendaMedia)}/mês.`;
    } else if (r.estrutura === 'em-fluxo' && r.veredito === 'acima') {
        kw = 'Livre zerado'; txt = `o que vem a caminho cobre só ${fmt(r.cobCaminho)} d (abaixo de ${DIAS_MIN_SAUDAVEL}). A reposição não acompanha.`;
    } else if (r.estrutura === 'em-fluxo') {
        kw = 'Em reposição'; txt = `livre zerado, mas o caminho recompõe ${fmt(r.cobCaminho)} d de cobertura.`;
    } else if (r.veredito === 'acima') {
        kw = dir || 'Reprimida'; txt = `a saída supera o estoque livre; cobertura abaixo de ${DIAS_MIN_SAUDAVEL} d.`;
    } else if (r.veredito === 'abaixo') {
        kw = dir || 'Sobrando'; txt = `estoque sobrando para o giro; cobertura acima de ${DIAS_ENCALHADO} d.`;
    } else if (r.veredito === 'sem-giro') {
        kw = 'Sem giro'; txt = `sem venda na janela; cobertura indefinida.`;
    } else {
        kw = dir || 'No ritmo'; txt = `entrada e saída equilibradas; cobertura na faixa ${DIAS_MIN_SAUDAVEL}–${DIAS_ENCALHADO} d.`;
    }
    const verdHtml = `<div class="rit-verd" style="border-left-color:${cor};"><b style="color:${cor};">${kw}</b> · ${txt}</div>`;

    // 3) Compra: a conta inteira em uma linha (já descontando livre + a caminho).
    const mc = memoriaCompra(p, RITMO_META_DIAS);
    const base = `meta ${mc.diasMeta} d → ${fmt(mc.necessario)} un · tem ${fmt(mc.disponivelFuturo)} (${fmt(p.estoqueLivre)}+${fmt(mc.aCaminho)})`;
    let compraV;
    if (mc.comprar <= 0) {
        compraV = `${base} · <strong style="color:var(--saudavel);">0 a comprar</strong>`;
    } else {
        const fecho = mc.pisoMinimo ? `<strong class="rit-buy">${fmt(mc.comprar)} un</strong> (piso 1 cx)`
            : mc.temMultiplo ? `<strong class="rit-buy">${fmt(mc.comprar)} un</strong> (${mc.caixas} cx de ${fmt(mc.multiplo)})`
                : `<strong class="rit-buy">${fmt(mc.comprar)} un</strong>`;
        compraV = `${base} · faltam ${fmt(mc.comprarBruto)} → ${fecho}`;
    }
    const compraHtml = `<div class="rit-line"><span class="rit-line-rot">Compra</span><span class="rit-line-v">${compraV}</span></div>`;

    // 4) CDs que mais vendem este SKU (com o livre de cada um).
    const topCDs = originalProducts
        .filter(o => o.sku === p.sku)
        .sort((a, b) => (b.vendaMedia || 0) - (a.vendaMedia || 0))
        .slice(0, 3);
    const cdsHtml = topCDs.length
        ? `<div class="rit-line"><span class="rit-line-rot">CDs</span><span class="rit-cds">${topCDs.map(c => `<span class="rit-cd"><b>${escapeHtml(c.cd)}</b> ${Math.round(c.vendaMedia).toLocaleString('pt-BR')}/mês · ${Math.round(c.estoqueLivre).toLocaleString('pt-BR')} livre</span>`).join('')}</span></div>`
        : '';

    return `
    <div class="rec-memoria ritmo-memoria">
        ${cobHtml}
        ${verdHtml}
        ${compraHtml}
        ${cdsHtml}
        <div class="rec-memoria-vendas"><span class="rec-memoria-label">Vendas/mês</span>${vendasPorMesHtml(p)}</div>
    </div>`;
}

window.toggleRitmoDetail = function (idx) {
    const row = document.getElementById('ritdet-' + idx);
    const exp = document.getElementById('ritexp-' + idx);
    if (!row) return;
    const aberto = row.style.display !== 'none';
    row.style.display = aberto ? 'none' : 'block';
    if (exp) exp.textContent = aberto ? '▸' : '▾';
};

function buildDetailCdOptions(cds) {
    const dd = document.getElementById('msCDDropdown');
    dd.innerHTML = cds.map(cd =>
        `<label class="ms-opt"><input type="checkbox" value="${escapeHtml(cd)}" data-mscd> <span>${escapeHtml(cd)}</span></label>`
    ).join('');
    dd.querySelectorAll('input[data-mscd]').forEach(chk => {
        chk.addEventListener('change', () => {
            if (chk.checked) detailCDs.add(chk.value); else detailCDs.delete(chk.value);
            renderDetailChips();
            updateDetailLabels();
            buildDetailProductOptions(document.getElementById('msProdSearch').value);
            renderDetailBlocks();
        });
    });
}

// Monta a lista de checkboxes de Produto (SKUs únicos), filtrada pela busca
// e limitada aos CDs selecionados (se houver).
function buildDetailProductOptions(filtro) {
    const opts = document.getElementById('msProdOptions');
    if (!dashboardData) { opts.innerHTML = ''; return; }
    const termo = (filtro || '').toLowerCase().trim();

    // SKUs disponíveis: respeita o filtro GLOBAL (CD/Curva/Status) e os CDs do próprio Detalhes
    const base = originalProducts.filter(p => dentroDoFiltroGlobal(p) && (detailCDs.size === 0 || detailCDs.has(p.cd)));
    const porSku = {};
    base.forEach(p => { if (!porSku[p.sku]) porSku[p.sku] = p; });
    let skus = Object.values(porSku);
    if (termo) {
        skus = skus.filter(p =>
            p.material.toLowerCase().includes(termo) ||
            String(p.codigoSAP).toLowerCase().includes(termo) ||
            String(p.ean).toLowerCase().includes(termo)
        );
    }
    skus.sort((a, b) => a.material.localeCompare(b.material) || String(a.codigoSAP).localeCompare(String(b.codigoSAP)));

    const limite = 300;
    const lista = skus.slice(0, limite);
    opts.innerHTML = lista.map(p => {
        const lbl = rotuloProduto(p);
        return `<label class="ms-opt"><input type="checkbox" value="${escapeHtml(p.sku)}" data-msprod${detailProducts.has(p.sku) ? ' checked' : ''}> <span title="${escapeHtml(lbl)}">${escapeHtml(lbl)}</span></label>`;
    }).join('') || '<div class="ms-empty">Nenhum produto encontrado</div>';

    if (skus.length > limite) {
        opts.innerHTML += `<div class="ms-empty">Mostrando ${limite} de ${skus.length}. Refine a busca.</div>`;
    }

    opts.querySelectorAll('input[data-msprod]').forEach(chk => {
        chk.addEventListener('change', () => {
            if (chk.checked) detailProducts.add(chk.value); else detailProducts.delete(chk.value);
            renderDetailChips();
            updateDetailLabels();
            renderDetailBlocks();
        });
    });
}

// Rótulo de um SKU a partir da sua chave (busca um ponto representativo)
function labelDoSku(sku) {
    const p = originalProducts.find(x => x.sku === sku);
    return p ? rotuloProduto(p) : sku;
}

// Atualiza os rótulos dos botões
function updateDetailLabels() {
    const cdLbl = document.getElementById('msCDLabel');
    const prodLbl = document.getElementById('msProdLabel');
    if (cdLbl) cdLbl.textContent = detailCDs.size === 0 ? 'Todos os CDs' : `${detailCDs.size} CD${detailCDs.size > 1 ? 's' : ''} selecionado${detailCDs.size > 1 ? 's' : ''}`;
    if (prodLbl) prodLbl.textContent = detailProducts.size === 0 ? 'Nenhum produto selecionado' : `${detailProducts.size} produto${detailProducts.size > 1 ? 's' : ''} selecionado${detailProducts.size > 1 ? 's' : ''}`;
}

// Renderiza os chips de seleção (com X para remover)
function renderDetailChips() {
    const cdChips = document.getElementById('msCDChips');
    const prodChips = document.getElementById('msProdChips');

    cdChips.innerHTML = [...detailCDs].sort().map(cd =>
        `<span class="af-chip">${escapeHtml(cd)} <span class="af-x" data-rmcd="${escapeHtml(cd)}">✕</span></span>`
    ).join('');
    cdChips.querySelectorAll('[data-rmcd]').forEach(x => {
        x.addEventListener('click', () => {
            detailCDs.delete(x.getAttribute('data-rmcd'));
            const chk = document.querySelector(`#msCDDropdown input[value="${cssEscape(x.getAttribute('data-rmcd'))}"]`);
            if (chk) chk.checked = false;
            renderDetailChips(); updateDetailLabels();
            buildDetailProductOptions(document.getElementById('msProdSearch').value);
            renderDetailBlocks();
        });
    });

    // Produtos: resumo compacto + lista recolhível (evita ocupar muito espaço com 20+ itens)
    if (detailProducts.size === 0) {
        prodChips.innerHTML = '';
        return;
    }
    const skus = [...detailProducts];
    const chipsHtml = skus.map(sku =>
        `<span class="af-chip">${escapeHtml(truncate(labelDoSku(sku), 28))} <span class="af-x" data-rmprod="${escapeHtml(sku)}">✕</span></span>`
    ).join('');
    prodChips.innerHTML = `
        <div class="chips-summary">
            <span class="chips-count">${skus.length} produto${skus.length > 1 ? 's' : ''} selecionado${skus.length > 1 ? 's' : ''}</span>
            <button type="button" class="chips-toggle" aria-expanded="${prodChipsAberto}">${prodChipsAberto ? 'ocultar lista ▴' : 'ver lista ▾'}</button>
            <button type="button" class="chips-clear-all">limpar todos</button>
        </div>
        <div class="chips-list" style="display:${prodChipsAberto ? 'flex' : 'none'};">${chipsHtml}</div>`;

    prodChips.querySelector('.chips-toggle').addEventListener('click', () => {
        prodChipsAberto = !prodChipsAberto;
        renderDetailChips();
    });
    prodChips.querySelector('.chips-clear-all').addEventListener('click', () => {
        detailProducts.clear();
        document.querySelectorAll('#msProdOptions input:checked').forEach(c => { c.checked = false; });
        renderDetailChips(); updateDetailLabels(); renderDetailBlocks();
    });
    prodChips.querySelectorAll('[data-rmprod]').forEach(x => {
        x.addEventListener('click', () => {
            detailProducts.delete(x.getAttribute('data-rmprod'));
            const chk = document.querySelector(`#msProdOptions input[value="${cssEscape(x.getAttribute('data-rmprod'))}"]`);
            if (chk) chk.checked = false;
            renderDetailChips(); updateDetailLabels(); renderDetailBlocks();
        });
    });
}

// Estado de expansão da lista de chips de produtos (recolhida por padrão)
let prodChipsAberto = false;

// Escapa um valor para uso seguro em querySelector
function cssEscape(s) {
    if (window.CSS && CSS.escape) return CSS.escape(s);
    return String(s).replace(/["\\]/g, '\\$&');
}

// Renderiza um bloco de detalhe por (produto × CD) selecionado
function renderDetailBlocks() {
    const container = document.getElementById('detailBlocks');
    const aviso = document.getElementById('noDetailSelected');
    if (!dashboardData) return;

    if (detailProducts.size === 0) {
        container.innerHTML = '';
        aviso.style.display = 'block';
        aviso.textContent = 'Selecione ao menos um produto para ver detalhes';
        return;
    }

    // Pontos correspondentes (CD+SKU) que batem com a seleção e com o filtro global
    const pontos = originalProducts.filter(p =>
        detailProducts.has(p.sku) && dentroDoFiltroGlobal(p) && (detailCDs.size === 0 || detailCDs.has(p.cd))
    ).sort((a, b) => a.material.localeCompare(b.material) || a.cd.localeCompare(b.cd));

    if (pontos.length === 0) {
        container.innerHTML = '';
        aviso.style.display = 'block';
        aviso.textContent = 'Nenhum ponto de estoque para essa combinação de CD e produto';
        return;
    }

    aviso.style.display = 'none';
    const ordenados = aplicarOrdenacaoDetalhe(pontos);
    container.innerHTML = renderDetailList(ordenados);
    // guarda os pontos atuais (já ordenados) para o toggle de expansão e edição de meta
    detailPontosAtuais = ordenados;
}

// Estado de ordenação da tabela de Detalhes (clique no cabeçalho)
// key = coluna; dir = 'asc' | 'desc'. key null = ordem padrão (produto, CD).
let detailSort = { key: null, dir: 'desc' };

// Valor de uma linha para a coluna escolhida
function valorOrdenacao(p, key) {
    switch (key) {
        case 'produto': return rotuloProduto(p).toLowerCase();
        case 'cd': return p.cd;
        case 'curva': return p.curva;
        case 'venda': return p.vendaMedia;
        case 'estoque': return p.estoqueLivre;
        case 'pendencia': return (p.pendenciaTransito || 0) + (p.pendenciaEntrega || 0);
        case 'dias': return p.diasLivre;
        case 'meta': return getMetaDias(p);
        case 'sugestao': return memoriaCompra(p).comprar;
        default: return 0;
    }
}

const COLS_TEXTO_DETALHE = ['produto', 'cd', 'curva'];

function aplicarOrdenacaoDetalhe(pontos) {
    if (!detailSort.key) return pontos;
    const isTexto = COLS_TEXTO_DETALHE.includes(detailSort.key);
    const fator = detailSort.dir === 'asc' ? 1 : -1;
    // sort estável: empates mantêm a ordem padrão (produto, CD) — útil ao agrupar por curva
    return pontos.slice().sort((a, b) => {
        const va = valorOrdenacao(a, detailSort.key);
        const vb = valorOrdenacao(b, detailSort.key);
        let cmp = isTexto ? String(va).localeCompare(String(vb), 'pt-BR') : (va - vb);
        if (cmp === 0) cmp = a.material.localeCompare(b.material) || a.cd.localeCompare(b.cd);
        return fator * cmp;
    });
}

// Clique no cabeçalho: define/inverte a ordenação e redesenha
function ordenarDetalhePor(key) {
    if (detailSort.key === key) {
        detailSort.dir = detailSort.dir === 'asc' ? 'desc' : 'asc';
    } else {
        detailSort.key = key;
        // colunas numéricas começam do maior para o menor; texto começa A→Z
        detailSort.dir = COLS_TEXTO_DETALHE.includes(key) ? 'asc' : 'desc';
    }
    renderDetailBlocks();
}

// Pontos atualmente listados na aba Detalhes (para expandir sob demanda)
let detailPontosAtuais = [];

// Lista compacta: giro, estoque e sugestão de compra por linha. Detalhe completo abre ao clicar.
function renderDetailList(pontos) {
    const rows = pontos.map((p, idx) => {
        const memo = memoriaCompra(p);
        const dias = Math.round(p.diasLivre);
        const caixasTxt = (memo.temMultiplo && memo.caixas > 0)
            ? `<div style="font-size:0.72rem;color:var(--text-secondary);">${memo.caixas} cx × ${memo.multiplo}</div>`
            : (memo.temMultiplo ? '' : `<div style="font-size:0.7rem;color:var(--excesso);">sem múltiplo</div>`);
        const sugCol = p.semGiro
            ? '<span style="color:#94a3b8;">—</span>'
            : `<strong style="color:var(--primary);">${memo.comprar.toLocaleString('pt-BR')} un</strong>${caixasTxt}`;
        const metaCol = p.semGiro
            ? '<span style="color:#cbd5e1;">—</span>'
            : `<input type="number" class="row-meta-input" min="${META_DIAS_MIN}" max="${META_DIAS_MAX}" step="1" value="${getMetaDias(p)}" data-det-idx="${idx}" title="Dias de cobertura desta linha">`;
        const aCaminho = (p.pendenciaTransito || 0) + (p.pendenciaEntrega || 0);
        const pendCol = aCaminho > 0
            ? `<span style="color:#0891b2;font-weight:600;">${aCaminho.toLocaleString('pt-BR')}</span>`
            : '<span style="color:#cbd5e1;">—</span>';
        return `
        <tr class="rec-row" onclick="toggleDetailRow(${idx})">
            <td>${idx + 1}</td>
            <td title="${escapeHtml(rotuloProduto(p))}">${truncate(rotuloProduto(p), 34)} <span class="rec-expand" id="detexp-${idx}">▸</span></td>
            <td><strong>${p.cd}</strong></td>
            <td><span class="curva-badge curva-${p.curva}">${p.curva}</span></td>
            <td style="text-align:right;">${p.vendaMedia.toLocaleString('pt-BR')}</td>
            <td style="text-align:right;">${p.estoqueLivre.toFixed(0)}</td>
            <td style="text-align:right;" title="Pendência de trânsito + entrega (a caminho)">${pendCol}</td>
            <td style="text-align:right;"><span class="status-badge ${statusGrupo(p.status)}" style="padding:0.2rem 0.5rem;">${dias}</span></td>
            <td style="text-align:center;" onclick="event.stopPropagation()">${metaCol}</td>
            <td style="text-align:right;" id="detsug-${idx}">${sugCol}</td>
        </tr>
        <tr class="rec-detail-row" id="detdet-${idx}" style="display:none;">
            <td colspan="10">${detailBlockHtml(p)}</td>
        </tr>`;
    }).join('');

    return `
    <div class="rec-table-hint">Clique no cabeçalho para ordenar (▼ maior→menor, ▲ menor→maior; clique em <strong>Curva</strong> para agrupar A/B/C). Clique numa linha para abrir os detalhes. Edite <strong>Meta</strong> para recalcular a sugestão. <strong>Pendência</strong> = trânsito + entrega.</div>
    <div class="table-container">
        <table class="products-table rec-table">
            <thead>
                <tr>
                    <th>#</th>
                    ${thOrd('produto', 'Produto', 'left')}
                    ${thOrd('cd', 'CD', 'left')}
                    ${thOrd('curva', 'Curva', 'left')}
                    ${thOrd('venda', 'Venda/mês', 'right')}
                    ${thOrd('estoque', 'Estoque', 'right')}
                    ${thOrd('pendencia', 'Pendência', 'right')}
                    ${thOrd('dias', 'Dias', 'right')}
                    ${thOrd('meta', 'Meta<br><span style="font-weight:400;font-size:0.7rem;color:var(--text-secondary);">dias</span>', 'center')}
                    ${thOrd('sugestao', 'Sugestão', 'right')}
                </tr>
            </thead>
            <tbody>${rows}</tbody>
        </table>
    </div>`;
}

// Monta um <th> clicável para ordenar a tabela de Detalhes, com indicador de seta
function thOrd(key, label, align) {
    const ativo = detailSort.key === key;
    const seta = ativo
        ? `<span class="sort-arrow active">${detailSort.dir === 'asc' ? '▲' : '▼'}</span>`
        : '<span class="sort-arrow">⇅</span>';
    return `<th class="sortable${ativo ? ' sort-active' : ''}" data-sort-key="${key}" style="text-align:${align};">${label} ${seta}</th>`;
}

// Recalcula uma linha da aba Detalhes quando o usuário muda a meta de dias
function onDetMetaChange(idx, valor) {
    const p = detailPontosAtuais[idx];
    if (!p) return;
    const v = setMetaDiasLinha(p, valor);
    const inp = document.querySelector(`.row-meta-input[data-det-idx="${idx}"]`);
    if (inp) inp.value = v;
    const memo = memoriaCompra(p);
    const caixasTxt = (memo.temMultiplo && memo.caixas > 0)
        ? `<div style="font-size:0.72rem;color:var(--text-secondary);">${memo.caixas} cx × ${memo.multiplo}</div>`
        : (memo.temMultiplo ? '' : `<div style="font-size:0.7rem;color:var(--excesso);">sem múltiplo</div>`);
    const sugCell = document.getElementById('detsug-' + idx);
    if (sugCell) sugCell.innerHTML = `<strong style="color:var(--primary);">${memo.comprar.toLocaleString('pt-BR')} un</strong>${caixasTxt}`;
    // Se o bloco completo dessa linha estiver aberto, atualiza-o também
    const det = document.getElementById('detdet-' + idx);
    if (det && det.style.display !== 'none') {
        const cell = det.querySelector('td');
        if (cell) cell.innerHTML = detailBlockHtml(p);
    }
}

// Alterna a linha de detalhe completo na lista da aba Detalhes
window.toggleDetailRow = function (idx) {
    const row = document.getElementById('detdet-' + idx);
    const arrow = document.getElementById('detexp-' + idx);
    if (!row) return;
    const aberto = row.style.display !== 'none';
    row.style.display = aberto ? 'none' : 'table-row';
    if (arrow) arrow.textContent = aberto ? '▸' : '▾';
};

// HTML completo de um bloco de detalhe para um ponto (CD+Material)
function detailBlockHtml(p) {
    const hist = JANELA.map((mes, i) => ({
        mes: mes.abbr + (mes.parcial ? '*' : ''),
        qtde: p.vendas[i],
        valor: p.vendasRS[i]
    }));

    const chartUn = createBarChartSimple(hist.map(h => h.mes), hist.map(h => h.qtde), '#3b82f6') +
        `<p class="chart-note">* ${mesVigenteNome()} parcial (até dia ${dashboardData.diasCorridos}). Projeção mês cheio: <strong>${Math.round(p.vendaParcialProjetada)} un</strong></p>`;
    const chartRS = createBarChartSimple(hist.map(h => h.mes), hist.map(h => h.valor), '#10b981', true);

    const memo = memoriaCompra(p);
    const sapTxt = String(p.codigoSAP || '').replace(/\.0$/, '').trim();

    return `
    <div class="detail-panel">
        <h3>${escapeHtml(p.material)} <span class="status-badge ${statusGrupo(p.status)}">${formatStatus(p.status)}</span>
            ${sapTxt ? `<span class="detail-sap-tag">SAP ${escapeHtml(sapTxt)}</span>` : ''}
            <span class="detail-cd-tag">${escapeHtml(p.cd)}</span></h3>
        <div class="detail-grid">
            <div class="detail-section">
                <h4>📊 Histórico de Vendas (Unidades)</h4>
                ${chartUn}
            </div>
            <div class="detail-section">
                <h4>💰 Histórico de Vendas (R$)</h4>
                ${chartRS}
            </div>
            <div class="detail-section">
                <h4>📈 Estoques</h4>
                <ul>
                    <li><span class="label">Estoque Livre:</span> <span class="value">${p.estoqueLivre.toFixed(0)} un</span></li>
                    <li><span class="label">Estoque Qualidade:</span> <span class="value">${p.estoqueQualidade.toFixed(0)} un</span></li>
                    <li><span class="label">Estoque Bloqueado:</span> <span class="value">${p.estoqueBloqueado.toFixed(0)} un</span></li>
                    <li><span class="label">Estoque Total:</span> <span class="value">${p.estoqueTotal.toFixed(0)} un</span></li>
                    <li><span class="label">Pendência Trânsito:</span> <span class="value">${p.pendenciaTransito.toFixed(0)} un</span></li>
                    <li><span class="label">Pendência Entrega:</span> <span class="value">${p.pendenciaEntrega.toFixed(0)} un</span></li>
                    <li><span class="label">Estoque Livre R$:</span> <span class="value">${formatBRL(p.estoqueLivreRS)}</span></li>
                </ul>
            </div>
            <div class="detail-section">
                <h4>⚡ Análise de Giro</h4>
                <ul>
                    <li><span class="label">Venda Média Ponderada:</span> <span class="value">${p.vendaMedia.toLocaleString('pt-BR')} un/mês</span></li>
                    <li><span class="label">Equivale a:</span> <span class="value">${memo.giroDia.toFixed(0)} un/dia</span></li>
                    <li><span class="label">Dias de Estoque (livre):</span> <span class="value">${p.diasLivre.toFixed(1)}</span></li>
                    <li><span class="label">Dias de Estoque (total):</span> <span class="value">${p.diasTotal.toFixed(1)}</span></li>
                    <li><span class="label">Curva ABC:</span> <span class="value">${p.curva}</span></li>
                    <li><span class="label">PF:</span> <span class="value">${formatBRL(p.pf)}</span></li>
                    <li><span class="label">Última Entrada:</span> <span class="value">${p.dataUltimaEntrada || '—'}</span></li>
                    <li style="border-top:2px solid var(--border);margin-top:8px;padding-top:8px;"><span class="label">Caixa de embarque:</span> <span class="value">${memo.temMultiplo ? memo.multiplo.toLocaleString('pt-BR') + ' un' : '<span style="color:var(--excesso);">não cadastrada</span>'}</span></li>
                    <li><span class="label">Sugestão p/ ${memo.diasMeta} dias:</span> <span class="value" style="color:var(--primary);font-size:1.1rem;">${memo.comprar.toLocaleString('pt-BR')} un${memo.temMultiplo && memo.caixas > 0 ? ` <span style="font-size:0.8rem;color:var(--text-secondary);">(${memo.caixas} ${memo.caixas === 1 ? 'cx' : 'cx'})</span>` : ''}</span></li>
                </ul>
            </div>
        </div>
    </div>`;
}

// Liga os toggles dos dropdowns e o fechamento ao clicar fora
function setupDetailMultiselects() {
    const cdToggle = document.getElementById('msCDToggle');
    const cdDropdown = document.getElementById('msCDDropdown');
    const prodToggle = document.getElementById('msProdToggle');
    const prodDropdown = document.getElementById('msProdDropdown');
    const prodSearch = document.getElementById('msProdSearch');

    cdToggle.addEventListener('click', (e) => {
        e.stopPropagation();
        const open = cdDropdown.classList.toggle('open');
        cdToggle.classList.toggle('active', open);
        prodDropdown.classList.remove('open');
        prodToggle.classList.remove('active');
    });

    prodToggle.addEventListener('click', (e) => {
        e.stopPropagation();
        const open = prodDropdown.classList.toggle('open');
        prodToggle.classList.toggle('active', open);
        cdDropdown.classList.remove('open');
        cdToggle.classList.remove('active');
        if (open) setTimeout(() => prodSearch.focus(), 50);
    });

    prodSearch.addEventListener('input', () => buildDetailProductOptions(prodSearch.value));
    prodSearch.addEventListener('click', (e) => e.stopPropagation());
    cdDropdown.addEventListener('click', (e) => e.stopPropagation());
    prodDropdown.addEventListener('click', (e) => e.stopPropagation());

    // Fecha dropdowns ao clicar fora
    document.addEventListener('click', () => {
        cdDropdown.classList.remove('open'); cdToggle.classList.remove('active');
        prodDropdown.classList.remove('open'); prodToggle.classList.remove('active');
    });
}

// Mini gráfico de barras para o painel de detalhes
function createBarChartSimple(labels, values, color, isMoney) {
    const maxV = Math.max(...values, 1);
    const w = 320, h = 180, padB = 30, padT = 24, padL = 10;
    const slot = (w - padL * 2) / labels.length;
    const plotH = h - padB - padT;

    let bars = '';
    values.forEach((v, i) => {
        const bh = (v / maxV) * plotH;
        const x = padL + i * slot + slot * 0.2;
        const y = padT + plotH - bh;
        const bw = slot * 0.6;
        bars += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bw.toFixed(1)}" height="${bh.toFixed(1)}" fill="${color}" opacity="0.85" rx="3"/>`;
        bars += `<text x="${(x + bw / 2).toFixed(1)}" y="${h - padB + 16}" text-anchor="middle" font-size="11" fill="#64748b">${labels[i]}</text>`;
        const lbl = isMoney ? formatBRLShort(v) : Math.round(v).toLocaleString('pt-BR');
        bars += `<text x="${(x + bw / 2).toFixed(1)}" y="${(y - 5).toFixed(1)}" text-anchor="middle" font-size="9.5" font-weight="bold" fill="#1e293b">${lbl}</text>`;
    });

    return `<svg width="100%" height="${h}" viewBox="0 0 ${w} ${h}" style="max-width:100%;">
        <line x1="${padL}" y1="${padT + plotH}" x2="${w - padL}" y2="${padT + plotH}" stroke="#e2e8f0" stroke-width="1"/>
        ${bars}
    </svg>`;
}

// ============================================
// ABA: VENDAS DO CD (histórico de saída mês a mês)
// ============================================
// Soma as vendas (unidades e R$) de todos os produtos do filtro global,
// mês a mês. Mostra o total da distribuidora sempre; o detalhe de um CD
// específico abre só quando escolhido no seletor.
let cdSalesPorCD = {};
function renderVendasCD() {
    const totalBox = document.getElementById('cdSalesTotal');
    const sel = document.getElementById('cdSalesSelect');
    const detBox = document.getElementById('cdSalesDetalhe');
    if (!totalBox || !sel || !detBox) return;
    if (!dashboardData) {
        totalBox.innerHTML = '<div class="empty-state">Carregue uma planilha para ver o histórico de vendas.</div>';
        sel.innerHTML = '<option value="">Selecione um CD...</option>';
        detBox.innerHTML = '';
        cdSalesPorCD = {};
        return;
    }
    const produtos = dashboardData.products.filter(p => dentroDoFiltroGlobal(p));
    if (!produtos.length) {
        totalBox.innerHTML = '<div class="empty-state">Nenhum produto no filtro atual.</div>';
        sel.innerHTML = '<option value="">Selecione um CD...</option>';
        detBox.innerHTML = '';
        cdSalesPorCD = {};
        return;
    }
    const totalUn = [0, 0, 0, 0, 0], totalRS = [0, 0, 0, 0, 0];
    const porCD = {};
    produtos.forEach(p => {
        for (let i = 0; i < 5; i++) { totalUn[i] += p.vendas[i]; totalRS[i] += p.vendasRS[i]; }
        const cd = p.cd || '—';
        if (!porCD[cd]) porCD[cd] = { un: [0, 0, 0, 0, 0], rs: [0, 0, 0, 0, 0] };
        for (let i = 0; i < 5; i++) { porCD[cd].un[i] += p.vendas[i]; porCD[cd].rs[i] += p.vendasRS[i]; }
    });
    cdSalesPorCD = porCD;
    const cds = Object.keys(porCD);
    const tituloTotal = cds.length === 1 ? cds[0] : 'Total da Distribuidora';
    totalBox.innerHTML = blocoVendasCD(tituloTotal, totalUn, totalRS, true);

    // Opções do seletor, em ordem alfabética de CD (média mensal no rótulo)
    const volume = cd => porCD[cd].un.slice(0, 4).reduce((x, y) => x + y, 0);
    const ord = cds.slice().sort((a, b) => a.localeCompare(b, 'pt-BR'));
    const anterior = sel.value;
    sel.innerHTML = '<option value="">Selecione um CD para ver o histórico individual...</option>' +
        ord.map(cd => `<option value="${escapeHtml(cd)}">${escapeHtml(cd)} — média ${Math.round(volume(cd) / 4).toLocaleString('pt-BR')} un/mês</option>`).join('');
    if (anterior && porCD[anterior]) sel.value = anterior; // preserva escolha entre atualizações
    renderVendasCDDetalhe();
}

// Abre o bloco do CD selecionado (ou uma dica, se nenhum)
function renderVendasCDDetalhe() {
    const sel = document.getElementById('cdSalesSelect');
    const detBox = document.getElementById('cdSalesDetalhe');
    if (!sel || !detBox) return;
    const cd = sel.value;
    if (!cd || !cdSalesPorCD[cd]) {
        detBox.innerHTML = '<p class="cd-sales-hint">Selecione um CD acima para abrir o histórico individual dele.</p>';
        return;
    }
    detBox.innerHTML = blocoVendasCD(cd, cdSalesPorCD[cd].un, cdSalesPorCD[cd].rs, false);
}

function blocoVendasCD(titulo, un, rs, destaque) {
    const labels = JANELA.map(m => m.abbr + (m.parcial ? '*' : ''));
    const dias = dashboardData.diasCorridos || 0;
    const parcial = un[4];
    const proj = dias > 0 ? Math.round(parcial / dias * 30) : Math.round(parcial);
    const media4 = Math.round((un[0] + un[1] + un[2] + un[3]) / 4);
    const mediaRS4 = (rs[0] + rs[1] + rs[2] + rs[3]) / 4;
    const chartUn = createBarChartSimple(labels, un.map(v => Math.round(v)), '#3b82f6') +
        `<p class="chart-note">* ${mesVigenteNome()} parcial (até dia ${dias}). Projeção mês cheio: <strong>${proj.toLocaleString('pt-BR')} un</strong></p>`;
    const chartRS = createBarChartSimple(labels, rs, '#10b981', true);
    return `
    <div class="detail-panel cd-sales-panel${destaque ? ' cd-sales-total' : ''}">
        <h3>${escapeHtml(titulo)}
            <span class="cd-sales-meta">média ${media4.toLocaleString('pt-BR')} un/mês • ${formatBRL(mediaRS4)}/mês nos 4 meses cheios</span></h3>
        <div class="detail-grid">
            <div class="detail-section"><h4>📊 Histórico de Vendas (Unidades)</h4>${chartUn}</div>
            <div class="detail-section"><h4>💰 Histórico de Vendas (R$)</h4>${chartRS}</div>
        </div>
    </div>`;
}

// ============================================
// ABAS
// ============================================

function handleTabClick(e) {
    const tab = e.target.getAttribute('data-tab');
    tabButtons.forEach(b => b.classList.remove('active'));
    tabContents.forEach(c => c.classList.remove('active'));
    e.target.classList.add('active');
    document.getElementById(tab).classList.add('active');
    if (tab === 'cd-sales') renderVendasCD();
    if (tab === 'ritmo') updateRitmo();
}

// ============================================
// UTILIDADES
// ============================================

function truncate(str, len) {
    str = String(str || '');
    return str.length > len ? str.substring(0, len) + '…' : str;
}

function escapeHtml(str) {
    return String(str || '').replace(/[&<>"']/g, c => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    })[c]);
}

function formatStatus(status) {
    return { deficit: 'Compra', saudavel: 'Compra', excesso: 'Atenção', 'sem-giro': 'Problema' }[status] || status;
}

// Classe de cor pela faixa EXIBIDA (compra/atencao/problema), para o badge não sair
// com duas cores no "Compra" (déficit vermelho + saudável verde).
function statusGrupo(status) {
    return { deficit: 'compra', saudavel: 'compra', excesso: 'atencao', 'sem-giro': 'problema' }[status] || status;
}

// ============================================
// EXPORTAR PARA EXCEL — apenas o que está em tela (filtrado), conforme a aba ativa
// ============================================

// Retorna a lista de produtos atualmente visível, de acordo com a aba ativa e seus filtros
function produtosEmTela() {
    const tab = document.querySelector('.tab-content.active')?.id || 'overview';
    if (tab === 'recommendations') return { lista: recListaAtual, contexto: 'Sugestoes' };
    if (tab === 'details') return { lista: detailPontosAtuais, contexto: 'Detalhes' };
    // overview / products / cd-analysis: base filtrada pelos filtros de CD/Curva/Status/Busca
    const base = (allProducts.length > 0 || isFilterActive()) ? allProducts : baseAtual();
    return { lista: base, contexto: 'Produtos' };
}

// --- Estilos do Excel exportado (cores iguais ao dashboard; requer xlsx-js-style) ---
const XLS_BORDA = { style: 'thin', color: { rgb: 'FFE2E8F0' } };
const XLS_BORDAS = { top: XLS_BORDA, bottom: XLS_BORDA, left: XLS_BORDA, right: XLS_BORDA };
const XLS_HEADER = {
    fill: { patternType: 'solid', fgColor: { rgb: 'FF2563EB' } },
    font: { color: { rgb: 'FFFFFFFF' }, bold: true },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: XLS_BORDAS
};
// Status: fundo claro (tinta achatada sobre branco) + texto colorido, como o badge da tela
function xlsEstiloStatus(statusKey) {
    // Cor pela faixa EXIBIDA (Compra / Atenção / Problema), não pelo status interno,
    // para "Compra" (déficit + saudável) não sair com duas cores (vermelho e verde).
    const grupo = (statusKey === 'deficit' || statusKey === 'saudavel') ? 'compra'
        : (statusKey === 'excesso') ? 'atencao'
            : (statusKey === 'sem-giro') ? 'problema' : 'outro';
    const c = {
        'compra':   { bg: 'DBEAFE', fg: '1D4ED8' },  // azul  -> necessidade de compra (< 60)
        'atencao':  { bg: 'FDEEE7', fg: 'EA580C' },  // laranja -> atenção (60-100)
        'problema': { bg: 'FCE9E9', fg: 'DC2626' }   // vermelho -> problema (> 100)
    }[grupo] || { bg: 'FFFFFF', fg: '1E293B' };
    return {
        fill: { patternType: 'solid', fgColor: { rgb: 'FF' + c.bg } },
        font: { color: { rgb: 'FF' + c.fg }, bold: true },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: XLS_BORDAS
    };
}
// Curva: preenchimento sólido + texto branco (A azul, B roxo, C cinza), como o badge da tela
function xlsEstiloCurva(curva) {
    const bg = { 'A': '2563EB', 'B': '8B5CF6', 'C': '64748B' }[curva] || '94A3B8';
    return {
        fill: { patternType: 'solid', fgColor: { rgb: 'FF' + bg } },
        font: { color: { rgb: 'FFFFFFFF' }, bold: true },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: XLS_BORDAS
    };
}
const xlsEstiloTexto = alinha => ({ alignment: { horizontal: alinha || 'left', vertical: 'center' }, border: XLS_BORDAS });
const xlsEstiloNum = fmt => ({ numFmt: fmt, alignment: { horizontal: 'right', vertical: 'center' }, border: XLS_BORDAS });

function exportarExcel() {
    if (!dashboardData) {
        alert('Carregue uma planilha antes de exportar.');
        return;
    }
    if (typeof XLSX === 'undefined') {
        alert('Biblioteca de planilha (SheetJS) não carregada.');
        return;
    }

    const { lista, contexto } = produtosEmTela();
    if (!lista || lista.length === 0) {
        alert('Não há linhas em tela para exportar com os filtros atuais.');
        return;
    }

    // Ordem: CD → Material (mantida na exportação)
    const ordenada = lista
        .slice()
        .sort((a, b) => a.cd.localeCompare(b.cd) || a.material.localeCompare(b.material));

    const linhas = ordenada.map(p => {
        const memo = memoriaCompra(p);
        const sap = String(p.codigoSAP || '').replace(/\.0$/, '').trim();
        return {
            'CD': p.cd,
            'Material': p.material,
            'Fornecedor': p.fornecedor,
            'SAP': sap,
            'Curva': p.curva,
            'Status': formatStatus(p.status),
            'Venda Média (un/mês)': Math.round(p.vendaMedia),
            'Estoque Livre (un)': Math.round(p.estoqueLivre),
            'Estoque Livre (R$)': +(p.estoqueLivreRS || 0).toFixed(2),
            'A Caminho (un)': Math.round((p.pendenciaTransito || 0) + (p.pendenciaEntrega || 0)),
            'Dias Estoque': Math.round(p.diasLivre),
            'Meta Cobertura (dias)': memo.diasMeta,
            'Sugestão Compra (un)': memo.comprar,
            'Caixas': memo.temMultiplo ? memo.caixas : '',
            'Múltiplo Caixa': memo.temMultiplo ? memo.multiplo : ''
        };
    });

    const ws = XLSX.utils.json_to_sheet(linhas);
    ws['!cols'] = [
        { wch: 8 }, { wch: 38 }, { wch: 22 }, { wch: 10 }, { wch: 6 },
        { wch: 11 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 14 },
        { wch: 12 }, { wch: 16 }, { wch: 18 }, { wch: 9 }, { wch: 13 }
    ];
    ws['!autofilter'] = { ref: `A1:O${linhas.length + 1}` };
    ws['!rows'] = [{ hpt: 26 }]; // cabeçalho um pouco mais alto

    // --- Formatação e cores (igual ao dashboard) ---
    const FMT_INT = '#,##0', FMT_BRL = 'R$ #,##0.00';
    const COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'];

    // Cabeçalho (linha 1)
    COLS.forEach(c => { const cel = ws[c + '1']; if (cel) cel.s = XLS_HEADER; });

    // Linhas de dados (a partir da linha 2)
    ordenada.forEach((p, i) => {
        const r = i + 2;
        const set = (col, estilo) => { const ref = col + r; if (ws[ref]) ws[ref].s = estilo; };
        set('A', xlsEstiloTexto('center'));   // CD
        set('B', xlsEstiloTexto('left'));     // Material
        set('C', xlsEstiloTexto('left'));     // Fornecedor
        set('D', xlsEstiloTexto('center'));   // SAP
        set('E', xlsEstiloCurva(p.curva));    // Curva (cor)
        set('F', xlsEstiloStatus(p.status));  // Status (cor)
        set('G', xlsEstiloNum(FMT_INT));      // Venda Média
        set('H', xlsEstiloNum(FMT_INT));      // Estoque Livre (un)
        set('I', xlsEstiloNum(FMT_BRL));      // Estoque Livre (R$)
        set('J', xlsEstiloNum(FMT_INT));      // A Caminho
        set('K', xlsEstiloNum(FMT_INT));      // Dias Estoque
        set('L', xlsEstiloNum(FMT_INT));      // Meta Cobertura
        set('M', xlsEstiloNum(FMT_INT));      // Sugestão Compra
        set('N', xlsEstiloNum(FMT_INT));      // Caixas
        set('O', xlsEstiloNum(FMT_INT));      // Múltiplo Caixa
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Análise de Estoque');

    const ref = String(dashboardData.dataReferencia || '').replace(/\//g, '-');
    XLSX.writeFile(wb, `Analise_Estoque_SantaCruz_${contexto}_${ref || 'export'}.xlsx`);
}

function formatBRL(v) {
    return 'R$ ' + (v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatBRLShort(v) {
    if (v >= 1000000) return 'R$' + (v / 1000000).toFixed(1) + 'M';
    if (v >= 1000) return 'R$' + (v / 1000).toFixed(1) + 'k';
    return 'R$' + Math.round(v);
}

// KPIs: valor cheio, todos os dígitos, sem abreviar "mil"/"mi". Ex.: R$ 60.482
function formatBRLCheio(v) {
    return 'R$ ' + Math.round(v || 0).toLocaleString('pt-BR');
}

console.log('Dashboard pronto. SheetJS:', typeof XLSX !== 'undefined' ? 'carregado' : 'NÃO carregado');

// Ao abrir o link, recarrega a última planilha salva (se houver). Some apenas em "Limpar Dados".
(async function restaurarPlanilha() {
    if (typeof XLSX === 'undefined') { console.warn('SheetJS indisponível; restauração ignorada.'); return; }
    try {
        const buffer = await idbCarregar();
        if (!buffer) return;
        let nome = 'Planilha salva';
        try { nome = localStorage.getItem(LS_NOME) || nome; } catch (e) {}
        carregarPlanilha(buffer, nome);
        console.log('Planilha restaurada do armazenamento local.');
    } catch (e) {
        console.warn('Falha ao restaurar planilha salva:', e);
    }
})();
