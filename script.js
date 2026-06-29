// ============================================
// DASHBOARD DE ANÁLISE DE ESTOQUE - SANTA CRUZ
// Lê arquivos .xlsx nativamente via SheetJS (offline)
// ============================================

// CONFIGURAÇÕES DE FAIXAS DE DIAS DE ESTOQUE
const DIAS_MIN_SAUDAVEL = 45;   // abaixo disso = déficit
const DIAS_MAX_SAUDAVEL = 60;   // 60 ou mais = excesso (alinha com o filtro >= 60 dias)
const DIAS_ENCALHADO = 100;     // acima disso = encalhado (estoque parado), mesmo com giro lento
// Teto da faixa "No ritmo" SÓ da aba Ritmo Entrada × Saída. Desacoplado do DIAS_ENCALHADO
// global (100, usado em classificar() / Visão Geral / Produtos): aqui a faixa saudável é mais
// apertada. No ritmo = cobertura entre DIAS_MIN_SAUDAVEL (45) e este teto.
const RITMO_TETO_DIAS = 90;

// Aba Ritmo Entrada × Saída: variação % do giro recente vs base abaixo da qual o
// movimento é "Estável" (ruído), e giro mínimo (un/mês) para confiar na tendência.
const BANDA_TENDENCIA = 0.12;
const MIN_GIRO_TENDENCIA = 5;

// Aba Detalhes: acima deste nº de produtos, o "Selecionar todos" pede confirmação
// antes de montar todos os blocos de detalhe (cada produto × CD vira um bloco).
// É só uma trava de desempenho p/ não congelar a aba — dá p/ prosseguir no aviso.
const MAX_DETALHE_CONFIRMA = 80;

// Reserva considerada "estoque livre do CD". Só esta reserva entra nos totais de
// estoque (livre/qualidade/bloqueado/total/R$), nas pendências E nas vendas/giro.
// As demais (Colaborativa, Loja a Loja, OL Indústria) NÃO são estoque livre
// disponível e não devem inflar a velocidade: os dias de estoque livre são
// Estoque Livre (Regular) ÷ Venda (Regular), idêntico à coluna do Excel.
// Use '' (string vazia) para voltar a somar TODAS as reservas.
const RESERVA_LIVRE = 'Regular';

// Quantos dias de cobertura mirar nas sugestões de compra (PADRÃO para todas as linhas).
// 45 dias de cobertura-alvo. O cálculo já desconta o que vem a caminho (trânsito +
// entrega), pois a posição projetada de fim de mês parte do Estoque Total, que inclui
// as pendências de entrega. Ajustável por linha nas abas Sugestões e Detalhes.
const META_DIAS_COBERTURA = 45;
const META_DIAS_MIN = 7;
const META_DIAS_MAX = 365;

// Meta FIXA usada nas comparações entre bases (Cockpit, "O que mudou", deltas por CD).
// O retrato de cada base é sempre congelado nesta meta, independente da meta que estiver
// na tela, para que duas bases nunca sejam comparadas em prazos diferentes.
const META_REF_COMPARATIVO = 45;

// Tolerância do arredondamento por caixa de embarque (fração de 1 caixa).
// A compra é arredondada PARA CIMA ao múltiplo da caixa, MAS uma caixa extra só é
// contada se a sobra exceder esta fração. Ex. com 0.10: necessidade de 3,02 caixas
// (sobra 2%) → 3 caixas; 3,30 caixas (sobra 30%) → 4 caixas. Evita comprar 1 caixa
// inteira só para cobrir uma folga ínfima. Use 0 para voltar ao ceil puro.
const TOLERANCIA_CAIXA = 0.10;

// Meta de cobertura POR LINHA (CD + produto). Vazio = usa a meta global (metaDiasGlobal).
// O usuário pode sobrescrever individualmente cada linha nas abas Sugestões e Detalhes.
const metaDiasPorLinha = {};
// Meta padrão aplicada a TODAS as linhas que não têm ajuste individual. Começa em
// META_DIAS_COBERTURA e pode ser redefinida de uma vez no campo "Meta de dias para todos
// os produtos" da aba Sugestões de Compra.
let metaDiasGlobal = META_DIAS_COBERTURA;
function chaveLinha(p) { return p.cd + '||' + p.sku; }
function getMetaDias(p) {
    const v = metaDiasPorLinha[chaveLinha(p)];
    return (typeof v === 'number' && !isNaN(v)) ? v : metaDiasGlobal;
}
function setMetaDiasLinha(p, valor) {
    let v = parseInt(valor, 10);
    if (isNaN(v)) v = metaDiasGlobal;
    v = Math.max(META_DIAS_MIN, Math.min(META_DIAS_MAX, v));
    metaDiasPorLinha[chaveLinha(p)] = v;
    return v;
}
// Redefine a meta de dias de TODAS as linhas de uma vez. Zera os ajustes individuais
// para que todo o portfólio passe a seguir o novo valor; ajustes por linha podem ser
// refeitos depois na coluna Meta.
function setMetaDiasTodos(valor) {
    let v = parseInt(valor, 10);
    if (isNaN(v)) v = META_DIAS_COBERTURA;
    v = Math.max(META_DIAS_MIN, Math.min(META_DIAS_MAX, v));
    metaDiasGlobal = v;
    Object.keys(metaDiasPorLinha).forEach(k => delete metaDiasPorLinha[k]);
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
    saudavel: 'Saudável (45-60 dias)',
    reposicao: 'Em reposição (coberto pelo que vem)',
    excesso: 'Atenção (60-100 dias)',
    'sem-giro': 'Problema (> 100 dias)',
    'sem-sinal': 'Sem sinal de venda',
    'parado': 'Parado (sem giro recente)'
};
// Várias chaves apontam para o MESMO conjunto global (ex.: cd, prodCd e recCd = filtroCD).
const MS_FILTROS = {
    cd:       { set: filtroCD,     toggle: 'msCdToggle',       drop: 'msCdDrop',       opts: 'msCdOpts',       vazio: 'Todos os CDs',    rotulo: v => 'CD: ' + v,              plural: n => n + ' CDs',    onChange: aplicarTudo },
    prodCd:   { set: filtroCD,     toggle: 'msProdCdToggle',   drop: 'msProdCdDrop',   opts: 'msProdCdOpts',   vazio: 'Todos os CDs',    rotulo: v => 'CD: ' + v,              plural: n => n + ' CDs',    onChange: aplicarTudo },
    recCd:    { set: filtroCD,     toggle: 'msRecCdToggle',    drop: 'msRecCdDrop',    opts: 'msRecCdOpts',    vazio: 'Todos os CDs',    rotulo: v => 'CD: ' + v,              plural: n => n + ' CDs',    onChange: aplicarTudo },
    ritmoCd:  { set: filtroCD,     toggle: 'msRitmoCdToggle',  drop: 'msRitmoCdDrop',  opts: 'msRitmoCdOpts',  vazio: 'Todos os CDs',    rotulo: v => 'CD: ' + v,              plural: n => n + ' CDs',    onChange: ritmoCdChange },
    globalCurva: { set: filtroCurva, toggle: 'msGlobalCurvaToggle', drop: 'msGlobalCurvaDrop', opts: 'msGlobalCurvaOpts', vazio: 'Todas as Curvas', rotulo: v => 'Curva ' + v, plural: n => n + ' curvas', onChange: aplicarTudo },
    status:   { set: filtroStatus, toggle: 'msStatusToggle',   drop: 'msStatusDrop',   opts: 'msStatusOpts',   vazio: 'Todos',           rotulo: v => MS_STATUS_LABEL[v] || v, plural: n => n + ' status', onChange: aplicarTudo }
};
const MS_CD_KEYS = ['cd', 'prodCd', 'recCd', 'ritmoCd'];
const MS_CURVA_KEYS = ['curva', 'recCurva', 'ritmoCurva', 'mapaCurva'];

// Aplica e sincroniza TUDO quando qualquer filtro global muda
function aplicarTudo() {
    Object.keys(MS_FILTROS).forEach(q => msSincroniza(q)); // espelha o conjunto em todos os seletores
    applyFilters();          // Visão Geral + Produtos
    updateRecommendations(); // Sugestões
    atualizarDetalhes();     // Detalhes
    renderVendasCD();        // Vendas do CD
    updateRitmo();           // Ritmo Entrada × Saída
    updateCockpit();         // Cockpit
    updateGargalo();         // Gargalo por CD
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
    // Opções fixas de Curva (filtro global, vale para todas as abas)
    const curvas = [{ value: 'A', label: 'Curva A' }, { value: 'B', label: 'Curva B' }, { value: 'C', label: 'Curva C' }];
    msMontaOpcoes('globalCurva', curvas);
    msMontaOpcoes('status', [
        { value: 'deficit', label: MS_STATUS_LABEL.deficit },
        { value: 'saudavel', label: MS_STATUS_LABEL.saudavel },
        { value: 'reposicao', label: MS_STATUS_LABEL.reposicao },
        { value: 'excesso', label: MS_STATUS_LABEL.excesso },
        { value: 'sem-giro', label: MS_STATUS_LABEL['sem-giro'] }
    ]);
}
setupFiltrosMulti();

// Exportar para Excel
exportBtn.addEventListener('click', exportarExcel);

// Edição da meta de dias POR LINHA (delegação de eventos — a tabela é redesenhada)
document.getElementById('recommendationsContainer').addEventListener('change', (e) => {
    // Campo "Meta de dias para todos os produtos": redefine a meta de todas as linhas
    if (e.target.id === 'recMetaTodos') {
        setMetaDiasTodos(e.target.value);
        updateRecommendations();
        return;
    }
    const inp = e.target.closest('.row-meta-input');
    if (inp && inp.dataset.recIdx !== undefined) onRecMetaChange(parseInt(inp.dataset.recIdx, 10), inp.value);
});
// Ordenação por clique no cabeçalho da tabela de Sugestões
document.getElementById('recommendationsContainer').addEventListener('click', (e) => {
    // Botão "Aplicar a todos" do campo de meta em massa
    if (e.target.closest('#recMetaTodosBtn')) {
        const inp = document.getElementById('recMetaTodos');
        if (inp) { setMetaDiasTodos(inp.value); updateRecommendations(); }
        return;
    }
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
const IDB_SNAPS = 'snapshots';   // histórico compacto por data-base (para a aba "O que mudou")
const IDB_KEY = 'planilha';
const LS_NOME = 'sc_nome_arquivo';
const MAX_SNAPSHOTS = 12;         // mantém os 12 retratos mais recentes (um por data-base)

function idbOpen() {
    return new Promise((resolve, reject) => {
        // v2: acrescenta o store de snapshots (o store 'arquivo' é preservado).
        const req = indexedDB.open(IDB_DB, 2);
        req.onupgradeneeded = (e) => {
            const db = req.result;
            if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE);
            if (!db.objectStoreNames.contains(IDB_SNAPS)) db.createObjectStore(IDB_SNAPS, { keyPath: 'dataKey' });
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}
// Salva (ou sobrescreve) o retrato de uma data-base e poda o histórico ao limite.
async function snapSalvar(snap) {
    if (!snap) return;
    try {
        const db = await idbOpen();
        await new Promise((resolve, reject) => {
            const tx = db.transaction(IDB_SNAPS, 'readwrite');
            tx.objectStore(IDB_SNAPS).put(snap);
            tx.oncomplete = resolve;
            tx.onerror = () => reject(tx.error);
        });
        // Poda: mantém só os MAX_SNAPSHOTS mais recentes por dataKey
        const todos = await snapListar();
        if (todos.length > MAX_SNAPSHOTS) {
            const apagar = todos.slice(0, todos.length - MAX_SNAPSHOTS).map(s => s.dataKey);
            await new Promise((resolve) => {
                const tx = db.transaction(IDB_SNAPS, 'readwrite');
                apagar.forEach(k => tx.objectStore(IDB_SNAPS).delete(k));
                tx.oncomplete = resolve;
                tx.onerror = resolve;
            });
        }
        db.close();
    } catch (e) { console.warn('Não foi possível salvar o retrato da base:', e); }
}
// Lista todos os retratos ordenados por data-base (mais antigo → mais recente).
async function snapListar() {
    try {
        const db = await idbOpen();
        const lista = await new Promise((resolve, reject) => {
            const tx = db.transaction(IDB_SNAPS, 'readonly');
            const req = tx.objectStore(IDB_SNAPS).getAll();
            req.onsuccess = () => resolve(req.result || []);
            req.onerror = () => reject(req.error);
        });
        db.close();
        return lista.sort((a, b) => (a.dataKey < b.dataKey ? -1 : a.dataKey > b.dataKey ? 1 : 0));
    } catch (e) { console.warn('Não foi possível ler o histórico de bases:', e); return []; }
}
async function snapApagarTudo() {
    try {
        const db = await idbOpen();
        await new Promise((resolve) => {
            const tx = db.transaction(IDB_SNAPS, 'readwrite');
            tx.objectStore(IDB_SNAPS).clear();
            tx.oncomplete = resolve;
            tx.onerror = resolve;
        });
        db.close();
    } catch (e) { /* silencioso */ }
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
    sincronizarHistorico();   // monta o retrato da base, compara com a anterior e redesenha Cockpit + "O que mudou"
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
            // PROJEÇÃO UNDS (coluna AQ): venda do mês vigente extrapolada por dias úteis,
            // já calculada pelo gerador da planilha. Base da posição projetada de fim de mês.
            projecaoMes: 0,

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
        o.projecaoMes += _num(row['PROJEÇÃO UNDS']);

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

    // Venda média ponderada <= 0 não é déficit: é AUSÊNCIA DE SINAL. Acontece quando a
    // janela só tem devolução (média negativa, ex.: BA03) ou nenhuma venda na reserva
    // Regular (média zero, com ou sem estoque parado). Sem demanda positiva os "dias de
    // cobertura" ficam indefinidos (a planilha devolve 0), e 0 < 45 jogava esses pontos em
    // 'deficit' por engano. memoriaCompra já sugere 0 para eles (necessário <= 0), então
    // marcá-los como 'sem-sinal' não altera nenhuma compra: só remove o falso vermelho.
    p.status = p.vendaMedia > 0 ? classificar(p.diasLivre) : 'sem-sinal';

    // GIRO RECENTE: vendeu no último mês cheio (vendas[3]=mai) OU no parcial corrente
    // (vendas[4]=jun). Cobertura baixa (déficit/saudável) SEM giro recente não é necessidade
    // de compra: é ponto PARADO. A média ponderada ainda positiva é histórico que não se
    // repete (ex.: ATIP no RS13, vendeu fev/mar e zerou). Vira 'parado' e não dispara pedido.
    // Pontos com estoque alto seguem em excesso/encalhado (a foto já diz "sobra"); só os de
    // cobertura baixa, que disparariam compra à toa, são reclassificados aqui.
    const giroRecente = (p.vendas[3] + p.vendas[4]) > 0;
    if (p.vendaMedia > 0 && !giroRecente && (p.status === 'deficit' || p.status === 'saudavel')) {
        p.status = 'parado';
    }

    // Reconcilia o status de COMPRA com o que já está a caminho (pendência de
    // trânsito + entrega = mercadoria comprada que vai entrar no CD). Um item com
    // cobertura LIVRE baixa (déficit/saudável) mas cuja compra LÍQUIDA — depois de
    // abater o que vem — é zero não precisa de compra nova: está "em reposição".
    // Usa a MESMA conta da aba Sugestões (memoriaCompra), sem inventar critério.
    if (p.status === 'deficit' || p.status === 'saudavel') {
        const aCaminho = (p.pendenciaTransito || 0) + (p.pendenciaEntrega || 0);
        if (aCaminho > 0 && memoriaCompra(p).comprar === 0) {
            p.status = 'reposicao';
        }
    }
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
                deficit: 0, saudavel: 0, reposicao: 0, excesso: 0, semGiro: 0, semSinal: 0, parado: 0
            };
        }
        const c = cdMap[p.cd];
        c.totalProdutos++;
        c.estoqueLivreTotal += p.estoqueLivre;
        c.estoqueLivreRSTotal += p.estoqueLivreRS;
        c.vendaMediaTotal += p.vendaMedia;
        if (p.status === 'deficit') c.deficit++;
        else if (p.status === 'saudavel') c.saudavel++;
        else if (p.status === 'reposicao') c.reposicao++;
        else if (p.status === 'excesso') c.excesso++;
        else if (p.status === 'sem-giro') c.semGiro++;
        else if (p.status === 'sem-sinal') c.semSinal++;
        else if (p.status === 'parado') c.parado++;
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
    msSincroniza('globalCurva');
    msSincroniza('status');
    fileInput.value = '';
    fileName.textContent = 'Selecionar Planilha Excel';
    clearBtn.style.display = 'none';
    exportBtn.style.display = 'none';
    Object.keys(metaDiasPorLinha).forEach(k => delete metaDiasPorLinha[k]);
    metaDiasGlobal = META_DIAS_COBERTURA;
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
    msSincroniza('globalCurva'); msSincroniza('status');
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

    // Telas novas: Cockpit, O que mudou
    cmpAtual = null;
    cmpAnterior = null;
    snapApagarTudo();   // "Limpar Dados" zera também o histórico de comparação
    const ckp = document.getElementById('cockpitContent');
    if (ckp) ckp.innerHTML = '<div class="empty-state">Carregue uma planilha para abrir o cockpit.</div>';
    const mud = document.getElementById('mudancasContent');
    if (mud) mud.innerHTML = '<div class="empty-state">Carregue uma planilha para comparar com a base anterior.</div>';
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
    const statusLabels = { deficit: 'Necessidade de compra (<45 dias)', saudavel: 'Saudável (45-60 dias)', reposicao: 'Em reposição (a caminho)', excesso: 'Atenção (60-100 dias)', 'sem-giro': 'Problema (>100 dias)', 'sem-sinal': 'Sem sinal de venda', 'parado': 'Parado (sem giro recente)' };

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
    updateCockpit();
    updateGargalo();
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
    //   necessidade de compra = < 45 dias (a meta de cobertura; 45-60d já está saudável)
    //   atenção               = 60 a 100 dias (status excesso)
    //   problema              = > 100 dias (status sem-giro)
    const estTotal = prods.reduce((s, p) => s + p.estoqueLivreRS, 0);
    // Faixa por COBERTURA LIVRE < 45 dias (a meta). 45-60d fica de fora: está acima da
    // meta, é saudável. Inclui "reposicao" com cobertura livre < 45 (livre baixo, já coberto
    // pelo que vem): a faixa mede capital por cobertura livre, então pertence aqui.
    const estCompra = prods.filter(p => (p.status === 'deficit' || p.status === 'saudavel' || p.status === 'reposicao') && p.diasLivre < DIAS_MIN_SAUDAVEL)
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

    // 4o card: Necessidade de compra (< 45 dias de cobertura — a meta)
    document.getElementById('kpiCds').textContent = formatBRLCheio(estCompra);
    document.getElementById('kpiCdsLabel').innerHTML = `Necessidade de compra (&lt; 45 dias) &middot; ${pctCompra.toFixed(0)}% do capital`;
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
    tbody.innerHTML = prods.slice(0, limite).map(p => {
        const aCaminho = (p.pendenciaTransito || 0) + (p.pendenciaEntrega || 0);
        return `
        <tr>
            <td><strong>${isSku ? '🌐 ' + p.nCDs + ' CDs' : p.cd}</strong></td>
            <td title="${escapeHtml(rotuloProduto(p))}">${truncate(rotuloProduto(p), 36)}</td>
            <td>${aCaminho > 0 ? aCaminho.toLocaleString('pt-BR') : '<span style="color:var(--text-secondary);">—</span>'}</td>
            <td><span class="curva-badge curva-${p.curva}">${p.curva}</span></td>
            <td>${fmtUn(p.vendaMedia, 1)}</td>
            <td>${fmtUn(p.estoqueLivre)}</td>
            <td><strong>${fmtUn(p.diasLivre)}</strong></td>
            <td><span class="status-badge ${statusGrupo(p.status)}">${formatStatus(p.status)}</span></td>
        </tr>
    `;
    }).join('');

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
                <span class="status-badge saudavel" title="Saudável (45-60 dias)">${c.saudavel} ok</span>
                ${c.reposicao > 0 ? `<span class="status-badge reposicao" title="Em reposição: estoque livre baixo, mas já coberto pelo que está a caminho">${c.reposicao} reposição</span>` : ''}
                <span class="status-badge excesso" title="Atenção (60-100 dias)">${c.excesso} atenção</span>
                ${c.semGiro > 0 ? `<span class="status-badge sem-giro" title="Problema (>100 dias)">${c.semGiro} problema</span>` : ''}
                ${c.semSinal > 0 ? `<span class="status-badge sem-sinal" title="Sem sinal de venda (média ponderada <= 0)">${c.semSinal} sem sinal</span>` : ''}
                ${c.parado > 0 ? `<span class="status-badge parado" title="Parado: vendeu na janela mas sem giro no último mês cheio nem no parcial">${c.parado} parado</span>` : ''}
            </div>
        </div>`;
    }).join('');
}

// ============================================
// SUGESTÕES DE COMPRA
// ============================================

function updateRecommendations() {
    if (!dashboardData) return;
    const filter = recFilter.value || 'todos';
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
        intro = `Produtos que precisam de compra para sustentar a meta de ${metaDiasGlobal} dias de cobertura.`;
        // Inclui "reposicao" (déficit/saudável já coberto pelo que vem): assim eles
        // seguem entrando na conta de "já com reposição a caminho" (jaRepostos) e saem
        // da lista de compra pelo filtro comprar>0 logo abaixo — comportamento idêntico.
        const emDeficit = prods.filter(p => p.status === 'deficit' || p.status === 'saudavel' || p.status === 'reposicao');
        lista = emDeficit.filter(p => memoriaCompra(p).comprar > 0)
            .sort((a, b) => a.diasLivre - b.diasLivre);
        jaRepostos = emDeficit.length - lista.length;
    } else if (filter === 'urgent') {
        intro = 'Compra urgente: cobertura abaixo de 30 dias e sem reposição suficiente a caminho.';
        const criticos = prods.filter(p => Math.round(p.diasLivre) < 30);
        lista = criticos.filter(p => memoriaCompra(p).comprar > 0)
            .sort((a, b) => (b.vendaMedia / (b.diasLivre + 1)) - (a.vendaMedia / (a.diasLivre + 1)));
        jaRepostos = criticos.length - lista.length;
    } else if (filter === 'performance') {
        intro = 'Estoque parado: atenção (60-100 dias) e problema (&gt;100 dias). Avaliar redução ou remanejamento.';
        lista = prods.filter(p => p.status === 'excesso' || p.status === 'sem-giro')
            .sort((a, b) => {
                const da = a.diasLivre;
                const db = b.diasLivre;
                return db - da;
            });
    } else if (filter === 'todos') {
        intro = 'Todos os produtos do escopo, do menor para o maior em cobertura. Itens já cobertos aparecem com 0 un.';
        lista = prods.slice().sort((a, b) => a.diasLivre - b.diasLivre);
        jaRepostos = 0;
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
    // Conta "sem múltiplo" só entre os itens que de fato têm compra > 0 (em "Todos", a
    // lista inclui itens já cobertos com 0 un, que não devem aparecer nesse aviso).
    const semMultiplo = lista.filter(p => { const m = memoriaCompra(p); return m.comprar > 0 && !m.temMultiplo; }).length;
    const semMultTxt = semMultiplo > 0 ? ` · ${semMultiplo} item(ns) sem múltiplo cadastrado` : '';
    const caixasTxt = totalCaixas > 0 ? ` (${totalCaixas.toLocaleString('pt-BR')} caixas)` : '';
    const repostoTxt = jaRepostos > 0 ? ` · ${jaRepostos} já com reposição a caminho (fora da lista)` : '';
    const itensTxt = filter === 'todos'
        ? `<strong>${lista.length} itens</strong> no total em ${escopo}`
        : `<strong>${lista.length} itens</strong> a comprar em ${escopo}`;
    const introHtml = `<div class="rec-intro">${intro}<br>${itensTxt}${filter !== 'performance' ? ` · sugestão total: <span class="rec-total"><strong>${totalComprar.toLocaleString('pt-BR')} un</strong>${caixasTxt}</span>${repostoTxt}${semMultTxt}` : ''}</div>`;

    // Campo de meta em massa: aparece para os critérios que mostram "Comprar" (não em parado)
    const bulkHtml = filter !== 'performance' ? renderBulkMetaControl() : '';

    // Lista compacta (tabela). Sem teto baixo: todos os pontos abaixo de 60 dias com
    // necessidade de compra aparecem (o corte antigo de 300 escondia a maioria).
    recListaAtual = aplicarOrdenacaoRec(lista).slice(0, 5000);
    container.innerHTML = bulkHtml + introHtml + renderRecTable(recListaAtual, filter, filtroCD.size === 1);
}

// Campo "Meta de dias para todos os produtos" no topo da aba Sugestões. Redefine a meta
// de cobertura de TODAS as linhas de uma vez (zera ajustes individuais).
function renderBulkMetaControl() {
    return `<div class="rec-bulk-meta">
        <label for="recMetaTodos">Meta de dias para <strong>todos os produtos</strong>:</label>
        <input type="number" id="recMetaTodos" class="rec-bulk-input" min="${META_DIAS_MIN}" max="${META_DIAS_MAX}" step="1" value="${metaDiasGlobal}">
        <button type="button" class="rec-bulk-btn" id="recMetaTodosBtn">Aplicar a todos</button>
        <span class="rec-bulk-hint">Zera os ajustes feitos por linha.</span>
    </div>`;
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
// RACIONAL (igual à coluna "Sugestão Compra" do mapa de estoque oficial): a sugestão
// não desconta a foto de hoje, e sim a POSIÇÃO PROJETADA PARA O FIM DO MÊS — o estoque
// total atual menos a saída que ainda resta no mês vigente (projeção cheia − já vendido).
// O que está a caminho (trânsito + entrega) já entra porque está dentro do Estoque Total.
function memoriaCompra(p, metaOverride) {
    const diasMeta = (typeof metaOverride === 'number' && metaOverride > 0) ? metaOverride : getMetaDias(p);
    const mesesCobertura = diasMeta / 30;
    // Alvo de cobertura, em unidades: venda média ponderada × (dias meta ÷ 30). Sem
    // arredondar aqui — o ajuste ao múltiplo de caixa vem no fim, igual ao mapa.
    const necessario = p.vendaMedia * mesesCobertura;
    const aCaminho = (p.pendenciaTransito || 0) + (p.pendenciaEntrega || 0);
    // Projeção do mês vigente (un): usa a coluna PROJEÇÃO UNDS da planilha, que já traz a
    // extrapolação por dias ÚTEIS embutida pelo gerador do arquivo. Se a coluna não vier,
    // cai no projetado por dias corridos (venda parcial ÷ dia do mês × 30).
    const projMes = (p.projecaoMes > 0) ? p.projecaoMes : p.vendaParcialProjetada;
    // Saída que ainda resta até o fim do mês (projeção cheia − o já vendido no mês).
    const restoMes = projMes - p.vendaParcial;
    // Posição projetada para o fim do mês: Estoque Total de hoje (livre + qualidade +
    // bloqueado + trânsito + entrega) menos a saída restante. É a base prospectiva que a
    // sugestão desconta — não a foto atual de estoque livre.
    const estoqueProjFim = p.estoqueTotal - restoMes;
    const comprarBruto = Math.max(0, necessario - estoqueProjFim);
    // Arredonda PARA CIMA ao múltiplo da caixa de embarque (não se compra caixa fracionada)
    const mult = ajustarAoMultiplo(p, comprarBruto);
    let comprar = mult.ajustado;   // valor final usado em todo o dashboard
    let caixas = mult.caixas;
    let pisoMinimo = false;
    // PISO DE 1 CAIXA DE EMBARQUE: produto em déficit COM necessidade líquida real
    // (comprarBruto > 0) nunca sugere menos que 1 caixa de embarque.
    // SÓ se aplica quando ainda falta comprar algo. Se a posição projetada de fim de mês
    // já cobre a meta (comprarBruto = 0), a sugestão fica 0 — não faz sentido mandar
    // comprar 1 caixa de um item que o estoque projetado já cobre.
    if (mult.temMultiplo && p.vendaMedia > 0 && p.status === 'deficit'
        && comprarBruto > 0 && comprar < mult.multiplo) {
        comprar = mult.multiplo;
        caixas = 1;
        pisoMinimo = true;
    }
    // SEM GIRO RECENTE: zero venda no último mês cheio (vendas[3]) e no parcial (vendas[4]).
    // A média ponderada ainda positiva é histórico que não se repete; não justifica reabastecer
    // um item que parou de vender (ex.: ATIP no RS13). Zera a sugestão qualquer que seja o
    // tamanho da necessidade — inclusive acima de 1 caixa — para não comprar item parado. O
    // ponto vira 'parado' na classificação; aqui garante-se que a recompra também é zero.
    const semGiroRecente = p.vendaMedia > 0 && (p.vendas[3] + p.vendas[4]) <= 0;
    if (semGiroRecente) { comprar = 0; caixas = 0; pisoMinimo = false; }
    // Média simples dos meses cheios da janela (exclui o mês vigente parcial)
    const mediaMesesCheios = (p.vendas[0] + p.vendas[1] + p.vendas[2] + p.vendas[3]) / 4;
    const giroDia = p.vendaMedia / 30;
    // cobertura da posição projetada de fim de mês, em dias
    const diasFuturo = p.vendaMedia > 0 ? (estoqueProjFim / p.vendaMedia) * 30 : 0;
    // disponivelFuturo mantido como alias de estoqueProjFim (base que a compra desconta)
    return {
        necessario, comprar, comprarBruto, diasMeta, pisoMinimo, semGiroRecente,
        multiplo: mult.multiplo, caixas, temMultiplo: mult.temMultiplo,
        mediaMesesCheios, giroDia, mesesCobertura, aCaminho,
        projMes, restoMes, estoqueProjFim, disponivelFuturo: estoqueProjFim, diasFuturo
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
    const { necessario, comprar, comprarBruto, multiplo, caixas, temMultiplo, mediaMesesCheios, giroDia, mesesCobertura, aCaminho, projMes, restoMes, estoqueProjFim, diasFuturo, diasMeta, pisoMinimo, semGiroRecente } = memoriaCompra(p);
    const mesesTxt = Number.isInteger(mesesCobertura) ? `${mesesCobertura} meses` : `${mesesCobertura.toFixed(1)} meses`;
    const r = n => Math.round(n).toLocaleString('pt-BR');
    // Rótulo dos meses cheios da janela (exclui o mês vigente parcial)
    const mesesCheiosTxt = JANELA.slice(0, 4).map(m => m.cap).join('–');

    const jaCoberto = comprarBruto <= 0 && !pisoMinimo;
    const caminhoNota = aCaminho > 0 ? ` (inclui ${r(aCaminho)} un a caminho — trânsito + entrega — já dentro do total)` : '';

    const formulaBruto = jaCoberto
        ? `Necessidade líquida = ${r(necessario)} − ${r(estoqueProjFim)} = <strong>${r(necessario - estoqueProjFim)} un</strong> → <strong style="color:var(--saudavel);">0 un a comprar</strong>. A posição projetada de fim de mês já cobre a meta de ${diasMeta} dias (cobertura ≈ ${diasFuturo.toFixed(0)} dias).`
        : `Necessidade líquida = ${r(necessario)} − ${r(estoqueProjFim)} = <strong>${r(comprarBruto)} un</strong>.`;

    // Ajuste ao múltiplo da caixa de embarque
    let blocoCaixa;
    if (semGiroRecente) {
        blocoCaixa = `Sem venda no último mês cheio (${JANELA[3].cap}) nem no parcial (${JANELA[4].cap}): item <strong>parado</strong>, sem giro recente. A média ponderada (${p.vendaMedia.toLocaleString('pt-BR')} un/mês) é histórico que não se repete — <strong style="color:var(--saudavel);">0 un a comprar</strong>. Reavaliar quando voltar a vender.`;
    } else if (jaCoberto) {
        blocoCaixa = ''; // já coberto: nada a comprar, sem arredondamento de caixa
    } else if (pisoMinimo) {
        blocoCaixa = `Caixa de embarque = <strong>${multiplo.toLocaleString('pt-BR')} un</strong>. A necessidade líquida (${r(comprarBruto)} un) é menor que 1 caixa, mas o item está em déficit — aplicado o <strong>mínimo de 1 caixa</strong> = <strong style="color:var(--primary);">${comprar.toLocaleString('pt-BR')} un</strong>.`;
    } else if (temMultiplo) {
        blocoCaixa = `Caixa de embarque = <strong>${multiplo.toLocaleString('pt-BR')} un</strong>. Arredondando para cima: ${r(comprarBruto)} ÷ ${multiplo.toLocaleString('pt-BR')} → <strong>${caixas.toLocaleString('pt-BR')} ${caixas === 1 ? 'caixa' : 'caixas'}</strong> = <strong style="color:var(--primary);">${comprar.toLocaleString('pt-BR')} un</strong>.`;
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
                Venda média de <strong>${p.vendaMedia.toLocaleString('pt-BR')} un/mês</strong> (média ponderada; média simples ${mesesCheiosTxt} = ${r(mediaMesesCheios)} un/mês ≈ ${giroDia.toFixed(0)} un/dia).
                Para cobrir <strong>${diasMeta} dias (${mesesTxt})</strong>: ${p.vendaMedia.toLocaleString('pt-BR')} × ${mesesCobertura.toFixed(mesesCobertura % 1 ? 1 : 0)} = <strong>${r(necessario)} un</strong> necessárias.
                ${mesVigenteNome()} projetado em <strong>${r(projMes)} un</strong> (parcial ${p.vendaParcial.toLocaleString('pt-BR')} un até dia ${dashboardData.diasCorridos}); ainda saem <strong>${r(restoMes)} un</strong> até o fim do mês.
                Estoque total hoje <strong>${p.estoqueTotal.toLocaleString('pt-BR')} un</strong>${caminhoNota} → posição projetada de fim de mês = ${p.estoqueTotal.toLocaleString('pt-BR')} − ${r(restoMes)} = <strong>${r(estoqueProjFim)} un</strong> (cobertura ≈ ${diasFuturo.toFixed(0)} dias).
                ${formulaBruto}
                ${blocoCaixa}
                <span class="jun-nota">* Projeção do mês vigente pela coluna PROJEÇÃO UNDS da planilha (extrapolação por dias úteis).</span>
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
            ? `<span style="color:var(--excesso);font-weight:600;">${fmtUn(p.estoqueLivre)} un</span>`
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
            <td style="text-align:right;">${fmtUn(p.estoqueLivre)}</td>
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
    const TETO = RITMO_TETO_DIAS;     // 90 — acima = estoque travado (teto próprio do Ritmo, ver topo)
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

    // Camada 1: estrutura do estoque
    let estrutura;
    if (p.estoqueLivre > 0) estrutura = 'com-estoque';
    else if (aCaminho > 0) estrutura = 'em-fluxo';   // zerado, mas reposição a caminho
    else if (vm > 0) estrutura = 'ruptura';          // zerado, nada vindo, mas vende
    else estrutura = 'sem-giro';                     // sem venda: cobertura indefinida

    // Camada 3: veredito (acima / no ritmo / abaixo)
    // Risco de FALTA unificado: havendo reposição a caminho, o piso de cobertura é medido
    // sobre o pipeline (livre + a caminho), não só sobre o livre de hoje. Assim com-estoque
    // passa a enxergar a reposição igual ao em-fluxo. O EXCESSO continua olhando só o que já
    // está parado na mão (livre), porque o que vem a caminho ainda não é estoque travado.
    let veredito;
    if (estrutura === 'sem-giro') veredito = 'sem-giro';
    else if (estrutura === 'ruptura') veredito = 'acima';
    else if (estrutura === 'em-fluxo') veredito = cobCaminho < PISO ? 'acima' : 'ritmo';
    else if (cobAtual > TETO) veredito = 'abaixo';
    else {
        const cobEfetiva = aCaminho > 0 ? cobCaminho : cobAtual;
        veredito = cobEfetiva < PISO ? 'acima' : 'ritmo';
    }

    // "Em reposição": estoque na mão magro hoje (abaixo do piso), mas o que vem a caminho
    // recompõe a cobertura até o piso. Não é caso de compra, é de monitorar a entrega.
    // Equivale à sugestão = 0 (livre + a caminho já cobre RITMO_META_DIAS, que é igual a PISO).
    const emReposicao = aCaminho > 0 && veredito === 'ritmo' && cobAtual < PISO;

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
        estrutura, tendencia, varTend, veredito, emReposicao, sBase, sRec, aCaminho,
        cobAtual, cobCaminho, diasSemEntrada, prioridade
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
    const foco = (ritmoFilter && ritmoFilter.value) || 'tudo';
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
    const reposicao = todos.filter(o => o.r.emReposicao);
    const ritmo = todos.filter(o => o.r.veredito === 'ritmo' && !o.r.emReposicao);
    const abaixo = todos.filter(o => o.r.veredito === 'abaixo');
    const ruptura = acima.filter(o => o.r.estrutura === 'ruptura').length;
    const rsParado = abaixo.reduce((s, o) => s + (o.p.estoqueLivreRS || 0), 0);
    if (kpisEl) kpisEl.innerHTML = renderRitmoKpis({ acima: acima.length, ruptura, reposicao: reposicao.length, ritmo: ritmo.length, abaixo: abaixo.length, rsParado }, sc);

    let lista;
    if (foco === 'acima') lista = acima.slice();
    else if (foco === 'reposicao') lista = reposicao.slice();
    else if (foco === 'abaixo') lista = abaixo.slice();
    else if (foco === 'ritmo') lista = ritmo.slice();
    else lista = todos.slice();

    // Ordenação: 'auto' = recomendada por foco; senão a escolha explícita do usuário.
    const ordem = (ritmoSort && ritmoSort.value) || 'auto';
    if (ordem === 'compra') lista.sort((a, b) => (b.compra - a.compra) || (a.r.cobAtual - b.r.cobAtual));
    else if (ordem === 'nome') lista.sort((a, b) => rotuloProduto(a.p).localeCompare(rotuloProduto(b.p), 'pt-BR'));
    else if (ordem === 'reprimido') lista.sort((a, b) => (b.r.prioridade - a.r.prioridade) || (a.r.cobAtual - b.r.cobAtual));
    else if (foco === 'abaixo') lista.sort((a, b) => b.r.cobAtual - a.r.cobAtual);
    else if (foco === 'ritmo' || foco === 'reposicao') lista.sort((a, b) => a.r.cobAtual - b.r.cobAtual);
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
        acima: `<strong>Demanda reprimida</strong>: a venda supera o estoque livre, cobertura abaixo de ${DIAS_MIN_SAUDAVEL} dias, e o que vem a caminho (se vem) não cobre. Risco de ruptura, prioridade de compra. A sugestão repõe até ${RITMO_META_DIAS} dias, já descontando o que vem. Os mais reprimidos primeiro.`,
        reposicao: `<strong>Em reposição</strong>: estoque na mão magro hoje (abaixo de ${DIAS_MIN_SAUDAVEL} dias), mas a reposição a caminho já recompõe a meta. Não é compra, é monitorar a entrega: se atrasar, vira ruptura. Os mais magros hoje primeiro.`,
        abaixo: `<strong>Estoque sobrando</strong>: cobertura acima de ${RITMO_TETO_DIAS} dias. Capital parado, avaliar frear reposição. Maior cobertura primeiro.`,
        ritmo: `<strong>No ritmo</strong>: cobertura entre ${DIAS_MIN_SAUDAVEL} e ${RITMO_TETO_DIAS} dias com o estoque na mão. Equilibrado, sem ação.`,
        tudo: 'Todos os produtos com giro, do mais reprimido ao mais parado.'
    };
    const intro = `<div class="rec-intro">${introMap[foco] || introMap.acima}<br><strong>${lista.length} ${noun}</strong> em ${escopo}.</div>`;
    container.innerHTML = intro + renderRitmoTable(lista, sc || filtroCD.size === 1, sc);
}

function renderRitmoKpis(k, sc) {
    const totalReal = k.acima + (k.reposicao || 0) + k.ritmo + k.abaixo;
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
            ${seg(k.reposicao || 0, 'var(--reposicao)', 'Em reposição')}
            ${seg(k.ritmo, 'var(--saudavel)', 'No ritmo')}
            ${seg(k.abaixo, 'var(--excesso)', 'Estoque sobrando')}
        </div>
        <div class="ritmo-portfolio-leg">
            ${leg('var(--deficit)', 'Demanda reprimida', k.acima)}
            ${leg('var(--reposicao)', 'Em reposição', k.reposicao || 0)}
            ${leg('var(--saudavel)', 'No ritmo', k.ritmo)}
            ${leg('var(--excesso)', 'Estoque sobrando', k.abaixo)}
        </div>
        ${cap.length ? `<div class="ritmo-portfolio-cap">${cap.join(' · ')}</div>` : ''}
    </div>`;
}

function ritmoZona(d) { return d < DIAS_MIN_SAUDAVEL ? 'acima' : d <= RITMO_TETO_DIAS ? 'ritmo' : 'abaixo'; }
const RITMO_COR = { acima: 'var(--deficit)', ritmo: 'var(--saudavel)', abaixo: 'var(--excesso)' };
const RITMO_META_DIAS = 45;   // meta da sugestão de compra NESTA visão (alinhada à linha de alerta de 45 d e à meta padrão da aba de Sugestões).

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

// Linha de leitura direta: bate o olho e lê o veredito + a cobertura em dias.
function renderRitmoTable(lista, hideCD, sc) {
    const rows = lista.map((o, idx) => {
        const p = o.p, r = o.r;

        const cdTxt = sc
            ? `Santa Cruz · ${(p.nCDs || 1)} CD${(p.nCDs || 1) > 1 ? 's' : ''}`
            : (hideCD ? '' : escapeHtml(p.cd));
        const meta = [cdTxt, `curva ${p.curva}`].filter(Boolean).join(' · ');

        // VENDA/MÊS (coluna do ritmo de saída): a venda média mensal, sóbria. Substitui as
        // mini-barras Entrada × Saída — agora a velocidade vem em número, e a cobertura
        // (coluna ao lado) é esse estoque traduzido em dias.
        const vm = Math.max(p.vendaMedia || 0, 0);
        const vendaHtml = `<div class="ritmo-rvenda"><b>${Math.round(vm).toLocaleString('pt-BR')}</b><i>/mês</i></div>`;

        // DIAGNÓSTICO (bloco do meio): a cobertura em dias (sóbria) + a situação numa etiqueta.
        // A cor agora vive na etiqueta (pill) da Situação, não no número da cobertura.
        // "hoje" = só o estoque livre; "previsto" = livre + a caminho.
        const semGiro = r.veredito === 'sem-giro';
        const cobHoje = Math.max(0, Math.round(r.cobAtual));
        const cobFut = Math.max(0, Math.round(r.cobCaminho));
        const temCaminho = r.aCaminho > 0;
        // Status + cor (borda da linha) + tom da etiqueta (pill) por estado. "Em reposição"
        // (cyan) = magro hoje mas o que vem cobre a meta. "Reposição insuficiente" = vem algo
        // mas não cobre (segue vermelho, é compra).
        let statusWord, cor, tone;
        if (r.estrutura === 'ruptura') { statusWord = 'Ruptura'; cor = 'var(--deficit)'; tone = 'deficit'; }
        else if (r.emReposicao) { statusWord = 'Em reposição'; cor = 'var(--reposicao)'; tone = 'reposicao'; }
        else if (r.veredito === 'acima') { statusWord = temCaminho ? 'Reposição insuficiente' : 'Demanda reprimida'; cor = 'var(--deficit)'; tone = 'deficit'; }
        else if (r.veredito === 'abaixo') { statusWord = 'Estoque sobrando'; cor = 'var(--excesso)'; tone = 'excesso'; }
        else if (semGiro) { statusWord = 'Sem giro'; cor = 'var(--text-secondary)'; tone = 'neutro'; }
        else { statusWord = 'No ritmo'; cor = 'var(--saudavel)'; tone = 'saudavel'; }
        // Cobertura (coluna do meio): "hoje" = só o estoque livre; "previsto" = livre + a caminho.
        // Sóbria de propósito: a cor da situação fica na etiqueta, não aqui.
        const cobTip = semGiro
            ? 'Sem venda no período: cobertura indefinida.'
            : `Cobertura com o estoque livre de hoje: ${cobHoje} dias.${temCaminho ? ` Prevista com o que vem a caminho (${r.aCaminho.toLocaleString('pt-BR')} un em trânsito + entrega): ${cobFut} dias.` : ''}`;
        const cobSub = cobFut > cobHoje ? `${cobFut.toLocaleString('pt-BR')} d previsto` : 'sem reposição';
        // Quanto de reposição vem a caminho (trânsito + entrega), logo abaixo do "previsto".
        // Só aparece quando há reposição que de fato melhora a cobertura (evita "+0 un").
        const repoQtd = (cobFut > cobHoje && r.aCaminho > 0)
            ? `<div class="ritmo-rcob-q">+${r.aCaminho.toLocaleString('pt-BR')} un</div>`
            : '';
        const cobHtml = semGiro
            ? `<div class="ritmo-rcob" title="${cobTip}"><div class="ritmo-rcob-h"><span class="na">sem venda</span></div></div>`
            : `<div class="ritmo-rcob" title="${cobTip}">
                    <div class="ritmo-rcob-h"><b>${cobHoje.toLocaleString('pt-BR')}</b> d <span class="lbl">hoje</span></div>
                    <div class="ritmo-rcob-f">${cobSub}</div>
                    ${repoQtd}
                </div>`;

        // Compra (coluna da direita): a sugestão, sempre em azul. Mesma conta da aba Sugestões
        // (meta de 45 dias, descontando a posição projetada de fim de mês). Zero fica apagado.
        const mc = memoriaCompra(p, RITMO_META_DIAS);
        const compra = mc.comprar || 0;
        const projFimTxt = Math.round(mc.estoqueProjFim).toLocaleString('pt-BR');
        const compraTip = compra > 0
            ? `Comprar para ${mc.diasMeta} dias, descontada a posição projetada de fim de mês (${projFimTxt} un = estoque total − saída restante).`
            : `Posição projetada de fim de mês (${projFimTxt} un) já cobre a meta de ${mc.diasMeta} dias.`;
        const buyHtml = compra > 0
            ? `<div class="ritmo-rbuy" title="${compraTip}"><b>${compra.toLocaleString('pt-BR')}</b><i>un</i></div>`
            : `<div class="ritmo-rbuy zero" title="${compraTip}"><b>0</b></div>`;

        return `
        <div class="ritmo-item">
            <div class="ritmo-rrow" onclick="toggleRitmoDetail(${idx})" style="border-left-color:${cor};">
                <div class="ritmo-rname">
                    <div class="ritmo-rn" title="${escapeHtml(rotuloProduto(p))}">${truncate(rotuloProduto(p), 44)}</div>
                    <div class="ritmo-rm">${meta}</div>
                </div>
                ${vendaHtml}
                ${cobHtml}
                <div class="ritmo-rsit"><span class="ritmo-pill ${tone}"><span class="ritmo-pill-dot"></span>${statusWord}</span></div>
                ${buyHtml}
                <span class="ritmo-rexp" id="ritexp-${idx}">▸</span>
            </div>
            <div class="ritmo-detalhe" id="ritdet-${idx}" style="display:none;">${explicacaoRitmo(p, r)}</div>
        </div>`;
    }).join('');

    const base = sc
        ? 'Somado entre todos os CDs da Santa Cruz. Clique numa linha para a memória de cálculo.'
        : 'Cada linha é um produto num CD. Clique para a memória de cálculo.';
    const dica = `${base} Venda/mês é o ritmo de saída; cobertura é esse estoque traduzido em dias.`;
    const head = `<div class="ritmo-rhead">
            <span class="rh-name">Produto</span>
            <span class="rh-venda">Venda/mês</span>
            <span class="rh-cob">Cobertura</span>
            <span class="rh-sit">Situação</span>
            <span class="rh-buy">Comprar</span>
            <span class="rh-exp"></span>
        </div>`;
    return `<div class="ritmo-dica">${dica}</div><div class="ritmo-list">${head}${rows}</div>`;
}

// Memória do ritmo — compacta: os números do produto, o gap de cobertura lado a lado,
// o veredito em 1 linha, a conta de compra inteira e onde o SKU vende.
function explicacaoRitmo(p, r) {
    const fmt = n => isFinite(n) ? Math.round(n).toLocaleString('pt-BR') : '∞';
    const cor = r.emReposicao ? 'var(--reposicao)' : (RITMO_COR[r.veredito] || 'var(--text-secondary)');
    const dir = (r.tendencia === 'baixo' || r.varTend == null) ? null
        : r.tendencia === 'acelerando' ? 'Acelerando'
        : r.tendencia === 'desacelerando' ? 'Desacelerando'
        : 'Estável';

    // 0) Números do produto (saíram da linha compacta; aqui ficam por extenso).
    const vmF = Math.round(p.vendaMedia || 0).toLocaleString('pt-BR');
    const livreF = Math.round(p.estoqueLivre || 0).toLocaleString('pt-BR');
    const camNum = (p.pendenciaTransito || 0) + (p.pendenciaEntrega || 0);
    const numsHtml = `<div class="rit-line"><span class="rit-line-rot">Fluxo</span><span class="rit-line-v">Saída <b>${vmF}</b>/mês · livre na mão <b>${livreF}</b>${camNum > 0 ? ` · a caminho <b>${camNum.toLocaleString('pt-BR')}</b>` : ' · nada a caminho'}</span></div>`;

    // 1) Coberturas lado a lado: o gap entre elas é o sinal do ritmo.
    const cob = [];
    cob.push(`<div class="rit-cob-cell"><div class="rit-cob-n" style="color:${cor};">${fmt(r.cobAtual)}<span class="rit-cob-u">d</span></div><div class="rit-cob-l">livre hoje</div></div>`);
    if (r.aCaminho > 0)
        cob.push(`<div class="rit-cob-cell"><div class="rit-cob-n">${fmt(r.cobCaminho)}<span class="rit-cob-u">d</span></div><div class="rit-cob-l">previsto</div></div>`);
    if (r.diasSemEntrada != null)
        cob.push(`<div class="rit-cob-cell rit-cob-ctx"><div class="rit-cob-n">${r.diasSemEntrada}<span class="rit-cob-u">d</span></div><div class="rit-cob-l">desde a entrada</div></div>`);
    const cobHtml = `<div class="rit-cob">${cob.join('')}</div>`;

    // 2) Veredito em uma linha: palavra-chave colorida + leitura curta.
    let kw, txt;
    if (r.estrutura === 'ruptura') {
        kw = 'Ruptura'; txt = `livre zerado e nada a caminho; vende ${fmt(p.vendaMedia)}/mês.`;
    } else if (r.emReposicao) {
        kw = 'Em reposição'; txt = `livre cobre só ${fmt(r.cobAtual)} d hoje, mas o que vem a caminho recompõe ${fmt(r.cobCaminho)} d (acima de ${DIAS_MIN_SAUDAVEL}). Monitorar a entrega.`;
    } else if (r.veredito === 'acima' && r.aCaminho > 0) {
        kw = 'Reposição insuficiente'; txt = `livre cobre ${fmt(r.cobAtual)} d e o que vem leva só a ${fmt(r.cobCaminho)} d (abaixo de ${DIAS_MIN_SAUDAVEL}). A reposição não acompanha.`;
    } else if (r.veredito === 'acima') {
        kw = dir || 'Reprimida'; txt = `a saída supera o estoque livre; cobertura abaixo de ${DIAS_MIN_SAUDAVEL} d e nada a caminho.`;
    } else if (r.veredito === 'abaixo') {
        kw = dir || 'Sobrando'; txt = `estoque sobrando para o giro; cobertura acima de ${RITMO_TETO_DIAS} d.`;
    } else if (r.veredito === 'sem-giro') {
        kw = 'Sem giro'; txt = `sem venda na janela; cobertura indefinida.`;
    } else {
        kw = dir || 'No ritmo'; txt = `entrada e saída equilibradas; cobertura na faixa ${DIAS_MIN_SAUDAVEL}–${RITMO_TETO_DIAS} d.`;
    }
    const verdHtml = `<div class="rit-verd" style="border-left-color:${cor};"><b style="color:${cor};">${kw}</b> · ${txt}</div>`;

    // 3) Compra: a conta inteira em uma linha (descontando a projeção de fim de mês).
    const mc = memoriaCompra(p, RITMO_META_DIAS);
    const base = `meta ${mc.diasMeta} d → ${fmt(mc.necessario)} un · projeção fim do mês ${fmt(mc.estoqueProjFim)} (total ${fmt(p.estoqueTotal)} − ${fmt(mc.restoMes)} resto)`;
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
        ${numsHtml}
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

// Retorna os SKUs únicos disponíveis na aba Detalhes para um termo de busca,
// já filtrados (filtro GLOBAL + CDs do Detalhes) e ordenados. É a MESMA base
// usada para montar a lista E para o "Selecionar todos", então os dois andam
// sempre em sincronia (o botão seleciona exatamente o que o filtro mostra).
function skusDisponiveisDetalhe(filtro) {
    if (!dashboardData) return [];
    const termo = (filtro || '').toLowerCase().trim();
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
    return skus;
}

// Monta a lista de checkboxes de Produto (SKUs únicos), filtrada pela busca
// e limitada aos CDs selecionados (se houver).
function buildDetailProductOptions(filtro) {
    const opts = document.getElementById('msProdOptions');
    if (!dashboardData) { opts.innerHTML = ''; return; }
    const skus = skusDisponiveisDetalhe(filtro);

    const limite = 300;
    const lista = skus.slice(0, limite);
    opts.innerHTML = lista.map(p => {
        const lbl = rotuloProduto(p);
        return `<label class="ms-opt"><input type="checkbox" value="${escapeHtml(p.sku)}" data-msprod${detailProducts.has(p.sku) ? ' checked' : ''}> <span title="${escapeHtml(lbl)}">${escapeHtml(lbl)}</span></label>`;
    }).join('') || '<div class="ms-empty">Nenhum produto encontrado</div>';

    if (skus.length > limite) {
        opts.innerHTML += `<div class="ms-empty">Mostrando ${limite} de ${skus.length}. Refine a busca.</div>`;
    }

    // Reflete no botão quantos produtos o "Selecionar todos" vai marcar agora
    // (respeita a busca e os CDs do Detalhes; inclui os que passam do limite acima).
    const btnAll = document.getElementById('msProdSelectAll');
    if (btnAll) btnAll.textContent = skus.length ? `Selecionar todos (${skus.length})` : 'Selecionar todos';

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
            <td style="text-align:right;">${fmtUn(p.estoqueLivre)}</td>
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
                    <li><span class="label">Estoque Livre:</span> <span class="value">${fmtUn(p.estoqueLivre)} un</span></li>
                    <li><span class="label">Estoque Qualidade:</span> <span class="value">${fmtUn(p.estoqueQualidade)} un</span></li>
                    <li><span class="label">Estoque Bloqueado:</span> <span class="value">${fmtUn(p.estoqueBloqueado)} un</span></li>
                    <li><span class="label">Estoque Total:</span> <span class="value">${fmtUn(p.estoqueTotal)} un</span></li>
                    <li><span class="label">Pendência Trânsito:</span> <span class="value">${fmtUn(p.pendenciaTransito)} un</span></li>
                    <li><span class="label">Pendência Entrega:</span> <span class="value">${fmtUn(p.pendenciaEntrega)} un</span></li>
                    <li><span class="label">Estoque Livre R$:</span> <span class="value">${formatBRL(p.estoqueLivreRS)}</span></li>
                </ul>
            </div>
            <div class="detail-section">
                <h4>⚡ Análise de Giro</h4>
                <ul>
                    <li><span class="label">Venda Média Ponderada:</span> <span class="value">${p.vendaMedia.toLocaleString('pt-BR')} un/mês</span></li>
                    <li><span class="label">Equivale a:</span> <span class="value">${fmtUn(memo.giroDia)} un/dia</span></li>
                    <li><span class="label">Dias de Estoque (livre):</span> <span class="value">${fmtUn(p.diasLivre, 1)}</span></li>
                    <li><span class="label">Dias de Estoque (total):</span> <span class="value">${fmtUn(p.diasTotal, 1)}</span></li>
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

    // "Selecionar todos" / "Limpar" da lista de produtos (aba Detalhes).
    const prodSelectAll = document.getElementById('msProdSelectAll');
    const prodClearAll = document.getElementById('msProdClearAll');

    if (prodSelectAll) prodSelectAll.addEventListener('click', (e) => {
        e.stopPropagation();
        const skus = skusDisponiveisDetalhe(prodSearch.value);
        if (!skus.length) return;
        // Cada produto vira um (ou mais) bloco de detalhe — um nº muito grande pode
        // deixar a aba pesada. Acima do limite, confirma antes (mas deixa seguir).
        if (skus.length > MAX_DETALHE_CONFIRMA &&
            !confirm(`Selecionar ${skus.length} produtos? Carregar todos os detalhes de uma vez pode deixar a aba lenta.`)) {
            return;
        }
        skus.forEach(p => detailProducts.add(p.sku));
        buildDetailProductOptions(prodSearch.value);
        renderDetailChips();
        updateDetailLabels();
        renderDetailBlocks();
    });

    if (prodClearAll) prodClearAll.addEventListener('click', (e) => {
        e.stopPropagation();
        if (!detailProducts.size) return;
        detailProducts.clear();
        document.querySelectorAll('#msProdOptions input[data-msprod]:checked').forEach(c => { c.checked = false; });
        renderDetailChips();
        updateDetailLabels();
        renderDetailBlocks();
    });

    // Fecha dropdowns ao clicar fora
    document.addEventListener('click', () => {
        cdDropdown.classList.remove('open'); cdToggle.classList.remove('active');
        prodDropdown.classList.remove('open'); prodToggle.classList.remove('active');
    });
}

// Mini gráfico de barras para o painel de detalhes
// ============================================
// TOP 5 MARCAS POR MÊS (observação de cada barra)
// ============================================
// Para cada um dos 5 meses da janela, soma as vendas por marca (= Material),
// ordena e devolve as 5 maiores, junto com o % que cada uma representou no mês.
// Lê os mesmos vetores (p.vendas / p.vendasRS) que desenham as barras, então o
// total das marcas sempre fecha com a barra do mês.
let _chartTipReg = {};   // id do gráfico -> [htmlMes0, htmlMes1, ... htmlMes4]
let _chartSeq = 0;       // gera ids únicos a cada gráfico desenhado

function topMarcasMes(produtos, chave) {
    const meses = [];
    for (let i = 0; i < 5; i++) {
        const acc = {};
        let total = 0;
        produtos.forEach(p => {
            const arr = p[chave];
            const v = arr ? (arr[i] || 0) : 0;
            total += v;                                  // total líquido = igual à barra do mês
            acc[p.material] = (acc[p.material] || 0) + v; // venda líquida por marca (soma os CDs)
        });
        const top = Object.keys(acc)
            .map(nome => ({ nome, valor: acc[nome] }))
            .filter(o => o.valor > 0)                    // só marcas com venda líquida positiva
            .sort((a, b) => b.valor - a.valor)
            .slice(0, 5)
            .map(o => ({ nome: o.nome, valor: o.valor, share: total > 0 ? o.valor / total : 0 }));
        meses.push({ total, top });
    }
    return meses;
}

// HTML do tooltip de um mês: a lista ranqueada das 5 marcas
function tipMarcasHtml(mesInfo, label, parcial, isMoney) {
    const fmt = v => isMoney
        ? 'R$ ' + Math.round(v).toLocaleString('pt-BR')
        : Math.round(v).toLocaleString('pt-BR') + ' un';
    const tag = parcial ? ' <span class="tip-parcial">parcial</span>' : '';
    if (!mesInfo.top.length) {
        return `<div class="tip-head"><span>Top 5 marcas — ${label}${tag}</span></div>` +
               `<div class="tip-empty">Sem vendas registradas neste mês.</div>`;
    }
    const cor = isMoney ? 'var(--saudavel)' : 'var(--primary)';
    const rows = mesInfo.top.map((t, k) => `
        <div class="tip-row">
            <span class="tip-rank">${k + 1}</span>
            <span class="tip-name" title="${escapeHtml(t.nome)}">
                <span class="tip-name-txt">${escapeHtml(t.nome)}</span>
                <span class="tip-bar"><i style="width:${(t.share * 100).toFixed(1)}%;background:${cor}"></i></span>
            </span>
            <span class="tip-val">${fmt(t.valor)}<small>${(t.share * 100).toFixed(1)}%</small></span>
        </div>`).join('');
    return `<div class="tip-head"><span>Top 5 marcas — ${label}${tag}</span>` +
           `<span class="tip-total">${fmt(mesInfo.total)}</span></div>` +
           `<div class="tip-list">${rows}</div>`;
}

// --- Tooltip flutuante único, compartilhado por todos os gráficos ---
let _tipPinned = false;
function _ensureTipEl() {
    let t = document.getElementById('chartTip');
    if (!t) {
        t = document.createElement('div');
        t.id = 'chartTip';
        t.className = 'chart-tip';
        document.body.appendChild(t);
    }
    return t;
}
function _posTip(t, x, y) {
    const pad = 12;
    const r = t.getBoundingClientRect();
    let left = x - r.width / 2;
    let top = y - r.height - 14;            // acima do ponteiro
    if (top < pad) top = y + 20;            // se não couber acima, abre abaixo
    left = Math.max(pad, Math.min(left, window.innerWidth - r.width - pad));
    top = Math.max(pad, Math.min(top, window.innerHeight - r.height - pad));
    t.style.left = left + 'px';
    t.style.top = top + 'px';
}
function _showTip(cid, idx, x, y) {
    const tips = _chartTipReg[cid];
    if (!tips || tips[idx] == null) return;
    const t = _ensureTipEl();
    t.innerHTML = tips[idx];
    t.classList.add('show');
    _posTip(t, x, y);
}
function _hideTip() {
    const t = document.getElementById('chartTip');
    if (t) t.classList.remove('show');
    _tipPinned = false;
}
function _hitDe(e) {
    const el = e.target;
    return el && el.closest ? el.closest('.bar-hit') : null;
}
let _chartTipsReady = false;
function initChartTips() {
    if (_chartTipsReady) return;
    _chartTipsReady = true;
    document.addEventListener('mouseover', e => {
        if (_tipPinned) return;
        const hit = _hitDe(e);
        if (hit) _showTip(hit.getAttribute('data-chart'), +hit.getAttribute('data-idx'), e.clientX, e.clientY);
    });
    document.addEventListener('mousemove', e => {
        if (_tipPinned) return;
        const hit = _hitDe(e);
        const t = document.getElementById('chartTip');
        if (hit && t && t.classList.contains('show')) _posTip(t, e.clientX, e.clientY);
    });
    document.addEventListener('mouseout', e => {
        if (_tipPinned) return;
        const hit = _hitDe(e);
        if (!hit) return;
        const to = e.relatedTarget;
        if (!to || !(to.closest && to.closest('.bar-hit'))) _hideTip();
    });
    // toque / clique: fixa o mês tocado (essencial no mobile); tocar fora fecha
    document.addEventListener('click', e => {
        const hit = _hitDe(e);
        if (hit) {
            _tipPinned = false;
            const r = hit.getBoundingClientRect();
            _showTip(hit.getAttribute('data-chart'), +hit.getAttribute('data-idx'), r.left + r.width / 2, r.top + 10);
            _tipPinned = true;
        } else if (_tipPinned) {
            _hideTip();
        }
    });
    window.addEventListener('scroll', _hideTip, true);
    window.addEventListener('resize', _hideTip);
}
initChartTips();   // ativa os tooltips de "Top 5 marcas" nas barras (uma vez, no load)

function createBarChartSimple(labels, values, color, isMoney, tips) {
    const maxV = Math.max(...values, 1);
    const w = 320, h = 180, padB = 30, padT = 24, padL = 10;
    const slot = (w - padL * 2) / labels.length;
    const plotH = h - padB - padT;
    const interativo = Array.isArray(tips) && tips.length === labels.length;
    const cid = interativo ? 'c' + (++_chartSeq) : '';
    if (interativo) _chartTipReg[cid] = tips;

    let bars = '', hits = '';
    values.forEach((v, i) => {
        const bh = (v / maxV) * plotH;
        const x = padL + i * slot + slot * 0.2;
        const y = padT + plotH - bh;
        const bw = slot * 0.6;
        bars += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bw.toFixed(1)}" height="${bh.toFixed(1)}" fill="${color}" opacity="0.85" rx="3"/>`;
        bars += `<text x="${(x + bw / 2).toFixed(1)}" y="${h - padB + 16}" text-anchor="middle" font-size="11" fill="#64748b">${labels[i]}</text>`;
        const lbl = isMoney ? formatBRLShort(v) : Math.round(v).toLocaleString('pt-BR');
        bars += `<text x="${(x + bw / 2).toFixed(1)}" y="${(y - 5).toFixed(1)}" text-anchor="middle" font-size="9.5" font-weight="bold" fill="#1e293b">${lbl}</text>`;
        if (interativo) {
            // área sensível cobrindo a coluna inteira (transparente, por cima das barras)
            const hx = padL + i * slot;
            hits += `<rect class="bar-hit" data-chart="${cid}" data-idx="${i}" x="${hx.toFixed(1)}" y="${padT}" width="${slot.toFixed(1)}" height="${plotH.toFixed(1)}" fill="transparent"/>`;
        }
    });

    const svg = `<svg width="100%" height="${h}" viewBox="0 0 ${w} ${h}" style="max-width:100%;">
        <line x1="${padL}" y1="${padT + plotH}" x2="${w - padL}" y2="${padT + plotH}" stroke="#e2e8f0" stroke-width="1"/>
        ${bars}
        ${hits}
    </svg>`;
    return interativo ? `<div class="chart-wrap">${svg}</div>` : svg;
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
        if (!porCD[cd]) porCD[cd] = { un: [0, 0, 0, 0, 0], rs: [0, 0, 0, 0, 0], prods: [] };
        porCD[cd].prods.push(p);
        for (let i = 0; i < 5; i++) { porCD[cd].un[i] += p.vendas[i]; porCD[cd].rs[i] += p.vendasRS[i]; }
    });
    cdSalesPorCD = porCD;
    _chartTipReg = {};   // zera os tooltips a cada atualização para não acumular
    const cds = Object.keys(porCD);
    const tituloTotal = cds.length === 1 ? cds[0] : 'Total da Distribuidora';
    totalBox.innerHTML = blocoVendasCD(tituloTotal, totalUn, totalRS, true, produtos);

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
    detBox.innerHTML = blocoVendasCD(cd, cdSalesPorCD[cd].un, cdSalesPorCD[cd].rs, false, cdSalesPorCD[cd].prods);
}

function blocoVendasCD(titulo, un, rs, destaque, prods) {
    const labels = JANELA.map(m => m.abbr + (m.parcial ? '*' : ''));
    const dias = dashboardData.diasCorridos || 0;
    const parcial = un[4];
    const proj = dias > 0 ? Math.round(parcial / dias * 30) : Math.round(parcial);
    const media4 = Math.round((un[0] + un[1] + un[2] + un[3]) / 4);
    const mediaRS4 = (rs[0] + rs[1] + rs[2] + rs[3]) / 4;

    // Observação de cada mês: as 5 marcas (produtos) que mais venderam — em
    // unidades no gráfico de UN e em R$ no gráfico de R$. Cada mês tem seu ranking.
    const lista = prods || [];
    const topUn = topMarcasMes(lista, 'vendas');
    const topRS = topMarcasMes(lista, 'vendasRS');
    const tipsUn = JANELA.map((m, i) => tipMarcasHtml(topUn[i], m.cap, m.parcial, false));
    const tipsRS = JANELA.map((m, i) => tipMarcasHtml(topRS[i], m.cap, m.parcial, true));
    const dica = '<p class="chart-hint">🏆 Passe o mouse ou toque em um mês para ver as 5 marcas que mais venderam.</p>';

    const chartUn = createBarChartSimple(labels, un.map(v => Math.round(v)), '#3b82f6', false, tipsUn) +
        `<p class="chart-note">* ${mesVigenteNome()} parcial (até dia ${dias}). Projeção mês cheio: <strong>${proj.toLocaleString('pt-BR')} un</strong></p>` + dica;
    const chartRS = createBarChartSimple(labels, rs, '#10b981', true, tipsRS) + dica;
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

// ============================================
// COCKPIT · O QUE MUDOU
// Tese única: "recompra represada" = quanto de faturamento Supera está travado agora,
// medido pela MESMA conta da aba Sugestões (memoriaCompra → un a comprar × PF).
// ============================================

// Recompra represada de um produto (un a comprar × preço de fábrica).
// Usa memoriaCompra (projeção de fim de mês + múltiplo de caixa + meta global) para
// bater 1:1 com a aba Sugestões de Compra — não inventa um segundo racional.
function recompraDe(p) {
    const comprar = memoriaCompra(p).comprar;
    return { comprar, un: comprar, rs: comprar * (p.pf || 0) };
}

// "DD/MM/AAAA" -> "AAAAMMDD" (string ordenável por data-base)
function chaveDataRef(ref) {
    const m = String(ref || '').match(/(\d{2})\/(\d{2})\/(\d{4})/);
    return m ? (m[3] + m[2] + m[1]) : '00000000';
}

// Há algum filtro global ativo (CD, Curva ou Status)? O Cockpit mostra a comparação
// temporal só na visão cheia (sem filtro), pois o retrato salvo é da base inteira.
function temFiltroGlobal() {
    return filtroCD.size > 0 || filtroCurva.size > 0 || filtroStatus.size > 0;
}

// Produtos do Cockpit: base por CD+SKU dentro do filtro global (CD/Curva/Status).
function produtosCockpit() {
    return originalProducts.filter(p => dentroDoFiltroGlobal(p));
}

// Agrega uma lista de produtos nos números-chave do painel (recompra, capital por faixa,
// contagem por status). Base única dos três painéis e do snapshot.
function agregarNumeros(prods) {
    let recompraRS = 0, recompraUn = 0, itensCompra = 0;
    let capTotal = 0, capAtencao = 0, capProblema = 0, capCompra = 0;
    // Itens que ainda precisam de PEDIDO NOVO para bater a meta de 45 dias. Mesma conta da
    // aba Sugestões (filtro "deficit"): população < 60d (deficit/saudavel/reposicao) com
    // comprar > 0. compraDef = já em déficit (< 45d); compra4560 = ainda 45-60d mas a
    // projeção de fim de mês cai abaixo de 45d. reposicao tem comprar 0 (já coberto).
    let compraNec = 0, compraDef = 0, compra4560 = 0;
    const cont = { deficit: 0, saudavel: 0, reposicao: 0, excesso: 0, 'sem-giro': 0 };
    prods.forEach(p => {
        capTotal += p.estoqueLivreRS;
        if (p.status === 'excesso') capAtencao += p.estoqueLivreRS;
        else if (p.status === 'sem-giro') capProblema += p.estoqueLivreRS;
        if (p.status === 'deficit' || p.status === 'saudavel' || p.status === 'reposicao') capCompra += p.estoqueLivreRS;
        cont[p.status] = (cont[p.status] || 0) + 1;
        const r = recompraDe(p);
        if (r.un > 0) {
            recompraRS += r.rs; recompraUn += r.un; itensCompra++;
            if (p.status === 'deficit' || p.status === 'saudavel' || p.status === 'reposicao') {
                compraNec++;
                if (p.status === 'deficit') compraDef++;
                else if (p.status === 'saudavel') compra4560++;
            }
        }
    });
    return { recompraRS, recompraUn, itensCompra, capTotal, capAtencao, capProblema, capCompra, compraNec, compraDef, compra4560, cont, n: prods.length };
}

// Recompra represada por faixa de status (quem dispara o pedido).
function recompraPorFaixa(prods) {
    const f = { deficit: 0, saudavel: 0, reposicao: 0, excesso: 0, 'sem-giro': 0 };
    prods.forEach(p => { const r = recompraDe(p); if (r.un > 0) f[p.status] = (f[p.status] || 0) + r.rs; });
    return f;
}

// Total a caminho (pendência de trânsito + entrega) em unidades e em R$ (× PF), separado
// por tipo. Já está dentro do Estoque Total, então a recompra represada é líquida disto.
function pendenciaTotal(prods) {
    let transito = 0, entrega = 0, rs = 0;
    prods.forEach(p => {
        const t = p.pendenciaTransito || 0, e = p.pendenciaEntrega || 0;
        transito += t; entrega += e;
        rs += (t + e) * (p.pf || 0);
    });
    return { transito, entrega, un: transito + entrega, rs };
}

// Recompra represada quebrada por curva (A/B/C). Cada curva traz R$, unidades, pontos
// e a parcela que é déficit crítico (< 45 dias).
function recompraPorCurva(prods) {
    const base = () => ({ rs: 0, un: 0, itens: 0, deficit: 0 });
    const c = { A: base(), B: base(), C: base() };
    prods.forEach(p => {
        const k = (p.curva === 'A' || p.curva === 'B' || p.curva === 'C') ? p.curva : null;
        if (!k) return;
        const r = recompraDe(p);
        if (r.un > 0) {
            c[k].rs += r.rs; c[k].un += r.un; c[k].itens++;
            if (p.status === 'deficit') c[k].deficit += r.rs;
        }
    });
    return c;
}

// Produtos de uma curva CONSOLIDADOS por marca (uma linha por SKU, somando TODOS os CDs).
// A sugestão de compra de cada linha é a soma das necessidades de cada CD daquela marca
// (un e R$), pois a reposição é por CD; aqui apresentamos só o total. Mantém apenas marcas
// com recompra represada > 0 (as que têm sugestão). Respeita o filtro global, igual aos
// cards do Cockpit. A soma de un/R$ desta lista é idêntica ao card da curva.
function produtosDaCurvaConsolidado(curvaLetra) {
    const prods = produtosCockpit().filter(p => p.curva === curvaLetra);
    const porSku = {};
    prods.forEach(p => {
        const r = recompraDe(p);
        const s = porSku[p.sku] || (porSku[p.sku] = {
            sku: p.sku, produto: rotuloProduto(p),
            un: 0, rs: 0, estoqueLivre: 0, vendaMedia: 0, nCD: 0, defCD: 0, cds: []
        });
        s.un += r.un; s.rs += r.rs;
        s.estoqueLivre += p.estoqueLivre; s.vendaMedia += p.vendaMedia;
        s.nCD++;
        if (p.status === 'deficit') s.defCD++;
        // Linha por CD daquela marca (para o drill-down: quais CDs estão em déficit)
        s.cds.push({ cd: p.cd, status: p.status, diasLivre: p.diasLivre, un: r.un, rs: r.rs });
    });
    return Object.values(porSku)
        .filter(s => s.un > 0)
        .map(s => {
            s.diasLivre = s.vendaMedia > 0 ? (s.estoqueLivre / s.vendaMedia) * 30 : 0;
            s.status = s.vendaMedia > 0 ? classificar(s.diasLivre) : 'sem-sinal';
            return s;
        })
        .sort((a, b) => b.rs - a.rs);
}

// Fecha o modal de curva com a tecla Esc
function cvModalEsc(e) { if (e.key === 'Escape') fecharModalCurva(); }

function fecharModalCurva() {
    const back = document.getElementById('cvModalBackdrop');
    if (back) back.remove();
    document.body.style.overflow = '';
    document.removeEventListener('keydown', cvModalEsc);
}

// Abre o modal com os produtos da curva e suas sugestões, consolidados no total (todos os CDs).
function abrirModalCurva(curvaLetra) {
    fecharModalCurva();   // instância única
    const lista = produtosDaCurvaConsolidado(curvaLetra);
    const totUn = lista.reduce((s, x) => s + x.un, 0);
    const totRS = lista.reduce((s, x) => s + x.rs, 0);
    const metaTxt = metaDiasGlobal;

    const linhas = lista.map((s, i) => {
        const cob = s.vendaMedia > 0 ? `≈ ${Math.round(s.diasLivre).toLocaleString('pt-BR')} d` : 'sem giro';
        const cobCls = s.status === 'deficit' ? 'cvm-cob-def'
            : (s.status === 'sem-giro' || s.status === 'excesso') ? 'cvm-cob-exc' : 'cvm-cob-ok';
        const expansivel = s.defCD > 0;
        const cdTxt = expansivel
            ? `${s.nCD} ${s.nCD === 1 ? 'CD' : 'CDs'} · <strong>${s.defCD} em déficit</strong>`
            : `${s.nCD} ${s.nCD === 1 ? 'CD' : 'CDs'}`;

        const linha = `<tr class="cvm-row${expansivel ? ' cvm-expandable' : ''}"${expansivel ? ` data-i="${i}"` : ''}>
            <td class="cvm-rank">${expansivel ? '<span class="cvm-caret">▸</span>' : ''}${i + 1}</td>
            <td class="cvm-prod">${escapeHtml(s.produto)}</td>
            <td class="cvm-un">${Math.round(s.un).toLocaleString('pt-BR')}</td>
            <td class="cvm-rs">${formatBRLCheio(s.rs)}</td>
            <td class="cvm-cob ${cobCls}">${cob}</td>
            <td class="cvm-cd">${cdTxt}</td>
        </tr>`;

        if (!expansivel) return linha;

        // Drill-down: CDs em déficit (cobertura de estoque livre < 45 dias), mais críticos primeiro.
        // Cada CD é uma linha na MESMA tabela, alinhada sob Un a comprar / R$ / Cobertura.
        const defCds = s.cds.filter(c => c.status === 'deficit').sort((a, b) => a.diasLivre - b.diasLivre);
        const unDef = defCds.reduce((x, c) => x + c.un, 0);
        const unFora = Math.round(s.un - unDef);
        const nFora = s.cds.filter(c => c.status !== 'deficit' && c.un > 0).length;

        const cabecalho = `<tr class="cvm-sub cvm-sub-head" data-parent="${i}">
            <td></td>
            <td colspan="5" class="cvm-sub-label">${defCds.length} ${defCds.length === 1 ? 'CD em déficit' : 'CDs em déficit'} <span class="cvm-sub-hint">cobertura abaixo de 45 dias, mais crítico primeiro</span></td>
        </tr>`;

        const linhasCD = defCds.map(c => `<tr class="cvm-sub cvm-sub-cd" data-parent="${i}">
            <td></td>
            <td class="cvm-prod cvm-sub-cd-name">${escapeHtml(c.cd)}</td>
            <td class="cvm-un cvm-sub-un">${c.un === 0 ? '<span class="cvm-cd-tag" title="cobertura abaixo de 45 dias, mas o que está a caminho já cobre">coberto</span>' : Math.round(c.un).toLocaleString('pt-BR')}</td>
            <td class="cvm-rs cvm-sub-rs">${formatBRLCheio(c.rs)}</td>
            <td class="cvm-cob cvm-cob-def">≈ ${Math.round(c.diasLivre).toLocaleString('pt-BR')} d</td>
            <td class="cvm-cd"></td>
        </tr>`).join('');

        const rodape = unFora > 0 ? `<tr class="cvm-sub cvm-sub-foot" data-parent="${i}">
            <td></td>
            <td colspan="5">+ ${unFora.toLocaleString('pt-BR')} un a comprar em ${nFora} ${nFora === 1 ? 'CD' : 'CDs'} fora de déficit (45-60d ou em reposição)</td>
        </tr>` : '';

        return linha + cabecalho + linhasCD + rodape;
    }).join('');

    const corpo = lista.length
        ? `<table class="cv-modal-table">
                <thead><tr>
                    <th class="cvm-rank">#</th>
                    <th>Produto</th>
                    <th class="cvm-num">Un a comprar</th>
                    <th class="cvm-num">R$ represado</th>
                    <th class="cvm-num">Cobertura</th>
                    <th>Presença</th>
                </tr></thead>
                <tbody>${linhas}</tbody>
            </table>`
        : `<div class="cv-modal-vazio">Nenhum produto com recompra represada nesta curva${temFiltroGlobal() ? ' dentro do filtro atual' : ''}.</div>`;

    const html = `
        <div class="cv-modal" role="dialog" aria-modal="true" aria-label="Produtos da Curva ${curvaLetra}">
            <div class="cv-modal-head">
                <div class="cv-modal-head-top">
                    <div>
                        <div class="cv-modal-eyebrow">Recompra represada · Curva ${curvaLetra}</div>
                        <div class="cv-modal-tot">${formatBRLCheio(totRS)}</div>
                    </div>
                    <button type="button" class="cv-modal-x" aria-label="Fechar">×</button>
                </div>
                <div class="cv-modal-sub">${lista.length.toLocaleString('pt-BR')} ${lista.length === 1 ? 'produto' : 'produtos'} · ${Math.round(totUn).toLocaleString('pt-BR')} un a comprar · soma de todos os CDs · meta ${metaTxt} dias</div>
            </div>
            <div class="cv-modal-body">${corpo}</div>
            <div class="cv-modal-foot">Clique num produto com déficit para ver os CDs. Sugestão por produto = soma das necessidades de cada CD (a reposição é por CD; aqui mostra-se só o total). A cobertura é consolidada nos CDs. Mesma conta da aba Sugestões de Compra.</div>
        </div>`;

    const back = document.createElement('div');
    back.className = 'cv-modal-backdrop';
    back.id = 'cvModalBackdrop';
    back.innerHTML = html;
    document.body.appendChild(back);
    document.body.style.overflow = 'hidden';

    back.addEventListener('click', e => { if (e.target === back) fecharModalCurva(); });
    back.querySelector('.cv-modal-x').addEventListener('click', fecharModalCurva);
    document.addEventListener('keydown', cvModalEsc);

    // Expandir/recolher os CDs em déficit de cada produto
    back.querySelectorAll('.cvm-row.cvm-expandable').forEach(row => {
        row.addEventListener('click', () => {
            const i = row.getAttribute('data-i');
            const aberto = row.classList.toggle('open');
            back.querySelectorAll('.cvm-sub[data-parent="' + i + '"]').forEach(sub => {
                sub.classList.toggle('open', aberto);
            });
        });
    });
}

// Composição da recompra/déficit por GIRO. Produtos sem histórico de venda
// (vendaMedia <= 0) caem em "déficit" por terem 0 dias de cobertura, mas a conta de
// compra os zera (não há demanda a repor). Este corte separa o déficit real (com giro)
// do ruído (sem giro), e confirma que a recompra represada vem só de itens com giro.
function composicaoGiro(prods) {
    let rsComGiro = 0, rsSemGiro = 0, unComGiro = 0, unSemGiro = 0;
    let defComGiro = 0, defSemGiro = 0;
    prods.forEach(p => {
        const temGiro = p.vendaMedia > 0;
        if (p.status === 'deficit') { temGiro ? defComGiro++ : defSemGiro++; }
        const r = recompraDe(p);
        if (r.un > 0) {
            if (temGiro) { rsComGiro += r.rs; unComGiro += r.un; }
            else { rsSemGiro += r.rs; unSemGiro += r.un; }
        }
    });
    const totRS = rsComGiro + rsSemGiro;
    return {
        rsComGiro, rsSemGiro, unComGiro, unSemGiro, defComGiro, defSemGiro,
        pctComGiro: totRS > 0 ? (rsComGiro / totRS * 100) : 100
    };
}

// ---- Snapshot compacto da base inteira (para "O que mudou") ----
function montarSnapshot() {
    if (!dashboardData) return null;
    const prods = originalProducts;             // base completa, SEM filtro
    // Congela o retrato SEMPRE na meta de referência (45 dias), seja qual for a meta na
    // tela. Assim duas bases são sempre comparadas no mesmo prazo. Restaura no fim.
    const metaSalva = metaDiasGlobal;
    metaDiasGlobal = META_REF_COMPARATIVO;
    let snap;
    try {
        const tot = agregarNumeros(prods);
        const porCD = {};
        const porSku = {};
        prods.forEach(p => {
            const r = recompraDe(p);
            const c = porCD[p.cd] || (porCD[p.cd] = { cd: p.cd, recompraRS: 0, un: 0, capTotal: 0, capProblema: 0, capAtencao: 0, estoqueLivre: 0, vendaMedia: 0, def: 0 });
            c.capTotal += p.estoqueLivreRS; c.estoqueLivre += p.estoqueLivre; c.vendaMedia += p.vendaMedia;
            if (p.status === 'excesso') c.capAtencao += p.estoqueLivreRS;
            else if (p.status === 'sem-giro') c.capProblema += p.estoqueLivreRS;
            if (p.status === 'deficit') c.def++;
            if (r.un > 0) { c.recompraRS += r.rs; c.un += r.un; }

            const s = porSku[p.sku] || (porSku[p.sku] = { sku: p.sku, material: truncate(p.material, 44), curva: p.curva, recompraRS: 0, un: 0, capRS: 0, estoqueLivre: 0, vendaMedia: 0, nCD: 0, defCD: 0 });
            s.recompraRS += r.rs; s.un += r.un; s.capRS += p.estoqueLivreRS;
            s.estoqueLivre += p.estoqueLivre; s.vendaMedia += p.vendaMedia; s.nCD++;
            if (p.status === 'deficit') s.defCD++;
        });
        Object.values(porCD).forEach(c => { c.diasMedios = c.vendaMedia > 0 ? (c.estoqueLivre / c.vendaMedia) * 30 : 0; });
        Object.values(porSku).forEach(s => {
            s.diasLivre = s.vendaMedia > 0 ? (s.estoqueLivre / s.vendaMedia) * 30 : 0;
            s.status = s.vendaMedia > 0 ? classificar(s.diasLivre) : 'sem-sinal';
            s.sellout = Math.round(s.vendaMedia);            // saída média ponderada (un/mês)
            s.recompraRS = Math.round(s.recompraRS); s.un = Math.round(s.un); s.capRS = Math.round(s.capRS);
            delete s.estoqueLivre; delete s.vendaMedia;      // poda o que não é usado na comparação
        });
        snap = {
            dataKey: chaveDataRef(dashboardData.dataReferencia),
            dataRef: dashboardData.dataReferencia,
            ts: Date.now(),
            metaRef: META_REF_COMPARATIVO,   // prazo em que este retrato foi congelado
            totais: tot,
            porCD,
            porSku
        };
    } finally {
        metaDiasGlobal = metaSalva;          // devolve a meta da tela
    }
    return snap;
}

// Estado da comparação (preenchido por sincronizarHistorico, assíncrono)
let cmpAtual = null;
let cmpAnterior = null;

// Monta o retrato da base atual, encontra o retrato anterior (data-base distinta e mais
// recente antes desta), salva o atual e redesenha Cockpit + "O que mudou".
async function sincronizarHistorico() {
    if (!dashboardData) return;
    cmpAtual = montarSnapshot();
    try {
        const hist = await snapListar();
        const anteriores = hist.filter(s => s.dataKey < cmpAtual.dataKey);
        cmpAnterior = anteriores.length ? anteriores[anteriores.length - 1] : null;
        await snapSalvar(cmpAtual);
    } catch (e) {
        console.warn('Histórico de comparação indisponível:', e);
        cmpAnterior = null;
    }
    updateCockpit();    // agora com o delta (se houver base anterior)
    updateMudancas();
}

// ---- Formatação de variação ----
function fmtSinalRS(v) {
    const s = v > 0 ? '+' : v < 0 ? '−' : '';
    return s + formatBRLCheio(Math.abs(v)).replace('R$ ', 'R$ ');
}
function fmtSinalUn(v) {
    const s = v > 0 ? '+' : v < 0 ? '−' : '';
    return s + Math.abs(Math.round(v)).toLocaleString('pt-BR');
}
function fmtPct(v) {
    if (!isFinite(v)) return '—';
    const s = v > 0 ? '+' : v < 0 ? '−' : '';
    return s + Math.abs(v).toFixed(1).replace('.', ',') + '%';
}
// Classe de cor para um delta. Em recompra/capital travado, SUBIR é ruim (vermelho).
function classeDelta(v, subirEhRuim) {
    if (Math.abs(v) < 0.5) return 'delta-flat';
    const ruim = subirEhRuim ? v > 0 : v < 0;
    return ruim ? 'delta-ruim' : 'delta-bom';
}

// ============================================
// COCKPIT
// ============================================
function updateCockpit() {
    const box = document.getElementById('cockpitContent');
    if (!box) return;
    if (!dashboardData) { box.innerHTML = '<div class="empty-state">Carregue uma planilha para abrir o cockpit.</div>'; return; }

    const filtrado = temFiltroGlobal();
    const prods = produtosCockpit();
    const a = agregarNumeros(prods);
    const faixa = recompraPorFaixa(prods);
    const curva = recompraPorCurva(prods);
    const giro = composicaoGiro(prods);
    const pend = pendenciaTotal(prods);
    const totalFaixa = faixa.deficit + faixa.saudavel + faixa.reposicao + faixa.excesso + faixa['sem-giro'];
    const pctDef = a.recompraRS > 0 ? (faixa.deficit / a.recompraRS * 100) : 0;

    // Delta do número-herói vs base anterior (só na visão cheia). A comparação é SEMPRE
    // no mesmo prazo (META_REF_COMPARATIVO = 45 dias): usa os retratos congelados das duas
    // bases, não o número da tela. Se a base anterior foi salva numa versão antiga (sem
    // metaRef), avisa para reabri-la em vez de subtrair prazos diferentes.
    let heroDelta = '';
    if (!filtrado && cmpAtual && cmpAnterior) {
        const compativel = cmpAtual.metaRef === META_REF_COMPARATIVO && cmpAnterior.metaRef === META_REF_COMPARATIVO;
        if (compativel) {
            const atualRef = cmpAtual.totais.recompraRS;
            const d = atualRef - cmpAnterior.totais.recompraRS;
            const p = cmpAnterior.totais.recompraRS > 0 ? (d / cmpAnterior.totais.recompraRS * 100) : Infinity;
            heroDelta = `<div class="cockpit-hero-delta ${classeDelta(d, true)}">
                ${fmtSinalRS(d)} <span class="cockpit-hero-delta-pct">(${fmtPct(p)})</span>
                <span class="cockpit-hero-delta-base">vs base ${escapeHtml(cmpAnterior.dataRef)} · ${META_REF_COMPARATIVO} dias</span>
            </div>`;
        } else {
            heroDelta = `<div class="cockpit-hero-delta delta-incompat">
                <span class="cockpit-hero-delta-base">base ${escapeHtml(cmpAnterior.dataRef)} foi salva em outra meta — reabra essa planilha uma vez para comparar a ${META_REF_COMPARATIVO} dias</span>
            </div>`;
        }
    }

    const metaTxt = metaDiasGlobal;
    const avisoFiltro = filtrado
        ? `<div class="cockpit-filtro-aviso">Números filtrados pela seleção atual (CD/Curva/Status). A comparação com a base anterior aparece na visão sem filtro. <button type="button" class="cockpit-limpa-filtro" id="cockpitLimpaFiltro">Limpar filtros</button></div>`
        : '';

    // Barra de composição do represada por faixa
    const seg = (v, cor, label) => {
        if (v <= 0 || totalFaixa <= 0) return '';
        const w = v / totalFaixa * 100;
        const txt = w >= 10 ? `${label} · ${formatBRLCheio(v).replace('R$ ', '')}` : '';
        return `<div class="cockpit-seg" style="flex:0 0 ${w.toFixed(2)}%;background:${cor};" title="${label}: ${formatBRLCheio(v)}">${txt}</div>`;
    };

    box.innerHTML = `
        ${avisoFiltro}
        <div class="cockpit-hero">
            <div class="cockpit-hero-eyebrow">Recompra represada · potencial de faturamento Supera</div>
            <div class="cockpit-hero-value">${formatBRLCheio(a.recompraRS)}</div>
            ${heroDelta}
            <div class="cockpit-hero-sub">
                ${a.recompraUn.toLocaleString('pt-BR')} unidades a comprar em ${a.itensCompra.toLocaleString('pt-BR')} pontos (CD × produto),
                para levar a cobertura à meta de <strong>${metaTxt} dias</strong> (ajustável na aba Sugestões).
            </div>
            ${pend.un > 0 ? `<div class="cockpit-hero-pend">
                <span class="cockpit-hero-pend-dot"></span>
                <span><strong>${Math.round(pend.un).toLocaleString('pt-BR')} un</strong> já a caminho — ${Math.round(pend.transito).toLocaleString('pt-BR')} em trânsito + ${Math.round(pend.entrega).toLocaleString('pt-BR')} em entrega · <strong>${formatBRLCheio(pend.rs)}</strong> em PF. Já dentro do Estoque Total, então a recompra acima é líquida do que a SC já pediu.</span>
            </div>` : ''}
            <div class="cockpit-hero-nucleo ${pctDef >= 50 ? 'nucleo-alto' : ''}">
                <strong>${formatBRLCheio(faixa.deficit)}</strong> disso (${pctDef.toFixed(0)}%) é déficit crítico — itens abaixo de 45 dias, a parte que a SC deveria estar disparando agora.
            </div>
        </div>

        <div class="cockpit-giro">
            <span class="cockpit-giro-ok"><span class="cockpit-giro-dot"></span><strong>${giro.pctComGiro.toFixed(0)}%</strong> da recompra represada é de produtos com giro real (${giro.unComGiro.toLocaleString('pt-BR')} un)</span>
            ${giro.defSemGiro > 0 ? `<span class="cockpit-giro-alerta">${giro.defSemGiro.toLocaleString('pt-BR')} ${giro.defSemGiro === 1 ? 'item aparece' : 'itens aparecem'} em déficit sem histórico de venda — ficam de fora da conta de compra</span>` : ''}
        </div>

        <div class="cockpit-comp">
            <div class="cockpit-comp-head">Composição do represada por faixa de cobertura</div>
            <div class="cockpit-comp-bar">
                ${seg(faixa.deficit, 'var(--deficit)', 'Déficit < 45d')}
                ${seg(faixa.saudavel, 'var(--excesso-light)', 'Top-up 45–60d')}
                ${seg(faixa.reposicao, 'var(--reposicao)', 'Em reposição')}
                ${seg(faixa.excesso, 'var(--saudavel-light)', 'Ajuste fino')}
            </div>
            <div class="cockpit-comp-leg">
                <span class="cockpit-leg"><span class="cockpit-leg-dot" style="background:var(--deficit)"></span>Déficit &lt; 45d: <strong>${formatBRLCheio(faixa.deficit)}</strong></span>
                <span class="cockpit-leg"><span class="cockpit-leg-dot" style="background:var(--excesso-light)"></span>Top-up 45–60d: <strong>${formatBRLCheio(faixa.saudavel)}</strong></span>
                ${faixa.reposicao > 0 ? `<span class="cockpit-leg"><span class="cockpit-leg-dot" style="background:var(--reposicao)"></span>Em reposição: <strong>${formatBRLCheio(faixa.reposicao)}</strong></span>` : ''}
            </div>
        </div>

        <div class="cockpit-curva">
            <div class="cockpit-curva-head">Recompra represada por curva</div>
            <div class="cockpit-curva-cards">
                ${['A', 'B', 'C'].map(k => {
                    const cv = curva[k];
                    const pct = a.recompraRS > 0 ? (cv.rs / a.recompraRS * 100) : 0;
                    const clicavel = cv.un > 0;
                    return `<div class="cockpit-curva-card cv-${k}${clicavel ? ' is-link' : ''}"${clicavel ? ` data-curva="${k}" role="button" tabindex="0"` : ''}>
                        <div class="cockpit-curva-letra">Curva ${k}</div>
                        <div class="cockpit-curva-val">${formatBRLCheio(cv.rs)}</div>
                        <div class="cockpit-curva-meta">${pct.toFixed(0)}% do total · ${cv.itens.toLocaleString('pt-BR')} itens · ${cv.un.toLocaleString('pt-BR')} un</div>
                        <div class="cockpit-curva-def">déficit ${formatBRLCheio(cv.deficit)}</div>
                        ${clicavel ? '<div class="cockpit-curva-cta">ver produtos →</div>' : ''}
                    </div>`;
                }).join('')}
            </div>
        </div>

        <div class="cockpit-cards">
            <div class="cockpit-card cockpit-card-link" data-goto="problema">
                <div class="cockpit-card-val sem-giro">${formatBRLCheio(a.capProblema)}</div>
                <div class="cockpit-card-lbl">Capital travado · problema (&gt; 100 dias)</div>
                <div class="cockpit-card-meta">${a.cont['sem-giro'].toLocaleString('pt-BR')} itens · ${a.capTotal > 0 ? (a.capProblema / a.capTotal * 100).toFixed(0) : 0}% do capital livre</div>
                ${deltaCard(!filtrado, a.capProblema, cmpAnterior && cmpAnterior.totais.capProblema, true)}
            </div>
            <div class="cockpit-card cockpit-card-link" data-goto="atencao">
                <div class="cockpit-card-val excesso">${formatBRLCheio(a.capAtencao)}</div>
                <div class="cockpit-card-lbl">Capital em atenção (60–100 dias)</div>
                <div class="cockpit-card-meta">${a.cont.excesso.toLocaleString('pt-BR')} itens · ${a.capTotal > 0 ? (a.capAtencao / a.capTotal * 100).toFixed(0) : 0}% do capital livre</div>
                ${deltaCard(!filtrado, a.capAtencao, cmpAnterior && cmpAnterior.totais.capAtencao, true)}
            </div>
            <div class="cockpit-card cockpit-card-link" data-goto="compra">
                <div class="cockpit-card-val">${a.compraNec.toLocaleString('pt-BR')}</div>
                <div class="cockpit-card-lbl">Itens com necessidade de compra (meta 45 dias)</div>
                <div class="cockpit-card-meta">déficit &lt; 45d ${a.compraDef.toLocaleString('pt-BR')} · 45–60d projetando queda ${a.compra4560.toLocaleString('pt-BR')} · + ${(a.cont.deficit + a.cont.saudavel + a.cont.reposicao - a.compraNec).toLocaleString('pt-BR')} já com reposição a caminho</div>
            </div>
            <div class="cockpit-card">
                <div class="cockpit-card-val">${formatBRLCheio(a.capTotal)}</div>
                <div class="cockpit-card-lbl">Capital em estoque livre (total)</div>
                <div class="cockpit-card-meta">${a.n.toLocaleString('pt-BR')} pontos de estoque ativos</div>
            </div>
        </div>

        <div class="cockpit-foot">Base ${escapeHtml(dashboardData.dataReferencia)} · ${a.cont.deficit + a.cont.saudavel + a.cont.reposicao + a.cont.excesso + a.cont['sem-giro'] + (a.cont['sem-sinal'] || 0) + (a.cont['parado'] || 0)} itens classificados. A recompra represada usa a mesma conta da aba Sugestões de Compra.</div>
    `;

    // Navegação dos cards
    box.querySelectorAll('.cockpit-card-link').forEach(card => {
        card.addEventListener('click', () => {
            const goto = card.getAttribute('data-goto');
            if (goto === 'compra') {
                const recF = document.getElementById('recFilter');
                if (recF) recF.value = 'deficit';
                updateRecommendations();
                ativarAba('recommendations');
            } else {
                filtroStatus.clear();
                filtroStatus.add(goto === 'problema' ? 'sem-giro' : 'excesso');
                sortBy.value = 'dias-desc';
                aplicarTudo();
                ativarAba('products');
            }
        });
    });
    // Clique nos cards de curva abre o modal de produtos consolidados (todos os CDs somados)
    box.querySelectorAll('.cockpit-curva-card.is-link').forEach(card => {
        const abrir = () => abrirModalCurva(card.getAttribute('data-curva'));
        card.addEventListener('click', abrir);
        card.addEventListener('keydown', e => {
            if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); abrir(); }
        });
    });

    const limpa = document.getElementById('cockpitLimpaFiltro');
    if (limpa) limpa.addEventListener('click', () => {
        filtroCD.clear(); filtroCurva.clear(); filtroStatus.clear();
        productSearch.value = '';
        aplicarTudo();
    });
}

// Mini-delta para os cards do cockpit (capital). Só aparece na visão cheia e com base anterior.
function deltaCard(mostrar, atual, anterior, subirEhRuim) {
    if (!mostrar || anterior == null) return '';
    const d = atual - anterior;
    if (Math.abs(d) < 0.5) return `<div class="cockpit-card-delta delta-flat">sem variação vs base anterior</div>`;
    const p = anterior > 0 ? (d / anterior * 100) : Infinity;
    return `<div class="cockpit-card-delta ${classeDelta(d, subirEhRuim)}">${fmtSinalRS(d)} (${fmtPct(p)}) vs anterior</div>`;
}

// ============================================
// MAPA DE GARGALO (déficit crítico × capital travado por CD)
// ============================================
function updateGargalo() {
    const box = document.getElementById('gargaloContent');
    if (!box) return;
    if (!dashboardData) { box.innerHTML = '<div class="empty-state">Carregue uma planilha para mapear os gargalos por CD.</div>'; return; }

    // Respeita só o filtro de Curva (global), igual às demais visões por CD.
    const prods = originalProducts.filter(p => !filtroCurva.size || filtroCurva.has(p.curva));
    const porCD = {};
    prods.forEach(p => {
        const c = porCD[p.cd] || (porCD[p.cd] = { cd: p.cd, defRS: 0, travado: 0, nDef: 0, nEnc: 0 });
        if (p.status === 'sem-giro') { c.travado += p.estoqueLivreRS; c.nEnc++; }
        const r = recompraDe(p);
        if (r.un > 0 && p.status === 'deficit') { c.defRS += r.rs; c.nDef++; }
    });

    // Só entram CDs com pelo menos uma das duas dores (senão poluem o quadrante).
    const lista = Object.values(porCD).filter(c => c.defRS > 0 || c.travado > 0);
    if (!lista.length) { box.innerHTML = '<div class="empty-state">Nenhum CD com déficit ou capital travado nesta seleção.</div>'; return; }

    const mediana = arr => {
        const s = [...arr].sort((a, b) => a - b);
        const m = Math.floor(s.length / 2);
        return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
    };
    const medDef = mediana(lista.map(c => c.defRS));
    const medTrv = mediana(lista.map(c => c.travado));
    const maxDef = Math.max(1, ...lista.map(c => c.defRS));
    const maxTrv = Math.max(1, ...lista.map(c => c.travado));

    lista.forEach(c => {
        const altoDef = c.defRS >= medDef;
        const altoTrv = c.travado >= medTrv;
        c.quad = altoDef && altoTrv ? 'gargalo' : altoDef ? 'comprar' : altoTrv ? 'destravar' : 'equilibrado';
        // Índice 0-100: média geométrica dos dois eixos normalizados. Só fica alto quando
        // AS DUAS dores são altas ao mesmo tempo — é exatamente o que caracteriza o gargalo.
        c.idx = Math.sqrt((c.defRS / maxDef) * (c.travado / maxTrv)) * 100;
    });

    const META = {
        gargalo:     { lbl: 'Gargalo',     acao: 'comprar urgente E destravar o parado' },
        comprar:     { lbl: 'Só comprar',  acao: 'déficit por volume, estoque limpo — repor' },
        destravar:   { lbl: 'Destravar',   acao: 'parado sem ruptura — liquidar / remanejar' },
        equilibrado: { lbl: 'Equilibrado', acao: 'sem dor relevante nos dois eixos' }
    };

    const celula = q => {
        const g = lista.filter(c => c.quad === q).sort((a, b) => b.idx - a.idx);
        const linhas = g.map(c => `
            <div class="gq-cd gq-cd-link" data-cd="${escapeHtml(c.cd)}" title="Ver produtos do CD ${escapeHtml(c.cd)}">
                <span class="gq-cd-nome">${escapeHtml(c.cd)}</span>
                <span class="gq-cd-nums">
                    <span class="gq-cd-def" title="${c.nDef} ${c.nDef === 1 ? 'item' : 'itens'} em déficit (&lt; 45d)">▲ ${formatBRLCheio(c.defRS)}</span>
                    <span class="gq-cd-trv" title="${c.nEnc} ${c.nEnc === 1 ? 'item encalhado' : 'itens encalhados'} (&gt; 100d)">■ ${formatBRLCheio(c.travado)}</span>
                </span>
            </div>`).join('') || '<div class="gq-vazio">nenhum CD aqui</div>';
        return `<div class="gq-cell gq-${q}">
            <div class="gq-cell-head">
                <span class="gq-cell-lbl">${META[q].lbl}</span>
                <span class="gq-cell-cnt">${g.length}</span>
            </div>
            <div class="gq-cell-acao">${META[q].acao}</div>
            ${g.length ? '<div class="gq-cell-cols"><span class="gq-cd-def">▲ déficit</span><span class="gq-cd-trv">■ excesso</span></div>' : ''}
            <div class="gq-cell-body">${linhas}</div>
        </div>`;
    };

    const garg = lista.filter(c => c.quad === 'gargalo');
    const presoG = garg.reduce((s, c) => s + c.travado, 0);
    const defG = garg.reduce((s, c) => s + c.defRS, 0);

    box.innerHTML = `
        <div class="gq-resumo ${garg.length ? 'gq-resumo-alto' : ''}">
            <div class="gq-resumo-n">${garg.length}</div>
            <div class="gq-resumo-txt">
                <strong>${garg.length === 1 ? 'CD em gargalo' : 'CDs em gargalo'}</strong>: déficit crítico e capital parado convivendo no mesmo CD.
                <span class="gq-resumo-cifras">${formatBRLCheio(presoG)} travado &middot; ${formatBRLCheio(defG)} a recomprar</span>
            </div>
        </div>
        <div class="gq-legenda">
            <span><span class="gq-tag-def">▲</span> déficit crítico a recomprar (itens &lt; 45 dias)</span>
            <span><span class="gq-tag-trv">■</span> excesso / capital travado (&gt; 100 dias)</span>
            <span class="gq-legenda-corte">corte = mediana de cada eixo &middot; déficit ${formatBRLCheio(medDef)} &middot; travado ${formatBRLCheio(medTrv)} &middot; clique num CD para ver seus produtos</span>
        </div>
        <div class="gq-matriz">
            ${celula('comprar')}
            ${celula('gargalo')}
            ${celula('equilibrado')}
            ${celula('destravar')}
        </div>`;

    // Clique num CD -> abre um modal com os produtos do CD SEPARADOS em dois blocos
    // (necessidade de compra × excesso travado), com distinção visual clara.
    box.querySelectorAll('.gq-cd-link').forEach(el => {
        el.addEventListener('click', () => {
            const cd = el.getAttribute('data-cd');
            if (cd) abrirGargaloCD(cd);
        });
    });
}

// Drill-down do Gargalo: ao clicar num CD, abre um modal com os produtos SEPARADOS em
// dois blocos — necessidade de compra (déficit < 45d) e excesso travado (> 100d) — em vez
// de uma lista única ordenada. Mesma definição das colunas ▲/■ da matriz.
function gqModalEsc(e) { if (e.key === 'Escape') fecharGargaloModal(); }
function fecharGargaloModal() {
    const back = document.getElementById('gqModalBackdrop');
    if (back) back.remove();
    document.body.style.overflow = '';
    document.removeEventListener('keydown', gqModalEsc);
}
function abrirGargaloCD(cd) {
    fecharGargaloModal();   // instância única
    const prods = originalProducts.filter(p => p.cd === cd && (!filtroCurva.size || filtroCurva.has(p.curva)));

    // BLOCO 1: necessidade de compra = déficit (< 45d) com recompra > 0, mais crítico primeiro.
    const compra = prods.filter(p => p.status === 'deficit' && recompraDe(p).un > 0)
        .map(p => ({ p, m: recompraDe(p) }))
        .sort((a, b) => a.p.diasLivre - b.p.diasLivre);
    // BLOCO 2: excesso travado = sem-giro (> 100d), maior capital parado primeiro.
    const excesso = prods.filter(p => p.status === 'sem-giro')
        .sort((a, b) => b.estoqueLivreRS - a.estoqueLivreRS);

    const totCompraRS = compra.reduce((s, x) => s + x.m.rs, 0);
    const totCompraUn = compra.reduce((s, x) => s + x.m.un, 0);
    const totExcRS = excesso.reduce((s, p) => s + p.estoqueLivreRS, 0);

    const tabCompra = compra.length ? `
        <table class="cv-modal-table gq-sec-table">
            <thead><tr>
                <th>Produto</th>
                <th class="cvm-num">Un a comprar</th>
                <th class="cvm-num">R$ represado</th>
                <th class="cvm-num">Cobertura</th>
            </tr></thead>
            <tbody>${compra.map(x => `
                <tr>
                    <td class="cvm-prod">${escapeHtml(rotuloProduto(x.p))}</td>
                    <td class="cvm-un">${fmtUn(x.m.un)}</td>
                    <td class="cvm-rs">${formatBRLCheio(x.m.rs)}</td>
                    <td class="cvm-cob cvm-cob-def">≈ ${fmtUn(x.p.diasLivre)} d</td>
                </tr>`).join('')}</tbody>
        </table>` : '<div class="gq-sec-vazio">Nenhum item em déficit a comprar neste CD.</div>';

    const tabExcesso = excesso.length ? `
        <table class="cv-modal-table gq-sec-table">
            <thead><tr>
                <th>Produto</th>
                <th class="cvm-num">Estoque livre</th>
                <th class="cvm-num">R$ travado</th>
                <th class="cvm-num">Cobertura</th>
            </tr></thead>
            <tbody>${excesso.map(p => `
                <tr>
                    <td class="cvm-prod">${escapeHtml(rotuloProduto(p))}</td>
                    <td class="cvm-un gq-un-neutro">${fmtUn(p.estoqueLivre)} un</td>
                    <td class="cvm-rs">${formatBRLCheio(p.estoqueLivreRS)}</td>
                    <td class="cvm-cob cvm-cob-exc">≈ ${fmtUn(p.diasLivre)} d</td>
                </tr>`).join('')}</tbody>
        </table>` : '<div class="gq-sec-vazio">Nenhum item travado (> 100 dias) neste CD.</div>';

    const html = `
        <div class="cv-modal" role="dialog" aria-modal="true" aria-label="Gargalo do CD ${escapeHtml(cd)}">
            <div class="cv-modal-head">
                <div class="cv-modal-head-top">
                    <div>
                        <div class="cv-modal-eyebrow">Gargalo · CD</div>
                        <div class="cv-modal-tot">${escapeHtml(cd)}</div>
                    </div>
                    <button type="button" class="cv-modal-x" aria-label="Fechar">×</button>
                </div>
                <div class="gq-tabs" role="tablist">
                    <button type="button" class="gq-tab gq-tab-def active" data-sec="compra" role="tab"><span class="gq-tab-ico">▲</span> <strong>${compra.length}</strong> a comprar <span class="gq-tab-rs">${formatBRLCheio(totCompraRS)}</span></button>
                    <button type="button" class="gq-tab gq-tab-trv" data-sec="excesso" role="tab"><span class="gq-tab-ico">■</span> <strong>${excesso.length}</strong> em excesso <span class="gq-tab-rs">${formatBRLCheio(totExcRS)}</span></button>
                </div>
            </div>
            <div class="cv-modal-body">
                <div class="gq-sec gq-sec-compra" data-sec="compra">
                    <div class="gq-sec-hint">déficit &lt; 45 dias · mais crítico primeiro</div>
                    ${tabCompra}
                </div>
                <div class="gq-sec gq-sec-excesso gq-sec-hidden" data-sec="excesso">
                    <div class="gq-sec-hint">&gt; 100 dias · maior capital primeiro</div>
                    ${tabExcesso}
                </div>
            </div>
            <div class="cv-modal-foot">Os dois lados do gargalo, separados: o que falta repor (▲) e o que está parado (■). Mesma conta da aba Sugestões de Compra.</div>
        </div>`;

    const back = document.createElement('div');
    back.className = 'cv-modal-backdrop';
    back.id = 'gqModalBackdrop';
    back.innerHTML = html;
    document.body.appendChild(back);
    document.body.style.overflow = 'hidden';

    back.addEventListener('click', e => { if (e.target === back) fecharGargaloModal(); });
    back.querySelector('.cv-modal-x').addEventListener('click', fecharGargaloModal);
    document.addEventListener('keydown', gqModalEsc);

    // Abas: clicar em "a comprar" / "em excesso" troca qual bloco aparece (sem rolar até o fim).
    back.querySelectorAll('.gq-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            const sec = tab.getAttribute('data-sec');
            back.querySelectorAll('.gq-tab').forEach(t => t.classList.toggle('active', t === tab));
            back.querySelectorAll('.gq-sec').forEach(s => s.classList.toggle('gq-sec-hidden', s.getAttribute('data-sec') !== sec));
            back.querySelector('.cv-modal-body').scrollTop = 0;
        });
    });
}

// ============================================
// O QUE MUDOU (atual × base anterior)
// ============================================
function updateMudancas() {
    const box = document.getElementById('mudancasContent');
    if (!box) return;
    if (!dashboardData) { box.innerHTML = '<div class="empty-state">Carregue uma planilha para comparar com a base anterior.</div>'; return; }
    if (!cmpAtual) { box.innerHTML = '<div class="empty-state">Preparando comparação…</div>'; return; }
    if (!cmpAnterior) {
        box.innerHTML = `<div class="mudancas-primeira">
            <div class="mudancas-primeira-ic">🗂️</div>
            <div class="mudancas-primeira-tit">Esta é a primeira base guardada neste navegador</div>
            <div class="mudancas-primeira-txt">O painel guardou o retrato da base <strong>${escapeHtml(cmpAtual.dataRef)}</strong>. Quando você carregar uma base mais nova aqui, esta aba mostra exatamente o que mudou: o que entrou em déficit, o que drenou, e como mexeu a recompra represada.</div>
        </div>`;
        return;
    }

    if (cmpAnterior.metaRef !== META_REF_COMPARATIVO || cmpAtual.metaRef !== META_REF_COMPARATIVO) {
        box.innerHTML = `<div class="mudancas-primeira">
            <div class="mudancas-primeira-ic">🔄</div>
            <div class="mudancas-primeira-tit">Base anterior salva em outra meta</div>
            <div class="mudancas-primeira-txt">O retrato da base <strong>${escapeHtml(cmpAnterior.dataRef)}</strong> foi guardado numa versão anterior, em prazo diferente de ${META_REF_COMPARATIVO} dias. Reabra essa planilha uma vez no Antigravity para regravar o retrato a ${META_REF_COMPARATIVO} dias — aí a comparação volta a aparecer no mesmo prazo.</div>
        </div>`;
        return;
    }

    const A = cmpAtual, B = cmpAnterior;   // A = atual, B = anterior
    const dRecompra = A.totais.recompraRS - B.totais.recompraRS;
    const dProblema = A.totais.capProblema - B.totais.capProblema;
    const dAtencao = A.totais.capAtencao - B.totais.capAtencao;
    const pRecompra = B.totais.recompraRS > 0 ? (dRecompra / B.totais.recompraRS * 100) : Infinity;
    const pProblema = B.totais.capProblema > 0 ? (dProblema / B.totais.capProblema * 100) : Infinity;
    const pAtencao = B.totais.capAtencao > 0 ? (dAtencao / B.totais.capAtencao * 100) : Infinity;

    // União de SKUs e transições de faixa
    const skus = new Set([...Object.keys(A.porSku), ...Object.keys(B.porSku)]);
    let entItens = 0, entRS = 0;       // entraram em déficit
    let saiItens = 0, saiRS = 0;       // saíram de déficit
    let virouEnc = 0, virouEncRS = 0;  // viraram encalhado (>100d)
    let drenouEnc = 0, drenouEncRS = 0;// drenaram (saíram de encalhado)
    const subiu = [], desceu = [], novos = [], sumiu = [];

    skus.forEach(k => {
        const a = A.porSku[k], b = B.porSku[k];
        const sa = a ? a.status : null, sb = b ? b.status : null;
        const rsA = a ? a.recompraRS : 0, rsB = b ? b.recompraRS : 0;
        const dRS = rsA - rsB;

        if (a && !b) { if (rsA > 0) novos.push({ ...a, dRS: rsA }); }
        else if (!a && b) { if (rsB > 0) sumiu.push({ ...b, dRS: -rsB }); }
        else {
            const ehDef = s => s === 'deficit';
            const ehEnc = s => s === 'sem-giro';
            if (!ehDef(sb) && ehDef(sa)) { entItens++; entRS += rsA; }
            if (ehDef(sb) && !ehDef(sa)) { saiItens++; saiRS += rsB; }
            if (!ehEnc(sb) && ehEnc(sa)) { virouEnc++; virouEncRS += a.capRS; }
            if (ehEnc(sb) && !ehEnc(sa)) { drenouEnc++; drenouEncRS += b.capRS; }
            if (dRS > 0) subiu.push({ ...a, dRS, statusAnt: sb });
            else if (dRS < 0) desceu.push({ ...a, dRS, statusAnt: sb, rsB });
        }
    });

    subiu.sort((x, y) => y.dRS - x.dRS);
    desceu.sort((x, y) => x.dRS - y.dRS);
    novos.sort((x, y) => y.dRS - x.dRS);
    sumiu.sort((x, y) => x.dRS - y.dRS);

    const STBADGE = {
        deficit: ['Déficit', 'st-deficit'], saudavel: ['45–60d', 'st-saud'], reposicao: ['Reposição', 'st-repo'],
        excesso: ['Atenção', 'st-exc'], 'sem-giro': ['Encalhado', 'st-enc'], 'sem-sinal': ['Sem sinal', 'st-sinal'], 'parado': ['Parado', 'st-parado']
    };
    const badge = st => { const b = STBADGE[st] || [st, '']; return `<span class="mud-badge ${b[1]}">${b[0]}</span>`; };

    const linhaSku = (s, mostraAnt) => `
        <div class="mud-item">
            <div class="mud-item-main">
                <span class="mud-item-nome">${escapeHtml(s.material)}</span>
                <span class="mud-item-trans">${mostraAnt && s.statusAnt ? badge(s.statusAnt) + '<span class="mud-arrow">→</span>' : ''}${badge(s.status)}</span>
            </div>
            <div class="mud-item-side">
                <span class="mud-item-rs ${classeDelta(s.dRS, true)}">${fmtSinalRS(s.dRS)}</span>
                <span class="mud-item-sub">${s.nCD || 0} CDs · sell-out ${(s.sellout || 0).toLocaleString('pt-BR')}/mês</span>
            </div>
        </div>`;

    const listaOuVazio = (arr, render, vazio, mostraAnt) =>
        arr.length ? arr.slice(0, 12).map(s => render(s, mostraAnt)).join('') : `<div class="mud-vazio">${vazio}</div>`;

    // Movimento por CD
    const cdsKeys = new Set([...Object.keys(A.porCD), ...Object.keys(B.porCD)]);
    const cdMov = [];
    cdsKeys.forEach(k => {
        const a = A.porCD[k], b = B.porCD[k];
        const d = (a ? a.recompraRS : 0) - (b ? b.recompraRS : 0);
        if (Math.abs(d) >= 1) cdMov.push({ cd: k, d, atual: a ? a.recompraRS : 0 });
    });
    cdMov.sort((x, y) => Math.abs(y.d) - Math.abs(x.d));

    box.innerHTML = `
        <div class="mud-cab">
            Comparando <strong>${escapeHtml(A.dataRef)}</strong> com <strong>${escapeHtml(B.dataRef)}</strong>
            <span class="mud-cab-sub">retratos guardados automaticamente neste navegador</span>
        </div>

        <div class="mud-kpis">
            <div class="mud-kpi">
                <div class="mud-kpi-lbl">Recompra represada</div>
                <div class="mud-kpi-val">${formatBRLCheio(A.totais.recompraRS)}</div>
                <div class="mud-kpi-delta ${classeDelta(dRecompra, true)}">${fmtSinalRS(dRecompra)} <span>(${fmtPct(pRecompra)})</span></div>
            </div>
            <div class="mud-kpi">
                <div class="mud-kpi-lbl">Capital travado (&gt; 100d)</div>
                <div class="mud-kpi-val">${formatBRLCheio(A.totais.capProblema)}</div>
                <div class="mud-kpi-delta ${classeDelta(dProblema, true)}">${fmtSinalRS(dProblema)} <span>(${fmtPct(pProblema)})</span></div>
            </div>
            <div class="mud-kpi">
                <div class="mud-kpi-lbl">Capital em atenção (60–100d)</div>
                <div class="mud-kpi-val">${formatBRLCheio(A.totais.capAtencao)}</div>
                <div class="mud-kpi-delta ${classeDelta(dAtencao, true)}">${fmtSinalRS(dAtencao)} <span>(${fmtPct(pAtencao)})</span></div>
            </div>
        </div>

        <div class="mud-transicoes">
            <div class="mud-trans-card mud-trans-ruim">
                <div class="mud-trans-num">${entItens}</div>
                <div class="mud-trans-lbl">entraram em déficit</div>
                <div class="mud-trans-rs">+${formatBRLCheio(entRS).replace('R$ ', 'R$ ')} de recompra nova</div>
            </div>
            <div class="mud-trans-card mud-trans-bom">
                <div class="mud-trans-num">${saiItens}</div>
                <div class="mud-trans-lbl">saíram de déficit</div>
                <div class="mud-trans-rs">−${formatBRLCheio(saiRS).replace('R$ ', 'R$ ')} de recompra resolvida</div>
            </div>
            <div class="mud-trans-card mud-trans-ruim">
                <div class="mud-trans-num">${virouEnc}</div>
                <div class="mud-trans-lbl">viraram encalhado (&gt;100d)</div>
                <div class="mud-trans-rs">${formatBRLCheio(virouEncRS)} de capital novo travado</div>
            </div>
            <div class="mud-trans-card mud-trans-bom">
                <div class="mud-trans-num">${drenouEnc}</div>
                <div class="mud-trans-lbl">drenaram (saíram de encalhado)</div>
                <div class="mud-trans-rs">${formatBRLCheio(drenouEncRS)} de capital destravado</div>
            </div>
        </div>

        <div class="mud-colunas">
            <div class="mud-col">
                <div class="mud-col-tit"><span class="delta-ruim">▲</span> Subiu a recompra represada</div>
                <div class="mud-col-sub">SKUs que mais aumentaram a necessidade de compra</div>
                ${listaOuVazio(subiu, linhaSku, 'Nada subiu de forma relevante.', true)}
            </div>
            <div class="mud-col">
                <div class="mud-col-tit"><span class="delta-bom">▼</span> Drenou / melhorou</div>
                <div class="mud-col-sub">SKUs que mais reduziram a recompra represada</div>
                ${listaOuVazio(desceu, linhaSku, 'Nada drenou de forma relevante.', true)}
            </div>
        </div>

        ${(novos.length || sumiu.length) ? `
        <div class="mud-colunas">
            <div class="mud-col">
                <div class="mud-col-tit">Apareceram na base</div>
                <div class="mud-col-sub">SKUs com recompra que não existiam na base anterior</div>
                ${listaOuVazio(novos, linhaSku, 'Nenhum SKU novo com recompra.', false)}
            </div>
            <div class="mud-col">
                <div class="mud-col-tit">Saíram da base</div>
                <div class="mud-col-sub">SKUs que tinham recompra e não estão mais</div>
                ${listaOuVazio(sumiu, linhaSku, 'Nenhum SKU sumiu.', false)}
            </div>
        </div>` : ''}

        ${cdMov.length ? `
        <div class="mud-cdmov">
            <div class="mud-col-tit">Movimento por CD (recompra represada)</div>
            <div class="mud-cdmov-grid">
                ${cdMov.slice(0, 12).map(c => `
                    <div class="mud-cdmov-item">
                        <span class="mud-cdmov-cd">${escapeHtml(c.cd)}</span>
                        <span class="mud-cdmov-d ${classeDelta(c.d, true)}">${fmtSinalRS(c.d)}</span>
                    </div>`).join('')}
            </div>
        </div>` : ''}
    `;
}

function handleTabClick(e) {
    const tab = e.target.getAttribute('data-tab');
    tabButtons.forEach(b => b.classList.remove('active'));
    tabContents.forEach(c => c.classList.remove('active'));
    e.target.classList.add('active');
    document.getElementById(tab).classList.add('active');
    if (tab === 'cd-sales') renderVendasCD();
    if (tab === 'ritmo') updateRitmo();
    if (tab === 'cockpit') updateCockpit();
    if (tab === 'gargalo') updateGargalo();
    if (tab === 'mudancas') updateMudancas();
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
    return { deficit: 'Compra', saudavel: 'Saudável', reposicao: 'Em reposição', excesso: 'Atenção', 'sem-giro': 'Problema', 'sem-sinal': 'Sem sinal', 'parado': 'Parado' }[status] || status;
}

// Classe de cor pela faixa EXIBIDA. Déficit (< 45) = azul "Compra"; saudável (45-60) =
// verde (acima da meta de 45d); reposição = ciano; atenção/problema laranja/vermelho.
function statusGrupo(status) {
    return { deficit: 'compra', saudavel: 'saudavel', reposicao: 'reposicao', excesso: 'atencao', 'sem-giro': 'problema', 'sem-sinal': 'sem-sinal', 'parado': 'parado' }[status] || status;
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

// Formata quantidade em UNIDADES no padrão pt-BR: separador de milhar (ponto) só depois
// do 3o dígito; vírgula decimal apenas quando há fração. Ex.: 12729 -> "12.729";
// 63 -> "63"; 63,4 -> "63,4". Substitui toFixed, que não agrupa milhar e força ".0".
function fmtUn(x, dec = 0) {
    return Number(x || 0).toLocaleString('pt-BR', { maximumFractionDigits: dec, minimumFractionDigits: 0 });
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
