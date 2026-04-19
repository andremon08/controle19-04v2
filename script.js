// ─── CONFIGURAÇÃO ───────────────────────────────────────────────────────────
const BASE_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSJq1BdeNlo6gvM1vBhtgD88MRevuRrODf2NmVESwH5CMQ6VBkuZMUaNEr8xCoHeJlmnlsJaDV_Cj9L/pub';

const URL_VERBAS     = BASE_URL + '?gid=1303157015&single=true&output=csv';
const URL_SERVIDORES = BASE_URL + '?gid=1533392322&single=true&output=csv';
const URL_CARGOS     = BASE_URL + '?gid=1823673227&single=true&output=csv';
const URL_ESTRUTURAS = BASE_URL + '?gid=46958645&single=true&output=csv';

// ─── REGRAS DE NEGÓCIO ──────────────────────────────────────────────────────
const TETO_VEREADOR  = 92998.45;
const TOLERANCIA     = 0.13;
const MAX_SERVIDORES = 9;

// Cargo especial: não consome vaga, verba nem aparece na estrutura
const CARGO_CEDIDO = 'CEDIDOS DE OUTRAS ENTIDADES SEM ÔNUS';
const isCedido = cargo => cargo.trim().toUpperCase().includes('CEDIDOS DE OUTRAS ENTIDADES');

// Lista exata de lotações especiais com estrutura fixa definida por lei.
// TUDO que não estiver aqui e não começar com "Bloco" ou "Liderança" = Gabinete de Vereador.
const LOTACOES_ESPECIAIS = {
    'GABINETE DA PRESIDÊNCIA':         ['CC-1', 'CC-5', 'CC-6', 'CC-7'],
    'GABINETE DA 1ª VICE-PRESIDÊNCIA': ['CC-4', 'CC-6'],
    'GABINETE DA 2ª VICE-PRESIDÊNCIA': ['CC-4', 'CC-7'],
    'GABINETE DA 1ª SECRETARIA':       ['CC-3', 'CC-6', 'CC-7'],
    'GABINETE DA 2ª SECRETARIA':       ['CC-4', 'CC-6', 'CC-7'],
    'GABINETE DA 3ª SECRETARIA':       ['CC-5', 'CC-7'],
    'GABINETE DA 4ª SECRETARIA':       ['CC-5', 'CC-7'],
};

// ─── ESTADO GLOBAL ──────────────────────────────────────────────────────────
let dadosVerbas     = [];
let dadosServidores = [];
let tabelaCargos    = {};  // { 'CC-1': 12000.00, ... }
let dadosEstruturas = {};  // { 'Gabinete X': ['CC-1','CC-3','CC-6'], ... }
let _todasSugestoes = []; // sugestões geradas para o filtro em tempo real
let saldo_atual     = 0;  // saldo do gabinete atual, usado no filtro

// Estado atual para exportação
let _exportEstado = {
    mes: '', gab: '', tipo: '',
    servidores: [], estrutura: [], responsavel: ''
};

// ─── INICIALIZAÇÃO ──────────────────────────────────────────────────────────
function iniciar() {
    setStatus('', 'Carregando...');

    Promise.all([
        carregarCSV(URL_VERBAS,     'verbas'),
        carregarCSV(URL_SERVIDORES, 'servidores'),
        carregarCSV(URL_CARGOS,     'cargos'),
        carregarCSV(URL_ESTRUTURAS, 'estruturas'),
    ])
    .then(([verbas, servidores, cargos, estruturas]) => {
        try {
            dadosVerbas     = verbas;
            dadosServidores = servidores;
            tabelaCargos    = construirTabelaCargos(cargos);
            dadosEstruturas = construirEstruturas(estruturas);
            preencherFiltros();
            setStatus('ok', 'Dados carregados');
        } catch (err) {
            console.error('[Erro ao processar dados]', err);
            setStatus('erro', 'Erro ao processar dados');
        }
    })
    .catch(err => {
        console.error('[Erro ao carregar]', err);
        setStatus('erro', typeof err === 'string' ? err : 'Erro ao carregar dados');
    });
}

function carregarCSV(url, nome) {
    return new Promise((resolve, reject) => {
        // Timeout de 15s para não ficar pendurado
        const timer = setTimeout(() => {
            reject(`Timeout ao carregar "${nome}". Verifique se a planilha está publicada.`);
        }, 15000);

        Papa.parse(url, {
            download: true,
            header: true,
            skipEmptyLines: true,
            complete: r => {
                clearTimeout(timer);
                if (!r.data || r.data.length === 0) {
                    console.warn(`[${nome}] Aba vazia ou sem dados.`);
                }
                resolve(r.data || []);
            },
            error: e => {
                clearTimeout(timer);
                console.error(`[${nome}] Erro PapaParse:`, e);
                // Resolve com array vazio em vez de rejeitar, para não travar tudo
                // caso uma aba ainda não tenha dados
                resolve([]);
            },
        });
    });
}

function construirTabelaCargos(linhas) {
    const tabela = {};
    linhas.forEach(l => {
        const cargo   = (l['Cargo'] || '').trim();
        const salario = parseMoeda(l['Salário'] || l['Salario'] || '0');
        if (cargo) tabela[cargo] = salario;
    });
    return tabela;
}

// Aba estruturas: colunas Gabinete | Cargo
// Uma linha por cargo da estrutura (com repetições para mesmo CC)
function construirEstruturas(linhas) {
    const estruturas = {};
    linhas.forEach(l => {
        const gab   = (l['Gabinete'] || '').trim();
        const cargo = (l['Cargo'] || '').trim();
        if (!gab || !cargo) return;
        // Extrai código CC do nome completo (ex: "CHEFE DE GABINETE - CC-1" → "CC-1")
        const match = cargo.toUpperCase().match(/CC-\d+/);
        const cc = match ? match[0] : cargo.toUpperCase();
        if (!estruturas[gab]) estruturas[gab] = [];
        estruturas[gab].push(cc);
    });
    return estruturas;
}

// ─── CLASSIFICAÇÃO ──────────────────────────────────────────────────────────
// Retorna 'mesa_diretora', 'bloco' ou 'vereador'
function classificarTipo(gabinete) {
    const g = gabinete.trim().toUpperCase();
    // Verifica lista exata de lotações especiais (case-insensitive)
    for (const nome of Object.keys(LOTACOES_ESPECIAIS)) {
        if (g === nome.toUpperCase()) return 'mesa_diretora';
    }
    // Blocos e lideranças: nome contém essas palavras
    if (/\bbloco\b/i.test(g) || /\blideran/i.test(g)) return 'bloco';
    // Todo o resto = gabinete de vereador
    return 'vereador';
}

// ─── FILTROS ────────────────────────────────────────────────────────────────
function preencherFiltros() {
    const mesSelect = document.getElementById('mesSelect');
    const gabSelect = document.getElementById('gabineteSelect');

    const mesesSet = new Set();
    const gabSet   = new Set();

    dadosVerbas.forEach(l => {
        if (l['Mês'])      mesesSet.add(l['Mês'].trim());
        if (l['Gabinete']) gabSet.add(l['Gabinete'].trim());
    });

    const mesesOrdenados = [...mesesSet].sort((a, b) => {
        const [ma, aa] = a.split('/').map(Number);
        const [mb, ab] = b.split('/').map(Number);
        return aa !== ab ? aa - ab : ma - mb;
    });
    const gabsOrdenados = [...gabSet].sort((a, b) => a.localeCompare(b, 'pt-BR'));

    mesSelect.innerHTML = '<option value="">Selecione o mês...</option>';
    gabSelect.innerHTML = '<option value="">Selecione a lotação...</option>';

    mesesOrdenados.forEach(m => mesSelect.innerHTML += `<option value="${m}">${m}</option>`);
    gabsOrdenados.forEach(g  => gabSelect.innerHTML += `<option value="${g}">${g}</option>`);

    mesSelect.addEventListener('change', atualizarPainel);
    gabSelect.addEventListener('change', atualizarPainel);
}


// ─── ATUALIZAÇÃO PRINCIPAL ──────────────────────────────────────────────────
function atualizarPainel() {
    const mes = document.getElementById('mesSelect').value.trim();
    const gab = document.getElementById('gabineteSelect').value.trim();

    ocultarTudo();
    if (!mes || !gab) return;

    const { inicio, fim } = intervaloMes(mes);

    const verbaMes = dadosVerbas.find(
        l => l['Mês']?.trim() === mes && l['Gabinete']?.trim() === gab
    ) || {};

    // Filtra servidores ativos no mês via datas
    const servidoresMes = dadosServidores
        .filter(l => (l['Gabinete'] || '').trim() === gab)
        .map(l => {
            const admissao       = parseData(l['Admissão']   || l['Admissao']   || '');
            const exoneracao     = parseData(l['Exoneração'] || l['Exoneracao'] || '');
            const ativo          = estaAtivo(admissao, exoneracao, inicio, fim);
            const exoneradoNoMes = !!(exoneracao && exoneracao >= inicio && exoneracao <= fim);
            return { ...l, admissao, exoneracao, ativo, exoneradoNoMes };
        })
        .filter(l => l.ativo);

    const tipo = classificarTipo(gab);
    atualizarTopbarBadges(mes, tipo);

    // Salva estado para exportação
    _exportEstado.mes  = mes;
    _exportEstado.gab  = gab;
    _exportEstado.tipo = tipo;

    if (tipo === 'vereador') {
        renderizarVereador(gab, verbaMes, servidoresMes);
    } else {
        renderizarEspecial(gab, tipo, verbaMes, servidoresMes);
    }
}

// ─── TOPBAR BADGES ──────────────────────────────────────────────────────────
function atualizarTopbarBadges(mes, tipo) {
    let tipoBadge = '';
    if (tipo === 'vereador')       tipoBadge = `<span class="badge badge-tipo-vereador">Gabinete de Vereador</span>`;
    else if (tipo === 'mesa_diretora') tipoBadge = `<span class="badge badge-tipo-especial">Mesa Diretora</span>`;
    else                           tipoBadge = `<span class="badge badge-tipo-especial">Bloco / Liderança</span>`;
    document.getElementById('topbarBadges').innerHTML =
        `<span class="badge badge-mes">${mes}</span>${tipoBadge}`;
}

// ─── PAINEL VEREADOR ────────────────────────────────────────────────────────
function renderizarVereador(gab, verbaMes, servidores) {
    document.getElementById('painelVereador').classList.remove('escondido');

    const responsavel = (verbaMes['Responsável'] || verbaMes['Responsavel'] || '').trim();

    // Separa servidores ativos (não exonerados) dos exonerados no mês
    // Cedidos de outras entidades sem ônus são excluídos de verba, contagem e estrutura
    const servsAtivos     = servidores.filter(s => !s.exoneradoNoMes && !isCedido(s['Cargo'] || ''));
    const servsExonerados = servidores.filter(s => s.exoneradoNoMes);

    // Verba utilizada considera apenas ativos (exonerado = vaga liberada)
    let verbaUtil = 0;
    servsAtivos.forEach(s => {
        const cargo = (s['Cargo'] || '').trim();
        verbaUtil += tabelaCargos[cargo] || 0;
    });

    // ── Estrutura do gabinete ──
    const estrutura = dadosEstruturas[gab] || [];
    const totalVagasEstrutura = estrutura.length;

    const saldo  = TETO_VEREADOR - verbaUtil;
    const pct    = TETO_VEREADOR > 0 ? Math.min((verbaUtil / TETO_VEREADOR) * 100, 100) : 0;
    const nServs = servsAtivos.filter(s => (s['Nome do Servidor'] || '').trim()).length;

    // Vagas baseadas na estrutura; exonerados no mês liberam vagas
    const totalRef = totalVagasEstrutura > 0 ? totalVagasEstrutura : MAX_SERVIDORES;
    const vagasLiv = totalRef - nServs;

    document.getElementById('vVerbaTotal').textContent    = moeda(TETO_VEREADOR);
    document.getElementById('vVerbaUtil').textContent     = moeda(verbaUtil);
    document.getElementById('vSaldo').textContent         = moeda(Math.abs(saldo));
    document.getElementById('vServidores').textContent    = `${nServs} / ${totalRef}`;
    document.getElementById('vServidoresSub').textContent = vagasLiv === 1 ? '1 vaga disponível' : vagasLiv > 1 ? `${vagasLiv} vagas disponíveis` : 'sem vagas disponíveis';
    document.getElementById('vVerbaUtilSub').textContent  = responsavel ? `Resp.: ${responsavel}` : 'soma dos salários';

    const cardSaldo = document.getElementById('vSaldo').closest('.card');
    cardSaldo.classList.remove('card-saldo-negativo');
    if (saldo < 0) {
        cardSaldo.classList.add('card-saldo-negativo');
        document.getElementById('vSaldoSub').textContent = 'acima do teto';
    } else {
        document.getElementById('vSaldoSub').textContent = 'disponível no teto';
    }

    // Barra de progresso
    const fill  = document.getElementById('vProgressoFill');
    const pctEl = document.getElementById('vProgressoPct');
    fill.style.width  = pct.toFixed(1) + '%';
    pctEl.textContent = pct.toFixed(1) + '%';
    fill.classList.remove('aviso', 'perigo');
    if (pct >= 100)     fill.classList.add('perigo');
    else if (pct >= 85) fill.classList.add('aviso');

    // Alertas
    const temCC1 = servsAtivos.some(s => {
        const c = (s['Cargo'] || '').trim().toUpperCase();
        return (c.endsWith('CC-1') || c.endsWith('- CC-1') || c === 'CC-1');
    });
    const alertas = [];
    if (!temCC1) alertas.push({ tipo: 'erro', msg: 'Chefe de Gabinete (CC-1) não está lotado. Este cargo é obrigatório.' });
    if (nServs > MAX_SERVIDORES) alertas.push({ tipo: 'erro', msg: `O gabinete possui ${nServs} servidores, acima do limite legal de ${MAX_SERVIDORES}.` });
    if (verbaUtil > TETO_VEREADOR + TOLERANCIA) {
        alertas.push({ tipo: 'erro', msg: `Verba utilizada (${moeda(verbaUtil)}) excede o teto legal de ${moeda(TETO_VEREADOR)}.` });
    } else if (verbaUtil > TETO_VEREADOR) {
        alertas.push({ tipo: 'aviso', msg: 'Verba dentro da margem de tolerância de R$ 0,13.' });
    }
    if (servsExonerados.length > 0) {
        const nomes = servsExonerados.map(s => (s['Nome do Servidor'] || '').trim()).join(', ');
        alertas.push({ tipo: 'aviso', msg: `Exonerado(s) neste mês: ${nomes}.` });
    }
    if (alertas.length === 0) alertas.push({ tipo: 'ok', msg: 'Gabinete em conformidade com as regras legais.' });
    renderizarAlertas('alertasVereador', alertas);

    // ── Estrutura com cards agrupados ──
    const resumoEl = document.getElementById('vEstruturaResumo');
    const grade    = document.getElementById('gradeEstrutura');

    if (estrutura.length === 0) {
        grade.innerHTML = `<p style="font-size:14px;color:var(--muted);font-style:italic;padding:4px 0">Estrutura não cadastrada para este gabinete.</p>`;
        resumoEl.textContent = '';
    } else {
        // Constrói lista de CCs ativos (cópia fresca, não consumida)
        const cargosAtivosCC = servsAtivos.map(s => extrairCC(s['Cargo'] || ''));

        // Constrói estrutura efetiva: começa com os slots formais e adiciona
        // slots extras se houver mais servidores ativos do que slots na estrutura
        const estruturaEfetiva = [...estrutura];
        const consumivelCheck = [...cargosAtivosCC];
        // Marca quais slots da estrutura formal estão ocupados
        estruturaEfetiva.forEach((cc, i) => {
            const idx = consumivelCheck.indexOf(cc);
            if (idx !== -1) consumivelCheck.splice(idx, 1);
        });
        // Sobram em consumivelCheck os servidores que não têm slot na estrutura — adiciona como extra
        consumivelCheck.forEach(cc => estruturaEfetiva.push(cc));

        // Conta vagos (slots não preenchidos)
        const consumivelVagos = [...cargosAtivosCC];
        const vagosCount = estruturaEfetiva.filter(cc => {
            const idx = consumivelVagos.indexOf(cc);
            if (idx !== -1) { consumivelVagos.splice(idx, 1); return false; }
            return true;
        }).length;

        const totalSlots = estruturaEfetiva.length;
        resumoEl.textContent = vagosCount === 1
            ? `1 vaga livre de ${totalSlots}`
            : vagosCount > 1
                ? `${vagosCount} vagas livres de ${totalSlots}`
                : `${totalSlots} de ${totalSlots} preenchidos`;

        grade.innerHTML = renderizarEstruturaCCs(estruturaEfetiva, [...cargosAtivosCC]);
    }

    // ── Tabela servidores ──
    const tbody = document.getElementById('corpoTabelaVereador');
    const tfoot = document.getElementById('rodapeTabelaVereador');
    tbody.innerHTML = '';

    // Verba total tabela inclui todos (ativos + exonerados no mês para histórico)
    let verbaTotalTabela = 0;
    servidores.forEach(s => {
        const nome      = (s['Nome do Servidor'] || '').trim();
        const cargo     = (s['Cargo'] || '').trim();
        const matricula = (s['Matrícula'] || s['Matricula'] || '').trim();
        const sal       = tabelaCargos[cargo] || 0;
        const admStr    = s.admissao   ? formatarData(s.admissao)   : '—';
        const exoStr    = s.exoneracao ? formatarData(s.exoneracao) : '—';
        if (!nome) return;
        if (!s.exoneradoNoMes) verbaTotalTabela += sal;
        const exoTag = s.exoneradoNoMes ? `<span class="tag-exonerado">Exonerado</span>` : '';
        tbody.innerHTML += `
            <tr class="${s.exoneradoNoMes ? 'tr-exonerado' : ''}">
                <td class="col-matricula">${matricula || '—'}</td>
                <td>${nome}${exoTag}</td>
                <td>${cargo || '—'}</td>
                <td class="col-salario">${sal > 0 ? moeda(sal) : '—'}</td>
                <td class="col-data">${admStr}</td>
                <td class="col-data">${exoStr}</td>
            </tr>`;
    });

    tfoot.innerHTML = `
        <tr>
            <td colspan="3">Total em vigor</td>
            <td class="col-salario col-salario-total">${moeda(verbaUtil)}</td>
            <td colspan="2"></td>
        </tr>`;

    // ── Sugestão de Composição ──
    renderizarSugestao(saldo, nServs);

    // Salva estado para exportação
    _exportEstado.servidores  = servsAtivos;
    _exportEstado.estrutura   = dadosEstruturas[gab] || [];
    _exportEstado.responsavel = responsavel;
}

// ─── PAINEL ESPECIAL ────────────────────────────────────────────────────────
function renderizarEspecial(gab, tipo, verbaMes, servidores) {
    document.getElementById('painelEspecial').classList.remove('escondido');

    const responsavel = (verbaMes['Responsável'] || verbaMes['Responsavel'] || '').trim();
    const elResp       = document.getElementById('eResponsavel');
    const elDestaque   = document.getElementById('eResponsavelDestaque');
    elResp.textContent = responsavel || 'Não informado';
    if (responsavel) {
        elDestaque.style.display = 'flex';
    } else {
        elDestaque.style.display = 'none';
    }

    // Estrutura esperada
    let estrutura = [];
    if (tipo === 'mesa_diretora') {
        const chave = Object.keys(LOTACOES_ESPECIAIS).find(
            k => k.toUpperCase() === gab.trim().toUpperCase()
        );
        estrutura = chave ? LOTACOES_ESPECIAIS[chave] : [];
    } else {
        estrutura = ['CC-8'];
    }

    // Alertas
    const servsAtivosEspecial = servidores.filter(s => !s.exoneradoNoMes);
    const ccsAtivos = servsAtivosEspecial.map(s => extrairCC(s['Cargo'] || ''));
    const vagasLivres = estrutura.filter(cc => {
        const idx = ccsAtivos.indexOf(cc);
        if (idx !== -1) { ccsAtivos.splice(idx, 1); return false; }
        return true;
    });
    const alertas = [];
    if (tipo === 'bloco') {
        const blocoAtivo = servsAtivosEspecial.some(s => extrairCC(s['Cargo'] || '') === 'CC-8');
        if (!blocoAtivo) alertas.push({ tipo: 'aviso', msg: 'O cargo CC-8 desta lotação está vago.' });
        else             alertas.push({ tipo: 'ok',   msg: 'Lotação regularmente ocupada.' });
    } else {
        if (vagasLivres.length === estrutura.length)
            alertas.push({ tipo: 'aviso', msg: 'Nenhum cargo desta lotação está ocupado.' });
        else if (vagasLivres.length > 0)
            alertas.push({ tipo: 'aviso', msg: `${vagasLivres.length} cargo(s) com vaga disponível: ${vagasLivres.join(', ')}.` });
        else
            alertas.push({ tipo: 'ok', msg: 'Todos os cargos da estrutura estão ocupados.' });
    }
    const exoneradosNoMes = servidores.filter(s => s.exoneradoNoMes);
    if (exoneradosNoMes.length > 0) {
        const nomes = exoneradosNoMes.map(s => (s['Nome do Servidor'] || '').trim()).join(', ');
        alertas.push({ tipo: 'aviso', msg: `Exonerado(s) neste mês: ${nomes}.` });
    }
    renderizarAlertas('alertasEspecial', alertas);

    // ── Tabela primeiro ──
    const tbody = document.getElementById('corpoTabelaEspecial');
    tbody.innerHTML = '';
    servidores.forEach(s => {
        const nome      = (s['Nome do Servidor'] || '').trim();
        const cargo     = (s['Cargo'] || '').trim();
        const matricula = (s['Matrícula'] || s['Matricula'] || '').trim();
        const sal       = tabelaCargos[cargo] || 0;
        const admStr    = s.admissao   ? formatarData(s.admissao)   : '—';
        const exoStr    = s.exoneracao ? formatarData(s.exoneracao) : '—';
        if (!nome) return;
        const exoTag = s.exoneradoNoMes ? `<span class="tag-exonerado">Exonerado</span>` : '';
        tbody.innerHTML += `
            <tr class="${s.exoneradoNoMes ? 'tr-exonerado' : ''}">
                <td class="col-matricula">${matricula || '—'}</td>
                <td>${nome}${exoTag}</td>
                <td>${cargo || '—'}</td>
                <td class="col-salario">${sal > 0 ? moeda(sal) : '—'}</td>
                <td class="col-data">${admStr}</td>
                <td class="col-data">${exoStr}</td>
            </tr>`;
    });
    if (!tbody.innerHTML) {
        tbody.innerHTML = `<tr><td colspan="6" style="color:var(--muted);font-style:italic;text-align:center;padding:16px">Nenhum servidor lotado neste período</td></tr>`;
    }

    // ── Estrutura da Lotação com cards ──
    const resumoEl = document.getElementById('eEstruturaResumo');
    const grade    = document.getElementById('gradeEspecial');

    if (estrutura.length === 0) {
        grade.innerHTML = `<p style="font-size:14px;color:var(--muted);font-style:italic;padding:4px 0">Estrutura não definida para esta lotação.</p>`;
        resumoEl.textContent = '';
    } else {
        // Reconstrói cargos ativos para o render (ccsAtivos foi consumido acima)
        const cargosAtivosRender = servsAtivosEspecial.map(s => extrairCC(s['Cargo'] || ''));

        // Conta vagos para o resumo
        const tempVagos = [...cargosAtivosRender];
        const vagosCount = estrutura.filter(cc => {
            const idx = tempVagos.indexOf(cc);
            if (idx !== -1) { tempVagos.splice(idx, 1); return false; }
            return true;
        }).length;

        resumoEl.textContent = vagosCount === 1
            ? `1 vaga livre de ${estrutura.length}`
            : vagosCount > 1
                ? `${vagosCount} vagas livres de ${estrutura.length}`
                : `${estrutura.length} de ${estrutura.length} preenchidos`;

        grade.innerHTML = renderizarEstruturaCCs(estrutura, cargosAtivosRender);
    }

    // Salva estado para exportação
    _exportEstado.servidores  = servsAtivosEspecial;
    _exportEstado.estrutura   = estrutura;
    _exportEstado.responsavel = responsavel;
}

// ─── SUGESTÃO DE COMPOSIÇÃO ─────────────────────────────────────────────────
function renderizarSugestao(saldo, nServsAtivos) {
    saldo_atual = saldo; // salva para uso no filtro em tempo real
    const secao      = document.getElementById('secaoSugestao');
    const introEl    = document.getElementById('sugestaoIntro');
    const gridEl     = document.getElementById('sugestaoGrid');
    const resumo     = document.getElementById('sugestaoResumo');
    const filtroWrap = document.getElementById('sugestaoFiltroWrap');
    const filtroInput = document.getElementById('sugestaoFiltro');

    const vagasRestantes = MAX_SERVIDORES - nServsAtivos;

    if (vagasRestantes <= 0 || saldo <= 0) {
        secao.classList.add('escondido');
        return;
    }
    secao.classList.remove('escondido');

    const vagasStr = vagasRestantes === 1 ? '1 vaga disponível' : `${vagasRestantes} vagas disponíveis`;
    resumo.textContent = `${vagasStr} — saldo ${moeda(saldo)}`;

    // CCs disponíveis excluindo CC-1
    const ccMap = {};
    Object.entries(tabelaCargos).forEach(([nome, sal]) => {
        const cc = extrairCC(nome);
        if (!cc.startsWith('CC-') || cc === 'CC-1') return;
        if (!ccMap[cc] || sal < ccMap[cc]) ccMap[cc] = sal;
    });
    const ccs = Object.entries(ccMap)
        .map(([cc, sal]) => ({ cc, sal }))
        .sort((a, b) => a.sal - b.sal);

    if (ccs.length === 0) {
        introEl.innerHTML = `<p class="sugestao-intro">Tabela de cargos não carregada.</p>`;
        filtroWrap.style.display = 'none';
        return;
    }

    // Gera todas as combinações com repetição de tamanho 1..vagasRestantes
    const sugestoes = [];
    const vistos = new Set();

    function combinar(inicio, atual, custoAtual) {
        if (atual.length > 0) {
            const chave = [...atual].sort().join('+');
            if (!vistos.has(chave)) {
                vistos.add(chave);
                sugestoes.push({ ccs: [...atual], custo: custoAtual });
            }
        }
        if (atual.length === vagasRestantes) return;
        for (let i = inicio; i < ccs.length; i++) {
            const novoCusto = custoAtual + ccs[i].sal;
            if (novoCusto > saldo + TOLERANCIA) break;
            combinar(i, [...atual, ccs[i].cc], novoCusto);
        }
    }
    combinar(0, [], 0);

    if (sugestoes.length === 0) {
        introEl.innerHTML = `<p class="sugestao-intro">Não há cargos que caibam no saldo disponível de ${moeda(saldo)}.</p>`;
        filtroWrap.style.display = 'none';
        gridEl.innerHTML = ''; // limpa cards de gabinete anterior
        _todasSugestoes = [];
        return;
    }

    // Ordena: mais cargos primeiro, depois por custo decrescente
    sugestoes.sort((a, b) =>
        b.ccs.length !== a.ccs.length ? b.ccs.length - a.ccs.length : b.custo - a.custo
    );
    _todasSugestoes = sugestoes;

    const vagasIntro = vagasRestantes === 1 ? '1 vaga disponível' : `${vagasRestantes} vagas disponíveis`;
    introEl.innerHTML = `<p class="sugestao-intro">
        Com <strong>${vagasIntro}</strong> e saldo de <strong>${moeda(saldo)}</strong>,
        abaixo estão todas as combinações possíveis dentro do teto legal de ${moeda(TETO_VEREADOR)}:
    </p>`;

    // Mostra filtro e reseta
    filtroWrap.style.display = 'block';
    filtroInput.value = '';

    renderizarCardsSugestao('');

    // Filtro em tempo real — remove e recria o listener para evitar duplicatas
    const novoInput = filtroInput.cloneNode(true);
    filtroInput.parentNode.replaceChild(novoInput, filtroInput);
    novoInput.addEventListener('input', () => {
        renderizarCardsSugestao(novoInput.value.trim());
    });
}

function renderizarCardsSugestao(filtro) {
    const gridEl = document.getElementById('sugestaoGrid');

    // Suporte a múltiplos termos separados por vírgula
    // Ex: "CC2, CC3" → filtra cards que contenham CC-2 E CC-3
    const termos = filtro
        .split(',')
        .map(t => t.trim().toUpperCase().replace(/\s/g, ''))
        .filter(t => t.length > 0);

    const normalizar = cc => cc.replace('-', '');

    const filtradas = termos.length === 0
        ? _todasSugestoes
        : _todasSugestoes.filter(s => {
            // Cada termo deve estar presente pelo menos uma vez na combinação
            const ccsNorm = s.ccs.map(normalizar);
            return termos.every(termo => {
                const termoNorm = termo.replace('-', '');
                return ccsNorm.some(cc => cc.includes(termoNorm));
            });
        });

    if (filtradas.length === 0) {
        const termosLabel = termos.join(', ');
        gridEl.innerHTML = `<p class="sugestao-sem-resultado">Nenhuma combinação encontrada para "${termosLabel}".</p>`;
        return;
    }

    gridEl.innerHTML = filtradas.map(s => {
        const n = s.ccs.length;
        const saldoPos = saldo_atual - s.custo;
        const contagem = {};
        s.ccs.forEach(cc => { contagem[cc] = (contagem[cc] || 0) + 1; });
        const ccsLabel = Object.entries(contagem)
            .map(([cc, qtd]) => qtd > 1 ? `${qtd}× ${cc}` : cc)
            .join(' + ');
        const cargoStr = n === 1 ? '1 cargo' : `${n} cargos`;
        return `
            <div class="sugestao-card">
                <div class="sugestao-card-titulo">${cargoStr}</div>
                <div class="sugestao-card-ccs">${ccsLabel}</div>
                <div class="sugestao-card-total">+ ${moeda(s.custo)}</div>
                <div class="sugestao-card-saldo">Saldo restante: ${moeda(saldoPos)}</div>
            </div>`;
    }).join('');
}



// Extrai o código CC do nome completo do cargo
// "CHEFE DE GABINETE PARLAMENTAR - CC-1" → "CC-1"
function extrairCC(nomeCargo) {
    const match = nomeCargo.trim().toUpperCase().match(/CC-\d+/);
    return match ? match[0] : nomeCargo.trim().toUpperCase();
}

// Renderiza a estrutura do gabinete agrupada por CC com disposição triangular
function renderizarEstruturaCCs(estrutura, cargosAtivosCC) {
    const consumivel = [...cargosAtivosCC];

    const slots = estrutura.map(cc => {
        const idx = consumivel.indexOf(cc);
        if (idx !== -1) { consumivel.splice(idx, 1); return { cc, estado: 'ocupado' }; }
        return { cc, estado: 'vago' };
    });

    // Agrupa por CC preservando ordem de primeiro aparecimento
    const ordem = [];
    const grupos = {};
    slots.forEach(s => {
        if (!grupos[s.cc]) { grupos[s.cc] = []; ordem.push(s.cc); }
        grupos[s.cc].push(s.estado);
    });

    // Número de linhas = máximo de repetições de qualquer CC
    const maxSlots = Math.max(...ordem.map(cc => grupos[cc].length));

    // Grid: 1 coluna por tipo de CC, gap uniforme entre todos
    // Cada CC ocupa sempre a mesma coluna, repetições vão para linhas abaixo
    const numCols = ordem.length;
    const templateCols = Array(numCols).fill('102px').join(' ');

    let cellsHTML = '';
    ordem.forEach((cc, colIdx) => {
        const n = grupos[cc].length;
        // Preenche os slots ocupados/vagos
        for (let row = 0; row < n; row++) {
            const estado = grupos[cc][row];
            cellsHTML += `<div class="cc-card ${estado}" style="grid-column:${colIdx+1};grid-row:${row+1}">
                <span class="cc-card-label">${cc}</span>
                <span class="cc-card-status">${estado === 'ocupado' ? '● ocupado' : '○ vago'}</span>
            </div>`;
        }
        // Linhas acima do máximo ficam vazias (não precisam de placeholder — grid-template-rows cuida)
    });

    return `<div class="grade-cc-grid" style="grid-template-columns:${templateCols};grid-template-rows:repeat(${maxSlots},82px)">${cellsHTML}</div>`;
}


function ocultarTudo() {
    document.getElementById('estadoInicial').style.display = 'none';
    document.getElementById('painelVereador').classList.add('escondido');
    document.getElementById('painelEspecial').classList.add('escondido');
    document.getElementById('secaoSugestao').classList.add('escondido');
    document.getElementById('topbarBadges').innerHTML = '';
    const mes = document.getElementById('mesSelect').value;
    const gab = document.getElementById('gabineteSelect').value;
    if (!mes || !gab) document.getElementById('estadoInicial').style.display = 'flex';
}

function renderizarAlertas(idEl, alertas) {
    const el = document.getElementById(idEl);
    el.innerHTML = '';
    const icons = {
        erro:  `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>`,
        aviso: `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`,
        ok:    `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>`,
    };
    alertas.forEach(a => {
        el.innerHTML += `<div class="alerta alerta-${a.tipo}">${icons[a.tipo]}<span>${a.msg}</span></div>`;
    });
}

// "DD/MM/AAAA" → Date
function parseData(str) {
    if (!str || !str.trim()) return null;
    const partes = str.trim().split('/');
    if (partes.length !== 3) return null;
    const [d, m, a] = partes.map(Number);
    if (!d || !m || !a) return null;
    return new Date(a, m - 1, d);
}

// "MM/AAAA" → { inicio, fim }
function intervaloMes(mesAno) {
    const [m, a] = mesAno.split('/').map(Number);
    return { inicio: new Date(a, m - 1, 1), fim: new Date(a, m, 0) };
}

function estaAtivo(admissao, exoneracao, inicio, fim) {
    if (!admissao) return false;
    if (admissao > fim) return false;
    if (exoneracao && exoneracao < inicio) return false;
    return true;
}

function moeda(valor) {
    return valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function parseMoeda(str) {
    // Aceita "12.000,00" e "12000.00"
    const s = str.toString().trim();
    if (s.includes(',')) return parseFloat(s.replace(/\./g, '').replace(',', '.')) || 0;
    return parseFloat(s) || 0;
}

function formatarData(date) {
    if (!date) return '—';
    return date.toLocaleDateString('pt-BR');
}

function setStatus(tipo, msg) {
    const dot = document.querySelector('.status-dot');
    const txt = document.getElementById('statusConexao');
    dot.className = 'status-dot' + (tipo ? ' ' + tipo : '');
    txt.textContent = msg;
}

iniciar();

// ─── EXPORTAÇÃO ─────────────────────────────────────────────────────────────

function dadosParaExport() {
    const e       = _exportEstado;
    const isVer   = e.tipo === 'vereador';
    const titulo  = isVer ? 'Relatório do Gabinete' : 'Relatório de Lotação';
    const secComp = isVer ? 'Composição do Gabinete' : 'Composição da Lotação';

    const servsLimpos = e.servidores.filter(s =>
        !s.exoneradoNoMes && !isCedido(s['Cargo'] || '')
    );

    const linhasServs = servsLimpos.map(s => ({
        'Matrícula': (s['Matrícula'] || s['Matricula'] || '').trim() || '—',
        'Nome':      (s['Nome do Servidor'] || '').trim(),
        'Cargo':     (s['Cargo'] || '').trim(),
        'Salário':   moeda(tabelaCargos[(s['Cargo'] || '').trim()] || 0),
        'Admissão':  s.admissao ? formatarData(s.admissao) : '—',
    }));

    const consumivel = servsLimpos.map(s => extrairCC(s['Cargo'] || ''));
    const linhasComp = e.estrutura.map(cc => {
        const idx = consumivel.indexOf(cc);
        const ocupado = idx !== -1;
        if (ocupado) consumivel.splice(idx, 1);
        return { 'Cargo': cc, 'Status': ocupado ? 'Ocupado' : 'Vago' };
    });

    return { titulo, secComp, linhasServs, linhasComp, e };
}

function exportarCSV() {
    const { linhasServs, linhasComp, e } = dadosParaExport();
    const blob = new Blob(
        [`Mês: ${e.mes}\nLotação: ${e.gab}\n\nServidores\n${Papa.unparse(linhasServs)}\n\nComposição\n${Papa.unparse(linhasComp)}`],
        { type: 'text/csv;charset=utf-8;' }
    );
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${e.gab} - ${e.mes}.csv`.replace(/[/\\?%*:|"<>]/g, '-');
    a.click();
}

function exportarXLSX() {
    const { linhasServs, linhasComp, e } = dadosParaExport();
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(linhasServs), 'Servidores');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(linhasComp),  'Composição');
    XLSX.writeFile(wb, `${e.gab} - ${e.mes}.xlsx`.replace(/[/\\?%*:|"<>]/g, '-'));
}

function exportarPDF() {
    const { jsPDF } = window.jspdf;
    const { titulo, secComp, linhasServs, linhasComp, e } = dadosParaExport();
    const isVer = e.tipo === 'vereador';
    const doc   = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const PW = 210, PH = 297, ML = 18, MR = 18;

    // Logo como PNG embutida
    const LOGO = "data:image/png;base64,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";
    // Dimensões da logo no PDF: proporcional ao original 1920x1080 → cabe em ~80x45mm
    const LW = 72, LH = 40;

    // ── CABEÇALHO ──
    // Fundo branco puro no topo
    doc.setFillColor(255, 255, 255);
    doc.rect(0, 0, PW, 48, 'F');

    // Logo no canto superior direito
    doc.addImage(LOGO, 'PNG', PW - MR - LW, 4, LW, LH);

    // Linha institucional à esquerda
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.setTextColor(40, 40, 40);
    doc.text('CONTROLE ORÇAMENTÁRIO LEGISLATIVO', ML, 12);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.setTextColor(100, 100, 100);
    doc.text(isVer ? 'Relatório do Gabinete' : 'Relatório de Lotação', ML, 18);

    // Linha divisória
    doc.setDrawColor(30, 30, 30);
    doc.setLineWidth(0.6);
    doc.line(ML, 46, PW - MR, 46);
    doc.setLineWidth(0.15);
    doc.setDrawColor(180, 180, 180);
    doc.line(ML, 47.5, PW - MR, 47.5);

    // ── BLOCO IDENTIFICAÇÃO ──
    let y = 56;

    // Nome da lotação — maior destaque
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.setTextColor(15, 15, 15);
    doc.text(e.gab, ML, y);
    y += 7;

    // Responsável (só para especiais)
    if (!isVer && e.responsavel) {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(9);
        doc.setTextColor(80, 80, 80);
        doc.text('Vereador Responsável: ', ML, y);
        const labelW = doc.getTextWidth('Vereador Responsável: ');
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(30, 30, 30);
        doc.text(e.responsavel, ML + labelW, y);
        y += 5;
    }

    // Linha separadora leve
    doc.setDrawColor(210, 210, 210);
    doc.setLineWidth(0.2);
    doc.line(ML, y + 1, PW - MR, y + 1);
    y += 7;

    // ── TABELA SERVIDORES ──
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.setTextColor(30, 30, 30);
    doc.text('SERVIDORES LOTADOS', ML, y);
    y += 3;

    doc.autoTable({
        startY: y,
        margin: { left: ML, right: MR },
        head: [['Matrícula', 'Nome do Servidor', 'Cargo', 'Salário', 'Admissão']],
        body: linhasServs.map(r => [r['Matrícula'], r['Nome'], r['Cargo'], r['Salário'], r['Admissão']]),
        styles: {
            font: 'helvetica', fontSize: 8, cellPadding: 3,
            textColor: [25, 25, 25], lineColor: [200, 200, 200], lineWidth: 0.15,
        },
        headStyles: {
            fillColor: [35, 35, 35], textColor: [255, 255, 255],
            fontStyle: 'bold', fontSize: 7.5,
        },
        alternateRowStyles: { fillColor: [245, 245, 245] },
        columnStyles: {
            0: { cellWidth: 22 },
            1: { cellWidth: 'auto' },
            2: { cellWidth: 30 },
            3: { cellWidth: 26, halign: 'right' },
            4: { cellWidth: 22 },
        },
    });

    y = doc.lastAutoTable.finalY + 10;

    // ── COMPOSIÇÃO ──
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.setTextColor(30, 30, 30);
    doc.text(secComp.toUpperCase(), ML, y);
    y += 3;

    doc.autoTable({
        startY: y,
        margin: { left: ML, right: MR },
        tableWidth: 80,
        head: [['Cargo (CC)', 'Status']],
        body: linhasComp.map(r => [r['Cargo'], r['Status']]),
        styles: {
            font: 'helvetica', fontSize: 8, cellPadding: 3,
            textColor: [25, 25, 25], lineColor: [200, 200, 200], lineWidth: 0.15,
        },
        headStyles: {
            fillColor: [35, 35, 35], textColor: [255, 255, 255],
            fontStyle: 'bold', fontSize: 7.5,
        },
        alternateRowStyles: { fillColor: [245, 245, 245] },
        columnStyles: { 0: { cellWidth: 40 }, 1: { cellWidth: 38 } },
        didParseCell(data) {
            if (data.column.index === 1 && data.section === 'body') {
                data.cell.styles.textColor = data.cell.raw === 'Ocupado' ? [30, 110, 60] : [140, 100, 10];
                data.cell.styles.fontStyle = 'bold';
            }
        },
    });

    // ── RODAPÉ ──
    const today  = new Date();
    const dataStr = today.toLocaleDateString('pt-BR');
    const total  = doc.internal.getNumberOfPages();
    const FLW = 44, FLH = 25; // logo menor no rodapé

    for (let i = 1; i <= total; i++) {
        doc.setPage(i);

        // Linha dupla antes do rodapé
        doc.setDrawColor(30, 30, 30);
        doc.setLineWidth(0.6);
        doc.line(ML, PH - 22, PW - MR, PH - 22);
        doc.setLineWidth(0.15);
        doc.setDrawColor(180, 180, 180);
        doc.line(ML, PH - 21, PW - MR, PH - 21);

        // Logo pequena à direita do rodapé
        doc.addImage(LOGO, 'PNG', PW - MR - FLW, PH - 19, FLW, FLH);

        // Texto do rodapé à esquerda
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(7);
        doc.setTextColor(80, 80, 80);
        doc.text('Câmara Municipal de Curitiba', ML, PH - 14);
        doc.text('Controle Orçamentário Legislativo', ML, PH - 10);
        doc.setTextColor(130, 130, 130);
        doc.text(`Gerado em ${dataStr}   ·   Página ${i} de ${total}`, ML, PH - 6);
    }

    // Nome do arquivo: lotação + data AAAA-MM-DD
    const ymd = today.getFullYear() + '-' + String(today.getMonth()+1).padStart(2,'0') + '-' + String(today.getDate()).padStart(2,'0');
    doc.save(`${e.gab} - ${ymd}.pdf`.replace(/[/\?%*:|"<>]/g, '-'));
}
