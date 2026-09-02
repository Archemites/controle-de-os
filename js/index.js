// Colunas exatas da planilha original (na ordem). "Cobrança Medição" foi removida:
// não é relevante para o time e não é lida nem gravada pelo app.
const COLUMNS = [
  "O.S", "Mês Envio Medição", "Vistoria Projeto", "Ano", "GAP", "Envio Medição",
  "Sla Envio Medição", "Valor Medição Enviadas", "Engenheiro Acesso", "Projetista",
  "Regional", "Pedido de Compra", "Valor Faturado", "Medição Terceiros",
  "Custos Diversos", "Custo ART", "Custo Total"
];

const CURRENCY_COLS = new Set([
  "Valor Medição Enviadas", "Valor Faturado", "Medição Terceiros", "Custos Diversos", "Custo ART", "Custo Total"
]);
const DATE_COLS = new Set(["Vistoria Projeto", "Envio Medição"]);
// Colunas de texto/código normalizadas para MAIÚSCULO em todo cadastro e
// importação — datas e valores numéricos ficam de fora.
const COLS_MAIUSCULAS = new Set([
  "O.S", "Mês Envio Medição", "GAP", "Engenheiro Acesso", "Projetista", "Regional", "Pedido de Compra"
]);

let workbook = null;
let sheetName = null;
let headerRowIndex = 0; // 0-based row index of the header line inside the sheet
let lastOriginalRow = 0; // 0-based index of the last row that already has data in the original sheet
let rows = []; // array of row objects keyed by COLUMNS (only ORIGINAL rows, for display)
let newRows = []; // rows added by the user this session, not yet written into the sheet
// __linha das linhas do arquivo original que estão na lixeira (persistido no
// autosave). Impede que a linha reapareça quando o .xlsx original é relido.
let linhasExcluidas = new Set();
// Lixeira: registros excluídos ficam aqui por 30 dias antes do descarte definitivo.
// Cada item é { linha: <objeto da medição>, excluidoEm: <timestamp ms> }.
let lixeira = [];
const LIXEIRA_DIAS = 30;
const LIXEIRA_MS = LIXEIRA_DIAS * 86400000;
let activeTab = 'comPedido'; // 'comPedido' | 'semPedido' | 'aCobrar' | 'envios' | 'lixeira'
// Nome de cada aba, usado no rótulo do botão de download e no nome do arquivo exportado.
const NOMES_ABAS = {
  comPedido: 'Com Pedido (Faturados)',
  semPedido: 'Sem Pedido (Não Faturados)',
  envios: 'Envios',
  aCobrar: 'A Cobrar',
  lixeira: 'Lixeira'
};
const SLA_A_COBRAR_DIAS = 10;
let editingRow = null; // referência à linha (objeto em `rows`) sendo editada, ou null se cadastrando nova
let loadedFileName = null;

// --- Autosave (IndexedDB) ---------------------------------------------------
// Guarda os bytes originais do .xlsx aberto + as linhas novas/editadas desta
// sessão, para que um refresh acidental da página não perca o trabalho feito.
// Ao carregar a página, se houver um autosave salvo, ele é restaurado sozinho.
const AUTOSAVE_DB = 'medicoes_autosave';
const AUTOSAVE_STORE = 'estado';
const AUTOSAVE_KEY = 'atual';

function abrirAutosaveDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(AUTOSAVE_DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(AUTOSAVE_STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

let autosaveTimer = null;
function agendarAutosave() {
  clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(salvarAutosave, 400);
}

async function salvarAutosave() {
  if (!workbook || !originalFileBytes) return;
  try {
    const db = await abrirAutosaveDB();
    const estado = {
      fileName: loadedFileName,
      fileBytes: originalFileBytes,
      newRows,
      lixeira,
      linhasExcluidas: [...linhasExcluidas],
      activeTab,
      savedAt: Date.now()
    };
    await new Promise((resolve, reject) => {
      const tx = db.transaction(AUTOSAVE_STORE, 'readwrite');
      tx.objectStore(AUTOSAVE_STORE).put(estado, AUTOSAVE_KEY);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.error('Falha ao salvar autosave:', err);
  }
}

async function carregarAutosave() {
  try {
    const db = await abrirAutosaveDB();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(AUTOSAVE_STORE, 'readonly');
      const req = tx.objectStore(AUTOSAVE_STORE).get(AUTOSAVE_KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  } catch (err) {
    console.error('Falha ao ler autosave:', err);
    return null;
  }
}

async function limparAutosave() {
  try {
    const db = await abrirAutosaveDB();
    const tx = db.transaction(AUTOSAVE_STORE, 'readwrite');
    tx.objectStore(AUTOSAVE_STORE).delete(AUTOSAVE_KEY);
  } catch (err) {
    console.error('Falha ao limpar autosave:', err);
  }
}

let originalFileBytes = null; // Uint8Array do arquivo .xlsx original tal como aberto pelo usuário

const fileInput = document.getElementById('fileInput');
const openBtn = document.getElementById('openBtn');
const fileStatus = document.getElementById('fileStatus');
const cadastrarBtn = document.getElementById('cadastrarBtn');
const formModalOverlay = document.getElementById('formModalOverlay');
const formPanel = document.getElementById('formPanel');
const formPanelTitle = document.getElementById('formPanelTitle');
const cancelBtn = document.getElementById('cancelBtn');

formModalOverlay.addEventListener('click', (e) => {
  if (e.target === formModalOverlay) closeForm();
});
const saveRowBtn = document.getElementById('saveRowBtn');
const formMsg = document.getElementById('formMsg');
const downloadBtn = document.getElementById('downloadBtn');
const emptyState = document.getElementById('emptyState');
const dataTable = document.getElementById('dataTable');
const totalsRow = document.getElementById('totalsRow');
const headerRow = document.getElementById('headerRow');
const tableBody = document.getElementById('tableBody');
const rowCount = document.getElementById('rowCount');
const statusCount = document.getElementById('statusCount');
const subTabs = document.getElementById('subTabs');
const tabComPedido = document.getElementById('tabComPedido');
const tabSemPedido = document.getElementById('tabSemPedido');
const tabACobrar = document.getElementById('tabACobrar');
const tabEnvios = document.getElementById('tabEnvios');
const tabLixeira = document.getElementById('tabLixeira');
const lixeiraCount = document.getElementById('lixeiraCount');
const aCobrarBadge = document.getElementById('aCobrarBadge');
const comPedidoCount = document.getElementById('comPedidoCount');
const semPedidoCount = document.getElementById('semPedidoCount');
const enviosCount = document.getElementById('enviosCount');
const searchBar = document.getElementById('searchBar');
const searchInput = document.getElementById('searchInput');

let searchTerm = '';
let searchFields = new Set(['O.S', 'GAP', 'Pedido de Compra']);
const filterChips = document.querySelectorAll('.filter-chip');

// --- Tema claro/escuro (persistido no navegador) ----------------------------
const THEME_KEY = 'medicoes_tema';
const themeToggle = document.getElementById('themeToggle');

function aplicarTema(tema) {
  const dark = tema === 'dark';
  document.body.classList.toggle('dark', dark);
  themeToggle.title = dark ? 'Mudar para tema claro' : 'Mudar para tema escuro';
}

aplicarTema(localStorage.getItem(THEME_KEY) || 'light');

themeToggle.addEventListener('click', () => {
  const novo = document.body.classList.contains('dark') ? 'light' : 'dark';
  localStorage.setItem(THEME_KEY, novo);
  aplicarTema(novo);
});

// --- Colunas exibidas (persistido no navegador) -----------------------------
const COLUMNS_VISIBILITY_KEY = 'medicoes_colunas_visiveis';
const columnsPanel = document.getElementById('columnsPanel');
const columnsList = document.getElementById('columnsList');
let visibleColumns = new Set(COLUMNS);

function carregarVisibilidadeColunas() {
  try {
    const saved = JSON.parse(localStorage.getItem(COLUMNS_VISIBILITY_KEY));
    if (Array.isArray(saved)) {
      visibleColumns = new Set(saved.filter(c => COLUMNS.includes(c)));
    }
  } catch (e) {}
}

function salvarVisibilidadeColunas() {
  localStorage.setItem(COLUMNS_VISIBILITY_KEY, JSON.stringify([...visibleColumns]));
}

// A exportação só inclui as colunas marcadas no filtro; sem nenhuma marcada
// não há o que exportar, então o botão de download fica desabilitado.
function atualizarEstadoDownload() {
  const nenhumaColuna = visibleColumns.size === 0;
  // A lixeira não é exportável: são registros fora da lista, aguardando descarte.
  const naLixeira = activeTab === 'lixeira';
  downloadBtn.disabled = !workbook || nenhumaColuna || naLixeira;
  const nomeAba = NOMES_ABAS[activeTab] || '';
  downloadBtn.textContent = `Baixar planilha — ${nomeAba}`;
  downloadBtn.title = naLixeira
    ? 'A aba "Lixeira" não é exportada — restaure os registros para incluí-los em uma planilha'
    : nenhumaColuna
      ? 'Selecione ao menos uma coluna no filtro de colunas para baixar a planilha'
      : `Baixa somente as linhas exibidas na aba "${nomeAba}", com as colunas marcadas no filtro`;
}

function renderColumnsPanel() {
  columnsList.innerHTML = COLUMNS.map((col, i) => `
    <div class="column-toggle">
      <input type="checkbox" id="colToggle_${i}" data-col="${col}" ${visibleColumns.has(col) ? 'checked' : ''}>
      <label for="colToggle_${i}">${col}</label>
    </div>
  `).join('');
  columnsList.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.addEventListener('change', () => {
      if (cb.checked) visibleColumns.add(cb.dataset.col);
      else visibleColumns.delete(cb.dataset.col);
      salvarVisibilidadeColunas();
      atualizarEstadoDownload();
      renderTable();
    });
  });
}

document.getElementById('columnsShowAll').addEventListener('click', () => {
  visibleColumns = new Set(COLUMNS);
  salvarVisibilidadeColunas();
  renderColumnsPanel();
  atualizarEstadoDownload();
  renderTable();
});
document.getElementById('columnsHideAll').addEventListener('click', () => {
  visibleColumns = new Set();
  salvarVisibilidadeColunas();
  renderColumnsPanel();
  atualizarEstadoDownload();
  renderTable();
});

carregarVisibilidadeColunas();
renderColumnsPanel();
atualizarEstadoDownload();

const columnsPanelWrap = document.getElementById('columnsPanelWrap');

searchInput.addEventListener('input', () => {
  searchTerm = searchInput.value.trim().toLowerCase();
  renderTable();
});

filterChips.forEach(chip => {
  chip.addEventListener('click', () => {
    const field = chip.dataset.field;
    if (searchFields.has(field)) {
      // Não permite desativar o último filtro ativo, para sempre haver ao menos um campo pesquisável.
      if (searchFields.size === 1) return;
      searchFields.delete(field);
      chip.classList.remove('active');
    } else {
      searchFields.add(field);
      chip.classList.add('active');
    }
    renderTable();
  });
});

function hasPedido(r) {
  const v = r['Pedido de Compra'];
  return v !== null && v !== undefined && String(v).trim() !== '';
}

function isACobrar(r) {
  if (hasPedido(r)) return false;
  const dias = slaAtual(r);
  return typeof dias === 'number' && dias > SLA_A_COBRAR_DIAS;
}

function setActiveTab(tab) {
  activeTab = tab;
  tabComPedido.classList.toggle('active', tab === 'comPedido');
  tabSemPedido.classList.toggle('active', tab === 'semPedido');
  tabACobrar.classList.toggle('active', tab === 'aCobrar');
  tabEnvios.classList.toggle('active', tab === 'envios');
  tabLixeira.classList.toggle('active', tab === 'lixeira');
  atualizarEstadoDownload(); // o rótulo do botão traz o nome da aba ativa
  renderTable();
  agendarAutosave();
}

tabComPedido.addEventListener('click', () => setActiveTab('comPedido'));
tabSemPedido.addEventListener('click', () => setActiveTab('semPedido'));
tabACobrar.addEventListener('click', () => setActiveTab('aCobrar'));
tabEnvios.addEventListener('click', () => setActiveTab('envios'));
tabLixeira.addEventListener('click', () => setActiveTab('lixeira'));

openBtn.addEventListener('click', () => fileInput.click());
cadastrarBtn.addEventListener('click', openForm);
cancelBtn.addEventListener('click', closeForm);
saveRowBtn.addEventListener('click', salvarMedicao);
downloadBtn.addEventListener('click', baixarPlanilha);

document.getElementById('f_envio').addEventListener('change', atualizarSla);
document.getElementById('f_custosDiversos').addEventListener('input', atualizarCustoTotal);
document.getElementById('f_custoArt').addEventListener('input', atualizarCustoTotal);

// Todos os campos do formulário (exceto os somente-leitura, calculados automaticamente)
// recebem destaque visual quando vazios, para o usuário identificar rapidamente o que falta.
const camposComDestaqueVazio = Array.from(document.querySelectorAll('#formPanel .field input:not([readonly]), #formPanel .field select'))
  .map(el => el.id);
camposComDestaqueVazio.forEach(id => {
  const el = document.getElementById(id);
  if (el) {
    el.addEventListener('input', () => {
      el.classList.toggle('empty-required', el.value === '');
    });
    el.addEventListener('change', () => {
      el.classList.toggle('empty-required', el.value === '');
    });
  }
});

// "O.S" e "Pedido de Compra" são zero-padded na planilha original via number_format
// customizado (ex: "0000000000" faz 21214728 exibir como "0021214728"). O SheetJS
// (sheet_to_json com raw:true) devolve o valor numérico puro (v), descartando o
// padding â€” por isso lemos essas duas colunas célula a célula, priorizando o texto
// formatado (w) quando a célula tem esse tipo de máscara, com fallback pro raw.
const COLS_ZERO_PAD = new Set(["O.S", "Pedido de Compra"]);

function valorCelulaComPadding(ws, rowIdx, colIdx, col, rawVal) {
  if (!COLS_ZERO_PAD.has(col)) return rawVal;
  const cell = ws[XLSX.utils.encode_cell({ r: rowIdx, c: colIdx })];
  if (cell && cell.t === 'n' && typeof cell.w === 'string' && /^0\d/.test(cell.w)) {
    return cell.w;
  }
  return rawVal;
}

// Registro sem nenhum campo preenchido (só nulos, vazios ou espaços) — OU sem
// nenhum dos campos que realmente identificam uma medição (O.S, Envio Medição,
// Pedido de Compra, Valor Medição Enviadas, Valor Faturado). Sobra na planilha
// original um rastro de milhares de linhas de fórmula "fantasma" que só carregam
// lixo em "Ano"/"Sla Envio Medição" (ex.: Ano=2026, Sla=46231, resto tudo nulo) —
// tecnicamente não são 100% vazias, mas não representam registro nenhum. Ambos os
// casos são sempre descartados — na importação, no cadastro/edição e antes de exportar.
function linhaVazia(row) {
  if (!row) return true;
  if (COLUMNS.every(col => {
    const v = row[col];
    return v === null || v === undefined || String(v).trim() === '';
  })) return true;
  const semIdentidade = ['O.S', 'Envio Medição', 'Pedido de Compra', 'Valor Medição Enviadas', 'Valor Faturado']
    .every(col => {
      const v = row[col];
      return v === null || v === undefined || String(v).trim() === '';
    });
  return semIdentidade;
}

// Constrói os objetos de linha (chaveados por COLUMNS) a partir de uma matriz já
// extraída via sheet_to_json, reaplicando o padding de zeros perdido no raw:true.
// Toda O.S numérica é gravada com 10 dígitos, zeros à esquerda inclusos
// ("21214728" -> "0021214728"): é assim que ela existe na planilha de origem
// (921 de 921 O.S numéricas) e é o que o time espera ver na tela e no export.
// Valores que não são só dígitos (ex: "GAP-36613", "0021214728/0021214746")
// passam intactos — completá-los com zeros só destruiria a informação.
function normalizarOs(v) {
  if (v === null || v === undefined) return v;
  const s = String(v).trim();
  return /^\d+$/.test(s) ? s.padStart(10, '0') : v;
}

function construirLinhas(ws, json, headerIdx) {
  const linhas = [];
  let ultimaLinha = headerIdx;
  for (let i = headerIdx + 1; i < json.length; i++) {
    const r = json[i];
    if (!r || r.every(c => c === null || c === undefined || c === '')) continue;
    const obj = {};
    COLUMNS.forEach((col, idx) => {
      const val = r[idx] !== undefined ? r[idx] : null;
      const lida = valorCelulaComPadding(ws, i, idx, col, val);
      const norm = (COLS_MAIUSCULAS.has(col) && typeof lida === 'string') ? lida.toUpperCase() : lida;
      obj[col] = col === 'O.S' ? normalizarOs(norm) : norm;
    });
    // A linha pode ter conteúdo só em colunas fora de COLUMNS (ex: "Cobrança
    // Medição"), o que a deixaria vazia como registro — nesse caso, descarta.
    if (linhaVazia(obj)) continue;
    // Índice da linha na planilha de origem: é a identidade estável usada para
    // reconciliar o autosave (uma linha original editada volta como "extra" e
    // precisa SUBSTITUIR a original, não ser acrescentada de novo).
    obj.__linha = i;
    linhas.push(obj);
    ultimaLinha = i;
  }
  return { linhas, ultimaLinha };
}

function carregarWorkbookDeBytes(data, fileName, rowsExtras, excluidas, lixeiraSalva) {
  workbook = XLSX.read(data, { type: 'array', cellDates: true });
  sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];

  // A build community do SheetJS não preserva de forma confiável as larguras de
  // coluna (!cols) nem o number_format customizado de "O.S" (zero-padded) no ciclo
  // read->write. Reforçamos manualmente aqui, logo após a leitura, com os valores
  // exatos do arquivo modelo, para o download sair idêntico ao original.
  ws['!cols'] = [
    { wch: 23.608073 }, { wch: 22.166667 }, { wch: 19.053385 }, { wch: 9.053385 },
    { wch: 19.498698 }, { wch: 18.166667 }, { wch: 21.276042 }, { wch: 26.608073 },
    { wch: 22.053385 }, { wch: 19.608073 }, { wch: 13.166667 }, { wch: 21.276042 },
    { wch: 18.608073 }, { wch: 21.608073 }, { wch: 19.385417 }, { wch: 14.166667 },
    { wch: 15.276042 }, { wch: 22.053385 }
  ];

  const json = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  // Find header row: the row that contains "O.S" as first cell
  headerRowIndex = json.findIndex(r => r && r[0] === 'O.S');
  if (headerRowIndex === -1) headerRowIndex = 1; // fallback matches known layout

  newRows = [];
  // Exclusões restauradas do autosave: as linhas do arquivo original que o usuário
  // apagou não podem reaparecer só porque o .xlsx é relido do zero.
  linhasExcluidas = new Set(Array.isArray(excluidas) ? excluidas : []);
  lixeira = Array.isArray(lixeiraSalva) ? lixeiraSalva.filter(i => i && i.linha) : [];
  // Itens vencidos (mais de LIXEIRA_DIAS na lixeira) são descartados aqui. Os que
  // vieram do arquivo continuam em `linhasExcluidas` — seguem fora da lista, só
  // deixam de ser recuperáveis.
  purgarLixeira();
  const extraidas = construirLinhas(ws, json, headerRowIndex);
  rows = linhasExcluidas.size
    ? extraidas.linhas.filter(r => !linhasExcluidas.has(r.__linha))
    : extraidas.linhas;
  lastOriginalRow = extraidas.ultimaLinha;

  // Reaplica linhas novas/editadas restauradas do autosave. As que vieram de uma
  // linha do arquivo (têm __linha) substituem a original recém-lida — só as
  // cadastradas nesta sessão é que são acrescentadas, senão a edição de uma
  // medição existente viraria uma duplicata a cada recarregamento.
  if (rowsExtras && rowsExtras.length) {
    rowsExtras.forEach(r => {
      if (r && r.__linha !== undefined && linhasExcluidas.has(r.__linha)) return;
      const pos = r && r.__linha !== undefined ? rows.findIndex(o => o.__linha === r.__linha) : -1;
      if (pos !== -1) rows[pos] = r;
      else rows.push(r);
      newRows.push(r);
    });
  }

  loadedFileName = fileName;
  originalFileBytes = data;

  // Toda importação de planilha base reexibe todas as colunas, mesmo que o
  // usuário tenha desmarcado algumas antes.
  visibleColumns = new Set(COLUMNS);
  salvarVisibilidadeColunas();
  renderColumnsPanel();

  activeTab = 'comPedido';
  tabComPedido.classList.add('active');
  tabSemPedido.classList.remove('active');
  tabEnvios.classList.remove('active');

  cadastrarBtn.disabled = false;
  atualizarEstadoDownload();
  renderTable();

  // Conferência automática contra o Controle de Pedidos de Compra. Em setTimeout
  // para o chamador terminar de escrever a mensagem em `fileStatus` antes de a
  // sincronização acrescentar a dela.
  setTimeout(() => sincronizarComPedidos(), 0);
}

// Lê linhas de dados de um arquivo .xlsx sem torná-lo o workbook ativo,
// usado no modo "adicionar à lista existente" da importação.
function extrairLinhasDeBytes(data) {
  const wb = XLSX.read(data, { type: 'array', cellDates: true });
  const sn = wb.SheetNames[0];
  const ws = wb.Sheets[sn];
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  let hIdx = json.findIndex(r => r && r[0] === 'O.S');
  if (hIdx === -1) hIdx = 1;
  return construirLinhas(ws, json, hIdx).linhas;
}

let arquivoImportPendente = null;

// Colunas consideradas na comparação de duplicidade "por campos em comum"
// (ex.: mesma O.S. ou mesmo Pedido de Compra já são fortes indícios; exclui
// "Custo Total" por ser derivado automaticamente de outros campos).
const CAMPOS_COMPARACAO_DUP = COLUMNS.filter(c => c !== 'Custo Total');

function valorComparavel(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return v.getTime();
  return String(v).trim().toLowerCase();
}

// Conta quantos campos têm valores iguais (e não vazios) entre duas linhas.
function contarCamposIguais(a, b) {
  let count = 0;
  for (const col of CAMPOS_COMPARACAO_DUP) {
    const va = valorComparavel(a[col]);
    const vb = valorComparavel(b[col]);
    if (va !== null && vb !== null && va === vb) count++;
  }
  return count;
}

// Um registro é tratado como duplicado do outro se compartilharem 2 ou mais
// campos iguais (ex.: mesmo Pedido de Compra E mesma Regional), não só a O.S.
const MIN_CAMPOS_IGUAIS_DUP = 2;

function mesmaOs(a, b) {
  const oa = valorComparavel(a['O.S']);
  const ob = valorComparavel(b['O.S']);
  return oa !== null && ob !== null && oa === ob;
}

// Checa a O.S primeiro (é a identidade de fato da medição em todo o resto do
// app — hasPedido/linhasPelaOs/sincronizarComPedidos etc. giram em torno
// dela), só caindo pro heurístico de "N campos em comum" quando a O.S não bate
// com nada. Sem essa prioridade, num re-import do arquivo inteiro 932 de 933
// linhas casavam com a linha ERRADA: em ~1000 registros, "Mês Envio Medição" +
// "Ano" + "Regional" + os dois "R$ -" de Custos Diversos/Custo ART são valores
// tão comuns que qualquer par de linhas atinge o mínimo de 2 campos iguais por
// pura coincidência, elegendo a primeira linha da planilha como "duplicata" de
// quase tudo (ver TESTE_DE_FOGO.md).
function encontrarLinhaDuplicada(nova, listaRows) {
  for (let idx = 0; idx < listaRows.length; idx++) {
    if (mesmaOs(nova, listaRows[idx])) return idx;
  }
  for (let idx = 0; idx < listaRows.length; idx++) {
    if (contarCamposIguais(nova, listaRows[idx]) >= MIN_CAMPOS_IGUAIS_DUP) return idx;
  }
  return -1;
}

function processarImportacao(file, data, modo) {
  try {
    if (modo === 'append' && workbook) {
      const extraidas = extrairLinhasDeBytes(data);

      const semDuplicata = [];
      const fila = [];
      extraidas.forEach(nova => {
        const idxExistente = encontrarLinhaDuplicada(nova, rows);
        if (idxExistente !== -1) {
          fila.push({ nova, idxExistente });
        } else {
          const idxNaSemDuplicata = encontrarLinhaDuplicada(nova, semDuplicata);
          if (idxNaSemDuplicata !== -1) {
            fila.push({ nova, idxExistente: rows.length + idxNaSemDuplicata });
          } else {
            semDuplicata.push(nova);
          }
        }
      });

      semDuplicata.forEach(r => {
        // Vieram de OUTRO arquivo: o __linha delas não corresponde à planilha
        // ativa e confundiria a reconciliação do autosave.
        delete r.__linha;
        rows.push(r);
        newRows.push(r);
      });

      if (fila.length > 0) {
        iniciarRevisaoDuplicados(fila, file.name, semDuplicata.length);
      } else {
        renderTable();
        fileStatus.textContent = `Adicionadas ${extraidas.length} medições de "${file.name}" à lista existente.`;
        salvarAutosave();
      }
    } else {
      carregarWorkbookDeBytes(data, file.name, null);
      fileStatus.textContent = `Carregado: "${file.name}" â€” ${rows.length} medições encontradas.`;
      salvarAutosave();
    }
  } catch (err) {
    fileStatus.textContent = 'Erro ao ler o arquivo: ' + err.message;
  }
}

// --- Revisão de duplicados (um a um) ---
const dupModalOverlay = document.getElementById('dupModalOverlay');
const dupPanelSummary = document.getElementById('dupPanelSummary');
const dupItemInfo = document.getElementById('dupItemInfo');
const dupSubstituirBtn = document.getElementById('dupSubstituirBtn');
const dupAdicionarBtn = document.getElementById('dupAdicionarBtn');
const dupIgnorarBtn = document.getElementById('dupIgnorarBtn');
const dupAplicarTodos = document.getElementById('dupAplicarTodos');

let filaDuplicados = [];
let filaIndice = 0;
let filaTotalSemDuplicata = 0;
let filaNomeArquivo = '';
let filaAdicionadasCount = 0;
let filaAcaoForcada = null; // quando "aplicar a todos" é marcado

const CAMPOS_RESUMO_DUP = ["O.S", "Pedido de Compra", "GAP", "Envio Medição", "Valor Medição Enviadas", "Engenheiro Acesso"];

function camposIguaisEntre(a, b) {
  return CAMPOS_COMPARACAO_DUP.filter(col => {
    const va = valorComparavel(a[col]);
    const vb = valorComparavel(b[col]);
    return va !== null && vb !== null && va === vb;
  });
}

function iniciarRevisaoDuplicados(fila, nomeArquivo, totalSemDuplicata) {
  filaDuplicados = fila;
  filaIndice = 0;
  filaTotalSemDuplicata = totalSemDuplicata;
  filaNomeArquivo = nomeArquivo;
  filaAdicionadasCount = totalSemDuplicata;
  filaAcaoForcada = null;
  dupAplicarTodos.checked = false;
  mostrarProximoDuplicado();
}

function renderCampoDup(row) {
  return CAMPOS_RESUMO_DUP.map(col => {
    let val = row[col];
    if (DATE_COLS.has(col)) val = fmtDate(val);
    else if (CURRENCY_COLS.has(col)) val = fmtCurrency(val);
    return `<div><b>${col}:</b> ${val === null || val === undefined || val === '' ? 'â€”' : escapeHtml(val)}</div>`;
  }).join('');
}

function mostrarProximoDuplicado() {
  // "Aplicar a todos": resolve o resto da fila num laço, não em recursão. Um
  // arquivo com muitos duplicados (milhares — plausível ao consolidar
  // planilhas históricas) estourava a pilha de chamadas, já que
  // mostrarProximoDuplicado()/aplicarAcaoDuplicado() se chamavam um ao outro a
  // cada item (ver TESTE_DE_FOGO.md).
  if (filaAcaoForcada) {
    while (filaIndice < filaDuplicados.length) aplicarAcaoDuplicado(filaAcaoForcada);
  }
  if (filaIndice >= filaDuplicados.length) {
    finalizarRevisaoDuplicados();
    return;
  }
  const { nova, idxExistente } = filaDuplicados[filaIndice];
  const existente = rows[idxExistente];
  const camposIguais = camposIguaisEntre(nova, existente);
  dupPanelSummary.textContent = `Duplicado ${filaIndice + 1} de ${filaDuplicados.length} (mesmo(s) valor(es) em: ${camposIguais.join(', ')}). O que deseja fazer com este item?`;
  dupItemInfo.innerHTML = `<b>Existente na lista:</b>${renderCampoDup(existente)}<br><b>Novo, da planilha importada:</b>${renderCampoDup(nova)}`;
  dupModalOverlay.classList.add('open');
}

function aplicarAcaoDuplicado(acao) {
  const { nova, idxExistente } = filaDuplicados[filaIndice];
  if (acao === 'substituir') {
    const antiga = rows[idxExistente];
    // Herda a identidade da linha substituída (ver __linha em construirLinhas);
    // sem isso o autosave reinseriria a original ao lado da substituta.
    if (antiga && antiga.__linha !== undefined) nova.__linha = antiga.__linha;
    else delete nova.__linha;
    rows[idxExistente] = nova;
    const posNewRows = newRows.indexOf(antiga);
    if (posNewRows !== -1) newRows[posNewRows] = nova;
    else newRows.push(nova);
    filaAdicionadasCount++;
  } else if (acao === 'adicionar') {
    delete nova.__linha; // veio de outro arquivo — não é uma linha da planilha ativa
    rows.push(nova);
    newRows.push(nova);
    filaAdicionadasCount++;
  }
  // 'ignorar': não faz nada com o item novo.
  filaIndice++;
}

function finalizarRevisaoDuplicados() {
  dupModalOverlay.classList.remove('open');
  renderTable();
  fileStatus.textContent = `Adicionadas ${filaAdicionadasCount} medições de "${filaNomeArquivo}" à lista existente (${filaDuplicados.length} duplicado(s) revisado(s)).`;
  salvarAutosave();
  filaDuplicados = [];
  filaIndice = 0;
  filaAcaoForcada = null;
}

function escolherAcaoDuplicado(acao) {
  if (dupAplicarTodos.checked) filaAcaoForcada = acao;
  aplicarAcaoDuplicado(acao);
  mostrarProximoDuplicado();
}

dupSubstituirBtn.addEventListener('click', () => escolherAcaoDuplicado('substituir'));
dupAdicionarBtn.addEventListener('click', () => escolherAcaoDuplicado('adicionar'));
dupIgnorarBtn.addEventListener('click', () => escolherAcaoDuplicado('ignorar'));

const importModalOverlay = document.getElementById('importModalOverlay');
const importReplaceBtn = document.getElementById('importReplaceBtn');
const importAppendBtn = document.getElementById('importAppendBtn');
const importCancelBtn = document.getElementById('importCancelBtn');

function fecharImportModal() {
  importModalOverlay.classList.remove('open');
  arquivoImportPendente = null;
}

importModalOverlay.addEventListener('click', (e) => {
  if (e.target === importModalOverlay) fecharImportModal();
});
importCancelBtn.addEventListener('click', fecharImportModal);
importReplaceBtn.addEventListener('click', () => {
  if (!arquivoImportPendente) return;
  const { file, data } = arquivoImportPendente;
  fecharImportModal();
  processarImportacao(file, data, 'replace');
});
importAppendBtn.addEventListener('click', () => {
  if (!arquivoImportPendente) return;
  const { file, data } = arquivoImportPendente;
  fecharImportModal();
  processarImportacao(file, data, 'append');
});

fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    if (workbook && rows.length > 0) {
      arquivoImportPendente = { file, data };
      importModalOverlay.classList.add('open');
    } else {
      processarImportacao(file, data, 'replace');
    }
  };
  reader.readAsArrayBuffer(file);
  fileInput.value = '';
});

(async function restaurarAutosaveNaAbertura() {
  const estado = await carregarAutosave();
  if (!estado || !estado.fileBytes) return;
  try {
    carregarWorkbookDeBytes(estado.fileBytes, estado.fileName, estado.newRows || [], estado.linhasExcluidas || [], estado.lixeira || []);
    if (estado.activeTab) setActiveTab(estado.activeTab);
    const dataSalva = new Date(estado.savedAt).toLocaleString('pt-BR');
    fileStatus.textContent = `Sessão restaurada automaticamente: "${estado.fileName}" â€” ${rows.length} medições (autosave de ${dataSalva}).`;
  } catch (err) {
    console.error('Falha ao restaurar autosave:', err);
  }
})();

function toDateInputValue(v) {
  if (!v) return '';
  const d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return '';
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

function marcarCamposVazios() {
  camposComDestaqueVazio.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      const vazio = el.value === '' || el.value === null;
      el.classList.toggle('empty-required', vazio);
    }
  });
}

// ---------------------------------------------------------------------------
// Leitura de PDF
// Extrai o texto do PDF via pdf.js (100% no navegador, nenhum upload sai da
// máquina) e procura, por regex, os campos que aparecem nas ordens/notas:
// O.S, GAP, Pedido de Compra (= Nº do Pedido), Valor Faturado (= Valor Líquido),
// Data e Cliente.
// O resultado é só uma leitura assistida — nada é gravado sem o usuário mandar;
// o botão "Usar no cadastro de medição" pré-preenche o formulário de cadastro.
// ---------------------------------------------------------------------------
const pdfBtn = document.getElementById('pdfBtn');
const pdfInput = document.getElementById('pdfInput');
const pdfModalOverlay = document.getElementById('pdfModalOverlay');
const pdfPanelText = document.getElementById('pdfPanelText');
const pdfResultados = document.getElementById('pdfResultados');
const pdfUsarBtn = document.getElementById('pdfUsarBtn');
const pdfOutroBtn = document.getElementById('pdfOutroBtn');
const pdfFecharBtn = document.getElementById('pdfFecharBtn');

let pdfCamposLidos = null; // último resultado da extração, ou null
let pdfNomeArquivo = ''; // nome do PDF lido, citado nas mensagens/observação

// O "Valor Líquido" do pedido não vem como "rótulo: valor" — é a ÚLTIMA coluna
// da tabela de itens, cujo cabeçalho é "... Preço Líquido  Valor Líquido" e cujos
// números só aparecem na linha do item, depois de Qt.Pedido/Dt.Remessa/Preço:
//   00010 750002891 7.03 SERV ... 28,985 UN 01.04.2026 23,899948 692,74
// Um regex ancorado no rótulo pegaria o primeiro número do item (a quantidade).
// Então: recorta o bloco de itens (do cabeçalho até os tributos/fim) e soma os
// números com EXATAMENTE 2 casas decimais — Qt.Pedido tem 3 e Preço Líquido 6,
// de modo que cada linha de item contribui só com o seu Valor Líquido.
function extrairValorLiquido(texto) {
  const cab = texto.match(/valor\s+l[íi]quido/i);
  if (cab) {
    const resto = texto.slice(cab.index + cab[0].length);
    const bloco = resto.split(/\b(?:CBS\d|IB\dS|IBS|Total\s|Observa)/i)[0];
    const valores = bloco.match(/\d{1,3}(?:\.\d{3})*,\d{2}(?!\d)/g) || [];
    if (valores.length) {
      const soma = valores.reduce((acc, v) => acc + valorNumerico(v), 0);
      return soma.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
  }
  // Fallback para notas/documentos com rótulo explícito ("Valor Faturado: R$ ...").
  const m = texto.match(/valor\s+(?:faturado|total|da\s+nota)\s*[:\-]?\s*(?:R\$)?\s*([\d.\s]+,\d{2})/i);
  return m ? m[1].trim() : '';
}

function extrairData(texto) {
  const rotulada = texto.match(/\bdata\s*(?:de\s*)?(?:emiss[ãa]o|faturamento|do\s+pedido)?\s*[:\-]\s*(\d{2}[\/.\-]\d{2}[\/.\-]\d{4})/i);
  const qualquer = rotulada || texto.match(/(\d{2}[\/.\-]\d{2}[\/.\-]\d{4})/);
  return qualquer ? qualquer[1].replace(/[.\-]/g, '/') : '';
}

// O nome do cliente vem sempre depois do protocolo, na linha do circuito:
//   Circuito: 010354387 - PROTOCOLO: 202530482590-1 - UNIDAS LOCADORA S A
// Ou seja: o que sobra depois de "PROTOCOLO: <código> - ", até o próximo bloco
// do documento. Fallback para os rótulos usuais ("Cliente:", "Razão Social:").
function extrairCliente(texto) {
  const PARADA = /\s+(?:Valor|Pre[çc]o|Circuito|Item|Material|Denomina|Qt\.|Dt\.|CNPJ|CPF|Endere[çc]o|O\.?S\b|GAP|Pedido|Total|Observa)/i;
  const prot = texto.match(/PROTOCOLO\s*:?\s*[\w.\-]+\s*[-–\/]\s*(.{3,80})/i);
  // Só vale o nome que vem depois do protocolo. Em qualquer outro formato o
  // campo fica em branco (exibido como "não encontrado") em vez de chutar.
  if (!prot) return '';
  return prot[1].split(PARADA)[0].trim().replace(/[-–\s]+$/, '');
}

// Procura na lista de medições carregada a O.S correspondente a uma GAP lida do
// PDF. Serve tanto para preencher a O.S quando o PDF não a traz quanto para
// sinalizar divergência entre a O.S do PDF e a da planilha.
function osPelaGap(gap) {
  const alvo = String(gap || '').replace(/\D/g, '');
  if (!alvo) return '';
  const achada = rows.find(r => String(r['GAP'] || '').replace(/\D/g, '') === alvo && String(r['O.S'] || '').trim() !== '');
  return achada ? String(achada['O.S']).trim() : '';
}

function padOs(v) {
  const d = String(v || '').replace(/\D/g, '');
  return d ? d.padStart(10, '0') : d; // sempre 10 dígitos, como `normalizarOs()`
}

// Rótulo exibido → como o campo é procurado no texto do PDF.
// `re` roda sobre o texto normalizado (espaços colapsados, uma linha só);
// campos com layout tabular usam `extrai(texto)` no lugar do regex.
const PDF_CAMPOS = [
  // A O.S é exibida com zeros à esquerda ("0019408611"); quando o PDF traz o
  // número "pelado", padOs() repõe o padding até 10 dígitos.
  { chave: 'O.S', re: /\bO\.?\s?S\.?\s*(?:n[ºo°]?|:|-|n[úu]mero)?\s*[:\-]?\s*(\d{6,12})/i, formata: padOs },
  { chave: 'GAP', re: /\bGAP\s*[-–:]?\s*(\d{3,8})/i, formata: v => 'GAP-' + v },
  // "Pedido de Compra" e "Nº do Pedido" são o mesmo campo no PDF; idem
  // "Valor Faturado" e "Valor Líquido" — por isso cada um é um único item aqui.
  { chave: 'Pedido de Compra', re: /(?:pedido\s+de\s+compra|n[úu]mero\s+do\s+pedido|n[ºo°]\s*(?:do\s*)?pedido|pedido\s*n[ºo°])\s*[:\-]?\s*([A-Za-z0-9\/\-]{4,20})/i },
  { chave: 'Valor Faturado', extrai: extrairValorLiquido },
  // A data do pedido vem logo abaixo do "Nº do Pedido", rotulada "Data:" e com
  // pontos (18.04.2026). Prioriza o campo rotulado; só cai para "qualquer data
  // no documento" se não houver rótulo. Normalizada para dd/mm/aaaa.
  { chave: 'Data', extrai: extrairData },
  { chave: 'Cliente', extrai: extrairCliente }
];

function abrirLeituraPdf() {
  pdfCamposLidos = null;
  pdfNomeArquivo = '';
  pdfResultados.innerHTML = '';
  pdfUsarBtn.disabled = true;
  pdfPanelText.textContent = 'Selecione um arquivo .pdf para procurar O.S, GAP, Pedido de Compra (Nº do Pedido), Valor Faturado (Valor Líquido), Data e Cliente.';
  pdfModalOverlay.classList.add('open');
  pdfInput.click();
}

function fecharLeituraPdf() {
  pdfModalOverlay.classList.remove('open');
}

pdfBtn.addEventListener('click', abrirLeituraPdf);
pdfOutroBtn.addEventListener('click', () => pdfInput.click());
pdfFecharBtn.addEventListener('click', fecharLeituraPdf);
pdfModalOverlay.addEventListener('click', (e) => {
  if (e.target === pdfModalOverlay) fecharLeituraPdf();
});

pdfInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  pdfInput.value = ''; // permite reabrir o mesmo arquivo depois
  if (!file) return;

  pdfModalOverlay.classList.add('open');
  pdfResultados.innerHTML = '';
  pdfUsarBtn.disabled = true;
  pdfNomeArquivo = file.name;
  pdfPanelText.textContent = `Lendo "${file.name}"...`;

  try {
    const texto = await extrairTextoPdf(await file.arrayBuffer());
    pdfCamposLidos = procurarCamposNoTexto(texto);
    // A GAP identifica a O.S na planilha já carregada: usa isso para preencher a
    // O.S quando o PDF não a traz, ou para avisar quando as duas divergem.
    const osLista = osPelaGap(pdfCamposLidos['GAP']);
    if (osLista) {
      if (!pdfCamposLidos['O.S']) {
        pdfCamposLidos['O.S'] = osLista;
        pdfCamposLidos['_osOrigem'] = 'identificada pela GAP na lista';
      } else if (pdfCamposLidos['O.S'].replace(/\D/g, '') !== osLista.replace(/\D/g, '')) {
        pdfCamposLidos['_osOrigem'] = `atenção: a GAP aponta para a O.S ${osLista} na lista`;
      }
    }
    const achados = PDF_CAMPOS.filter(({ chave }) => pdfCamposLidos[chave]).length;
    pdfPanelText.textContent = `"${file.name}" — ${achados} de ${PDF_CAMPOS.length} campo(s) encontrado(s).`;
    pdfResultados.innerHTML = PDF_CAMPOS.map(({ chave }) => {
      const val = pdfCamposLidos[chave];
      // A O.S pode vir da planilha (via GAP) em vez do PDF — a origem é anotada ao lado.
      const nota = (chave === 'O.S' && pdfCamposLidos['_osOrigem'])
        ? ` <em>(${escapeHtml(pdfCamposLidos['_osOrigem'])})</em>` : '';
      return `<div class="pdf-campo${val ? '' : ' vazio'}"><strong>${chave}</strong><span class="pdf-valor">${val ? escapeHtml(val) : 'não encontrado'}${nota}</span></div>`;
    }).join('');
    // Sem medição NÃO FATURADA com essa O.S não há o que atualizar: o botão fica
    // desabilitado e o motivo é explicado no próprio modal.
    const candidatas = linhasPelaOs(pdfCamposLidos['O.S']);
    const semOs = candidatas.length === 0;
    pdfUsarBtn.disabled = achados === 0 || semOs;
    pdfUsarBtn.title = semOs
      ? 'Nenhuma medição não faturada com essa O.S — não há registro para atualizar'
      : '';
    if (semOs) {
      const qual = pdfCamposLidos['O.S']
        ? `a O.S ${pdfCamposLidos['O.S']} não foi encontrada entre as medições não faturadas`
        : 'o PDF não trouxe nenhuma O.S';
      pdfResultados.innerHTML += `<div class="pdf-campo vazio"><strong>Sem O.S para atualizar</strong><span class="pdf-valor">${escapeHtml(qual)}</span></div>`;
    }
  } catch (err) {
    pdfCamposLidos = null;
    pdfPanelText.textContent = 'Erro ao ler o PDF: ' + err.message;
    console.error(err);
  }
});

// Concatena o texto de todas as páginas do PDF numa única string normalizada.
async function extrairTextoPdf(arrayBuffer) {
  if (typeof pdfjsLib === 'undefined') {
    throw new Error('biblioteca pdf.js não carregada (verifique a conexão com a internet).');
  }
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
  const doc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let texto = '';
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    texto += content.items.map(it => it.str).join(' ') + ' ';
  }
  return texto.replace(/\s+/g, ' ').trim();
}

function procurarCamposNoTexto(texto) {
  const out = {};
  PDF_CAMPOS.forEach(({ chave, re, formata, extrai }) => {
    let val;
    if (extrai) {
      val = extrai(texto) || '';
    } else {
      const m = texto.match(re);
      val = m ? m[1].trim() : '';
    }
    if (val && formata) val = formata(val);
    out[chave] = val;
  });
  return out;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

// Procura o registro com a MESMA O.S do PDF (comparando só os dígitos, já que a
// O.S aparece com zeros à esquerda em uns lugares e sem em outros).
// A busca é restrita aos NÃO FATURADOS (`!hasPedido`): são justamente os que
// estão esperando o pedido de compra do PDF. Linhas já faturadas nunca são
// alteradas pela leitura de PDF, mesmo que a O.S bata.
function linhasPelaOs(os) {
  const alvo = String(os || '').replace(/\D/g, '');
  if (!alvo) return [];
  return rows.filter(r => !hasPedido(r) && String(r['O.S'] || '').replace(/\D/g, '') === alvo);
}

// Grava um novo registro direto no autosave do Controle de Pedidos de Compra
// (outro banco IndexedDB, mesma origem). O `manualRows` de lá é exatamente onde
// moram os pedidos cadastrados à mão, então o registro aparece assim que aquela
// aba for aberta/recarregada. Se o Pedidos ainda não tiver planilha importada,
// não há autosave para complementar e a função avisa devolvendo false.
async function criarPedidoNoControle(campos) {
  const valor = valorNumerico(campos['Valor Faturado']);
  let recebimento = null;
  if (campos['Data']) {
    const [d, m, a] = campos['Data'].split('/');
    const dt = new Date(`${a}-${m}-${d}T00:00:00`);
    if (!isNaN(dt.getTime())) recebimento = dt;
  }
  const novoPedido = {
    "Recebimento Pedido": recebimento,
    "Número Pedido": campos['Pedido de Compra'] ? String(campos['Pedido de Compra']).toUpperCase() : null,
    "Valor": isNaN(valor) ? null : valor,
    "O.S": campos['O.S'] ? String(campos['O.S']).toUpperCase() : null,
    "Cliente": campos['Cliente'] ? String(campos['Cliente']).toUpperCase() : null,
    "Gap": campos['GAP'] ? String(campos['GAP']).toUpperCase() : null,
    "Observação": `Importado do PDF ${pdfNomeArquivo}`.toUpperCase(),
  };
  try {
    const db = await new Promise((resolve, reject) => {
      const req = indexedDB.open('pedidos_autosave', 1);
      req.onupgradeneeded = () => req.result.createObjectStore('estado');
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    const estado = await new Promise((resolve, reject) => {
      const tx = db.transaction('estado', 'readonly');
      const req = tx.objectStore('estado').get('atual');
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
    if (!estado || !estado.fileBytes) return { criado: false, motivo: 'sem-planilha' };
    estado.manualRows = Array.isArray(estado.manualRows) ? estado.manualRows : [];

    // Duplicata: já existe um pedido cadastrado à mão com o mesmo Nº do Pedido
    // ou a mesma O.S. Só os manuais dão para conferir daqui (as linhas
    // importadas ficam nos bytes do .xlsx, que não são re-parseados aqui).
    const digitos = v => String(v || '').replace(/\D/g, '');
    const dup = estado.manualRows.find(p =>
      (digitos(novoPedido['Número Pedido']) && digitos(p['Número Pedido']) === digitos(novoPedido['Número Pedido'])) ||
      (digitos(novoPedido['O.S']) && digitos(p['O.S']) === digitos(novoPedido['O.S']))
    );
    if (dup) {
      const ok = confirm(
        `Já existe um pedido cadastrado com Nº ${dup['Número Pedido'] || '-'} / O.S ${dup['O.S'] || '-'} no Controle de Pedidos de Compra.

` +
        'Adicionar mesmo assim (mantendo os dois)?'
      );
      if (!ok) return { criado: false, motivo: 'duplicado-cancelado' };
    }

    estado.manualRows.push(novoPedido);
    estado.savedAt = Date.now();
    await new Promise((resolve, reject) => {
      const tx = db.transaction('estado', 'readwrite');
      tx.objectStore('estado').put(estado, 'atual');
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
    return { criado: true, duplicado: !!dup };
  } catch (err) {
    console.error('Falha ao criar o pedido no Controle de Pedidos de Compra:', err);
    return { criado: false, motivo: 'erro' };
  }
}

// Ação principal da leitura, em duas partes:
// 1) complementa o registro NÃO FATURADO de mesma O.S com GAP, Pedido de Compra
//    e Valor Faturado (sem mexer nos demais campos);
// 2) sempre cria um registro novo no Controle de Pedidos de Compra com o que foi
//    lido — mesmo quando nenhuma O.S bate, aproveitando só o que o PDF trouxe.
// Sem O.S correspondente, o cadastro de medição pré-preenchido ainda é aberto.
pdfUsarBtn.addEventListener('click', async () => {
  if (!pdfCamposLidos) return;

  const camposPdf = pdfCamposLidos;
  const res = await criarPedidoNoControle(camposPdf);
  let avisoPedido;
  if (res.criado) {
    avisoPedido = res.duplicado
      ? ' Um pedido DUPLICADO foi criado no Controle de Pedidos de Compra (confirmado por você).'
      : ' Um pedido foi criado no Controle de Pedidos de Compra.';
  } else if (res.motivo === 'duplicado-cancelado') {
    avisoPedido = ' Nenhum pedido criado no Controle de Pedidos de Compra (duplicata cancelada por você).';
  } else if (res.motivo === 'sem-planilha') {
    avisoPedido = ' (Não foi possível criar o pedido no Controle de Pedidos de Compra: importe a planilha de pedidos naquela aba primeiro.)';
  } else {
    avisoPedido = ' (Falha ao criar o pedido no Controle de Pedidos de Compra — veja o console.)';
  }

  // Mais de uma medição não faturada com a mesma O.S: o usuário precisa saber
  // qual será alterada (a primeira da lista) antes de seguir.
  const candidatas = linhasPelaOs(camposPdf['O.S']);
  if (candidatas.length > 1) {
    const seguir = confirm(
      `Atenção: há ${candidatas.length} medições NÃO FATURADAS com a O.S ${camposPdf['O.S']}.

` +
      'Só a primeira delas será atualizada. Continuar?'
    );
    if (!seguir) {
      fecharLeituraPdf();
      fileStatus.textContent = `Atualização cancelada: ${candidatas.length} medições não faturadas com a O.S ${camposPdf['O.S']}.` + avisoPedido;
      return;
    }
  }

  const existente = candidatas[0];
  if (existente) {
    const valorPdf = valorNumerico(pdfCamposLidos['Valor Faturado']);
    const atualizados = [];
    if (pdfCamposLidos['GAP']) { existente['GAP'] = String(pdfCamposLidos['GAP']).toUpperCase(); atualizados.push('GAP'); }
    if (pdfCamposLidos['Pedido de Compra']) { existente['Pedido de Compra'] = String(pdfCamposLidos['Pedido de Compra']).toUpperCase(); atualizados.push('Pedido de Compra'); }
    if (!isNaN(valorPdf)) { existente['Valor Faturado'] = valorPdf; atualizados.push('Valor Faturado'); }

    // A linha pode ser original da planilha: garante que ela também seja
    // reescrita na exportação, entrando em newRows se ainda não estiver lá.
    if (atualizados.length && !newRows.includes(existente)) newRows.push(existente);

    fecharLeituraPdf();
    renderTable();
    agendarAutosave();
    fileStatus.textContent = (atualizados.length
      ? `O.S ${existente['O.S']} (não faturada) atualizada a partir de "${pdfNomeArquivo}": ${atualizados.join(', ')}.`
      : `O.S ${existente['O.S']} encontrada, mas o PDF não trouxe GAP, Pedido de Compra nem Valor Faturado.`) + avisoPedido;
    sincronizarComPedidos({ avisar: false });
    return;
  }

  // Caminho defensivo: o botão já fica desabilitado quando não há medição não
  // faturada com essa O.S, então aqui só resta fechar e explicar.
  fecharLeituraPdf();
  fileStatus.textContent = `Nenhuma medição não faturada com a O.S ${camposPdf['O.S'] || '(não encontrada)'} — nada foi atualizado.` + avisoPedido;
  sincronizarComPedidos({ avisar: false });
});

// --- Sincronização com o Controle de Pedidos de Compra -----------------------
// Sempre que a lista de medições é (re)carregada, os registros do Controle de
// Pedidos de Compra são conferidos contra TODAS as medições — faturadas e não
// faturadas — e os campos equivalentes que estiverem VAZIOS na medição são
// preenchidos com o que o pedido já traz. Nada que já esteja preenchido é
// sobrescrito, e a planilha de pedidos nunca é alterada por aqui.
//
// Os pedidos vêm do autosave da outra aba (`pedidos_autosave`, mesma origem):
// os bytes do .xlsx importado lá são reparseados aqui (aplicando exclusões e
// edições salvas) e somados aos `manualRows` cadastrados à mão/por PDF.
const CAMPOS_PEDIDO_MEDICAO = [
  { pedido: 'Número Pedido', medicao: 'Pedido de Compra' },
  { pedido: 'Valor', medicao: 'Valor Faturado' },
  { pedido: 'Gap', medicao: 'GAP' },
  { pedido: 'O.S', medicao: 'O.S' }
];

// Só os dígitos E sem zeros à esquerda: a mesma O.S aparece como "0021214728" na
// planilha de Medições e "21214728" na de Pedidos (e vice-versa para o nº do
// pedido). Comparar só `\D` fazia esses pares deixarem de casar.
const soDigitos = v => String(v === null || v === undefined ? '' : v).replace(/\D/g, '').replace(/^0+/, '');
const campoVazio = v => v === null || v === undefined || String(v).trim() === '';

async function lerPedidosDoControle() {
  const db = await new Promise((resolve, reject) => {
    const req = indexedDB.open('pedidos_autosave', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('estado');
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  const estado = await new Promise((resolve, reject) => {
    const tx = db.transaction('estado', 'readonly');
    const req = tx.objectStore('estado').get('atual');
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
  if (!estado) return [];

  const COLS_PEDIDOS = ["Recebimento Pedido", "Número Pedido", "Valor", "O.S", "Cliente", "Gap", "Observação"];
  let doArquivo = [];
  if (estado.fileBytes) {
    const wb = XLSX.read(estado.fileBytes, { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const txt = v => (v === null || v === undefined) ? '' : String(v).trim();
    let hIdx = json.findIndex(r => r && (txt(r[3]) === 'O.S' || txt(r[1]) === 'Número Pedido'));
    if (hIdx === -1) hIdx = 2;
    const excluidas = new Set(Array.isArray(estado.linhasExcluidas) ? estado.linhasExcluidas : []);
    doArquivo = json.slice(hIdx + 1)
      .map((r, idx) => {
        if (!r || !r.some(c => c !== undefined && c !== null && c !== '')) return null;
        const obj = { __linha: idx };
        COLS_PEDIDOS.forEach((col, i) => {
          const val = r[i];
          obj[col] = (val === undefined || val === null) ? null : val;
        });
        if (!obj['Número Pedido'] && !obj['O.S'] && !obj['Cliente']) return null;
        return obj;
      })
      .filter(p => p && !excluidas.has(p.__linha));
    (Array.isArray(estado.linhasEditadas) ? estado.linhasEditadas : []).forEach(e => {
      const pos = doArquivo.findIndex(p => p.__linha === e.__linha);
      if (pos !== -1) doArquivo[pos] = e;
    });
  }
  // `filter(Boolean)`/typeof: um autosave antigo ou corrompido pode trazer
  // null/strings dentro de manualRows, e um único item inválido derrubava toda a
  // sincronização com TypeError.
  const manuais = (Array.isArray(estado.manualRows) ? estado.manualRows : [])
    .filter(p => p && typeof p === 'object');
  return doArquivo.concat(manuais);
}

// Índices de busca: um pedido é ligado a uma medição pela O.S, pelo número do
// pedido ou pela GAP — nessa ordem de prioridade.
function indexarPedidos(pedidos) {
  const porOs = new Map(), porNumero = new Map(), porGap = new Map();
  pedidos.forEach(p => {
    const os = soDigitos(p['O.S']);
    const num = soDigitos(p['Número Pedido']);
    const gap = String(p['Gap'] || '').trim().toUpperCase();
    if (os && !porOs.has(os)) porOs.set(os, p);
    if (num && !porNumero.has(num)) porNumero.set(num, p);
    if (gap && !porGap.has(gap)) porGap.set(gap, p);
  });
  return { porOs, porNumero, porGap };
}

function pedidoDaMedicao(row, idx) {
  const os = soDigitos(row['O.S']);
  if (os && idx.porOs.has(os)) return idx.porOs.get(os);

  // Casamento por nº do pedido ou por GAP só vale se as O.S não se contradisserem:
  // GAPs se repetem entre O.S diferentes na base real, e sem essa guarda uma
  // medição herdava o Pedido de Compra de outra O.S.
  const conflita = p => os && soDigitos(p['O.S']) && soDigitos(p['O.S']) !== os;

  const num = soDigitos(row['Pedido de Compra']);
  if (num && idx.porNumero.has(num)) {
    const p = idx.porNumero.get(num);
    if (!conflita(p)) return p;
  }
  const gap = String(row['GAP'] || '').trim().toUpperCase();
  if (gap && idx.porGap.has(gap)) {
    const p = idx.porGap.get(gap);
    if (!conflita(p)) return p;
  }
  return null;
}

// Retorna quantas medições foram complementadas. Silenciosa quando não há nada
// a fazer (ou quando a aba de Pedidos ainda não tem planilha importada).
async function sincronizarComPedidos({ avisar = true } = {}) {
  if (!workbook || !rows.length) return 0;
  let pedidos;
  try {
    pedidos = await lerPedidosDoControle();
  } catch (err) {
    console.error('Falha ao ler o Controle de Pedidos de Compra:', err);
    return 0;
  }
  if (!pedidos.length) return 0;

  const idx = indexarPedidos(pedidos);
  let alteradas = 0, campos = 0;
  rows.forEach(row => {
    const p = pedidoDaMedicao(row, idx);
    if (!p) return;
    let mudou = false;
    CAMPOS_PEDIDO_MEDICAO.forEach(({ pedido, medicao }) => {
      if (!campoVazio(row[medicao])) return; // nunca sobrescreve o que já existe
      const origem = p[pedido];
      if (campoVazio(origem)) return;
      if (CURRENCY_COLS.has(medicao)) {
        const n = valorNumerico(origem);
        if (isNaN(n)) return;
        row[medicao] = n;
      } else {
        const texto = String(origem).trim().toUpperCase();
        row[medicao] = medicao === 'O.S' ? normalizarOs(texto) : texto;
      }
      mudou = true;
      campos++;
      // Marca a célula como preenchida pela sincronização, só pra destacar na
      // tela (renderTable) — não é lido em nenhum outro lugar, não afeta export.
      if (!row.__sincronizados) row.__sincronizados = [];
      if (!row.__sincronizados.includes(medicao)) row.__sincronizados.push(medicao);
    });
    if (mudou) {
      alteradas++;
      // Linha do arquivo original também precisa sair na exportação.
      if (!newRows.includes(row)) newRows.push(row);
    }
  });

  if (alteradas) {
    renderTable();
    agendarAutosave();
    if (avisar) {
      fileStatus.textContent =
        `${fileStatus.textContent} ${alteradas} medição(ões) complementada(s) com dados do Controle de Pedidos de Compra (${campos} campo(s) preenchido(s)).`.trim();
    }
  }
  return alteradas;
}

// --- Lixeira -----------------------------------------------------------------
// Excluir NÃO apaga o registro: move para a aba "Lixeira", onde ele fica por
// LIXEIRA_DIAS dias antes de ser descartado de vez. Enquanto está na lixeira o
// registro sai de `rows`/`newRows`, então não aparece nas outras abas, nos
// totais nem nas planilhas baixadas.
function rotuloLinha(row) {
  return row['O.S'] ? `O.S ${row['O.S']}` : 'este registro';
}

function diasRestantesLixeira(item) {
  return Math.max(0, Math.ceil((item.excluidoEm + LIXEIRA_MS - Date.now()) / 86400000));
}

// Descarta definitivamente o que já passou dos 30 dias. Chamado ao restaurar o
// autosave e a cada renderização, para o prazo valer mesmo sem interação.
function purgarLixeira() {
  const limite = Date.now() - LIXEIRA_MS;
  const antes = lixeira.length;
  lixeira = lixeira.filter(item => item && item.excluidoEm > limite);
  return antes - lixeira.length;
}

function excluirLinha(row) {
  const os = rotuloLinha(row);
  if (!window.confirm(`Mover ${os} para a lixeira?\n\nO registro sai da lista e das planilhas baixadas, mas fica ${LIXEIRA_DIAS} dias na aba "Lixeira", de onde pode ser restaurado. Depois disso é apagado por completo.`)) return;

  if (row.__linha !== undefined) linhasExcluidas.add(row.__linha);
  rows = rows.filter(r => r !== row);
  newRows = newRows.filter(r => r !== row);
  lixeira.push({ linha: row, excluidoEm: Date.now() });
  if (editingRow === row) closeForm();
  renderTable();
  fileStatus.textContent = `${os} movido para a lixeira — será apagado definitivamente em ${LIXEIRA_DIAS} dias.`;
  agendarAutosave();
}

function restaurarDaLixeira(item) {
  const row = item.linha;
  if (row.__linha !== undefined) linhasExcluidas.delete(row.__linha);
  lixeira = lixeira.filter(i => i !== item);
  rows.push(row);
  // Volta marcada como pendente de gravação: ela não está mais no arquivo que o
  // usuário tem em mãos (ou pode nunca ter estado, se foi cadastrada aqui).
  if (!newRows.includes(row)) newRows.push(row);
  renderTable();
  fileStatus.textContent = `${rotuloLinha(row)} restaurado da lixeira.`;
  agendarAutosave();
}

function excluirDefinitivo(item) {
  const os = rotuloLinha(item.linha);
  if (!window.confirm(`Apagar ${os} definitivamente?\n\nEsta ação não pode ser desfeita.`)) return;
  lixeira = lixeira.filter(i => i !== item);
  renderTable();
  fileStatus.textContent = `${os} apagado definitivamente.`;
  agendarAutosave();
}

function esvaziarLixeira() {
  if (!lixeira.length) return;
  if (!window.confirm(`Esvaziar a lixeira (${lixeira.length} registro(s))?\n\nEsta ação não pode ser desfeita.`)) return;
  lixeira = [];
  renderTable();
  fileStatus.textContent = 'Lixeira esvaziada.';
  agendarAutosave();
}

function openForm() {
  editingRow = null;
  saveRowBtn.textContent = 'Salvar Medição';
  formPanelTitle.textContent = 'Cadastrar Medição';
  formModalOverlay.classList.add('open');
  formMsg.textContent = '';
  formMsg.className = '';
  const today = new Date().toISOString().split('T')[0];
  document.getElementById('f_envio').value = today;
  document.getElementById('f_ano').value = new Date().getFullYear();
  atualizarSla();
  atualizarCustoTotal();
  marcarCamposVazios();
}

function openFormParaEdicao(row) {
  editingRow = row;
  saveRowBtn.textContent = 'Salvar Alterações';
  formPanelTitle.textContent = 'Editar Medição';
  formModalOverlay.classList.add('open');
  formMsg.textContent = '';
  formMsg.className = '';

  document.getElementById('f_os').value = row['O.S'] ?? '';
  document.getElementById('f_mes').value = row['Mês Envio Medição'] || 'Janeiro';
  document.getElementById('f_vistoria').value = toDateInputValue(row['Vistoria Projeto']);
  document.getElementById('f_ano').value = row['Ano'] ?? new Date().getFullYear();
  document.getElementById('f_gap').value = row['GAP'] ?? '';
  document.getElementById('f_envio').value = toDateInputValue(row['Envio Medição']);
  document.getElementById('f_valorMedicao').value = row['Valor Medição Enviadas'] ?? '';
  document.getElementById('f_engenheiro').value = row['Engenheiro Acesso'] ?? '';
  document.getElementById('f_projetista').value = row['Projetista'] ?? '';
  document.getElementById('f_regional').value = row['Regional'] ?? '';
  document.getElementById('f_pedido').value = row['Pedido de Compra'] ?? '';
  document.getElementById('f_valorFaturado').value = row['Valor Faturado'] ?? '';
  document.getElementById('f_medicaoTerceiros').value = row['Medição Terceiros'] ?? '';
  document.getElementById('f_custosDiversos').value = row['Custos Diversos'] ?? '';
  document.getElementById('f_custoArt').value = row['Custo ART'] ?? '';

  atualizarSla();
  atualizarCustoTotal();
  marcarCamposVazios();
}

function closeForm() {
  formModalOverlay.classList.remove('open');
  camposComDestaqueVazio.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove('empty-required');
  });
  document.querySelectorAll('#formPanel input').forEach(i => { if (i.type !== 'date') i.value = ''; });
  document.getElementById('f_mes').selectedIndex = 0;
  editingRow = null;
  saveRowBtn.textContent = 'Salvar Medição';
}

function diasDesde(dataStr) {
  if (!dataStr) return null;
  const inicio = new Date(dataStr + 'T00:00:00');
  const hoje = new Date();
  hoje.setHours(0,0,0,0);
  const diffMs = hoje - inicio;
  return Math.floor(diffMs / 86400000);
}

function atualizarSla() {
  const envio = document.getElementById('f_envio').value;
  const dias = diasDesde(envio);
  document.getElementById('f_sla').value = dias !== null ? dias : '';
}

function parseNum(v) {
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function atualizarCustoTotal() {
  const diversos = parseNum(document.getElementById('f_custosDiversos').value);
  const art = parseNum(document.getElementById('f_custoArt').value);
  document.getElementById('f_custoTotal').value = (diversos + art).toFixed(2);
}

function campoPreenchido(v) {
  if (v === null || v === undefined) return false;
  if (typeof v === 'number') return !isNaN(v);
  return String(v).trim() !== '';
}

function salvarMedicao() {
  const os = normalizarOs(document.getElementById('f_os').value.trim().toUpperCase());
  const envio = document.getElementById('f_envio').value;
  if (!os) {
    formMsg.textContent = 'Preencha o campo O.S.';
    formMsg.className = 'error';
    return;
  }
  if (!envio) {
    formMsg.textContent = 'Preencha a data de Envio Medição.';
    formMsg.className = 'error';
    return;
  }

  const novoPedido = document.getElementById('f_pedido').value.trim().toUpperCase();

  // Se está editando uma medição que ainda não tinha Pedido de Compra e o usuário
  // acabou de preenchê-lo agora, a contagem de dias (SLA) deve parar na data de hoje
  // (data do preenchimento) em vez de continuar recalculando em tempo real depois.
  const estavaSemPedido = editingRow && !hasPedido(editingRow);
  const congelandoPedidoAgora = estavaSemPedido && novoPedido !== '';

  if (congelandoPedidoAgora) {
    const confirmado = window.confirm(
      `Confirmar preenchimento do Pedido de Compra "${novoPedido}"?\n\n` +
      `A contagem de dias será congelada na data de hoje e a medição voltará para a aba "Com Pedido".`
    );
    if (!confirmado) return;
  }

  const dadosForm = {
    "O.S": os,
    "Mês Envio Medição": document.getElementById('f_mes').value.toUpperCase(),
    "Vistoria Projeto": document.getElementById('f_vistoria').value ? new Date(document.getElementById('f_vistoria').value + 'T00:00:00') : null,
    "Ano": parseInt(document.getElementById('f_ano').value) || new Date().getFullYear(),
    "GAP": document.getElementById('f_gap').value.trim().toUpperCase(),
    "Envio Medição": new Date(envio + 'T00:00:00'),
    "Sla Envio Medição": diasDesde(envio),
    "Valor Medição Enviadas": parseNum(document.getElementById('f_valorMedicao').value),
    "Engenheiro Acesso": document.getElementById('f_engenheiro').value.trim().toUpperCase(),
    "Projetista": document.getElementById('f_projetista').value.trim().toUpperCase(),
    "Regional": document.getElementById('f_regional').value.trim().toUpperCase(),
    "Pedido de Compra": novoPedido,
    "Valor Faturado": parseNum(document.getElementById('f_valorFaturado').value),
    "Medição Terceiros": parseNum(document.getElementById('f_medicaoTerceiros').value),
    "Custos Diversos": parseNum(document.getElementById('f_custosDiversos').value),
    "Custo ART": parseNum(document.getElementById('f_custoArt').value),
    "Custo Total": parseNum(document.getElementById('f_custoTotal').value),
  };

  if (editingRow) {
    Object.assign(editingRow, dadosForm);
    editingRow.__isNew = true; // marca como alterada para refletir no download da planilha
    // O formulário acabou de reescrever todos os campos à mão — a marca de
    // "veio da sincronização" (destaque visual na célula) não faz mais sentido.
    delete editingRow.__sincronizados;

    if (congelandoPedidoAgora) {
      // Congela o SLA no número de dias já contados até hoje (data do preenchimento
      // do Pedido de Compra), para não continuar recalculando em tempo real depois.
      editingRow.__slaCongelado = diasDesde(envio);
    }

    if (!newRows.includes(editingRow)) newRows.push(editingRow);

    const promovida = hasPedido(editingRow);
    renderTable();
    formMsg.textContent = promovida
      ? 'Medição atualizada e movida para "Com Pedido"! Clique em "Baixar planilha atualizada" para salvar no arquivo.'
      : 'Medição atualizada! Clique em "Baixar planilha atualizada" para salvar no arquivo.';
    formMsg.className = 'ok';
    closeForm();
    if (promovida) setActiveTab('comPedido');
    agendarAutosave();
    // Confere na hora se o Controle de Pedidos já tem o que falta nessa medição
    // (antes só rodava no import/autosave/leitura de PDF — cadastro e edição
    // manuais ficavam esperando o próximo refresh pra serem complementados).
    sincronizarComPedidos({ avisar: false });
  } else {
    const novaLinha = { ...dadosForm, __isNew: true };
    rows.push(novaLinha);
    newRows.push(novaLinha);
    renderTable();
    formMsg.textContent = 'Medição adicionada! Clique em "Baixar planilha atualizada" para salvar no arquivo.';
    formMsg.className = 'ok';
    closeForm();
    document.getElementById('tableWrap').scrollTop = document.getElementById('tableWrap').scrollHeight;
    agendarAutosave();
    sincronizarComPedidos({ avisar: false });
  }
}

// Converte um valor de moeda para número, reaproveitando a limpeza usada em
// fmtCurrency (parte das linhas antigas guarda o valor já como texto formatado).
// Retorna NaN quando não é conversível.
function valorNumerico(v) {
  if (v === null || v === undefined || v === '') return NaN;
  if (typeof v === 'number') return v;
  const bruto = String(v).replace(/R\$/gi, '').trim();
  // "R$ -" é a formatação contábil de zero usada na planilha, não um valor inválido.
  if (/^-+$/.test(bruto)) return 0;
  return parseFloat(bruto.replace(/[\s .]/g, '').replace(',', '.'));
}

function fmtCurrency(v) {
  if (v === null || v === undefined || v === '') return '';
  let n;
  if (typeof v === 'number') {
    n = v;
  } else {
    // Remove "R$", espaços (inclusive não separáveis) e separador de milhar,
    // depois troca a vírgula decimal por ponto (ex: "R$ 1 278,82" -> "1278.82").
    const limpo = String(v).replace(/R\$/gi, '').replace(/[\s .]/g, '').replace(',', '.');
    n = parseFloat(limpo);
  }
  if (isNaN(n)) return v;
  if (n === 0) return '';
  return n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtDate(v) {
  if (!v) return '';
  const d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return v;
  return d.toLocaleDateString('pt-BR');
}

function slaAtual(r) {
  // Se o Pedido de Compra foi preenchido nesta sessão, a contagem foi congelada
  // no número de dias já decorridos até a data do preenchimento.
  if (r.__slaCongelado !== undefined && r.__slaCongelado !== null) return r.__slaCongelado;

  // Caso contrário, calculado em tempo real: dias corridos entre "Envio Medição" e
  // hoje, ignorando qualquer valor de SLA gravado na planilha original.
  const envio = r['Envio Medição'];
  if (!envio) return '';
  const d = (envio instanceof Date) ? envio : new Date(envio);
  if (isNaN(d.getTime())) return '';
  const inicio = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  return Math.floor((hoje - inicio) / 86400000);
}

function matchesSearch(r) {
  if (!searchTerm) return true;
  return [...searchFields].some(field => {
    const v = r[field];
    return v !== null && v !== undefined && String(v).toLowerCase().includes(searchTerm);
  });
}

// Ordenação por coluna (clique no cabeçalho alterna crescente â†’ decrescente â†’ sem ordenação)
const MESES_ORDEM = {
  'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3, 'abril': 4, 'maio': 5,
  'junho': 6, 'julho': 7, 'agosto': 8, 'setembro': 9, 'outubro': 10,
  'novembro': 11, 'dezembro': 12
};
let sortCol = null;
let sortDir = 1; // 1 = crescente, -1 = decrescente

function sortKey(r, col) {
  const val = col === 'Sla Envio Medição' ? slaAtual(r) : r[col];
  if (val === null || val === undefined || val === '') return null;
  if (DATE_COLS.has(col)) {
    const d = (val instanceof Date) ? val : new Date(val);
    return isNaN(d.getTime()) ? null : d.getTime();
  }
  if (col === 'Mês Envio Medição') {
    const m = MESES_ORDEM[String(val).trim().toLowerCase()];
    return m !== undefined ? m : null;
  }
  if (CURRENCY_COLS.has(col)) {
    // Reaproveita valorNumerico() (e não um regex próprio) para que "R$ -" — a
    // formatação contábil de zero — ordene como 0 em vez de cair pro fim da
    // lista como se fosse vazio.
    const n = typeof val === 'number' ? val : valorNumerico(val);
    return isNaN(n) ? null : n;
  }
  if (NUMERIC_CODE_COLS.has(col) || col === 'Ano' || col === 'Sla Envio Medição') {
    const n = typeof val === 'number' ? val : parseFloat(String(val).replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.'));
    return isNaN(n) ? null : n;
  }
  return String(val).toLowerCase();
}

function compareRows(a, b, col) {
  const ka = sortKey(a, col), kb = sortKey(b, col);
  if (ka === null && kb === null) return 0;
  if (ka === null) return 1;  // vazios sempre no fim
  if (kb === null) return -1;
  if (typeof ka === 'string' || typeof kb === 'string') {
    return String(ka).localeCompare(String(kb), 'pt-BR', { numeric: true }) * sortDir;
  }
  return (ka - kb) * sortDir;
}

function toggleSort(col) {
  if (sortCol !== col) { sortCol = col; sortDir = 1; }
  else if (sortDir === 1) { sortDir = -1; }
  else { sortCol = null; sortDir = 1; }
  renderTable();
}

// Linhas efetivamente exibidas: filtro da aba ativa + busca, na ordenação atual.
// Compartilhado por renderTable() e baixarPlanilha(), para que cada aba exporte
// exatamente o que está na tela.
function linhasVisiveis() {
  // Registros completamente vazios nunca são exibidos nem exportados — some com
  // eles de vez, para não voltarem no autosave nem na próxima planilha baixada.
  const vazias = rows.filter(linhaVazia);
  if (vazias.length) {
    rows = rows.filter(r => !linhaVazia(r));
    newRows = newRows.filter(r => !linhaVazia(r));
  }
  const lista = rows.filter(r => {
    if (activeTab === 'comPedido' && !hasPedido(r)) return false;
    if (activeTab === 'semPedido' && hasPedido(r)) return false;
    if (activeTab === 'aCobrar' && !isACobrar(r)) return false;
    return matchesSearch(r);
  });
  if (sortCol) lista.sort((a, b) => compareRows(a, b, sortCol));
  return lista;
}

// Ícones da coluna "Ações" — SVG inline (traço em currentColor, então herdam a
// cor do botão: neutra no lápis, vermelha na lixeira, em ambos os temas).
const ICONE_LAPIS = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4z"/></svg>';
const ICONE_LIXEIRA = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/></svg>';
const ICONE_RESTAURAR = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 12a9 9 0 1 0 3-6.7"/><path d="M3 4v5h5"/></svg>';

function renderTable() {
  purgarLixeira(); // o prazo de 30 dias vale mesmo sem o usuário mexer em nada
  const naLixeira = activeTab === 'lixeira';
  const visibleCols = COLUMNS.filter(c => visibleColumns.has(c));
  headerRow.innerHTML = visibleCols.map(c => {
    const isSorted = sortCol === c;
    const ind = isSorted ? (sortDir === 1 ? '▲' : '▼') : '⇅';
    return `<th class="sortable${isSorted ? ' sorted' : ''}" data-col="${c}" title="Ordenar por ${c}">${c}<span class="sort-ind">${ind}</span></th>`;
  }).join('') + (naLixeira ? '<th class="acoes-col">Excluído em</th><th class="acoes-col">Ações</th>' : '<th class="acoes-col">Ações</th>');
  headerRow.querySelectorAll('th.sortable').forEach(th => {
    th.addEventListener('click', () => toggleSort(th.dataset.col));
  });
  dataTable.classList.toggle('tabela-a-cobrar', activeTab === 'aCobrar');

  const visibleRows = naLixeira ? [] : linhasVisiveis();

  totalsRow.innerHTML = visibleCols.map(col => {
    if (!COLS_SOMATORIO.has(col)) return '<td></td>';
    const soma = visibleRows.reduce((acc, r) => {
      const n = valorNumerico(r[col]);
      return isNaN(n) ? acc : acc + n;
    }, 0);
    return `<td>${fmtCurrency(soma)}</td>`;
  }).join('') + (naLixeira ? '<td></td><td></td>' : '<td></td>');
  totalsRow.style.display = naLixeira ? 'none' : '';

  tableBody.innerHTML = '';

  if (naLixeira) renderLixeira(visibleCols);
  else visibleRows.forEach((r, i) => {
    const tr = document.createElement('tr');
    // Edição por duplo clique e botões de ação valem em TODAS as abas.
    tr.classList.add('editable-row');
    tr.title = 'Clique duas vezes para editar';
    tr.addEventListener('dblclick', () => openFormParaEdicao(r));
    tr.innerHTML = visibleCols.map(col => {
      let val = col === 'Sla Envio Medição' ? slaAtual(r) : r[col];
      if (DATE_COLS.has(col)) val = fmtDate(val);
      else if (CURRENCY_COLS.has(col)) val = fmtCurrency(val);
      else if (val === null || val === undefined) val = '';
      // Célula pode vir de uma planilha importada ou de digitação manual — nunca
      // confiável o bastante pra entrar em innerHTML sem escapar (ver TESTE_DE_FOGO.md).
      // Destaque sutil na célula que a sincronização com Pedidos preencheu sozinha.
      const sincronizada = Array.isArray(r.__sincronizados) && r.__sincronizados.includes(col);
      const attrs = sincronizada ? ' class="cel-sincronizada" title="Preenchido automaticamente pela sincronização com o Controle de Pedidos de Compra"' : '';
      return `<td${attrs}>${escapeHtml(val)}</td>`;
    }).join('') +
      '<td class="acoes-cell">' +
      `<button type="button" class="row-btn editar" title="Editar registro" aria-label="Editar registro">${ICONE_LAPIS}</button>` +
      `<button type="button" class="row-btn excluir" title="Excluir registro" aria-label="Excluir registro">${ICONE_LIXEIRA}</button>` +
      '</td>';
    const btnEditar = tr.querySelector('.row-btn.editar');
    const btnExcluir = tr.querySelector('.row-btn.excluir');
    btnEditar.addEventListener('click', ev => { ev.stopPropagation(); openFormParaEdicao(r); });
    btnExcluir.addEventListener('click', ev => { ev.stopPropagation(); excluirLinha(r); });
    tableBody.appendChild(tr);
  });

  const temConteudo = rows.length || lixeira.length;
  subTabs.classList.toggle('initially-hidden', !temConteudo);
  subTabs.style.display = temConteudo ? 'flex' : 'none';
  searchBar.style.display = temConteudo && !naLixeira ? 'block' : 'none';
  columnsPanelWrap.style.display = temConteudo ? 'flex' : 'none';
  emptyState.style.display = 'none';
  dataTable.classList.remove('initially-hidden');
  dataTable.style.display = '';
  rowCount.textContent = naLixeira
    ? `${lixeira.length} registro(s) na lixeira`
    : `${visibleRows.length} de ${rows.length} registros`;

  const faturadas = rows.filter(hasPedido).length;
  const naoFaturadas = rows.length - faturadas;
  statusCount.textContent = `${faturadas} Faturadas · ${naoFaturadas} Não Faturadas`;

  const totalACobrar = rows.filter(isACobrar).length;
  tabACobrar.classList.toggle('alert', totalACobrar > 0);
  aCobrarBadge.classList.remove('initially-hidden');
  aCobrarBadge.style.display = totalACobrar > 0 ? 'inline-flex' : 'none';
  aCobrarBadge.textContent = totalACobrar;

  comPedidoCount.textContent = `(${faturadas})`;
  semPedidoCount.textContent = `(${naoFaturadas})`;
  enviosCount.textContent = `(${rows.length})`;
  lixeiraCount.textContent = `(${lixeira.length})`;
}

// Corpo da aba "Lixeira": os mesmos campos das outras abas, mais a data de
// exclusão / prazo restante, e ações próprias (restaurar ou apagar de vez).
// Não entra em `rows`, então não afeta buscas, totais nem planilhas baixadas.
function renderLixeira(visibleCols) {
  const itens = [...lixeira].sort((a, b) => b.excluidoEm - a.excluidoEm);
  if (!itens.length) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="${visibleCols.length + 2}" style="text-align:center; padding:18px; opacity:0.7;">Lixeira vazia. Registros excluídos ficam aqui por ${LIXEIRA_DIAS} dias.</td>`;
    tableBody.appendChild(tr);
    return;
  }
  itens.forEach(item => {
    const r = item.linha;
    const tr = document.createElement('tr');
    const dias = diasRestantesLixeira(item);
    tr.innerHTML = visibleCols.map(col => {
      let val = col === 'Sla Envio Medição' ? (r['Sla Envio Medição'] ?? '') : r[col];
      if (DATE_COLS.has(col)) val = fmtDate(val);
      else if (CURRENCY_COLS.has(col)) val = fmtCurrency(val);
      else if (val === null || val === undefined) val = '';
      return `<td>${escapeHtml(val)}</td>`;
    }).join('') +
      `<td>${new Date(item.excluidoEm).toLocaleDateString('pt-BR')} · ${dias} dia(s)</td>` +
      '<td class="acoes-cell">' +
      `<button type="button" class="row-btn restaurar" title="Restaurar registro" aria-label="Restaurar registro">${ICONE_RESTAURAR}</button>` +
      `<button type="button" class="row-btn excluir" title="Apagar definitivamente" aria-label="Apagar definitivamente">${ICONE_LIXEIRA}</button>` +
      '</td>';
    tr.querySelector('.row-btn.restaurar').addEventListener('click', () => restaurarDaLixeira(item));
    tr.querySelector('.row-btn.excluir').addEventListener('click', () => excluirDefinitivo(item));
    tableBody.appendChild(tr);
  });
}

// GAP é o único código realmente textual (ex: "GAP-33607"). O.S e Pedido de Compra
// são numéricos no arquivo original (evita notação científica via largura de coluna
// reforçada em !cols, não via conversão para texto â€” ver fileInput handler).
const TEXT_COLS = new Set(["GAP"]);
const NUMERIC_CODE_COLS = new Set(["O.S", "Pedido de Compra"]);
// Colunas somadas na linha de totais (na tela e na exportação).
const COLS_SOMATORIO = new Set(["Valor Medição Enviadas", "Valor Faturado", "Custo Total", "Medição Terceiros", "Custos Diversos", "Custo ART"]);

function excelSerialFromDate(date) {
  // Converte Date JS para serial de data do Excel (mesma base usada pelo arquivo original)
  const epoch = Date.UTC(1899, 11, 30);
  return (Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - epoch) / 86400000;
}

// Aplica alinhamento centralizado (horizontal e vertical) em todas as células
// dentro do !ref da planilha. Mantém qualquer estilo já existente na célula.
//
// Só estiliza células que EXISTEM no objeto da planilha. Células vazias (campos
// em branco, que o app grava como "sem célula") continuam com o alinhamento
// padrão do Excel até que alguém digite algo nelas. Materializá-las só para
// centralizar levaria o arquivo de ~2,8 MB para ~10,4 MB (199.750 células
// vazias a mais) e quebraria a convenção de célula-em-branco da planilha.
function centralizarCelulas(ws) {
  if (!ws || !ws['!ref']) return;
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell) continue;
      cell.s = Object.assign({}, cell.s, {
        alignment: Object.assign({}, cell.s && cell.s.alignment, {
          horizontal: 'center', vertical: 'center', wrapText: false
        })
      });
    }
  }
}

// Larguras de coluna originais da planilha, indexadas por COLUMNS.
const LARGURAS_COLUNAS = [
  23.608073, 22.166667, 19.053385, 9.053385,
  19.498698, 18.166667, 21.276042, 26.608073,
  22.053385, 19.608073, 13.166667, 21.276042,
  18.608073, 21.608073, 19.385417, 14.166667,
  15.276042, 22.053385
];

// Exporta apenas as colunas marcadas no filtro de colunas. Como a exportação
// vira um subconjunto de colunas, não dá para reaproveitar a planilha original
// (as colunas ocultas deixariam buracos), então montamos uma planilha nova a
// partir das linhas visíveis — reforçando number_format/larguras por célula,
// igual ao que a planilha original trazia. Cada aba exporta somente as linhas
// que ela exibe (filtro da aba + busca), num arquivo com nome próprio.
function baixarPlanilha() {
  if (!workbook) return;

  const colsExport = COLUMNS.filter(c => visibleColumns.has(c));
  if (colsExport.length === 0) return;

  const rowsExport = linhasVisiveis();
  const nomeAba = NOMES_ABAS[activeTab] || '';

  // Linha 1 = totais (fórmulas SUM), linha 2 = cabeçalho, dados a partir da linha 3.
  const ws = XLSX.utils.aoa_to_sheet([[], colsExport]);

  const primeiraLinhaDados = 3;
  const ultimaLinhaDados = primeiraLinhaDados + rowsExport.length - 1;
  colsExport.forEach((col, c) => {
    if (!COLS_SOMATORIO.has(col)) return;
    const letra = XLSX.utils.encode_col(c);
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    // O `v` (soma pré-calculada) é obrigatório: o SheetJS descarta na escrita
    // qualquer célula que tenha só `f`, sem valor em cache.
    const soma = rowsExport.reduce((acc, r) => {
      const n = valorNumerico(r[col]);
      return isNaN(n) ? acc : acc + n;
    }, 0);
    const cell = { t: 'n', v: soma, z: 'R$ #,##0.00' };
    if (rowsExport.length) cell.f = `SUM(${letra}${primeiraLinhaDados}:${letra}${ultimaLinhaDados})`;
    ws[addr] = cell;
  });

  rowsExport.forEach((rowObj, i) => {
    const r = i + 2; // linha 0 = totais, linha 1 = cabeçalho
    colsExport.forEach((col, c) => {
      const addr = XLSX.utils.encode_cell({ r, c });
      // O SLA é sempre recalculado (dias entre "Envio Medição" e hoje, ou o valor
      // congelado no preenchimento do pedido) — o número gravado na planilha de
      // origem está desatualizado e ia pra exportação como 0/vazio.
      const val = col === 'Sla Envio Medição' ? slaAtual(rowObj) : rowObj[col];
      let cell;

      if (val === null || val === undefined || val === '') {
        cell = undefined; // deixa a célula vazia, igual ao padrão da planilha
      } else if (DATE_COLS.has(col)) {
        const d = (val instanceof Date) ? val : new Date(val);
        cell = isNaN(d.getTime()) ? { t: 's', v: String(val) }
                                  : { t: 'n', v: excelSerialFromDate(d), z: 'dd/mm/yyyy' };
      } else if (TEXT_COLS.has(col)) {
        cell = { t: 's', v: String(val) };
      } else if (CURRENCY_COLS.has(col)) {
        // Linhas antigas podem guardar o valor como texto formatado ("R$ 1 278,82").
        const n = valorNumerico(val);
        cell = isNaN(n) ? undefined : { t: 'n', v: n, z: 'R$ #,##0.00' };
      } else if (NUMERIC_CODE_COLS.has(col)) {
        // Exporta exatamente o texto da célula (inclusive os zeros à esquerda),
        // como string com formato texto — sem reconverter pra número/máscara.
        cell = { t: 's', v: String(val), z: '@' };
      } else if (col === 'Ano' || col === 'Sla Envio Medição') {
        const n = typeof val === 'number' ? val : parseFloat(String(val).replace(/\D/g, ''));
        cell = isNaN(n) ? { t: 's', v: String(val) } : { t: 'n', v: n };
      } else {
        cell = { t: 's', v: String(val) };
      }

      if (cell) ws[addr] = cell;
    });
  });

  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rowsExport.length + 1, c: colsExport.length - 1 } });
  ws['!cols'] = colsExport.map(col => ({ wch: LARGURAS_COLUNAS[COLUMNS.indexOf(col)] || 18 }));

  centralizarCelulas(ws);

  const wbOut = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbOut, ws, (nomeAba || sheetName).slice(0, 31));

  try {
    XLSX.writeFile(wbOut, `Medições Enviadas - ${nomeAba} - Teraup Soluções.xlsx`, { bookType: 'xlsx', cellStyles: true });
    formMsg.textContent = '';
    // O autosave não é limpo aqui: a exportação é parcial (só a aba atual),
    // então o rascunho local continua sendo a única cópia completa.
    fileStatus.textContent = `Planilha da aba "${nomeAba}" baixada com ${rowsExport.length} de ${rows.length} linha(s) e ${colsExport.length} de ${COLUMNS.length} coluna(s).`;
  } catch (err) {
    formMsg.textContent = 'Erro ao gerar o download: ' + err.message;
    formMsg.className = 'error';
    console.error(err);
  }
}
