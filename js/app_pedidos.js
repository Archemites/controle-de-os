// Colunas exatas da planilha "Controle Pedidos de Compra.xlsx" (aba "PEDIDOS DE COMPRA"), na ordem.
// A 7ª coluna não tem cabeçalho no arquivo original; funciona como observação livre
// (ex: "NÃO FATURADO VIA SISTEMA").
const COLUMNS = ["Recebimento Pedido", "Número Pedido", "Valor", "O.S", "Cliente", "Gap", "Observação"];
const CURRENCY_COLS = new Set(["Valor"]);
const DATE_COLS = new Set(["Recebimento Pedido"]);

let workbook = null;
let sheetName = null;
let headerRowIndex = 0;
let rows = [];
let loadedFileName = null;
let originalFileBytes = null;
let manualRows = []; // pedidos adicionados manualmente pelo usuário (não vêm do .xlsx importado)
// Identidade (__linha) das linhas do arquivo original que estão na lixeira, e
// cópias das linhas do arquivo que foram editadas — ambas persistidas no autosave.
let linhasExcluidas = new Set();
let linhasEditadas = [];
// Lixeira: pedidos excluídos ficam aqui por 30 dias antes do descarte definitivo.
// Cada item é { linha: <objeto do pedido>, excluidoEm: <timestamp ms> }.
let lixeira = [];
const LIXEIRA_DIAS = 30;
const LIXEIRA_MS = LIXEIRA_DIAS * 86400000;
let abaAtiva = 'pedidos'; // 'pedidos' | 'lixeira'

const fileInput = document.getElementById('fileInput');
const openBtn = document.getElementById('openBtn');
const fileStatus = document.getElementById('fileStatus');
const emptyState = document.getElementById('emptyState');
const dataTable = document.getElementById('dataTable');
const totalsRow = document.getElementById('totalsRow');
const headerRow = document.getElementById('headerRow');
const tableBody = document.getElementById('tableBody');
const rowCount = document.getElementById('rowCount');
const searchBar = document.getElementById('searchBar');
// Declarado aqui (e não junto do bloco de exportação) porque renderTable() já o
// usa para desabilitar o download enquanto a aba "Lixeira" está aberta.
const downloadBtn = document.getElementById('downloadBtn');
const searchInput = document.getElementById('searchInput');

let searchTerm = '';
let searchFields = new Set(['O.S', 'Número Pedido', 'Gap', 'Cliente']);
const filterChips = document.querySelectorAll('.filter-chip');

// --- Tema claro/escuro (mesma chave usada pelo Controle de Medições, para
// manter o tema sincronizado entre as duas abas do app) ----------------------
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

filterChips.forEach(chip => {
  chip.addEventListener('click', () => {
    const field = chip.dataset.field;
    if (searchFields.has(field)) {
      if (searchFields.size === 1) return; // mantém ao menos um filtro ativo
      searchFields.delete(field);
      chip.classList.remove('active');
    } else {
      searchFields.add(field);
      chip.classList.add('active');
    }
    renderTable();
  });
});

searchInput.addEventListener('input', () => {
  searchTerm = searchInput.value.trim().toLowerCase();
  renderTable();
});

// --- Autosave (IndexedDB) ---------------------------------------------------
// Guarda os bytes do .xlsx aberto para restaurar automaticamente a sessão em
// um refresh, do mesmo jeito que o Controle de Medições faz (banco próprio,
// não compartilhado, já que é outra planilha).
const AUTOSAVE_DB = 'pedidos_autosave';
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

async function salvarAutosave() {
  if (!workbook || !originalFileBytes) return;
  try {
    const db = await abrirAutosaveDB();
    const estado = {
      fileName: loadedFileName, fileBytes: originalFileBytes, manualRows,
      linhasExcluidas: [...linhasExcluidas], linhasEditadas, lixeira,
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

// --- Leitura da planilha -----------------------------------------------------
// O.S e Número Pedido vêm com tipos de célula inconsistentes no arquivo original
// (ora texto, ora número) — sempre convertidos para string ao ler, preservando
// o padding de zeros já embutido como texto literal (ao contrário da planilha
// de Medições, aqui não há number_format a recuperar).
function valorTexto(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

// Mesma regra do Controle de Medições: O.S numérica sempre com 10 dígitos
// (zeros à esquerda), valores mistos como "AMP-4866/0030871575" intactos.
function normalizarOs(v) {
  if (v === null || v === undefined) return v;
  const s = String(v).trim();
  return /^\d+$/.test(s) ? s.padStart(10, '0') : v;
}

// `__linha` é a posição da linha dentro das linhas de dados do arquivo, e serve de
// identidade estável para reaplicar edições e exclusões depois que o autosave
// relê o .xlsx original do zero (mesma ideia do __linha em app_medicoes.html).
function construirLinhas(json) {
  return json
    .map((r, idx) => {
      if (!r.some(c => c !== undefined && c !== null && c !== '')) return null;
      const obj = { __linha: idx };
      COLUMNS.forEach((col, i) => {
        let val = r[i];
        if (val === undefined) val = null;
        if (col === 'O.S' || col === 'Número Pedido' || col === 'Cliente' || col === 'Gap' || col === 'Observação') {
          val = val === null ? null : valorTexto(val).toUpperCase();
          if (val === '') val = null;
          if (col === 'O.S') val = normalizarOs(val);
        }
        obj[col] = val;
      });
      return obj;
    })
    .filter(obj => obj && !linhaVazia(obj));
}

// Registro sem nenhum campo preenchido — sempre descartado, na importação e
// antes de exibir/exportar. Também descarta linhas de subtotal que a planilha
// intercala entre os dados (só "Valor" preenchido, sem Número Pedido/O.S/Cliente
// que identifiquem um pedido de fato).
function linhaVazia(row) {
  if (!row) return true;
  if (COLUMNS.every(col => {
    const v = row[col];
    return v === null || v === undefined || String(v).trim() === '';
  })) return true;
  const semIdentificacao = !row['Número Pedido'] && !row['O.S'] && !row['Cliente'];
  return semIdentificacao;
}

function carregarWorkbookDeBytes(bytes, fileName, excluidas, editadas, lixeiraSalva) {
  const wb = XLSX.read(bytes, { type: 'array', cellDates: true });
  const wsName = wb.SheetNames[0];
  const ws = wb.Sheets[wsName];
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let hIdx = json.findIndex(r => valorTexto(r[3]) === 'O.S' || valorTexto(r[1]) === 'Número Pedido');
  if (hIdx === -1) hIdx = 2; // fallback: linha 3 (0-based 2), como no arquivo original

  workbook = wb;
  sheetName = wsName;
  headerRowIndex = hIdx;
  loadedFileName = fileName;
  originalFileBytes = bytes;
  manualRows = [];
  // Exclusões/edições restauradas do autosave: sem isso, uma linha do arquivo
  // apagada ou alterada voltaria ao estado original a cada recarregamento.
  linhasExcluidas = new Set(Array.isArray(excluidas) ? excluidas : []);
  linhasEditadas = Array.isArray(editadas) ? editadas : [];
  lixeira = Array.isArray(lixeiraSalva) ? lixeiraSalva.filter(i => i && i.linha) : [];
  purgarLixeira(); // descarta o que já passou dos 30 dias
  rows = construirLinhas(json.slice(hIdx + 1)).filter(r => !linhasExcluidas.has(r.__linha));
  linhasEditadas.forEach(e => {
    const pos = rows.findIndex(o => o.__linha === e.__linha);
    if (pos !== -1) rows[pos] = e;
  });
}

fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    carregarWorkbookDeBytes(data, file.name);
    fileStatus.textContent = `Planilha carregada: "${file.name}" — ${rows.length} pedidos.`;
    renderTable();
    salvarAutosave();
  };
  reader.readAsArrayBuffer(file);
  fileInput.value = '';
});

openBtn.addEventListener('click', () => fileInput.click());

// --- Adicionar pedido manualmente ------------------------------------------
// Este viewer não grava de volta no .xlsx (só leitura/importação), então
// pedidos adicionados manualmente existem apenas em memória + autosave
// (IndexedDB), reaparecendo ao recarregar a página mas sem alterar o arquivo original.
const addBtn = document.getElementById('addBtn');
const addModalOverlay = document.getElementById('addModalOverlay');
const addFormMsg = document.getElementById('addFormMsg');

const addPanelTitle = document.getElementById('addPanelTitle');
const addSaveBtn = document.getElementById('addSaveBtn');
let editingRow = null; // linha (objeto em `rows`) sendo editada, ou null ao cadastrar

const CAMPOS_FORM = ['f_recebimento', 'f_numeroPedido', 'f_valor', 'f_osPedido', 'f_cliente', 'f_gapPedido', 'f_observacao'];

function abrirModalAdicionar() {
  if (!workbook) {
    alert('Importe uma planilha primeiro clicando em "Importar Pedidos".');
    return;
  }
  editingRow = null;
  addPanelTitle.textContent = 'Adicionar Pedido';
  addSaveBtn.textContent = 'Salvar';
  CAMPOS_FORM.forEach(id => document.getElementById(id).value = '');
  addFormMsg.textContent = '';
  addFormMsg.className = '';
  addModalOverlay.classList.add('open');
  document.getElementById('f_recebimento').focus();
}

function toDateInputValue(v) {
  if (!v) return '';
  const d = v instanceof Date ? v : new Date(v);
  if (isNaN(d.getTime())) return '';
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${d.getFullYear()}-${mm}-${dd}`;
}

function abrirModalEdicao(row) {
  editingRow = row;
  addPanelTitle.textContent = 'Editar Pedido';
  addSaveBtn.textContent = 'Salvar Alterações';
  document.getElementById('f_recebimento').value = toDateInputValue(row['Recebimento Pedido']);
  document.getElementById('f_numeroPedido').value = row['Número Pedido'] ?? '';
  const valor = row['Valor'];
  document.getElementById('f_valor').value = (valor === null || valor === undefined || valor === '') ? '' : String(valor).replace('.', ',');
  document.getElementById('f_osPedido').value = row['O.S'] ?? '';
  document.getElementById('f_cliente').value = row['Cliente'] ?? '';
  document.getElementById('f_gapPedido').value = row['Gap'] ?? '';
  document.getElementById('f_observacao').value = row['Observação'] ?? '';
  addFormMsg.textContent = '';
  addFormMsg.className = '';
  addModalOverlay.classList.add('open');
}

// --- Lixeira -----------------------------------------------------------------
// Excluir NÃO apaga o pedido: move para a aba "Lixeira", onde ele fica por
// LIXEIRA_DIAS dias antes do descarte definitivo. Enquanto está lá, o registro
// sai de `rows`/`manualRows` — não aparece na lista, nos totais nem na planilha
// baixada. Linhas vindas do .xlsx (com __linha) também entram em
// `linhasExcluidas`, senão voltariam quando o autosave relesse o arquivo.
function rotuloPedido(row) {
  return row['Número Pedido'] ? `o pedido nº ${row['Número Pedido']}` : 'este registro';
}

function diasRestantesLixeira(item) {
  return Math.max(0, Math.ceil((item.excluidoEm + LIXEIRA_MS - Date.now()) / 86400000));
}

// Descarta definitivamente o que já passou dos 30 dias. Chamado na restauração
// do autosave e a cada renderização, para o prazo valer mesmo sem interação.
function purgarLixeira() {
  const limite = Date.now() - LIXEIRA_MS;
  lixeira = lixeira.filter(item => item && item.excluidoEm > limite);
}

function excluirPedido(row) {
  const rotulo = rotuloPedido(row);
  if (!window.confirm(`Mover ${rotulo} para a lixeira?\n\nO registro sai da lista e das planilhas baixadas, mas fica ${LIXEIRA_DIAS} dias na aba "Lixeira", de onde pode ser restaurado. Depois disso é apagado por completo.`)) return;

  if (row.__linha !== undefined) linhasExcluidas.add(row.__linha);
  rows = rows.filter(r => r !== row);
  manualRows = manualRows.filter(r => r !== row);
  linhasEditadas = linhasEditadas.filter(r => r !== row);
  lixeira.push({ linha: row, excluidoEm: Date.now() });
  if (editingRow === row) fecharModalAdicionar();
  renderTable();
  salvarAutosave();
  fileStatus.textContent = `${rotulo} movido para a lixeira — será apagado definitivamente em ${LIXEIRA_DIAS} dias.`;
}

function restaurarDaLixeira(item) {
  const row = item.linha;
  if (row.__linha !== undefined) {
    linhasExcluidas.delete(row.__linha);
    // Voltou de uma linha do arquivo: registra como editada para o autosave
    // reaplicá-la sobre o .xlsx relido (inclusive se tinha sido alterada antes).
    if (!linhasEditadas.includes(row)) linhasEditadas.push(row);
  } else if (!manualRows.includes(row)) {
    manualRows.push(row);
  }
  lixeira = lixeira.filter(i => i !== item);
  rows.push(row);
  renderTable();
  salvarAutosave();
  fileStatus.textContent = `${rotuloPedido(row)} restaurado da lixeira.`;
}

function excluirDefinitivo(item) {
  const rotulo = rotuloPedido(item.linha);
  if (!window.confirm(`Apagar ${rotulo} definitivamente?\n\nEsta ação não pode ser desfeita.`)) return;
  lixeira = lixeira.filter(i => i !== item);
  renderTable();
  salvarAutosave();
  fileStatus.textContent = `${rotulo} apagado definitivamente.`;
}

function fecharModalAdicionar() {
  addModalOverlay.classList.remove('open');
  editingRow = null;
}

addBtn.addEventListener('click', abrirModalAdicionar);
document.getElementById('addCancelBtn').addEventListener('click', fecharModalAdicionar);
addModalOverlay.addEventListener('click', (e) => {
  if (e.target === addModalOverlay) fecharModalAdicionar();
});

document.getElementById('addSaveBtn').addEventListener('click', () => {
  const recebimento = document.getElementById('f_recebimento').value;
  const numeroPedido = document.getElementById('f_numeroPedido').value.trim().toUpperCase();
  const valorStr = document.getElementById('f_valor').value.trim();
  const os = normalizarOs(document.getElementById('f_osPedido').value.trim().toUpperCase());
  const cliente = document.getElementById('f_cliente').value.trim().toUpperCase();
  const gap = document.getElementById('f_gapPedido').value.trim().toUpperCase();
  const observacao = document.getElementById('f_observacao').value.trim().toUpperCase();

  if (!numeroPedido) {
    addFormMsg.textContent = 'Preencha o campo Número Pedido.';
    addFormMsg.className = 'error';
    return;
  }

  const novaLinha = {
    "Recebimento Pedido": recebimento ? new Date(recebimento + 'T00:00:00') : null,
    "Número Pedido": numeroPedido,
    "Valor": valorStr ? valorNumerico(valorStr) : null,
    "O.S": os || null,
    "Cliente": cliente || null,
    "Gap": gap || null,
    "Observação": observacao || null,
  };

  if (editingRow) {
    Object.assign(editingRow, novaLinha);
    // Edição de uma linha do arquivo original: guarda a cópia alterada para o
    // autosave reaplicá-la sobre o .xlsx relido. Linhas manuais já são salvas
    // inteiras em `manualRows`.
    if (editingRow.__linha !== undefined && !linhasEditadas.includes(editingRow)) {
      linhasEditadas.push(editingRow);
    }
  } else {
    rows.push(novaLinha);
    manualRows.push(novaLinha);
  }
  renderTable();
  salvarAutosave();
  fecharModalAdicionar();
});

// ---------------------------------------------------------------------------
// Leitura de PDF
// Extrai o texto do PDF via pdf.js (100% no navegador, nenhum upload sai da
// máquina) e procura, por regex, os campos que aparecem nas ordens/notas:
// O.S, GAP, Pedido de Compra (= Nº do Pedido), Valor Faturado (= Valor Líquido),
// Data e Cliente.
// O resultado é só uma leitura assistida — nada é gravado sem o usuário mandar;
// o botão "Adicionar como novo pedido" cria a linha na lista de pedidos.
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

// Procura na lista de pedidos carregada a O.S correspondente a uma GAP lida do
// PDF. Serve tanto para preencher a O.S quando o PDF não a traz quanto para
// sinalizar divergência entre a O.S do PDF e a da planilha.
function osPelaGap(gap) {
  const alvo = String(gap || '').replace(/\D/g, '');
  if (!alvo) return '';
  const achada = rows.find(r => String(r['Gap'] || '').replace(/\D/g, '') === alvo && String(r['O.S'] || '').trim() !== '');
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
    pdfUsarBtn.disabled = achados === 0;
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

// Ação principal da leitura: o Controle de Pedidos de Compra sempre ganha um
// REGISTRO NOVO com o que o PDF trouxe (é no Controle de Medições que a leitura
// complementa a linha não faturada de mesma O.S). O pedido entra em `rows`
// (exibição) e em `manualRows` (persistido no autosave), sem tocar no .xlsx.
pdfUsarBtn.addEventListener('click', () => {
  if (!pdfCamposLidos) return;
  if (!workbook) {
    pdfPanelText.textContent = 'Importe uma planilha primeiro clicando em "Importar Pedidos".';
    return;
  }

  const data = pdfCamposLidos['Data'];
  let recebimento = null;
  if (data) {
    const [d, m, a] = data.split('/');
    recebimento = new Date(`${a}-${m}-${d}T00:00:00`);
    if (isNaN(recebimento.getTime())) recebimento = null;
  }
  const valor = valorNumerico(pdfCamposLidos['Valor Faturado']);
  const novaLinha = {
    "Recebimento Pedido": recebimento,
    "Número Pedido": pdfCamposLidos['Pedido de Compra'] ? String(pdfCamposLidos['Pedido de Compra']).toUpperCase() : null,
    "Valor": isNaN(valor) ? null : valor,
    "O.S": pdfCamposLidos['O.S'] ? normalizarOs(String(pdfCamposLidos['O.S']).toUpperCase()) : null,
    "Cliente": pdfCamposLidos['Cliente'] ? String(pdfCamposLidos['Cliente']).toUpperCase() : null,
    "Gap": pdfCamposLidos['GAP'] ? String(pdfCamposLidos['GAP']).toUpperCase() : null,
    "Observação": `Importado do PDF ${pdfNomeArquivo}`.toUpperCase(),
  };
  // Duplicata: mesmo Nº do Pedido ou mesma O.S já na lista (comparando só os
  // dígitos). Avisa e deixa o usuário decidir se mantém os dois.
  const digitos = v => String(v || '').replace(/\D/g, '');
  const dup = rows.find(p =>
    (digitos(novaLinha['Número Pedido']) && digitos(p['Número Pedido']) === digitos(novaLinha['Número Pedido'])) ||
    (digitos(novaLinha['O.S']) && digitos(p['O.S']) === digitos(novaLinha['O.S']))
  );
  if (dup) {
    const ok = confirm(
      `Já existe um pedido com Nº ${dup['Número Pedido'] || '-'} / O.S ${dup['O.S'] || '-'} na lista.\n\n` +
      'Adicionar mesmo assim (mantendo os dois)?'
    );
    if (!ok) {
      fecharLeituraPdf();
      fileStatus.textContent = `Pedido de "${pdfNomeArquivo}" não adicionado: duplicata do Nº ${dup['Número Pedido'] || '-'} / O.S ${dup['O.S'] || '-'}.`;
      return;
    }
  }

  rows.push(novaLinha);
  manualRows.push(novaLinha);
  renderTable();
  salvarAutosave();
  fecharLeituraPdf();
  fileStatus.textContent = `Pedido adicionado a partir de "${pdfNomeArquivo}"${dup ? ' (duplicata confirmada por você)' : ''} — ${rows.length} pedidos na lista.`;
});

(async function restaurarAutosaveNaAbertura() {
  const estado = await carregarAutosave();
  if (!estado || !estado.fileBytes) return;
  try {
    carregarWorkbookDeBytes(estado.fileBytes, estado.fileName, estado.linhasExcluidas || [], estado.linhasEditadas || [], estado.lixeira || []);
    if (Array.isArray(estado.manualRows) && estado.manualRows.length) {
      manualRows = estado.manualRows;
      rows.push(...manualRows);
    }
    const dataSalva = new Date(estado.savedAt).toLocaleString('pt-BR');
    fileStatus.textContent = `Sessão restaurada automaticamente: "${estado.fileName}" — ${rows.length} pedidos (autosave de ${dataSalva}).`;
    renderTable();
  } catch (err) {
    console.error('Falha ao restaurar autosave:', err);
  }
})();

// --- Formatação ---------------------------------------------------------------
function valorNumerico(v) {
  if (v === null || v === undefined || v === '') return NaN;
  if (typeof v === 'number') return v;
  const bruto = String(v).replace(/R\$/gi, '').trim();
  if (/^-+$/.test(bruto)) return 0;
  return parseFloat(bruto.replace(/[\s .]/g, '').replace(',', '.'));
}

function fmtCurrency(v) {
  if (v === null || v === undefined || v === '') return '';
  const n = valorNumerico(v);
  if (isNaN(n)) return String(v);
  if (n === 0) return '';
  return n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function fmtDate(v) {
  if (!v) return '';
  const d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return String(v);
  return d.toLocaleDateString('pt-BR');
}

function matchesSearch(r) {
  if (!searchTerm) return true;
  return [...searchFields].some(field => {
    const v = r[field];
    return v !== null && v !== undefined && String(v).toLowerCase().includes(searchTerm);
  });
}

// --- Ordenação (clique no cabeçalho alterna crescente → decrescente → sem ordenação) ---
let sortCol = null;
let sortDir = 1;

function sortKey(r, col) {
  const val = r[col];
  if (val === null || val === undefined || val === '') return null;
  if (DATE_COLS.has(col)) {
    const d = (val instanceof Date) ? val : new Date(val);
    return isNaN(d.getTime()) ? null : d.getTime();
  }
  if (CURRENCY_COLS.has(col)) {
    const n = valorNumerico(val);
    return isNaN(n) ? null : n;
  }
  if (col === 'O.S' || col === 'Número Pedido') {
    const n = parseFloat(String(val).replace(/\D/g, ''));
    return isNaN(n) ? String(val).toLowerCase() : n;
  }
  return String(val).toLowerCase();
}

function compareRows(a, b, col) {
  const ka = sortKey(a, col), kb = sortKey(b, col);
  if (ka === null && kb === null) return 0;
  if (ka === null) return 1;
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

// Ícones da coluna "Ações" — SVG inline (traço em currentColor, então herdam a
// cor do botão: neutra no lápis, vermelha na lixeira, em ambos os temas).
const ICONE_LAPIS = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4z"/></svg>';
const ICONE_LIXEIRA = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/></svg>';
const ICONE_RESTAURAR = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 12a9 9 0 1 0 3-6.7"/><path d="M3 4v5h5"/></svg>';

const subTabs = document.getElementById('subTabs');
const tabPedidos = document.getElementById('tabPedidos');
const tabLixeira = document.getElementById('tabLixeira');
const pedidosCount = document.getElementById('pedidosCount');
const lixeiraCount = document.getElementById('lixeiraCount');

function setActiveTab(aba) {
  abaAtiva = aba;
  tabPedidos.classList.toggle('active', aba === 'pedidos');
  tabLixeira.classList.toggle('active', aba === 'lixeira');
  renderTable();
}

tabPedidos.addEventListener('click', () => setActiveTab('pedidos'));
tabLixeira.addEventListener('click', () => setActiveTab('lixeira'));

function renderTable() {
  purgarLixeira(); // o prazo de 30 dias vale mesmo sem o usuário mexer em nada
  const naLixeira = abaAtiva === 'lixeira';
  headerRow.innerHTML = COLUMNS.map(c => {
    const isSorted = sortCol === c;
    const ind = isSorted ? (sortDir === 1 ? '▲' : '▼') : '⇅';
    const label = c === 'Observação' ? 'Observação' : c;
    return `<th class="sortable${isSorted ? ' sorted' : ''}" data-col="${c}" title="Ordenar por ${label}">${label}<span class="sort-ind">${ind}</span></th>`;
  }).join('') + (naLixeira ? '<th class="acoes-col">Excluído em</th><th class="acoes-col">Ações</th>' : '<th class="acoes-col">Ações</th>');
  headerRow.querySelectorAll('th.sortable').forEach(th => {
    th.addEventListener('click', () => toggleSort(th.dataset.col));
  });

  if (rows.some(linhaVazia)) {
    rows = rows.filter(r => !linhaVazia(r));
    manualRows = manualRows.filter(r => !linhaVazia(r));
  }
  const visibleRows = naLixeira ? [] : rows.filter(matchesSearch);
  if (sortCol) visibleRows.sort((a, b) => compareRows(a, b, sortCol));

  totalsRow.style.display = naLixeira ? 'none' : '';
  totalsRow.innerHTML = COLUMNS.map(col => {
    if (col !== 'Valor') return '<td></td>';
    const soma = visibleRows.reduce((acc, r) => {
      const n = valorNumerico(r[col]);
      return isNaN(n) ? acc : acc + n;
    }, 0);
    return `<td>${fmtCurrency(soma)}</td>`;
  }).join('') + (naLixeira ? '<td></td><td></td>' : '<td></td>');

  tableBody.innerHTML = '';
  if (naLixeira) renderLixeira();
  else visibleRows.forEach(r => {
    const tr = document.createElement('tr');
    tr.className = 'editable-row';
    tr.title = 'Clique duas vezes para editar';
    tr.addEventListener('dblclick', () => abrirModalEdicao(r));
    tr.innerHTML = COLUMNS.map(col => {
      let val = r[col];
      if (DATE_COLS.has(col)) val = fmtDate(val);
      else if (CURRENCY_COLS.has(col)) val = fmtCurrency(val);
      else if (val === null || val === undefined) val = '';
      // Célula pode vir de uma planilha importada ou de digitação manual — nunca
      // confiável o bastante pra entrar em innerHTML sem escapar (ver TESTE_DE_FOGO.md).
      return `<td>${escapeHtml(val)}</td>`;
    }).join('') +
      '<td class="acoes-cell">' +
      `<button type="button" class="row-btn editar" title="Editar registro" aria-label="Editar registro">${ICONE_LAPIS}</button>` +
      `<button type="button" class="row-btn excluir" title="Excluir registro" aria-label="Excluir registro">${ICONE_LIXEIRA}</button>` +
      '</td>';
    tr.querySelector('.row-btn.editar').addEventListener('click', ev => { ev.stopPropagation(); abrirModalEdicao(r); });
    tr.querySelector('.row-btn.excluir').addEventListener('click', ev => { ev.stopPropagation(); excluirPedido(r); });
    tableBody.appendChild(tr);
  });

  const temConteudo = rows.length || lixeira.length;
  subTabs.style.display = temConteudo ? 'flex' : 'none';
  searchBar.style.display = temConteudo && !naLixeira ? 'block' : 'none';
  emptyState.style.display = temConteudo ? 'none' : 'block';
  dataTable.classList.remove('initially-hidden');
  dataTable.style.display = temConteudo ? '' : 'none';
  rowCount.textContent = !temConteudo ? ''
    : naLixeira ? `${lixeira.length} registro(s) na lixeira`
    : `${visibleRows.length} de ${rows.length} pedidos`;
  pedidosCount.textContent = `(${rows.length})`;
  lixeiraCount.textContent = `(${lixeira.length})`;

  // A lixeira não é exportável: são registros fora da lista, aguardando descarte.
  downloadBtn.disabled = !workbook || naLixeira;
  downloadBtn.title = naLixeira
    ? 'A aba "Lixeira" não é exportada — restaure os registros para incluí-los na planilha'
    : '';
}

// Corpo da aba "Lixeira": os mesmos campos da lista, mais a data de exclusão e o
// prazo restante, com ações próprias (restaurar ou apagar de vez). Estes
// registros não estão em `rows`, então não entram em busca, totais nem export.
function renderLixeira() {
  const itens = [...lixeira].sort((a, b) => b.excluidoEm - a.excluidoEm);
  if (!itens.length) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="${COLUMNS.length + 2}" style="text-align:center; padding:18px; opacity:0.7;">Lixeira vazia. Registros excluídos ficam aqui por ${LIXEIRA_DIAS} dias.</td>`;
    tableBody.appendChild(tr);
    return;
  }
  itens.forEach(item => {
    const r = item.linha;
    const tr = document.createElement('tr');
    tr.innerHTML = COLUMNS.map(col => {
      let val = r[col];
      if (DATE_COLS.has(col)) val = fmtDate(val);
      else if (CURRENCY_COLS.has(col)) val = fmtCurrency(val);
      else if (val === null || val === undefined) val = '';
      return `<td>${escapeHtml(val)}</td>`;
    }).join('') +
      `<td>${new Date(item.excluidoEm).toLocaleDateString('pt-BR')} · ${diasRestantesLixeira(item)} dia(s)</td>` +
      '<td class="acoes-cell">' +
      `<button type="button" class="row-btn restaurar" title="Restaurar registro" aria-label="Restaurar registro">${ICONE_RESTAURAR}</button>` +
      `<button type="button" class="row-btn excluir" title="Apagar definitivamente" aria-label="Apagar definitivamente">${ICONE_LIXEIRA}</button>` +
      '</td>';
    tr.querySelector('.row-btn.restaurar').addEventListener('click', () => restaurarDaLixeira(item));
    tr.querySelector('.row-btn.excluir').addEventListener('click', () => excluirDefinitivo(item));
    tableBody.appendChild(tr);
  });
}

renderTable();

// --- Baixar planilha "Controle de Pedidos de Compra.xlsx" -------------------
// Gera uma planilha nova (não é o mesmo objeto workbook importado) contendo
// todos os `rows` atualmente carregados (linhas do arquivo original + linhas
// adicionadas manualmente), na mesma formatação visual usada no restante do
// app: cabeçalho azul, data dd/mm/aaaa, moeda "R$ #,##0.00", O.S como texto
// (preserva o zero-padding literal já embutido na string).

function excelSerialFromDate(date) {
  const epoch = Date.UTC(1899, 11, 30);
  return (Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - epoch) / 86400000;
}

// Aplica alinhamento centralizado (horizontal e vertical) em todas as células
// dentro do !ref da planilha. Mantém qualquer estilo já existente na célula.
// Só estiliza células que existem no objeto da planilha (ver nota em app_medicoes.html).
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

function baixarPlanilhaPedidos() {
  if (!rows.length) return;

  // Linha de totais fica logo após o cabeçalho (linha 2, 1-based) e antes dos
  // registros, que começam na linha 3. Colunas sem soma (não-CURRENCY_COLS)
  // ficam em branco nessa linha.
  const totalsRowAoa = COLUMNS.map(() => null);
  const aoa = [COLUMNS, totalsRowAoa, ...rows.map(r => COLUMNS.map(col => {
    const val = r[col];
    if (val === null || val === undefined) return null;
    if (DATE_COLS.has(col)) return (val instanceof Date) ? val : new Date(val);
    if (CURRENCY_COLS.has(col)) return valorNumerico(val);
    return String(val);
  }))];

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  const firstDataRow = 3; // 1-based: linha 1 = cabeçalho, linha 2 = totais, linha 3 = primeiro registro
  const lastDataRow = firstDataRow + rows.length - 1;

  rows.forEach((r, i) => {
    const rIdx = i + 2; // +2: linha 0 = cabeçalho, linha 1 = totais, dados a partir da linha 2 (0-based)
    COLUMNS.forEach((col, cIdx) => {
      const addr = XLSX.utils.encode_cell({ r: rIdx, c: cIdx });
      const cell = ws[addr];
      if (!cell) return;
      if (DATE_COLS.has(col)) {
        const d = r[col] instanceof Date ? r[col] : new Date(r[col]);
        if (!isNaN(d.getTime())) { cell.t = 'n'; cell.v = excelSerialFromDate(d); cell.z = 'dd/mm/yyyy'; }
      } else if (CURRENCY_COLS.has(col)) {
        cell.z = 'R$ #,##0.00';
      }
    });
  });

  // Fórmulas de soma na linha de totais (linha 2, 0-based r:1), uma por coluna de moeda.
  COLUMNS.forEach((col, cIdx) => {
    if (!CURRENCY_COLS.has(col)) return;
    const colLetter = XLSX.utils.encode_col(cIdx);
    const addr = XLSX.utils.encode_cell({ r: 1, c: cIdx });
    // O `v` (soma pré-calculada) é obrigatório: o SheetJS descarta na escrita
    // qualquer célula que tenha só `f`, sem valor em cache.
    const soma = rows.reduce((acc, r) => {
      const n = valorNumerico(r[col]);
      return isNaN(n) ? acc : acc + n;
    }, 0);
    const cell = { t: 'n', v: soma, z: 'R$ #,##0.00' };
    if (rows.length) cell.f = `SUM(${colLetter}${firstDataRow}:${colLetter}${lastDataRow})`;
    ws[addr] = cell;
  });

  ws['!cols'] = [
    { wch: 18 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 32 }, { wch: 14 }, { wch: 28 }
  ];

  centralizarCelulas(ws);

  const wbOut = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbOut, ws, 'PEDIDOS DE COMPRA');

  try {
    XLSX.writeFile(wbOut, 'Controle de Pedidos de Compra.xlsx', { bookType: 'xlsx', cellStyles: true });
    fileStatus.textContent = `Planilha baixada com ${rows.length} pedido(s).`;
  } catch (err) {
    fileStatus.textContent = 'Erro ao gerar o download: ' + err.message;
    console.error(err);
  }
}

downloadBtn.addEventListener('click', baixarPlanilhaPedidos);
