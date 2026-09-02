// Replicates app_medicoes.html logic: read original workbook, add one test row via the SAME
// cell-writing logic as baixarPlanilha(), and write output file with XLSX.writeFile.
const XLSX = require("xlsx");
const path = require("path");

const SRC = path.join(__dirname, "xlsx", "Medições enviadas - Teraup Soluções.xlsx");
const OUT = path.join(__dirname, "output_simulado.xlsx");

const COLUMNS = [
  "O.S", "Mês Envio Medição", "Vistoria Projeto", "Ano", "GAP", "Envio Medição",
  "Sla Envio Medição", "Valor Medição Enviadas", "Engenheiro Acesso", "Projetista",
  "Regional", "Pedido de Compra", "Valor Faturado", "Medição Terceiros",
  "Custos Diversos", "Custo ART", "Custo Total", "Cobrança Medição"
];
const CURRENCY_COLS = new Set([
  "Valor Medição Enviadas", "Valor Faturado", "Medição Terceiros", "Custos Diversos", "Custo ART", "Custo Total"
]);
const DATE_COLS = new Set(["Vistoria Projeto", "Envio Medição"]);
const TEXT_COLS = new Set(["O.S", "GAP", "Pedido de Compra"]);

// --- Mimic file load (fileInput change handler) ---
const buf = require("fs").readFileSync(SRC);
const workbook = XLSX.read(buf, { type: "buffer", cellDates: true });
const sheetName = workbook.SheetNames[0];
const ws = workbook.Sheets[sheetName];
const json = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

let headerRowIndex = json.findIndex(r => r && r[0] === "O.S");
if (headerRowIndex === -1) headerRowIndex = 1;

let lastOriginalRow = headerRowIndex;
for (let i = headerRowIndex + 1; i < json.length; i++) {
  const r = json[i];
  if (!r || r.every(c => c === null || c === undefined || c === "")) continue;
  lastOriginalRow = i;
}
console.log("headerRowIndex (0-based):", headerRowIndex, "-> Excel row", headerRowIndex + 1);
console.log("lastOriginalRow (0-based):", lastOriginalRow, "-> Excel row", lastOriginalRow + 1);

// --- Mimic salvarMedicao() building one new row object ---
function excelSerialFromDate(date) {
  const epoch = Date.UTC(1899, 11, 30);
  return (Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) - epoch) / 86400000;
}
function diasDesde(dataStr) {
  if (!dataStr) return null;
  const inicio = new Date(dataStr + "T00:00:00");
  const hoje = new Date();
  hoje.setHours(0,0,0,0);
  return Math.floor((hoje - inicio) / 86400000);
}

const envioStr = "2026-07-28";
const novaLinha = {
  "O.S": "19999999",
  "Mês Envio Medição": "Julho",
  "Vistoria Projeto": new Date("2026-07-20T00:00:00"),
  "Ano": 2026,
  "GAP": "GAP-99999",
  "Envio Medição": new Date(envioStr + "T00:00:00"),
  "Sla Envio Medição": diasDesde(envioStr),
  "Valor Medição Enviadas": 500,
  "Engenheiro Acesso": "TESTE ENGENHEIRO",
  "Projetista": "TESTE PROJETISTA",
  "Regional": "TESTE REGIONAL",
  "Pedido de Compra": "4500999999",
  "Valor Faturado": 500,
  "Medição Terceiros": 100,
  "Custos Diversos": 50,
  "Custo ART": 50,
  "Custo Total": 100,
  "Cobrança Medição": null,
};
const newRows = [novaLinha];

// --- Mimic baixarPlanilha() exactly ---
const refRow = lastOriginalRow;
let rowCursor = lastOriginalRow + 1;

newRows.forEach(rowObj => {
  COLUMNS.forEach((col, colIdx) => {
    const addr = XLSX.utils.encode_cell({ r: rowCursor, c: colIdx });
    const refAddr = XLSX.utils.encode_cell({ r: refRow, c: colIdx });
    const refCell = ws[refAddr];
    let val = rowObj[col];
    let cell;

    if (val === null || val === undefined || val === "") {
      cell = undefined;
    } else if (DATE_COLS.has(col)) {
      const d = (val instanceof Date) ? val : new Date(val);
      cell = { t: "n", v: excelSerialFromDate(d), z: (refCell && refCell.z) || "mm-dd-yy" };
    } else if (TEXT_COLS.has(col)) {
      cell = { t: "s", v: String(val) };
    } else if (col === "Ano" || CURRENCY_COLS.has(col) || col === "Sla Envio Medição") {
      const n = typeof val === "number" ? val : parseFloat(val);
      cell = { t: "n", v: isNaN(n) ? 0 : n };
      if (CURRENCY_COLS.has(col)) cell.z = "R$ #,##0.00";
      else if (refCell && refCell.z) cell.z = refCell.z;
    } else {
      cell = { t: "s", v: String(val) };
    }

    if (cell) {
      ws[addr] = cell;
    } else {
      delete ws[addr];
    }
  });
  rowCursor++;
});

const range = XLSX.utils.decode_range(ws["!ref"]);
range.e.r = Math.max(range.e.r, rowCursor - 1);
ws["!ref"] = XLSX.utils.encode_range(range);

XLSX.writeFile(workbook, OUT, { bookType: "xlsx" });
console.log("Wrote:", OUT);
console.log("New row written at Excel row:", refRow + 1 + 1); // rowCursor start = lastOriginalRow+1 (0-based) -> +1 for 1-based
