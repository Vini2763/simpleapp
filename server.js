const express = require("express");
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");

const app = express();

// 🔴 PORTA: obrigatório pro Render
const PORT = process.env.PORT || 3000;

// caminho do Excel
const EXCEL_PATH = path.join(__dirname, "assets", "plan.xlsx");

app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// ========================
// TESTE BACKEND
// ========================
app.get("/api/ping", (req, res) => {
  res.json({ message: "backend funcionando!" });
});

// ========================
// GARANTIR QUE O EXCEL EXISTE
// ========================
function ensureExcelExists() {
  if (!fs.existsSync(EXCEL_PATH)) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ["Coluna A", "Coluna B", "Coluna C"],
      ["Exemplo 1", "Exemplo 2", "Exemplo 3"],
    ]);
    XLSX.utils.book_append_sheet(wb, ws, "Planilha1");
    XLSX.writeFile(wb, EXCEL_PATH);
  }
}

ensureExcelExists();

// ========================
// LER EXCEL
// ========================
app.get("/api/excel", (req, res) => {
  try {
    const workbook = XLSX.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    res.json({ data });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Erro ao ler Excel" });
  }
});

// ========================
// SALVAR EXCEL
// ========================
app.post("/api/excel", (req, res) => {
  try {
    const { data } = req.body;

    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Planilha1");

    XLSX.writeFile(workbook, EXCEL_PATH);

    res.json({ status: "salvo" });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Erro ao salvar Excel" });
  }
});

// ========================
// START
// ========================
app.listen(PORT, "0.0.0.0", () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});
