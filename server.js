const express = require("express");
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");

const app = express();
const PORT = 3000;
const EXCEL_PATH = path.join(__dirname, "assets", "plan.xlsx");

app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// TESTE BACKEND
app.get("/api/ping", (req, res) => {
  res.json({ message: "backend funcionando!" });
});

// LER PLANILHA
app.get("/api/excel", (req, res) => {
  try {
    const workbook = XLSX.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    res.json({ data });
  } catch (err) {
    res.status(500).json({ error: "Erro ao ler Excel" });
  }
});

// SALVAR PLANILHA
app.post("/api/excel", (req, res) => {
  try {
    const { data } = req.body;

    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Planilha1");

    XLSX.writeFile(workbook, EXCEL_PATH);

    res.json({ status: "salvo" });
  } catch (err) {
    res.status(500).json({ error: "Erro ao salvar Excel" });
  }
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Servidor rodando em http://localhost:${PORT}`);
});
