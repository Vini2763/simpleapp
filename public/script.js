const label = document.getElementById("label");
const excelContainer = document.getElementById("excelContainer");

document.getElementById("btnPing").onclick = async () => {
  const res = await fetch("/api/ping");
  const data = await res.json();
  label.textContent = data.message;
};

// CARREGAR EXCEL
async function loadExcel() {
  const res = await fetch("/api/excel");
  const json = await res.json();
  renderTable(json.data);
}

function renderTable(data) {
  excelContainer.innerHTML = "";

  const table = document.createElement("table");

  data.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");

    row.forEach((cell, colIndex) => {
      const td = document.createElement("td");
      td.contentEditable = true;
      td.textContent = cell ?? "";
      tr.appendChild(td);
    });

    table.appendChild(tr);
  });

  excelContainer.appendChild(table);
}

// SALVAR
document.getElementById("btnSave").onclick = async () => {
  const table = document.querySelector("table");
  const data = [];

  for (const row of table.rows) {
    const rowData = [];
    for (const cell of row.cells) {
      rowData.push(cell.textContent);
    }
    data.push(rowData);
  }

  await fetch("/api/excel", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ data }),
  });

  alert("Planilha salva!");
};

loadExcel();
