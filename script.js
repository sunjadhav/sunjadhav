const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const dashboard = document.getElementById("dashboard");
const sheetTabs = document.getElementById("sheetTabs");
const themeToggle = document.getElementById("themeToggle");

themeToggle.addEventListener("click", () => document.body.classList.toggle("dark"));

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.style.borderColor = "green";
});
dropZone.addEventListener("dragleave", () => {
  dropZone.style.borderColor = "";
});
dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener("change", (e) => handleFile(e.target.files[0]));

function handleFile(file) {
  if (!file) return;
  const reader = new FileReader();
  if (file.name.endsWith(".csv")) {
    reader.onload = e => processCSV(e.target.result);
    reader.readAsText(file);
  } else {
    reader.onload = e => {
      const workbook = XLSX.read(e.target.result, { type: "binary" });
      loadSheets(workbook);
    };
    reader.readAsBinaryString(file);
  }
}

function processCSV(csvText) {
  const lines = csvText.split("\n");
  const headers = lines[0].split(",");
  const data = lines.slice(1).map(row => {
    const values = row.split(",");
    const obj = {};
    headers.forEach((h, i) => obj[h.trim()] = values[i]?.trim());
    return obj;
  });
  renderDashboard("CSV Sheet", data);
}

function loadSheets(workbook) {
  sheetTabs.innerHTML = "";
  workbook.SheetNames.forEach(name => {
    const btn = document.createElement("button");
    btn.textContent = name;
    btn.onclick = () => {
      const sheet = workbook.Sheets[name];
      const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      renderDashboard(name, data);
    };
    sheetTabs.appendChild(btn);
  });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  renderDashboard(workbook.SheetNames[0], XLSX.utils.sheet_to_json(firstSheet, { defval: "" }));
}

function renderDashboard(name, data) {
  dashboard.innerHTML = "";
  if (!data.length) return;
  const card = document.createElement("div");
  card.className = "card";
  card.innerHTML = `<h2>${name}</h2>`;

  const keys = Object.keys(data[0]);
  const numeric = keys.filter(k => !isNaN(parseFloat(data[0][k])));
  const categoricals = keys.filter(k => !numeric.includes(k));

  // Pie/Donut for top category breakdown
  if (categoricals.length && numeric.length) {
    const catKey = categoricals[0];
    const numKey = numeric[0];
    const grouped = {};
    data.forEach(row => {
      grouped[row[catKey]] = (grouped[row[catKey]] || 0) + parseFloat(row[numKey] || 0);
    });
    const labels = Object.keys(grouped);
    const values = Object.values(grouped);

    createChart(card, labels, values, `Top ${catKey} by ${numKey}`, "pie");
  }

  // Line chart for time series
  const dateKey = keys.find(k => /date/i.test(k));
  if (dateKey && numeric.length) {
    const sorted = [...data].sort((a, b) => new Date(a[dateKey]) - new Date(b[dateKey]));
    const labels = sorted.map(d => d[dateKey]);
    const values = sorted.map(d => parseFloat(d[numeric[0]]) || 0);
    createChart(card, labels, values, `${numeric[0]} Over Time`, "line");
  }

  // Bar charts for top numeric fields
  numeric.forEach(numKey => {
    const labels = data.map((_, i) => `Row ${i + 1}`);
    const values = data.map(d => parseFloat(d[numKey]) || 0);
    createChart(card, labels, values, numKey, "bar");
  });

  dashboard.appendChild(card);
}

function createChart(container, labels, values, title, type) {
  const canvas = document.createElement("canvas");
  const download = document.createElement("button");
  download.textContent = "Download Chart";
  download.className = "download-btn";
  download.onclick = () => {
    const link = document.createElement("a");
    link.download = `${title}_chart.png`;
    link.href = canvas.toDataURL();
    link.click();
  };
  container.appendChild(download);
  container.appendChild(canvas);

  new Chart(canvas.getContext("2d"), {
    type,
    data: {
      labels,
      datasets: [{
        label: title,
        data: values,
        backgroundColor: type === "pie" ? generateColors(labels.length) : "rgba(75, 192, 192, 0.6)",
        borderColor: "rgba(0,0,0,0.1)",
        borderWidth: 1
      }]
    },
    options: { responsive: true, plugins: { legend: { display: type === "pie" } } }
  });
}

function generateColors(n) {
  return Array.from({ length: n }, (_, i) =>
    `hsl(${(i * 360) / n}, 70%, 60%)`
  );
}
