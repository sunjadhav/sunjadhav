const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const dashboard = document.getElementById("dashboard");
const sheetTabs = document.getElementById("sheetTabs");
const themeToggle = document.getElementById("themeToggle");

// Dark mode toggle
themeToggle.addEventListener("click", () => {
  document.body.classList.toggle("dark");
});

// Drag and drop
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

// File upload
fileInput.addEventListener("change", (e) => {
  handleFile(e.target.files[0]);
});

// File handler
function handleFile(file) {
  if (!file) return;
  const reader = new FileReader();
  if (file.name.endsWith(".csv")) {
    reader.onload = e => processCSV(e.target.result);
    reader.readAsText(file);
  } else if (file.name.endsWith(".xlsx")) {
    reader.onload = e => {
      const workbook = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
      displaySheets(workbook);
    };
    reader.readAsArrayBuffer(file);
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

function displaySheets(workbook) {
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

  // Auto-load first sheet
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const firstData = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
  renderDashboard(workbook.SheetNames[0], firstData);
}

function renderDashboard(sheetName, data) {
  dashboard.innerHTML = "";
  if (!data.length) return;

  const keys = Object.keys(data[0]);
  const numeric = keys.filter(k => !isNaN(parseFloat(data[0][k])));
  const card = document.createElement("div");
  card.className = "card";

  card.innerHTML = `<h2>${sheetName}</h2>`;

  // Summary
  let summary = "<h3>Summary</h3><table><tr><th>Column</th><th>Mean</th><th>Median</th><th>Min</th><th>Max</th><th>Std Dev</th></tr>";
  numeric.forEach(col => {
    const vals = data.map(d => parseFloat(d[col])).filter(v => !isNaN(v));
    if (!vals.length) return;
    const mean = (vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(2);
    const median = vals.sort((a, b) => a - b)[Math.floor(vals.length / 2)];
    const std = Math.sqrt(vals.map(v => (v - mean) ** 2).reduce((a, b) => a + b, 0) / vals.length).toFixed(2);
    summary += `<tr><td>${col}</td><td>${mean}</td><td>${median}</td><td>${Math.min(...vals)}</td><td>${Math.max(...vals)}</td><td>${std}</td></tr>`;
  });
  summary += "</table>";
  card.innerHTML += summary;

  // Charts
  numeric.slice(0, 2).forEach(col => {
    const labels = data.map((_, i) => `Row ${i + 1}`);
    const values = data.map(d => parseFloat(d[col]) || 0);

    const canvas = document.createElement("canvas");
    const download = document.createElement("button");
    download.textContent = "Download Chart";
    download.className = "download-btn";
    download.onclick = () => {
      const link = document.createElement("a");
      link.download = `${col}_chart.png`;
      link.href = canvas.toDataURL();
      link.click();
    };

    card.appendChild(download);
    card.appendChild(canvas);

    new Chart(canvas.getContext("2d"), {
      type: "bar",
      data: {
        labels,
        datasets: [{
          label: col,
          data: values,
          backgroundColor: "rgba(75, 192, 192, 0.6)"
        }]
      },
      options: { responsive: true }
    });
  });

  dashboard.appendChild(card);
}
