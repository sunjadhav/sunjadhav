document.getElementById("fileInput").addEventListener("change", handleFile, false);

function handleFile(event) {
  const file = event.target.files[0];
  const reader = new FileReader();

  if (file.name.endsWith(".csv")) {
    reader.onload = (e) => processCSV(e.target.result);
    reader.readAsText(file);
  } else if (file.name.endsWith(".xlsx")) {
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      generateDashboard(json);
    };
    reader.readAsArrayBuffer(file);
  }
}

function processCSV(csvText) {
  const lines = csvText.split("\n");
  const headers = lines[0].split(",");
  const data = lines.slice(1).map(line => {
    const values = line.split(",");
    const obj = {};
    headers.forEach((h, i) => (obj[h.trim()] = values[i]?.trim()));
    return obj;
  });
  generateDashboard(data);
}

function generateDashboard(data) {
  const dashboard = document.getElementById("dashboard");
  dashboard.innerHTML = "";

  if (data.length === 0) return;

  // Show summary stats
  const keys = Object.keys(data[0]);
  const numericKeys = keys.filter(k => !isNaN(data[0][k]));

  let summaryHTML = "<h2>Summary</h2><table><thead><tr><th>Column</th><th>Mean</th><th>Count</th></tr></thead><tbody>";
  numericKeys.forEach(key => {
    const values = data.map(row => parseFloat(row[key])).filter(v => !isNaN(v));
    const mean = (values.reduce((a, b) => a + b, 0) / values.length).toFixed(2);
    summaryHTML += `<tr><td>${key}</td><td>${mean}</td><td>${values.length}</td></tr>`;
  });
  summaryHTML += "</tbody></table>";
  dashboard.innerHTML += summaryHTML;

  // Show data table
  let tableHTML = "<h2>Data Preview</h2><table><thead><tr>";
  keys.forEach(k => tableHTML += `<th>${k}</th>`);
  tableHTML += "</tr></thead><tbody>";
  data.slice(0, 10).forEach(row => {
    tableHTML += "<tr>";
    keys.forEach(k => tableHTML += `<td>${row[k]}</td>`);
    tableHTML += "</tr>";
  });
  tableHTML += "</tbody></table>";
  dashboard.innerHTML += tableHTML;

  // Generate a chart using the first numeric column
  if (numericKeys.length > 0) {
    const ctx = document.getElementById("chartCanvas").getContext("2d");
    new Chart(ctx, {
      type: "bar",
      data: {
        labels: data.map((row, i) => `Row ${i + 1}`),
        datasets: [{
          label: numericKeys[0],
          data: data.map(row => parseFloat(row[numericKeys[0]]) || 0),
          backgroundColor: "rgba(75, 192, 192, 0.5)",
          borderColor: "rgba(75, 192, 192, 1)",
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        scales: {
          y: { beginAtZero: true }
        }
      }
    });
  }
}
