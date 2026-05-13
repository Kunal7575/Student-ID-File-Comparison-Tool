let missingRecords = [];

const mainFileInput = document.getElementById("mainFile");
const secondaryFileInput = document.getElementById("secondaryFile");

const mainDropZone = document.getElementById("mainDropZone");
const secondaryDropZone = document.getElementById("secondaryDropZone");

const mainFileName = document.getElementById("mainFileName");
const secondaryFileName = document.getElementById("secondaryFileName");

const compareBtn = document.getElementById("compareBtn");
const downloadBtn = document.getElementById("downloadBtn");

const messageBox = document.getElementById("message");
const summaryBox = document.getElementById("summary");
const previewSection = document.getElementById("previewSection");
const previewTable = document.getElementById("previewTable");

const mainCount = document.getElementById("mainCount");
const secondaryCount = document.getElementById("secondaryCount");
const missingCount = document.getElementById("missingCount");

setupDropZone(mainDropZone, mainFileInput, mainFileName);
setupDropZone(secondaryDropZone, secondaryFileInput, secondaryFileName);

document.querySelectorAll(".browse-btn").forEach(button => {
  button.addEventListener("click", () => {
    const targetId = button.dataset.target;
    document.getElementById(targetId).click();
  });
});

mainFileInput.addEventListener("change", () => {
  mainFileName.textContent = mainFileInput.files[0]?.name || "No file selected";
});

secondaryFileInput.addEventListener("change", () => {
  secondaryFileName.textContent = secondaryFileInput.files[0]?.name || "No file selected";
});

compareBtn.addEventListener("click", async () => {
  resetUI();

  const mainFile = mainFileInput.files[0];
  const secondaryFile = secondaryFileInput.files[0];

  if (!mainFile || !secondaryFile) {
    showMessage("Please upload both the main file and the secondary file.", "error");
    return;
  }

  try {
    const mainData = await readExcelFile(mainFile);
    const secondaryData = await readExcelFile(secondaryFile);

    if (mainData.length === 0 || secondaryData.length === 0) {
      showMessage("One of the files appears to be empty.", "error");
      return;
    }

    if (secondaryData.length > mainData.length) {
      showMessage(
        "Warning: The secondary file has more records than the main file. Please check if the files were uploaded in the wrong order.",
        "warning"
      );
      return;
    }

    const mainIdColumn = findStudentIdColumn(mainData);
    const secondaryIdColumn = findStudentIdColumn(secondaryData);

    if (!mainIdColumn || !secondaryIdColumn) {
      showMessage(
        "Could not find the Student ID column. Please make sure both files have a column like 'Student ID #' or 'Student ID'.",
        "error"
      );
      return;
    }

    const mainIds = new Set(
      mainData
        .map(row => cleanStudentId(row[mainIdColumn]))
        .filter(id => id !== "")
    );

    missingRecords = secondaryData.filter(row => {
      const id = cleanStudentId(row[secondaryIdColumn]);
      return id !== "" && !mainIds.has(id);
    });

    mainCount.textContent = mainData.length;
    secondaryCount.textContent = secondaryData.length;
    missingCount.textContent = missingRecords.length;
    summaryBox.classList.remove("hidden");

    if (missingRecords.length === 0) {
      showMessage(
        "Congrats, all secondary records exist in the main file. No missing records found.",
        "success"
      );
    } else {
      showMessage(
        `${missingRecords.length} missing record(s) found. You can download the report below.`,
        "warning"
      );

      downloadBtn.classList.remove("hidden");
      showPreview(missingRecords);
    }

  } catch (error) {
    console.error(error);
    showMessage(
      "Something went wrong while reading the files. Please check the Excel format and try again.",
      "error"
    );
  }
});

downloadBtn.addEventListener("click", () => {
  if (missingRecords.length === 0) {
    showMessage("No missing records to download.", "success");
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(missingRecords);
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, worksheet, "Missing Records");
  XLSX.writeFile(workbook, "missing_records_report.xlsx");
});

function setupDropZone(dropZone, input, fileNameLabel) {
  dropZone.addEventListener("dragover", event => {
    event.preventDefault();
    dropZone.classList.add("dragover");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragover");
  });

  dropZone.addEventListener("drop", event => {
    event.preventDefault();
    dropZone.classList.remove("dragover");

    const file = event.dataTransfer.files[0];

    if (!file) return;

    const allowedExtensions = [".xlsx", ".xls", ".csv"];
    const fileName = file.name.toLowerCase();

    const isAllowed = allowedExtensions.some(extension =>
      fileName.endsWith(extension)
    );

    if (!isAllowed) {
      showMessage("Please upload only Excel or CSV files.", "error");
      return;
    }

    input.files = event.dataTransfer.files;
    fileNameLabel.textContent = file.name;
  });
}

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = event => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
          defval: "",
          raw: false
        });

        resolve(jsonData);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function findStudentIdColumn(data) {
  if (!data || data.length === 0) return null;

  const columns = Object.keys(data[0]);

  const possibleNames = [
    "student id #",
    "student id",
    "student number",
    "student number #",
    "studentid",
    "studentnumber",
    "id"
  ];

  return columns.find(column => {
    const cleanedColumn = column
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();

    const compactColumn = cleanedColumn.replace(/[^a-z0-9]/g, "");

    return (
      possibleNames.includes(cleanedColumn) ||
      possibleNames.includes(compactColumn)
    );
  });
}

function cleanStudentId(value) {
  if (value === null || value === undefined) return "";

  let id = String(value)
    .trim()
    .replace(/\.0$/, "")
    .replace(/\s+/g, "");

  id = id.replace(/[^0-9]/g, "");

  id = id.replace(/^0+/, "");

  return id;
}

function showMessage(text, type) {
  messageBox.textContent = text;
  messageBox.className = `message ${type}`;
  messageBox.classList.remove("hidden");
}

function showPreview(records) {
  previewSection.classList.remove("hidden");
  previewTable.innerHTML = "";

  const previewRecords = records.slice(0, 20);

  if (previewRecords.length === 0) return;

  const columns = Object.keys(previewRecords[0]);

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");

  columns.forEach(column => {
    const th = document.createElement("th");
    th.textContent = column;
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);
  previewTable.appendChild(thead);

  const tbody = document.createElement("tbody");

  previewRecords.forEach(record => {
    const row = document.createElement("tr");

    columns.forEach(column => {
      const td = document.createElement("td");
      td.textContent = record[column] ?? "";
      row.appendChild(td);
    });

    tbody.appendChild(row);
  });

  previewTable.appendChild(tbody);
}

function resetUI() {
  missingRecords = [];

  messageBox.className = "message hidden";
  messageBox.textContent = "";

  summaryBox.classList.add("hidden");
  previewSection.classList.add("hidden");
  downloadBtn.classList.add("hidden");

  previewTable.innerHTML = "";

  mainCount.textContent = "0";
  secondaryCount.textContent = "0";
  missingCount.textContent = "0";
}
