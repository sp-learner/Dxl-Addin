/* eslint-disable no-undef */
// Initialize Modal with enhanced column handling
function openFormatSheetModal() {
  const modal = document.getElementById("formatSheetModal");
  const overlay = document.getElementById("modalOverlay");

  modal.style.display = "block";
  overlay.style.display = "block";

  initializeColumns();
  loadFormatDropdown();
  resetForm();
}

// Modified function to open the Format Sheet modal
document.getElementById("FormatSheet").addEventListener("click", function() {
  document.getElementById("formatSheet").style.display = "block";
  document.getElementById("modalOverlay").style.display = "block";
  displaySavedFormats(); // Load and display saved formats instead of columns
});

document.getElementById("closeModal2").addEventListener("click", closeModal2);
document.getElementById("modalOverlay").addEventListener("click", closeModal2);

function closeModal2() {
  document.getElementById("formatSheet").style.display = "none";
  document.getElementById("modalOverlay").style.display = "none";
}

// Updated function to display saved formats as buttons
function displaySavedFormats() {
  const container = document.getElementById("savedFormatsContainer");
  container.innerHTML = "";

  const savedFormats = getAllSavedFormats();

  savedFormats.forEach(format => {
    const formatBtn = document.createElement("button");
    formatBtn.className = "format-btn";
    formatBtn.textContent = format.name;
    formatBtn.onclick = async function () {
      await loadFormatIntoWorkspace(format.name); // Updated function call
    };
    container.appendChild(formatBtn);
  });
}

async function loadFormatIntoWorkspace(formatName) {
  try {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
      throw new Error('Excel integration not supported');
    }

    const savedFormat = JSON.parse(localStorage.getItem(`FORMAT_${formatName}`));
    if (!savedFormat) {
      showToastNotification(`Format "${formatName}" not found!`);
      return;
    }

    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const currentSheet = context.workbook.worksheets.getActiveWorksheet();

      // Create or clear the target sheet
      let targetSheet;
      try {
        targetSheet = workbook.worksheets.add(formatName);
      } catch (error) {
        targetSheet = workbook.worksheets.getItem(formatName);
        targetSheet.activate();
        targetSheet.getUsedRange().clear(Excel.ClearApplyTo.contents);
      }

      // Step 1: Get data from current sheet
      const currentUsedRange = currentSheet.getUsedRange();
      currentUsedRange.load("values, columnCount, rowCount");
      await context.sync();

      const currentData = currentUsedRange.values;
      const currentHeaders = currentData[0];
      const currentRows = currentData.slice(1); // Data rows

      // Step 2: Map format columns to current sheet columns (using synonyms)
      const newFormatColumns = savedFormat.columns;
      const headerMap = {};

      newFormatColumns.forEach((newHeader, newIndex) => {
        // Check for exact match first
        const exactMatchIndex = currentHeaders.indexOf(newHeader);
        if (exactMatchIndex !== -1) {
          headerMap[newIndex] = exactMatchIndex;
          return;
        }

        // Check synonyms if no exact match
        const synonymKey = Object.keys(FormatSynonyms).find(key => 
          FormatSynonyms[key].includes(newHeader) || key === newHeader
        );

        if (synonymKey) {
          // Find the first matching synonym in current headers
          for (let i = 0; i < currentHeaders.length; i++) {
            const currentHeader = currentHeaders[i];
            if (
              currentHeader === synonymKey || 
              FormatSynonyms[synonymKey]?.includes(currentHeader)
            ) {
              headerMap[newIndex] = i;
              break;
            }
          }
        }
      });

      // Step 3: Write data to target sheet
      // Use CURRENT SHEET HEADERS (not standardized names) for the new sheet
      const headersToApply = newFormatColumns.map((_, newIndex) => {
        const currentColIndex = headerMap[newIndex];
        return currentColIndex !== undefined ? currentHeaders[currentColIndex] : newFormatColumns[newIndex];
      });

      // Apply headers (original names from current sheet)
      headersToApply.forEach((header, index) => {
        targetSheet.getRange("1:1").getCell(0, index).values = [[header]];
      });

      // Copy matching data
      if (currentRows.length > 0) {
        currentRows.forEach((row, rowIndex) => {
          Object.keys(headerMap).forEach(newCol => {
            const oldCol = headerMap[newCol];
            targetSheet.getRangeByIndexes(rowIndex + 1, parseInt(newCol), 1, 1).values = [[row[oldCol]]];
          });
        });
      }

      // Auto-fit columns
      targetSheet.getUsedRange().format.autofitColumns();
      await context.sync();


      showToastNotification(`"${formatName}" format applied Successfully`);
      localStorage.setItem('ACTIVE_FORMAT', formatName);
      closeModal2();
    });
  } catch (error) {
    console.error(`Error loading format "${formatName}":`, error);
    showToastNotification(`Failed: ${error.message}`);
  }
}

function showToastNotification(message, type = "info") {
  const toast = document.createElement("div");
  toast.className = `toast-notification toast-${type}`;
  toast.innerHTML = `
    <div class="toast-icon"></div>
    <div class="toast-message">${message}</div>
  `;

  document.body.appendChild(toast);

  setTimeout(() => {
    toast.classList.add("fade-out");
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}

// Helper function to get all saved formats (already in your code)
function getAllSavedFormats() {
  const savedFormats = [];
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key.startsWith("FORMAT_")) {
      try {
        const format = JSON.parse(localStorage.getItem(key));
        savedFormats.push(format);
      } catch (e) {
        console.error("Error parsing format:", key, e);
      }
    }
  }
  return savedFormats;
}

function closeModal(event) {
  if (event.target.id === "closeModal" || event.target === document.getElementById("formatSheetModal")) {
    document.getElementById("formatSheetModal").style.display = "none";
    document.getElementById("modalOverlay").style.display = "none";
  }
}

function initializeColumns() {
  const available = document.getElementById("availableColumns");

  if (available.options.length === 0) {
    const columns = [
      "Packet No", "Status", "Shape", "Weight", "Color",
      "PinneyColor", "Clarity", "Rate",
      "Disc", "NetRate", "Amount", "Cut"
    ];

    columns.forEach((col, index) => {
      available.add(new Option(`${index + 1}. ${col}`, col));
    });
  }
}

// Load saved formats into dropdown
function loadFormatDropdown() {
  const formatDropdown = document.getElementById("formatDropdown");
  formatDropdown.innerHTML = '<option value="">Select a format</option>';

  const savedFormats = getAllSavedFormats();
  savedFormats.forEach(format => {
    const option = document.createElement("option");
    option.value = format.name;
    option.textContent = format.name;
    formatDropdown.appendChild(option);
  });
}

// Get all saved formats from localStorage
function getAllSavedFormats() {
  const savedFormats = [];
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key.startsWith("FORMAT_")) {
      try {
        const format = JSON.parse(localStorage.getItem(key));
        savedFormats.push(format);
      } catch (e) {
        console.error("Error parsing format:", key, e);
      }
    }
  }
  return savedFormats;
}

function loadSelectedFormat() {
  // Clear all existing highlights first
  const available = document.getElementById("availableColumns");
  Array.from(available.options).forEach(option => {
    option.classList.remove("transferred");
  });

  const formatName = document.getElementById("formatDropdown").value;
  if (!formatName) return;

  const savedFormat = JSON.parse(localStorage.getItem(`FORMAT_${formatName}`));
  if (savedFormat) {
    document.getElementById("formatName").value = savedFormat.name || "";
    resetSelectedColumns();

    if (savedFormat.columns && savedFormat.columns.length > 0) {
      savedFormat.columns.forEach(col => {
        // Find and highlight the corresponding option in available columns
        const option = Array.from(available.options).find(opt => opt.value === col);
        if (option) {
          option.classList.add("transferred");
        }
        addColumnToSelection(col);
      });
    }
  }
}

// Reset the form to empty state
function resetForm() {
  document.getElementById("formatName").value = "";
  resetSelectedColumns();
  document.getElementById("formatDropdown").value = "";

  // Clear all transferred highlights
  const available = document.getElementById("availableColumns");
  Array.from(available.options).forEach(option => {
    option.classList.remove("transferred");
  });
}

// Reset selected columns section
function resetSelectedColumns() {
  const selectedContainer = document.getElementById("selectedColumnsContainer");
  if (selectedContainer) selectedContainer.innerHTML = "";
}

function transferColumns() {
  const available = document.getElementById("availableColumns");
  const selectedContainer = document.getElementById("selectedColumnsContainer");

  if (!available || !selectedContainer) return;

  const selectedOptions = Array.from(available.selectedOptions);

  selectedOptions.forEach(option => {
    const existingRow = Array.from(selectedContainer.querySelectorAll('tr')).find(
      row => row.dataset.column === option.value
    );

    if (!existingRow) {
      addColumnToSelection(option.value, option.text);
      option.classList.add('transferred');
      option.classList.add('selected-to-transfer');
      setTimeout(() => option.classList.remove('selected-to-transfer'), 300);
    }
  });

  updateButtonStates();
}

function addColumnToTable(columnValue, columnText) {
  const selectedContainer = document.getElementById("selectedColumnsContainer");

  if (!selectedContainer) return;

  const row = document.createElement('tr');
  row.dataset.column = columnValue;

  const srCell = document.createElement('td');
  srCell.textContent = selectedContainer.querySelectorAll('tr').length + 1;

  const nameCell = document.createElement('td');
  nameCell.textContent = columnText.split('. ')[1] || columnText;

  row.appendChild(srCell);
  row.appendChild(nameCell);
  selectedContainer.appendChild(row);

  updateSerialNumbers();
}

function updateSerialNumbers() {
  const rows = document.querySelectorAll('#selectedColumnsContainer tr');
  rows.forEach((row, index) => {
    row.cells[0].textContent = index + 1;
  });
}

function returnColumns() {
  const available = document.getElementById("availableColumns");
  const selectedContainer = document.getElementById("selectedColumnsContainer");
  if (!selectedContainer) return;

  const selectedRows = Array.from(selectedContainer.querySelectorAll('tr.selected-row'));

  selectedRows.forEach(row => {
    row.remove();
    const option = Array.from(available.options).find(opt => opt.value === row.dataset.column);
    if (option) {
      option.classList.remove('transferred');
      option.classList.add('returned-option');
      setTimeout(() => option.classList.remove('returned-option'), 300);
    }
  });

  updateSerialNumbers();
  updateButtonStates();
}

function updateButtonStates() {
  const available = document.getElementById("availableColumns");
  const selectedContainer = document.getElementById("selectedColumnsContainer");

  document.getElementById("addColumn").disabled = !available?.selectedOptions?.length;
  document.getElementById("removeColumn").disabled = !selectedContainer?.querySelector('tr.selected-row');
}

function saveFormat() {
  try {
    const formatName = document.getElementById("formatName").value.trim();
    const selectedContainer = document.getElementById("selectedColumnsContainer");

    if (!formatName) {
      showToastNotification("Please enter a format name", "error");
      return;
    }

    const selectedColumns = selectedContainer ? 
      Array.from(selectedContainer.querySelectorAll('tr')).map(row => row.dataset.column) : [];

    if (selectedColumns.length === 0) {
      showToastNotification("Please select at least one column", "warning");
      return;
    }

    const formatData = {
      name: formatName,
      columns: selectedColumns,
      timestamp: new Date().toISOString()
    };

    // Check if format already exists
    const existingFormat = localStorage.getItem(`FORMAT_${formatName}`);
    const isUpdate = existingFormat !== null;

    // Save with unique key
    localStorage.setItem(`FORMAT_${formatName}`, JSON.stringify(formatData));

    loadFormatDropdown();

    // Reset form for new format
    resetForm();

    showToastNotification(
      isUpdate ? `"${formatName}" updated successfully!` : `"${formatName}" created successfully!`,
      "success"
    );
  } catch (error) {
    console.error("Error saving format:", error);
    showToastNotification("Failed to save format", "error");
  }
}

function deleteFormat() {
  try {
    const formatDropdown = document.getElementById("formatDropdown");
    if (!formatDropdown) {
      showToastNotification("Format dropdown element not found", "error");
      return;
    }

    const formatName = formatDropdown.value;
    if (!formatName) {
      showToastNotification("Please select a format to delete", "error");
      return;
    }

    // // Confirm deletion with a custom dialog (better than native confirm)
    // if (!confirmDelete(formatName)) {
    //   return;
    // }

    const storageKey = `FORMAT_${formatName}`;
    if (!localStorage.getItem(storageKey)) {
      showToastNotification(`Format "${formatName}" not found`, "error");
      return;
    }

    // Perform deletion
    localStorage.removeItem(storageKey);

    // Refresh UI
    loadFormatDropdown();
    resetForm();

    // Show success notification
    showToastNotification(`"${formatName}" deleted successfully`, "success");
  } catch (error) {
    console.error("Delete format error:", error);
    showToastNotification("Failed to delete format", "error");
  }
}

function addColumnToSelection(columnValue) {
  const available = document.getElementById("availableColumns");
  const selectedContainer = document.getElementById("selectedColumnsContainer");

  if (!available || !selectedContainer) return;

  for (let option of available.options) {
    if (option.value === columnValue) {
      const exists = Array.from(selectedContainer.querySelectorAll('tr')).some(
        row => row.dataset.column === columnValue
      );
      if (!exists) {
        addColumnToTable(option.value, option.text);
        option.classList.add('transferred');
      }
      break;
    }
  }
}

// Initialize event listeners
function initFormatSheet() {
  // Modal events
  document.getElementById("closeModal")?.addEventListener("click", closeModal);
  document.getElementById("formatSheetModal")?.addEventListener("click", closeModal);
  document.getElementById("Format Sheet")?.addEventListener("click", closeModal2);
  document.getElementById("closeModal2")?.addEventListener("click", closeModal2);

  // Format operations
  document.getElementById("addColumn")?.addEventListener("click", transferColumns);
  document.getElementById("removeColumn")?.addEventListener("click", returnColumns);
  document.getElementById("saveFormat")?.addEventListener("click", saveFormat);
  document.getElementById("deleteFormat")?.addEventListener("click", deleteFormat);
  document.getElementById("formatDropdown")?.addEventListener("change", loadSelectedFormat);
  document.getElementById("newFormatBtn")?.addEventListener("click", resetForm);

  document.getElementById("availableColumns")?.addEventListener("change", updateButtonStates);

  // Row selection
  document.addEventListener('click', function(e) {
    const selectedContainer = document.getElementById("selectedColumnsContainer");
    if (!selectedContainer) return;

    if (e.target.closest('#selectedColumnsContainer tr')) {
      const row = e.target.closest('tr');
      row.classList.toggle('selected-row');
      updateButtonStates();
    }
  });

  updateButtonStates();
}

document.addEventListener("DOMContentLoaded", initFormatSheet);