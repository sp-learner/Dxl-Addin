function openFormatModal(mode) {
  const modal = document.getElementById("formatSheetModal");
  const overlay = document.getElementById("modalOverlay");

 
  if (mode === "custom") {
    modal.querySelector(".modal-header span").textContent = "Custom Format Options";
  } else {
    modal.querySelector(".modal-header span").textContent = "Format Sheet Options";
  }

  modal.style.display = "block";
  overlay.style.display = "block";
}

// Close modal handler
document.getElementById("closeFormatModal").addEventListener("click", function () {
  document.getElementById("formatSheetModal").style.display = "none";
  document.getElementById("modalOverlay").style.display = "none";
});
