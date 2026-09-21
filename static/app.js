let selectedFile = null;

const fileInput = document.getElementById("file-input");
const fileDrop = document.getElementById("file-drop");
const fileDropLabel = document.getElementById("file-drop-label");
const step2 = document.getElementById("step2");
const inspectError = document.getElementById("inspect-error");
const convertError = document.getElementById("convert-error");
const convertStatus = document.getElementById("convert-status");

function showError(el, message) {
  el.textContent = message;
  el.classList.remove("hidden");
}
function hideError(el) {
  el.classList.add("hidden");
}

fileDrop.addEventListener("click", () => fileInput.click());
fileDrop.addEventListener("dragover", (e) => { e.preventDefault(); fileDrop.classList.add("drag"); });
fileDrop.addEventListener("dragleave", () => fileDrop.classList.remove("drag"));
fileDrop.addEventListener("drop", (e) => {
  e.preventDefault();
  fileDrop.classList.remove("drag");
  if (e.dataTransfer.files.length) {
    fileInput.files = e.dataTransfer.files;
    handleFile(e.dataTransfer.files[0]);
  }
});
fileInput.addEventListener("change", (e) => {
  if (e.target.files.length) handleFile(e.target.files[0]);
});

async function handleFile(file) {
  selectedFile = file;
  fileDropLabel.textContent = file.name;
  hideError(inspectError);
  step2.classList.add("hidden");

  const formData = new FormData();
  formData.append("file", file);

  try {
    const res = await fetch("/api/inspect", { method: "POST", body: formData });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || "Xəta baş verdi");

    const sheetSelect = document.getElementById("sheet-select");
    sheetSelect.innerHTML = "";
    for (const s of data.sheets) {
      const opt = document.createElement("option");
      opt.value = s.name;
      opt.textContent = `${s.name} (${s.row_count} sətir)`;
      sheetSelect.appendChild(opt);
    }

    document.getElementById("contractor-input").value = data.suggested_contractor || "";
    document.getElementById("filename-input").value = data.suggested_filename || "converted";

    step2.classList.remove("hidden");
  } catch (err) {
    showError(inspectError, err.message);
  }
}

document.getElementById("convert-btn").addEventListener("click", async () => {
  if (!selectedFile) return;
  hideError(convertError);
  convertStatus.classList.add("hidden");

  const filenameRaw = document.getElementById("filename-input").value.trim();
  if (!filenameRaw) {
    showError(convertError, "Fayl üçün ad daxil edin");
    return;
  }
  const contractorName = document.getElementById("contractor-input").value.trim();
  if (!contractorName) {
    showError(convertError, "Podratçının adını daxil edin");
    return;
  }

  const btn = document.getElementById("convert-btn");
  btn.disabled = true;
  btn.textContent = "Çevrilir...";

  try {
    const formData = new FormData();
    formData.append("file", selectedFile);
    formData.append("sheet_name", document.getElementById("sheet-select").value);
    formData.append("contractor_name", contractorName);
    formData.append("client_name", document.getElementById("client-input").value.trim() || "Salyan Oil");

    const res = await fetch("/api/convert", { method: "POST", body: formData });
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.detail || "Çevirmə xətası");
    }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const cleanName = filenameRaw.replace(/\.xlsx$/i, "");
    a.download = `${cleanName}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    convertStatus.textContent = `Hazırdır: "${cleanName}.xlsx" yükləndi.`;
    convertStatus.classList.remove("hidden");
  } catch (err) {
    showError(convertError, err.message);
  } finally {
    btn.disabled = false;
    btn.textContent = "Çevir və yüklə";
  }
});
