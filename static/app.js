const uploadForm = document.querySelector("#upload-form");
const uploadButton = document.querySelector("#upload-button");
const fileInput = document.querySelector("#document-file");
const utilityNameInput = document.querySelector("#utility-name");
const uploadStatus = document.querySelector("#upload-status");
const uploadCount = document.querySelector("#upload-count");
const uploadCountBadge = document.querySelector("#upload-count-badge");
const databaseMode = document.querySelector("#database-mode");
const setupWarning = document.querySelector("#setup-warning");
const clearButton = document.querySelector("#clear-uploads");
const refreshButton = document.querySelector("#refresh-map");
const mapFrame = document.querySelector("#map-frame");
const importsToggle = document.querySelector("#imports-toggle");
const uploadDrawer = document.querySelector("#upload-drawer");
const drawerBackdrop = document.querySelector("#drawer-backdrop");
const drawerClose = document.querySelector("#drawer-close");
const totalCount = document.querySelector("#total-count");
console.log("totalCount:", totalCount);

function openDrawer() {
  uploadDrawer.classList.add("is-open");
  drawerBackdrop.hidden = false;
  importsToggle.setAttribute("aria-expanded", "true");
}

function closeDrawer() {
  uploadDrawer.classList.remove("is-open");
  drawerBackdrop.hidden = true;
  importsToggle.setAttribute("aria-expanded", "false");
}

importsToggle.addEventListener("click", () => {
  if (uploadDrawer.classList.contains("is-open")) {
    closeDrawer();
  } else {
    openDrawer();
  }
});
drawerClose.addEventListener("click", closeDrawer);
drawerBackdrop.addEventListener("click", closeDrawer);
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") { closeDrawer(); }
});

async function requestJson(url, options = {}) {
  const response = await fetch(url, options);
  const payload = await response.json();

  if (!response.ok) {
    throw new Error(payload.error || "The request could not be completed.");
  }

  return payload;
}

function setUploadStatus(message, isError = false) {
  uploadStatus.textContent = message;
  uploadStatus.classList.toggle("is-error", isError);
}

async function refreshStatus() {
  try {
    const status = await requestJson("/api/status");
    const count = status.uploaded_count;
    uploadCount.textContent = `${count} ${count === 1 ? "project" : "projects"}`;
    uploadCountBadge.textContent = String(count);
    totalCount.textContent = `${status.total_count} mapped projects`;
    databaseMode.textContent = status.database_backend;
    clearButton.disabled = count === 0;
    uploadButton.disabled = !status.gemini_configured;
    setupWarning.hidden = status.gemini_configured;
  } catch (error) {
    setUploadStatus(error.message, true);
  }
}

function refreshMap() {
  mapFrame.src = `/map?refresh=${Date.now()}`;
}

uploadForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (!fileInput.files.length) {
    setUploadStatus("Choose a document first.", true);
    return;
  }

  const utilityName = utilityNameInput.value.trim();
  if (!utilityName) {
    setUploadStatus("Enter the utility name first.", true);
    return;
  }

  const formData = new FormData();
  formData.set("document", fileInput.files[0]);
  formData.set("utility_name", utilityName);
  uploadButton.disabled = true;
  uploadButton.textContent = "Sending to Gemini...";
  setUploadStatus("Reading the document and mapping its projects.");

  try {
    const result = await requestJson("/api/upload", {
      method: "POST",
      body: formData,
    });
    const summary = [`Added ${result.added} ${result.added === 1 ? "project" : "projects"} (of ${result.extracted} extracted).`];

    if (result.already_imported) {
      summary.push(`${result.already_imported} already imported.`);
    }
    if (result.not_mappable) {
      summary.push(`${result.not_mappable} could not be geocoded -- check the server console for why.`);
    }
    if (result.geocoded) {
      summary.push(`${result.geocoded} location lookups used OpenStreetMap.`);
    }

    setUploadStatus(summary.join(" "));
    uploadForm.reset();
    refreshMap();
    closeDrawer();
    await refreshStatus();
  } catch (error) {
    setUploadStatus(error.message, true);
  } finally {
    uploadButton.textContent = "Parse and add projects";
    await refreshStatus();
  }
});

clearButton.addEventListener("click", async () => {
  if (!window.confirm("Remove all locally imported projects from the map?")) {
    return;
  }

  clearButton.disabled = true;
  try {
    await requestJson("/api/uploaded-projects", { method: "DELETE" });
    setUploadStatus("Imported projects removed.");
    refreshMap();
    await refreshStatus();
  } catch (error) {
    setUploadStatus(error.message, true);
    await refreshStatus();
  }
});

refreshButton.addEventListener("click", refreshMap);
refreshStatus();