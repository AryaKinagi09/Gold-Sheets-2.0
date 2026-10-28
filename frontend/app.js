const combo = document.getElementById("builder-combo");
const builderCard = document.getElementById("builder-card");
const builderLabel = document.getElementById("builder-label");
const builderSearch = document.getElementById("builder-search");
const builderOptions = document.getElementById("builder-options");
const builderNote = document.getElementById("builder-note");
const divisionsCard = document.getElementById("divisions-card");
const divisionSearch = document.getElementById("division-search");
const regionsEl = document.getElementById("regions");
const selectAll = document.getElementById("select-all");
const divisionNote = document.getElementById("division-note");
const generateBtn = document.getElementById("generate");
const statusEl = document.getElementById("status");
const modeNote = document.getElementById("mode-note");
const modeButtons = [...document.querySelectorAll(".mode")];

let builders = [];
let selectedKey = null;
let sheetType = "corporate";
let currentDivisions = [];

function displayBuilder(builder) {
  return builder.has_prior_gold_sheet
    ? builder.builder_name
    : `${builder.builder_name} (no prior sheet)`;
}

function selectedBuilder() {
  return builders.find((b) => b.builder_key === selectedKey) || null;
}

function buildersForMode() {
  if (sheetType === "corporate") {
    // Every main builder (one row per builder_key), not division accounts.
    // Do not require has_corporate — many HQ rows are not named "Corporate".
    return builders;
  }
  return builders.filter((b) => b.region_count > 0);
}

function selectedRegions() {
  return [...regionsEl.querySelectorAll("input[type=checkbox]:checked")].map(
    (box) => box.value
  );
}

function setMode(mode) {
  sheetType = mode;
  modeButtons.forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.mode === mode);
  });
  const isDivision = mode === "division";
  divisionsCard.hidden = !isDivision;
  builderLabel.textContent = "Builder";
  builderSearch.placeholder = "Select a builder";
  modeNote.textContent = isDivision
    ? "Divisions uses pages 3–4. Pick a builder, then choose one or more of its divisions."
    : "Corporate uses pages 1–2. Pick a builder; all of its divisions are listed on the sheet.";

  const stillValid = buildersForMode().some((b) => b.builder_key === selectedKey);
  if (!stillValid) {
    selectedKey = null;
    builderSearch.value = "";
    builderNote.textContent = "";
    currentDivisions = [];
    regionsEl.innerHTML = `<p class="muted">Select a builder to see its divisions.</p>`;
    divisionNote.textContent = "";
    divisionSearch.value = "";
  } else if (isDivision) {
    loadDivisionsForBuilder();
  }
  syncGenerateState();
}

function builderListQuery() {
  const builder = selectedBuilder();
  const text = builderSearch.value.trim().toLowerCase();
  if (builder && text === displayBuilder(builder).toLowerCase()) {
    return "";
  }
  return text;
}

async function ensureBuildersLoaded() {
  if (builders.length > 0) {
    return true;
  }
  try {
    const res = await fetch("/builders", { cache: "no-store" });
    if (!res.ok) {
      throw new Error(`Could not load builders (${res.status})`);
    }
    const data = await res.json();
    if (!Array.isArray(data)) {
      throw new Error("Builder list was not an array");
    }
    builders = data;
    statusEl.textContent =
      data.length > 0
        ? `${data.length} builders loaded.`
        : "No builders found in the database.";
    return data.length > 0;
  } catch (err) {
    builders = [];
    statusEl.textContent =
      err.message || "Could not load builders. Is the API running?";
    return false;
  }
}

async function openBuilderMenu() {
  builderOptions.hidden = false;
  builderSearch.setAttribute("aria-expanded", "true");
  builderOptions.innerHTML =
    '<li class="muted" style="padding:8px 10px">Loading builders…</li>';

  await ensureBuildersLoaded();

  const query = builderListQuery();
  const matches = buildersForMode().filter((b) =>
    displayBuilder(b).toLowerCase().includes(query)
  );
  builderOptions.innerHTML = "";
  matches.forEach((builder) => {
    const item = document.createElement("li");
    const button = document.createElement("button");
    button.type = "button";
    button.setAttribute("role", "option");
    button.textContent = displayBuilder(builder);
    button.setAttribute(
      "aria-selected",
      builder.builder_key === selectedKey ? "true" : "false"
    );
    button.addEventListener("click", () => chooseBuilder(builder.builder_key));
    item.appendChild(button);
    builderOptions.appendChild(item);
  });
  if (!matches.length) {
    const item = document.createElement("li");
    item.className = "muted";
    item.textContent =
      builders.length === 0
        ? "Could not load builders"
        : sheetType === "corporate"
          ? "No matching builders"
          : "No matching builders with divisions";
    item.style.padding = "8px 10px";
    builderOptions.appendChild(item);
  }
}

function closeBuilderMenu() {
  builderOptions.hidden = true;
  builderSearch.setAttribute("aria-expanded", "false");
}

async function chooseBuilder(builderKey) {
  const builder = builders.find((b) => b.builder_key === builderKey);
  selectedKey = builder ? builderKey : null;
  builderSearch.value = builder ? displayBuilder(builder) : "";
  closeBuilderMenu();

  if (!builder) {
    builderNote.textContent = "";
    currentDivisions = [];
    regionsEl.innerHTML = `<p class="muted">Select a builder to see its divisions.</p>`;
    syncGenerateState();
    return;
  }

  if (sheetType === "corporate") {
    builderNote.textContent = builder.has_prior_gold_sheet
      ? `${builder.region_count} divisions will be listed on the corporate sheet`
      : "No prior Gold Sheet yet, so a PDF cannot be generated.";
    syncGenerateState();
    return;
  }

  builderNote.textContent = builder.has_prior_gold_sheet
    ? `${builder.region_count} divisions for this builder`
    : "No prior Gold Sheet yet, so a PDF cannot be generated.";
  await loadDivisionsForBuilder();
}

async function loadDivisionsForBuilder() {
  const builder = selectedBuilder();
  selectAll.checked = false;
  divisionSearch.value = "";
  currentDivisions = [];
  regionsEl.innerHTML = `<p class="muted">Loading divisions…</p>`;
  divisionNote.textContent = "";

  if (!builder) {
    regionsEl.innerHTML = `<p class="muted">Select a builder to see its divisions.</p>`;
    syncGenerateState();
    return;
  }

  const res = await fetch(
    `/builders/${encodeURIComponent(builder.builder_key)}/regions`
  );
  if (!res.ok) {
    regionsEl.innerHTML = `<p class="muted">Could not load divisions.</p>`;
    syncGenerateState();
    return;
  }
  currentDivisions = await res.json();
  renderDivisions();
}

function renderDivisions() {
  const builder = selectedBuilder();
  const query = divisionSearch.value.trim().toLowerCase();
  const matches = currentDivisions.filter((d) =>
    String(d.region_label || "")
      .toLowerCase()
      .includes(query)
  );

  regionsEl.innerHTML = "";
  if (!builder) {
    regionsEl.innerHTML = `<p class="muted">Select a builder to see its divisions.</p>`;
    divisionNote.textContent = "";
    syncGenerateState();
    return;
  }
  if (!matches.length) {
    regionsEl.innerHTML = `<p class="muted">No divisions match that search.</p>`;
    divisionNote.textContent = "";
    syncGenerateState();
    return;
  }

  matches.forEach((division) => {
    const label = document.createElement("label");
    label.className = "region";
    const box = document.createElement("input");
    box.type = "checkbox";
    box.value = division.region_label;
    box.disabled = !builder.has_prior_gold_sheet;
    box.addEventListener("change", syncGenerateState);
    label.append(
      box,
      document.createTextNode(` ${division.region_label}`)
    );
    regionsEl.appendChild(label);
  });
  divisionNote.textContent = `${matches.length} division(s) for ${builder.builder_name}`;
  syncGenerateState();
}

function syncGenerateState() {
  const builder = selectedBuilder();
  const boxes = [
    ...regionsEl.querySelectorAll("input[type=checkbox]:not(:disabled)"),
  ];
  const checked = boxes.filter((box) => box.checked);
  selectAll.checked = boxes.length > 0 && checked.length === boxes.length;

  if (sheetType === "corporate") {
    generateBtn.disabled = !(builder && builder.has_prior_gold_sheet);
    return;
  }
  generateBtn.disabled = !(
    builder &&
    builder.has_prior_gold_sheet &&
    selectedRegions().length > 0
  );
}

selectAll.addEventListener("change", () => {
  regionsEl
    .querySelectorAll("input[type=checkbox]:not(:disabled)")
    .forEach((box) => {
      box.checked = selectAll.checked;
    });
  syncGenerateState();
});

modeButtons.forEach((btn) => {
  btn.addEventListener("click", () => setMode(btn.dataset.mode));
});

builderSearch.addEventListener("focus", openBuilderMenu);
builderSearch.addEventListener("click", openBuilderMenu);
builderSearch.addEventListener("input", () => {
  const builder = selectedBuilder();
  if (!builder || builderSearch.value.trim() !== displayBuilder(builder)) {
    selectedKey = null;
    builderNote.textContent = "";
    currentDivisions = [];
    if (sheetType === "division") {
      regionsEl.innerHTML = `<p class="muted">Select a builder to see its divisions.</p>`;
      divisionNote.textContent = "";
    }
  }
  openBuilderMenu();
  syncGenerateState();
});
builderSearch.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeBuilderMenu();
  }
});

divisionSearch.addEventListener("input", renderDivisions);

document.addEventListener("click", (event) => {
  if (!combo.contains(event.target)) {
    closeBuilderMenu();
  }
});

generateBtn.addEventListener("click", async () => {
  const builder = selectedBuilder();
  statusEl.textContent = "Generating…";
  generateBtn.disabled = true;
  try {
    const regions =
      sheetType === "division" ? selectedRegions() : null;
    const res = await fetch("/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        builder_key: builder.builder_key,
        sheet_type: sheetType,
        regions,
        year: 2026,
        dry_run: false,
      }),
    });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.detail || "Generate failed");
    }
    const blob = await res.blob();
    const disposition = res.headers.get("content-disposition") || "";
    const match = disposition.match(/filename="?([^"]+)"?/i);
    const fallback =
      sheetType === "corporate"
        ? `GS_${builder.builder_key}_2026_corporate.pdf`
        : regions.length > 1
          ? `GS_${builder.builder_key}_2026_divisions.zip`
          : `GS_${builder.builder_key}_2026_division.pdf`;
    const filename = match ? match[1] : fallback;
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
    statusEl.textContent =
      sheetType === "corporate"
        ? "Corporate PDF downloaded (pages 1–2)."
        : regions.length > 1
          ? "Division PDFs downloaded as a zip (pages 3–4 each)."
          : "Division PDF downloaded (pages 3–4).";
  } catch (err) {
    statusEl.textContent = err.message;
  } finally {
    syncGenerateState();
  }
});

setMode("corporate");
ensureBuildersLoaded();
