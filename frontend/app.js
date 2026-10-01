const combo = document.getElementById("builder-combo");
const builderSearch = document.getElementById("builder-search");
const builderOptions = document.getElementById("builder-options");
const builderNote = document.getElementById("builder-note");
const regionsEl = document.getElementById("regions");
const selectAll = document.getElementById("select-all");
const generateBtn = document.getElementById("generate");
const statusEl = document.getElementById("status");

let builders = [];
let selectedKey = null;

function displayName(builder) {
  return builder.has_prior_gold_sheet
    ? builder.builder_name
    : `${builder.builder_name} (no prior sheet)`;
}

function selectedBuilder() {
  return builders.find((b) => b.builder_key === selectedKey) || null;
}

function selectedRegions() {
  return [...regionsEl.querySelectorAll("input[type=checkbox]:checked")].map(
    (box) => box.value
  );
}

function listQuery() {
  const builder = selectedBuilder();
  const text = builderSearch.value.trim().toLowerCase();
  if (builder && text === displayName(builder).toLowerCase()) {
    return "";
  }
  return text;
}

function openMenu() {
  const query = listQuery();
  const matches = builders.filter((b) =>
    displayName(b).toLowerCase().includes(query)
  );
  builderOptions.innerHTML = "";
  matches.forEach((builder) => {
    const item = document.createElement("li");
    const button = document.createElement("button");
    button.type = "button";
    button.setAttribute("role", "option");
    button.textContent = displayName(builder);
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
    item.textContent = "No matching builders";
    item.style.padding = "8px 10px";
    builderOptions.appendChild(item);
  }
  builderOptions.hidden = false;
  builderSearch.setAttribute("aria-expanded", "true");
}

function closeMenu() {
  builderOptions.hidden = true;
  builderSearch.setAttribute("aria-expanded", "false");
}

function chooseBuilder(builderKey) {
  const builder = builders.find((b) => b.builder_key === builderKey);
  selectedKey = builder ? builderKey : null;
  builderSearch.value = builder ? displayName(builder) : "";
  closeMenu();
  loadRegions();
}

async function loadRegions() {
  const builder = selectedBuilder();
  regionsEl.innerHTML = "";
  selectAll.checked = false;
  generateBtn.disabled = true;
  if (!builder) {
    builderNote.textContent = "";
    regionsEl.innerHTML = `<p class="muted">Select a builder to see regions.</p>`;
    return;
  }
  builderNote.textContent = builder.has_prior_gold_sheet
    ? `${builder.region_count} regions`
    : "This builder has account rows, but no prior Gold Sheet yet, so a PDF cannot be generated.";

  const res = await fetch(`/builders/${encodeURIComponent(builder.builder_key)}/regions`);
  const regions = await res.json();
  if (!regions.length) {
    regionsEl.innerHTML = `<p class="muted">No regions in the Salesforce extract.</p>`;
    return;
  }
  regions.forEach((region) => {
    const label = document.createElement("label");
    label.className = "region";
    const box = document.createElement("input");
    box.type = "checkbox";
    box.value = region.region_label;
    box.addEventListener("change", syncSelectAll);
    label.append(box, document.createTextNode(` ${region.region_label}`));
    regionsEl.appendChild(label);
  });
  selectAll.disabled = !builder.has_prior_gold_sheet;
}

function syncSelectAll() {
  const boxes = [...regionsEl.querySelectorAll("input[type=checkbox]")];
  selectAll.checked = boxes.length > 0 && boxes.every((box) => box.checked);
  const builder = selectedBuilder();
  generateBtn.disabled = !(
    builder &&
    builder.has_prior_gold_sheet &&
    selectedRegions().length
  );
}

selectAll.addEventListener("change", () => {
  regionsEl.querySelectorAll("input[type=checkbox]").forEach((box) => {
    box.checked = selectAll.checked;
  });
  syncSelectAll();
});

builderSearch.addEventListener("focus", openMenu);
builderSearch.addEventListener("click", openMenu);
builderSearch.addEventListener("input", () => {
  const builder = selectedBuilder();
  if (!builder || builderSearch.value.trim() !== displayName(builder)) {
    selectedKey = null;
    loadRegions();
  }
  openMenu();
});
builderSearch.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeMenu();
  }
});

document.addEventListener("click", (event) => {
  if (!combo.contains(event.target)) {
    closeMenu();
  }
});

generateBtn.addEventListener("click", async () => {
  const builder = selectedBuilder();
  const regions = selectedRegions();
  const boxes = [...regionsEl.querySelectorAll("input[type=checkbox]")];
  statusEl.textContent = "Generating…";
  generateBtn.disabled = true;
  try {
    const res = await fetch("/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        builder_key: builder.builder_key,
        regions: regions.length === boxes.length ? null : regions,
        year: 2026,
        dry_run: false,
      }),
    });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.detail || "Generate failed");
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `GS_${builder.builder_key}_2026.pdf`;
    link.click();
    URL.revokeObjectURL(url);
    statusEl.textContent = "PDF downloaded.";
  } catch (err) {
    statusEl.textContent = err.message;
  } finally {
    syncSelectAll();
  }
});

fetch("/builders")
  .then((res) => res.json())
  .then((data) => {
    builders = data;
  })
  .catch(() => {
    statusEl.textContent = "Could not load builders. Is the API running?";
  });
