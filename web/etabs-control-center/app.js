const state = {
  current: null,
  parsed: null,
  parseError: null,
  hasChanges: false,
  parseTimer: null,
};

const elements = {
  statusDot: document.getElementById("statusDot"),
  statusText: document.getElementById("statusText"),
  clockText: document.getElementById("clockText"),
  modelTitle: document.getElementById("modelTitle"),
  currentBaseElevation: document.getElementById("currentBaseElevation"),
  currentStoryCount: document.getElementById("currentStoryCount"),
  physicalCounts: document.getElementById("physicalCounts"),
  presentUnits: document.getElementById("presentUnits"),
  storyEditor: document.getElementById("storyEditor"),
  saveToggle: document.getElementById("saveToggle"),
  unlockToggle: document.getElementById("unlockToggle"),
  previewBadge: document.getElementById("previewBadge"),
  previewSummary: document.getElementById("previewSummary"),
  previewTableBody: document.getElementById("previewTableBody"),
  logOutput: document.getElementById("logOutput"),
  refreshButton: document.getElementById("refreshButton"),
  resetButton: document.getElementById("resetButton"),
  copyCurrentButton: document.getElementById("copyCurrentButton"),
  applyButton: document.getElementById("applyButton"),
};

function setClock() {
  elements.clockText.textContent = new Date().toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
}

function log(message, tone = "info") {
  const line = document.createElement("div");
  line.className = "log-line";

  const time = document.createElement("span");
  time.className = "log-time";
  time.textContent = new Date().toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });

  const text = document.createElement("span");
  text.className = `log-message ${tone}`;
  text.textContent = message;

  line.append(time, text);
  elements.logOutput.prepend(line);
}

function setStatus(text, tone = "loading") {
  elements.statusText.textContent = text.toUpperCase();
  elements.statusDot.classList.remove("ready", "error");
  if (tone === "ready") {
    elements.statusDot.classList.add("ready");
  } else if (tone === "error") {
    elements.statusDot.classList.add("error");
  }
}

function formatNumber(value) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) {
    return "-";
  }

  return Number(value).toFixed(3).replace(/\.?0+$/, "");
}

function unitsLabel(unitsEnum) {
  const lookup = {
    1: "LB-IN",
    2: "LB-FT",
    3: "KIP-IN",
    4: "KIP-FT",
    5: "KN-MM",
    6: "KN-CM",
    7: "KN-M",
    8: "N-MM",
    9: "N-CM",
    10: "N-M",
    11: "TF-M",
    12: "TF-CM",
    13: "TF-MM",
    14: "KGF-M",
    15: "KGF-CM",
    16: "KGF-MM",
  };

  return lookup[unitsEnum] || String(unitsEnum ?? "-");
}

function buildEditorText(stories) {
  if (!state.current) {
    return "";
  }

  const lines = [`Base Elevation Ft\t${formatNumber(state.current.BaseElevation)}`, "Story\tElevation Ft"];
  (stories || []).forEach((story) => {
    lines.push(`${story.Name ?? story.name}\t${formatNumber(story.Elevation ?? story.elevation)}`);
  });
  return lines.join("\n");
}

function buildTableRows(stories, emptyText) {
  if (!stories || stories.length === 0) {
    return `<tr><td colspan="4" class="empty-cell">${emptyText}</td></tr>`;
  }

  return stories
    .map((story, index) => {
      const name = story.Name ?? story.name;
      const elevation = story.Elevation ?? story.elevation;
      const height = story.Height ?? story.height;
      const rowIndex = story.Index ?? story.index ?? index;

      return `
        <tr>
          <td>${rowIndex + 1}</td>
          <td>${name}</td>
          <td class="num">${formatNumber(elevation)}</td>
          <td class="num">${formatNumber(height)}</td>
        </tr>
      `;
    })
    .join("");
}

function nearlyEqual(left, right, tolerance = 1e-6) {
  return Math.abs(Number(left) - Number(right)) <= tolerance;
}

function areStoriesEquivalent(parsedStories, currentStories) {
  if (!parsedStories || !currentStories || parsedStories.length !== currentStories.length) {
    return false;
  }

  for (let index = 0; index < parsedStories.length; index += 1) {
    const parsed = parsedStories[index];
    const current = currentStories[index];
    if ((parsed.name ?? "").trim() !== (current.Name ?? "").trim()) {
      return false;
    }
    if (!nearlyEqual(parsed.elevation, current.Elevation)) {
      return false;
    }
  }

  return true;
}

function computeHasChanges() {
  if (!state.current || !state.parsed || !state.parsed.stories) {
    return false;
  }

  if (state.parsed.baseElevation === null || state.parsed.baseElevation === undefined) {
    return false;
  }

  return !nearlyEqual(state.parsed.baseElevation, state.current.BaseElevation) ||
    !areStoriesEquivalent(state.parsed.stories, state.current.Stories);
}

function renderCurrentMeta() {
  if (!state.current) {
    return;
  }

  elements.modelTitle.textContent = state.current.ModelPath || "UNSAVED ETABS MODEL";
  elements.currentBaseElevation.textContent = formatNumber(state.current.BaseElevation);
  elements.currentStoryCount.textContent = String(state.current.StoryCount);
  elements.physicalCounts.textContent = `F ${state.current.PhysicalCounts.Frames}  A ${state.current.PhysicalCounts.Areas}  P ${state.current.PhysicalCounts.Points}`;
  elements.presentUnits.textContent = (state.current.LengthUnit || unitsLabel(state.current.PresentUnitsEnum)).toUpperCase();
}

function renderPreview() {
  if (state.parseError) {
    elements.previewBadge.textContent = "PARSE ERROR";
    elements.previewSummary.textContent = state.parseError.toUpperCase();
    elements.previewTableBody.innerHTML = buildTableRows([], "EDITOR CONTENT COULD NOT BE PARSED.");
    return;
  }

  if (!state.parsed || !state.parsed.stories || state.parsed.stories.length === 0) {
    elements.previewBadge.textContent = "NO ROWS";
    elements.previewSummary.textContent = "EDIT THE CURRENT STORIES PANEL TO PREVIEW A VALIDATED STORY SET.";
    elements.previewTableBody.innerHTML = buildTableRows([], "NO PARSED STORY DATA YET.");
    return;
  }

  const badgeText = state.hasChanges ? `${state.parsed.storyCount} ROWS CHANGED` : `${state.parsed.storyCount} ROWS LIVE`;
  elements.previewBadge.textContent = badgeText;

  const notes = [];
  if (state.parsed.baseElevation !== null && state.parsed.baseElevation !== undefined) {
    notes.push(`BASE ${formatNumber(state.parsed.baseElevation)} FT`);
  }
  if (state.parsed.orderChanged) {
    notes.push("REORDERED LOW TO HIGH");
  }
  if (state.hasChanges) {
    notes.push("UPDATE BUTTON ARMED");
  } else {
    notes.push("NO DIFF FROM LIVE MODEL");
  }
  if (state.parsed.warnings && state.parsed.warnings.length) {
    notes.push(...state.parsed.warnings.map((warning) => warning.toUpperCase()));
  }
  elements.previewSummary.textContent = notes.join(" | ");
  elements.previewTableBody.innerHTML = buildTableRows(state.parsed.stories, "NO PARSED STORY DATA YET.");
}

function renderApplyButton() {
  const canApply = !state.parseError && state.parsed && state.parsed.stories && state.parsed.stories.length > 0 && state.hasChanges;
  elements.applyButton.disabled = !canApply;
  elements.applyButton.classList.toggle("button-armed", canApply);
}

async function requestJson(url, options = {}) {
  const response = await fetch(url, {
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
    ...options,
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || `Request failed: ${response.status}`);
  }
  return payload;
}

function loadEditorFromCurrent(logMessage = true) {
  if (!state.current) {
    return;
  }

  elements.storyEditor.value = buildEditorText(state.current.Stories);
  if (logMessage) {
    log("Loaded live ETABS stories into the editable current-stories surface.");
  }
}

async function analyzeEditor() {
  const text = elements.storyEditor.value;
  const baseElevation = state.current ? Number(state.current.BaseElevation) : null;

  if (!text.trim()) {
    state.parsed = null;
    state.parseError = "The current stories editor is blank.";
    state.hasChanges = false;
    renderPreview();
    renderApplyButton();
    return;
  }

  try {
    const parsed = await requestJson("/api/stories/parse", {
      method: "POST",
      body: JSON.stringify({ text, baseElevation }),
    });
    state.parsed = parsed;
    state.parseError = null;
    state.hasChanges = computeHasChanges();
  } catch (error) {
    state.parsed = null;
    state.parseError = error.message;
    state.hasChanges = false;
  }

  renderPreview();
  renderApplyButton();
}

function queueAnalyzeEditor() {
  if (state.parseTimer) {
    window.clearTimeout(state.parseTimer);
  }

  state.parseTimer = window.setTimeout(() => {
    analyzeEditor().catch((error) => {
      state.parseError = error.message;
      state.parsed = null;
      state.hasChanges = false;
      renderPreview();
      renderApplyButton();
    });
  }, 140);
}

async function refreshCurrentStories({ overwriteEditor = true } = {}) {
  setStatus("READING LIVE ETABS STORIES", "loading");
  const current = await requestJson("/api/stories/current");
  state.current = current;
  renderCurrentMeta();

  if (overwriteEditor) {
    loadEditorFromCurrent(false);
  }

  await analyzeEditor();
  setStatus("CONNECTED TO LIVE ETABS SESSION", "ready");
  log(`Loaded ${current.StoryCount} current ETABS stories from the active model.`);
}

async function applyStories() {
  if (!state.parsed || !state.parsed.stories || state.parsed.stories.length === 0) {
    await analyzeEditor();
  }

  if (state.parseError || !state.hasChanges) {
    return;
  }

  const baseElevation = Number(state.parsed.baseElevation);
  const confirmation = window.confirm(
    `Update the ETABS model with ${state.parsed.storyCount} stories at base elevation ${formatNumber(baseElevation)} ft?`
  );
  if (!confirmation) {
    return;
  }

  log("Updating ETABS story definitions from the edited current-stories surface.");
  const result = await requestJson("/api/stories/apply", {
    method: "POST",
    body: JSON.stringify({
      stories: state.parsed.stories,
      baseElevation,
      save: elements.saveToggle.checked,
      unlockIfLocked: elements.unlockToggle.checked,
    }),
  });

  log(`Updated ETABS to ${result.PostStoryCount} stories.`);
  if (result.BackupCreated && result.BackupPath) {
    log(`Backup created at ${result.BackupPath}.`);
  }

  await refreshCurrentStories({ overwriteEditor: true });
}

async function copyEditorStories() {
  await navigator.clipboard.writeText(elements.storyEditor.value);
  log("Copied the editable current-stories surface to the clipboard.");
}

elements.refreshButton.addEventListener("click", async () => {
  try {
    const shouldOverwrite = !state.hasChanges || window.confirm("Overwrite editor changes with the live ETABS story set?");
    if (!shouldOverwrite) {
      return;
    }
    await refreshCurrentStories({ overwriteEditor: true });
  } catch (error) {
    setStatus("ETABS CONNECTION FAILED", "error");
    log(error.message, "error");
  }
});

elements.resetButton.addEventListener("click", async () => {
  if (!state.current) {
    return;
  }
  loadEditorFromCurrent();
  await analyzeEditor();
});

elements.copyCurrentButton.addEventListener("click", async () => {
  try {
    await copyEditorStories();
  } catch (error) {
    log(error.message, "error");
  }
});

elements.applyButton.addEventListener("click", async () => {
  try {
    await applyStories();
  } catch (error) {
    log(error.message, "error");
  }
});

elements.storyEditor.addEventListener("input", () => {
  queueAnalyzeEditor();
});

window.addEventListener("load", async () => {
  setClock();
  window.setInterval(setClock, 1000);

  try {
    log("ETABS control center ready. Refreshing live model context.");
    await refreshCurrentStories({ overwriteEditor: true });
  } catch (error) {
    setStatus("ETABS CONNECTION FAILED", "error");
    log(error.message, "error");
    renderPreview();
    renderApplyButton();
  }
});
