const DEFAULT_HEIGHT = 340;
const extensionApi = globalThis.browser ?? globalThis.chrome;
const slider = document.querySelector("#height");
const value = document.querySelector("#value");
const heightOptions = document.querySelector("#height-options");
const modes = [...document.querySelectorAll('input[name="mode"]')];
const saveButton = document.querySelector("#save");
const saveStatus = document.querySelector("#save-status");

function show(height) {
  slider.value = height;
  value.textContent = `${height} px`;
}

function selectedMode() {
  return modes.find((input) => input.checked)?.value || "preview";
}

function showMode(mode) {
  modes.forEach((input) => input.checked = input.value === mode);
  heightOptions.hidden = mode !== "preview";
}

function setStatus(message, state = "") {
  saveStatus.textContent = message;
  saveStatus.dataset.state = state;
}

extensionApi.storage.local.get(["readingPaneHeight", "readerMode"]).then((stored) => {
  show(stored.readingPaneHeight || DEFAULT_HEIGHT);
  showMode(stored.readerMode === "accordion" ? "accordion" : "preview");
}).catch(() => {
  show(DEFAULT_HEIGHT);
  showMode("preview");
  setStatus("Could not read saved preferences.", "error");
});

slider.addEventListener("input", () => {
  show(Number(slider.value));
  setStatus("");
});

modes.forEach((input) => input.addEventListener("change", () => {
  if (!input.checked) return;
  showMode(input.value);
  setStatus("");
}));

saveButton.addEventListener("click", async () => {
  const preferences = {
    readingPaneHeight: Number(slider.value),
    readerMode: selectedMode()
  };

  saveButton.disabled = true;
  setStatus("Saving…");

  try {
    await extensionApi.storage.local.set(preferences);
    const saved = await extensionApi.storage.local.get(["readingPaneHeight", "readerMode"]);

    if (
      saved.readingPaneHeight !== preferences.readingPaneHeight ||
      saved.readerMode !== preferences.readerMode
    ) {
      throw new Error("Saved values could not be verified.");
    }

    setStatus("Saved in this browser.", "saved");
  } catch (error) {
    setStatus("Could not save preferences. Please try again.", "error");
  } finally {
    saveButton.disabled = false;
  }
});
