const DEFAULT_HEIGHT = 340;
const slider = document.querySelector("#height");
const value = document.querySelector("#value");
const heightOptions = document.querySelector("#height-options");
const modes = [...document.querySelectorAll('input[name="mode"]')];

function show(height) {
  slider.value = height;
  value.textContent = `${height} px`;
}

function showMode(mode) {
  modes.forEach((input) => input.checked = input.value === mode);
  heightOptions.hidden = mode !== "preview";
}

browser.storage.local.get(["readingPaneHeight", "readerMode"]).then((stored) => {
  show(stored.readingPaneHeight || DEFAULT_HEIGHT);
  showMode(stored.readerMode === "accordion" ? "accordion" : "preview");
});

slider.addEventListener("input", () => {
  const height = Number(slider.value);
  show(height);
  browser.storage.local.set({ readingPaneHeight: height });
});

modes.forEach((input) => input.addEventListener("change", () => {
  if (!input.checked) return;
  showMode(input.value);
  browser.storage.local.set({ readerMode: input.value });
}));
