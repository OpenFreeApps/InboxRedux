(() => {
  "use strict";

  // Firefox exposes the WebExtension API as `browser`; Chrome uses `chrome`.
  const extensionApi = globalThis.browser ?? globalThis.chrome;

  const DEFAULT_HEIGHT = 340;
  const MIN_HEIGHT = 180;
  const MIN_LIST_HEIGHT = 180;
  const STORAGE_KEY = "readingPaneHeight";
  const MODE_KEY = "readerMode";
  const PREVIEW_MODE = "preview";
  const ACCORDION_MODE = "accordion";
  const attachedDividers = new WeakSet();
  let preferredHeight = DEFAULT_HEIGHT;
  let readerMode = PREVIEW_MODE;
  let persistTimer;
  let activeAccordion;

  function getParts() {
    const readingPane = document.querySelector("#ReadingPaneContainerId");
    const readingRegion = readingPane?.parentElement;
    const layout = readingRegion?.parentElement;

    if (!readingPane || !readingRegion || !layout) return null;

    // Outlook has changed the wrappers around the list and the reading pane
    // several times. Derive the sibling parts from the reading pane instead
    // of relying on a fixed three-child layout.
    const children = [...layout.children];
    const readingIndex = children.indexOf(readingRegion);
    if (readingIndex < 2 || getComputedStyle(layout).flexDirection !== "column") return null;

    const divider = children[readingIndex - 1];
    const messageList = children.slice(0, readingIndex - 1)[0];
    if (!messageList || !divider) return null;

    return { readingRegion, layout, messageList, divider };
  }

  function clampHeight(parts, height) {
    const maximum = Math.max(MIN_HEIGHT, parts.layout.clientHeight - MIN_LIST_HEIGHT);
    return Math.round(Math.max(MIN_HEIGHT, Math.min(Number(height) || DEFAULT_HEIGHT, maximum)));
  }

  function applyHeight(parts, height = preferredHeight) {
    const safeHeight = clampHeight(parts, height);

    // The message list must consume the remaining vertical space; otherwise a
    // fixed reading-pane height can leave unused space below the message list.
    parts.messageList.style.setProperty("flex", "1 1 0px", "important");
    parts.messageList.style.setProperty("min-height", "0", "important");
    parts.readingRegion.style.setProperty("flex", `0 0 ${safeHeight}px`, "important");
    parts.readingRegion.style.setProperty("height", `${safeHeight}px`, "important");
  }

  function saveHeight(height) {
    preferredHeight = height;
    clearTimeout(persistTimer);
    persistTimer = setTimeout(() => {
      extensionApi.storage.local.set({ [STORAGE_KEY]: preferredHeight });
    }, 150);
  }

  function attachDivider(parts) {
    if (attachedDividers.has(parts.divider)) return;
    attachedDividers.add(parts.divider);

    let drag = null;

    parts.divider.addEventListener("pointerdown", (event) => {
      if (readerMode !== PREVIEW_MODE) return;
      if (event.button !== 0) return;

      const currentParts = getParts();
      if (!currentParts || currentParts.divider !== parts.divider) return;

      drag = { pointerId: event.pointerId, parts: currentParts };
      parts.divider.setPointerCapture?.(event.pointerId);

      // Outlook's own resizer cannot update values deliberately pinned by this
      // extension, so the divider updates our saved value instead.
      event.preventDefault();
      event.stopImmediatePropagation();
    }, true);

    parts.divider.addEventListener("pointermove", (event) => {
      if (!drag || event.pointerId !== drag.pointerId) return;

      const currentParts = getParts();
      if (!currentParts) return;

      const layoutRect = currentParts.layout.getBoundingClientRect();
      const height = clampHeight(currentParts, layoutRect.bottom - event.clientY);
      applyHeight(currentParts, height);
      preferredHeight = height;
      event.preventDefault();
      event.stopImmediatePropagation();
    }, true);

    const finishDrag = (event) => {
      if (!drag || event.pointerId !== drag.pointerId) return;
      saveHeight(preferredHeight);
      drag = null;
    };

    parts.divider.addEventListener("pointerup", finishDrag, true);
    parts.divider.addEventListener("pointercancel", finishDrag, true);
  }

  function removeAccordion() {
    activeAccordion?.remove();
    activeAccordion = null;
  }

  function getMessageRow(target, eventPath = []) {
    // Outlook has moved its message rows into web-component shadow trees in
    // some builds. target.closest() cannot cross that boundary, but a
    // composed event path retains the original row element.
    const rowSelector = '[role="option"], [data-convid], [data-message-id], [data-item-id]';

    const directRow = target instanceof Element ? target.closest(rowSelector) : null;
    if (directRow) return directRow;

    for (const node of eventPath) {
      if (!(node instanceof Element)) continue;
      if (node.matches(rowSelector)) return node;
      const ancestor = node.closest(rowSelector);
      if (ancestor) return ancestor;
    }

    return null;
  }

  function cloneReadingPane() {
    const pane = document.querySelector("#ReadingPaneContainerId");
    const source = pane?.firstElementChild;
    if (!source) return null;

    const copy = source.cloneNode(true);
    copy.querySelectorAll("[id]").forEach((element) => element.removeAttribute("id"));
    copy.querySelectorAll("button, a, input, textarea, select").forEach((element) => {
      element.setAttribute("tabindex", "-1");
    });
    return copy;
  }

  function openAccordion(row) {
    if (readerMode !== ACCORDION_MODE || !row.isConnected) return;
    if (activeAccordion?.previousElementSibling === row) {
      removeAccordion();
      return;
    }

    const content = cloneReadingPane();
    if (!content) return;

    removeAccordion();
    const accordion = document.createElement("section");
    accordion.className = "oh-accordion-reader";
    accordion.setAttribute("aria-label", "Opened message");
    accordion.innerHTML = '<div class="oh-accordion-toolbar"><span>Message preview</span><button type="button" class="oh-accordion-close" aria-label="Close message preview">×</button></div>';
    accordion.append(content);
    accordion.querySelector(".oh-accordion-close").addEventListener("click", removeAccordion);
    row.insertAdjacentElement("afterend", accordion);
    activeAccordion = accordion;
  }

  function attachAccordionListener() {
    if (document.documentElement.dataset.ohAccordionListener === "1") return;
    document.documentElement.dataset.ohAccordionListener = "1";

    document.addEventListener("dblclick", (event) => {
      if (readerMode !== ACCORDION_MODE) return;
      const row = getMessageRow(event.target, event.composedPath());
      if (!row) return;

      // Outlook's default double-click action opens a compose-like pop-out.
      // Accordion Reader owns this gesture, while normal single-click and
      // multi-select behavior remain untouched.
      event.preventDefault();
      event.stopImmediatePropagation();

      // The first click selects the message and Outlook renders its native
      // reading pane. Clone that already-rendered content after it settles.
      window.setTimeout(() => openAccordion(row), 450);
    }, true);
  }

  function applyAccordionMode(parts) {
    // Attach regardless of whether Outlook's current wrappers match our
    // resizable-pane layout. That keeps the reader usable across OWA updates.
    attachAccordionListener();

    if (!parts) {
      const readingRegion = document.querySelector("#ReadingPaneContainerId")?.parentElement;
      readingRegion?.style.setProperty("display", "none", "important");
      return;
    }

    parts.readingRegion.style.setProperty("display", "none", "important");
    parts.divider.style.setProperty("display", "none", "important");
    parts.messageList.style.setProperty("flex", "1 1 0px", "important");
    parts.messageList.style.setProperty("min-height", "0", "important");
  }

  function applyPreviewMode(parts) {
    removeAccordion();
    const readingRegion = document.querySelector("#ReadingPaneContainerId")?.parentElement;
    readingRegion?.style.removeProperty("display");
    if (!parts) return;

    parts.divider.style.removeProperty("display");
    applyHeight(parts);
    attachDivider(parts);
  }

  function update() {
    const parts = getParts();
    if (readerMode === ACCORDION_MODE) applyAccordionMode(parts);
    else applyPreviewMode(parts);
  }

  extensionApi.storage.local.get([STORAGE_KEY, MODE_KEY]).then((stored) => {
    if (Number.isFinite(stored[STORAGE_KEY])) preferredHeight = stored[STORAGE_KEY];
    if (stored[MODE_KEY] === ACCORDION_MODE) readerMode = ACCORDION_MODE;
    update();
  });

  extensionApi.storage.onChanged.addListener((changes, areaName) => {
    if (areaName !== "local") return;
    if (changes[STORAGE_KEY]) preferredHeight = changes[STORAGE_KEY].newValue || DEFAULT_HEIGHT;
    if (changes[MODE_KEY]) readerMode = changes[MODE_KEY].newValue === ACCORDION_MODE ? ACCORDION_MODE : PREVIEW_MODE;
    update();
  });

  // Outlook replaces large portions of its DOM during navigation and message
  // selection. Reapply only after that work settles.
  let updateTimer;
  new MutationObserver(() => {
    clearTimeout(updateTimer);
    updateTimer = setTimeout(update, 80);
  }).observe(document.documentElement, { childList: true, subtree: true });

  window.addEventListener("resize", update);
})();
