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
  const attachedLiveScrollRegions = new WeakSet();
  let preferredHeight = DEFAULT_HEIGHT;
  let readerMode = PREVIEW_MODE;
  let persistTimer;
  let activeAccordion;
  let activeAccordionKey;
  let activeAccordionRow;
  let liveReader;
  let accordionPausedForOutlookAction = false;

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

  function restoreLiveReader() {
    if (!liveReader) return;

    const { readingRegion } = liveReader;
    readingRegion.classList.remove("oh-accordion-live-region");
    for (const property of [
      "display", "position", "inset", "top", "left", "right", "bottom", "width",
      "height", "visibility", "pointer-events", "z-index"
    ]) {
      readingRegion.style.removeProperty(property);
    }
    liveReader = null;
  }

  function removeAccordion() {
    restoreLiveReader();
    activeAccordion?.remove();
    activeAccordion = null;
    activeAccordionKey = null;
    activeAccordionRow = null;
  }

  function restoreNormalOutlookLayout() {
    const parts = getParts();
    const readingRegion = document.querySelector("#ReadingPaneContainerId")?.parentElement;

    readingRegion?.style.removeProperty("display");
    readingRegion?.style.removeProperty("position");
    readingRegion?.style.removeProperty("inset");
    readingRegion?.style.removeProperty("top");
    readingRegion?.style.removeProperty("left");
    readingRegion?.style.removeProperty("right");
    readingRegion?.style.removeProperty("bottom");
    readingRegion?.style.removeProperty("width");
    readingRegion?.style.removeProperty("height");
    readingRegion?.style.removeProperty("visibility");
    readingRegion?.style.removeProperty("pointer-events");
    readingRegion?.style.removeProperty("z-index");

    if (!parts) return;
    parts.layout.style.removeProperty("position");
    parts.divider.style.removeProperty("display");
    for (const property of ["flex", "height", "max-height", "min-height"]) {
      parts.messageList.style.removeProperty(property);
    }
  }

  function handOffToOutlook() {
    accordionPausedForOutlookAction = true;
    removeAccordion();
    restoreNormalOutlookLayout();
  }


  function getMessageRowKey(row) {
    for (const attribute of ["data-convid", "data-message-id", "data-item-id"]) {
      const value = row.getAttribute(attribute);
      if (value) return { attribute, value };
    }

    return null;
  }

  function matchesMessageRowKey(row, key) {
    return Boolean(key && row.getAttribute(key.attribute) === key.value);
  }

  function findMessageRow(key) {
    if (!key) return null;

    return [...document.querySelectorAll(`[${key.attribute}]`)]
      .find((row) => matchesMessageRowKey(row, key)) || null;
  }

  function restoreAccordionIfNeeded() {
    if (!activeAccordion || activeAccordion.isConnected) return;

    const row = findMessageRow(activeAccordionKey);
    if (row) row.insertAdjacentElement("afterend", activeAccordion);
  }

  function isActiveAccordionRow(row) {
    return activeAccordionRow === row ||
      matchesMessageRowKey(row, activeAccordionKey);
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

  function getScrollableElement(root) {
    const candidates = [root, ...root.querySelectorAll("*")];

    return candidates.find((element) => {
      const overflowY = getComputedStyle(element).overflowY;
      return element.scrollHeight > element.clientHeight &&
        (overflowY === "auto" || overflowY === "scroll");
    }) || root;
  }

  function canScroll(element, deltaY) {
    const maximum = Math.max(0, element.scrollHeight - element.clientHeight);
    return deltaY < 0 ? element.scrollTop > 0 : element.scrollTop < maximum - 1;
  }

  function getAccordionMetrics(parts, row) {
    const layoutRect = parts.layout.getBoundingClientRect();
    const rowRect = row.getBoundingClientRect();
    const listRect = parts.messageList.getBoundingClientRect();
    const listScroller = getScrollableElement(parts.messageList);
    const scrollerRect = listScroller.getBoundingClientRect();
    const scrollbarWidth = Math.max(0, listScroller.offsetWidth - listScroller.clientWidth);
    const toolbarHeight = 33;
    const gutter = 0;
    const left = Math.max(0, listRect.left - layoutRect.left);
    // Use the scrollable content edge. The wrapper's rectangle includes the
    // scrollbar track, which Outlook leaves outside the message rows.
    const contentRight = Math.min(layoutRect.right, scrollerRect.right - scrollbarWidth);
    const width = Math.max(0, contentRight - listRect.left);
    const top = Math.max(0, rowRect.bottom - layoutRect.top);
    const availableHeight = Math.max(
      MIN_HEIGHT,
      Math.min(680, layoutRect.height - top - toolbarHeight - gutter)
    );

    return { top, left, width, toolbarHeight, availableHeight, gutter };
  }

  function positionLiveReader(parts, row) {
    if (!liveReader || !row?.isConnected) return;

    const { top, left, width, toolbarHeight, availableHeight } = getAccordionMetrics(parts, row);
    activeAccordion.style.setProperty("top", `${top}px`, "important");
    activeAccordion.style.setProperty("left", `${left}px`, "important");
    activeAccordion.style.setProperty("right", "auto", "important");
    activeAccordion.style.setProperty("width", `${width}px`, "important");

    const readerTop = top + toolbarHeight;
    liveReader.readingRegion.style.setProperty("top", `${readerTop}px`, "important");
    liveReader.readingRegion.style.setProperty("left", `${left}px`, "important");
    liveReader.readingRegion.style.setProperty("right", "auto", "important");
    liveReader.readingRegion.style.setProperty("width", `${width}px`, "important");
    liveReader.readingRegion.style.setProperty("height", `${availableHeight}px`, "important");
  }

  function attachLiveReaderScroll(parts) {
    if (attachedLiveScrollRegions.has(parts.readingRegion)) return;
    attachedLiveScrollRegions.add(parts.readingRegion);

    parts.readingRegion.addEventListener("wheel", (event) => {
      if (
        readerMode !== ACCORDION_MODE ||
        event.ctrlKey ||
        liveReader?.readingRegion !== parts.readingRegion
      ) return;

      const readerScrollTarget = getScrollableElement(parts.readingRegion);
      if (canScroll(readerScrollTarget, event.deltaY)) return;

      const inboxScrollTarget = getScrollableElement(parts.messageList);
      if (!canScroll(inboxScrollTarget, event.deltaY)) return;

      inboxScrollTarget.scrollTop += event.deltaY;
      event.preventDefault();
      event.stopPropagation();
    }, { capture: true, passive: false });
  }

  function openAccordion(row) {
    if (readerMode !== ACCORDION_MODE || !row.isConnected) return;
    if (isActiveAccordionRow(row)) {
      removeAccordion();
      return;
    }

    const parts = getParts();
    if (!parts) return;

    removeAccordion();
    const accordion = document.createElement("section");
    accordion.className = "oh-accordion-reader";
    accordion.setAttribute("aria-label", "Opened message");
    accordion.innerHTML = '<div class="oh-accordion-toolbar"><span>Message preview</span><button type="button" class="oh-accordion-close" aria-label="Close message preview">×</button></div>';
    accordion.querySelector(".oh-accordion-close").addEventListener("click", () => {
      removeAccordion();
      // With no opened item, remove the native reader from layout entirely.
      // This avoids leaving an invisible absolute pane over the message list.
      const currentParts = getParts();
      currentParts?.readingRegion.style.setProperty("display", "none", "important");
    });
    parts.layout.append(accordion);

    // Keep the actual Outlook reader in its original DOM location. This
    // preserves built-in controls, selection, and message-specific behavior.
    parts.readingRegion.classList.add("oh-accordion-live-region");
    parts.readingRegion.style.setProperty("display", "block", "important");
    parts.readingRegion.style.setProperty("position", "absolute", "important");
    parts.readingRegion.style.setProperty("inset", "auto", "important");
    parts.readingRegion.style.setProperty("visibility", "visible", "important");
    parts.readingRegion.style.setProperty("pointer-events", "auto", "important");
    parts.readingRegion.style.setProperty("z-index", "6", "important");

    activeAccordion = accordion;
    activeAccordionKey = getMessageRowKey(row);
    activeAccordionRow = row;
    liveReader = { parts, readingRegion: parts.readingRegion };
    positionLiveReader(parts, row);
    attachLiveReaderScroll(parts);
  }
  function attachAccordionListener() {
    if (document.documentElement.dataset.ohAccordionListener === "1") return;
    document.documentElement.dataset.ohAccordionListener = "1";

    document.addEventListener("click", (event) => {
      if (readerMode !== ACCORDION_MODE || !activeAccordion) return;
      const row = getMessageRow(event.target, event.composedPath());
      if (row && !isActiveAccordionRow(row)) removeAccordion();
    }, true);

    document.addEventListener("click", (event) => {
      if (readerMode !== ACCORDION_MODE || !liveReader) return;
      const control = event.composedPath().find((node) =>
        node instanceof Element && node.matches("button, [role=button], a")
      ) || (event.target instanceof Element
        ? event.target.closest("button, [role=button], a")
        : null);
      const label = [
        control?.getAttribute("aria-label"),
        control?.getAttribute("title"),
        control?.textContent,
        ...event.composedPath()
          .filter((node) => node instanceof Element)
          .map((node) => node.getAttribute?.("aria-label") || node.getAttribute?.("title"))
      ].filter(Boolean).join(" ").trim();

      // These actions replace the reader with Outlook's compose surface.
      // Restore its normal layout before Outlook handles the command.
      if (/\b(reply(?: all)?|forward)\b/i.test(label)) handOffToOutlook();
    }, true);

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
      accordionPausedForOutlookAction = false;
      window.setTimeout(() => openAccordion(row), 450);
    }, true);
  }

  function applyAccordionMode(parts) {
    // Attach regardless of whether Outlook's current wrappers match our
    // resizable-pane layout. That keeps the reader usable across OWA updates.
    attachAccordionListener();

    if (accordionPausedForOutlookAction) {
      restoreNormalOutlookLayout();
      return;
    }

    if (!parts) {
      const readingRegion = document.querySelector("#ReadingPaneContainerId")?.parentElement;
      readingRegion?.style.setProperty("display", "none", "important");
      return;
    }

    parts.layout.style.setProperty("position", "relative", "important");
    parts.divider.style.setProperty("display", "none", "important");
    parts.messageList.style.setProperty("flex", "1 1 0px", "important");
    parts.messageList.style.setProperty("height", "auto", "important");
    parts.messageList.style.setProperty("max-height", "none", "important");
    parts.messageList.style.setProperty("min-height", "0", "important");

    // Once opened, the live reader belongs to the accordion until the user
    // closes it or opens another message.
    if (liveReader?.readingRegion === parts.readingRegion) {
      positionLiveReader(parts, activeAccordionRow);
      return;
    }

    // With no accordion open, remove the native reader from layout. It is
    // restored by openAccordion after Outlook selects the double-clicked row.
    parts.readingRegion.style.setProperty("display", "none", "important");
    parts.readingRegion.style.removeProperty("position");
    parts.readingRegion.style.removeProperty("inset");
    parts.readingRegion.style.removeProperty("width");
    parts.readingRegion.style.removeProperty("height");
    parts.readingRegion.style.removeProperty("visibility");
    parts.readingRegion.style.removeProperty("pointer-events");
  }

  function applyPreviewMode(parts) {
    removeAccordion();
    const readingRegion = document.querySelector("#ReadingPaneContainerId")?.parentElement;
    readingRegion?.style.removeProperty("display");
    readingRegion?.style.removeProperty("position");
    readingRegion?.style.removeProperty("inset");
    readingRegion?.style.removeProperty("height");
    readingRegion?.style.removeProperty("width");
    readingRegion?.style.removeProperty("visibility");
    readingRegion?.style.removeProperty("pointer-events");
    if (!parts) return;

    parts.layout.style.removeProperty("position");
    parts.divider.style.removeProperty("display");
    parts.messageList.style.removeProperty("height");
    parts.messageList.style.removeProperty("max-height");
    applyHeight(parts);
    attachDivider(parts);
  }

  function update() {
    const parts = getParts();
    if (readerMode === ACCORDION_MODE) {
      applyAccordionMode(parts);
      restoreAccordionIfNeeded();
    } else {
      applyPreviewMode(parts);
    }
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
