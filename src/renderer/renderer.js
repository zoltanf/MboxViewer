const openButton = document.getElementById("openButton");
const remoteContentButton = document.getElementById("remoteContentButton");
const searchInput = document.getElementById("searchInput");
const searchWrap = document.getElementById("searchWrap");
const searchClearButton = document.getElementById("searchClearButton");
const filterToolsIcon = document.getElementById("filterToolsIcon");
const dateFilterContainer = document.getElementById("dateFilterContainer");
const dateFilterButton = document.getElementById("dateFilterButton");
const dateFilter = document.getElementById("dateFilter");
const dateFilterClearButton = document.getElementById("dateFilterClearButton");
const fromFilterContainer = document.getElementById("fromFilterContainer");
const fromFilterButton = document.getElementById("fromFilterButton");
const fromFilter = document.getElementById("fromFilter");
const fromFilterInput = document.getElementById("fromFilterInput");
const fromFilterClearButton = document.getElementById("fromFilterClearButton");
const toFilterContainer = document.getElementById("toFilterContainer");
const toFilterButton = document.getElementById("toFilterButton");
const toFilter = document.getElementById("toFilter");
const toFilterInput = document.getElementById("toFilterInput");
const toFilterClearButton = document.getElementById("toFilterClearButton");
const subjectFilterContainer = document.getElementById("subjectFilterContainer");
const subjectFilterButton = document.getElementById("subjectFilterButton");
const subjectFilter = document.getElementById("subjectFilter");
const subjectFilterInput = document.getElementById("subjectFilterInput");
const subjectFilterClearButton = document.getElementById("subjectFilterClearButton");
const attachmentToggleContainer = document.getElementById("attachmentToggleContainer");
const attachmentToggleButton = document.getElementById("attachmentToggleButton");
const dateBoundsLabel = document.getElementById("dateBoundsLabel");
const dateFromLabel = document.getElementById("dateFromLabel");
const dateToLabel = document.getElementById("dateToLabel");
const dateRangeFill = document.getElementById("dateRangeFill");
const dateFromRange = document.getElementById("dateFromRange");
const dateToRange = document.getElementById("dateToRange");
const statusMessage = document.getElementById("statusMessage");
const statusMeta = document.getElementById("statusMeta");
const openProgress = document.getElementById("openProgress");
const openProgressBar = document.getElementById("openProgressBar");
const mailListPanel = document.querySelector(".mail-list-panel");
const mailList = document.getElementById("mailList");
const messageView = document.getElementById("messageView");
const externalLinkModal = document.getElementById("externalLinkModal");
const externalLinkUrl = document.getElementById("externalLinkUrl");
const externalLinkCancel = document.getElementById("externalLinkCancel");
const externalLinkCopy = document.getElementById("externalLinkCopy");
const externalLinkOpen = document.getElementById("externalLinkOpen");
const emlSourceModal = document.getElementById("emlSourceModal");
const emlSourceClose = document.getElementById("emlSourceClose");
const emlSourceContent = document.getElementById("emlSourceContent");
const layout = document.querySelector(".layout");
const splitter = document.getElementById("splitter");

const PAGE_SIZE = 200;
const SEARCH_DEBOUNCE_MS = 220;
const DAY_MS = 24 * 60 * 60 * 1000;

let dbPath = "";
let mboxPath = "";
let totalMessages = 0;
let totalResults = 0;
let currentQuery = "";
let currentOffset = 0;
let selectedMessageId = null;
let currentPageMessages = [];
let currentStandaloneMessage = null;
let resultIndexById = new Map();
let searchDebounceTimer = null;
let requestToken = 0;
let messageRequestToken = 0;
let openingInProgress = false;
let loadingMoreMessages = false;
let dateRangeMinDayTs = null;
let dateRangeMaxDayTs = null;
let dateFromDayOffset = 0;
let dateToDayOffset = 0;
let dateFilterPopoverOpen = false;
let fromFilterPopoverOpen = false;
let toFilterPopoverOpen = false;
let subjectFilterPopoverOpen = false;
let attachmentsOnlyFilterEnabled = false;
let remoteContentEnabled = false;
let externalLinkModalOpen = false;
let emlSourceModalOpen = false;
let pendingExternalUrl = "";

openButton.addEventListener("click", openMbox);
if (remoteContentButton) {
  remoteContentButton.addEventListener("click", toggleRemoteContent);
}
searchInput.addEventListener("input", onSearchInput);
if (searchClearButton) {
  searchClearButton.addEventListener("click", onSearchClearClick);
}
if (dateFilterButton) {
  dateFilterButton.addEventListener("click", onDateFilterButtonClick);
}
if (dateFilterClearButton) {
  dateFilterClearButton.addEventListener("click", clearActiveDateFilter);
}
if (fromFilterButton) {
  fromFilterButton.addEventListener("click", () => onTextFilterButtonClick("from"));
}
if (toFilterButton) {
  toFilterButton.addEventListener("click", () => onTextFilterButtonClick("to"));
}
if (subjectFilterButton) {
  subjectFilterButton.addEventListener("click", () => onTextFilterButtonClick("subject"));
}
if (dateFromRange) {
  dateFromRange.addEventListener("input", onDateRangeInput);
}
if (dateToRange) {
  dateToRange.addEventListener("input", onDateRangeInput);
}
if (fromFilterInput) {
  fromFilterInput.addEventListener("input", () => onTextFilterInput("from"));
}
if (toFilterInput) {
  toFilterInput.addEventListener("input", () => onTextFilterInput("to"));
}
if (subjectFilterInput) {
  subjectFilterInput.addEventListener("input", () => onTextFilterInput("subject"));
}
if (fromFilterClearButton) {
  fromFilterClearButton.addEventListener("click", () => clearTextFilter("from"));
}
if (toFilterClearButton) {
  toFilterClearButton.addEventListener("click", () => clearTextFilter("to"));
}
if (subjectFilterClearButton) {
  subjectFilterClearButton.addEventListener("click", () => clearTextFilter("subject"));
}
if (attachmentToggleButton) {
  attachmentToggleButton.addEventListener("click", toggleAttachmentsOnlyFilter);
}
document.addEventListener("pointerdown", handleDocumentPointerDown);
document.addEventListener("keydown", handleGlobalKeyDown);
if (mailListPanel) {
  mailListPanel.addEventListener("scroll", handleMailListScroll, { passive: true });
}
if (externalLinkModal) {
  externalLinkModal.addEventListener("click", handleExternalLinkModalClick);
}
if (externalLinkCancel) {
  externalLinkCancel.addEventListener("click", () => setExternalLinkModalOpen(false));
}
if (externalLinkCopy) {
  externalLinkCopy.addEventListener("click", copyPendingExternalLink);
}
if (externalLinkOpen) {
  externalLinkOpen.addEventListener("click", confirmOpenExternalLink);
}
if (emlSourceModal) {
  emlSourceModal.addEventListener("click", handleEmlSourceModalClick);
}
if (emlSourceClose) {
  emlSourceClose.addEventListener("click", () => setEmlSourceModalOpen(false));
}

window.addEventListener("resize", () => {
  const frame = document.getElementById("messageFrame");
  if (frame) {
    fitMessageFrameToContent(frame);
  }
});

const removeProgressListener = window.mboxApi.onIndexProgress(handleIndexProgress);
const removeOpenMailboxRequestListener = window.mboxApi.onOpenMailboxRequest((payload) => {
  const filePath = typeof payload?.filePath === "string" ? payload.filePath : "";
  if (!filePath) {
    return;
  }
  void openMailboxPath(filePath);
});
window.addEventListener("beforeunload", () => {
  if (typeof removeProgressListener === "function") {
    removeProgressListener();
  }
  if (typeof removeOpenMailboxRequestListener === "function") {
    removeOpenMailboxRequestListener();
  }
});

initSplitter();
updateRemoteContentButtonState();
updateSearchUiState();
setSearchVisible(false);
setTextFiltersVisible(false);
setRemoteContentVisible(false);
void consumePendingMailboxOpen();

async function openMbox() {
  return openMailboxRequest(() => window.mboxApi.openMbox());
}

async function openMailboxPath(filePath) {
  const normalizedPath = String(filePath || "").trim();
  if (!normalizedPath) {
    return;
  }

  return openMailboxRequest(() => window.mboxApi.openMailboxPath({ filePath: normalizedPath }));
}

async function consumePendingMailboxOpen() {
  try {
    const payload = await window.mboxApi.consumePendingOpenFile();
    const filePath = typeof payload?.filePath === "string" ? payload.filePath : "";
    if (filePath) {
      await openMailboxPath(filePath);
    }
  } catch (error) {
    console.error("Failed to consume pending mailbox open request.", error);
  }
}

async function openMailboxRequest(loader) {
  requestToken += 1;
  messageRequestToken += 1;
  if (searchDebounceTimer) {
    clearTimeout(searchDebounceTimer);
    searchDebounceTimer = null;
  }
  openingInProgress = true;
  setOpenButtonBusy(true);
  setOpenProgress({ visible: true, indeterminate: true, value: 8 });
  setStatusMessage("Opening file...");

  try {
    const result = await loader();
    if (!result || result.canceled) {
      setStatusMessage(result?.error || "Open cancelled.");
      setOpenProgress({ visible: false });
      return;
    }

    dbPath = result.dbPath || "";
    mboxPath = result.filePath || "";
    totalMessages = Number.isInteger(result.total) ? result.total : 0;
    currentStandaloneMessage = result?.standaloneMessage || null;
    setSearchVisible(Boolean(dbPath));
    setTextFiltersVisible(Boolean(dbPath));
    setRemoteContentVisible(Boolean(dbPath || currentStandaloneMessage));
    currentQuery = "";
    searchInput.value = "";
    resetTextFilters();
    updateSearchUiState();
    currentOffset = 0;
    selectedMessageId = null;
    configureDateFilter(result?.dateRange || null);

    applyPageResult(result);
    setOpenProgress({ visible: true, indeterminate: false, value: 100 });
    await loadSelectedMessage();
    setStatusMessage(`Loaded ${totalMessages} email${totalMessages === 1 ? "" : "s"} from ${mboxPath}`);
  } catch (error) {
    setStatusMessage("Failed to open file.");
    setOpenProgress({ visible: false });
    console.error(error);
  } finally {
    openingInProgress = false;
    setOpenButtonBusy(false);
    refreshStatusMeta();
    if (dbPath || currentStandaloneMessage) {
      setTimeout(() => {
        if (!openingInProgress) {
          setOpenProgress({ visible: false });
        }
      }, 220);
    } else {
      setOpenProgress({ visible: false });
    }
  }
}

function onSearchInput() {
  currentQuery = searchInput.value.trim();
  updateSearchUiState();
  currentOffset = 0;
  scheduleLoadPage();
}

function onSearchClearClick() {
  if (!searchInput || !searchInput.value) {
    return;
  }

  searchInput.value = "";
  searchInput.focus();
  currentQuery = "";
  currentOffset = 0;
  updateSearchUiState();
  scheduleLoadPage();
}

function updateSearchUiState() {
  const hasValue = Boolean(searchInput && searchInput.value.length > 0);
  if (searchWrap) {
    searchWrap.classList.toggle("has-value", hasValue);
  }
  if (searchClearButton) {
    searchClearButton.hidden = !hasValue;
  }
}

function setSearchVisible(visible) {
  if (!searchWrap) {
    return;
  }
  searchWrap.hidden = !visible;
}

function setRemoteContentVisible(visible) {
  if (!remoteContentButton) {
    return;
  }
  remoteContentButton.hidden = !visible;
}

function toggleRemoteContent() {
  remoteContentEnabled = !remoteContentEnabled;
  updateRemoteContentButtonState();
  setStatusMessage(
    remoteContentEnabled
      ? "Remote content is enabled for HTML messages."
      : "Remote content is blocked for HTML messages."
  );
  if (selectedMessageId) {
    void loadSelectedMessage();
  }
}

function updateRemoteContentButtonState() {
  if (!remoteContentButton) {
    return;
  }

  remoteContentButton.classList.toggle("active", remoteContentEnabled);
  remoteContentButton.setAttribute("aria-pressed", remoteContentEnabled ? "true" : "false");
  remoteContentButton.setAttribute(
    "aria-label",
    remoteContentEnabled ? "Remote content enabled" : "Remote content blocked"
  );
  remoteContentButton.title = remoteContentEnabled ? "Remote content enabled" : "Remote content blocked";
}

function onDateFilterButtonClick() {
  if (!dateFilterButton || dateFilterButton.hidden) {
    return;
  }
  closeTextFilterPopovers();
  setDateFilterPopoverOpen(!dateFilterPopoverOpen);
}

function onTextFilterButtonClick(filterKey) {
  const nextOpen = !isTextFilterPopoverOpen(filterKey);
  setDateFilterPopoverOpen(false);
  closeTextFilterPopovers();
  setTextFilterPopoverOpen(filterKey, nextOpen);
}

function handleDocumentPointerDown(event) {
  if (!isAnyToolbarPopoverOpen()) {
    return;
  }
  const target = event.target;
  const containers = [dateFilterContainer, fromFilterContainer, toFilterContainer, subjectFilterContainer].filter(Boolean);
  if (containers.some((container) => container.contains(target))) {
    return;
  }
  closeAllToolbarPopovers();
}

function handleGlobalKeyDown(event) {
  if (event.key === "Escape") {
    if (emlSourceModalOpen) {
      setEmlSourceModalOpen(false);
      return;
    }

    if (externalLinkModalOpen) {
      setExternalLinkModalOpen(false);
      return;
    }

    if (isAnyToolbarPopoverOpen()) {
      closeAllToolbarPopovers();
      return;
    }
    return;
  }

  if (externalLinkModalOpen || emlSourceModalOpen || event.defaultPrevented) {
    return;
  }
  if (event.metaKey || event.ctrlKey || event.altKey) {
    return;
  }
  if (isTypingTarget(event.target)) {
    return;
  }

  if (event.key === "PageDown") {
    event.preventDefault();
    void jumpMessageSelectionByPage(1);
    return;
  }

  if (event.key === "PageUp") {
    event.preventDefault();
    void jumpMessageSelectionByPage(-1);
    return;
  }

  if (event.key === "Home") {
    event.preventDefault();
    void jumpToAbsoluteMessageIndex(0);
    return;
  }

  if (event.key === "End") {
    event.preventDefault();
    void jumpToAbsoluteMessageIndex(totalResults - 1);
    return;
  }

  const key = String(event.key || "");
  const direction = key === "ArrowDown" || key === "j" || key === "J" ? 1 : key === "ArrowUp" || key === "k" || key === "K" ? -1 : 0;
  if (!direction) {
    return;
  }

  event.preventDefault();
  moveMessageSelection(direction);
}

function isTypingTarget(target) {
  if (!target) {
    return false;
  }

  if (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement
  ) {
    return true;
  }

  if (target instanceof HTMLElement) {
    if (target.isContentEditable) {
      return true;
    }
    if (typeof target.closest === "function" && target.closest("[contenteditable='true']")) {
      return true;
    }
  }

  return false;
}

function moveMessageSelection(direction) {
  if (!dbPath || currentPageMessages.length === 0) {
    return;
  }

  const currentIndex = currentPageMessages.findIndex((message) => message.id === selectedMessageId);
  if (currentIndex === -1) {
    const nextIndex = direction > 0 ? 0 : currentPageMessages.length - 1;
    const nextMessage = currentPageMessages[nextIndex];
    if (!nextMessage) {
      return;
    }
    selectedMessageId = nextMessage.id;
    renderList();
    scrollSelectedMessageIntoView();
    void loadSelectedMessage();
    return;
  }

  const nextIndex = currentIndex + direction;
  if (nextIndex >= 0 && nextIndex < currentPageMessages.length) {
    const nextMessage = currentPageMessages[nextIndex];
    if (!nextMessage || nextMessage.id === selectedMessageId) {
      return;
    }

    selectedMessageId = nextMessage.id;
    renderList();
    scrollSelectedMessageIntoView();
    void loadSelectedMessage();
    return;
  }

  if (direction > 0) {
    if (getLoadedMessageEndOffset() >= totalResults) {
      return;
    }
    void loadMoreMessages({ edgeSelection: "first-new" });
    return;
  }

  if (direction < 0) {
    return;
  }
}

async function jumpMessageSelectionByPage(direction) {
  if (!dbPath || totalResults <= 0) {
    return;
  }

  const currentAbsoluteIndex = getCurrentAbsoluteMessageIndex();
  const baseIndex = currentAbsoluteIndex >= 0 ? currentAbsoluteIndex : 0;
  const jumpSize = getVisibleMessageJumpCount();
  await jumpToAbsoluteMessageIndex(baseIndex + direction * jumpSize);
}

async function jumpToAbsoluteMessageIndex(targetIndex) {
  if (!dbPath || totalResults <= 0) {
    return;
  }

  const clampedIndex = Math.max(0, Math.min(totalResults - 1, Number(targetIndex) || 0));
  const localIndex = clampedIndex - currentOffset;
  if (localIndex >= 0 && localIndex < currentPageMessages.length) {
    const nextMessage = currentPageMessages[localIndex];
    if (!nextMessage || nextMessage.id === selectedMessageId) {
      return;
    }

    selectedMessageId = nextMessage.id;
    renderList();
    scrollSelectedMessageIntoView();
    await loadSelectedMessage();
    return;
  }

  const pageOffset = Math.floor(clampedIndex / PAGE_SIZE) * PAGE_SIZE;
  currentOffset = pageOffset;
  await loadPage({
    offset: pageOffset,
    selectAbsoluteIndex: clampedIndex,
    scrollToSelected: true
  });
}

function getCurrentAbsoluteMessageIndex() {
  if (selectedMessageId === null || selectedMessageId === undefined) {
    return -1;
  }

  const localIndex = currentPageMessages.findIndex((message) => message.id === selectedMessageId);
  if (localIndex === -1) {
    return -1;
  }

  return currentOffset + localIndex;
}

function getLoadedMessageEndOffset() {
  return currentOffset + currentPageMessages.length;
}

function getVisibleMessageJumpCount() {
  if (!mailListPanel) {
    return 10;
  }

  const firstItem = mailList.querySelector(".mail-item");
  if (!firstItem) {
    return 10;
  }

  const itemHeight = Math.max(1, firstItem.getBoundingClientRect().height + 6);
  const panelHeight = Math.max(1, mailListPanel.clientHeight);
  return Math.max(1, Math.floor(panelHeight / itemHeight));
}

function scrollSelectedMessageIntoView() {
  const selectedItem = mailList.querySelector(".mail-item.active");
  if (!selectedItem || typeof selectedItem.scrollIntoView !== "function") {
    return;
  }
  selectedItem.scrollIntoView({ block: "nearest" });
}

function setDateFilterPopoverOpen(nextOpen) {
  dateFilterPopoverOpen = Boolean(nextOpen);
  if (dateFilter) {
    dateFilter.hidden = !dateFilterPopoverOpen;
  }
  if (dateFilterButton) {
    dateFilterButton.setAttribute("aria-expanded", dateFilterPopoverOpen ? "true" : "false");
  }
}

function setTextFilterPopoverOpen(filterKey, nextOpen) {
  const open = Boolean(nextOpen);
  const config = getTextFilterConfig(filterKey);
  if (!config) {
    return;
  }

  if (filterKey === "from") {
    fromFilterPopoverOpen = open;
  } else if (filterKey === "to") {
    toFilterPopoverOpen = open;
  } else if (filterKey === "subject") {
    subjectFilterPopoverOpen = open;
  }

  if (config.popover) {
    config.popover.hidden = !open;
  }
  if (config.button) {
    config.button.setAttribute("aria-expanded", open ? "true" : "false");
  }
  if (open && config.input) {
    setTimeout(() => config.input.focus(), 0);
  }
}

function closeTextFilterPopovers() {
  setTextFilterPopoverOpen("from", false);
  setTextFilterPopoverOpen("to", false);
  setTextFilterPopoverOpen("subject", false);
}

function closeAllToolbarPopovers() {
  setDateFilterPopoverOpen(false);
  closeTextFilterPopovers();
}

function isAnyToolbarPopoverOpen() {
  return dateFilterPopoverOpen || fromFilterPopoverOpen || toFilterPopoverOpen || subjectFilterPopoverOpen;
}

function isTextFilterPopoverOpen(filterKey) {
  if (filterKey === "from") {
    return fromFilterPopoverOpen;
  }
  if (filterKey === "to") {
    return toFilterPopoverOpen;
  }
  if (filterKey === "subject") {
    return subjectFilterPopoverOpen;
  }
  return false;
}

function handleExternalLinkModalClick(event) {
  const target = event.target;
  if (!(target instanceof Element)) {
    return;
  }

  if (target.getAttribute("data-close-external-link") === "true") {
    setExternalLinkModalOpen(false);
  }
}

function handleEmlSourceModalClick(event) {
  const target = event.target;
  if (!(target instanceof Element)) {
    return;
  }

  if (target.getAttribute("data-close-eml-source") === "true") {
    setEmlSourceModalOpen(false);
  }
}

function setExternalLinkModalOpen(nextOpen) {
  externalLinkModalOpen = Boolean(nextOpen);
  if (externalLinkModal) {
    externalLinkModal.hidden = !externalLinkModalOpen;
  }
  if (!externalLinkModalOpen) {
    pendingExternalUrl = "";
  }
  if (externalLinkModalOpen && externalLinkOpen) {
    setTimeout(() => externalLinkOpen.focus(), 0);
  }
}

function setEmlSourceModalOpen(nextOpen) {
  emlSourceModalOpen = Boolean(nextOpen);
  if (emlSourceModal) {
    emlSourceModal.hidden = !emlSourceModalOpen;
  }
  if (!emlSourceModalOpen && emlSourceContent) {
    emlSourceContent.textContent = "";
  }
  if (emlSourceModalOpen && emlSourceClose) {
    setTimeout(() => emlSourceClose.focus(), 0);
  }
}

function requestExternalLinkOpen(url) {
  pendingExternalUrl = String(url || "").trim();
  if (!pendingExternalUrl) {
    return;
  }

  renderExternalLinkPreview(pendingExternalUrl);
  setExternalLinkModalOpen(true);
}

async function confirmOpenExternalLink() {
  if (!pendingExternalUrl) {
    setExternalLinkModalOpen(false);
    return;
  }

  const urlToOpen = pendingExternalUrl;
  pendingExternalUrl = "";
  setExternalLinkModalOpen(false);
  try {
    const result = await window.mboxApi.openExternal({ url: urlToOpen });
    if (!result || !result.opened) {
      setStatusMessage("Could not open external link.");
    }
  } catch (error) {
    setStatusMessage("Could not open external link.");
    console.error(error);
  }
}

async function copyPendingExternalLink() {
  if (!pendingExternalUrl) {
    return;
  }

  try {
    const result = await window.mboxApi.copyToClipboard({ text: pendingExternalUrl });
    if (!result || !result.copied) {
      setStatusMessage("Could not copy link to clipboard.");
      return;
    }
    setStatusMessage("Link copied to clipboard.");
  } catch (error) {
    setStatusMessage("Could not copy link to clipboard.");
    console.error(error);
  }
}

function openEmlSourcePreview(emlSource) {
  if (!emlSourceContent) {
    return;
  }

  emlSourceContent.textContent = String(emlSource || "").replace(/\r\n/g, "\n");
  setEmlSourceModalOpen(true);
}

function renderExternalLinkPreview(urlValue) {
  if (!externalLinkUrl) {
    return;
  }

  const parts = splitUrlForDomainHighlight(urlValue);
  externalLinkUrl.innerHTML = `${escapeHtml(parts.before)}<strong>${escapeHtml(parts.domain)}</strong>${escapeHtml(
    parts.after
  )}`;
}

async function openAttachmentPreview(attachment) {
  if (!isPreviewableAttachment(attachment)) {
    setStatusMessage("Preview is only available for image and PDF attachments.");
    return;
  }
  if (!attachment?.base64) {
    setStatusMessage("Preview data is not available for this attachment.");
    return;
  }

  setStatusMessage(`Opening preview for ${attachment.fileName || "attachment"}...`);
  try {
    const result = await window.mboxApi.openAttachmentPreview({
      fileName: attachment.fileName || "attachment",
      contentType: attachment.contentType || "application/octet-stream",
      base64: attachment.base64 || ""
    });
    if (!result || !result.opened) {
      setStatusMessage("Could not open attachment preview.");
      return;
    }
    setStatusMessage(`Preview opened for ${attachment.fileName || "attachment"}.`);
  } catch (error) {
    setStatusMessage("Could not open attachment preview.");
    console.error(error);
  }
}

function splitUrlForDomainHighlight(urlValue) {
  const full = String(urlValue || "");
  try {
    const parsed = new URL(full);
    const protocol = parsed.protocol.toLowerCase();

    if (protocol === "http:" || protocol === "https:") {
      const prefix = `${parsed.protocol}//`;
      let domainStart = full.indexOf(prefix);
      if (domainStart !== -1) {
        domainStart += prefix.length;
      } else {
        domainStart = 0;
      }

      const afterScheme = full.slice(domainStart);
      const authEnd = afterScheme.lastIndexOf("@");
      if (authEnd !== -1) {
        domainStart += authEnd + 1;
      }

      const hostName = parsed.hostname || "";
      const highlightedDomain = getRegistrableDomain(hostName);
      const hostIndex = hostName ? full.indexOf(hostName, domainStart) : -1;
      if (hostIndex !== -1 && highlightedDomain) {
        const domainIndex = hostIndex + Math.max(0, hostName.length - highlightedDomain.length);
        return {
          before: full.slice(0, domainIndex),
          domain: highlightedDomain,
          after: full.slice(domainIndex + highlightedDomain.length)
        };
      }
    }

    if (protocol === "mailto:") {
      const queryIndex = full.indexOf("?");
      const body = queryIndex === -1 ? full : full.slice(0, queryIndex);
      const atIndex = body.lastIndexOf("@");
      if (atIndex !== -1 && atIndex + 1 < body.length) {
        const domainStart = atIndex + 1;
        const domainEnd = body.length;
        const mailDomain = full.slice(domainStart, domainEnd);
        const highlightedDomain = getRegistrableDomain(mailDomain);
        const highlightStart = domainStart + Math.max(0, mailDomain.length - highlightedDomain.length);
        return {
          before: full.slice(0, highlightStart),
          domain: highlightedDomain,
          after: full.slice(highlightStart + highlightedDomain.length)
        };
      }
    }
  } catch {
    // Fallback to full URL without parsing.
  }

  return { before: "", domain: full, after: "" };
}

function getRegistrableDomain(hostValue) {
  const host = String(hostValue || "").trim().toLowerCase().replace(/\.+$/, "");
  if (!host || !host.includes(".") || host.includes(":")) {
    return host;
  }

  const parts = host.split(".").filter(Boolean);
  if (parts.length <= 2) {
    return host;
  }

  const commonSecondLevelSuffixes = new Set([
    "ac.uk",
    "co.jp",
    "co.kr",
    "co.nz",
    "co.uk",
    "com.au",
    "com.br",
    "com.cn",
    "com.hk",
    "com.mx",
    "com.sg",
    "com.tr",
    "edu.au",
    "gov.uk",
    "net.au",
    "org.au",
    "org.uk"
  ]);
  const trailingPair = `${parts[parts.length - 2]}.${parts[parts.length - 1]}`;
  if (commonSecondLevelSuffixes.has(trailingPair) && parts.length >= 3) {
    return parts.slice(-3).join(".");
  }

  if (parts[parts.length - 1].length === 2 && parts[parts.length - 2].length <= 3 && parts.length >= 3) {
    return parts.slice(-3).join(".");
  }

  return parts.slice(-2).join(".");
}

function onDateRangeInput(event) {
  if (dateRangeMinDayTs === null || dateRangeMaxDayTs === null) {
    return;
  }

  const maxOffset = getMaxDateDayOffset();
  const nextFrom = clampNumberInRange(dateFromRange?.value, 0, maxOffset, dateFromDayOffset);
  const nextTo = clampNumberInRange(dateToRange?.value, 0, maxOffset, dateToDayOffset);

  if (event?.target === dateFromRange && nextFrom > nextTo) {
    dateFromDayOffset = nextFrom;
    dateToDayOffset = nextFrom;
  } else if (event?.target === dateToRange && nextTo < nextFrom) {
    dateFromDayOffset = nextTo;
    dateToDayOffset = nextTo;
  } else {
    dateFromDayOffset = Math.min(nextFrom, nextTo);
    dateToDayOffset = Math.max(nextFrom, nextTo);
  }

  syncDateRangeInputs();
  updateDateFilterLabels();
  updateDateFilterButtonState();
  currentOffset = 0;
  scheduleLoadPage();
}

function onTextFilterInput(filterKey) {
  updateTextFilterButtonState();
  updateTextFilterClearButtons();
  currentOffset = 0;
  scheduleLoadPage();
}

function scheduleLoadPage() {
  if (searchDebounceTimer) {
    clearTimeout(searchDebounceTimer);
  }

  searchDebounceTimer = setTimeout(() => {
    searchDebounceTimer = null;
    void loadPage();
  }, SEARCH_DEBOUNCE_MS);
}

async function loadPage(options = {}) {
  if (!dbPath) {
    return;
  }

  const dateFilterPayload = getActiveDateFilterPayload();
  const textFiltersPayload = getActiveTextFiltersPayload();
  const hasTextQuery = Boolean(currentQuery);
  const hasDateFilter = dateFilterPayload !== null;
  const hasFieldFilter = textFiltersPayload !== null;
  const token = ++requestToken;
  const queryLabel =
    hasTextQuery || hasDateFilter || hasFieldFilter
      ? "Applying filters..."
      : "Loading messages...";
  setStatusMessage(queryLabel);

  try {
    const requestedOffset = Number.isInteger(options.offset) ? options.offset : currentOffset;
    const pageOffset = options.append ? getLoadedMessageEndOffset() : Math.max(0, requestedOffset);
    const page = await window.mboxApi.searchMessages({
      dbPath,
      query: currentQuery,
      limit: PAGE_SIZE,
      offset: pageOffset,
      dateFrom: dateFilterPayload?.dateFrom ?? null,
      dateTo: dateFilterPayload?.dateTo ?? null,
      senderQuery: textFiltersPayload?.senderQuery ?? null,
      recipientQuery: textFiltersPayload?.recipientQuery ?? null,
      subjectQuery: textFiltersPayload?.subjectQuery ?? null,
      attachmentsOnly: textFiltersPayload?.attachmentsOnly ?? false
    });

    if (token !== requestToken) {
      return;
    }

    applyPageResult(page, options);
    await loadSelectedMessage(token);
    if (options.scrollToSelected) {
      scrollSelectedMessageIntoView();
    }

    if (hasTextQuery || hasDateFilter || hasFieldFilter) {
      setStatusMessage(`Showing ${totalResults} matching emails in ${mboxPath}`);
    } else {
      setStatusMessage(`Showing emails from ${mboxPath}`);
    }
  } catch (error) {
    if (token !== requestToken) {
      return;
    }
    setStatusMessage("Search failed.");
    console.error(error);
  }
}

function applyPageResult(result, options = {}) {
  totalResults = Number.isInteger(result?.total) ? result.total : 0;
  const append = options.append === true;
  const incomingMessages = Array.isArray(result?.messages) ? result.messages : [];
  const previousScrollTop = mailListPanel ? mailListPanel.scrollTop : 0;

  if (append) {
    const seenIds = new Set(currentPageMessages.map((message) => message.id));
    const appendedMessages = incomingMessages.filter((message) => !seenIds.has(message.id));
    currentPageMessages = currentPageMessages.concat(appendedMessages);
    resultIndexById = new Map(currentPageMessages.map((message) => [message.id, message.resultIndex || null]));

    if (options?.edgeSelection === "first-new") {
      selectedMessageId = appendedMessages[0]?.id ?? selectedMessageId;
    }
  } else {
    currentOffset = Number.isInteger(result?.offset) ? result.offset : currentOffset;
    currentPageMessages = incomingMessages;
    resultIndexById = new Map(currentPageMessages.map((message) => [message.id, message.resultIndex || null]));
  }

  const selectAbsoluteIndex = Number.isInteger(options.selectAbsoluteIndex) ? options.selectAbsoluteIndex : null;
  if (selectAbsoluteIndex !== null) {
    const selectLocalIndex = selectAbsoluteIndex - currentOffset;
    selectedMessageId = currentPageMessages[selectLocalIndex]?.id ?? selectedMessageId;
  }

  const edgeSelection = options?.edgeSelection;
  if (!append && edgeSelection === "first") {
    selectedMessageId = currentPageMessages[0]?.id ?? null;
  } else if (!append && edgeSelection === "last") {
    selectedMessageId = currentPageMessages[currentPageMessages.length - 1]?.id ?? null;
  }

  if (!currentPageMessages.some((message) => message.id === selectedMessageId)) {
    selectedMessageId = currentPageMessages[0]?.id ?? null;
  }

  renderList();
  if (append && mailListPanel) {
    mailListPanel.scrollTop = previousScrollTop;
  } else if (!append && mailListPanel) {
    mailListPanel.scrollTop = 0;
  }
  refreshStatusMeta();
}

function handleMailListScroll() {
  if (!shouldLoadMoreMessages()) {
    return;
  }
  void loadMoreMessages();
}

function shouldLoadMoreMessages() {
  if (!mailListPanel || !dbPath || loadingMoreMessages) {
    return false;
  }
  if (currentPageMessages.length === 0 || getLoadedMessageEndOffset() >= totalResults) {
    return false;
  }

  const thresholdPx = 160;
  return mailListPanel.scrollTop + mailListPanel.clientHeight >= mailListPanel.scrollHeight - thresholdPx;
}

async function loadMoreMessages(options = {}) {
  if (!shouldLoadMoreMessages() && options.edgeSelection !== "first-new") {
    return;
  }
  if (loadingMoreMessages || !dbPath || getLoadedMessageEndOffset() >= totalResults) {
    return;
  }

  loadingMoreMessages = true;
  try {
    await loadPage({ append: true, edgeSelection: options.edgeSelection || null });
    if (options.edgeSelection === "first-new") {
      scrollSelectedMessageIntoView();
    }
  } finally {
    loadingMoreMessages = false;
    if (shouldLoadMoreMessages()) {
      setTimeout(() => {
        if (shouldLoadMoreMessages()) {
          void loadMoreMessages();
        }
      }, 0);
    }
  }
}

async function loadSelectedMessage(expectedPageToken = requestToken) {
  const currentMessageToken = ++messageRequestToken;

  if (!selectedMessageId) {
    renderMessage(null);
    return;
  }

  if (!dbPath) {
    if (currentStandaloneMessage && currentStandaloneMessage.id === selectedMessageId) {
      currentStandaloneMessage.resultIndex = resultIndexById.get(selectedMessageId) || 1;
      renderMessage(currentStandaloneMessage);
      return;
    }
    renderMessage(null);
    return;
  }

  messageView.innerHTML = `
    <h2>Loading message...</h2>
    <p>Please wait while the message is loaded from disk.</p>
  `;

  try {
    const message = await window.mboxApi.getMessage({
      dbPath,
      id: selectedMessageId
    });

    if (expectedPageToken !== requestToken || currentMessageToken !== messageRequestToken) {
      return;
    }

    if (!message) {
      renderMessage(null);
      return;
    }

    message.resultIndex = resultIndexById.get(message.id) || null;
    renderMessage(message);
  } catch (error) {
    if (expectedPageToken !== requestToken || currentMessageToken !== messageRequestToken) {
      return;
    }
    renderMessage(null);
    setStatusMessage("Failed to load selected message.");
    console.error(error);
  }
}

function renderList() {
  mailList.innerHTML = "";

  if (!dbPath && currentPageMessages.length === 0) {
    const empty = document.createElement("li");
    empty.className = "empty";
    empty.textContent = "Open an mbox, pst, or eml file to start browsing messages.";
    mailList.appendChild(empty);
    return;
  }

  if (currentPageMessages.length === 0) {
    const empty = document.createElement("li");
    empty.className = "empty";
    empty.textContent = currentQuery || isDateFilterActive() || isAnyTextFilterActive()
      ? "No emails match the current filters."
      : "No emails found in this page.";
    mailList.appendChild(empty);
    return;
  }

  for (const msg of currentPageMessages) {
    const sender = getSenderDisplay(msg.from || "Unknown sender");
    const dateLabel = formatListDate(msg.date || "");
    const snippet = compactText(msg.snippet || "");
    const attachmentBadge = msg.hasAttachments
      ? `
        <span class="mail-attachment-indicator" aria-label="Has attachment" title="Has attachment">
          <svg viewBox="0 0 24 24" aria-hidden="true">
            <path
              d="M8.5 12.5v4.25a3.25 3.25 0 1 0 6.5 0v-7.5a2 2 0 1 0-4 0v6.75a.75.75 0 0 0 1.5 0v-5.75a1 1 0 1 1 2 0v5.75a2.75 2.75 0 0 1-5.5 0V9.25a4.5 4.5 0 0 1 9 0v7.5a5 5 0 1 1-10 0V12.5a.75.75 0 0 1 1.5 0Z"
              fill="currentColor"
            ></path>
          </svg>
        </span>
      `
      : "";
    const item = document.createElement("li");
    item.className = `mail-item${msg.id === selectedMessageId ? " active" : ""}`;
    item.innerHTML = `
      <div class="mail-row-top">
        <div class="mail-from" title="${escapeHtml(msg.from || "Unknown sender")}">${escapeHtml(sender)}</div>
        <div class="mail-date-wrap">
          ${attachmentBadge}
          <span class="mail-date">${escapeHtml(dateLabel)}</span>
        </div>
      </div>
      <p class="mail-subject-line">${escapeHtml(msg.subject || "(No Subject)")}</p>
      <div class="mail-snippet">${escapeHtml(snippet)}</div>
    `;

    item.addEventListener("click", () => {
      selectedMessageId = msg.id;
      renderList();
      void loadSelectedMessage();
    });

    mailList.appendChild(item);
  }
}

function renderMessage(msg) {
  if (!msg) {
    messageView.innerHTML = `
      <h2>No message selected</h2>
      <p>Open an mbox file or adjust your search query.</p>
    `;
    refreshStatusMeta();
    return;
  }

  const attachments = Array.isArray(msg.attachments) ? msg.attachments : [];
  const bodyHtml = resolveCidUrls(msg.bodyHtml || "", attachments);
  const sanitizedBody = sanitizeEmailHtml(bodyHtml, { allowRemoteContent: remoteContentEnabled });
  const remoteContentWarningMarkup = sanitizedBody.blockedRemoteContent
    ? `
      <div class="message-privacy-banner" data-role="remote-content-warning">
        To protect your privacy, remote content was blocked in this message.
      </div>
    `
    : "";
  const attachmentsMarkup = attachments.length
    ? `
      <section class="attachments">
        <h3>Attachments (${attachments.length})</h3>
        <ul class="attachment-list">
          ${attachments
            .map(
              (att) => `
            <li class="attachment-item">
              <div class="attachment-meta">
                <div class="attachment-name">${escapeHtml(att.fileName || "attachment")}</div>
                <div class="attachment-sub">${escapeHtml(att.contentType || "application/octet-stream")}${
                  att.size != null ? ` | ${formatBytes(att.size)}` : ""
                }${att.isInline ? " | inline" : ""}</div>
              </div>
              <div class="attachment-actions">
                ${
                  isPreviewableAttachment(att)
                    ? `<button class="attachment-preview" data-attachment-id="${escapeHtml(att.id)}">Preview</button>`
                    : ""
                }
                <button class="attachment-download" data-attachment-id="${escapeHtml(att.id)}">Download</button>
              </div>
            </li>
          `
            )
            .join("")}
        </ul>
      </section>
    `
    : "";

  messageView.innerHTML = `
    <section class="message-header">
      <div class="message-header-top">
        <h2>${escapeHtml(msg.subject || "(No Subject)")}</h2>
        <div class="message-header-actions">
          <button id="previewEmlButton" class="message-preview-eml" type="button" aria-label="Preview original EML source" title="Preview original EML source">
            <svg viewBox="0 0 24 24" aria-hidden="true">
              <path
                d="M4 5.5A2.5 2.5 0 0 1 6.5 3h7.88a2.5 2.5 0 0 1 1.77.73l3.12 3.12A2.5 2.5 0 0 1 20 8.62V18.5A2.5 2.5 0 0 1 17.5 21h-11A2.5 2.5 0 0 1 4 18.5v-13Zm2.5-.5a.5.5 0 0 0-.5.5v13a.5.5 0 0 0 .5.5h11a.5.5 0 0 0 .5-.5V9h-3a2 2 0 0 1-2-2V5H6.5Zm8.5.41V7h1.59L15 5.41ZM8 11a1 1 0 0 1 1-1h6a1 1 0 1 1 0 2H9a1 1 0 0 1-1-1Zm0 4a1 1 0 0 1 1-1h6a1 1 0 1 1 0 2H9a1 1 0 0 1-1-1Z"
                fill="currentColor"
              ></path>
            </svg>
          </button>
          <button id="downloadEmlButton" class="message-download-eml" type="button" aria-label="Download message as EML" title="Download message as EML">
            <svg viewBox="0 0 24 24" aria-hidden="true">
              <path
                d="M12 3a1 1 0 0 1 1 1v8.59l2.3-2.29a1 1 0 1 1 1.4 1.41l-4 3.99a1 1 0 0 1-1.4 0l-4-3.99a1 1 0 1 1 1.4-1.41L11 12.59V4a1 1 0 0 1 1-1Zm-7 13a1 1 0 0 1 1 1v1.5a.5.5 0 0 0 .5.5h11a.5.5 0 0 0 .5-.5V17a1 1 0 1 1 2 0v1.5a2.5 2.5 0 0 1-2.5 2.5h-11A2.5 2.5 0 0 1 4 18.5V17a1 1 0 0 1 1-1Z"
                fill="currentColor"
              ></path>
            </svg>
          </button>
        </div>
      </div>
      <p><strong>From:</strong> ${escapeHtml(msg.from || "")}</p>
      <p><strong>To:</strong> ${escapeHtml(msg.to || "")}</p>
      <p><strong>Date:</strong> ${escapeHtml(msg.date || "")}</p>
    </section>
    ${attachmentsMarkup}
    <section class="message-body">
      ${remoteContentWarningMarkup}
      <iframe
        id="messageFrame"
        class="message-frame"
        sandbox="allow-same-origin allow-popups allow-popups-to-escape-sandbox"
        scrolling="no"
      ></iframe>
    </section>
  `;

  for (const button of messageView.querySelectorAll(".attachment-download")) {
    button.addEventListener("click", async () => {
      const attachmentId = button.dataset.attachmentId;
      const attachment = attachments.find((item) => item.id === attachmentId);
      if (!attachment) {
        return;
      }

      setStatusMessage(`Saving ${attachment.fileName || "attachment"}...`);
      try {
        const result = await window.mboxApi.saveAttachment({
          fileName: attachment.fileName || "attachment.bin",
          base64: attachment.base64 || ""
        });

        if (!result || result.canceled) {
          setStatusMessage("Save cancelled.");
          return;
        }

        setStatusMessage(`Saved attachment to ${result.filePath}`);
      } catch (error) {
        setStatusMessage("Failed to save attachment.");
        console.error(error);
      }
    });
  }

  for (const button of messageView.querySelectorAll(".attachment-preview")) {
    button.addEventListener("click", () => {
      const attachmentId = button.dataset.attachmentId;
      const attachment = attachments.find((item) => item.id === attachmentId);
      if (!attachment) {
        return;
      }
      void openAttachmentPreview(attachment);
    });
  }

  const emlButton = document.getElementById("downloadEmlButton");
  if (emlButton) {
    emlButton.addEventListener("click", async () => {
      setStatusMessage("Saving .eml file...");
      try {
        const result = await window.mboxApi.saveMessageEml({
          fileName: buildEmlFileName(msg),
          emlSource: msg.emlSource || ""
        });

        if (!result || result.canceled) {
          setStatusMessage("Save cancelled.");
          return;
        }

        setStatusMessage(`Saved .eml to ${result.filePath}`);
      } catch (error) {
        setStatusMessage("Failed to save .eml.");
        console.error(error);
      }
    });
  }

  const previewEmlButton = document.getElementById("previewEmlButton");
  if (previewEmlButton) {
    previewEmlButton.addEventListener("click", () => {
      openEmlSourcePreview(msg.emlSource || "");
    });
  }

  const frame = document.getElementById("messageFrame");
  frame.addEventListener("load", () => {
    bindExternalLinksInFrame(frame);
    fitMessageFrameToContent(frame);
  });
  frame.srcdoc = buildFrameDocument(sanitizedBody.html);
  refreshStatusMeta();
}

function handleIndexProgress(payload) {
  if (!openingInProgress || !payload) {
    return;
  }

  const phase = payload.phase || "";
  if (phase === "preparing") {
    setOpenProgress({ visible: true, indeterminate: true, value: 12 });
    setStatusMessage("Preparing SQLite index...");
    return;
  }

  if (phase === "converting-pst") {
    const convertedMessages = Number(payload.messagesConverted) || 0;
    setOpenProgress({ visible: true, indeterminate: true, value: 18 });
    setStatusMessage(`Converting PST messages${convertedMessages > 0 ? ` | ${convertedMessages} converted` : "..."}`);
    return;
  }

  if (phase === "indexing") {
    const bytesRead = Number(payload.bytesRead) || 0;
    const totalBytes = Number(payload.totalBytes) || 0;
    const messagesIndexed = Number(payload.messagesIndexed) || 0;
    const percent = totalBytes > 0 ? Math.min(100, (bytesRead / totalBytes) * 100) : 0;
    setOpenProgress({ visible: true, indeterminate: false, value: percent });
    setStatusMessage(
      `Indexing ${(percent || 0).toFixed(1)}% | ${formatBytes(bytesRead)} / ${formatBytes(totalBytes)} | ${messagesIndexed} emails`
    );
    return;
  }

  if (phase === "ready") {
    const total = Number(payload.messagesIndexed) || 0;
    setOpenProgress({ visible: true, indeterminate: false, value: 100 });
    if (payload.reused) {
      setStatusMessage(`Using existing SQLite index (${total} emails).`);
    } else {
      setStatusMessage(`Index built (${total} emails).`);
    }
  }
}

function configureDateFilter(range) {
  const from = Number(range?.from);
  const to = Number(range?.to);
  const hasRange = Number.isFinite(from) && Number.isFinite(to) && to >= from;

  if (!hasRange || !dateFilter || !dateFromRange || !dateToRange) {
    resetDateFilter();
    return;
  }

  dateRangeMinDayTs = startOfLocalDay(from);
  dateRangeMaxDayTs = startOfLocalDay(to);

  if (dateRangeMaxDayTs < dateRangeMinDayTs) {
    const fallback = dateRangeMinDayTs;
    dateRangeMinDayTs = dateRangeMaxDayTs;
    dateRangeMaxDayTs = fallback;
  }

  dateFromDayOffset = 0;
  dateToDayOffset = getMaxDateDayOffset();
  syncDateRangeInputs();
  updateDateFilterLabels();
  if (dateFilterButton) {
    dateFilterButton.hidden = false;
  }
  setDateFilterPopoverOpen(false);
  updateDateFilterButtonState();
}

function resetDateFilter() {
  dateRangeMinDayTs = null;
  dateRangeMaxDayTs = null;
  dateFromDayOffset = 0;
  dateToDayOffset = 0;
  setDateFilterPopoverOpen(false);
  if (dateFilterButton) {
    dateFilterButton.hidden = true;
  }
  if (dateBoundsLabel) {
    dateBoundsLabel.textContent = "";
  }
  if (dateFromLabel) {
    dateFromLabel.textContent = "From: Any";
  }
  if (dateToLabel) {
    dateToLabel.textContent = "To: Any";
  }
  if (dateFromRange) {
    dateFromRange.min = "0";
    dateFromRange.max = "0";
    dateFromRange.value = "0";
  }
  if (dateToRange) {
    dateToRange.min = "0";
    dateToRange.max = "0";
    dateToRange.value = "0";
  }
  if (dateRangeFill) {
    dateRangeFill.style.left = "0%";
    dateRangeFill.style.width = "100%";
  }
  updateDateFilterButtonState();
}

function resetTextFilters() {
  if (fromFilterInput) {
    fromFilterInput.value = "";
  }
  if (toFilterInput) {
    toFilterInput.value = "";
  }
  if (subjectFilterInput) {
    subjectFilterInput.value = "";
  }
  attachmentsOnlyFilterEnabled = false;
  closeTextFilterPopovers();
  updateTextFilterButtonState();
  updateTextFilterClearButtons();
}

function clearActiveDateFilter() {
  if (dateRangeMinDayTs === null || dateRangeMaxDayTs === null) {
    return;
  }

  const maxOffset = getMaxDateDayOffset();
  if (dateFromDayOffset === 0 && dateToDayOffset === maxOffset) {
    return;
  }

  dateFromDayOffset = 0;
  dateToDayOffset = maxOffset;
  syncDateRangeInputs();
  updateDateFilterLabels();
  updateDateFilterButtonState();
  currentOffset = 0;
  scheduleLoadPage();
}

function getMaxDateDayOffset() {
  if (dateRangeMinDayTs === null || dateRangeMaxDayTs === null) {
    return 0;
  }
  const diff = dateRangeMaxDayTs - dateRangeMinDayTs;
  return Math.max(0, Math.round(diff / DAY_MS));
}

function syncDateRangeInputs() {
  if (!dateFromRange || !dateToRange) {
    return;
  }

  const maxOffset = getMaxDateDayOffset();
  dateFromDayOffset = clampNumberInRange(dateFromDayOffset, 0, maxOffset, 0);
  dateToDayOffset = clampNumberInRange(dateToDayOffset, 0, maxOffset, maxOffset);
  if (dateFromDayOffset > dateToDayOffset) {
    dateFromDayOffset = dateToDayOffset;
  }

  const maxText = String(maxOffset);
  dateFromRange.min = "0";
  dateFromRange.max = maxText;
  dateFromRange.step = "1";
  dateFromRange.value = String(dateFromDayOffset);

  dateToRange.min = "0";
  dateToRange.max = maxText;
  dateToRange.step = "1";
  dateToRange.value = String(dateToDayOffset);

  updateDateRangeFill();
}

function updateDateFilterLabels() {
  if (dateRangeMinDayTs === null || dateRangeMaxDayTs === null) {
    return;
  }

  const selectedFromTs = getDateFilterDayTimestamp(dateFromDayOffset);
  const selectedToTs = getDateFilterDayTimestamp(dateToDayOffset);
  if (dateBoundsLabel) {
    dateBoundsLabel.textContent = `${formatDateRangeLabel(dateRangeMinDayTs)} to ${formatDateRangeLabel(
      dateRangeMaxDayTs
    )}`;
  }
  if (dateFromLabel) {
    dateFromLabel.textContent = `From: ${formatDateRangeLabel(selectedFromTs)}`;
  }
  if (dateToLabel) {
    dateToLabel.textContent = `To: ${formatDateRangeLabel(selectedToTs)}`;
  }
}

function getDateFilterDayTimestamp(dayOffset) {
  if (dateRangeMinDayTs === null) {
    return null;
  }
  return dateRangeMinDayTs + dayOffset * DAY_MS;
}

function isDateFilterActive() {
  if (dateRangeMinDayTs === null || dateRangeMaxDayTs === null) {
    return false;
  }
  return dateFromDayOffset > 0 || dateToDayOffset < getMaxDateDayOffset();
}

function updateDateFilterButtonState() {
  if (!dateFilterButton) {
    if (dateFilterClearButton) {
      dateFilterClearButton.disabled = !isDateFilterActive();
    }
    return;
  }
  const active = isDateFilterActive();
  dateFilterButton.classList.toggle("active", active);
  if (dateFilterClearButton) {
    dateFilterClearButton.disabled = !active;
  }
}

function updateTextFilterButtonState() {
  const fromActive = Boolean(fromFilterInput && fromFilterInput.value.trim());
  const toActive = Boolean(toFilterInput && toFilterInput.value.trim());
  const subjectActive = Boolean(subjectFilterInput && subjectFilterInput.value.trim());

  if (fromFilterButton) {
    fromFilterButton.classList.toggle("active", fromActive);
  }
  if (toFilterButton) {
    toFilterButton.classList.toggle("active", toActive);
  }
  if (subjectFilterButton) {
    subjectFilterButton.classList.toggle("active", subjectActive);
  }
  if (attachmentToggleButton) {
    attachmentToggleButton.classList.toggle("active", attachmentsOnlyFilterEnabled);
    attachmentToggleButton.setAttribute("aria-pressed", attachmentsOnlyFilterEnabled ? "true" : "false");
  }
}

function updateTextFilterClearButtons() {
  if (fromFilterClearButton) {
    fromFilterClearButton.disabled = !Boolean(fromFilterInput && fromFilterInput.value.trim());
  }
  if (toFilterClearButton) {
    toFilterClearButton.disabled = !Boolean(toFilterInput && toFilterInput.value.trim());
  }
  if (subjectFilterClearButton) {
    subjectFilterClearButton.disabled = !Boolean(subjectFilterInput && subjectFilterInput.value.trim());
  }
}

function clearTextFilter(filterKey) {
  const config = getTextFilterConfig(filterKey);
  if (!config || !config.input || !config.input.value) {
    return;
  }

  config.input.value = "";
  updateTextFilterButtonState();
  updateTextFilterClearButtons();
  currentOffset = 0;
  scheduleLoadPage();
  config.input.focus();
}

function getActiveTextFiltersPayload() {
  const senderQuery = String(fromFilterInput?.value || "").trim();
  const recipientQuery = String(toFilterInput?.value || "").trim();
  const subjectQuery = String(subjectFilterInput?.value || "").trim();

  if (!senderQuery && !recipientQuery && !subjectQuery && !attachmentsOnlyFilterEnabled) {
    return null;
  }

  return {
    senderQuery: senderQuery || null,
    recipientQuery: recipientQuery || null,
    subjectQuery: subjectQuery || null,
    attachmentsOnly: attachmentsOnlyFilterEnabled
  };
}

function isAnyTextFilterActive() {
  return getActiveTextFiltersPayload() !== null;
}

function getTextFilterConfig(filterKey) {
  if (filterKey === "from") {
    return {
      container: fromFilterContainer,
      button: fromFilterButton,
      popover: fromFilter,
      input: fromFilterInput,
      clearButton: fromFilterClearButton
    };
  }
  if (filterKey === "to") {
    return {
      container: toFilterContainer,
      button: toFilterButton,
      popover: toFilter,
      input: toFilterInput,
      clearButton: toFilterClearButton
    };
  }
  if (filterKey === "subject") {
    return {
      container: subjectFilterContainer,
      button: subjectFilterButton,
      popover: subjectFilter,
      input: subjectFilterInput,
      clearButton: subjectFilterClearButton
    };
  }
  return null;
}

function setTextFiltersVisible(visible) {
  if (filterToolsIcon) {
    filterToolsIcon.hidden = !visible;
  }
  for (const element of [fromFilterContainer, toFilterContainer, subjectFilterContainer, attachmentToggleContainer]) {
    if (element) {
      element.hidden = !visible;
    }
  }
  if (!visible) {
    resetTextFilters();
  } else {
    updateTextFilterButtonState();
    updateTextFilterClearButtons();
  }
}

function toggleAttachmentsOnlyFilter() {
  attachmentsOnlyFilterEnabled = !attachmentsOnlyFilterEnabled;
  updateTextFilterButtonState();
  currentOffset = 0;
  scheduleLoadPage();
}

function getActiveDateFilterPayload() {
  if (!isDateFilterActive()) {
    return null;
  }

  const selectedFromTs = getDateFilterDayTimestamp(dateFromDayOffset);
  const selectedToTs = getDateFilterDayTimestamp(dateToDayOffset);
  if (selectedFromTs === null || selectedToTs === null) {
    return null;
  }

  return {
    dateFrom: selectedFromTs,
    dateTo: selectedToTs + DAY_MS - 1
  };
}

function formatDateRangeLabel(timestamp) {
  if (!Number.isFinite(timestamp)) {
    return "";
  }
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric"
  }).format(new Date(timestamp));
}

function updateDateRangeFill() {
  if (!dateRangeFill) {
    return;
  }

  const maxOffset = getMaxDateDayOffset();
  if (maxOffset <= 0) {
    dateRangeFill.style.left = "0%";
    dateRangeFill.style.width = "100%";
    return;
  }

  const leftPct = (dateFromDayOffset / maxOffset) * 100;
  const rightPct = (dateToDayOffset / maxOffset) * 100;
  dateRangeFill.style.left = `${Math.max(0, Math.min(100, leftPct))}%`;
  dateRangeFill.style.width = `${Math.max(0, Math.min(100, rightPct - leftPct))}%`;
}

function clampNumberInRange(value, min, max, fallback) {
  const parsed = Number.parseInt(String(value ?? ""), 10);
  if (!Number.isFinite(parsed)) {
    return fallback;
  }
  if (parsed < min) {
    return min;
  }
  if (parsed > max) {
    return max;
  }
  return parsed;
}

function startOfLocalDay(timestamp) {
  const date = new Date(timestamp);
  date.setHours(0, 0, 0, 0);
  return date.getTime();
}

function buildFrameDocument(contentHtml) {
  const themeMeta = '<meta name="color-scheme" content="light dark" />';
  const themeStyle =
    "<style>:root{color-scheme:light dark}html,body{background:#ffffff;color:#111827}" +
    "@media (prefers-color-scheme: dark){html,body{background:#0f1622;color:#e5edf7}}" +
    "</style>";

  if (/<html[\s>]/i.test(contentHtml)) {
    let html = String(contentHtml);
    if (/<head[\s>]/i.test(html)) {
      if (!/name=["']color-scheme["']/i.test(html)) {
        html = html.replace(/<head([^>]*)>/i, `<head$1>${themeMeta}${themeStyle}`);
      } else {
        html = html.replace(/<head([^>]*)>/i, `<head$1>${themeStyle}`);
      }
      return html;
    }

    return html.replace(/<html([^>]*)>/i, `<html$1><head>${themeMeta}${themeStyle}</head>`);
  }

  return `<!doctype html>
<html>
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="color-scheme" content="light dark" />
    <style>:root{color-scheme:light dark}html,body{background:#ffffff;color:#111827}@media (prefers-color-scheme: dark){html,body{background:#0f1622;color:#e5edf7}}</style>
    <base target="_blank" />
  </head>
  <body>${contentHtml}</body>
</html>`;
}

function resolveCidUrls(html, attachments) {
  const map = new Map();

  for (const attachment of attachments || []) {
    if (!attachment.contentId || !attachment.base64) {
      continue;
    }

    const key = normalizeCid(attachment.contentId);
    const mime = attachment.contentType || "application/octet-stream";
    map.set(key, `data:${mime};base64,${attachment.base64}`);
  }

  let output = String(html || "");
  output = output.replace(/\b(src|href)\s*=\s*(['"])cid:([^'"\s>]+)\2/gi, (match, attr, quote, cid) => {
    const replacement = map.get(normalizeCid(cid));
    if (!replacement) {
      return match;
    }
    return `${attr}=${quote}${replacement}${quote}`;
  });

  output = output.replace(/url\((['"]?)cid:([^\)'"\s]+)\1\)/gi, (match, quote, cid) => {
    const replacement = map.get(normalizeCid(cid));
    if (!replacement) {
      return match;
    }
    return `url(${quote}${replacement}${quote})`;
  });

  return output;
}

function normalizeCid(value) {
  return String(value || "").trim().replace(/^<|>$/g, "").toLowerCase();
}

function bindExternalLinksInFrame(frame) {
  try {
    const doc = frame.contentDocument;
    if (!doc) {
      return;
    }

    for (const link of doc.querySelectorAll("a[href], area[href]")) {
      link.setAttribute("rel", "noopener noreferrer");
      link.setAttribute("target", "_blank");
      link.addEventListener("click", handleFrameLinkClick, true);
    }
  } catch {
    // Ignore iframe states where document access is unavailable.
  }
}

function handleFrameLinkClick(event) {
  const link = findFrameLinkFromEvent(event);
  if (!link) {
    return;
  }

  const rawHref = String(link.getAttribute("href") || "").trim();
  if (!rawHref) {
    return;
  }

  const resolvedUrl = resolveExternalUrl(rawHref, link.baseURI);
  if (!resolvedUrl || !isSupportedExternalUrl(resolvedUrl)) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  requestExternalLinkOpen(resolvedUrl);
}

function findFrameLinkFromEvent(event) {
  const currentTarget = event?.currentTarget;
  if (isAnchorLikeWithHref(currentTarget)) {
    return currentTarget;
  }

  const target = event?.target;
  if (target && typeof target.closest === "function") {
    const closest = target.closest("a[href], area[href]");
    if (isAnchorLikeWithHref(closest)) {
      return closest;
    }
  }

  if (typeof event?.composedPath === "function") {
    for (const node of event.composedPath()) {
      if (isAnchorLikeWithHref(node)) {
        return node;
      }
      if (node && typeof node.closest === "function") {
        const closest = node.closest("a[href], area[href]");
        if (isAnchorLikeWithHref(closest)) {
          return closest;
        }
      }
    }
  }

  return null;
}

function isAnchorLikeWithHref(value) {
  return Boolean(
    value &&
      value.nodeType === 1 &&
      typeof value.matches === "function" &&
      value.matches("a[href], area[href]")
  );
}

function resolveExternalUrl(rawHref, baseUri) {
  try {
    return new URL(rawHref, baseUri || undefined).toString();
  } catch {
    return "";
  }
}

function isSupportedExternalUrl(urlValue) {
  try {
    const protocol = new URL(urlValue).protocol.toLowerCase();
    return protocol === "http:" || protocol === "https:" || protocol === "mailto:" || protocol === "tel:";
  } catch {
    return false;
  }
}

function sanitizeEmailHtml(html, options = {}) {
  const { allowRemoteContent = true } = options;
  const template = document.createElement("template");
  template.innerHTML = String(html || "");
  let blockedRemoteContent = false;

  const blockedTags = ["script", "iframe", "object", "embed", "form"];
  for (const tag of blockedTags) {
    for (const node of template.content.querySelectorAll(tag)) {
      node.remove();
    }
  }

  const elements = template.content.querySelectorAll("*");
  for (const element of elements) {
    if (!allowRemoteContent && element.tagName === "STYLE") {
      const sanitizedCss = stripRemoteContentFromCss(element.textContent || "");
      if (sanitizedCss.changed) {
        blockedRemoteContent = true;
        if (sanitizedCss.value) {
          element.textContent = sanitizedCss.value;
        } else {
          element.remove();
          continue;
        }
      }
    }

    for (const attr of Array.from(element.attributes)) {
      const attrName = attr.name.toLowerCase();
      const rawAttrValue = attr.value.trim();
      const attrValue = rawAttrValue.toLowerCase();

      if (attrName.startsWith("on")) {
        element.removeAttribute(attr.name);
        continue;
      }

      if ((attrName === "href" || attrName === "src" || attrName === "srcset") && isUnsafeUrl(attrValue)) {
        element.removeAttribute(attr.name);
        continue;
      }

      if (!allowRemoteContent) {
        if (attrName === "style") {
          const sanitizedStyle = stripRemoteContentFromCss(rawAttrValue);
          if (sanitizedStyle.changed) {
            blockedRemoteContent = true;
            if (sanitizedStyle.value) {
              element.setAttribute(attr.name, sanitizedStyle.value);
            } else {
              element.removeAttribute(attr.name);
            }
          }
          continue;
        }

        if (attrName === "poster" || attrName === "background") {
          if (isRemoteResourceUrl(rawAttrValue)) {
            element.removeAttribute(attr.name);
            blockedRemoteContent = true;
          }
          continue;
        }

        if (attrName === "src" || attrName === "srcset") {
          if (containsRemoteResource(attrName, rawAttrValue)) {
            element.removeAttribute(attr.name);
            blockedRemoteContent = true;
          }
          continue;
        }

        if (attrName === "href" && isRemoteResourceUrl(rawAttrValue)) {
          const tagName = element.tagName;
          if (tagName === "LINK" || tagName === "BASE") {
            element.removeAttribute(attr.name);
            blockedRemoteContent = true;
            continue;
          }

          if (tagName !== "A" && tagName !== "AREA") {
            element.removeAttribute(attr.name);
            blockedRemoteContent = true;
          }
        }
      }
    }
  }

  return {
    html: template.innerHTML,
    blockedRemoteContent
  };
}

function isUnsafeUrl(urlValue) {
  return (
    urlValue.startsWith("javascript:") ||
    urlValue.startsWith("vbscript:") ||
    urlValue.startsWith("data:text/html")
  );
}

function containsRemoteResource(attrName, value) {
  if (attrName === "srcset") {
    return String(value || "")
      .split(",")
      .some((entry) => {
        const candidate = entry.trim().split(/\s+/)[0] || "";
        return isRemoteResourceUrl(candidate);
      });
  }

  return isRemoteResourceUrl(value);
}

function isRemoteResourceUrl(urlValue) {
  const value = String(urlValue || "").trim().toLowerCase();
  return (
    value.startsWith("http://") ||
    value.startsWith("https://") ||
    value.startsWith("//") ||
    value.startsWith("ftp://")
  );
}

function stripRemoteContentFromCss(cssText) {
  let value = String(cssText || "");
  let changed = false;

  value = value.replace(/@import\s+(?:url\()?\s*(['"]?)([^'")\s]+)\1\s*\)?[^;]*;/gi, (match, quote, url) => {
    if (!isRemoteResourceUrl(url)) {
      return match;
    }
    changed = true;
    return "";
  });

  value = value.replace(/url\(\s*(['"]?)([^'")]+)\1\s*\)/gi, (match, quote, url) => {
    if (!isRemoteResourceUrl(url)) {
      return match;
    }
    changed = true;
    return "url()";
  });

  return {
    value: value.trim(),
    changed
  };
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function isPreviewableAttachment(attachment) {
  const contentType = String(attachment?.contentType || "").toLowerCase();
  if (!contentType || !attachment?.base64) {
    return false;
  }
  return contentType.startsWith("image/") || contentType === "application/pdf";
}

function formatBytes(bytes) {
  const value = Number(bytes) || 0;
  if (value < 1024) {
    return `${value} B`;
  }
  const kb = value / 1024;
  if (kb < 1024) {
    return `${kb.toFixed(1)} KB`;
  }
  const mb = kb / 1024;
  if (mb < 1024) {
    return `${mb.toFixed(1)} MB`;
  }
  return `${(mb / 1024).toFixed(2)} GB`;
}

function getSenderDisplay(fromValue) {
  const raw = String(fromValue || "").trim();
  if (!raw) {
    return "Unknown sender";
  }

  const match = raw.match(/^(.*?)\s*<([^>]+)>$/);
  if (match) {
    const name = match[1].replace(/^"|"$/g, "").trim();
    return name || match[2].trim();
  }

  return raw.replace(/^"|"$/g, "");
}

function formatListDate(dateValue) {
  const raw = String(dateValue || "").trim();
  if (!raw) {
    return "";
  }

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    return raw;
  }

  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  }).format(parsed);
}

function compactText(value) {
  let text = String(value || "");
  text = text
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ");
  text = decodeHtmlEntities(text);
  text = decodeQuotedPrintablePreview(text);
  text = text
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, " ")
    .replace(/\b[A-Za-z0-9+/]{120,}={0,2}\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return text;
}

function decodeHtmlEntities(value) {
  const parser = document.createElement("textarea");
  parser.innerHTML = String(value || "");
  return parser.value || "";
}

function decodeQuotedPrintablePreview(value) {
  const input = String(value || "");
  if (!/=[A-Fa-f0-9]{2}|=\r?\n/.test(input)) {
    return input;
  }

  return input
    .replace(/=\r?\n/g, "")
    .replace(/=([A-Fa-f0-9]{2})/g, (_, hex) => String.fromCharCode(Number.parseInt(hex, 16)));
}

function buildEmlFileName(msg) {
  const subject = (msg?.subject || "message").trim() || "message";
  const date = (msg?.date || "").trim();
  const base = date ? `${subject} ${date}` : subject;
  const sanitized = base
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
  return `${sanitized || "message"}.eml`;
}

function fitMessageFrameToContent(frame) {
  try {
    const doc = frame.contentDocument;
    if (!doc) {
      return;
    }

    const recalc = () => {
      try {
        const body = doc.body;
        const html = doc.documentElement;
        const nextHeight = Math.max(
          260,
          body ? body.scrollHeight : 0,
          body ? body.offsetHeight : 0,
          html ? html.scrollHeight : 0,
          html ? html.offsetHeight : 0
        );
        frame.style.height = `${nextHeight}px`;
      } catch {
        // Ignore dynamic resize failures.
      }
    };

    recalc();
    window.requestAnimationFrame(recalc);
    setTimeout(recalc, 50);
    setTimeout(recalc, 250);

    if (doc.fonts && doc.fonts.ready) {
      doc.fonts.ready.then(recalc).catch(() => {});
    }

    for (const image of Array.from(doc.images || [])) {
      if (!image.complete) {
        image.addEventListener("load", recalc, { once: true });
        image.addEventListener("error", recalc, { once: true });
      }
    }
  } catch {
    // Access can fail for unusual iframe states.
  }
}

function setStatusMessage(value) {
  statusMessage.textContent = value;
  refreshStatusMeta();
}

function setOpenProgress({ visible, indeterminate = false, value = 0 }) {
  if (!openProgress || !openProgressBar) {
    return;
  }

  const shouldShow = Boolean(visible);
  openProgress.hidden = !shouldShow;
  openProgress.classList.toggle("indeterminate", shouldShow && indeterminate);

  if (!shouldShow) {
    openProgress.style.removeProperty("--open-progress-value");
    openProgress.setAttribute("aria-valuenow", "0");
    openProgressBar.style.transform = "scaleX(0)";
    return;
  }

  const clamped = Math.max(0, Math.min(100, Number(value) || 0));
  if (indeterminate) {
    openProgress.removeAttribute("aria-valuenow");
    openProgressBar.style.transform = "";
    return;
  }

  openProgress.setAttribute("aria-valuenow", clamped.toFixed(0));
  openProgressBar.style.transform = `scaleX(${clamped / 100})`;
}

function setOpenButtonBusy(isBusy) {
  if (!openButton) {
    return;
  }
  openButton.disabled = Boolean(isBusy);
}

function refreshStatusMeta() {
  if (openingInProgress) {
    statusMeta.textContent = "Opening...";
    return;
  }

  if (!dbPath) {
    if (!currentPageMessages.length) {
      statusMeta.textContent = "0 emails";
      return;
    }

    const position = resultIndexById.get(selectedMessageId) || (selectedMessageId ? 1 : 0);
    const positionText = position > 0 ? position : "-";
    statusMeta.textContent = `Position ${positionText} / ${totalResults || currentPageMessages.length} (${totalMessages} total messages)`;
    return;
  }

  if (totalResults === 0) {
    statusMeta.textContent = `Position - / 0 (${totalMessages} total messages)`;
    return;
  }

  const position = resultIndexById.get(selectedMessageId) || 0;
  const positionText = position > 0 ? position : "-";
  statusMeta.textContent = `Position ${positionText} / ${totalResults} (${totalMessages} total messages)`;
}

function initSplitter() {
  if (!layout || !splitter) {
    return;
  }

  const savedWidth = Number.parseInt(localStorage.getItem("leftPaneWidthPx") || "", 10);
  if (Number.isFinite(savedWidth) && savedWidth > 0) {
    layout.style.setProperty("--left-pane-width", `${savedWidth}px`);
  }

  splitter.addEventListener("pointerdown", (event) => {
    if (window.matchMedia("(max-width: 900px)").matches) {
      return;
    }

    event.preventDefault();
    splitter.setPointerCapture(event.pointerId);
    splitter.classList.add("dragging");
    document.body.classList.add("resizing");

    const onMove = (moveEvent) => {
      const nextWidth = calculateLeftPaneWidth(moveEvent.clientX);
      layout.style.setProperty("--left-pane-width", `${nextWidth}px`);
    };

    const onUp = () => {
      splitter.classList.remove("dragging");
      document.body.classList.remove("resizing");
      const currentWidth = getCurrentLeftPaneWidthPx();
      localStorage.setItem("leftPaneWidthPx", String(currentWidth));
      splitter.removeEventListener("pointermove", onMove);
      splitter.removeEventListener("pointerup", onUp);
      splitter.removeEventListener("pointercancel", onUp);
    };

    splitter.addEventListener("pointermove", onMove);
    splitter.addEventListener("pointerup", onUp);
    splitter.addEventListener("pointercancel", onUp);
  });

  splitter.addEventListener("keydown", (event) => {
    if (window.matchMedia("(max-width: 900px)").matches) {
      return;
    }

    if (event.key !== "ArrowLeft" && event.key !== "ArrowRight") {
      return;
    }

    event.preventDefault();
    const direction = event.key === "ArrowLeft" ? -1 : 1;
    const delta = event.shiftKey ? 40 : 20;
    const current = getCurrentLeftPaneWidthPx();
    const target = clampLeftPaneWidth(current + direction * delta);
    layout.style.setProperty("--left-pane-width", `${target}px`);
    localStorage.setItem("leftPaneWidthPx", String(target));
  });
}

function calculateLeftPaneWidth(pointerClientX) {
  const rect = layout.getBoundingClientRect();
  const splitterWidth = splitter.getBoundingClientRect().width || 10;
  return clampLeftPaneWidth(pointerClientX - rect.left, rect.width, splitterWidth);
}

function clampLeftPaneWidth(value, containerWidth, splitterWidth) {
  const totalWidth = containerWidth || layout.getBoundingClientRect().width;
  const draggerWidth = splitterWidth || splitter.getBoundingClientRect().width || 10;
  const minLeft = 260;
  const minRight = 360;
  const maxLeft = Math.max(minLeft, totalWidth - minRight - draggerWidth);
  return Math.min(maxLeft, Math.max(minLeft, value));
}

function getCurrentLeftPaneWidthPx() {
  const resolved = getComputedStyle(layout).getPropertyValue("--left-pane-width").trim();
  if (resolved.endsWith("px")) {
    const px = Number.parseFloat(resolved.slice(0, -2));
    if (Number.isFinite(px)) {
      return px;
    }
  }

  const leftPanel = document.querySelector(".mail-list-panel");
  return leftPanel ? Math.round(leftPanel.getBoundingClientRect().width) : 320;
}
