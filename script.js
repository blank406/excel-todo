const STORAGE_KEY = "calendar.events.v1";
const IMPORTS_STORAGE_KEY = "calendar.imports.v1";
const PERSONAL_SOURCE = Object.freeze({ sourceType: "personal", sourceName: "개인 일정" });

const calendarGrid = document.querySelector("#calendar-grid");
const calendarTitle = document.querySelector("#calendar-title");
const selectedWeekday = document.querySelector("#selected-weekday");
const selectedDay = document.querySelector("#selected-day");
const selectedMonth = document.querySelector("#selected-month");
const selectedEvents = document.querySelector("#selected-events");
const emptyState = document.querySelector("#empty-state");
const scheduleCount = document.querySelector(".schedule-count");
const eventModal = document.querySelector("#event-modal");
const eventForm = document.querySelector("#event-form");
const modalTitle = document.querySelector("#modal-title");
const deleteEventButton = document.querySelector("#delete-event");
const titleInput = document.querySelector("#event-title");
const dateInput = document.querySelector("#event-date");
const startTimeInput = document.querySelector("#event-start-time");
const endTimeInput = document.querySelector("#event-end-time");
const memoInput = document.querySelector("#event-memo");
const toast = document.querySelector("#toast");
const excelModal = document.querySelector("#excel-modal");
const excelUploadView = document.querySelector("#excel-upload-view");
const excelResultView = document.querySelector("#excel-result-view");
const excelDropZone = document.querySelector("#excel-drop-zone");
const excelFileInput = document.querySelector("#excel-file-input");
const excelError = document.querySelector("#excel-error");
const excelFileName = document.querySelector("#excel-file-name");
const excelSheetCount = document.querySelector("#excel-sheet-count");
const excelSheetList = document.querySelector("#excel-sheet-list");
const excelReviewView = document.querySelector("#excel-review-view");
const excelEditView = document.querySelector("#excel-edit-view");
const excelReviewSummary = document.querySelector("#excel-review-summary");
const duplicateNotice = document.querySelector("#duplicate-notice");
const monthFilters = document.querySelector("#month-filters");
const importCandidateList = document.querySelector("#import-candidate-list");
const selectionSummary = document.querySelector("#selection-summary");
const selectFilteredButton = document.querySelector("#select-filtered");
const deselectFilteredButton = document.querySelector("#deselect-filtered");
const importSelectedButton = document.querySelector("#excel-import-selected");
const candidateForm = document.querySelector("#excel-candidate-form");
const candidateTitleInput = document.querySelector("#candidate-title");
const candidateDateInput = document.querySelector("#candidate-date");
const candidateStartTimeInput = document.querySelector("#candidate-start-time");
const candidateEndTimeInput = document.querySelector("#candidate-end-time");
const candidateVenueInput = document.querySelector("#candidate-venue");
const candidateEventTypeInput = document.querySelector("#candidate-event-type");
const candidateMemoInput = document.querySelector("#candidate-memo");
const importsModal = document.querySelector("#imports-modal");
const importsList = document.querySelector("#imports-list");
const importsEmpty = document.querySelector("#imports-empty");

const today = startOfDay(new Date());
const state = {
  displayYear: today.getFullYear(),
  displayMonth: today.getMonth(),
  selectedDate: today,
  events: loadEvents(),
  imports: loadImports(),
  editingEventId: null,
  lastFocusedElement: null,
};

const excelState = {
  workbook: null,
  sheetData: {},
  fileName: "",
  lastFocusedElement: null,
  candidates: [],
  unparsedCandidates: [],
  monthFilter: 0,
  editingCandidateId: null,
};

const weekdayNames = ["일요일", "월요일", "화요일", "수요일", "목요일", "금요일", "토요일"];

function loadEvents() {
  try {
    const savedEvents = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    return Array.isArray(savedEvents) ? savedEvents.filter(isValidEvent) : [];
  } catch (error) {
    console.warn("저장된 일정을 불러오지 못했습니다.", error);
    return [];
  }
}

function saveEvents() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.events));
    return true;
  } catch (error) {
    console.warn("일정을 저장하지 못했습니다.", error);
    showToast("일정을 저장하지 못했습니다.");
    return false;
  }
}

function loadImports() {
  try {
    const savedImports = JSON.parse(localStorage.getItem(IMPORTS_STORAGE_KEY) || "[]");
    return Array.isArray(savedImports) ? savedImports.filter(isValidImportRecord) : [];
  } catch (error) {
    console.warn("가져오기 기록을 불러오지 못했습니다.", error);
    return [];
  }
}

function saveImports() {
  try {
    localStorage.setItem(IMPORTS_STORAGE_KEY, JSON.stringify(state.imports));
    return true;
  } catch (error) {
    console.warn("가져오기 기록을 저장하지 못했습니다.", error);
    showToast("가져오기 기록을 저장하지 못했습니다.");
    return false;
  }
}

function isValidImportRecord(record) {
  return record && typeof record.id === "string" && typeof record.sourceName === "string"
    && typeof record.fileName === "string";
}

function isValidEvent(event) {
  return event && typeof event.id === "string" && typeof event.title === "string"
    && /^\d{4}-\d{2}-\d{2}$/.test(event.date);
}

function createEventId() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") return window.crypto.randomUUID();
  return `event-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function createImportId() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") return `import_${window.crypto.randomUUID()}`;
  return `import_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function parseDateKey(dateKey) {
  const [year, month, day] = dateKey.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function isSameDate(first, second) {
  return first.getFullYear() === second.getFullYear()
    && first.getMonth() === second.getMonth()
    && first.getDate() === second.getDate();
}

function formatDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function sortEvents(events) {
  return [...events].sort((first, second) => {
    const firstHasTime = Boolean(first.startTime);
    const secondHasTime = Boolean(second.startTime);
    if (firstHasTime !== secondHasTime) return firstHasTime ? -1 : 1;
    if (firstHasTime && first.startTime !== second.startTime) return first.startTime.localeCompare(second.startTime);
    return first.title.localeCompare(second.title, "ko");
  });
}

function getEventsForDate(dateKey) {
  return sortEvents(state.events.filter((event) => event.date === dateKey));
}

function createDateCell(date) {
  const button = document.createElement("button");
  const isCurrentMonth = date.getMonth() === state.displayMonth;
  const classes = ["date-cell"];
  if (!isCurrentMonth) classes.push("other-month");
  if (isSameDate(date, today)) classes.push("today");
  if (isSameDate(date, state.selectedDate)) classes.push("selected");

  button.type = "button";
  button.className = classes.join(" ");
  button.dataset.date = formatDateKey(date);
  button.setAttribute("role", "gridcell");
  button.setAttribute("aria-label", `${date.getFullYear()}년 ${date.getMonth() + 1}월 ${date.getDate()}일`);
  button.setAttribute("aria-selected", String(isSameDate(date, state.selectedDate)));

  const dateNumber = document.createElement("span");
  dateNumber.className = "date-number";
  dateNumber.textContent = date.getDate();
  const eventList = document.createElement("span");
  eventList.className = "event-list";
  eventList.setAttribute("aria-hidden", "true");
  button.append(dateNumber, eventList);
  button.addEventListener("click", () => selectDate(date));
  return button;
}

function renderCalendarEvents() {
  calendarGrid.querySelectorAll(".date-cell").forEach((cell) => {
    const eventList = cell.querySelector(".event-list");
    const events = getEventsForDate(cell.dataset.date);

    events.slice(0, 3).forEach((event) => {
      const eventRow = document.createElement("span");
      eventRow.className = "calendar-event";
      eventRow.dataset.source = event.sourceType || "unknown";

      const sourceIndicator = document.createElement("span");
      sourceIndicator.className = "source-indicator";
      sourceIndicator.setAttribute("aria-hidden", "true");
      eventRow.appendChild(sourceIndicator);

      const title = document.createElement("span");
      title.className = "event-title";
      title.textContent = event.title;
      eventRow.title = event.title;
      eventRow.appendChild(title);
      eventList.appendChild(eventRow);
    });

    const remainingCount = events.length - 3;
    if (remainingCount > 0) {
      const moreEvents = document.createElement("span");
      moreEvents.className = "more-events";
      moreEvents.textContent = `+${remainingCount}개`;
      eventList.appendChild(moreEvents);
    }
  });
}

function renderCalendar() {
  calendarGrid.replaceChildren();
  calendarTitle.textContent = `${state.displayYear}년 ${state.displayMonth + 1}월`;
  const firstDay = new Date(state.displayYear, state.displayMonth, 1);
  const lastDay = new Date(state.displayYear, state.displayMonth + 1, 0);
  const calendarStart = new Date(state.displayYear, state.displayMonth, 1 - firstDay.getDay());
  const calendarEnd = new Date(state.displayYear, state.displayMonth, lastDay.getDate() + (6 - lastDay.getDay()));
  const visibleDayCount = Math.round((calendarEnd - calendarStart) / 86400000) + 1;
  const fragment = document.createDocumentFragment();
  for (let index = 0; index < visibleDayCount; index += 1) {
    const date = new Date(calendarStart.getFullYear(), calendarStart.getMonth(), calendarStart.getDate() + index);
    fragment.appendChild(createDateCell(date));
  }
  calendarGrid.appendChild(fragment);
  renderCalendarEvents();
}

function createSelectedEvent(event) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "selected-event";
  button.setAttribute("aria-label", `${event.title} 일정 수정`);
  if (event.startTime) {
    const time = document.createElement("span");
    time.className = "selected-event-time";
    time.textContent = event.endTime ? `${event.startTime} - ${event.endTime}` : event.startTime;
    button.appendChild(time);
  }
  const title = document.createElement("span");
  title.className = "selected-event-title";
  title.textContent = event.title;
  button.appendChild(title);
  if (event.memo) {
    const memo = document.createElement("span");
    memo.className = "selected-event-memo";
    memo.textContent = event.memo;
    button.appendChild(memo);
  }
  const source = document.createElement("span");
  source.className = "selected-event-source";
  source.textContent = [event.sourceName || event.sourceType, event.eventType].filter(Boolean).join(" · ");
  if (event.venue) {
    const venue = document.createElement("span");
    venue.className = "selected-event-venue";
    venue.textContent = event.venue;
    button.appendChild(venue);
  }
  button.appendChild(source);
  button.addEventListener("click", () => openEventModal(event));
  return button;
}

function renderSelectedDateEvents() {
  const events = getEventsForDate(formatDateKey(state.selectedDate));
  selectedEvents.replaceChildren(...events.map(createSelectedEvent));
  scheduleCount.textContent = events.length;
  scheduleCount.setAttribute("aria-label", `등록된 일정 ${events.length}개`);
  emptyState.hidden = events.length > 0;
}

function renderSelectedDate() {
  const date = state.selectedDate;
  selectedWeekday.textContent = weekdayNames[date.getDay()];
  selectedDay.textContent = date.getDate();
  selectedMonth.textContent = `${date.getFullYear()}년 ${date.getMonth() + 1}월`;
  renderSelectedDateEvents();
}

function selectDate(date) {
  state.selectedDate = startOfDay(date);
  if (date.getFullYear() !== state.displayYear || date.getMonth() !== state.displayMonth) {
    state.displayYear = date.getFullYear();
    state.displayMonth = date.getMonth();
  }
  renderCalendar();
  renderSelectedDate();
}

function changeMonth(offset) {
  const nextMonth = new Date(state.displayYear, state.displayMonth + offset, 1);
  state.displayYear = nextMonth.getFullYear();
  state.displayMonth = nextMonth.getMonth();
  const lastDay = new Date(state.displayYear, state.displayMonth + 1, 0).getDate();
  state.selectedDate = new Date(state.displayYear, state.displayMonth, Math.min(state.selectedDate.getDate(), lastDay));
  renderCalendar();
  renderSelectedDate();
}

function openEventModal(event = null) {
  state.lastFocusedElement = document.activeElement;
  state.editingEventId = event ? event.id : null;
  eventForm.reset();
  modalTitle.textContent = event ? "일정 수정" : "일정 추가";
  deleteEventButton.hidden = !event;
  titleInput.value = event?.title || "";
  dateInput.value = event?.date || formatDateKey(state.selectedDate);
  startTimeInput.value = event?.startTime || "";
  endTimeInput.value = event?.endTime || "";
  memoInput.value = event?.memo || "";
  eventModal.hidden = false;
  document.body.classList.add("modal-open");
  window.setTimeout(() => titleInput.focus(), 0);
}

function closeEventModal() {
  eventModal.hidden = true;
  document.body.classList.remove("modal-open");
  state.editingEventId = null;
  state.lastFocusedElement?.focus();
}

function getEventFormValues() {
  return {
    title: titleInput.value.trim(),
    date: dateInput.value,
    startTime: startTimeInput.value,
    endTime: endTimeInput.value,
    memo: memoInput.value.trim(),
  };
}

function addEvent(values) {
  state.events.push({ id: createEventId(), ...values, ...PERSONAL_SOURCE });
}

function updateEvent(eventId, values) {
  const eventIndex = state.events.findIndex((event) => event.id === eventId);
  if (eventIndex !== -1) state.events[eventIndex] = { ...state.events[eventIndex], ...values };
}

function deleteEvent(eventId) {
  const event = state.events.find((item) => item.id === eventId);
  if (!event || !window.confirm(`“${event.title}” 일정을 삭제할까요?`)) return;
  state.events = state.events.filter((item) => item.id !== eventId);
  saveEvents();
  closeEventModal();
  renderCalendar();
  renderSelectedDateEvents();
  showToast("일정을 삭제했습니다.");
}

function handleEventSubmit(event) {
  event.preventDefault();
  const values = getEventFormValues();
  if (!values.title || !values.date) return;
  if (state.editingEventId) updateEvent(state.editingEventId, values);
  else addEvent(values);
  saveEvents();
  const savedDate = parseDateKey(values.date);
  closeEventModal();
  selectDate(savedDate);
  showToast("일정을 저장했습니다.");
}

function openExcelModal() {
  excelState.lastFocusedElement = document.activeElement;
  excelError.hidden = true;
  excelUploadView.hidden = Boolean(excelState.workbook);
  excelResultView.hidden = !excelState.workbook;
  excelReviewView.hidden = true;
  excelEditView.hidden = true;
  document.querySelector("#excel-modal-title").textContent = "Excel 일정 가져오기";
  excelModal.hidden = false;
  document.body.classList.add("modal-open");
}

function closeExcelModal() {
  excelModal.hidden = true;
  excelDropZone.classList.remove("drag-over");
  document.body.classList.remove("modal-open");
  excelState.lastFocusedElement?.focus();
}

function isExcelFile(file) {
  return Boolean(file) && /\.(xlsx|xls)$/i.test(file.name);
}

async function readWorkbook(file) {
  if (!window.XLSX) throw new Error("SheetJS를 불러오지 못했습니다.");
  const arrayBuffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(arrayBuffer);
  const sheetData = Object.fromEntries(
    workbook.SheetNames.map((sheetName) => [
      sheetName,
      window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 }),
    ]),
  );
  return { workbook, sheetData };
}

async function handleExcelFile(file) {
  excelError.hidden = true;

  if (!isExcelFile(file)) {
    excelError.textContent = "Excel 파일(.xlsx, .xls)만 업로드할 수 있습니다.";
    excelError.hidden = false;
    excelFileInput.value = "";
    return;
  }

  try {
    const { workbook, sheetData } = await readWorkbook(file);
    excelState.workbook = workbook;
    excelState.sheetData = sheetData;
    excelState.fileName = file.name;

    console.group(`Excel workbook: ${file.name}`);
    console.log("SheetNames", workbook.SheetNames);
    workbook.SheetNames.forEach((sheetName) => {
      console.log(sheetName, sheetData[sheetName]);
    });
    console.groupEnd();

    showExcelUploadResult();
  } catch (error) {
    console.error("Excel 파일 읽기 오류", error);
    excelError.textContent = "Excel 파일을 읽지 못했습니다. 파일을 확인해 주세요.";
    excelError.hidden = false;
  } finally {
    excelFileInput.value = "";
  }
}

function showExcelUploadResult() {
  const sheetNames = excelState.workbook.SheetNames;
  document.querySelector("#excel-modal-title").textContent = "Excel 일정 가져오기";
  excelFileName.textContent = excelState.fileName;
  excelSheetCount.textContent = `${sheetNames.length}개의 시트`;
  excelSheetList.replaceChildren(...sheetNames.map((sheetName) => {
    const item = document.createElement("li");
    item.textContent = sheetName;
    return item;
  }));
  excelUploadView.hidden = true;
  excelResultView.hidden = false;
  excelReviewView.hidden = true;
  excelEditView.hidden = true;
}

function analyzeWorkbook() {
  if (!excelState.workbook || !window.ExcelParser) {
    showToast("Excel 분석기를 사용할 수 없습니다.");
    return;
  }

  const result = window.ExcelParser.parseWorkbook(excelState.workbook, excelState.fileName);
  excelState.candidates = result.events.map((event, index) => ({
    ...event,
    candidateId: `candidate-${index + 1}`,
    selected: true,
    duplicate: false,
  }));
  excelState.unparsedCandidates = result.unparsedCandidates;
  excelState.monthFilter = 0;
  refreshDuplicateFlags();

  console.table(excelState.candidates.map(({ selected, duplicate, candidateId, ...event }) => event));
  console.log("unparsedCandidates", excelState.unparsedCandidates);
  showExcelReview();
}

function getEventSignature(event) {
  return [event.date, event.title.trim(), event.startTime || "", event.sourceName || ""].join("|").toLowerCase();
}

function refreshDuplicateFlags() {
  const existingSignatures = new Set(
    state.events.filter((event) => event.sourceType === "excel").map(getEventSignature),
  );
  const candidateSignatures = new Set();

  excelState.candidates.forEach((candidate) => {
    const signature = getEventSignature(candidate);
    candidate.duplicate = existingSignatures.has(signature) || candidateSignatures.has(signature);
    candidateSignatures.add(signature);
  });
}

function showExcelReview() {
  document.querySelector("#excel-modal-title").textContent = "일정 분석 및 검토";
  excelUploadView.hidden = true;
  excelResultView.hidden = true;
  excelEditView.hidden = true;
  excelReviewView.hidden = false;
  renderMonthFilters();
  renderImportCandidates();
}

function renderMonthFilters() {
  const filters = [{ value: 0, label: "전체" }];
  for (let month = 1; month <= 12; month += 1) filters.push({ value: month, label: `${month}월` });

  monthFilters.replaceChildren(...filters.map((filter) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `month-filter${excelState.monthFilter === filter.value ? " active" : ""}`;
    button.textContent = filter.label;
    button.addEventListener("click", () => {
      excelState.monthFilter = filter.value;
      renderMonthFilters();
      renderImportCandidates();
    });
    return button;
  }));
}

function renderImportCandidates() {
  const validCandidates = [];

  excelState.candidates.forEach((candidate) => {
    if (isValidCandidateDate(candidate.date)) validCandidates.push(candidate);
    else {
      console.warn("Invalid event date", candidate);
    }
  });

  const visibleCandidates = validCandidates.filter((candidate) => (
    !excelState.monthFilter || Number(candidate.date.slice(5, 7)) === excelState.monthFilter
  ));
  const duplicateCount = excelState.candidates.filter((candidate) => candidate.duplicate).length;
  const groupedCandidates = groupCandidatesByDate(visibleCandidates);

  excelReviewSummary.textContent = `총 ${excelState.candidates.length}개의 일정을 찾았습니다.`;
  duplicateNotice.hidden = duplicateCount === 0;
  duplicateNotice.textContent = `기존 일정과 동일하거나 중복 가능성이 있는 ${duplicateCount}개 일정은 가져오기에서 제외됩니다.`;
  importCandidateList.replaceChildren(...groupedCandidates.map(([date, candidates]) => (
    createImportDateGroup(date, candidates)
  )));
  updateSelectionUI();

  if (!visibleCandidates.length) {
    const empty = document.createElement("p");
    empty.className = "candidate-empty";
    empty.textContent = "이 월에서 찾은 일정이 없습니다.";
    importCandidateList.appendChild(empty);
  }
}

function getCandidatesForCurrentFilter() {
  if (!excelState.monthFilter) return excelState.candidates;
  return excelState.candidates.filter((candidate) => (
    isValidCandidateDate(candidate.date)
      && Number(candidate.date.slice(5, 7)) === excelState.monthFilter
  ));
}

function setFilteredCandidatesSelected(selected) {
  getCandidatesForCurrentFilter().forEach((candidate) => { candidate.selected = selected; });
  renderImportCandidates();
}

function updateSelectionUI() {
  const filteredCandidates = getCandidatesForCurrentFilter();
  const filteredSelectedCount = filteredCandidates.filter((candidate) => candidate.selected).length;
  const totalSelectedCount = excelState.candidates.filter((candidate) => candidate.selected).length;
  const filterLabel = excelState.monthFilter ? `${excelState.monthFilter}월` : "전체";

  selectionSummary.textContent = `${filterLabel} 일정 ${filteredCandidates.length}개 · ${filteredSelectedCount}개 선택`;
  selectFilteredButton.disabled = filteredCandidates.length === 0;
  deselectFilteredButton.disabled = filteredCandidates.length === 0;
  importSelectedButton.textContent = `선택한 일정 가져오기 (${totalSelectedCount})`;
  importSelectedButton.disabled = totalSelectedCount === 0;
}

function isValidCandidateDate(date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || "")) return false;
  const parsed = parseDateKey(date);
  return !Number.isNaN(parsed.getTime()) && formatDateKey(parsed) === date;
}

function groupCandidatesByDate(candidates) {
  const groups = new Map();
  [...candidates].sort((first, second) => first.date.localeCompare(second.date)).forEach((candidate) => {
    if (!groups.has(candidate.date)) groups.set(candidate.date, []);
    groups.get(candidate.date).push(candidate);
  });
  return [...groups.entries()].map(([date, items]) => [date, sortEvents(items)]);
}

function createImportDateGroup(dateKey, candidates) {
  const group = document.createElement("section");
  group.className = "import-date-group";

  const heading = document.createElement("h3");
  heading.className = "import-date-heading";
  const date = parseDateKey(dateKey);
  heading.textContent = `${date.getMonth() + 1}월 ${date.getDate()}일`;

  const items = document.createElement("div");
  items.className = "import-date-items";
  items.append(...candidates.map(createImportCandidate));
  group.append(heading, items);
  return group;
}

function createImportCandidate(candidate) {
  const row = document.createElement("div");
  row.className = "import-candidate";

  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.className = "candidate-check";
  checkbox.checked = candidate.selected;
  checkbox.setAttribute("aria-label", `${candidate.title} 선택`);
  checkbox.addEventListener("change", () => {
    candidate.selected = checkbox.checked;
    updateSelectionUI();
  });

  const content = document.createElement("button");
  content.type = "button";
  content.className = "candidate-content";
  content.setAttribute("aria-label", `${candidate.title} 수정`);

  const title = document.createElement("span");
  title.className = "candidate-title";
  title.textContent = candidate.title;
  content.appendChild(title);

  const metaText = [candidate.startTime, candidate.venue].filter(Boolean).join(" · ");
  if (metaText) {
    const meta = document.createElement("span");
    meta.className = "candidate-meta";
    meta.textContent = metaText;
    content.appendChild(meta);
  }

  const sourceText = [candidate.sourceName, candidate.eventType].filter(Boolean).join(" · ");
  if (sourceText) {
    const source = document.createElement("span");
    source.className = "candidate-source";
    source.textContent = sourceText;
    content.appendChild(source);
  }

  if (candidate.duplicate) {
    const duplicate = document.createElement("span");
    duplicate.className = "candidate-duplicate";
    duplicate.textContent = "중복 가능성 있음";
    content.appendChild(duplicate);
  }

  content.addEventListener("click", () => openCandidateEditor(candidate.candidateId));
  row.append(checkbox, content);
  return row;
}

function openCandidateEditor(candidateId) {
  const candidate = excelState.candidates.find((item) => item.candidateId === candidateId);
  if (!candidate) return;
  excelState.editingCandidateId = candidateId;
  document.querySelector("#excel-modal-title").textContent = "가져올 일정 수정";
  candidateTitleInput.value = candidate.title;
  candidateDateInput.value = candidate.date;
  candidateStartTimeInput.value = candidate.startTime || "";
  candidateEndTimeInput.value = candidate.endTime || "";
  candidateVenueInput.value = candidate.venue || "";
  candidateEventTypeInput.value = candidate.eventType || "";
  candidateMemoInput.value = candidate.memo || "";
  excelReviewView.hidden = true;
  excelEditView.hidden = false;
  candidateTitleInput.focus();
}

function closeCandidateEditor() {
  excelState.editingCandidateId = null;
  showExcelReview();
}

function updateImportCandidate(event) {
  event.preventDefault();
  const candidate = excelState.candidates.find((item) => item.candidateId === excelState.editingCandidateId);
  if (!candidate) return;
  Object.assign(candidate, {
    title: candidateTitleInput.value.trim(),
    date: candidateDateInput.value,
    startTime: candidateStartTimeInput.value,
    endTime: candidateEndTimeInput.value,
    venue: candidateVenueInput.value.trim(),
    eventType: candidateEventTypeInput.value.trim(),
    memo: candidateMemoInput.value.trim(),
  });
  refreshDuplicateFlags();
  closeCandidateEditor();
}

function importSelectedExcelEvents() {
  const selected = excelState.candidates.filter((candidate) => candidate.selected);
  const importable = selected.filter((candidate) => !candidate.duplicate);
  const skippedCount = selected.length - importable.length;

  if (!importable.length) {
    showToast(skippedCount ? "중복 일정을 제외하면 가져올 일정이 없습니다." : "가져올 일정을 선택해 주세요.");
    return;
  }

  const importId = createImportId();
  const importedEvents = importable.map(({ candidateId, selected: isSelected, duplicate, ...candidate }) => ({
    ...candidate,
    id: createEventId(),
    sourceType: "excel",
    sourceName: "영화의전당",
    importId,
  }));
  const importRecord = {
    id: importId,
    sourceName: "영화의전당",
    fileName: excelState.fileName,
    importedAt: new Date().toISOString(),
    eventCount: importedEvents.length,
  };
  state.events.push(...importedEvents);
  state.imports.push(importRecord);
  saveEvents();
  saveImports();

  const firstImportedDate = [...importedEvents].sort((first, second) => first.date.localeCompare(second.date))[0].date;
  resetExcelUpload();
  closeExcelModal();
  selectDate(parseDateKey(firstImportedDate));
  showToast(`${importedEvents.length}개의 일정을 가져왔습니다.${skippedCount ? ` 중복 ${skippedCount}개 제외` : ""}`);
}

function resetExcelUpload() {
  excelState.workbook = null;
  excelState.sheetData = {};
  excelState.fileName = "";
  excelState.candidates = [];
  excelState.unparsedCandidates = [];
  excelState.monthFilter = 0;
  excelState.editingCandidateId = null;
  excelError.hidden = true;
  excelSheetList.replaceChildren();
  excelResultView.hidden = true;
  excelReviewView.hidden = true;
  excelEditView.hidden = true;
  excelUploadView.hidden = false;
  document.querySelector("#excel-modal-title").textContent = "Excel 일정 가져오기";
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("visible");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => toast.classList.remove("visible"), 2200);
}

document.querySelector("#prev-month").addEventListener("click", () => changeMonth(-1));
document.querySelector("#next-month").addEventListener("click", () => changeMonth(1));
document.querySelector("[data-add-schedule]").addEventListener("click", () => openEventModal());
document.querySelector("#excel-import-button").addEventListener("click", openExcelModal);
document.querySelector("#modal-close").addEventListener("click", closeEventModal);
document.querySelector("#cancel-modal").addEventListener("click", closeEventModal);
document.querySelector("#excel-modal-close").addEventListener("click", closeExcelModal);
document.querySelector("#excel-cancel").addEventListener("click", closeExcelModal);
document.querySelector("#excel-reselect").addEventListener("click", resetExcelUpload);
document.querySelector("#excel-next").addEventListener("click", analyzeWorkbook);
document.querySelector("#excel-review-back").addEventListener("click", showExcelUploadResult);
importSelectedButton.addEventListener("click", importSelectedExcelEvents);
selectFilteredButton.addEventListener("click", () => setFilteredCandidatesSelected(true));
deselectFilteredButton.addEventListener("click", () => setFilteredCandidatesSelected(false));
document.querySelector("#candidate-edit-cancel").addEventListener("click", closeCandidateEditor);
candidateForm.addEventListener("submit", updateImportCandidate);
deleteEventButton.addEventListener("click", () => deleteEvent(state.editingEventId));
eventForm.addEventListener("submit", handleEventSubmit);
eventModal.addEventListener("click", (event) => { if (event.target === eventModal) closeEventModal(); });
excelModal.addEventListener("click", (event) => { if (event.target === excelModal) closeExcelModal(); });
excelFileInput.addEventListener("change", () => handleExcelFile(excelFileInput.files[0]));
["dragenter", "dragover"].forEach((eventName) => {
  excelDropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    excelDropZone.classList.add("drag-over");
  });
});
["dragleave", "drop"].forEach((eventName) => {
  excelDropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    excelDropZone.classList.remove("drag-over");
  });
});
excelDropZone.addEventListener("drop", (event) => handleExcelFile(event.dataTransfer.files[0]));
document.addEventListener("keydown", (event) => {
  if (event.key !== "Escape") return;
  if (!eventModal.hidden) closeEventModal();
  if (!excelModal.hidden) closeExcelModal();
});

renderCalendar();
renderSelectedDate();
