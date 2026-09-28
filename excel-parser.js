(function exposeExcelParser(global) {
  const EVENT_TYPE_MAP = Object.freeze({
    기: "기획",
    기획: "기획",
    대: "대관",
    대관: "대관",
    영: "영화상영",
    영화: "영화상영",
    영화상영: "영화상영",
  });

  const VENUE_MAP = Object.freeze({
    하: "하늘연극장",
    하늘: "하늘연극장",
    하늘연극장: "하늘연극장",
    야: "야외극장",
    야외: "야외극장",
    야외극장: "야외극장",
    광: "두레라움광장",
    광장: "두레라움광장",
    두레라움광장: "두레라움광장",
    리: "리허설룸",
    리허설: "리허설룸",
    리허설룸: "리허설룸",
  });

  const IGNORED_TEXT = new Set([
    "일", "월", "화", "수", "목", "금", "토", "일요일", "월요일", "화요일",
    "수요일", "목요일", "금요일", "토요일", "일정", "비고", "장소", "시간", "구분",
  ]);

  function parseWorkbook(workbook, fileName = "") {
    const year = detectWorkbookYear(workbook, fileName);
    const events = [];
    const unparsedCandidates = [];
    const analyzedSheets = [];

    workbook.SheetNames.forEach((sheetName) => {
      const month = getMonthFromSheetName(sheetName);
      if (!month) return;
      analyzedSheets.push(sheetName);
      const result = parseMonthlySheet(workbook.Sheets[sheetName], { year, month, sheetName });
      events.push(...result.events);
      unparsedCandidates.push(...result.unparsedCandidates);
    });

    return { year, events, unparsedCandidates, analyzedSheets };
  }

  function parseMonthlySheet(worksheet, context) {
    const rows = global.XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: "" });
    const dateRegions = extractCalendarDates(rows, worksheet["!merges"] || [], context);
    const events = [];
    const unparsedCandidates = [];

    dateRegions.forEach((region) => {
      const result = extractEventsFromDate(rows, region, context);
      events.push(...result.events);
      unparsedCandidates.push(...result.unparsedCandidates);
    });

    return { events, unparsedCandidates, dateRegions };
  }

  function extractCalendarDates(rows, merges, { year, month }) {
    const markerRows = [];

    rows.forEach((row, rowIndex) => {
      const candidates = row.map((value, columnIndex) => ({
        day: parseDayNumber(value),
        columnIndex,
      })).filter((item) => item.day && item.day <= daysInMonth(year, month));

      if (looksLikeCalendarWeek(candidates)) markerRows.push({ rowIndex, candidates });
    });

    if (!markerRows.length) return [];
    const rowGaps = markerRows.slice(1).map((row, index) => row.rowIndex - markerRows[index].rowIndex);
    const typicalRowSpan = median(rowGaps) || 6;
    const regions = [];

    markerRows.forEach((markerRow, markerIndex) => {
      const columnGaps = markerRow.candidates.slice(1).map((item, index) => (
        item.columnIndex - markerRow.candidates[index].columnIndex
      ));
      const typicalColumnSpan = median(columnGaps) || 1;
      const nextMarkerRow = markerRows[markerIndex + 1]?.rowIndex;
      const endRow = Math.max(
        markerRow.rowIndex,
        Math.min(rows.length - 1, nextMarkerRow ? nextMarkerRow - 1 : markerRow.rowIndex + typicalRowSpan - 1),
      );

      markerRow.candidates.forEach((candidate, candidateIndex) => {
        const merge = merges.find((range) => range.s.r <= markerRow.rowIndex
          && range.e.r >= markerRow.rowIndex
          && range.s.c <= candidate.columnIndex
          && range.e.c >= candidate.columnIndex);
        const nextColumn = markerRow.candidates[candidateIndex + 1]?.columnIndex;
        const startColumn = merge?.s.c ?? candidate.columnIndex;
        const endColumn = merge?.e.c ?? ((nextColumn ?? candidate.columnIndex + typicalColumnSpan) - 1);
        regions.push({
          date: formatDate(year, month, candidate.day),
          day: candidate.day,
          startRow: markerRow.rowIndex + 1,
          endRow,
          startColumn,
          endColumn: Math.max(startColumn, endColumn),
        });
      });
    });

    return uniqueBy(regions, (region) => region.date);
  }

  function extractEventsFromDate(rows, region, context) {
    const events = [];
    const unparsedCandidates = [];
    let pendingMetadata = null;

    for (let rowIndex = region.startRow; rowIndex <= region.endRow; rowIndex += 1) {
      const values = (rows[rowIndex] || [])
        .slice(region.startColumn, region.endColumn + 1)
        .map(cleanText)
        .filter(Boolean);
      const rawText = [...new Set(values)].join("\n");
      if (!rawText || shouldIgnore(rawText)) continue;

      const parsed = parseCandidateText(rawText);
      if (!parsed.title) {
        if (parsed.startTime || parsed.venue || parsed.eventType) {
          pendingMetadata = { ...pendingMetadata, ...withoutEmpty(parsed) };
        } else if (parsed.memo && events.length) {
          events[events.length - 1].memo = [events[events.length - 1].memo, parsed.memo].filter(Boolean).join(" · ");
          events[events.length - 1].rawText = [events[events.length - 1].rawText, rawText].filter(Boolean).join("\n");
        } else {
          unparsedCandidates.push({ date: region.date, sheetName: context.sheetName, rawText, reason: "제목 식별 불가" });
        }
        continue;
      }

      events.push({
        title: parsed.title,
        date: region.date,
        startTime: parsed.startTime || pendingMetadata?.startTime || "",
        endTime: parsed.endTime || pendingMetadata?.endTime || "",
        memo: parsed.memo || pendingMetadata?.memo || "",
        sourceType: "excel",
        sourceName: "영화의전당",
        eventType: parsed.eventType || pendingMetadata?.eventType || "",
        venue: parsed.venue || pendingMetadata?.venue || "",
        rawText: [pendingMetadata?.rawText, rawText].filter(Boolean).join("\n"),
      });
      pendingMetadata = null;
    }

    if (pendingMetadata) {
      unparsedCandidates.push({
        date: region.date,
        sheetName: context.sheetName,
        rawText: pendingMetadata.rawText || "",
        reason: "일정 정보만 있고 제목 없음",
      });
    }

    return { events, unparsedCandidates };
  }

  function parseCandidateText(rawText) {
    const normalizedTime = normalizeTime(rawText);
    const venue = normalizeVenue(rawText);
    const eventType = normalizeEventType(rawText);
    const operationalNotes = [...rawText.matchAll(/(?:^|[\s\n])(종일|셋업|철수)(?=$|[\s\n])/g)]
      .map((match) => match[1]);
    let title = rawText.replace(/\r?\n/g, " ");
    title = removeTimeExpressions(title);
    title = title.replace(/(?:^|\s)(?:기|대|영)\s*-\s*(?:하|야|광|리)(?=$|\s)/g, " ");
    title = title.replace(/(?:^|\s)(?:종일|셋업|철수)(?=$|\s)/g, " ");
    title = removeMappedTokens(title, VENUE_MAP);
    title = removeMappedTokens(title, EVENT_TYPE_MAP);
    title = title.replace(/[|·•]+/g, " ").replace(/\s{2,}/g, " ").trim();

    if (/^(종일|셋업|철수|setup)$/i.test(title)) title = "";
    return { ...normalizedTime, venue, eventType, title, memo: operationalNotes.join(" · "), rawText };
  }

  function normalizeTime(text) {
    const values = [];
    const timePattern = /(?<!\d)([01]?\d|2[0-3])(?:\s*[:시]\s*([0-5]?\d)?\s*분?)(?!\d)/g;
    let match;
    while ((match = timePattern.exec(text)) !== null) {
      const hours = String(Number(match[1])).padStart(2, "0");
      const minutes = String(Number(match[2] || 0)).padStart(2, "0");
      values.push(`${hours}:${minutes}`);
    }
    const serialMatches = text.match(/(?<![\d.])0\.\d+(?![\d.])/g) || [];
    serialMatches.forEach((serial) => {
      const totalMinutes = Math.round(Number(serial) * 24 * 60) % (24 * 60);
      const hours = String(Math.floor(totalMinutes / 60)).padStart(2, "0");
      const minutes = String(totalMinutes % 60).padStart(2, "0");
      values.push(`${hours}:${minutes}`);
    });
    const uniqueTimes = [...new Set(values)];
    const isExplicitRange = /(?:~|～|-|–|—)\s*(?:[01]?\d|2[0-3])(?:\s*[:시])/i.test(text);
    return {
      startTime: uniqueTimes[0] || "",
      endTime: isExplicitRange ? (uniqueTimes[1] || "") : "",
    };
  }

  function normalizeVenue(text) {
    const pair = text.match(/(?:^|[\s\n])(?:기|대|영)\s*-\s*(하|야|광|리)(?=$|[\s\n])/);
    if (pair) return VENUE_MAP[pair[1]] || "";
    return findMappedValue(text, VENUE_MAP);
  }

  function normalizeEventType(text) {
    const pair = text.match(/(?:^|[\s\n])(기|대|영)\s*-\s*(?:하|야|광|리)(?=$|[\s\n])/);
    if (pair) return EVENT_TYPE_MAP[pair[1]] || "";
    return findMappedValue(text, EVENT_TYPE_MAP);
  }

  function getMonthFromSheetName(sheetName) {
    const match = String(sheetName).trim().match(/^(1[0-2]|[1-9])\s*월(?:\s|$|[-_])/);
    return match ? Number(match[1]) : null;
  }

  function detectWorkbookYear(workbook, fileName) {
    const fileMatch = String(fileName).match(/(20\d{2})\s*년?/);
    if (fileMatch) return Number(fileMatch[1]);
    for (const sheetName of workbook.SheetNames) {
      const rows = global.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: "" });
      const sample = rows.slice(0, 15).flat().join(" ");
      const match = sample.match(/(20\d{2})\s*년/);
      if (match) return Number(match[1]);
    }
    return new Date().getFullYear();
  }

  function looksLikeCalendarWeek(candidates) {
    if (candidates.length < 2 || candidates.length > 7) return false;
    let sequentialPairs = 0;
    for (let index = 1; index < candidates.length; index += 1) {
      if (candidates[index].day === candidates[index - 1].day + 1) sequentialPairs += 1;
    }
    return sequentialPairs >= Math.max(1, candidates.length - 2);
  }

  function parseDayNumber(value) {
    const match = cleanText(value).match(/^([1-9]|[12]\d|3[01])\s*일?$/);
    return match ? Number(match[1]) : null;
  }

  function shouldIgnore(text) {
    const compact = cleanText(text);
    return IGNORED_TEXT.has(compact) || /^\d+$/.test(compact) || /^20\d{2}년/.test(compact);
  }

  function findMappedValue(text, mapping) {
    const tokens = text.split(/[\s\n()[\]{}\/|·,]+/).filter(Boolean);
    for (const token of tokens) {
      if (mapping[token]) return mapping[token];
    }
    return Object.values(mapping).find((value) => text.includes(value)) || "";
  }

  function removeMappedTokens(text, mapping) {
    return Object.keys(mapping).sort((a, b) => b.length - a.length).reduce((result, token) => {
      const escaped = token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      return result.replace(new RegExp(`(^|[\\s()[\\]{}\\/|·,])${escaped}(?=$|[\\s()[\\]{}\\/|·,])`, "g"), " ");
    }, text);
  }

  function removeTimeExpressions(text) {
    return text
      .replace(/(?<![\d.])0\.\d+(?![\d.])/g, " ")
      .replace(/(?<!\d)([01]?\d|2[0-3])\s*:\s*[0-5]\d(?!\d)/g, " ")
      .replace(/(?<!\d)([01]?\d|2[0-3])\s*시(?:\s*[0-5]?\d\s*분?)?(?!\d)/g, " ")
      .replace(/\s*(?:~|～|-|–|—|\/)\s*/g, " ");
  }

  function cleanText(value) {
    return String(value ?? "").replace(/\u00a0/g, " ").replace(/[ \t]+/g, " ").trim();
  }

  function formatDate(year, month, day) {
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  function daysInMonth(year, month) {
    return new Date(year, month, 0).getDate();
  }

  function median(values) {
    if (!values.length) return 0;
    const sorted = [...values].sort((a, b) => a - b);
    return sorted[Math.floor(sorted.length / 2)];
  }

  function uniqueBy(items, getKey) {
    return [...new Map(items.map((item) => [getKey(item), item])).values()];
  }

  function withoutEmpty(object) {
    return Object.fromEntries(Object.entries(object).filter(([, value]) => value));
  }

  global.ExcelParser = Object.freeze({
    EVENT_TYPE_MAP,
    VENUE_MAP,
    parseWorkbook,
    parseMonthlySheet,
    extractCalendarDates,
    extractEventsFromDate,
    normalizeTime,
    normalizeVenue,
    normalizeEventType,
  });
}(window));
