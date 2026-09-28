(function exposeExcelCompare(global) {
  const COMPARED_FIELDS = ["startTime", "endTime", "venue", "eventType", "memo", "sourceName"];

  function normalizeText(value, lowerCase = false) {
    const normalized = String(value ?? "").trim().replace(/\s+/g, " ");
    return lowerCase ? normalized.toLocaleLowerCase("ko") : normalized;
  }

  function normalizeEvent(event) {
    return {
      title: normalizeText(event.title, true),
      date: normalizeText(event.date),
      startTime: normalizeText(event.startTime),
      endTime: normalizeText(event.endTime),
      venue: normalizeText(event.venue, true),
      eventType: normalizeText(event.eventType, true),
      memo: normalizeText(event.memo),
      sourceName: normalizeText(event.sourceName, true),
    };
  }

  function compareEvents(oldEvents, newEvents) {
    const groups = createIdentityGroups(oldEvents, newEvents);
    const results = [];

    groups.forEach(({ oldItems, newItems }) => {
      compareIdentityGroup(oldItems, newItems, results);
    });

    results.sort((first, second) => {
      const dateOrder = getResultDate(first).localeCompare(getResultDate(second));
      if (dateOrder) return dateOrder;
      return getResultTitle(first).localeCompare(getResultTitle(second), "ko");
    });

    return {
      results,
      summary: results.reduce((counts, result) => {
        counts[result.status] += 1;
        return counts;
      }, { added: 0, changed: 0, removed: 0, unchanged: 0, ambiguous: 0 }),
    };
  }

  function createIdentityGroups(oldEvents, newEvents) {
    const groups = new Map();
    const add = (event, side) => {
      const normalized = normalizeEvent(event);
      const key = `${normalized.date}|${normalized.title}`;
      if (!groups.has(key)) groups.set(key, { oldItems: [], newItems: [] });
      groups.get(key)[side].push({ event, normalized });
    };
    oldEvents.forEach((event) => add(event, "oldItems"));
    newEvents.forEach((event) => add(event, "newItems"));
    return [...groups.values()];
  }

  function compareIdentityGroup(oldItems, newItems, results) {
    const remainingOld = [...oldItems];
    const remainingNew = [...newItems];

    for (let oldIndex = remainingOld.length - 1; oldIndex >= 0; oldIndex -= 1) {
      const newIndex = remainingNew.findIndex((newItem) => (
        isCompatibleIdentity(remainingOld[oldIndex].normalized, newItem.normalized)
          && getChanges(remainingOld[oldIndex], newItem).length === 0
      ));
      if (newIndex === -1) continue;
      results.push(createMatchedResult("unchanged", remainingOld[oldIndex], remainingNew[newIndex], []));
      remainingOld.splice(oldIndex, 1);
      remainingNew.splice(newIndex, 1);
    }

    let foundUniqueMatch = true;
    while (foundUniqueMatch) {
      foundUniqueMatch = false;
      for (let oldIndex = 0; oldIndex < remainingOld.length; oldIndex += 1) {
        const possibleNew = remainingNew
          .map((item, index) => ({ item, index }))
          .filter(({ item }) => isCompatibleIdentity(remainingOld[oldIndex].normalized, item.normalized));
        if (possibleNew.length !== 1) continue;
        const reverseMatches = remainingOld.filter((oldItem) => (
          isCompatibleIdentity(oldItem.normalized, possibleNew[0].item.normalized)
        ));
        if (reverseMatches.length !== 1) continue;

        const newIndex = possibleNew[0].index;
        const changes = getChanges(remainingOld[oldIndex], remainingNew[newIndex]);
        results.push(createMatchedResult(changes.length ? "changed" : "unchanged", remainingOld[oldIndex], remainingNew[newIndex], changes));
        remainingOld.splice(oldIndex, 1);
        remainingNew.splice(newIndex, 1);
        foundUniqueMatch = true;
        break;
      }
    }

    if (remainingOld.length === 1 && remainingNew.length === 1) {
      const changes = getChanges(remainingOld[0], remainingNew[0]);
      results.push(createMatchedResult(changes.length ? "changed" : "unchanged", remainingOld[0], remainingNew[0], changes));
      return;
    }

    if (remainingOld.length && remainingNew.length) {
      results.push({
        status: "ambiguous",
        oldEvents: remainingOld.map((item) => item.event),
        newEvents: remainingNew.map((item) => item.event),
        changes: [],
        wasImported: remainingOld.some((item) => item.event.wasImported === true),
      });
      return;
    }

    remainingOld.forEach((item) => results.push({
      status: "removed",
      oldEvent: item.event,
      newEvent: null,
      changes: [],
      wasImported: item.event.wasImported === true,
    }));
    remainingNew.forEach((item) => results.push({
      status: "added",
      oldEvent: null,
      newEvent: item.event,
      changes: [],
      wasImported: false,
    }));
  }

  function isCompatibleIdentity(oldEvent, newEvent) {
    const venueMatches = !oldEvent.venue || !newEvent.venue || oldEvent.venue === newEvent.venue;
    const sourceMatches = !oldEvent.sourceName || !newEvent.sourceName || oldEvent.sourceName === newEvent.sourceName;
    return oldEvent.date === newEvent.date && oldEvent.title === newEvent.title && venueMatches && sourceMatches;
  }

  function getChanges(oldItem, newItem) {
    return COMPARED_FIELDS.filter((field) => oldItem.normalized[field] !== newItem.normalized[field]).map((field) => ({
      field,
      oldValue: oldItem.event[field] ?? "",
      newValue: newItem.event[field] ?? "",
    }));
  }

  function createMatchedResult(status, oldItem, newItem, changes) {
    return {
      status,
      oldEvent: oldItem.event,
      newEvent: newItem.event,
      changes,
      wasImported: oldItem.event.wasImported === true,
    };
  }

  function getResultDate(result) {
    return result.newEvent?.date || result.oldEvent?.date || result.newEvents?.[0]?.date || result.oldEvents?.[0]?.date || "";
  }

  function getResultTitle(result) {
    return result.newEvent?.title || result.oldEvent?.title || result.newEvents?.[0]?.title || result.oldEvents?.[0]?.title || "";
  }

  global.ExcelCompare = Object.freeze({ normalizeText, normalizeEvent, compareEvents });
}(window));
