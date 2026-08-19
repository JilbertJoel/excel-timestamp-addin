/* global Office, Excel */

const CONFIG = Object.freeze({
  worksheetName: "Tracker",
  triggerColumnIndex: 3,   // Column D, zero-based.
  timestampColumnIndex: 4, // Column E, zero-based.
  firstDataRowIndex: 1,    // Row 2; row 1 contains headers.
  timeZone: "America/Santo_Domingo"
});

let handlerRegistered = false;
let changeQueue = Promise.resolve();

Office.onReady(async (info) => {
  wireButtons();

  if (info.host !== Office.HostType.Excel) {
    setStatus("error", "Excel required", "Open this add-in inside the Excel workbook.");
    return;
  }

  await startTimestampService(true);
});

function wireButtons() {
  document.getElementById("retry").addEventListener("click", () => startTimestampService(false));
  document.getElementById("enable-startup").addEventListener("click", enableAutomaticStartup);
  document.getElementById("disable-startup").addEventListener("click", disableAutomaticStartup);
}

async function startTimestampService(enableStartup) {
  setButtonsDisabled(true);
  setStatus("starting", "Starting…", "Connecting to the Tracker worksheet.");

  try {
    await registerChangeHandler();

    let startupMessage = "Automatic startup is already configured.";
    if (enableStartup) {
      startupMessage = await tryEnableAutomaticStartup();
    }

    setStatus(
      "active",
      "Active",
      `Future entries in column D will be timestamped in column E. ${startupMessage}`
    );
  } catch (error) {
    console.error(error);
    setStatus(
      "error",
      "Could not start",
      friendlyError(error)
    );
  } finally {
    setButtonsDisabled(false);
  }
}

async function registerChangeHandler() {
  if (handlerRegistered) return;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(CONFIG.worksheetName);
    sheet.load("name,isNullObject");
    await context.sync();

    if (sheet.isNullObject) {
      throw new Error(`Worksheet \"${CONFIG.worksheetName}\" was not found.`);
    }

    sheet.onChanged.add(queueChange);
    await context.sync();
  });

  handlerRegistered = true;
}

function queueChange(event) {
  // Each editor's add-in handles that editor's local change. Ignoring remote
  // events prevents two coauthors from trying to stamp the same edit.
  if (String(event.source).toLowerCase() === "remote") return;

  changeQueue = changeQueue
    .then(() => processChange(event))
    .catch((error) => {
      console.error("Timestamp event failed", error);
      setStatus("warning", "Active with a warning", friendlyError(error));
    });
}

async function processChange(event) {
  const addresses = localAddresses(event.address);
  if (addresses.length === 0) return;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(CONFIG.worksheetName);
    const intersections = addresses.map((address) => {
      const changedRange = sheet.getRange(address);
      const intersection = changedRange.getIntersectionOrNullObject("D:D");
      intersection.load("isNullObject,rowIndex,rowCount");
      return intersection;
    });

    await context.sync();

    const segments = mergeRowSegments(
      intersections
        .filter((range) => !range.isNullObject)
        .map((range) => ({
          start: Math.max(range.rowIndex, CONFIG.firstDataRowIndex),
          end: range.rowIndex + range.rowCount
        }))
        .filter((segment) => segment.end > segment.start)
    );

    if (segments.length === 0) return;

    const loadedSegments = segments.map((segment) => {
      const rowCount = segment.end - segment.start;
      const triggerRange = sheet.getRangeByIndexes(segment.start, CONFIG.triggerColumnIndex, rowCount, 1);
      const timestampRange = sheet.getRangeByIndexes(segment.start, CONFIG.timestampColumnIndex, rowCount, 1);
      triggerRange.load("values");
      timestampRange.load("values,formulas");
      return { ...segment, triggerRange, timestampRange };
    });

    await context.sync();

    // Excel stores real dates as numeric OLE Automation date values.
    // Build the serial from Dominican Republic wall-clock time so the cell
    // remains usable in sorting, filtering, formulas, and date arithmetic.
    const timestampSerial = dominicanExcelSerial(new Date());

    for (const segment of loadedSegments) {
      const triggerValues = segment.triggerRange.values;
      const timestampValues = segment.timestampRange.values;
      const timestampFormulas = segment.timestampRange.formulas;

      for (let offset = 0; offset < triggerValues.length; offset += 1) {
        const triggerHasValue = !isBlank(triggerValues[offset][0]);
        const timestampHasValue = !isBlank(timestampValues[offset][0]);
        const timestampHasFormula = typeof timestampFormulas[offset][0] === "string"
          && timestampFormulas[offset][0].startsWith("=");

        if (triggerHasValue && !timestampHasValue && !timestampHasFormula) {
          const cell = sheet.getCell(segment.start + offset, CONFIG.timestampColumnIndex);
          cell.values = [[timestampSerial]];
        }
      }
    }

    await context.sync();
  });
}

function localAddresses(eventAddress) {
  if (!eventAddress || typeof eventAddress !== "string") return [];

  return eventAddress
    .split(",")
    .map((address) => {
      const bang = address.lastIndexOf("!");
      return (bang >= 0 ? address.slice(bang + 1) : address).replaceAll("$", "").trim();
    })
    .filter(Boolean);
}

function mergeRowSegments(segments) {
  const ordered = segments.sort((a, b) => a.start - b.start);
  const merged = [];

  for (const segment of ordered) {
    const previous = merged[merged.length - 1];
    if (!previous || segment.start > previous.end) {
      merged.push({ ...segment });
    } else {
      previous.end = Math.max(previous.end, segment.end);
    }
  }

  return merged;
}

function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function dominicanExcelSerial(date) {
  const formatter = new Intl.DateTimeFormat("en-CA", {
    timeZone: CONFIG.timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hourCycle: "h23"
  });

  const parts = Object.fromEntries(
    formatter.formatToParts(date)
      .filter((part) => part.type !== "literal")
      .map((part) => [part.type, part.value])
  );

  const dominicanWallClockAsUtc = Date.UTC(
    Number(parts.year),
    Number(parts.month) - 1,
    Number(parts.day),
    Number(parts.hour),
    Number(parts.minute),
    Number(parts.second)
  );

  // Excel's standard 1900 date system uses 1899-12-30 as serial zero.
  return dominicanWallClockAsUtc / 86400000 + 25569;
}

async function tryEnableAutomaticStartup() {
  try {
    if (!Office.addin || !Office.addin.setStartupBehavior) {
      return "Keep the task pane open in this session.";
    }

    await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    return "Automatic startup is enabled for this workbook.";
  } catch (error) {
    console.warn("Automatic startup could not be enabled", error);
    return "Keep the task pane open in this session; automatic startup could not be confirmed.";
  }
}

async function enableAutomaticStartup() {
  setButtonsDisabled(true);
  const message = await tryEnableAutomaticStartup();
  setStatus("active", "Active", message);
  setButtonsDisabled(false);
}

async function disableAutomaticStartup() {
  setButtonsDisabled(true);

  try {
    await Office.addin.setStartupBehavior(Office.StartupBehavior.none);
    setStatus(
      "warning",
      "Automatic startup disabled",
      "Timestamps remain active for this session. Open the add-in manually next time."
    );
  } catch (error) {
    setStatus("error", "Could not change startup", friendlyError(error));
  } finally {
    setButtonsDisabled(false);
  }
}

function setStatus(type, title, message) {
  const card = document.getElementById("status-card");
  card.className = `status ${type}`;
  document.getElementById("status-title").textContent = title;
  document.getElementById("status-message").textContent = message;
}

function setButtonsDisabled(disabled) {
  document.querySelectorAll("button").forEach((button) => {
    button.disabled = disabled;
  });
}

function friendlyError(error) {
  const message = error && error.message ? error.message : String(error);
  if (message.includes("ItemNotFound") || message.includes("was not found")) {
    return `The worksheet \"${CONFIG.worksheetName}\" was not found. Confirm the tab name and retry.`;
  }
  return message;
}
