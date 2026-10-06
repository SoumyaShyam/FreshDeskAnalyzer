(() => {
  "use strict";

  const expectedPrefix = "FreshDesk_CRM_";
  const expectedDate = /FreshDesk_CRM_(\d{2})-(\d{2})-(\d{4})\.xlsx?$/i;
  const columns = [
    ["Ticket ID", ["ticketid", "ticket id", "id"]],
    ["Status", ["status", "ticket status"]],
    ["Subject", ["subject"]],
    ["Created time", ["createdtime", "created time", "createdat", "created date"]],
    ["ETA", ["eta", "estimatedtimeofarrival"]],
    ["Resolved time", ["resolvedtime", "resolved time", "resolvedat"]],
    ["Closed time", ["closedtime", "closed time", "closedat"]],
    ["Assignee", ["assignee", "assignedto", "assigned to"]],
    ["Full name", ["fullname", "full name", "requestername"]],
    ["Contact ID", ["contactid", "contact id", "email", "requesteremail"]],
    ["Group", ["group", "groupname", "group name"]]
  ];
  const reportColumns = ["Sl. No.", "Ticket ID", "Status", "Subject", "Created date-time", "Age (days)", "ETA", "Resolved date", "Closed date", "Assigned to", "Raised by", "Email"];
  const fieldMap = {
    "Ticket ID": "Ticket ID", Status: "Status", "Subject": "Subject", "Created time": "Created date-time", ETA: "ETA",
    "Resolved time": "Resolved date", "Closed time": "Closed date", "Assignee": "Assigned to",
    "Full name": "Raised by", "Contact ID": "Email"
  };

  let tickets = [];
  let groupNames = [];
  let activeGroup = "";
  let newestDumpDate = null;

  const $status = $("#status-line");
  const $statusText = $("#status-text");

  function setStatus(message, state = "") {
    $status.removeClass("is-loading is-success");
    if (state) $status.addClass(`is-${state}`);
    $statusText.text(message);
  }

  function showError(message) {
    $("#error-banner").text(message).prop("hidden", false);
  }

  function clearError() {
    $("#error-banner").text("").prop("hidden", true);
  }

  function parseDumpDate(fileName) {
    const match = fileName.match(expectedDate);
    if (!match) return null;
    const [, dayText, monthText, yearText] = match;
    const date = new Date(Number(yearText), Number(monthText) - 1, Number(dayText));
    if (date.getFullYear() !== Number(yearText) || date.getMonth() !== Number(monthText) - 1 || date.getDate() !== Number(dayText)) return null;
    return date;
  }

  function dateKey(date) {
    return date.getFullYear() * 10000 + (date.getMonth() + 1) * 100 + date.getDate();
  }

  function normalizeHeader(value) {
    return String(value ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
  }

  function excelDate(value) {
    if (value === null || value === undefined || value === "") return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === "number" && window.XLSX) {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, parsed.S || 0);
    }
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function formatDateTime(value) {
    const date = excelDate(value);
    if (!date) return value === null || value === undefined ? "" : String(value).trim();
    return new Intl.DateTimeFormat(undefined, { year: "numeric", month: "short", day: "2-digit", hour: "2-digit", minute: "2-digit" }).format(date);
  }

  function formatDate(value) {
    const date = excelDate(value);
    if (!date) return value === null || value === undefined ? "" : String(value).trim();
    return new Intl.DateTimeFormat(undefined, { year: "numeric", month: "short", day: "2-digit" }).format(date);
  }

  function ageInDays(value, asOf) {
    const created = excelDate(value);
    if (!created || !asOf) return "";
    const createdDay = Date.UTC(created.getFullYear(), created.getMonth(), created.getDate());
    const snapshotDay = Date.UTC(asOf.getFullYear(), asOf.getMonth(), asOf.getDate());
    return Math.max(0, Math.floor((snapshotDay - createdDay) / 86400000));
  }

  function findColumn(headers, aliases) {
    const normalized = headers.map(normalizeHeader);
    for (const alias of aliases) {
      const index = normalized.indexOf(normalizeHeader(alias));
      if (index !== -1) return headers[index];
    }
    return null;
  }

  function sheetToRecords(sheet) {
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
    const headerIndex = matrix.findIndex(row => row.some(cell => normalizeHeader(cell) === "ticketid"));
    if (headerIndex < 0) throw new Error('Could not find a "Ticket ID" header in the selected sheet.');
    const headers = matrix[headerIndex].map(cell => String(cell ?? "").trim());
    const resolved = new Map();
    for (const [field, aliases] of columns) {
      const header = findColumn(headers, aliases);
      if (header) resolved.set(field, headers.indexOf(header));
    }
    const required = ["Ticket ID", "Status", "Subject", "Created time", "Group"];
    const missing = required.filter(field => !resolved.has(field));
    if (missing.length) throw new Error(`Missing required column${missing.length > 1 ? "s" : ""}: ${missing.join(", ")}.`);

    return matrix.slice(headerIndex + 1).filter(row => row.some(value => value !== "")).map(row => {
      const record = {};
      for (const [field] of columns) record[field] = resolved.has(field) ? row[resolved.get(field)] : "";
      return record;
    });
  }

  async function readWorkbook(file) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
    if (!workbook.SheetNames.length) throw new Error("The workbook has no sheets.");
    const sheetName = workbook.SheetNames.find(name => name.trim().toLowerCase() === "main sheet") || workbook.SheetNames[0];
    return { sheetName, records: sheetToRecords(workbook.Sheets[sheetName]) };
  }

  function displayValue(ticket, field) {
    const value = ticket[field];
    if (field === "Created time") return formatDateTime(value);
    if (field === "ETA" && typeof value === "number") return formatDate(value);
    if (field === "Resolved time" || field === "Closed time") return formatDate(value);
    if (value instanceof Date) return formatDateTime(value);
    return value === null || value === undefined ? "" : String(value).trim();
  }

  function buildLatestTickets(fileData) {
    const latestById = new Map();
    for (const item of fileData) {
      const dumpDate = item.dumpDate;
      for (const record of item.records) {
        const id = String(record["Ticket ID"] ?? "").trim();
        if (!id) continue;
        const existing = latestById.get(id);
        if (!existing || dateKey(dumpDate) >= dateKey(existing.dumpDate)) {
          const group = String(record.Group ?? "").trim() || "Unassigned group";
          latestById.set(id, { ...record, Group: group, dumpDate });
        }
      }
    }
    tickets = Array.from(latestById.values()).map(ticket => ({
      ...ticket,
      age: String(ticket.Status ?? "").trim().toLowerCase() === "closed" ? "" : ageInDays(ticket["Created time"], newestDumpDate)
    }));
    groupNames = [...new Set(tickets.map(ticket => ticket.Group))].sort((a, b) => a.localeCompare(b));
    activeGroup = groupNames[0] || "";
  }

  function renderTabs() {
    const counts = tickets.reduce((result, ticket) => {
      result[ticket.Group] = (result[ticket.Group] || 0) + 1;
      return result;
    }, {});
    const $tabs = $("#group-tabs").empty();
    groupNames.forEach(group => {
      const $tab = $("<button>", { type: "button", role: "tab", class: "group-tab", "aria-selected": String(group === activeGroup), title: group });
      $tab.append(document.createTextNode(group));
      $tab.append($("<span>", { class: "tab-count" }).text(counts[group]));
      $tab.on("click", () => {
        activeGroup = group;
        renderTabs();
        renderRows();
      });
      $tabs.append($tab);
    });
  }

  function visibleTickets() {
    const query = $("#search-input").val().trim().toLowerCase();
    return tickets.filter(ticket => ticket.Group === activeGroup && (!query || [
      ticket["Ticket ID"], ticket.Status, ticket.Subject, ticket.Assignee, ticket["Full name"], ticket["Contact ID"]
    ].some(value => String(value ?? "").toLowerCase().includes(query))))
      .sort((left, right) => {
        if (left.age === "" && right.age !== "") return 1;
        if (right.age === "" && left.age !== "") return -1;
        if (left.age !== "" && right.age !== "" && left.age !== right.age) return left.age - right.age;
        return String(left["Ticket ID"]).localeCompare(String(right["Ticket ID"]), undefined, { numeric: true });
      });
  }

  function reportValues(ticket, index) {
    return reportColumns.map(label => {
      if (label === "Sl. No.") return String(index + 1);
      if (label === "Age (days)") return ticket.age === "" ? "" : String(ticket.age);
      const field = Object.keys(fieldMap).find(key => fieldMap[key] === label);
      return field ? displayValue(ticket, field) : "";
    });
  }

  function renderRows() {
    const rows = visibleTickets();
    const $body = $("#ticket-rows").empty();
    $("#result-count").text(`${rows.length} ${rows.length === 1 ? "ticket" : "tickets"}`);
    $("#active-group-label").text(activeGroup);
    $("#empty-state").prop("hidden", rows.length > 0);
    $(".table-wrap table").prop("hidden", rows.length === 0);

    rows.forEach((ticket, index) => {
      const values = reportValues(ticket, index);
      const $row = $("<tr>");
      values.forEach((value, index) => {
        const $cell = $("<td>").text(value).attr("title", value);
        if (index === 0) $cell.addClass("sl-number");
        if (index === 1) $cell.addClass("ticket-id");
        if (index === 2) $cell.addClass("status-cell");
        if (index === 3) $cell.addClass("subject-cell");
        if (reportColumns[index] === "Age (days)") $cell.addClass("age-cell");
        $row.append($cell);
      });
      $body.append($row);
    });
  }

  function setSummary(fileCount) {
    $("#metric-tickets").text(tickets.length.toLocaleString());
    $("#metric-groups").text(groupNames.length.toLocaleString());
    $("#metric-dumps").text(fileCount.toLocaleString());
    $("#metric-as-of").text(newestDumpDate ? formatDate(newestDumpDate) : "—");
    $("#snapshot-date-value").text(newestDumpDate ? formatDate(newestDumpDate) : "—");
    $("#snapshot-date").prop("hidden", !newestDumpDate);
  }

  async function analyzeFiles(files) {
    clearError();
    const candidates = Array.from(files).filter(file => /\.xlsx?$/i.test(file.name) && file.name.toLowerCase().startsWith(expectedPrefix.toLowerCase()));
    const rejected = Array.from(files).filter(file => /\.xlsx?$/i.test(file.name) && !file.name.toLowerCase().startsWith(expectedPrefix.toLowerCase()));
    if (!candidates.length) {
      showError(`No Excel files named ${expectedPrefix}<dd-mm-yyyy>.xlsx were found in the selection.`);
      setStatus("No matching Freshdesk exports found.");
      return;
    }

    const dated = candidates.map(file => ({ file, dumpDate: parseDumpDate(file.name) }));
    const invalid = dated.filter(item => !item.dumpDate);
    const valid = dated.filter(item => item.dumpDate).sort((a, b) => dateKey(a.dumpDate) - dateKey(b.dumpDate) || a.file.name.localeCompare(b.file.name));
    if (!valid.length) {
      showError("Matching filenames must use a valid date in DD-MM-YYYY format.");
      setStatus("No valid dated exports to analyze.");
      return;
    }

    $status.addClass("is-loading");
    setStatus(`Reading ${valid.length} Freshdesk ${valid.length === 1 ? "dump" : "dumps"}…`, "loading");
    const loaded = [];
    const failures = [];
    for (const item of valid) {
      try {
        const result = await readWorkbook(item.file);
        loaded.push({ ...item, ...result });
      } catch (error) {
        failures.push(`${item.file.name}: ${error.message}`);
      }
    }

    if (!loaded.length) {
      $("#report").prop("hidden", true);
      showError(failures.join(" ") || "No readable workbooks were found.");
      setStatus("Could not read any selected exports.");
      return;
    }

    newestDumpDate = loaded.reduce((latest, item) => !latest || dateKey(item.dumpDate) > dateKey(latest) ? item.dumpDate : latest, null);
    buildLatestTickets(loaded);
    setSummary(loaded.length);
    renderTabs();
    renderRows();
    $("#report").prop("hidden", false);
    const notes = [];
    if (invalid.length) notes.push(`${invalid.length} file${invalid.length === 1 ? "" : "s"} skipped for invalid date names`);
    if (rejected.length) notes.push(`${rejected.length} non-matching Excel file${rejected.length === 1 ? "" : "s"} ignored`);
    if (failures.length) notes.push(`${failures.length} unreadable dump${failures.length === 1 ? "" : "s"}`);
    setStatus(`${tickets.length.toLocaleString()} latest tickets from ${loaded.length} dump${loaded.length === 1 ? "" : "s"}${notes.length ? ` · ${notes.join(" · ")}` : ""}.`, "success");
    if (failures.length) showError(failures.join(" "));
  }

  function csvValue(value) {
    return `"${String(value ?? "").replace(/"/g, '""')}"`;
  }

  function exportCurrentGroup() {
    const rows = visibleTickets();
    const output = [reportColumns.map(csvValue).join(",")];
    rows.forEach((ticket, index) => {
      output.push(reportValues(ticket, index).map(csvValue).join(","));
    });
    const blob = new Blob(["\ufeff", output.join("\r\n")], { type: "text/csv;charset=utf-8" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `FreshDesk_${activeGroup.replace(/[^a-z0-9_-]+/gi, "_")}_${newestDumpDate ? `${String(newestDumpDate.getDate()).padStart(2, "0")}-${String(newestDumpDate.getMonth() + 1).padStart(2, "0")}-${newestDumpDate.getFullYear()}` : "report"}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  $("#choose-folder").on("click", async () => {
    clearError();
    if (!window.showDirectoryPicker) {
      showError("Folder selection is not supported by this browser. Use Select files to choose the Excel exports instead.");
      return;
    }
    try {
      const directory = await window.showDirectoryPicker();
      const files = [];
      for await (const entry of directory.values()) {
        if (entry.kind === "file" && /\.xlsx?$/i.test(entry.name)) files.push(await entry.getFile());
      }
      await analyzeFiles(files);
    } catch (error) {
      if (error.name !== "AbortError") showError(`Could not read the selected folder: ${error.message}`);
    }
  });

  async function loadDefaultFolder() {
    clearError();
    setStatus("Loading exports from the configured folder…", "loading");
    try {
      const response = await fetch("/api/default-folder", { cache: "no-store" });
      const result = await response.json();
      $("#default-folder-path").text(result.folder || "Configured folder").attr("title", result.folder || "");
      if (!response.ok) throw new Error(result.error || "Could not read the configured folder.");
      if (!result.files.length) {
        setStatus("No Freshdesk Excel exports found in the default folder.");
        return;
      }
      const files = await Promise.all(result.files.map(async name => {
        const fileResponse = await fetch(`/api/files/${encodeURIComponent(name)}`, { cache: "no-store" });
        if (!fileResponse.ok) throw new Error(`Could not load ${name}.`);
        return new File([await fileResponse.arrayBuffer()], name);
      }));
      await analyzeFiles(files);
    } catch (error) {
      $("#default-folder-path").text("Start with node server.js to load the configured folder.");
      showError(error.message);
      setStatus("Default-folder loading is unavailable; choose files or a folder manually.");
    }
  }

  $("#file-input").on("change", function () {
    if (this.files?.length) analyzeFiles(this.files);
    this.value = "";
  });
  $("#search-input").on("input", renderRows);
  $("#export-button").on("click", exportCurrentGroup);
  if (location.protocol === "http:" || location.protocol === "https:") {
    loadDefaultFolder();
  } else if (location.protocol === "file:") {
    showError("Browsers block automatic access to the configured Windows folder when index.html is opened directly. Double-click Start Freshdesk Analyzer.bat for automatic analysis, or choose the folder/files manually.");
    setStatus("Direct file opening cannot auto-load the configured folder.");
  }
})();