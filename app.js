(function () {
  "use strict";

  const STORAGE_KEY = "belomor_static_projects_v1";
  const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  const COLORS = {
    headerBg: "D9E1F2",
    leftHeaderBg: "F2F2F2",
    weekend: "E7E6E6",
    holiday: "F3CFC9",
    taskGreen: "C6EFCE",
    taskYellow: "FFF2CC",
    taskOrange: "FCE4D6",
    week12: "FFFF99",
    week34: "EFC057",
    week5: "FFA500",
    belomorHeader: "366092",
    belomorDiscuss: "F2DCDB",
    yellow: "FFFF00",
    border: "BFBFBF"
  };

  const TASKS = [
    { num: 1, name: "Получение ТЗ", type: "рабочие", duration: 5 },
    { num: 2, name: "Тест-фит, разработка и согласование", type: "рабочие", duration: 14 },
    { num: 3, name: "Расчёт стоимости", type: "рабочие", duration: 5 },
    { num: 4, name: "Согласование WPM", type: "рабочие", duration: 1 },
    { num: 5, name: "Согласование HRD/CFO/CEO", type: "рабочие", duration: 14 },
    { num: 6, name: "Согласование ДА с АРДД", type: "рабочие", duration: 21 },
    { num: 7, name: "Согласование ЗП на ДА (сущ. аренда)", type: "рабочие", duration: 21 },
    { num: 8, name: "Подписание ДА", type: "рабочие", duration: 7 },
    { num: 9, name: "Тендер на дизайн-проект", type: "рабочие", duration: 29 },
    { num: 10, name: "Согласование ЗП на проектирование и подписание договора", type: "рабочие", duration: 19 },
    { num: 11, name: "Проектирование и подписание договора", type: "рабочие", duration: 75 },
    { num: 12, name: "Тендер на Design&Build", type: "рабочие", duration: 29 },
    { num: 13, name: "Согласование ЗП на Design&Build", type: "рабочие", duration: 21 },
    { num: 14, name: "Design&Build", type: "календарные", duration: 140 },
    { num: 15, name: "Запуск", type: "рабочие", duration: 15 }
  ];

  const state = {
    excelMode: "belomor",
    newMode: "daily",
    editMode: "inc",
    currentIndex: null,
    projects: []
  };

  const els = {};

  document.addEventListener("DOMContentLoaded", () => {
    cacheElements();
    hydrateIcons();
    state.projects = loadProjects();
    fillTaskSelect();
    bindEvents();
    renderHistory();
    renderCurrentProject();
  });

  function cacheElements() {
    els.tabs = Array.from(document.querySelectorAll("[data-tab-target]"));
    els.views = Array.from(document.querySelectorAll("[data-tab-view]"));
    els.excelModeButtons = Array.from(document.querySelectorAll("[data-excel-mode]"));
    els.newModeButtons = Array.from(document.querySelectorAll("[data-new-mode]"));
    els.editModeButtons = Array.from(document.querySelectorAll("[data-edit-mode]"));
    els.excelFile = document.getElementById("excel-file");
    els.excelFileName = document.getElementById("excel-file-name");
    els.excelProcess = document.getElementById("excel-process");
    els.excelStatus = document.getElementById("excel-status");
    els.excelDownloads = document.getElementById("excel-downloads");
    els.projectStart = document.getElementById("project-start");
    els.projectName = document.getElementById("project-name");
    els.createProject = document.getElementById("create-project");
    els.projectSummary = document.getElementById("project-summary");
    els.historyList = document.getElementById("history-list");
    els.taskPreview = document.getElementById("task-preview");
    els.scheduleStatus = document.getElementById("schedule-status");
    els.scheduleDownloads = document.getElementById("schedule-downloads");
    els.taskSelect = document.getElementById("task-select");
    els.editDays = document.getElementById("edit-days");
    els.applyEdit = document.getElementById("apply-edit");
    els.deleteProject = document.getElementById("delete-project");
  }

  function hydrateIcons() {
    if (window.lucide) {
      window.lucide.createIcons();
    }
  }

  function bindEvents() {
    els.tabs.forEach((button) => {
      button.addEventListener("click", () => switchTab(button.dataset.tabTarget));
    });

    els.excelModeButtons.forEach((button) => {
      button.addEventListener("click", () => {
        state.excelMode = button.dataset.excelMode;
        setActive(els.excelModeButtons, button);
      });
    });

    els.newModeButtons.forEach((button) => {
      button.addEventListener("click", () => {
        state.newMode = button.dataset.newMode;
        setActive(els.newModeButtons, button);
      });
    });

    els.editModeButtons.forEach((button) => {
      button.addEventListener("click", () => {
        state.editMode = button.dataset.editMode;
        setActive(els.editModeButtons, button);
      });
    });

    els.excelFile.addEventListener("change", () => {
      const file = els.excelFile.files && els.excelFile.files[0];
      els.excelFileName.textContent = file ? file.name : "Выберите .xlsx или .xls";
    });

    els.excelProcess.addEventListener("click", processExcel);
    els.createProject.addEventListener("click", createProjectFromForm);
    els.applyEdit.addEventListener("click", applyDurationEdit);
    els.deleteProject.addEventListener("click", deleteCurrentProject);

    document.querySelectorAll("[data-generate-current]").forEach((button) => {
      button.addEventListener("click", () => generateCurrentProject(button.dataset.generateCurrent));
    });
  }

  function switchTab(tabName) {
    els.tabs.forEach((tab) => tab.classList.toggle("is-active", tab.dataset.tabTarget === tabName));
    els.views.forEach((view) => view.classList.toggle("is-active", view.dataset.tabView === tabName));
  }

  function setActive(buttons, activeButton) {
    buttons.forEach((button) => button.classList.toggle("is-active", button === activeButton));
  }

  function setStatus(element, message, tone) {
    element.textContent = message || "";
    element.classList.remove("ok", "error");
    if (tone) {
      element.classList.add(tone);
    }
  }

  function fillTaskSelect() {
    els.taskSelect.innerHTML = TASKS.map((task) => {
      return `<option value="${task.num}">${task.num}. ${escapeHtml(task.name)}</option>`;
    }).join("");
  }

  async function processExcel() {
    try {
      requireLibs(["XLSX", "ExcelJS"]);
      const file = els.excelFile.files && els.excelFile.files[0];
      if (!file) {
        setStatus(els.excelStatus, "Выберите Excel файл.", "error");
        return;
      }
      if (!/\.(xlsx|xls)$/i.test(file.name)) {
        setStatus(els.excelStatus, "Пожалуйста, отправьте файл в формате Excel (.xlsx или .xls).", "error");
        return;
      }

      setStatus(els.excelStatus, "Начинаю обработку файла...");
      els.excelDownloads.innerHTML = "";

      const rows = await readFirstSheet(file);
      const result = state.excelMode === "belomor"
        ? await buildBelomorFiles(rows)
        : await buildVvodFiles(rows);

      els.excelDownloads.innerHTML = "";
      renderDownload(result.excelBlob, result.excelName, "XLSX", els.excelDownloads);
      setStatus(els.excelStatus, "Файл успешно отформатирован.", "ok");
      hydrateIcons();
    } catch (error) {
      console.error(error);
      setStatus(els.excelStatus, `Произошла ошибка: ${error.message}`, "error");
    }
  }

  async function readFirstSheet(file) {
    const buffer = await file.arrayBuffer();
    const workbook = window.XLSX.read(buffer, { type: "array", cellDates: true });
    const firstName = workbook.SheetNames[0];
    if (!firstName) {
      throw new Error("В книге нет листов.");
    }
    const sheet = workbook.Sheets[firstName];
    const rows = window.XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: null,
      blankrows: true,
      raw: false,
      dateNF: "dd.mm.yyyy"
    });
    return normalizeRows(rows.length ? rows : [[]]);
  }

  async function buildBelomorFiles(inputRows) {
    const rows = transformBelomor(inputRows);
    const workbook = new window.ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet1");
    addRows(sheet, rows);
    styleBelomorSheet(sheet, rows);
    const excelBlob = await workbookToBlob(workbook);
    const baseName = `Беломор_${todayRu()}`;
    return {
      excelBlob,
      excelName: `${baseName}.xlsx`
    };
  }

  function transformBelomor(inputRows) {
    const maxCols = Math.max(1, ...inputRows.map((row) => row.length));
    let rows = inputRows.map((row, rowIndex) => {
      const next = normalizeRow(row, maxCols);
      next.unshift(rowIndex === 0 ? "К обсуждению" : "нет");
      return next;
    });

    const priorityIndex = findHeaderIndex(rows[0], "приоритет");
    if (priorityIndex >= 0) {
      rows = rows.map((row) => {
        const next = row.slice();
        next.splice(priorityIndex, 1);
        return next;
      });
    }

    const deadlineIndex = findHeaderIndex(rows[0], "дедлайн");
    if (deadlineIndex >= 0) {
      for (let i = 1; i < rows.length; i += 1) {
        rows[i][deadlineIndex] = normalizeRuDateValue(rows[i][deadlineIndex]);
      }
    }

    let performerIndex = findHeaderIndex(rows[0], "исполн");
    if (performerIndex < 0) {
      performerIndex = 4;
    }
    rows = ensureColumn(rows, performerIndex);

    for (let i = 1; i < rows.length; i += 1) {
      rows[i][performerIndex] = processPerformerCell(rows[i][performerIndex]);
    }

    rows = ensureColumn(rows, 2);
    for (let i = 0; i < rows.length; i += 1) {
      if (isBlank(rows[i][2])) {
        rows[i][2] = "TBD";
      }
    }
    return rows;
  }

  function styleBelomorSheet(sheet, rows) {
    const maxRow = rows.length;
    const maxCol = Math.max(1, ...rows.map((row) => row.length));
    const performerCol = Math.max(1, findHeaderIndex(rows[0], "исполн") + 1 || 5);
    const deadlineCol = findHeaderIndex(rows[0], "дедлайн") + 1;

    [13, 30, 10, 17, 143, 11].forEach((width, index) => {
      sheet.getColumn(index + 1).width = width;
    });

    for (let rowNum = 1; rowNum <= maxRow; rowNum += 1) {
      const row = sheet.getRow(rowNum);
      for (let colNum = 1; colNum <= maxCol; colNum += 1) {
        const cell = row.getCell(colNum);
        cell.border = thinBorder();
        if (rowNum === 1) {
          cell.font = { bold: true, size: 12, color: { argb: argb("FFFFFF") } };
          setFill(cell, COLORS.belomorHeader);
          cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        } else {
          cell.font = { size: 10 };
          if (colNum === 1 || colNum === 3) {
            cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
          } else {
            cell.alignment = { vertical: "middle", wrapText: true };
          }
        }
      }
      row.getCell(1).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      row.getCell(3).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      row.getCell(4).alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      row.getCell(performerCol).alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      if (deadlineCol > 0 && rowNum > 1) {
        row.getCell(deadlineCol).numFmt = "dd.mm.yyyy";
      }
    }

    for (let rowNum = 2; rowNum <= maxRow; rowNum += 1) {
      setFill(sheet.getRow(rowNum).getCell(1), COLORS.belomorDiscuss);
    }

    sheet.pageSetup = {
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 }
    };
  }

  async function buildVvodFiles(inputRows) {
    const rows = normalizeDateValues(normalizeRows(inputRows));
    const workbook = new window.ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet0", {
      views: [{ state: "frozen", ySplit: 1 }]
    });
    addRows(sheet, rows);
    styleVvodSheet(sheet, rows);
    const excelBlob = await workbookToBlob(workbook);
    const baseName = `Ввод_в_строй_${todayRu()}`;
    return {
      excelBlob,
      excelName: `${baseName}.xlsx`
    };
  }

  function styleVvodSheet(sheet, rows) {
    const maxRow = Math.max(1, rows.length);
    const maxCol = Math.max(1, ...rows.map((row) => row.length));

    for (let rowNum = 1; rowNum <= 65; rowNum += 1) {
      const row = sheet.getRow(rowNum);
      for (let colNum = 1; colNum <= 23; colNum += 1) {
        const cell = row.getCell(colNum);
        if (isBlank(cell.value)) {
          setFill(cell, COLORS.yellow);
        }
      }
    }

    for (let colNum = 1; colNum <= maxCol; colNum += 1) {
      sheet.getRow(1).getCell(colNum).font = { bold: true };
    }

    for (let rowNum = 1; rowNum <= maxRow; rowNum += 1) {
      const row = sheet.getRow(rowNum);
      for (let colNum = 1; colNum <= maxCol; colNum += 1) {
        const cell = row.getCell(colNum);
        if (colNum === 2) {
          cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
        } else if (colNum === 22) {
          cell.alignment = { horizontal: "left", vertical: "top", wrapText: true };
        } else {
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }
      }
    }

    const widths = { B: 39, S: 12.5, T: 20, U: 19.5, V: 69, W: 20 };
    Object.entries(widths).forEach(([letter, width]) => {
      sheet.getColumn(letter).width = width;
    });

    ["E", "H", "J", "K", "L", "M", "N", "O", "P", "S", "U", "W"].forEach((letter) => {
      const column = sheet.getColumn(letter);
      column.eachCell({ includeEmpty: true }, (cell) => {
        const current = cell.alignment || {};
        cell.alignment = { horizontal: current.horizontal, vertical: current.vertical, wrapText: true };
      });
    });

    ["X", "Y", "Z"].forEach((letter) => {
      sheet.getColumn(letter).outlineLevel = 1;
    });

    for (let colNum = 1; colNum <= maxCol; colNum += 1) {
      const letter = columnLetter(colNum);
      if (Object.prototype.hasOwnProperty.call(widths, letter)) {
        continue;
      }
      let maxLength = 0;
      for (let rowNum = 1; rowNum <= maxRow; rowNum += 1) {
        const value = sheet.getRow(rowNum).getCell(colNum).value;
        if (!isBlank(value)) {
          maxLength = Math.max(maxLength, String(value).length);
        }
      }
      sheet.getColumn(colNum).width = Math.min(maxLength + 2, 50);
    }

    const tableReady = rows.length > 0 && maxCol > 0;
    if (tableReady) {
      try {
        const headers = uniqueHeaders(normalizeRow(rows[0], maxCol));
        const bodyRows = rows.slice(1).map((row) => normalizeRow(row, maxCol));
        sheet.addTable({
          name: "FormattedTable",
          ref: "A1",
          headerRow: true,
          totalsRow: false,
          style: {
            theme: "TableStyleMedium2",
            showFirstColumn: false,
            showLastColumn: false,
            showRowStripes: true,
            showColumnStripes: false
          },
          columns: headers.map((name) => ({ name })),
          rows: bodyRows
        });
      } catch (error) {
        console.warn("Не удалось добавить Excel-таблицу", error);
      }
    }

    for (let rowNum = 1; rowNum <= maxRow; rowNum += 1) {
      const row = sheet.getRow(rowNum);
      for (let colNum = 1; colNum <= maxCol; colNum += 1) {
        row.getCell(colNum).border = thinBorder();
      }
    }

    sheet.pageSetup = {
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: {
        left: 0.6 / 2.54,
        right: 0.3 / 2.54,
        top: 0.5 / 2.54,
        bottom: 0.4 / 2.54,
        header: 0.3 / 2.54,
        footer: 0.3 / 2.54
      }
    };
  }

  function renderDownload(blob, fileName, label, container) {
    const url = URL.createObjectURL(blob);
    const card = document.createElement("div");
    card.className = "download-card";
    card.innerHTML = `
      <div>
        <div class="download-title">${escapeHtml(fileName)}</div>
        <div class="download-meta">${label} · ${formatBytes(blob.size)}</div>
      </div>
      <a class="download-link" href="${url}" download="${escapeAttribute(fileName)}">
        <i data-lucide="download"></i>
        Скачать
      </a>
    `;
    container.appendChild(card);
  }

  async function createProjectFromForm() {
    try {
      requireLibs(["ExcelJS"]);
      const start = parseRuDate(els.projectStart.value.trim());
      const name = els.projectName.value.trim();
      if (!start) {
        setStatus(els.scheduleStatus, "Неверный формат даты. Пример: 19.12.2026", "error");
        return;
      }
      if (!name) {
        setStatus(els.scheduleStatus, "Введите название проекта.", "error");
        return;
      }
      const record = {
        name,
        start: formatRuDate(start),
        files: { daily: null, weekly: null },
        durations: {}
      };
      state.projects.unshift(record);
      state.currentIndex = 0;
      saveProjects();
      renderHistory();
      renderCurrentProject();
      await generateCurrentProject(state.newMode);
      els.projectName.value = "";
    } catch (error) {
      console.error(error);
      setStatus(els.scheduleStatus, `Произошла ошибка: ${error.message}`, "error");
    }
  }

  async function generateCurrentProject(mode) {
    try {
      requireLibs(["ExcelJS"]);
      const record = getCurrentProject();
      if (!record) {
        setStatus(els.scheduleStatus, "Сначала создайте или выберите проект.", "error");
        return;
      }
      const start = parseRuDate(record.start);
      if (!start) {
        setStatus(els.scheduleStatus, "Некорректная дата в проекте.", "error");
        return;
      }

      setStatus(els.scheduleStatus, "Формирую график...");
      const tasks = buildTasksForRecord(record);
      const holidaySet = makeRuHolidaySet(start.getFullYear(), 3);
      const schedule = computeSchedule(start, tasks, holidaySet);
      const allDates = Object.values(schedule).flat();
      const endDate = allDates.length ? maxDate(allDates) : start;
      const workbook = mode === "weekly"
        ? makeWeeklyWorkbook(record.name, start, schedule, tasks, holidaySet)
        : makeDailyWorkbook(record.name, start, schedule, tasks, holidaySet);
      const blob = await workbookToBlob(workbook);
      const suffix = mode === "weekly" ? "_week" : "_day";
      const fileName = `${record.name}_с_${formatRuDate(start)}_по_${formatRuDate(endDate)}${suffix}.xlsx`;
      record.files[mode] = fileName;
      saveProjects();
      renderHistory();
      renderCurrentProject();
      els.scheduleDownloads.innerHTML = "";
      renderDownload(blob, fileName, "XLSX", els.scheduleDownloads);
      hydrateIcons();
      downloadBlob(blob, fileName);
      setStatus(els.scheduleStatus, "Готово. Можно выбрать другой формат, отредактировать или открыть историю.", "ok");
    } catch (error) {
      console.error(error);
      setStatus(els.scheduleStatus, `Произошла ошибка: ${error.message}`, "error");
    }
  }

  function applyDurationEdit() {
    const record = getCurrentProject();
    if (!record) {
      setStatus(els.scheduleStatus, "Сначала выберите проект из истории или создайте его.", "error");
      return;
    }
    const num = Number(els.taskSelect.value);
    const delta = Math.abs(Number.parseInt(els.editDays.value, 10));
    if (!delta) {
      setStatus(els.scheduleStatus, "Изменение на 0 дней ничего не меняет. Попробуйте другое значение.", "error");
      return;
    }
    const oldDuration = getDurationForRecord(record, num);
    const newDuration = state.editMode === "inc" ? oldDuration + delta : oldDuration - delta;
    if (newDuration < 1) {
      setStatus(els.scheduleStatus, `Нельзя сделать длительность меньше 1 дня. Сейчас ${oldDuration} дн.`, "error");
      return;
    }

    record.durations[String(num)] = newDuration;
    record.files.daily = null;
    record.files.weekly = null;
    saveProjects();
    renderHistory();
    renderCurrentProject();

    const taskName = getBaseTask(num).name;
    setStatus(
      els.scheduleStatus,
      `Длительность процесса ${num}. ${taskName} изменена с ${oldDuration} до ${newDuration} дней. Новые графики можно сформировать кнопками выше.`,
      "ok"
    );
  }

  function deleteCurrentProject() {
    if (state.currentIndex === null || !state.projects[state.currentIndex]) {
      setStatus(els.scheduleStatus, "Проект не выбран.", "error");
      return;
    }
    state.projects.splice(state.currentIndex, 1);
    state.currentIndex = null;
    saveProjects();
    renderHistory();
    renderCurrentProject();
    setStatus(els.scheduleStatus, "Проект удалён из истории.", "ok");
  }

  function renderHistory() {
    if (!state.projects.length) {
      els.historyList.innerHTML = `<p class="muted">История пуста.</p>`;
      return;
    }
    els.historyList.innerHTML = state.projects.slice(0, 50).map((record, index) => {
      const modes = [];
      if (record.files && record.files.daily) modes.push("дни");
      if (record.files && record.files.weekly) modes.push("недели");
      const modeText = modes.length ? modes.join(", ") : "пока нет файлов";
      return `
        <article class="history-item">
          <div>
            <div class="history-title">${index + 1}. ${escapeHtml(record.name || "Без имени")}</div>
            <div class="history-meta">${escapeHtml(record.start || "?")} · ${escapeHtml(modeText)}</div>
          </div>
          <div class="history-actions">
            <button class="small-btn" type="button" data-open-project="${index}" aria-label="Открыть ${index + 1}">
              <i data-lucide="folder-open"></i>
            </button>
          </div>
        </article>
      `;
    }).join("");
    els.historyList.querySelectorAll("[data-open-project]").forEach((button) => {
      button.addEventListener("click", () => {
        state.currentIndex = Number(button.dataset.openProject);
        renderCurrentProject();
        setStatus(els.scheduleStatus, "Проект открыт.", "ok");
      });
    });
    hydrateIcons();
  }

  function renderCurrentProject() {
    const record = getCurrentProject();
    if (!record) {
      els.projectSummary.innerHTML = `<p class="muted">Проект не выбран.</p>`;
      els.taskPreview.innerHTML = `<p class="muted">Нет расчёта.</p>`;
      els.scheduleDownloads.innerHTML = "";
      return;
    }
    els.projectSummary.innerHTML = `
      <div class="summary-title">${escapeHtml(record.name)}</div>
      <div class="summary-line">Старт: ${escapeHtml(record.start)}</div>
      <div class="summary-line">Файлы: ${escapeHtml(projectModesText(record))}</div>
    `;
    renderTaskPreview(record);
  }

  function renderTaskPreview(record) {
    const start = parseRuDate(record.start);
    if (!start) {
      els.taskPreview.innerHTML = `<p class="muted">Некорректная дата.</p>`;
      return;
    }
    const tasks = buildTasksForRecord(record);
    const holidaySet = makeRuHolidaySet(start.getFullYear(), 3);
    const schedule = computeSchedule(start, tasks, holidaySet);
    els.taskPreview.innerHTML = tasks.map((task) => {
      const dates = schedule[task.num] || [];
      return `
        <article class="task-row">
          <span class="task-num">${task.num}</span>
          <div>
            <div class="task-name">${escapeHtml(task.name)}</div>
            <div class="task-meta">${task.duration} дн. · ${escapeHtml(labelFor(dates))}</div>
          </div>
        </article>
      `;
    }).join("");
  }

  function projectModesText(record) {
    const modes = [];
    if (record.files && record.files.daily) modes.push("подневной");
    if (record.files && record.files.weekly) modes.push("понедельный");
    return modes.length ? modes.join(", ") : "пока нет файлов";
  }

  function getCurrentProject() {
    if (state.currentIndex === null) {
      return null;
    }
    const record = state.projects[state.currentIndex];
    if (!record) {
      return null;
    }
    record.files = record.files || { daily: null, weekly: null };
    record.files.daily = record.files.daily || null;
    record.files.weekly = record.files.weekly || null;
    record.durations = record.durations || {};
    return record;
  }

  function loadProjects() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      const parsed = raw ? JSON.parse(raw) : [];
      return Array.isArray(parsed) ? parsed : [];
    } catch (error) {
      console.warn("Не удалось прочитать историю", error);
      return [];
    }
  }

  function saveProjects() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.projects));
  }

  function buildTasksForRecord(record) {
    const overrides = record.durations || {};
    return TASKS.map((task) => {
      const override = Number.parseInt(overrides[String(task.num)], 10);
      return { ...task, duration: Number.isFinite(override) ? override : task.duration };
    });
  }

  function getDurationForRecord(record, num) {
    const task = buildTasksForRecord(record).find((item) => item.num === num);
    return task ? task.duration : getBaseTask(num).duration;
  }

  function getBaseTask(num) {
    const task = TASKS.find((item) => item.num === num);
    if (!task) {
      throw new Error(`Процесс ${num} не найден.`);
    }
    return task;
  }

  function computeSchedule(userStart, tasks, holidaySet) {
    const schedule = {};
    const startOf = {};
    const endOf = {};
    const taskMap = new Map(tasks.map((task) => [task.num, task]));
    const nums = Array.from(taskMap.keys()).sort((a, b) => a - b);

    const place = (startDay, type, duration) => {
      return String(type || "").toLowerCase().includes("календар")
        ? calendarSpan(startDay, duration)
        : workSpan(startDay, duration, holidaySet);
    };

    nums.forEach((num) => {
      const task = taskMap.get(num);
      const type = String(task.type || "").toLowerCase();
      const duration = Number(task.duration);
      let start;

      if (num === 1) {
        start = cloneDate(userStart);
        if (type.includes("рабоч") && !isWorkday(start, holidaySet)) {
          start = nextWorkday(start, holidaySet);
        }
      } else if (num >= 2 && num <= 6) {
        start = addDays(endOf[num - 1], 1);
        if (type.includes("рабоч")) start = nextWorkday(start, holidaySet);
      } else if (num === 7) {
        start = addDays(endOf[6], 1);
        if (type.includes("рабоч")) start = nextWorkday(start, holidaySet);
      } else if (num === 8) {
        start = addDays(endOf[7], 1);
        if (type.includes("рабоч")) start = nextWorkday(start, holidaySet);
      } else if (num === 9) {
        start = cloneDate(startOf[7]);
        if (type.includes("рабоч") && !isWorkday(start, holidaySet)) {
          start = nextWorkday(start, holidaySet);
        }
      } else if (num === 10) {
        start = addDays(endOf[9], 1);
        if (type.includes("рабоч")) start = nextWorkday(start, holidaySet);
      } else if (num === 11) {
        start = addDays(endOf[10], 1);
        if (type.includes("рабоч")) start = nextWorkday(start, holidaySet);
      } else if (num === 12) {
        start = addDays(endOf[11], 1);
        if (type.includes("рабоч")) start = nextWorkday(start, holidaySet);
      } else if (num === 13) {
        start = addDays(endOf[12], 1);
        if (type.includes("рабоч")) start = nextWorkday(start, holidaySet);
      } else if (num === 14) {
        start = addDays(endOf[13], 1);
      } else if (num === 15) {
        start = addDays(endOf[14], 1);
        if (type.includes("рабоч")) start = nextWorkday(start, holidaySet);
      } else {
        start = endOf[num - 1] ? addDays(endOf[num - 1], 1) : cloneDate(userStart);
        if (type.includes("рабоч")) start = nextWorkday(start, holidaySet);
      }

      const dates = place(start, type, duration);
      schedule[num] = dates;
      if (dates.length) {
        startOf[num] = minDate(dates);
        endOf[num] = maxDate(dates);
      }
    });

    return schedule;
  }

  function makeDailyWorkbook(projectName, userStart, allDates, tasks, holidaySet) {
    const flattened = Object.values(allDates).flat();
    const maxOffset = flattened.length ? Math.max(...flattened.map((d) => diffDays(d, userStart))) : 60;
    const totalDays = Math.max(60, maxOffset + 3);
    const days = Array.from({ length: totalDays }, (_, index) => addDays(userStart, index));
    const workbook = new window.ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("График (дни)");

    sheet.getColumn("A").width = 4;
    sheet.getColumn("B").width = 2.5;
    sheet.getColumn("E").width = 2.5;
    const dayCol0 = 6;
    for (let i = 0; i < totalDays; i += 1) {
      sheet.getColumn(dayCol0 + i).width = 4.2;
    }

    sheet.mergeCells(1, 1, 2, 4);
    const header = sheet.getCell("A1");
    header.value = "График работ (подневной)";
    header.alignment = { horizontal: "center", vertical: "middle" };
    header.font = { name: "Calibri", size: 9, bold: true };
    setFill(header, COLORS.headerBg);

    sheet.getCell("C4").value = "действие";
    sheet.getCell("D4").value = "Длительность";
    ["C4", "D4"].forEach((addr) => {
      const cell = sheet.getCell(addr);
      cell.font = { name: "Calibri", size: 9, bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      setFill(cell, COLORS.leftHeaderBg);
      cell.border = thinBorder();
    });

    let i = 0;
    while (i < totalDays) {
      const month = days[i].getMonth();
      const year = days[i].getFullYear();
      let j = i;
      while (j < totalDays && days[j].getMonth() === month) j += 1;
      sheet.mergeCells(2, dayCol0 + i, 2, dayCol0 + j - 1);
      const cell = sheet.getCell(2, dayCol0 + i);
      cell.value = `${monthNameRu(days[i])} ${year}`;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.font = { name: "Calibri", size: 9, bold: true };
      setFill(cell, COLORS.headerBg);
      i = j;
    }

    let currentWeek = isoWeek(days[0]).week;
    let startIndex = 0;
    for (let idx = 1; idx <= totalDays; idx += 1) {
      if (idx === totalDays || isoWeek(days[idx]).week !== currentWeek) {
        sheet.mergeCells(3, dayCol0 + startIndex, 3, dayCol0 + idx - 1);
        const cell = sheet.getCell(3, dayCol0 + startIndex);
        cell.value = currentWeek;
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.font = { name: "Calibri", size: 8, italic: true };
        cell.border = thinBorder();
        if (idx < totalDays) {
          currentWeek = isoWeek(days[idx]).week;
          startIndex = idx;
        }
      }
    }

    days.forEach((day, index) => {
      const cell = sheet.getCell(4, dayCol0 + index);
      cell.value = day.getDate();
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.font = { name: "Calibri", size: 9 };
      if (isHoliday(day, holidaySet)) {
        setFill(cell, COLORS.holiday);
      } else if (isWeekend(day)) {
        setFill(cell, COLORS.weekend);
      }
      cell.border = thinBorder();
    });

    for (let row = 2; row <= 4; row += 1) {
      for (let col = dayCol0; col < dayCol0 + totalDays; col += 1) {
        sheet.getCell(row, col).border = thinBorder();
      }
    }

    tasks.forEach((task) => {
      const rowNum = 4 + task.num;
      sheet.getCell(rowNum, 1).value = task.num;
      sheet.getCell(rowNum, 3).value = task.name;
      sheet.getCell(rowNum, 4).value = labelFor(allDates[task.num] || []);
      days.forEach((day, index) => {
        const cell = sheet.getCell(rowNum, dayCol0 + index);
        cell.border = thinBorder();
        if (isHoliday(day, holidaySet)) {
          setFill(cell, COLORS.holiday);
        } else if (isWeekend(day)) {
          setFill(cell, COLORS.weekend);
        }
      });

      const barColor = task.num <= 5 ? COLORS.taskGreen : task.num <= 10 ? COLORS.taskYellow : COLORS.taskOrange;
      (allDates[task.num] || []).forEach((day) => {
        const index = diffDays(day, userStart);
        if (index >= 0 && index < totalDays && !isHoliday(day, holidaySet)) {
          setFill(sheet.getCell(rowNum, dayCol0 + index), barColor);
        }
      });
    });

    autoFitColumn(sheet, "C", 14);
    autoFitColumn(sheet, "D", 14);
    sheet.views = [{ state: "frozen", xSplit: 5, ySplit: 4 }];
    sheet.pageSetup = { orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
    workbook.properties.title = projectName;
    return workbook;
  }

  function makeWeeklyWorkbook(projectName, userStart, allDates, tasks) {
    const flattened = Object.values(allDates).flat();
    const minD = flattened.length ? minDate(flattened) : userStart;
    const maxD = flattened.length ? maxDate(flattened) : addDays(userStart, 60);
    const mondayStart = addDays(minD, -(isoWeekday(minD) - 1));
    const mondayLast = addDays(maxD, -(isoWeekday(maxD) - 1));
    const weeks = [];
    const weekIndex = new Map();
    let cursor = cloneDate(mondayStart);
    while (cursor <= mondayLast) {
      const iso = isoWeek(cursor);
      const key = `${iso.year}-${iso.week}`;
      weekIndex.set(key, weeks.length);
      weeks.push({ key, monday: cloneDate(cursor), week: iso.week });
      cursor = addDays(cursor, 7);
    }

    const workbook = new window.ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("График (недели)");
    sheet.getColumn("A").width = 4;
    sheet.getColumn("B").width = 2.5;
    sheet.getColumn("E").width = 2.5;
    const weekCol0 = 6;
    weeks.forEach((_, index) => {
      sheet.getColumn(weekCol0 + index).width = 4.2;
    });

    sheet.mergeCells(1, 1, 2, 4);
    const header = sheet.getCell("A1");
    header.value = "График работ (понедельный)";
    header.alignment = { horizontal: "center", vertical: "middle" };
    header.font = { name: "Calibri", size: 9, bold: true };
    setFill(header, COLORS.headerBg);

    sheet.getCell("C4").value = "действие";
    sheet.getCell("D4").value = "Длительность";
    ["C4", "D4"].forEach((addr) => {
      const cell = sheet.getCell(addr);
      cell.font = { name: "Calibri", size: 9, bold: true };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      setFill(cell, COLORS.leftHeaderBg);
      cell.border = thinBorder();
    });

    let i = 0;
    while (i < weeks.length) {
      const month = weeks[i].monday.getMonth();
      const year = weeks[i].monday.getFullYear();
      let j = i;
      while (j < weeks.length && weeks[j].monday.getMonth() === month) j += 1;
      sheet.mergeCells(2, weekCol0 + i, 2, weekCol0 + j - 1);
      const cell = sheet.getCell(2, weekCol0 + i);
      cell.value = `${monthNameRu(weeks[i].monday)} ${year}`;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.font = { name: "Calibri", size: 9, bold: true };
      setFill(cell, COLORS.headerBg);
      i = j;
    }

    weeks.forEach((week, index) => {
      const cell = sheet.getCell(3, weekCol0 + index);
      cell.value = week.week;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.font = { name: "Calibri", size: 9 };
      cell.border = thinBorder();
    });

    for (let row = 2; row <= 3; row += 1) {
      for (let col = weekCol0; col < weekCol0 + weeks.length; col += 1) {
        sheet.getCell(row, col).border = thinBorder();
      }
    }

    tasks.forEach((task) => {
      const rowNum = 4 + task.num;
      sheet.getCell(rowNum, 1).value = task.num;
      sheet.getCell(rowNum, 3).value = task.name;
      sheet.getCell(rowNum, 4).value = labelFor(allDates[task.num] || []);
      weeks.forEach((_, index) => {
        sheet.getCell(rowNum, weekCol0 + index).border = thinBorder();
      });

      const counts = new Map();
      (allDates[task.num] || []).forEach((day) => {
        const iso = isoWeek(day);
        const index = weekIndex.get(`${iso.year}-${iso.week}`);
        if (index !== undefined) {
          counts.set(index, (counts.get(index) || 0) + 1);
        }
      });
      counts.forEach((count, index) => {
        const color = count <= 2 ? COLORS.week12 : count <= 4 ? COLORS.week34 : COLORS.week5;
        setFill(sheet.getCell(rowNum, weekCol0 + index), color);
      });
    });

    autoFitColumn(sheet, "C", 14);
    autoFitColumn(sheet, "D", 14);

    const baseRow = 38;
    sheet.getCell(baseRow, 3).value = "Легенда по количеству дней в неделе:";
    sheet.getCell(baseRow, 3).font = { name: "Calibri", size: 9, bold: true };
    [
      ["1–2 дня", COLORS.week12],
      ["3–4 дня", COLORS.week34],
      ["5+ дней", COLORS.week5]
    ].forEach(([label, color], index) => {
      sheet.getCell(baseRow + index + 1, 3).value = label;
      const cell = sheet.getCell(baseRow + index + 1, 4);
      setFill(cell, color);
      cell.border = thinBorder();
    });

    sheet.views = [{ state: "frozen", xSplit: 5, ySplit: 4 }];
    sheet.pageSetup = { orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
    workbook.properties.title = projectName;
    return workbook;
  }

  function addRows(sheet, rows) {
    rows.forEach((row) => sheet.addRow(row.map((value) => (value === undefined ? null : value))));
  }

  function processPerformerCell(value) {
    if (value === null || value === undefined) {
      return "";
    }
    return String(value).trim().replace(/(?<!\s)\(/, " (");
  }

  function normalizeDateValues(rows) {
    return rows.map((row, rowIndex) => {
      if (rowIndex === 0) {
        return row;
      }
      return row.map((value) => normalizeRuDateValue(value));
    });
  }

  function normalizeRuDateValue(value) {
    if (value instanceof Date) {
      return formatRuDate(value);
    }
    if (isBlank(value)) {
      return value;
    }

    const text = String(value).trim();
    const ruMatch = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/.exec(text);
    if (ruMatch) {
      return `${pad2(ruMatch[1])}.${pad2(ruMatch[2])}.${normalizeYear(ruMatch[3])}`;
    }

    const isoMatch = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/.exec(text);
    if (isoMatch) {
      return `${pad2(isoMatch[3])}.${pad2(isoMatch[2])}.${isoMatch[1]}`;
    }

    const usMatch = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(text);
    if (usMatch) {
      return `${pad2(usMatch[2])}.${pad2(usMatch[1])}.${normalizeYear(usMatch[3])}`;
    }

    return value;
  }

  function findHeaderIndex(headerRow, needle) {
    const lowerNeedle = needle.toLowerCase();
    return headerRow.findIndex((value) => typeof value === "string" && value.toLowerCase().includes(lowerNeedle));
  }

  function normalizeRows(rows) {
    const maxCols = Math.max(1, ...rows.map((row) => row.length));
    return rows.map((row) => normalizeRow(row, maxCols));
  }

  function normalizeRow(row, length) {
    const next = Array.isArray(row) ? row.slice(0, length) : [];
    while (next.length < length) {
      next.push(null);
    }
    return next;
  }

  function ensureColumn(rows, index) {
    return rows.map((row) => normalizeRow(row, Math.max(row.length, index + 1)));
  }

  function uniqueHeaders(row) {
    const seen = new Map();
    return row.map((value, index) => {
      const base = isBlank(value) ? `Column${index + 1}` : String(value);
      const count = seen.get(base) || 0;
      seen.set(base, count + 1);
      return count ? `${base}_${count + 1}` : base;
    });
  }

  function workbookToBlob(workbook) {
    return workbook.xlsx.writeBuffer().then((buffer) => new Blob([buffer], { type: XLSX_MIME }));
  }

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  function makeRuHolidaySet(startYear, spanYears) {
    const set = new Set();
    for (let year = startYear; year < startYear + spanYears; year += 1) {
      const fixed = [];
      for (let day = 1; day <= 8; day += 1) {
        fixed.push(new Date(year, 0, day));
      }
      fixed.push(
        new Date(year, 1, 23),
        new Date(year, 2, 8),
        new Date(year, 4, 1),
        new Date(year, 4, 9),
        new Date(year, 5, 12),
        new Date(year, 10, 4)
      );
      fixed.forEach((day) => set.add(dateKey(day)));
      fixed.forEach((day) => {
        if (!isWeekend(day)) {
          return;
        }
        let observed = addDays(day, 1);
        while (isWeekend(observed) || set.has(dateKey(observed))) {
          observed = addDays(observed, 1);
        }
        set.add(dateKey(observed));
      });
    }
    return set;
  }

  function isHoliday(date, holidaySet) {
    return holidaySet.has(dateKey(date));
  }

  function isWeekend(date) {
    return date.getDay() === 0 || date.getDay() === 6;
  }

  function isWorkday(date, holidaySet) {
    return !isWeekend(date) && !isHoliday(date, holidaySet);
  }

  function nextWorkday(date, holidaySet) {
    let current = cloneDate(date);
    while (!isWorkday(current, holidaySet)) {
      current = addDays(current, 1);
    }
    return current;
  }

  function workSpan(start, length, holidaySet) {
    let current = nextWorkday(start, holidaySet);
    const out = [];
    while (out.length < length) {
      if (isWorkday(current, holidaySet)) {
        out.push(cloneDate(current));
      }
      current = addDays(current, 1);
    }
    return out;
  }

  function calendarSpan(start, length) {
    return Array.from({ length }, (_, index) => addDays(start, index));
  }

  function labelFor(dates) {
    if (!dates.length) {
      return "";
    }
    const start = minDate(dates);
    const end = maxDate(dates);
    return `${formatRuDate(start)}–${formatRuDate(end)} (${diffDays(end, start) + 1} дн.)`;
  }

  function parseRuDate(value) {
    const match = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(value);
    if (!match) {
      return null;
    }
    const day = Number(match[1]);
    const month = Number(match[2]);
    const year = Number(match[3]);
    const date = new Date(year, month - 1, day);
    if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
      return null;
    }
    return date;
  }

  function formatRuDate(date) {
    return `${pad2(date.getDate())}.${pad2(date.getMonth() + 1)}.${date.getFullYear()}`;
  }

  function todayRu() {
    return formatRuDate(new Date());
  }

  function monthNameRu(date) {
    const months = [
      "Январь",
      "Февраль",
      "Март",
      "Апрель",
      "Май",
      "Июнь",
      "Июль",
      "Август",
      "Сентябрь",
      "Октябрь",
      "Ноябрь",
      "Декабрь"
    ];
    return months[date.getMonth()];
  }

  function isoWeekday(date) {
    return date.getDay() === 0 ? 7 : date.getDay();
  }

  function isoWeek(date) {
    const utc = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNumber = utc.getUTCDay() || 7;
    utc.setUTCDate(utc.getUTCDate() + 4 - dayNumber);
    const yearStart = new Date(Date.UTC(utc.getUTCFullYear(), 0, 1));
    const week = Math.ceil((((utc - yearStart) / 86400000) + 1) / 7);
    return { year: utc.getUTCFullYear(), week };
  }

  function dateKey(date) {
    return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
  }

  function addDays(date, days) {
    const next = cloneDate(date);
    next.setDate(next.getDate() + days);
    return next;
  }

  function cloneDate(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  function diffDays(a, b) {
    const left = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
    const right = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
    return Math.round((left - right) / 86400000);
  }

  function minDate(dates) {
    return dates.reduce((min, date) => (date < min ? date : min), dates[0]);
  }

  function maxDate(dates) {
    return dates.reduce((max, date) => (date > max ? date : max), dates[0]);
  }

  function isBlank(value) {
    return value === null || value === undefined || String(value).trim() === "";
  }

  function setFill(cell, color) {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb(color) } };
  }

  function thinBorder() {
    const side = { style: "thin", color: { argb: argb(COLORS.border) } };
    return { top: side, right: side, bottom: side, left: side };
  }

  function argb(rgb) {
    return rgb.length === 8 ? rgb : `FF${rgb}`;
  }

  function autoFitColumn(sheet, letter, minWidth) {
    let maxLength = 0;
    sheet.getColumn(letter).eachCell({ includeEmpty: false }, (cell) => {
      maxLength = Math.max(maxLength, String(cell.value || "").length);
    });
    sheet.getColumn(letter).width = Math.max(minWidth, Math.floor(maxLength * 1.1));
  }

  function columnLetter(index) {
    let dividend = index;
    let letter = "";
    while (dividend > 0) {
      const modulo = (dividend - 1) % 26;
      letter = String.fromCharCode(65 + modulo) + letter;
      dividend = Math.floor((dividend - modulo) / 26);
    }
    return letter;
  }

  function pad2(value) {
    return String(value).padStart(2, "0");
  }

  function normalizeYear(value) {
    const text = String(value);
    if (text.length === 4) {
      return text;
    }
    const year = Number(text);
    return String(year >= 70 ? 1900 + year : 2000 + year);
  }

  function formatBytes(bytes) {
    if (bytes < 1024) return `${bytes} Б`;
    if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} КБ`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} МБ`;
  }

  function requireLibs(names) {
    const missing = names.filter((name) => !window[name]);
    if (missing.length) {
      throw new Error(`Не загрузились библиотеки: ${missing.join(", ")}.`);
    }
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function escapeAttribute(value) {
    return escapeHtml(value).replace(/`/g, "&#096;");
  }
})();
