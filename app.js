(function () {
  "use strict";

  const GROUP_SIZE = 12;
  const GROUP_SIZE_OPTIONS = [10, 11, 12, 13, 14, 15, 16];
  const SHIRT_SIZES = [
    "KIDS SMALL",
    "KIDS MEDIUM",
    "KIDS LARGE",
    "KIDS X-LARGE",
    "ADULT SMALL",
    "ADULT MEDIUM",
    "ADULT LARGE",
  ];
  const SHORT_SIZES = [
    "YOUTH SMALL",
    "KIDS MEDIUM",
    "KIDS LARGE",
    "KIDS X-LARGE",
    "ADULT SMALL",
    "ADULT MEDIUM",
    "ADULT LARGE",
  ];
  const SOCK_SIZES = ["34-38", "39-42", "43-46", "47+"];
  const CAMP_NUMBERS = ["1", "2", "3"];

  const el = {
    appShell: document.querySelector("#appShell"),
    fileInput: document.querySelector("#fileInput"),
    loadStatus: document.querySelector("#loadStatus"),
    uploadPanel: document.querySelector("#uploadPanel"),
    sizeViewButton: document.querySelector("#sizeViewButton"),
    groupsViewButton: document.querySelector("#groupsViewButton"),
    controlsPanel: document.querySelector("#controlsPanel"),
    sheetSelect: document.querySelector("#sheetSelect"),
    sortSelect: document.querySelector("#sortSelect"),
    campSelect: document.querySelector("#campSelect"),
    exportButton: document.querySelector("#exportButton"),
    printDashboardButton: document.querySelector("#printDashboardButton"),
    csvDashboardButton: document.querySelector("#csvDashboardButton"),
    printGroupsButton: document.querySelector("#printGroupsButton"),
    csvGroupsButton: document.querySelector("#csvGroupsButton"),
    movePanel: document.querySelector("#movePanel"),
    selectedKidLabel: document.querySelector("#selectedKidLabel"),
    moveTargetSelect: document.querySelector("#moveTargetSelect"),
    moveKidButton: document.querySelector("#moveKidButton"),
    deleteSelectedKidButton: document.querySelector("#deleteSelectedKidButton"),
    clearSelectedKidButton: document.querySelector("#clearSelectedKidButton"),
    moveStatus: document.querySelector("#moveStatus"),
    resetButton: document.querySelector("#resetButton"),
    summaryGrid: document.querySelector("#summaryGrid"),
    kidCount: document.querySelector("#kidCount"),
    groupCount: document.querySelector("#groupCount"),
    activeSheetName: document.querySelector("#activeSheetName"),
    dashboardPanel: document.querySelector("#dashboardPanel"),
    dashboardTables: document.querySelector("#dashboardTables"),
    rosterPanel: document.querySelector("#rosterPanel"),
    groupsPreview: document.querySelector("#groupsPreview"),
  };

  const state = {
    workbook: null,
    rows: [],
    columnMap: {},
    summaries: {},
    kids: [],
    groups: [],
    activeSheet: "",
    fileLabel: "",
    coaches: {},
    campAssignments: {},
    groupSizes: {},
    groupSorts: {},
    selectedKid: null,
    activeView: "size",
  };

  function setStatus(message, isError) {
    el.loadStatus.textContent = message || "";
    el.loadStatus.style.color = isError ? "#b42318" : "";
  }

  function normalizeHeader(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/['’.]/g, "")
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function normalizeSize(value) {
    return String(value || "")
      .toUpperCase()
      .replace(/[‐‑‒–—−]/g, "-")
      .replace(/YOUTH/g, "KIDS")
      .replace(/CHILD/g, "KIDS")
      .replace(/\bEXTRA\s*-\s*LARGE\b/g, "X-LARGE")
      .replace(/\bEXTRA\s+LARGE\b/g, "X-LARGE")
      .replace(/\bEXTRALARGE\b/g, "X-LARGE")
      .replace(/\bX\s*-\s*LARGE\b/g, "X-LARGE")
      .replace(/\bX\s+LARGE\b/g, "X-LARGE")
      .replace(/\bX\s*LARGE\b/g, "X-LARGE")
      .replace(/\bXL\b/g, "X-LARGE")
      .replace(/\s+/g, " ")
      .trim();
  }

  function normalizeShortSize(value) {
    const size = normalizeSize(value);
    if (size === "KIDS SMALL") return "YOUTH SMALL";
    return size;
  }

  function findColumn(headers, candidates) {
    const normalized = headers.map(normalizeHeader);
    for (const candidate of candidates) {
      const wanted = normalizeHeader(candidate);
      const exactIndex = normalized.indexOf(wanted);
      if (exactIndex !== -1) return headers[exactIndex];
    }
    for (let index = 0; index < normalized.length; index += 1) {
      if (candidates.some((candidate) => normalized[index].includes(normalizeHeader(candidate)))) {
        return headers[index];
      }
    }
    return "";
  }

  function readDate(value) {
    if (!value) return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (!parsed) return null;
      return new Date(parsed.y, parsed.m - 1, parsed.d);
    }
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function getAge(date, today = new Date()) {
    if (!date) return "";
    let age = today.getFullYear() - date.getFullYear();
    const birthdayThisYear = new Date(today.getFullYear(), date.getMonth(), date.getDate());
    if (today < birthdayThisYear) age -= 1;
    return age;
  }

  function getSockBucket(value) {
    const match = String(value || "").match(/\d+(\.\d+)?/);
    if (!match) return "";
    const size = Number(match[0]);
    if (size <= 38) return "34-38";
    if (size <= 42) return "39-42";
    if (size <= 46) return "43-46";
    return "47+";
  }

  function countBy(rows, column, allowedSizes, normalizer) {
    const counts = Object.fromEntries(allowedSizes.map((size) => [size, 0]));
    rows.forEach((row) => {
      const size = normalizer(row[column]);
      if (Object.prototype.hasOwnProperty.call(counts, size)) counts[size] += 1;
    });
    return counts;
  }

  function buildColumnMap(rows) {
    const headers = Object.keys(rows[0] || {});
    return {
      camp: findColumn(headers, ["Choose a Camp", "Camp"]),
      firstName: findColumn(headers, ["Player First Name", "First Name"]),
      lastName: findColumn(headers, ["Player Last Name", "Last Name"]),
      gender: findColumn(headers, ["Gender"]),
      dob: findColumn(headers, ["Player Date of Birth", "Date of Birth", "DOB", "Birthday"]),
      shirt: findColumn(headers, ["Shirt size", "T Shirt Size", "T-Shirt Size"]),
      shorts: findColumn(headers, ["Shorts Size", "Short's Size", "Short Size"]),
      shoe: findColumn(headers, ["Shoe Size", "Shoe Size Eur", "Sock Size"]),
    };
  }

  function buildKids(rows, columnMap) {
    return rows
      .map((row, index) => {
        const dob = readDate(row[columnMap.dob]);
        return {
          id: `kid-${index}`,
          firstName: row[columnMap.firstName] || "",
          lastName: row[columnMap.lastName] || "",
          gender: row[columnMap.gender] || "",
          dob,
          age: getAge(dob),
          camp: row[columnMap.camp] || "",
        };
      })
      .filter((kid) => kid.firstName || kid.lastName);
  }

  function sortKids(kids) {
    const direction = el.sortSelect.value === "youngest" ? -1 : 1;
    return [...kids].sort((a, b) => {
      if (!a.dob && !b.dob) return `${a.lastName}${a.firstName}`.localeCompare(`${b.lastName}${b.firstName}`);
      if (!a.dob) return 1;
      if (!b.dob) return -1;
      return (a.dob.getTime() - b.dob.getTime()) * direction;
    });
  }

  function compareKidsByLastName(a, b) {
    const last = String(a.lastName || "").localeCompare(String(b.lastName || ""), undefined, { sensitivity: "base" });
    if (last) return last;
    return String(a.firstName || "").localeCompare(String(b.firstName || ""), undefined, { sensitivity: "base" });
  }

  function compareValues(a, b) {
    if (typeof a === "number" && typeof b === "number") return a - b;
    return String(a ?? "").localeCompare(String(b ?? ""), undefined, { numeric: true, sensitivity: "base" });
  }

  function getKidSortValue(kid, key, index) {
    if (key === "row") return index + 1;
    if (key === "firstName") return kid.firstName || "";
    if (key === "lastName") return kid.lastName || "";
    if (key === "age") return kid.age === "" ? -1 : Number(kid.age);
    if (key === "gender") return kid.gender || "";
    if (key === "group") return "";
    return "";
  }

  function groupLetterForIndex(index) {
    let number = index;
    let letters = "";
    do {
      letters = String.fromCharCode(65 + (number % 26)) + letters;
      number = Math.floor(number / 26) - 1;
    } while (number >= 0);
    return letters;
  }

  function groupLabel(group) {
    return group.letter;
  }

  function makeGroups(kids) {
    const sorted = sortKids(kids);
    const groups = [];
    for (let index = 0; index < sorted.length; index += GROUP_SIZE) {
      const groupIndex = groups.length;
      groups.push({
        number: groupIndex + 1,
        letter: groupLetterForIndex(groupIndex),
        kids: sorted.slice(index, index + GROUP_SIZE).sort(compareKidsByLastName),
      });
    }
    return groups;
  }

  function getGroupSort(group) {
    return state.groupSorts[group.letter] || { key: "lastName", direction: "asc" };
  }

  function sortedGroupKids(group) {
    const sort = getGroupSort(group);
    const direction = sort.direction === "desc" ? -1 : 1;
    return group.kids
      .map((kid, index) => ({ kid, index }))
      .sort((a, b) => {
        const value = compareValues(getKidSortValue(a.kid, sort.key, a.index), getKidSortValue(b.kid, sort.key, b.index));
        if (value) return value * direction;
        return compareKidsByLastName(a.kid, b.kid);
      })
      .map((item) => item.kid);
  }

  function setGroupSort(group, key) {
    const current = getGroupSort(group);
    state.groupSorts[group.letter] = {
      key,
      direction: current.key === key && current.direction === "asc" ? "desc" : "asc",
    };
  }

  function activeGroups() {
    if (!el.campSelect.value || el.campSelect.value === "ALL") return state.groups;
    if (el.campSelect.value === "UNASSIGNED") return state.groups.filter((group) => !getGroupCamp(group));
    return state.groups.filter((group) => getGroupCamp(group) === el.campSelect.value);
  }

  function coachKey(group, coachType) {
    const context = group.letter;
    return `camp-coach-${coachType}-${context}`;
  }

  function getCoach(group, coachType) {
    const key = coachKey(group, coachType);
    return state.coaches[key] ?? localStorage.getItem(key) ?? "";
  }

  function setCoach(group, coachType, value) {
    const key = coachKey(group, coachType);
    state.coaches[key] = value;
    if (value) localStorage.setItem(key, value);
    else localStorage.removeItem(key);
  }

  function captureCoachInputs() {
    document.querySelectorAll(".coach-input").forEach((input) => {
      const group = state.groups.find((candidate) => String(candidate.number) === input.dataset.group);
      if (group) setCoach(group, input.dataset.coach, input.value);
    });
  }

  function groupCampKey(group) {
    return `camp-assignment-${group.letter}`;
  }

  function getGroupCamp(group) {
    const key = groupCampKey(group);
    return state.campAssignments[key] ?? localStorage.getItem(key) ?? "";
  }

  function setGroupCamp(group, campNumber) {
    const key = groupCampKey(group);
    state.campAssignments[key] = campNumber;
    if (campNumber) localStorage.setItem(key, campNumber);
    else localStorage.removeItem(key);
  }

  function captureCampAssignments() {
    document.querySelectorAll(".camp-assignment-select").forEach((select) => {
      const group = state.groups.find((candidate) => String(candidate.number) === select.dataset.group);
      if (group) setGroupCamp(group, select.value);
    });
  }

  function groupSizeKey(group) {
    return `camp-group-size-${group.letter}`;
  }

  function getGroupSize(group) {
    const key = groupSizeKey(group);
    const saved = state.groupSizes[key] ?? localStorage.getItem(key);
    const parsed = Number(saved || GROUP_SIZE);
    return GROUP_SIZE_OPTIONS.includes(parsed) ? parsed : GROUP_SIZE;
  }

  function setGroupSize(group, size) {
    const parsed = Number(size);
    const nextSize = GROUP_SIZE_OPTIONS.includes(parsed) ? parsed : GROUP_SIZE;
    const key = groupSizeKey(group);
    state.groupSizes[key] = String(nextSize);
    localStorage.setItem(key, String(nextSize));
  }

  function captureGroupSizes() {
    document.querySelectorAll(".group-size-select").forEach((select) => {
      const group = state.groups.find((candidate) => String(candidate.number) === select.dataset.group);
      if (group) setGroupSize(group, select.value);
    });
  }

  function findSelectedKid() {
    if (!state.selectedKid) return null;
    const group = state.groups.find((candidate) => candidate.letter === state.selectedKid.groupLetter);
    const kid = group?.kids.find((candidate) => candidate.id === state.selectedKid.kidId);
    return group && kid ? { group, kid } : null;
  }

  function clearSelectedKid(message = "") {
    state.selectedKid = null;
    el.movePanel.hidden = true;
    el.selectedKidLabel.textContent = "No player selected";
    el.moveTargetSelect.innerHTML = "";
    el.moveStatus.textContent = message;
  }

  function updateMovePanel() {
    const selected = findSelectedKid();
    if (!selected) {
      clearSelectedKid();
      return;
    }

    el.movePanel.hidden = false;
    el.selectedKidLabel.textContent = `${selected.kid.firstName} ${selected.kid.lastName} from Group ${selected.group.letter}`;
    el.moveTargetSelect.innerHTML = state.groups
      .filter((group) => group.letter !== selected.group.letter)
      .map((group) => `<option value="${group.letter}">Group ${group.letter} (${group.kids.length}/${getGroupSize(group)})</option>`)
      .join("");
  }

  function moveSelectedKid() {
    const selected = findSelectedKid();
    const targetGroup = state.groups.find((group) => group.letter === el.moveTargetSelect.value);
    if (!selected || !targetGroup) return;

    if (targetGroup.kids.length >= getGroupSize(targetGroup)) {
      el.moveStatus.textContent = `Group ${targetGroup.letter} is full. Increase that group size or choose another group.`;
      return;
    }

    selected.group.kids = selected.group.kids.filter((kid) => kid.id !== selected.kid.id);
    targetGroup.kids.push(selected.kid);
    const targetLetter = targetGroup.letter;
    clearSelectedKid(`${selected.kid.firstName} ${selected.kid.lastName} moved to Group ${targetLetter}.`);
    renderGroups();
  }

  function updatePlayerCount() {
    el.kidCount.textContent = state.groups.reduce((sum, group) => sum + group.kids.length, 0) || state.kids.length;
  }

  function deleteSelectedPlayer() {
    const selected = findSelectedKid();
    if (!selected) return;

    const playerName = `${selected.kid.firstName} ${selected.kid.lastName}`.trim() || "Selected player";
    selected.group.kids = selected.group.kids.filter((kid) => kid.id !== selected.kid.id);
    state.kids = state.kids.filter((kid) => kid.id !== selected.kid.id);
    clearSelectedKid(`${playerName} deleted from Group ${selected.group.letter}.`);
    updatePlayerCount();
    renderGroups();
  }

  function stockKey(section, size) {
    return `camp-stock-${section}-${size}`;
  }

  function getStock(section, size) {
    const value = localStorage.getItem(stockKey(section, size));
    return value === null ? "" : value;
  }

  function setStock(section, size, value) {
    if (value === "") localStorage.removeItem(stockKey(section, size));
    else localStorage.setItem(stockKey(section, size), value);
  }

  function renderSizeTable(section, sizes, counts, showStock) {
    const total = Object.values(counts).reduce((sum, count) => sum + count, 0);
    const table = document.createElement("div");
    table.className = "table-wrap";
    const headerClass = (size) => (size.includes("ADULT") ? "adult-size" : "kid-size");
    table.innerHTML = `
      <table aria-label="${section} size summary">
        <thead>
          <tr>
            <th>${section}</th>
            ${sizes.map((size) => `<th class="${headerClass(size)}">${size}</th>`).join("")}
            <th>TOTAL</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>Customers</td>
            ${sizes.map((size) => `<td>${counts[size] || 0}</td>`).join("")}
            <td>${total}</td>
          </tr>
          ${
            showStock
              ? `<tr>
                  <td>Our Stock</td>
                  ${sizes
                    .map(
                      (size) =>
                        `<td><input class="stock-input" data-section="${section}" data-size="${size}" type="number" min="0" value="${getStock(
                          section,
                          size,
                        )}" aria-label="${section} ${size} stock"></td>`,
                    )
                    .join("")}
                  <td></td>
                </tr>
                <tr>
                  <td>Difference</td>
                  ${sizes.map((size) => `<td class="difference-cell" data-section="${section}" data-size="${size}"></td>`).join("")}
                  <td></td>
                </tr>`
              : `<tr><td>Our Stock</td>${sizes.map(() => "<td></td>").join("")}<td></td></tr>`
          }
        </tbody>
      </table>
    `;
    return table;
  }

  function updateDifferences() {
    document.querySelectorAll(".difference-cell").forEach((cell) => {
      const section = cell.dataset.section;
      const size = cell.dataset.size;
      const input = document.querySelector(`.stock-input[data-section="${section}"][data-size="${size}"]`);
      const stock = input.value === "" ? null : Number(input.value);
      const count = state.summaries[section]?.[size] || 0;
      cell.textContent = stock === null || Number.isNaN(stock) ? "" : stock - count;
      cell.classList.toggle("good", stock !== null && stock - count >= 0);
      cell.classList.toggle("bad", stock !== null && stock - count < 0);
    });
  }

  function renderDashboard() {
    el.dashboardTables.innerHTML = "";
    el.dashboardTables.append(
      renderSizeTable("SHIRTS", SHIRT_SIZES, state.summaries.SHIRTS, true),
      renderSizeTable("SHORTS", SHORT_SIZES, state.summaries.SHORTS, true),
      renderSizeTable("SOCKS", SOCK_SIZES, state.summaries.SOCKS, true),
    );
    document.querySelectorAll(".stock-input").forEach((input) => {
      input.addEventListener("input", () => {
        setStock(input.dataset.section, input.dataset.size, input.value);
        updateDifferences();
      });
    });
    updateDifferences();
  }

  function setActiveView(view) {
    state.activeView = view;
    el.sizeViewButton.classList.toggle("active", view === "size");
    el.groupsViewButton.classList.toggle("active", view === "groups");
    const hasWorkbook = Boolean(state.workbook);
    el.dashboardPanel.hidden = !hasWorkbook || view !== "size";
    el.rosterPanel.hidden = !hasWorkbook || view !== "groups";
  }

  function updateCampOptions() {
    const selected = el.campSelect.value || "ALL";
    el.campSelect.innerHTML = [
      '<option value="ALL">All groups</option>',
      ...CAMP_NUMBERS.map((campNumber) => `<option value="${campNumber}">Camp ${campNumber}</option>`),
      '<option value="UNASSIGNED">Unassigned</option>',
    ].join("");
    el.campSelect.value = selected === "UNASSIGNED" || selected === "ALL" || CAMP_NUMBERS.includes(selected) ? selected : "ALL";
  }

  function renderCampAssignmentOptions(selectedCamp) {
    return [
      '<option value="">Choose camp</option>',
      ...CAMP_NUMBERS.map((campNumber) => `<option value="${campNumber}" ${selectedCamp === campNumber ? "selected" : ""}>Camp ${campNumber}</option>`),
    ].join("");
  }

  function renderGroupSizeOptions(group) {
    const selectedSize = getGroupSize(group);
    return GROUP_SIZE_OPTIONS.map((size) => `<option value="${size}" ${selectedSize === size ? "selected" : ""}>${size}</option>`).join("");
  }

  function renderSortHeader(group, key, label) {
    const sort = getGroupSort(group);
    const active = sort.key === key;
    const arrow = active ? (sort.direction === "asc" ? " ▲" : " ▼") : "";
    return `<button class="sort-header ${active ? "active" : ""}" type="button" data-group="${group.number}" data-sort="${key}" aria-label="Sort group ${group.letter} by ${label}">${label}${arrow}</button>`;
  }

  function renderGroups() {
    el.groupsPreview.innerHTML = "";
    activeGroups().forEach((group) => {
      const card = document.createElement("article");
      card.className = "group-card";
      const groupSize = getGroupSize(group);
      const rows = Array.from({ length: Math.max(groupSize, group.kids.length) }, (_, index) => sortedGroupKids(group)[index] || null);
      const assignedCamp = getGroupCamp(group);
      card.innerHTML = `
        <div class="group-meta">
          <div>
            <label for="camp-${group.letter}">Camp</label>
            <select id="camp-${group.letter}" class="camp-assignment-select" data-group="${group.number}">
              ${renderCampAssignmentOptions(assignedCamp)}
            </select>
          </div>
          <div>
            <label for="size-${group.letter}">Group Size</label>
            <select id="size-${group.letter}" class="group-size-select" data-group="${group.number}">
              ${renderGroupSizeOptions(group)}
            </select>
          </div>
          <div>
            <label for="fcb-${group.letter}">FCB Coach</label>
            <input id="fcb-${group.letter}" class="coach-input" data-group="${group.number}" data-coach="fcb" type="text" value="${escapeHtml(
              getCoach(group, "fcb"),
            )}">
          </div>
          <div>
            <label for="aux-${group.letter}">Aux. Coach</label>
            <input id="aux-${group.letter}" class="coach-input" data-group="${group.number}" data-coach="aux" type="text" value="${escapeHtml(
              getCoach(group, "aux"),
            )}">
          </div>
        </div>
        <h3>Group ${escapeHtml(group.letter)}${assignedCamp ? ` · Camp ${escapeHtml(assignedCamp)}` : ""}</h3>
        <table>
          <thead>
            <tr>
              <th>${renderSortHeader(group, "row", "#")}</th>
              <th>${renderSortHeader(group, "firstName", "First Name")}</th>
              <th>${renderSortHeader(group, "lastName", "Last Name")}</th>
              <th>${renderSortHeader(group, "age", "Age")}</th>
              <th>${renderSortHeader(group, "gender", "Gender")}</th>
              <th>${renderSortHeader(group, "group", "Group")}</th>
            </tr>
          </thead>
          <tbody>
            ${rows
              .map(
                (kid, index) => `
                  <tr class="${kid ? "kid-row" : "empty-row"} ${
                    kid && state.selectedKid?.kidId === kid.id && state.selectedKid?.groupLetter === group.letter ? "selected-row" : ""
                  }" ${kid ? `data-kid-id="${kid.id}" data-group="${group.number}"` : ""}>
                    <td>${index + 1}</td>
                    <td>${kid?.firstName ? escapeHtml(kid.firstName) : "-"}</td>
                    <td>${kid?.lastName ? escapeHtml(kid.lastName) : "-"}</td>
                    <td>${kid?.age ?? "-"}</td>
                    <td>${kid?.gender ? escapeHtml(kid.gender) : "-"}</td>
                    <td>${groupLabel(group)}</td>
                  </tr>
                `,
              )
              .join("")}
          </tbody>
        </table>
      `;
      el.groupsPreview.append(card);
    });
    document.querySelectorAll(".coach-input").forEach((input) => {
      const saveInput = () => {
        const group = state.groups.find((candidate) => String(candidate.number) === input.dataset.group);
        if (group) setCoach(group, input.dataset.coach, input.value);
      };
      input.addEventListener("input", saveInput);
      input.addEventListener("change", saveInput);
    });
    document.querySelectorAll(".camp-assignment-select").forEach((select) => {
      select.addEventListener("change", () => {
        const group = state.groups.find((candidate) => String(candidate.number) === select.dataset.group);
        if (group) {
          setGroupCamp(group, select.value);
          renderGroups();
        }
      });
    });
    document.querySelectorAll(".group-size-select").forEach((select) => {
      select.addEventListener("change", () => {
        const group = state.groups.find((candidate) => String(candidate.number) === select.dataset.group);
        if (group) {
          setGroupSize(group, select.value);
          renderGroups();
        }
      });
    });
    document.querySelectorAll(".sort-header").forEach((button) => {
      button.addEventListener("click", () => {
        captureCoachInputs();
        captureCampAssignments();
        captureGroupSizes();
        const group = state.groups.find((candidate) => String(candidate.number) === button.dataset.group);
        if (group) {
          setGroupSort(group, button.dataset.sort);
          renderGroups();
        }
      });
    });
    document.querySelectorAll(".kid-row").forEach((row) => {
      row.addEventListener("click", () => {
        const group = state.groups.find((candidate) => String(candidate.number) === row.dataset.group);
        if (!group) return;
        state.selectedKid = { kidId: row.dataset.kidId, groupLetter: group.letter };
        el.moveStatus.textContent = "";
        renderGroups();
      });
    });
    updateMovePanel();
  }

  function refresh() {
    state.selectedKid = null;
    state.columnMap = buildColumnMap(state.rows);
    const missing = ["firstName", "lastName", "dob", "shirt", "shorts", "shoe"].filter((key) => !state.columnMap[key]);
    if (missing.length) {
      setStatus(`This tab is missing expected columns: ${missing.join(", ")}.`, true);
    } else {
      setStatus(`${state.rows.length} rows loaded from "${state.activeSheet}".`);
    }

    state.summaries = {
      SHIRTS: countBy(state.rows, state.columnMap.shirt, SHIRT_SIZES, normalizeSize),
      SHORTS: countBy(state.rows, state.columnMap.shorts, SHORT_SIZES, normalizeShortSize),
      SOCKS: countBy(state.rows, state.columnMap.shoe, SOCK_SIZES, getSockBucket),
    };
    state.kids = buildKids(state.rows, state.columnMap);
    state.groups = makeGroups(state.kids);
    updateCampOptions();

    updatePlayerCount();
    el.groupCount.textContent = state.groups.length;
    el.activeSheetName.textContent = state.activeSheet;
    el.summaryGrid.hidden = false;
    renderDashboard();
    renderGroups();
    setActiveView(state.activeView);
  }

  function loadSheet(sheetName) {
    state.activeSheet = sheetName;
    const sheet = state.workbook.Sheets[sheetName];
    state.rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
    refresh();
  }

  function loadWorkbookFromBuffer(buffer, label) {
    if (!window.XLSX) {
      throw new Error("Excel reader is still loading. Please wait a few seconds and choose the file again.");
    }

    state.workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    state.fileLabel = label || "";
    el.sheetSelect.innerHTML = state.workbook.SheetNames.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
    el.controlsPanel.hidden = false;
    el.resetButton.hidden = false;
    loadSheet(state.workbook.SheetNames[0]);
    setStatus(`${label} loaded. Choose a worksheet tab or download the report.`);
  }

  async function handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    try {
      setStatus("Reading workbook...");
      loadWorkbookFromBuffer(await file.arrayBuffer(), file.name);
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  async function loadSampleFromQuery() {
    const sample = new URLSearchParams(window.location.search).get("sample");
    if (!sample) return;
    try {
      setStatus("Loading sample workbook...");
      const response = await fetch(sample);
      if (!response.ok) throw new Error("Sample workbook could not be loaded.");
      loadWorkbookFromBuffer(await response.arrayBuffer(), sample);
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  function sheetDataForSummary(section, sizes, counts, includeStock) {
    const rows = [[section, ...sizes, "TOTAL"]];
    rows.push(["Customers", ...sizes.map((size) => counts[size] || 0), Object.values(counts).reduce((sum, count) => sum + count, 0)]);
    if (includeStock) {
      rows.push(["Our Stock", ...sizes.map((size) => Number(getStock(section, size)) || ""), ""]);
      rows.push([
        "Difference",
        ...sizes.map((size) => {
          const stock = Number(getStock(section, size));
          return getStock(section, size) === "" ? "" : stock - (counts[size] || 0);
        }),
        "",
      ]);
    } else {
      rows.push(["Our Stock", ...sizes.map(() => ""), ""]);
    }
    rows.push([]);
    return rows;
  }

  function groupSheetRows(group) {
    const assignedCamp = getGroupCamp(group);
    const rows = [
      ["", "", "CAMP", assignedCamp ? `Camp ${assignedCamp}` : "", "", ""],
      ["", "FCB Coach", getCoach(group, "fcb"), "", "Aux. Coach", getCoach(group, "aux")],
      [],
      ["#", "First Name", "Last Name", "Age", "Gender", "Group"],
    ];
    const rosterRows = Array.from({ length: Math.max(getGroupSize(group), group.kids.length) }, (_, index) => sortedGroupKids(group)[index] || null);
    rosterRows.forEach((kid, index) => {
      rows.push([index + 1, kid?.firstName || "", kid?.lastName || "", kid?.age ?? "", kid?.gender || "", groupLabel(group)]);
    });
    return rows;
  }

  function autoWidth(rows) {
    const widths = [];
    rows.forEach((row) => {
      row.forEach((cell, index) => {
        widths[index] = Math.max(widths[index] || 10, String(cell ?? "").length + 2);
      });
    });
    return widths.map((wch) => ({ wch: Math.min(Math.max(wch, 10), 26) }));
  }

  function safeSheetName(name) {
    return String(name).replace(/[\\/?*[\]:]/g, "-").slice(0, 31);
  }

  function exportWorkbook() {
    captureCoachInputs();
    captureCampAssignments();
    captureGroupSizes();
    const report = XLSX.utils.book_new();
    const groupsToExport = activeGroups();
    const dashboardRows = [
      ...sheetDataForSummary("SHIRTS", SHIRT_SIZES, state.summaries.SHIRTS, true),
      ...sheetDataForSummary("SHORTS", SHORT_SIZES, state.summaries.SHORTS, true),
      ...sheetDataForSummary("SOCKS", SOCK_SIZES, state.summaries.SOCKS, true),
    ];
    const dashboardSheet = XLSX.utils.aoa_to_sheet(dashboardRows);
    dashboardSheet["!cols"] = autoWidth(dashboardRows);
    XLSX.utils.book_append_sheet(report, dashboardSheet, "Dashboard");

    const allGroupsRows = [];
    groupsToExport.forEach((group, index) => {
      if (index) allGroupsRows.push([], []);
      allGroupsRows.push(...groupSheetRows(group));
    });
    const groupsSheet = XLSX.utils.aoa_to_sheet(allGroupsRows);
    groupsSheet["!cols"] = autoWidth(allGroupsRows);
    XLSX.utils.book_append_sheet(report, groupsSheet, "Groups");

    groupsToExport.forEach((group) => {
      const rows = groupSheetRows(group);
      const sheet = XLSX.utils.aoa_to_sheet(rows);
      sheet["!cols"] = autoWidth(rows);
      XLSX.utils.book_append_sheet(report, sheet, safeSheetName(`Group ${group.letter}`));
    });

    const selectedCamp =
      el.campSelect.value === "ALL" ? "all-groups" : el.campSelect.value === "UNASSIGNED" ? "unassigned" : `camp-${el.campSelect.value}`;
    const safeCamp = `${state.activeSheet || "camp"}-${selectedCamp}`.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "");
    XLSX.writeFile(report, `${safeCamp}-dashboard-groups.xlsx`);
  }

  function dashboardRows() {
    return [
      ...sheetDataForSummary("SHIRTS", SHIRT_SIZES, state.summaries.SHIRTS, true),
      ...sheetDataForSummary("SHORTS", SHORT_SIZES, state.summaries.SHORTS, true),
      ...sheetDataForSummary("SOCKS", SOCK_SIZES, state.summaries.SOCKS, true),
    ];
  }

  function csvEscape(value) {
    const text = String(value ?? "");
    return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  }

  function downloadDashboardCsv() {
    const csv = dashboardRows().map((row) => row.map(csvEscape).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const safeName = (state.activeSheet || "size-dashboard").replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "");
    link.href = url;
    link.download = `${safeName}-size-dashboard.csv`;
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function groupsRows() {
    captureCoachInputs();
    captureCampAssignments();
    captureGroupSizes();
    const rows = [];
    activeGroups().forEach((group) => {
      if (rows.length) rows.push([]);
      rows.push([`Camp ${getGroupCamp(group) || "Unassigned"}`, `Group ${group.letter}`]);
      rows.push(["FCB Coach", getCoach(group, "fcb"), "Aux. Coach", getCoach(group, "aux")]);
      rows.push(["#", "First Name", "Last Name", "Age", "Gender", "Group"]);
      const rosterRows = Array.from({ length: Math.max(getGroupSize(group), group.kids.length) }, (_, index) => sortedGroupKids(group)[index] || null);
      rosterRows.forEach((kid, index) => {
        rows.push([index + 1, kid?.firstName || "", kid?.lastName || "", kid?.age ?? "", kid?.gender || "", groupLabel(group)]);
      });
    });
    return rows;
  }

  function downloadGroupsCsv() {
    const csv = groupsRows().map((row) => row.map(csvEscape).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const selectedCamp =
      el.campSelect.value === "ALL" ? "all-groups" : el.campSelect.value === "UNASSIGNED" ? "unassigned" : `camp-${el.campSelect.value}`;
    const safeName = `${state.activeSheet || "groups"}-${selectedCamp}`.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "");
    link.href = url;
    link.download = `${safeName}.csv`;
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function printDashboard() {
    document.body.classList.remove("print-groups");
    document.body.classList.add("print-dashboard");
    window.print();
  }

  function printGroups() {
    captureCoachInputs();
    captureCampAssignments();
    captureGroupSizes();
    document.body.classList.remove("print-dashboard");
    document.body.classList.add("print-groups");
    window.print();
  }

  function reset() {
    state.workbook = null;
    state.rows = [];
    state.groups = [];
    state.fileLabel = "";
    state.selectedKid = null;
    clearSelectedKid();
    el.fileInput.value = "";
    el.sheetSelect.innerHTML = "";
    el.controlsPanel.hidden = true;
    el.summaryGrid.hidden = true;
    el.dashboardPanel.hidden = true;
    el.rosterPanel.hidden = true;
    el.resetButton.hidden = true;
    setActiveView(state.activeView);
    setStatus("");
  }

  el.fileInput.addEventListener("change", handleFile);
  el.sheetSelect.addEventListener("change", () => loadSheet(el.sheetSelect.value));
  el.sortSelect.addEventListener("change", () => {
    captureCoachInputs();
    captureCampAssignments();
    captureGroupSizes();
    state.groups = makeGroups(state.kids);
    state.selectedKid = null;
    updateCampOptions();
    renderGroups();
  });
  el.campSelect.addEventListener("change", () => {
    captureCoachInputs();
    captureCampAssignments();
    captureGroupSizes();
    renderGroups();
  });
  el.exportButton.addEventListener("click", exportWorkbook);
  el.printDashboardButton.addEventListener("click", printDashboard);
  el.csvDashboardButton.addEventListener("click", downloadDashboardCsv);
  el.printGroupsButton.addEventListener("click", printGroups);
  el.csvGroupsButton.addEventListener("click", downloadGroupsCsv);
  el.moveKidButton.addEventListener("click", moveSelectedKid);
  el.deleteSelectedKidButton.addEventListener("click", deleteSelectedPlayer);
  el.clearSelectedKidButton.addEventListener("click", () => {
    clearSelectedKid();
    renderGroups();
  });
  el.resetButton.addEventListener("click", reset);
  el.sizeViewButton.addEventListener("click", () => setActiveView("size"));
  el.groupsViewButton.addEventListener("click", () => setActiveView("groups"));
  window.addEventListener("afterprint", () => {
    document.body.classList.remove("print-dashboard", "print-groups");
  });
  loadSampleFromQuery();
  setActiveView(state.activeView);
})();
