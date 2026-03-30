const STORAGE_KEY = "fakturix_ch_erfassung_data_v2";
const state = loadState();
let editingCustomerId = null;
let editingEmployeeId = null;
let editingItemId = null;
let editingUnitId = null;
let editingProductTypeId = null;
let editingEntryId = null;
let customerSuggestionMap = new Map();
let employeeSuggestionMap = new Map();
let itemSuggestionMap = new Map();
let splitEmployeeQuantityMap = new Map();
ensureCleaningServiceExists();
enforceFixedUnitsCatalog();

const tabs = [...document.querySelectorAll(".tabs-main .tab")];
const panels = [...document.querySelectorAll(".panel")];

const customerForm = document.getElementById("customerForm");
const employeeForm = document.getElementById("employeeForm");
const itemForm = document.getElementById("itemForm");
const unitForm = document.getElementById("unitForm");
const productTypeForm = document.getElementById("productTypeForm");
const entryForm = document.getElementById("entryForm");
const settingsForm = document.getElementById("settingsForm");

const customerSubmitBtn = document.getElementById("customerSubmitBtn");
const customerCancelBtn = document.getElementById("customerCancelBtn");
const employeeSubmitBtn = document.getElementById("employeeSubmitBtn");
const employeeCancelBtn = document.getElementById("employeeCancelBtn");
const itemSubmitBtn = document.getElementById("itemSubmitBtn");
const itemCancelBtn = document.getElementById("itemCancelBtn");
const unitSubmitBtn = document.getElementById("unitSubmitBtn");
const unitCancelBtn = document.getElementById("unitCancelBtn");
const productTypeSubmitBtn = document.getElementById("productTypeSubmitBtn");
const productTypeCancelBtn = document.getElementById("productTypeCancelBtn");
const entrySubmitBtn = document.getElementById("entrySubmitBtn");
const entryCancelBtn = document.getElementById("entryCancelBtn");

const customerList = document.getElementById("customerList");
const deletedCustomerList = document.getElementById("deletedCustomerList");
const deletedCustomerSearch = document.getElementById("deletedCustomerSearch");
const deletedCustomerSearchClear = document.getElementById("deletedCustomerSearchClear");
const deletedCustomerSummary = document.getElementById("deletedCustomerSummary");
const customerSearch = document.getElementById("customerSearch");
const customerSearchClear = document.getElementById("customerSearchClear");
const employeeList = document.getElementById("employeeList");
const deletedEmployeeList = document.getElementById("deletedEmployeeList");
const deletedEmployeeSearch = document.getElementById("deletedEmployeeSearch");
const deletedEmployeeSearchClear = document.getElementById("deletedEmployeeSearchClear");
const deletedEmployeeSummary = document.getElementById("deletedEmployeeSummary");
const employeeSearch = document.getElementById("employeeSearch");
const employeeSearchClear = document.getElementById("employeeSearchClear");
const itemList = document.getElementById("itemList");
const deletedItemList = document.getElementById("deletedItemList");
const deletedItemSearch = document.getElementById("deletedItemSearch");
const deletedItemSearchClear = document.getElementById("deletedItemSearchClear");
const deletedItemSummary = document.getElementById("deletedItemSummary");
const itemSearch = document.getElementById("itemSearch");
const itemSearchClear = document.getElementById("itemSearchClear");
const unitList = document.getElementById("unitList");
const deletedUnitList = document.getElementById("deletedUnitList");
const deletedUnitSearch = document.getElementById("deletedUnitSearch");
const deletedUnitSearchClear = document.getElementById("deletedUnitSearchClear");
const deletedUnitSummary = document.getElementById("deletedUnitSummary");
const unitSearch = document.getElementById("unitSearch");
const unitSearchClear = document.getElementById("unitSearchClear");
const productTypeList = document.getElementById("productTypeList");
const deletedProductTypeList = document.getElementById("deletedProductTypeList");
const deletedProductTypeSearch = document.getElementById("deletedProductTypeSearch");
const deletedProductTypeSearchClear = document.getElementById("deletedProductTypeSearchClear");
const deletedProductTypeSummary = document.getElementById("deletedProductTypeSummary");
const productTypeSearch = document.getElementById("productTypeSearch");
const productTypeSearchClear = document.getElementById("productTypeSearchClear");
const entryListHeader = document.getElementById("entryListHeader");
const entryList = document.getElementById("entryList");
const entryListCustomerSearch = document.getElementById("entryListCustomerSearch");
const entryListCustomerClear = document.getElementById("entryListCustomerClear");

const entryCustomer = document.getElementById("entryCustomer");
const entryCustomerDisplay = document.getElementById("entryCustomerDisplay");
const entryCustomerSearch = document.getElementById("entryCustomerSearch");
const customerSuggestions = document.getElementById("customerSuggestions");
const entryEmployee = document.getElementById("entryEmployee");
const entryEmployeeDisplay = document.getElementById("entryEmployeeDisplay");
const entryEmployeeSearch = document.getElementById("entryEmployeeSearch");
const employeeSuggestions = document.getElementById("employeeSuggestions");
const entrySplitEmployees = document.getElementById("entrySplitEmployees");
const entryEmployeeSearchRow = document.getElementById("entryEmployeeSearchRow") || (entryEmployeeSearch ? entryEmployeeSearch.closest("label") : null);
const entryEmployeeSelectedRow = document.getElementById("entryEmployeeSelectedRow");
const entryEmployeesMultiRow = document.getElementById("entryEmployeesMultiRow");
const entryEmployeesMultiList = document.getElementById("entryEmployeesMultiList");
const entryEmployeesSplitDetails = document.getElementById("entryEmployeesSplitDetails");
const entryEmployeesSplitHint = document.getElementById("entryEmployeesSplitHint");
const entryItem = document.getElementById("entryItem");
const entryItemDisplay = document.getElementById("entryItemDisplay");
const entryItemSearch = document.getElementById("entryItemSearch");
const itemSuggestions = document.getElementById("itemSuggestions");
const entryDate = document.getElementById("entryDate");
const reportMonth = document.getElementById("reportMonth");
const reportCalendar = document.getElementById("reportCalendar");
const reportCalendarProducts = document.getElementById("reportCalendarProducts");
const entryQuantityInput = entryForm.querySelector('input[name="quantity"]');
const itemTypeSelect = document.getElementById("itemType");
const itemUnitSelect = document.getElementById("itemUnit");

const exportBtn = document.getElementById("exportBtn");
const exportStatus = document.getElementById("exportStatus");
const exportCleanupResult = document.getElementById("exportCleanupResult");
const exportCleanupList = document.getElementById("exportCleanupList");
const importBtn = document.getElementById("importBtn");
const importFile = document.getElementById("importFile");
const importStatus = document.getElementById("importStatus");
const currencyCode = document.getElementById("currencyCode");
const retentionMonths = document.getElementById("retentionMonths");
const settingsStatus = document.getElementById("settingsStatus");
const resetConfirmCheckbox = document.getElementById("resetConfirmCheckbox");
const resetDataBtn = document.getElementById("resetDataBtn");
const flashMessage = document.getElementById("flashMessage");
let flashMessageTimeoutId = null;

setDefaultDate();
wireTabs();
wireSubTabs();
wireForms();
wireCrudActions();
wireExport();
wireImport();
wireSearch();
wireCustomerSearch();
wireEmployeeSearch();
wireItemSearch();
wireProductTypeSearch();
wireDeletedMasterDataSearch();
wireSettings();
wireReportOverview();
wireResetConfirmation();
wireRequiredFieldStates();
renderAll();

function loadState() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) return normalizeStateData({ customers: [], employees: [], items: [], entries: [], units: [], productTypes: [], settings: { retentionMonths: 12 } });
    const parsed = JSON.parse(saved);
    return normalizeStateData(parsed);
  } catch {
    return normalizeStateData({ customers: [], employees: [], items: [], entries: [], units: [], productTypes: [], settings: { retentionMonths: 12 } });
  }
}


function enforceFixedUnitsCatalog() {
  state.units = [
    { id: "unit-h", name: "h" },
    { id: "unit-stk", name: "Stk" }
  ];

  state.items.forEach((item) => {
    const normalizedId = item.unitId === "unit-stk" ? "unit-stk" : "unit-h";
    item.unitId = normalizedId;
    item.unit = normalizedId === "unit-stk" ? "Stk" : "h";
  });
}
function normalizeStateData(parsed) {
  const customers = (Array.isArray(parsed?.customers) ? parsed.customers : []).map((c) => ({
    id: c?.id || generateId(),
    company: String(c?.company || ""),
    firstName: String(c?.firstName || ""),
    lastName: String(c?.lastName || ""),
    street: String(c?.street || ""),
    zip: String(c?.zip || ""),
    city: String(c?.city || ""),
    phone: String(c?.phone || ""),
    email: String(c?.email || ""),
    deleted: c?.deleted === true || String(c?.deleted || "").toLowerCase() === "true"
  }));
  const employees = (Array.isArray(parsed?.employees) ? parsed.employees : []).map((e) => ({
    id: e?.id || generateId(),
    firstName: String(e?.firstName || ""),
    lastName: String(e?.lastName || ""),
    street: String(e?.street || ""),
    zip: String(e?.zip || ""),
    city: String(e?.city || ""),
    phone: String(e?.phone || ""),
    email: String(e?.email || ""),
    iban: String(e?.iban || ""),
    hourlyWage: parseAmount(e?.hourlyWage),
    deleted: e?.deleted === true || String(e?.deleted || "").toLowerCase() === "true"
  }));

  const units = [
    { id: "unit-h", name: "h" },
    { id: "unit-stk", name: "Stk" }
  ];
  const productTypes = normalizeNamedCatalog(Array.isArray(parsed?.productTypes) ? parsed.productTypes : [], "");  ensureCatalogContains(productTypes, "Dienstleistung");
  ensureCatalogContains(productTypes, "Produkt");

  const items = (Array.isArray(parsed?.items) ? parsed.items : []).map((item) => {
    const normalized = {
      id: item?.id || generateId(),
      name: String(item?.name || ""),
      price: Number(item?.price) || 0,
      costPrice: parseAmount(item?.costPrice ?? item?.purchasePrice ?? item?.einstandspreis),
      deleted: item?.deleted === true || String(item?.deleted || "").toLowerCase() === "true",
      typeId: String(item?.typeId || ""),
      unitId: String(item?.unitId || ""),
      type: String(item?.type || ""),
      unit: String(item?.unit || "")
    };

    if (!normalized.typeId || !productTypes.some((t) => t.id === normalized.typeId)) {
      const fallbackType = normalizeCatalogName(normalized.type) || "Dienstleistung";
      normalized.typeId = ensureCatalogContains(productTypes, fallbackType).id;
    }
    if (!normalized.unitId || !units.some((u) => u.id === normalized.unitId)) {
      normalized.unitId = "unit-h";
    }

    normalized.type = getCatalogNameById(productTypes, normalized.typeId) || normalizeCatalogName(normalized.type) || "Dienstleistung";
    normalized.unit = normalized.unitId === "unit-stk" ? "Stk" : "h";
    return normalized;
  });

  const entries = (Array.isArray(parsed?.entries) ? parsed.entries : []).map((entry) => {
    const itemId = String(entry?.itemId || "");
    const item = items.find((i) => i.id === itemId);
    return {
      ...entry,
      id: entry?.id || generateId(),
      customerId: String(entry?.customerId || ""),
      employeeId: String(entry?.employeeId || ""),
      itemId,
      date: String(entry?.date || ""),
      quantity: Number(entry?.quantity) > 0 ? Number(entry.quantity) : 1,
      note: String(entry?.note || ""),
      unitPrice: parseAmount(entry?.unitPrice),
      costPrice: normalizeEntryCostPrice(entry, item)
    };
  });

  return {
    customers,
    employees,
    items,
    entries,
    units,
    productTypes,
    settings: {
      currency: normalizeCurrency(parsed?.settings?.currency),
      retentionMonths: normalizeMonths(parsed?.settings?.retentionMonths)
    }
  };
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function showFlashMessage(message) {
  if (!flashMessage) return;
  flashMessage.textContent = String(message || "");
  flashMessage.hidden = false;
  if (flashMessageTimeoutId) clearTimeout(flashMessageTimeoutId);
  flashMessageTimeoutId = setTimeout(() => {
    flashMessage.hidden = true;
  }, 1800);
}

function wireTabs() {
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      const target = tab.dataset.tab;
      tabs.forEach((t) => {
        const isActive = t === tab;
        t.classList.toggle("active", isActive);
        t.setAttribute("aria-selected", String(isActive));
      });

      panels.forEach((panel) => {
        const isActive = panel.id === target;
        panel.classList.toggle("active", isActive);
        panel.hidden = !isActive;
      });

      if (target === "erfassung") {
        setDefaultDate();
        resetEntryCustomerSelection();
        renderEntries();
      }

      if (target === "auswertung") {
        setDefaultReportMonth();
        renderReportOverview();
        renderProductReportOverview();
      }
    });
  });
}

function wireSubTabs() {
  initSubTabGroup(".panel#stammdaten");
  initSubTabGroup(".panel#wartung");
  initSubTabGroup(".panel#auswertung");
}

function initSubTabGroup(panelSelector) {
  const root = document.querySelector(panelSelector);
  if (!root) return;
  const subTabs = [...root.querySelectorAll(".subtab")];
  const subPanels = [...root.querySelectorAll(".subpanel")];
  subTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      const target = tab.dataset.subtab;
      subTabs.forEach((t) => {
        const active = t === tab;
        t.classList.toggle("active", active);
        t.setAttribute("aria-selected", String(active));
      });
      subPanels.forEach((p) => {
        const active = p.id === target;
        p.classList.toggle("active", active);
        p.hidden = !active;
      });
    });
  });
}

function wireForms() {
  customerForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const data = formToObject(customerForm);
    const isEditingCustomer = Boolean(editingCustomerId);

    const customer = {
      id: editingCustomerId || generateId(),
      company: data.company?.trim() || "",
      firstName: data.firstName.trim(),
      lastName: data.lastName.trim(),
      street: data.street?.trim() || "",
      zip: data.zip.trim(),
      city: data.city.trim(),
      phone: data.phone?.trim() || "",
      email: data.email?.trim() || "",
      deleted: false
    };

    if (editingCustomerId) {
      const index = state.customers.findIndex((c) => c.id === editingCustomerId);
      if (index >= 0) state.customers[index] = customer;
    } else {
      state.customers.push(customer);
    }

    resetCustomerForm();
    saveState();
    renderAll();
    showFlashMessage(isEditingCustomer ? "Kunde aktualisiert." : "Kunde gespeichert.");
  });

  customerCancelBtn.addEventListener("click", () => {
    resetCustomerForm();
  });

  employeeForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const data = formToObject(employeeForm);
    const isEditingEmployee = Boolean(editingEmployeeId);

    const employee = {
      id: editingEmployeeId || generateId(),
      firstName: data.firstName.trim(),
      lastName: data.lastName.trim(),
      street: data.street?.trim() || "",
      zip: data.zip.trim(),
      city: data.city.trim(),
      phone: data.phone?.trim() || "",
      email: data.email?.trim() || "",
      deleted: false,
      iban: data.iban?.trim() || "",
      hourlyWage: parseAmount(data.hourlyWage),
      deleted: false
    };

    if (editingEmployeeId) {
      const index = state.employees.findIndex((e) => e.id === editingEmployeeId);
      if (index >= 0) state.employees[index] = employee;
    } else {
      state.employees.push(employee);
    }

    resetEmployeeForm();
    saveState();
    renderAll();
    showFlashMessage(isEditingEmployee ? "Mitarbeiter aktualisiert." : "Mitarbeiter gespeichert.");
  });

  employeeCancelBtn.addEventListener("click", () => {
    resetEmployeeForm();
  });

  itemForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const data = formToObject(itemForm);
    const isEditingItem = Boolean(editingItemId);
    const typeId = String(data.typeId || "");
    const unitId = String(data.unitId || "");
    const typeName = getCatalogNameById(state.productTypes, typeId) || "Dienstleistung";
    const unitName = getCatalogNameById(state.units, unitId) || "h";

    const item = {
      id: editingItemId || generateId(),
      name: data.name.trim(),
      typeId,
      unitId,
      type: typeName,
      unit: unitName,
      price: Number(data.price),
      costPrice: parseAmount(data.costPrice),
      deleted: false
    };

    if (editingItemId) {
      const index = state.items.findIndex((i) => i.id === editingItemId);
      if (index >= 0) state.items[index] = item;
    } else {
      state.items.push(item);
    }

    resetItemForm();
    saveState();
    renderAll();
    showFlashMessage(isEditingItem ? "Produkt/Dienstleistung aktualisiert." : "Produkt/Dienstleistung gespeichert.");
  });

  itemCancelBtn.addEventListener("click", () => {
    resetItemForm();
  });

  unitForm?.addEventListener("submit", (event) => {
    event.preventDefault();
    const data = formToObject(unitForm);
    const name = normalizeCatalogName(data.name);
    if (!name) return;
    const isEditing = Boolean(editingUnitId);

    if (editingUnitId) {
      const index = state.units.findIndex((u) => u.id === editingUnitId);
      if (index >= 0) state.units[index] = { ...state.units[index], name };
      state.items.forEach((item) => {
        if (item.unitId === editingUnitId) item.unit = name;
      });
    } else {
      const existing = findCatalogByName(state.units, name);
      if (existing) {
        unitForm.name.value = existing.name;
        refreshRequiredFieldStates();
        return;
      }
      state.units.push({ id: generateId(), name });
    }

    resetUnitForm();
    saveState();
    renderAll();
    showFlashMessage(isEditing ? "Einheit aktualisiert." : "Einheit gespeichert.");
  });

  unitCancelBtn?.addEventListener("click", () => {
    resetUnitForm();
  });

  productTypeForm?.addEventListener("submit", (event) => {
    event.preventDefault();
    const data = formToObject(productTypeForm);
    const name = normalizeCatalogName(data.name);
    if (!name) return;
    const isEditing = Boolean(editingProductTypeId);

    if (editingProductTypeId) {
      const index = state.productTypes.findIndex((t) => t.id === editingProductTypeId);
      if (index >= 0) state.productTypes[index] = { ...state.productTypes[index], name };
      state.items.forEach((item) => {
        if (item.typeId === editingProductTypeId) item.type = name;
      });
    } else {
      const existing = findCatalogByName(state.productTypes, name);
      if (existing) {
        productTypeForm.name.value = existing.name;
        refreshRequiredFieldStates();
        return;
      }
      state.productTypes.push({ id: generateId(), name });
    }

    resetProductTypeForm();
    saveState();
    renderAll();
    showFlashMessage(isEditing ? "Produkttyp aktualisiert." : "Produkttyp gespeichert.");
  });

  productTypeCancelBtn?.addEventListener("click", () => {
    resetProductTypeForm();
  });

  entryCancelBtn?.addEventListener("click", () => {
    resetEntryForm();
  });

  entrySplitEmployees?.addEventListener("change", () => {
    syncEntryEmployeeAssignmentMode({ resetSelection: true });
    refreshRequiredFieldStates();
  });

  entryEmployeesMultiList?.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (!target.classList.contains("entry-employee-multi")) return;
    splitEmployeeQuantityMap = new Map();
    updateEntryMultiEmployeeSelectionState();
    renderSplitQuantityInputs();
    syncEntrySubmitButtonState();
  });

  entryDate.addEventListener("change", () => {
    renderEntries();
  });

  entryDate.addEventListener("input", () => {
    renderEntries();
  });

  entryItemSearch?.addEventListener("input", () => {
    updateEntryQuantityConstraints();
  });

  entryQuantityInput?.addEventListener("keydown", (event) => {
    if (!isPieceUnitSelected()) return;
    if ([".", ",", "e", "E", "+", "-"].includes(event.key)) {
      event.preventDefault();
    }
  });

  entryQuantityInput?.addEventListener("input", () => {
    if (!isPieceUnitSelected()) return;
    const raw = String(entryQuantityInput.value || "");
    const digitsOnly = raw.replace(/\D+/g, "");
    entryQuantityInput.value = digitsOnly;
  });

  entryQuantityInput?.addEventListener("input", () => {
    if (isSplitEmployeeAssignmentEnabled()) {
      syncSplitAllocationValidation();
      syncEntrySubmitButtonState();
    }
  });

  entryEmployeesSplitDetails?.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    if (!target.classList.contains("entry-employee-split-qty")) return;
    const employeeId = String(target.dataset.employeeId || "").trim();
    if (!employeeId) return;
    const normalized = parseSplitQuantityInputValue(target.value, isPieceUnitSelected());
    splitEmployeeQuantityMap.set(employeeId, normalized);
    target.value = normalized;
    syncSplitAllocationValidation();
    syncEntrySubmitButtonState();
  });

  entryForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const data = formToObject(entryForm);
    const isEditingEntry = Boolean(editingEntryId);
    const useSplitAcrossEmployees = isSplitEmployeeAssignmentEnabled() && !isEditingEntry;

    if (!data.customerId) {
      alert("Bitte zuerst einen Kunden auswählen.");
      return;
    }

    let assignedEmployeeIds = [];
    if (useSplitAcrossEmployees) {
      assignedEmployeeIds = getSelectedSplitEmployeeIds();
      if (assignedEmployeeIds.length < 2) {
        alert("Bitte mindestens zwei Mitarbeiter auswählen.");
        return;
      }
    } else {
      if (!data.employeeId) {
        alert("Bitte zuerst einen Mitarbeiter auswählen.");
        return;
      }
      assignedEmployeeIds = [data.employeeId];
    }

    if (!data.itemId) {
      alert("Bitte eine Dienstleistung / ein Produkt auswählen.");
      return;
    }

    const item = state.items.find((i) => i.id === data.itemId);
    const quantity = Number(data.quantity);
    if (!Number.isFinite(quantity) || quantity <= 0) {
      alert("Bitte eine gueltige Menge eingeben.");
      return;
    }

    const isPieceUnit = getItemUnitName(item) === "Stk";
    if (isPieceUnit && !Number.isInteger(quantity)) {
      alert("Bei Einheit 'Stk' sind nur ganze Zahlen erlaubt.");
      entryForm.quantity.focus();
      return;
    }

    const normalizedQuantity = isPieceUnit ? Math.trunc(quantity) : quantity;
    const splitAssignments = useSplitAcrossEmployees
      ? getSplitAssignmentsFromInputs(assignedEmployeeIds, normalizedQuantity, isPieceUnit)
      : { ok: true, assignments: [{ employeeId: assignedEmployeeIds[0], quantity: normalizedQuantity }], sum: normalizedQuantity };

    if (!splitAssignments.ok) {
      alert(splitAssignments.message || "Die Aufteilung auf Mitarbeiter ist ungültig.");
      entryForm.quantity.focus();
      return;
    }

    const existingEntry = editingEntryId ? state.entries.find((e) => e.id === editingEntryId) : null;
    const resolvedUnitPrice = resolveEntryUnitPrice(item);
    const resolvedCostPrice = resolveEntryCostPrice(item);
    const unitPrice = existingEntry && existingEntry.itemId === data.itemId && existingEntry.unitPrice != null
      ? existingEntry.unitPrice
      : resolvedUnitPrice;
    const costPrice = existingEntry && existingEntry.itemId === data.itemId && existingEntry.costPrice != null
      ? existingEntry.costPrice
      : resolvedCostPrice;

    let mergedCount = 0;
    let createdCount = 0;

    if (isEditingEntry) {
      const entry = {
        id: editingEntryId || generateId(),
        customerId: data.customerId,
        employeeId: data.employeeId,
        itemId: data.itemId,
        date: data.date,
        quantity: normalizedQuantity,
        note: data.note?.trim() || "",
        unitPrice,
        costPrice
      };

      const index = state.entries.findIndex((e) => e.id === editingEntryId);
      if (index >= 0) state.entries[index] = entry;
      else state.entries.push(entry);
    } else {
      splitAssignments.assignments.forEach(({ employeeId, quantity: employeeQuantity }) => {
        const entry = {
          id: generateId(),
          customerId: data.customerId,
          employeeId,
          itemId: data.itemId,
          date: data.date,
          quantity: employeeQuantity,
          note: data.note?.trim() || "",
          unitPrice,
          costPrice
        };
        const merged = mergeEntryIntoState(entry, item);
        if (merged) mergedCount += 1;
        else createdCount += 1;
      });
    }

    saveState();
    resetEntryForm();
    updateEntryQuantityConstraints();
    refreshRequiredFieldStates();
    renderEntries();

    if (isEditingEntry) {
      showFlashMessage("Erfassung aktualisiert");
    } else if (useSplitAcrossEmployees) {
      const details = [];
      details.push(`${assignedEmployeeIds.length} Mitarbeiter`);
      details.push(`${createdCount} neu`);
      if (mergedCount > 0) details.push(`${mergedCount} zusammengeführt`);
      showFlashMessage(`Erfassung aufgeteilt (${details.join(", ")})`);
    } else if (mergedCount > 0) {
      showFlashMessage("+ Menge zum bestehenden Eintrag addiert");
    } else {
      showFlashMessage("Erfassung gespeichert");
    }
  });
}
function wireCrudActions() {
  customerList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action][data-id]");
    if (!button) return;
    const customerId = button.dataset.id;
    const action = button.dataset.action;

    if (action === "edit") {
      editCustomer(customerId);
      return;
    }

    if (action === "delete") {
      deleteCustomer(customerId);
    }
    });

  deletedCustomerList?.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action][data-id]");
    if (!button) return;
    if (button.dataset.action === "restore") {
      restoreCustomer(button.dataset.id);
    }
  });

    employeeList.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-action][data-id]");
      if (!button) return;
      const employeeId = button.dataset.id;
      const action = button.dataset.action;

      if (action === "edit") {
        editEmployee(employeeId);
        return;
      }

      if (action === "delete") {
        deleteEmployee(employeeId);
      }
    });

    deletedEmployeeList?.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-action][data-id]");
      if (!button) return;
      if (button.dataset.action === "restore") {
        restoreEmployee(button.dataset.id);
      }
    });

    itemList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action][data-id]");
    if (!button) return;
    const itemId = button.dataset.id;
    const action = button.dataset.action;

    if (action === "edit") {
      editItem(itemId);
      return;
    }

    if (action === "delete") {
      deleteItem(itemId);
    }
  });

  deletedItemList?.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action][data-id]");
    if (!button) return;
    if (button.dataset.action === "restore") {
      restoreItem(button.dataset.id);
    }
  });

  unitList?.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action][data-id]");
    if (!button) return;
    const unitId = button.dataset.id;
    const action = button.dataset.action;

    if (action === "edit") {
      editUnit(unitId);
      return;
    }
    if (action === "delete") {
      deleteUnit(unitId);
    }
  });

  deletedUnitList?.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action][data-id]");
    if (!button) return;
    if (button.dataset.action === "restore") {
      restoreUnit(button.dataset.id);
    }
  });

  productTypeList?.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action][data-id]");
    if (!button) return;
    const typeId = button.dataset.id;
    const action = button.dataset.action;

    if (action === "edit") {
      editProductType(typeId);
      return;
    }
    if (action === "delete") {
      deleteProductType(typeId);
    }
  });

  deletedProductTypeList?.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action][data-id]");
    if (!button) return;
    if (button.dataset.action === "restore") {
      restoreProductType(button.dataset.id);
    }
  });

  entryList.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-action][data-id]");
      if (!button) return;
      const action = button.dataset.action;
      if (action === "edit-entry") {
        editEntry(button.dataset.id);
        return;
      }
      if (action === "delete-entry") {
        deleteEntry(button.dataset.id);
      }
    });
}

function wireExport() {
  exportBtn.addEventListener("click", async () => {
    const cleanup = pruneOldEntriesForExport();
    const removedCount = cleanup.entries.length;
    const removedItemCount = cleanup.items.length;
    const removedUnitCount = cleanup.units.length;
    const removedTypeCount = cleanup.productTypes.length;
    const removedCustomerCount = cleanup.customers.length;
    const removedEmployeeCount = cleanup.employees.length;
    const retentionMonthsCount = normalizeMonths(state?.settings?.retentionMonths);
            const payload = {
        app: "Fakturix CH Leistungserfassung",
        exportedAt: new Date().toISOString(),
        customers: state.customers,
        employees: state.employees,
        units: state.units,
        productTypes: state.productTypes,
        items: state.items.map((item) => ({
          ...item,
          type: getItemTypeName(item),
          unit: getItemUnitName(item)
        })),
        entries: state.entries,
        settings: state.settings
      };

    const formattedDate = new Date().toISOString().slice(0, 10);
    const fileName = `fakturix-ch-erfassung-export-${formattedDate}.json`;
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });

    try {
      if (navigator.share) {
        let canShareFiles = false;
        if (typeof navigator.canShare === "function" && typeof File !== "undefined") {
          try {
            canShareFiles = navigator.canShare({
              files: [new File([blob], fileName, { type: "application/json" })]
            });
          } catch {
            canShareFiles = false;
          }
        }

        if (canShareFiles) {
          try {
            const file = new File([blob], fileName, { type: "application/json" });
              await navigator.share({
                title: "Fakturix CH Leistungserfassung Export",
                text: "Export der mobilen Erfassungsdaten",
                files: [file]
              });
                exportStatus.textContent = `Export über Teilen-Menü erfolgreich. Alte Erfassungen älter als ${retentionMonthsCount} Monat(e) gelöscht: ${removedCount}. Gelöschte Stammdaten ohne Abhängigkeiten entfernt: Kunden ${removedCustomerCount}, Mitarbeiter ${removedEmployeeCount}, Produkte ${removedItemCount}, Einheiten ${removedUnitCount}, Produkttypen ${removedTypeCount}.`;
                showFlashMessage("Export war erfolgreich");
              renderExportCleanupPreview();
              return;
            } catch (shareError) {
              console.warn("Share fehlgeschlagen, wechsle auf Download-Fallback.", shareError);
            }
          }
        }

        downloadBlob(blob, fileName);
        exportStatus.textContent = `Datei "${fileName}" wurde heruntergeladen oder zum Speichern angeboten. Alte Erfassungen älter als ${retentionMonthsCount} Monat(e) gelöscht: ${removedCount}. Gelöschte Stammdaten ohne Abhängigkeiten entfernt: Kunden ${removedCustomerCount}, Mitarbeiter ${removedEmployeeCount}, Produkte ${removedItemCount}, Einheiten ${removedUnitCount}, Produkttypen ${removedTypeCount}.`;
        showFlashMessage("Export war erfolgreich");
        renderExportCleanupPreview();
      } catch (error) {
        exportStatus.textContent = "Export wurde abgebrochen oder ist fehlgeschlagen.";
        renderExportCleanupPreview();
        console.error(error);
      }
  });
}

function wireSettings() {
  settingsForm.addEventListener("submit", (event) => {
    event.preventDefault();
    state.settings.currency = normalizeCurrency(currencyCode?.value);
    state.settings.retentionMonths = normalizeMonths(retentionMonths.value);
    if (currencyCode) currencyCode.value = state.settings.currency;
    saveState();
    settingsStatus.textContent = `Gespeichert: Währung ${state.settings.currency}. Erfassungen werden ${state.settings.retentionMonths} Monate aufbewahrt.`;
    showFlashMessage("Einstellungen gespeichert");
    renderAll();
  });
}

function wireResetConfirmation() {
  if (!resetDataBtn || !resetConfirmCheckbox) return;

  const syncState = () => {
    resetDataBtn.disabled = !resetConfirmCheckbox.checked;
  };

  ["change", "input", "click"].forEach((eventName) => {
    resetConfirmCheckbox.addEventListener(eventName, syncState);
  });
  syncState();

  resetDataBtn.addEventListener("click", (event) => {
    event.preventDefault();
    if (resetDataBtn.disabled) return;

    if (!confirm("Alle Daten in der mobilen App löschen?")) return;
    if (!confirm("Sind Sie wirklich sicher?")) return;

    state.customers = [];
    state.employees = [];
    state.items = [];
    state.entries = [];
    state.units = [];
    state.productTypes = [];
    state.settings = { currency: "CHF", retentionMonths: 12 };
    ensureCleaningServiceExists();
    enforceFixedUnitsCatalog();

    resetCustomerForm();
    resetEmployeeForm();
    resetItemForm();
    resetUnitForm();
    resetProductTypeForm();
    resetEntryForm();
    saveState();
    renderAll();

    if (importFile) importFile.value = "";
    importStatus.textContent = "";
    exportStatus.textContent = "";
    settingsStatus.textContent = "Alle Daten wurden gelöscht.";
    showFlashMessage("Alle Daten wurden gelöscht");

    resetConfirmCheckbox.checked = false;
    syncState();
  });
}

function wireSearch() {
  const syncCustomerSelectionFromSearch = () => {
    let matchedId = getCustomerIdFromSuggestion(entryCustomerSearch.value);
    if (!matchedId) {
      renderCustomerSuggestions(entryCustomerSearch.value);
      matchedId = getCustomerIdFromSuggestion(entryCustomerSearch.value);
    }
    if (matchedId) {
      setSelectedCustomer(matchedId);
    } else if (!entryCustomerSearch.value.trim()) {
      clearSelectedCustomer();
    }
  };

  entryCustomerSearch.addEventListener("input", syncCustomerSelectionFromSearch);
  entryCustomerSearch.addEventListener("change", syncCustomerSelectionFromSearch);

  entryCustomerSearch.addEventListener("focus", () => {
    renderCustomerSuggestions(entryCustomerSearch.value);
  });
  const syncEmployeeSelectionFromSearch = () => {
    if (isSplitEmployeeAssignmentEnabled()) {
      clearSelectedEmployee();
      return;
    }
    let matchedId = getEmployeeIdFromSuggestion(entryEmployeeSearch.value);
    if (!matchedId) {
      renderEmployeeSuggestions(entryEmployeeSearch.value);
      matchedId = getEmployeeIdFromSuggestion(entryEmployeeSearch.value);
    }
    if (matchedId) {
      setSelectedEmployee(matchedId);
    } else if (!entryEmployeeSearch.value.trim()) {
      clearSelectedEmployee();
    }
  };

  entryEmployeeSearch.addEventListener("input", syncEmployeeSelectionFromSearch);
  entryEmployeeSearch.addEventListener("change", syncEmployeeSelectionFromSearch);

  entryEmployeeSearch.addEventListener("focus", () => {
    renderEmployeeSuggestions(entryEmployeeSearch.value);
  });

  const syncItemSelectionFromSearch = () => {
    let matchedId = getItemIdFromSuggestion(entryItemSearch.value);
    if (!matchedId) {
      renderItemSuggestions(entryItemSearch.value);
      matchedId = getItemIdFromSuggestion(entryItemSearch.value);
    }
    if (matchedId) {
      setSelectedItem(matchedId);
    } else if (!entryItemSearch.value.trim()) {
      clearSelectedItem();
    }
  };

  entryItemSearch.addEventListener("input", syncItemSelectionFromSearch);
  entryItemSearch.addEventListener("change", syncItemSelectionFromSearch);

  entryItemSearch.addEventListener("focus", () => {
    renderItemSuggestions(entryItemSearch.value);
  });

  entryListCustomerSearch?.addEventListener("input", () => {
    renderEntries();
  });

  entryListCustomerClear?.addEventListener("click", () => {
    if (entryListCustomerSearch) {
      entryListCustomerSearch.value = "";
      entryListCustomerSearch.focus();
    }
    renderEntries();
  });
}
function wireCustomerSearch() {
  customerSearch.addEventListener("input", () => {
    renderCustomers();
  });
  customerSearchClear?.addEventListener("click", () => {
    customerSearch.value = "";
    customerSearch.focus();
    renderCustomers();
  });
}

function wireEmployeeSearch() {
  employeeSearch.addEventListener("input", () => {
    renderEmployees();
  });
  employeeSearchClear?.addEventListener("click", () => {
    employeeSearch.value = "";
    employeeSearch.focus();
    renderEmployees();
  });
}

function wireItemSearch() {
  itemSearch.addEventListener("input", () => {
    renderItems();
  });
  itemSearchClear?.addEventListener("click", () => {
    itemSearch.value = "";
    itemSearch.focus();
    renderItems();
  });
}

function wireUnitSearch() {
  unitSearch?.addEventListener("input", () => {
    renderUnits();
  });
  unitSearchClear?.addEventListener("click", () => {
    if (unitSearch) {
      unitSearch.value = "";
      unitSearch.focus();
    }
    renderUnits();
  });
}

function wireProductTypeSearch() {
  productTypeSearch?.addEventListener("input", () => {
    renderProductTypes();
  });
  productTypeSearchClear?.addEventListener("click", () => {
    if (productTypeSearch) {
      productTypeSearch.value = "";
      productTypeSearch.focus();
    }
    renderProductTypes();
  });
}

function wireDeletedMasterDataSearch() {
  deletedCustomerSearch?.addEventListener("input", renderDeletedCustomers);
  deletedCustomerSearchClear?.addEventListener("click", () => {
    if (!deletedCustomerSearch) return;
    deletedCustomerSearch.value = "";
    deletedCustomerSearch.focus();
    renderDeletedCustomers();
  });

  deletedEmployeeSearch?.addEventListener("input", renderDeletedEmployees);
  deletedEmployeeSearchClear?.addEventListener("click", () => {
    if (!deletedEmployeeSearch) return;
    deletedEmployeeSearch.value = "";
    deletedEmployeeSearch.focus();
    renderDeletedEmployees();
  });

  deletedItemSearch?.addEventListener("input", renderDeletedItems);
  deletedItemSearchClear?.addEventListener("click", () => {
    if (!deletedItemSearch) return;
    deletedItemSearch.value = "";
    deletedItemSearch.focus();
    renderDeletedItems();
  });

  deletedUnitSearch?.addEventListener("input", renderDeletedUnits);
  deletedUnitSearchClear?.addEventListener("click", () => {
    if (!deletedUnitSearch) return;
    deletedUnitSearch.value = "";
    deletedUnitSearch.focus();
    renderDeletedUnits();
  });

  deletedProductTypeSearch?.addEventListener("input", renderDeletedProductTypes);
  deletedProductTypeSearchClear?.addEventListener("click", () => {
    if (!deletedProductTypeSearch) return;
    deletedProductTypeSearch.value = "";
    deletedProductTypeSearch.focus();
    renderDeletedProductTypes();
  });
}
function wireImport() {
  const syncImportButtonState = () => {
    const hasFile = Boolean(importFile.files?.[0]);
    importBtn.disabled = !hasFile;
    importBtn.classList.toggle("primary", hasFile);
    importBtn.classList.toggle("secondary", !hasFile);
  };
  importFile.addEventListener("change", syncImportButtonState);
  syncImportButtonState();

  importBtn.addEventListener("click", async () => {
    try {
      const file = importFile.files?.[0];
      if (!file) {
        importStatus.textContent = "Bitte zuerst eine JSON-Datei auswählen.";
        return;
      }

      const overwriteConfirmed = confirm(
        "Achtung: Beim Import werden alle aktuellen mobilen Daten überschrieben. Möchten Sie fortfahren?"
      );
      if (!overwriteConfirmed) {
        importStatus.textContent = "Import abgebrochen.";
        return;
      }

      const safetyConfirmed = confirm("Sind Sie wirklich sicher?");
      if (!safetyConfirmed) {
        importStatus.textContent = "Import abgebrochen.";
        return;
      }

      const raw = await file.text();
      const parsed = JSON.parse(raw);
      const normalized = normalizeImport(parsed);

        state.customers = normalized.customers;
        state.employees = normalized.employees;
        state.items = normalized.items;
        state.entries = normalized.entries;
        state.units = normalized.units;
        state.productTypes = normalized.productTypes;
        state.settings = normalized.settings;
        ensureCleaningServiceExists();
        enforceFixedUnitsCatalog();
      resetCustomerForm();
      resetEmployeeForm();
      resetItemForm();
      resetUnitForm();
      resetProductTypeForm();
      saveState();
      renderAll();

      importStatus.textContent = `Import erfolgreich: ${state.customers.length} Kunden, ${state.employees.length} Mitarbeiter, ${state.items.length} Einträge, ${state.entries.length} Erfassungen.`;
        showFlashMessage("Import war erfolgreich");
      importFile.value = "";
      syncImportButtonState();
    } catch (error) {
      importStatus.textContent = "Import fehlgeschlagen. Bitte eine gültige Export-Datei verwenden.";
      console.error(error);
    }
  });
}

function editCustomer(customerId) {
  const customer = state.customers.find((c) => c.id === customerId);
  if (!customer) return;

  customerForm.company.value = customer.company || "";
  customerForm.firstName.value = customer.firstName || "";
  customerForm.lastName.value = customer.lastName || "";
  if (customerForm.street) {
    customerForm.street.value = customer.street || "";
  }
  customerForm.zip.value = customer.zip || "";
  customerForm.city.value = customer.city || "";
  customerForm.phone.value = customer.phone || "";
  customerForm.email.value = customer.email || "";

  editingCustomerId = customerId;
  customerSubmitBtn.textContent = "Kunde aktualisieren";
  customerCancelBtn.hidden = false;
  refreshRequiredFieldStates();
  customerForm.firstName?.focus();
  customerForm.firstName?.select();
}

function editEmployee(employeeId) {
  const employee = state.employees.find((e) => e.id === employeeId);
  if (!employee) return;

  employeeForm.firstName.value = employee.firstName || "";
  employeeForm.lastName.value = employee.lastName || "";
  if (employeeForm.street) {
    employeeForm.street.value = employee.street || "";
  }
  employeeForm.zip.value = employee.zip || "";
  employeeForm.city.value = employee.city || "";
  employeeForm.phone.value = employee.phone || "";
  employeeForm.email.value = employee.email || "";
  if (employeeForm.iban) {
    employeeForm.iban.value = employee.iban || "";
  }
  employeeForm.hourlyWage.value = Number(employee.hourlyWage) > 0 ? String(employee.hourlyWage) : "";

  editingEmployeeId = employeeId;
  employeeSubmitBtn.textContent = "Mitarbeiter aktualisieren";
  employeeCancelBtn.hidden = false;
  refreshRequiredFieldStates();
  employeeForm.firstName?.focus();
  employeeForm.firstName?.select();
}

function deleteCustomer(customerId) {
  const customer = state.customers.find((c) => c.id === customerId);
  if (!customer) return;

  const confirmed = confirm("Kunde löschen? Bereits erfasste Leistungen bleiben erhalten. Der Kunde kann später wieder aktiviert werden.");
  if (!confirmed) return;

  customer.deleted = true;
  if (String(entryCustomer.value || "") === String(customerId)) {
    clearSelectedCustomer();
    entryCustomerSearch.value = "";
  }

  if (editingCustomerId === customerId) {
    resetCustomerForm();
  }

  saveState();
  renderAll();
}

function restoreCustomer(customerId) {
  const customer = state.customers.find((c) => c.id === customerId);
  if (!customer) return;
  customer.deleted = false;
  saveState();
  renderAll();
}

function deleteEmployee(employeeId) {
  const employee = state.employees.find((e) => e.id === employeeId);
  if (!employee) return;

  const confirmed = confirm("Mitarbeiter löschen? Bereits erfasste Leistungen bleiben erhalten. Der Mitarbeiter kann später wieder aktiviert werden.");
  if (!confirmed) return;

  employee.deleted = true;
  if (String(entryEmployee.value || "") === String(employeeId)) {
    clearSelectedEmployee();
    entryEmployeeSearch.value = "";
  }

  if (editingEmployeeId === employeeId) {
    resetEmployeeForm();
  }

  saveState();
  renderAll();
}

function restoreEmployee(employeeId) {
  const employee = state.employees.find((e) => e.id === employeeId);
  if (!employee) return;
  employee.deleted = false;
  saveState();
  renderAll();
}
function editItem(itemId) {
  const item = state.items.find((i) => i.id === itemId);
  if (!item) return;

  itemForm.name.value = item.name || "";
  const typeId = String(item.typeId || findOrCreateCatalogByName(state.productTypes, item.type || "Dienstleistung").id);
  const unitId = String(item.unitId || findOrCreateCatalogByName(state.units, item.unit || "h").id);
  if (itemTypeSelect) itemTypeSelect.value = typeId;
  if (itemUnitSelect) itemUnitSelect.value = unitId;
  itemForm.price.value = String(item.price ?? "");
  itemForm.costPrice.value = item.costPrice == null ? "" : String(item.costPrice);

  editingItemId = itemId;
  itemSubmitBtn.textContent = "Eintrag aktualisieren";
  itemCancelBtn.hidden = false;
  refreshRequiredFieldStates();
  itemForm.name?.focus();
  itemForm.name?.select();
}

function deleteItem(itemId) {
  const item = state.items.find((i) => i.id === itemId);
  if (!item) return;

  const confirmed = confirm("Produkt löschen? Bereits erfasste Leistungen bleiben erhalten. Das Produkt kann später wieder aktiviert werden.");
  if (!confirmed) return;

  item.deleted = true;
  if (String(entryItem.value || "") === String(itemId)) {
    clearSelectedItem();
    entryItemSearch.value = "";
  }

  if (editingItemId === itemId) {
    resetItemForm();
  }

  saveState();
  renderAll();
}

function restoreItem(itemId) {
  const item = state.items.find((i) => i.id === itemId);
  if (!item) return;
  item.deleted = false;
  saveState();
  renderAll();
}

function editUnit(unitId) {
  const unit = state.units.find((u) => u.id === unitId);
  if (!unit || !unitForm) return;

  unitForm.name.value = unit.name || "";
  editingUnitId = unitId;
  unitSubmitBtn.textContent = "Einheit aktualisieren";
  if (unitCancelBtn) unitCancelBtn.hidden = false;
  refreshRequiredFieldStates();
  unitForm.name?.focus();
  unitForm.name?.select();
}

function deleteUnit(unitId) {
  const unit = state.units.find((u) => u.id === unitId);
  if (!unit) return;

  const isDefault = ["h", "Stk"].includes(unit.name);
  if (isDefault) {
    alert("Standard-Einheiten können nicht gelöscht werden.");
    return;
  }

  const usedByProducts = state.items.filter((i) => i.unitId === unitId && !i.deleted).length;
  if (usedByProducts > 0) {
    alert(`Einheit kann nicht gelöscht werden. Sie wird noch von ${usedByProducts} Produkt(en) verwendet.`);
    return;
  }

  const confirmed = confirm("Einheit wirklich löschen?");
  if (!confirmed) return;

  unit.deleted = true;

  if (editingUnitId === unitId) resetUnitForm();
  saveState();
  renderAll();
}
function restoreUnit(unitId) {
  const unit = state.units.find((u) => u.id === unitId);
  if (!unit) return;
  unit.deleted = false;
  saveState();
  renderAll();
}

function editProductType(typeId) {
  const type = state.productTypes.find((t) => t.id === typeId);
  if (!type || !productTypeForm) return;

  productTypeForm.name.value = type.name || "";
  editingProductTypeId = typeId;
  productTypeSubmitBtn.textContent = "Produkttyp aktualisieren";
  if (productTypeCancelBtn) productTypeCancelBtn.hidden = false;
  refreshRequiredFieldStates();
  productTypeForm.name?.focus();
  productTypeForm.name?.select();
}

function deleteProductType(typeId) {
  const type = state.productTypes.find((t) => t.id === typeId);
  if (!type) return;

  const isDefault = ["Dienstleistung", "Produkt"].includes(type.name);
  if (isDefault) {
    alert("Standard-Produkttypen können nicht gelöscht werden.");
    return;
  }

  const usedByProducts = state.items.filter((i) => i.typeId === typeId && !i.deleted).length;
  if (usedByProducts > 0) {
    alert(`Produkttyp kann nicht gelöscht werden. Er wird noch von ${usedByProducts} Produkt(en) verwendet.`);
    return;
  }

  const confirmed = confirm("Produkttyp wirklich löschen?");
  if (!confirmed) return;

  type.deleted = true;

  if (editingProductTypeId === typeId) resetProductTypeForm();
  saveState();
  renderAll();
}
function restoreProductType(typeId) {
  const type = state.productTypes.find((t) => t.id === typeId);
  if (!type) return;
  type.deleted = false;
  saveState();
  renderAll();
}

function editEntry(entryId) {
  const entry = state.entries.find((e) => e.id === entryId);
  if (!entry) return;
  setSelectedCustomer(entry.customerId);
  if (entrySplitEmployees) {
    entrySplitEmployees.value = "no";
    entrySplitEmployees.disabled = true;
  }
  syncEntryEmployeeAssignmentMode({ resetSelection: true });
  setSelectedEmployee(entry.employeeId);
  const editCustomer = state.customers.find((c) => c.id === entry.customerId);
  entryCustomerSearch.value = editCustomer ? getCustomerDisplay(editCustomer) : "";
  entryEmployeeSearch.value = entryEmployeeDisplay.value || "";
  setSelectedItem(String(entry.itemId || ""));
  entryItemSearch.value = entryItemDisplay.value || "";
  entryDate.value = String(entry.date || entryDate.value || new Date().toISOString().slice(0, 10));
  entryForm.quantity.value = String(entry.quantity ?? "");
  entryForm.note.value = String(entry.note || "");
  editingEntryId = entryId;
  entrySubmitBtn.textContent = "Erfassung aktualisieren";
  if (entryCancelBtn) entryCancelBtn.hidden = false;
  updateEntryQuantityConstraints();
  refreshRequiredFieldStates();
  entryForm.quantity?.focus();
  entryForm.quantity?.select();
}

function deleteEntry(entryId) {
  const confirmed = confirm("Erfasste Leistung löschen?");
  if (!confirmed) return;
  state.entries = state.entries.filter((e) => e.id !== entryId);
  if (editingEntryId === entryId) {
    resetEntryForm();
  }
  saveState();
  renderEntries();
}

function resetCustomerForm() {
  editingCustomerId = null;
  customerForm.reset();
  customerSubmitBtn.textContent = "Kunde speichern";
  customerCancelBtn.hidden = true;
  refreshRequiredFieldStates();
}

function resetEmployeeForm() {
  editingEmployeeId = null;
  employeeForm.reset();
  employeeSubmitBtn.textContent = "Mitarbeiter speichern";
  employeeCancelBtn.hidden = true;
  refreshRequiredFieldStates();
}

function resetItemForm() {
  editingItemId = null;
  itemForm.reset();
  if (itemTypeSelect) {
    itemTypeSelect.value = "";
  }
  if (itemUnitSelect) {
    itemUnitSelect.value = "";
  }
  itemSubmitBtn.textContent = "Eintrag speichern";
  itemCancelBtn.hidden = true;
  refreshRequiredFieldStates();
}

function resetUnitForm() {
  editingUnitId = null;
  if (unitForm) unitForm.reset();
  if (unitSubmitBtn) unitSubmitBtn.textContent = "Einheit speichern";
  if (unitCancelBtn) unitCancelBtn.hidden = true;
  refreshRequiredFieldStates();
}

function resetProductTypeForm() {
  editingProductTypeId = null;
  if (productTypeForm) productTypeForm.reset();
  if (productTypeSubmitBtn) productTypeSubmitBtn.textContent = "Produkttyp speichern";
  if (productTypeCancelBtn) productTypeCancelBtn.hidden = true;
  refreshRequiredFieldStates();
}

function resetEntryForm() {
  editingEntryId = null;
  entrySubmitBtn.textContent = "Erfassung speichern";
  if (entryCancelBtn) entryCancelBtn.hidden = true;
  clearSelectedItem();
  entryItemSearch.value = "";
  entryForm.quantity.value = "";
  entryForm.note.value = "";
  clearSelectedCustomer();
  entryCustomerSearch.value = "";
  if (entrySplitEmployees) {
    entrySplitEmployees.value = "no";
    entrySplitEmployees.disabled = false;
  }
  clearSelectedEmployee();
  entryEmployeeSearch.value = "";
  clearSelectedSplitEmployeeIds();
  renderCustomerSuggestions("");
  renderEmployeeSuggestions("");
  renderEntryEmployeesMultiOptions();
  renderItemSuggestions("");
  syncEntryEmployeeAssignmentMode({ resetSelection: false });
  updateEntryQuantityConstraints();
  refreshRequiredFieldStates();
  entryForm.quantity?.focus();
  entryForm.quantity?.select();
}

function normalizeImport(parsed) {
  const importedCustomers = Array.isArray(parsed?.customers) ? parsed.customers : [];
  const importedEmployees = Array.isArray(parsed?.employees) ? parsed.employees : [];
  const importedItems = Array.isArray(parsed?.items) ? parsed.items : [];
  const importedEntries = Array.isArray(parsed?.entries) ? parsed.entries : [];
    const importedProductTypes = Array.isArray(parsed?.productTypes) ? parsed.productTypes : [];

  const customers = importedCustomers.map((c) => ({
    id: c?.id || generateId(),
    company: String(c?.company || ""),
    firstName: String(c?.firstName || ""),
    lastName: String(c?.lastName || ""),
    street: String(c?.street || ""),
    zip: String(c?.zip || ""),
    city: String(c?.city || ""),
    phone: String(c?.phone || ""),
    email: String(c?.email || ""),
    deleted: c?.deleted === true || String(c?.deleted || "").toLowerCase() === "true"
  }));

  const units = [
    { id: "unit-h", name: "h" },
    { id: "unit-stk", name: "Stk" }
  ];
  const productTypes = normalizeNamedCatalog(importedProductTypes, "");  ensureCatalogContains(productTypes, "Dienstleistung");
  ensureCatalogContains(productTypes, "Produkt");

  const items = importedItems.map((i) => {
    const typeName = normalizeCatalogName(i?.type) || "Dienstleistung";
        const typeId = String(i?.typeId || "");
    const unitId = String(i?.unitId || "");

    const resolvedTypeId = typeId && productTypes.some((t) => t.id === typeId)
      ? typeId
      : ensureCatalogContains(productTypes, typeName).id;

    const resolvedUnitId = unitId === "unit-stk" ? "unit-stk" : "unit-h";

    return {
      id: i?.id || generateId(),
      name: String(i?.name || ""),
      typeId: resolvedTypeId,
      unitId: resolvedUnitId,
      type: getCatalogNameById(productTypes, resolvedTypeId) || typeName,
      unit: resolvedUnitId === "unit-stk" ? "Stk" : "h",
      price: Number(i?.price) || 0,
      costPrice: parseAmount(i?.costPrice ?? i?.purchasePrice ?? i?.einstandspreis),
      deleted: i?.deleted === true || String(i?.deleted || "").toLowerCase() === "true"
    };
  });

  const employees = importedEmployees.map((e) => ({
    id: e?.id || generateId(),
    firstName: String(e?.firstName || ""),
    lastName: String(e?.lastName || ""),
    street: String(e?.street || ""),
    zip: String(e?.zip || ""),
    city: String(e?.city || ""),
    phone: String(e?.phone || ""),
    email: String(e?.email || ""),
    iban: String(e?.iban || ""),
    hourlyWage: parseAmount(e?.hourlyWage),
    deleted: e?.deleted === true || String(e?.deleted || "").toLowerCase() === "true"
  }));

  const customerIds = new Set(customers.map((c) => c.id));
  const employeeIds = new Set(employees.map((e) => e.id));
  const itemIds = new Set(items.map((i) => i.id));
  const today = new Date().toISOString().slice(0, 10);

  const entries = importedEntries
    .map((e) => {
      const itemId = String(e?.itemId || "");
      const item = items.find((i) => i.id === itemId);
      return {
        id: e?.id || generateId(),
        customerId: String(e?.customerId || ""),
        employeeId: String(e?.employeeId || ""),
        itemId,
        date: String(e?.date || today),
        quantity: Number(e?.quantity) > 0 ? Number(e.quantity) : 1,
        note: String(e?.note || ""),
        unitPrice: parseAmount(e?.unitPrice),
        costPrice: normalizeEntryCostPrice(e, item)
      };
    })
    .filter((e) => {
      const employeeOk = !e.employeeId || employeeIds.has(e.employeeId);
      return customerIds.has(e.customerId) && itemIds.has(e.itemId) && employeeOk;
    });

  const settings = {
    currency: normalizeCurrency(parsed?.settings?.currency),
    retentionMonths: normalizeMonths(parsed?.settings?.retentionMonths)
  };

  return { customers, employees, items, entries, units, productTypes, settings };
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function renderAll() {
  renderCustomers();
  renderDeletedCustomers();
  renderEmployees();
  renderDeletedEmployees();
  renderUnits();
  renderDeletedUnits();
  renderProductTypes();
  renderDeletedProductTypes();
  renderItems();
  renderDeletedItems();
  renderItemTypeUnitSelects();
  renderEntrySelects();
  renderEntries();
  renderReportOverview();
  renderProductReportOverview();
  renderSettings();
  renderExportCleanupPreview();
  refreshRequiredFieldStates();
}

function wireRequiredFieldStates() {
  const requiredFields = document.querySelectorAll("form input[required], form select[required], form textarea[required]");
  requiredFields.forEach((field) => {
    const refresh = () => {
      updateRequiredFieldState(field);
      syncEntrySubmitButtonState();
    };
    field.addEventListener("input", refresh);
    field.addEventListener("change", refresh);
    field.addEventListener("blur", refresh);
    refresh();
  });
  const customerCompanyField = customerForm?.company;
  if (customerCompanyField) {
    const refreshCustomerNameRequirement = () => {
      syncCustomerNameRequirement();
      syncEntrySubmitButtonState();
    };
    customerCompanyField.addEventListener("input", refreshCustomerNameRequirement);
    customerCompanyField.addEventListener("change", refreshCustomerNameRequirement);
    customerCompanyField.addEventListener("blur", refreshCustomerNameRequirement);
    refreshCustomerNameRequirement();
  }
  const optionalFields = document.querySelectorAll("form input:not([required]):not([type='hidden']), form select:not([required]), form textarea:not([required])");
  optionalFields.forEach((field) => {
    const refresh = () => updateOptionalFieldState(field);
    field.addEventListener("input", refresh);
    field.addEventListener("change", refresh);
    field.addEventListener("blur", refresh);
    refresh();
  });
}

function refreshRequiredFieldStates() {
  const requiredFields = document.querySelectorAll("form input[required], form select[required], form textarea[required]");
  requiredFields.forEach((field) => updateRequiredFieldState(field));
  syncCustomerNameRequirement();
  const optionalFields = document.querySelectorAll("form input:not([required]):not([type='hidden']), form select:not([required]), form textarea:not([required])");
  optionalFields.forEach((field) => updateOptionalFieldState(field));
  updateEntrySelectionDisplayStates();
  updateEntryMultiEmployeeSelectionState();
  syncEntrySubmitButtonState();
}

function updateRequiredFieldState(field) {
  const rawValue = String(field?.value ?? "").trim();
  const isRequired = Boolean(field?.required);
  if (!isRequired) {
    const isValidOptional = rawValue.length === 0 ? true : field.checkValidity();
    field.classList.toggle("required-empty", !isValidOptional);
    field.classList.toggle("required-filled", isValidOptional);
    const optionalLabel = field.closest("label");
    if (!optionalLabel) return;
    optionalLabel.classList.add("required-indicator");
    optionalLabel.classList.toggle("required-empty", !isValidOptional);
    optionalLabel.classList.toggle("required-filled", isValidOptional);
    return;
  }
  const hasValue = rawValue.length > 0;
  let isValid = hasValue && field.checkValidity();
  const linkedId = field.dataset?.requiredLinkedId;
  if (linkedId) {
    const linkedField = document.getElementById(linkedId);
    const linkedHasValue = String(linkedField?.value ?? "").trim().length > 0;
    isValid = isValid && linkedHasValue;
  }
  field.classList.toggle("required-empty", !isValid);
  field.classList.toggle("required-filled", isValid);
  const label = field.closest("label");
  if (!label) return;
  label.classList.add("required-indicator");
  label.classList.toggle("required-empty", !isValid);
  label.classList.toggle("required-filled", isValid);
}

function updateOptionalFieldState(field) {
  const rawValue = String(field?.value ?? "").trim();
  const isValid = rawValue.length === 0 ? true : field.checkValidity();
  field.classList.toggle("required-empty", !isValid);
  field.classList.toggle("required-filled", isValid);
  const label = field.closest("label");
  if (!label) return;
  label.classList.add("required-indicator");
  label.classList.toggle("required-empty", !isValid);
  label.classList.toggle("required-filled", isValid);
}

function updateEntrySelectionDisplayStates() {
  setDisplayFieldState(entryCustomerDisplay, String(entryCustomer.value || "").trim().length > 0);
  setDisplayFieldState(entryEmployeeDisplay, String(entryEmployee.value || "").trim().length > 0);
  setDisplayFieldState(entryItemDisplay, String(entryItem.value || "").trim().length > 0);
}


function isFormReadyForSubmit(form) {
  if (!form) return false;
  const requiredFields = form.querySelectorAll("input[required], select[required], textarea[required]");
  return [...requiredFields].every((field) => {
    const rawValue = String(field?.value ?? "").trim();
    if (!rawValue) return false;
    if (!field.checkValidity()) return false;
    const linkedId = field.dataset?.requiredLinkedId;
    if (!linkedId) return true;
    const linkedField = document.getElementById(linkedId);
    return String(linkedField?.value ?? "").trim().length > 0;
  });
}

function isEntryFormReadyForSubmit() {
  const baseReady = isFormReadyForSubmit(entryForm);
  if (!baseReady) return false;
  if (isSplitEmployeeAssignmentEnabled()) {
    return syncSplitAllocationValidation();
  }
  return true;
}

function syncEntrySubmitButtonState() {
  if (entrySubmitBtn) entrySubmitBtn.disabled = !isEntryFormReadyForSubmit();
  if (customerSubmitBtn) customerSubmitBtn.disabled = !isFormReadyForSubmit(customerForm);
  if (employeeSubmitBtn) employeeSubmitBtn.disabled = !isFormReadyForSubmit(employeeForm);
  if (itemSubmitBtn) itemSubmitBtn.disabled = !isFormReadyForSubmit(itemForm);
  if (unitSubmitBtn) unitSubmitBtn.disabled = !isFormReadyForSubmit(unitForm);
  if (productTypeSubmitBtn) productTypeSubmitBtn.disabled = !isFormReadyForSubmit(productTypeForm);
}
function syncCustomerNameRequirement() {
  if (!customerForm) return;
  const companyValue = String(customerForm.company?.value ?? "").trim();
  const requirePersonName = companyValue.length === 0;
  const firstNameField = customerForm.firstName;
  const lastNameField = customerForm.lastName;
  if (firstNameField) {
    firstNameField.required = requirePersonName;
    updateRequiredFieldState(firstNameField);
  }
  if (lastNameField) {
    lastNameField.required = requirePersonName;
    updateRequiredFieldState(lastNameField);
  }
}

function setDisplayFieldState(field, isValid) {
  if (!field) return;
  field.classList.toggle("required-empty", !isValid);
  field.classList.toggle("required-filled", isValid);
  const label = field.closest("label");
  if (!label) return;
  label.classList.add("required-indicator");
  label.classList.toggle("required-empty", !isValid);
  label.classList.toggle("required-filled", isValid);
}

function renderSettings() {
  if (currencyCode) currencyCode.value = normalizeCurrency(state?.settings?.currency);
  retentionMonths.value = String(normalizeMonths(state?.settings?.retentionMonths));
}
function renderCustomers() {
  const activeCustomers = state.customers.filter((c) => !c.deleted);
  if (!activeCustomers.length) {
    customerList.innerHTML = "<small>Noch keine aktiven Kunden erfasst.</small>";
    return;
  }

  const query = normalizeSearchText(customerSearch.value);
  const visibleCustomers = activeCustomers.filter((c) => {
    if (!query) return true;
    const haystack = normalizeSearchText([
      c.company,
      c.firstName,
      c.lastName,
      c.street,
      c.zip,
      c.city,
      c.phone,
      c.email
    ].join(" "));
    return haystack.includes(query);
  });

  if (!visibleCustomers.length) {
    customerList.innerHTML = "<small>Keine aktiven Kunden passend zur Suche gefunden.</small>";
    return;
  }

  customerList.innerHTML = visibleCustomers
    .map((c) => {
      const company = c.company ? `${escapeHtml(c.company)} - ` : "";
      return `
        <article class="card">
          <strong>${company}${escapeHtml(c.firstName)} ${escapeHtml(c.lastName)}</strong>
          <small class="customer-address">${escapeHtml(c.street)}, ${escapeHtml(c.zip)} ${escapeHtml(c.city)}</small>
          <small class="customer-meta">${escapeHtml(c.phone || "")}${c.email ? ` | ${escapeHtml(c.email)}` : ""}</small>
          <div class="card-actions">
            <button type="button" class="secondary" data-action="edit" data-id="${escapeHtml(c.id)}">Bearbeiten</button>
            <button type="button" class="danger" data-action="delete" data-id="${escapeHtml(c.id)}">Löschen</button>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderDeletedCustomers() {
  if (!deletedCustomerList) return;
  if (deletedCustomerSummary) deletedCustomerSummary.textContent = "Gelöschte aber in der Leistungserfassung noch verwendete Kunden (0)";
  const query = normalizeSearchText(deletedCustomerSearch?.value || "");
  const deletedCustomers = state.customers.filter((c) => {
    if (!c.deleted) return false;
    if (!query) return true;
    const haystack = normalizeSearchText([c.company, c.firstName, c.lastName, c.street, c.zip, c.city, c.phone, c.email].join(" "));
    return haystack.includes(query);
  });
  if (deletedCustomerSummary) deletedCustomerSummary.textContent = "Gelöschte aber in der Leistungserfassung noch verwendete Kunden (" + deletedCustomers.length + ")";
  if (!deletedCustomers.length) {
    deletedCustomerList.innerHTML = "<small>Keine gelöschten Kunden passend zur Suche gefunden.</small>";
    return;
  }
  deletedCustomerList.innerHTML = deletedCustomers
    .map((c) => {
      const company = c.company ? `${escapeHtml(c.company)} - ` : "";
      return `
        <article class="card">
          <strong>${company}${escapeHtml(c.firstName)} ${escapeHtml(c.lastName)}</strong>
          <div class="card-actions">
            <button type="button" class="secondary" data-action="restore" data-id="${escapeHtml(c.id)}">Wieder aktivieren</button>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderEmployees() {
  const activeEmployees = state.employees.filter((e) => !e.deleted);
  if (!activeEmployees.length) {
    employeeList.innerHTML = "<small>Noch keine aktiven Mitarbeiter erfasst.</small>";
    return;
  }

  const query = normalizeSearchText(employeeSearch.value);
  const visibleEmployees = activeEmployees.filter((e) => {
    if (!query) return true;
    const haystack = normalizeSearchText([
      e.firstName,
      e.lastName,
      e.street,
      e.zip,
      e.city,
      e.phone,
      e.email,
      e.iban
    ].join(" "));
    return haystack.includes(query);
  });

  if (!visibleEmployees.length) {
    employeeList.innerHTML = "<small>Keine aktiven Mitarbeiter passend zur Suche gefunden.</small>";
    return;
  }

  employeeList.innerHTML = visibleEmployees
    .map((e) => {
      const fullName = `${e.firstName || ""} ${e.lastName || ""}`.trim();
      return `
        <article class="card">
          <strong>${escapeHtml(fullName)}</strong>
          <small class="customer-address">${escapeHtml(e.street || "")}, ${escapeHtml(e.zip || "")} ${escapeHtml(e.city || "")}</small>
          <small class="customer-meta">${escapeHtml(e.phone || "")}${e.email ? ` | ${escapeHtml(e.email)}` : ""}${e.iban ? ` | ${escapeHtml(e.iban)}` : ""}${Number(e.hourlyWage) > 0 ? ` | Vereinbarter Stundenlohn: ${formatCurrency(e.hourlyWage)}` : ""}</small>
          <div class="card-actions">
            <button type="button" class="secondary" data-action="edit" data-id="${escapeHtml(e.id)}">Bearbeiten</button>
            <button type="button" class="danger" data-action="delete" data-id="${escapeHtml(e.id)}">Löschen</button>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderDeletedEmployees() {
  if (!deletedEmployeeList) return;
  if (deletedEmployeeSummary) deletedEmployeeSummary.textContent = "Gelöschte aber in der Leistungserfassung noch verwendete Mitarbeiter (0)";
  const query = normalizeSearchText(deletedEmployeeSearch?.value || "");
  const deletedEmployees = state.employees.filter((e) => {
    if (!e.deleted) return false;
    if (!query) return true;
    const haystack = normalizeSearchText([e.firstName, e.lastName, e.street, e.zip, e.city, e.phone, e.email, e.iban].join(" "));
    return haystack.includes(query);
  });
  if (deletedEmployeeSummary) deletedEmployeeSummary.textContent = "Gelöschte aber in der Leistungserfassung noch verwendete Mitarbeiter (" + deletedEmployees.length + ")";
  if (!deletedEmployees.length) {
    deletedEmployeeList.innerHTML = "<small>Keine gelöschten Mitarbeiter passend zur Suche gefunden.</small>";
    return;
  }
  deletedEmployeeList.innerHTML = deletedEmployees
    .map((e) => {
      const fullName = `${e.firstName || ""} ${e.lastName || ""}`.trim();
      return `
        <article class="card">
          <strong>${escapeHtml(fullName)}</strong>
          <div class="card-actions">
            <button type="button" class="secondary" data-action="restore" data-id="${escapeHtml(e.id)}">Wieder aktivieren</button>
          </div>
        </article>
      `;
    })
    .join("");
}
function renderUnits() {
  if (!unitList) return;
  const activeUnits = state.units.filter((u) => !u.deleted);
  if (!activeUnits.length) {
    unitList.innerHTML = "<small>Noch keine aktiven Einheiten erfasst.</small>";
    return;
  }

  const query = normalizeSearchText(unitSearch?.value);
  const visible = activeUnits.filter((u) => {
    if (!query) return true;
    return normalizeSearchText(u.name).includes(query);
  });

  if (!visible.length) {
    unitList.innerHTML = "<small>Keine aktiven Einheiten passend zur Suche gefunden.</small>";
    return;
  }

  unitList.innerHTML = visible
    .map((u) => `
      <article class="card">
        <strong>${escapeHtml(u.name)}</strong>
        <div class="card-actions">
          <button type="button" class="secondary" data-action="edit" data-id="${escapeHtml(u.id)}">Bearbeiten</button>
          <button type="button" class="danger" data-action="delete" data-id="${escapeHtml(u.id)}">Löschen</button>
        </div>
      </article>
    `)
    .join("");
}

function renderDeletedUnits() {
  if (!deletedUnitList) return;
  if (deletedUnitSummary) deletedUnitSummary.textContent = "Gelöschte aber in der Leistungserfassung noch verwendete Einheiten (0)";
  const query = normalizeSearchText(deletedUnitSearch?.value || "");
  const deletedUnits = state.units.filter((u) => u.deleted && (!query || normalizeSearchText(u.name).includes(query)));
  if (deletedUnitSummary) deletedUnitSummary.textContent = "Gelöschte aber in der Leistungserfassung noch verwendete Einheiten (" + deletedUnits.length + ")";
  if (!deletedUnits.length) {
    deletedUnitList.innerHTML = "<small>Keine gelöschten Einheiten passend zur Suche gefunden.</small>";
    return;
  }
  deletedUnitList.innerHTML = deletedUnits
    .map((u) => `
      <article class="card">
        <strong>${escapeHtml(u.name)}</strong>
        <div class="card-actions">
          <button type="button" class="secondary" data-action="restore" data-id="${escapeHtml(u.id)}">Wieder aktivieren</button>
        </div>
      </article>
    `)
    .join("");
}

function renderProductTypes() {
  if (!productTypeList) return;
  const activeTypes = state.productTypes.filter((t) => !t.deleted);
  if (!activeTypes.length) {
    productTypeList.innerHTML = "<small>Noch keine aktiven Produkttypen erfasst.</small>";
    return;
  }

  const query = normalizeSearchText(productTypeSearch?.value);
  const visible = activeTypes.filter((t) => {
    if (!query) return true;
    return normalizeSearchText(t.name).includes(query);
  });

  if (!visible.length) {
    productTypeList.innerHTML = "<small>Keine aktiven Produkttypen passend zur Suche gefunden.</small>";
    return;
  }

  productTypeList.innerHTML = visible
    .map((t) => `
      <article class="card">
        <strong>${escapeHtml(t.name)}</strong>
        <div class="card-actions">
          <button type="button" class="secondary" data-action="edit" data-id="${escapeHtml(t.id)}">Bearbeiten</button>
          <button type="button" class="danger" data-action="delete" data-id="${escapeHtml(t.id)}">Löschen</button>
        </div>
      </article>
    `)
    .join("");
}

function renderDeletedProductTypes() {
  if (!deletedProductTypeList) return;
  if (deletedProductTypeSummary) deletedProductTypeSummary.textContent = "Gelöschte aber in der Leistungserfassung noch verwendete Produkttypen (0)";
  const query = normalizeSearchText(deletedProductTypeSearch?.value || "");
  const deletedTypes = state.productTypes.filter((t) => t.deleted && (!query || normalizeSearchText(t.name).includes(query)));
  if (deletedProductTypeSummary) deletedProductTypeSummary.textContent = "Gelöschte aber in der Leistungserfassung noch verwendete Produkttypen (" + deletedTypes.length + ")";
  if (!deletedTypes.length) {
    deletedProductTypeList.innerHTML = "<small>Keine gelöschten Produkttypen passend zur Suche gefunden.</small>";
    return;
  }
  deletedProductTypeList.innerHTML = deletedTypes
    .map((t) => `
      <article class="card">
        <strong>${escapeHtml(t.name)}</strong>
        <div class="card-actions">
          <button type="button" class="secondary" data-action="restore" data-id="${escapeHtml(t.id)}">Wieder aktivieren</button>
        </div>
      </article>
    `)
    .join("");
}

function renderItemTypeUnitSelects() {
  const activeTypes = state.productTypes.filter((type) => !type.deleted);
  const activeUnits = state.units.filter((unit) => !unit.deleted);

  if (itemTypeSelect) {
    const previousTypeId = itemTypeSelect.value;
    itemTypeSelect.innerHTML = [
      '<option value="">Bitte Typ wählen</option>',
      ...activeTypes.map((type) => `<option value="${escapeHtml(type.id)}">${escapeHtml(type.name)}</option>`)
    ].join("");
    if (activeTypes.some((type) => type.id === previousTypeId)) {
      itemTypeSelect.value = previousTypeId;
    } else {
      itemTypeSelect.value = "";
    }
  }

  if (itemUnitSelect) {
    const previousUnitId = itemUnitSelect.value;
    itemUnitSelect.innerHTML = [
      '<option value="">Bitte Einheit wählen</option>',
      ...activeUnits.map((unit) => `<option value="${escapeHtml(unit.id)}">${escapeHtml(unit.name)}</option>`)
    ].join("");
    if (activeUnits.some((unit) => unit.id === previousUnitId)) {
      itemUnitSelect.value = previousUnitId;
    } else {
      itemUnitSelect.value = "";
    }
  }
}

function renderItems() {
  const activeItems = state.items.filter((i) => !i.deleted);
  if (!activeItems.length) {
    itemList.innerHTML = "<small>Noch keine aktiven Produkte erfasst.</small>";
    return;
  }

  const query = normalizeSearchText(itemSearch.value);
  const visibleItems = activeItems.filter((i) => {
    if (!query) return true;
    const haystack = normalizeSearchText([
      i.name,
      getItemTypeName(i),
      getItemUnitName(i),
      String(i.price),
      String(i.costPrice ?? "")
    ].join(" "));
    return haystack.includes(query);
  });

  if (!visibleItems.length) {
    itemList.innerHTML = "<small>Keine aktiven Produkte passend zur Suche gefunden.</small>";
    return;
  }

  itemList.innerHTML = visibleItems
    .map(
      (i) => `
      <article class="card">
        <strong>${escapeHtml(i.name)}</strong>
        <small>${escapeHtml(getItemTypeName(i))} | ${formatCurrency(i.price)} / ${escapeHtml(getItemUnitName(i))}${i.costPrice != null ? ` | Einstand: ${formatCurrency(i.costPrice)}` : ""}</small>
        <div class="card-actions">
          <button type="button" class="secondary" data-action="edit" data-id="${escapeHtml(i.id)}">Bearbeiten</button>
          <button type="button" class="danger" data-action="delete" data-id="${escapeHtml(i.id)}">Löschen</button>
        </div>
      </article>
    `
    )
    .join("");
}

function renderDeletedItems() {
  if (!deletedItemList) return;
  if (deletedItemSummary) deletedItemSummary.textContent = "Gelöschte aber in der Leistungserfassung noch verwendete Produkte (0)";
  const query = normalizeSearchText(deletedItemSearch?.value || "");
  const deletedItems = state.items.filter((i) => {
    if (!i.deleted) return false;
    if (!query) return true;
    const haystack = normalizeSearchText([i.name, getItemTypeName(i), getItemUnitName(i), String(i.price), String(i.costPrice ?? "")].join(" "));
    return haystack.includes(query);
  });
  if (deletedItemSummary) deletedItemSummary.textContent = "Gelöschte aber in der Leistungserfassung noch verwendete Produkte (" + deletedItems.length + ")";
  if (!deletedItems.length) {
    deletedItemList.innerHTML = "<small>Keine gelöschten Produkte passend zur Suche gefunden.</small>";
    return;
  }
  deletedItemList.innerHTML = deletedItems
    .map((i) => `
      <article class="card">
        <strong>${escapeHtml(i.name)}</strong>
        <small>${escapeHtml(getItemTypeName(i))} | ${formatCurrency(i.price)} / ${escapeHtml(getItemUnitName(i))}${i.costPrice != null ? ` | Einstand: ${formatCurrency(i.costPrice)}` : ""}</small>
        <div class="card-actions">
          <button type="button" class="secondary" data-action="restore" data-id="${escapeHtml(i.id)}">Wieder aktivieren</button>
        </div>
      </article>
    `)
    .join("");
}

function renderEntrySelects() {
  const selectedCustomerExists = state.customers.some((c) => !c.deleted && c.id === entryCustomer.value);
  if (!selectedCustomerExists) {
    clearSelectedCustomer();
  }
  const selectedEmployeeExists = state.employees.some((e) => !e.deleted && e.id === entryEmployee.value);
  if (!selectedEmployeeExists) {
    clearSelectedEmployee();
  }
  const selectedItemExists = state.items.some((item) => !item.deleted && item.id === entryItem.value);
  if (!selectedItemExists) {
    clearSelectedItem();
  }

  renderCustomerSuggestions(entryCustomerSearch.value);
  renderEmployeeSuggestions(entryEmployeeSearch.value);
  renderEntryEmployeesMultiOptions();
  renderItemSuggestions(entryItemSearch.value);
  syncEntryEmployeeAssignmentMode({ resetSelection: false });
  updateEntryQuantityConstraints();
}

function updateEntryQuantityConstraints() {
  if (!entryQuantityInput) return;
  const isPieceUnit = isPieceUnitSelected();
  if (isPieceUnit) {
    entryQuantityInput.step = "1";
    entryQuantityInput.min = "1";
    entryQuantityInput.inputMode = "numeric";
    const current = Number(entryQuantityInput.value);
    if (Number.isFinite(current) && current > 0 && !Number.isInteger(current)) {
      entryQuantityInput.value = String(Math.max(1, Math.round(current)));
    }
    return;
  }
  entryQuantityInput.step = "0.01";
  entryQuantityInput.min = "0.01";
  entryQuantityInput.inputMode = "decimal";
}

function isPieceUnitSelected() {
  const selectedItem = state.items.find((item) => item.id === entryItem.value);
  return getItemUnitName(selectedItem) === "Stk";
}

function isSplitEmployeeAssignmentEnabled() {
  return String(entrySplitEmployees?.value || "no").toLowerCase() === "yes";
}

function getSelectedSplitEmployeeIds() {
  if (!entryEmployeesMultiList) return [];
  return [...entryEmployeesMultiList.querySelectorAll("input.entry-employee-multi[type=\"checkbox\"]:checked")]
    .map((input) => String(input.value || "").trim())
    .filter(Boolean);
}

function clearSelectedSplitEmployeeIds() {
  if (!entryEmployeesMultiList) return;
  entryEmployeesMultiList.querySelectorAll("input.entry-employee-multi[type=\"checkbox\"]").forEach((input) => {
    input.checked = false;
  });
  splitEmployeeQuantityMap = new Map();
  if (entryEmployeesSplitDetails) {
    entryEmployeesSplitDetails.innerHTML = "";
    entryEmployeesSplitDetails.hidden = true;
  }
  if (entryEmployeesSplitHint) {
    entryEmployeesSplitHint.textContent = "";
    entryEmployeesSplitHint.hidden = true;
    entryEmployeesSplitHint.classList.remove("split-hint-ok", "split-hint-error");
  }
}

function getSplitQuantityScale(isPieceUnit) {
  return isPieceUnit ? 1 : 100;
}

function normalizeSplitQuantity(value, isPieceUnit) {
  const scale = getSplitQuantityScale(isPieceUnit);
  const scaled = Math.round(Number(value || 0) * scale);
  return scaled / scale;
}

function formatSplitQuantityValue(value, isPieceUnit) {
  const normalized = normalizeSplitQuantity(value, isPieceUnit);
  if (isPieceUnit) return String(Math.max(0, Math.trunc(normalized)));
  return normalized.toFixed(2);
}

function parseSplitQuantityInputValue(rawValue, isPieceUnit) {
  const normalizedRaw = String(rawValue ?? "").replace(",", ".").trim();
  if (!normalizedRaw) return "";
  const parsed = Number(normalizedRaw);
  if (!Number.isFinite(parsed) || parsed < 0) return "";
  if (isPieceUnit) return String(Math.max(0, Math.trunc(parsed)));
  return formatSplitQuantityValue(parsed, false);
}

function getMainQuantityValue(isPieceUnit) {
  const parsed = Number(String(entryQuantityInput?.value || "").replace(",", "."));
  if (!Number.isFinite(parsed) || parsed <= 0) return 0;
  return normalizeSplitQuantity(parsed, isPieceUnit);
}

function getDefaultSplitQuantities(mainQuantity, employeeIds, isPieceUnit) {
  const count = employeeIds.length;
  if (!count || mainQuantity <= 0) return new Map();
  const map = new Map();
  const scale = getSplitQuantityScale(isPieceUnit);
  const totalScaled = Math.round(mainQuantity * scale);
  const base = Math.floor(totalScaled / count);
  const remainder = totalScaled - base * count;

  employeeIds.forEach((employeeId, index) => {
    const scaledValue = base + (index < remainder ? 1 : 0);
    const value = scaledValue / scale;
    map.set(employeeId, formatSplitQuantityValue(value, isPieceUnit));
  });
  return map;
}
function updateEntryMultiEmployeeSelectionState() {
  if (!entryEmployeesMultiRow || !entryEmployeesMultiList) return;
  const isSplitMode = isSplitEmployeeAssignmentEnabled();
  const selectedCount = getSelectedSplitEmployeeIds().length;
  const isValid = !isSplitMode || selectedCount > 1;

  entryEmployeesMultiRow.classList.add("required-indicator");
  entryEmployeesMultiRow.classList.toggle("required-empty", !isValid);
  entryEmployeesMultiRow.classList.toggle("required-filled", isValid);
  entryEmployeesMultiList.classList.toggle("required-empty", !isValid);
  entryEmployeesMultiList.classList.remove("required-filled");
}

function renderEntryEmployeesMultiOptions() {
  if (!entryEmployeesMultiList) return;
  const selectedBeforeRender = new Set(getSelectedSplitEmployeeIds());
  const activeEmployees = state.employees
    .filter((employee) => !employee.deleted)
    .slice()
    .sort((a, b) => getEmployeeDisplay(a).localeCompare(getEmployeeDisplay(b), "de-CH"));

  if (!activeEmployees.length) {
    entryEmployeesMultiList.innerHTML = "<small>Keine aktiven Mitarbeiter vorhanden.</small>";
    updateEntryMultiEmployeeSelectionState();
    return;
  }

  entryEmployeesMultiList.innerHTML = activeEmployees
    .map((employee) => {
      const checked = selectedBeforeRender.has(String(employee.id || "")) ? " checked" : "";
      return `<label class="multi-employee-option"><input class="entry-employee-multi" type="checkbox" value="${escapeHtml(employee.id)}"${checked} /> <span>${escapeHtml(getEmployeeDisplay(employee))}</span></label>`;
    })
    .join("");

  updateEntryMultiEmployeeSelectionState();
}

function renderSplitQuantityInputs() {
  if (!entryEmployeesSplitDetails) return;
  const isSplitMode = isSplitEmployeeAssignmentEnabled();
  const selectedIds = getSelectedSplitEmployeeIds();

  if (!isSplitMode || !selectedIds.length) {
    entryEmployeesSplitDetails.innerHTML = "";
    entryEmployeesSplitDetails.hidden = true;
    if (entryEmployeesSplitHint) {
      entryEmployeesSplitHint.textContent = "";
      entryEmployeesSplitHint.hidden = true;
      entryEmployeesSplitHint.classList.remove("split-hint-ok", "split-hint-error");
    }
    return;
  }

  const isPieceUnit = isPieceUnitSelected();
  const mainQuantity = getMainQuantityValue(isPieceUnit);
  const defaultMap = getDefaultSplitQuantities(mainQuantity, selectedIds, isPieceUnit);

  const nextMap = new Map();
  selectedIds.forEach((employeeId) => {
    const existing = splitEmployeeQuantityMap.get(employeeId);
    const value = existing != null && String(existing).trim() !== ""
      ? parseSplitQuantityInputValue(existing, isPieceUnit)
      : String(defaultMap.get(employeeId) ?? "");
    nextMap.set(employeeId, value);
  });
  splitEmployeeQuantityMap = nextMap;

  entryEmployeesSplitDetails.innerHTML = selectedIds
    .map((employeeId) => {
      const employee = state.employees.find((e) => String(e.id) === String(employeeId));
      const label = employee ? getEmployeeDisplay(employee) : "Unbekannter Mitarbeiter";
      const value = String(splitEmployeeQuantityMap.get(employeeId) ?? "");
      const step = isPieceUnit ? "1" : "0.01";
      const inputMode = isPieceUnit ? "numeric" : "decimal";
      return `<div class="multi-employee-split-row"><span class="multi-employee-split-name">${escapeHtml(label)}</span><input class="entry-employee-split-qty" data-employee-id="${escapeHtml(employeeId)}" type="number" min="0" step="${step}" inputmode="${inputMode}" value="${escapeHtml(value)}" /></div>`;
    })
    .join("");

  entryEmployeesSplitDetails.hidden = false;
  syncSplitAllocationValidation();
}

function getSplitAssignmentsFromInputs(selectedEmployeeIds, mainQuantity, isPieceUnit) {
  const assignments = selectedEmployeeIds.map((employeeId) => {
    const raw = String(splitEmployeeQuantityMap.get(employeeId) ?? "").replace(",", ".").trim();
    const value = Number(raw);
    return { employeeId, quantity: normalizeSplitQuantity(value, isPieceUnit) };
  });

  if (assignments.length < 2) {
    return { ok: false, assignments: [], message: "Bitte mindestens zwei Mitarbeiter auswählen." };
  }

  if (assignments.some((item) => !Number.isFinite(item.quantity) || item.quantity <= 0)) {
    return { ok: false, assignments: [], message: "Bitte pro ausgewähltem Mitarbeiter eine Menge grösser als 0 eingeben." };
  }

  const scale = getSplitQuantityScale(isPieceUnit);
  const mainScaled = Math.round(mainQuantity * scale);
  const sumScaled = assignments.reduce((acc, item) => acc + Math.round(item.quantity * scale), 0);

  if (sumScaled !== mainScaled) {
    return { ok: false, assignments: [], message: "Die Summe der Mitarbeiter-Anteile muss genau der Hauptmenge entsprechen." };
  }

  const sum = sumScaled / scale;
  return { ok: true, assignments, sum };
}

function syncSplitAllocationValidation() {
  if (!entryEmployeesSplitDetails || !entryEmployeesMultiRow) return true;

  const isSplitMode = isSplitEmployeeAssignmentEnabled();
  if (!isSplitMode) {
    entryEmployeesSplitDetails.classList.remove("required-empty", "required-filled");
    if (entryEmployeesSplitHint) {
      entryEmployeesSplitHint.hidden = true;
      entryEmployeesSplitHint.textContent = "";
      entryEmployeesSplitHint.classList.remove("split-hint-error", "split-hint-ok");
    }
    return true;
  }

  const selectedIds = getSelectedSplitEmployeeIds();
  const isPieceUnit = isPieceUnitSelected();
  const mainQuantity = getMainQuantityValue(isPieceUnit);
  const result = getSplitAssignmentsFromInputs(selectedIds, mainQuantity, isPieceUnit);

  entryEmployeesSplitDetails.classList.toggle("required-empty", !result.ok);
  entryEmployeesSplitDetails.classList.toggle("required-filled", result.ok);

  if (entryEmployeesSplitHint) {
    entryEmployeesSplitHint.hidden = false;
    if (result.ok) {
      const sum = Number(result.sum || 0);
      entryEmployeesSplitHint.textContent = `Zugewiesen: ${sum} / Hauptmenge: ${mainQuantity} (exakt)`;
      entryEmployeesSplitHint.classList.add("split-hint-ok");
      entryEmployeesSplitHint.classList.remove("split-hint-error");
    } else {
      entryEmployeesSplitHint.textContent = result.message || "Aufteilung prüfen";
      entryEmployeesSplitHint.classList.add("split-hint-error");
      entryEmployeesSplitHint.classList.remove("split-hint-ok");
    }
  }

  return result.ok;
}

function syncEntryEmployeeAssignmentMode({ resetSelection = false } = {}) {
  const isSplitMode = isSplitEmployeeAssignmentEnabled();

  if (entryEmployeeSearchRow) entryEmployeeSearchRow.hidden = isSplitMode;
  if (entryEmployeeSelectedRow) entryEmployeeSelectedRow.hidden = true;
  if (entryEmployeesMultiRow) entryEmployeesMultiRow.hidden = !isSplitMode;

  if (entryEmployeeSearch) {
    entryEmployeeSearch.required = !isSplitMode;
  }

  if (isSplitMode) {
    if (resetSelection) {
      entryEmployeeSearch.value = "";
      clearSelectedEmployee();
      clearSelectedSplitEmployeeIds();
    }
    renderEntryEmployeesMultiOptions();
    renderSplitQuantityInputs();
  } else {
    if (resetSelection) {
      clearSelectedSplitEmployeeIds();
    }
    if (entryEmployeesSplitDetails) {
      entryEmployeesSplitDetails.hidden = true;
      entryEmployeesSplitDetails.innerHTML = "";
    }
    if (entryEmployeesSplitHint) {
      entryEmployeesSplitHint.hidden = true;
      entryEmployeesSplitHint.textContent = "";
      entryEmployeesSplitHint.classList.remove("split-hint-error", "split-hint-ok");
    }
  }

  updateEntryMultiEmployeeSelectionState();
  syncSplitAllocationValidation();
}

function mergeEntryIntoState(entry, item) {
  const normalizedNote = String(entry.note || "").trim();
  const mergeIndex = state.entries.findIndex((existing) =>
    String(existing.customerId || "") === String(entry.customerId || "") &&
    String(existing.employeeId || "") === String(entry.employeeId || "") &&
    String(existing.itemId || "") === String(entry.itemId || "") &&
    String(existing.date || "") === String(entry.date || "") &&
    Math.round((Number(existing.unitPrice) || 0) * 100) === Math.round((Number(entry.unitPrice) || 0) * 100) &&
    String(existing.note || "").trim() === normalizedNote
  );

  if (mergeIndex >= 0) {
    const mergedEntry = { ...state.entries[mergeIndex] };
    const mergedQty = (Number(mergedEntry.quantity) || 0) + (Number(entry.quantity) || 0);
    mergedEntry.quantity = getItemUnitName(item) === "Stk" ? Math.trunc(mergedQty) : Number(mergedQty.toFixed(4));
    mergedEntry.unitPrice = entry.unitPrice;
    mergedEntry.costPrice = entry.costPrice;
    state.entries[mergeIndex] = mergedEntry;
    return true;
  }

  state.entries.push(entry);
  return false;
}

function resetEntryCustomerSelection() {
  entryCustomerSearch.value = "";
  clearSelectedCustomer();
  if (entrySplitEmployees) {
    entrySplitEmployees.value = "no";
    entrySplitEmployees.disabled = false;
  }
  entryEmployeeSearch.value = "";
  clearSelectedEmployee();
  clearSelectedSplitEmployeeIds();
  entryItemSearch.value = "";
  clearSelectedItem();
  renderEntrySelects();
}

function renderEntries() {
  const selectedDate = entryDate.value;
  const weekday = selectedDate ? formatWeekdayCH(selectedDate) : "";
  const heading = selectedDate
    ? `<h3 class="entry-day-title">Erfasste Leistungen für ${escapeHtml(weekday)}, ${escapeHtml(formatDateCH(selectedDate))}:</h3>`
    : `<h3 class="entry-day-title">Erfasste Leistungen</h3>`;
  if (entryListHeader) {
    entryListHeader.innerHTML = heading;
  }

  const customerFilterRaw = entryListCustomerSearch?.value?.trim() || "";
  const customerFilter = normalizeSearchText(customerFilterRaw);
  const visibleEntries = selectedDate
    ? state.entries.filter((e) => e.date === selectedDate)
    : state.entries;
  const filteredEntries = customerFilter
    ? visibleEntries.filter((e) => {
      const customer = state.customers.find((c) => c.id === e.customerId);
      const haystack = normalizeSearchText([
        customer?.company,
        customer?.firstName,
        customer?.lastName,
        customer?.street,
        customer?.zip,
        customer?.city,
        customer?.phone,
        customer?.email
      ].filter(Boolean).join(" "));
      return haystack.includes(customerFilter);
    })
    : visibleEntries;

  if (!filteredEntries.length) {
    if (selectedDate) {
      const filterSuffix = customerFilter ? ` mit Kundenfilter "${escapeHtml(customerFilterRaw)}"` : "";
      entryList.innerHTML = `<small>Keine Erfassungen für ${escapeHtml(formatDateCH(selectedDate))}${filterSuffix} vorhanden.</small>`;
      return;
    }
    if (customerFilter) {
      entryList.innerHTML = `<small>Keine Erfassungen für Kundenfilter "${escapeHtml(customerFilterRaw)}" vorhanden.</small>`;
      return;
    }
    entryList.innerHTML = `<small>Noch keine Erfassungen vorhanden.</small>`;
    return;
  }

  const sortedEntries = filteredEntries
    .slice()
    .sort((a, b) => a.customerId.localeCompare(b.customerId));

  const groups = new Map();
  sortedEntries.forEach((e) => {
    const customer = state.customers.find((c) => c.id === e.customerId);
    const customerName = customer
      ? `${customer.company ? `${customer.company} - ` : ""}${customer.firstName} ${customer.lastName}`
      : "Unbekannter Kunde";
    const key = customer ? customer.id : `unknown-${customerName}`;
    if (!groups.has(key)) {
      groups.set(key, { customerName, entries: [] });
    }
    groups.get(key).entries.push(e);
  });

  const listHtml = [...groups.values()]
    .map((group) => {
      const cards = group.entries
        .map((e) => {
          const item = state.items.find((i) => i.id === e.itemId);
          const employee = state.employees.find((emp) => emp.id === e.employeeId);
          const itemName = item ? `${item.name}${item.deleted ? " (gelöscht)" : ""}` : "Unbekannter Eintrag";
          const employeeName = employee ? `${employee.firstName || ""} ${employee.lastName || ""}`.trim() : "Unbekannter Mitarbeiter";
          const unitPrice = e.unitPrice != null ? e.unitPrice : resolveEntryUnitPrice(item);
          const total = unitPrice * e.quantity;
          return `
            <article class="card">
              <p class="entry-main">${escapeHtml(itemName)} | ${e.quantity} ${escapeHtml(getItemUnitName(item))}</p>
              <p class="entry-sub">${escapeHtml(formatDateCH(e.date))} | MA: ${escapeHtml(employeeName)} | Ans.: ${formatCurrency(unitPrice)} | Total: ${formatCurrency(total)}${e.note ? ` | ${escapeHtml(e.note)}` : ""}</p>
              <div class="card-actions">
                <button type="button" class="secondary" data-action="edit-entry" data-id="${escapeHtml(e.id)}">Bearbeiten</button>
                <button type="button" class="danger" data-action="delete-entry" data-id="${escapeHtml(e.id)}">Löschen</button>
              </div>
            </article>
          `;
        })
        .join("");

      return `
        <section class="entry-group">
          <h3>${escapeHtml(group.customerName)}</h3>
          ${cards}
        </section>
      `;
    })
    .join("");

  entryList.innerHTML = listHtml;
}
function setDefaultDate() {
  entryDate.value = new Date().toISOString().slice(0, 10);
}

function formToObject(form) {
  return Object.fromEntries(new FormData(form).entries());
}

function formatCurrency(value) {
  const currency = normalizeCurrency(state?.settings?.currency);
  try {
    return new Intl.NumberFormat("de-CH", {
      style: "currency",
      currency
    }).format(value || 0);
  } catch {
    const amount = new Intl.NumberFormat("de-CH", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value || 0);
    return `${currency} ${amount}`;
  }
}

function parseAmount(value) {
  if (value === "" || value == null) return null;
  const n = Number(value);
  return Number.isFinite(n) && n >= 0 ? n : null;
}

function resolveEntryUnitPrice(item) {
  if (!item) return 0;
  return Number(item.price) || 0;
}

function resolveEntryCostPrice(item) {
  if (!item) return null;
  return parseAmount(item.costPrice);
}

function normalizeEntryCostPrice(entry, item) {
  const snapshot = parseAmount(
    entry?.costPrice ??
    entry?.unitCostPrice ??
    entry?.costPriceSnapshot ??
    entry?.einstandspreis ??
    entry?.purchasePrice
  );
  if (snapshot != null) return snapshot;
  return resolveEntryCostPrice(item);
}



function wireReportOverview() {
  if (!reportMonth) return;
  const rerender = () => {
    renderReportOverview();
    renderProductReportOverview();
  };
  reportMonth.addEventListener("change", rerender);
  reportMonth.addEventListener("input", rerender);
  setDefaultReportMonth();
}

function setDefaultReportMonth() {
  if (!reportMonth) return;
  if (!String(reportMonth.value || "").trim()) {
    reportMonth.value = new Date().toISOString().slice(0, 7);
  }
}

function formatMonthCH(month) {
  const [year, m] = String(month || "").split("-");
  if (!year || !m) return String(month || "");
  return `${m}.${year}`;
}

function renderReportOverview() {
  if (!reportCalendar || !reportMonth) return;
  const month = String(reportMonth.value || "").trim();
  if (!month) {
    reportCalendar.innerHTML = "<small>Bitte Monat wählen.</small>";
    return;
  }

  const monthMatch = month.match(/^(\d{4})-(\d{2})$/);
  if (!monthMatch) {
    reportCalendar.innerHTML = "<small>Ungültiger Monat.</small>";
    return;
  }

  const year = Number(monthMatch[1]);
  const monthIndex = Number(monthMatch[2]) - 1;
  const lastDate = new Date(year, monthIndex + 1, 0).getDate();

  const byDay = new Map();
  const monthEntries = state.entries.filter((entry) => String(entry?.date || "").startsWith(`${month}-`));

  const getCustomerShort = (customer) => {
    if (!customer) return "Unbekannter Kunde";
    const name = `${customer.firstName || ""} ${customer.lastName || ""}`.trim();
    if (customer.company && name) return `${customer.company} - ${name}`;
    return customer.company || name || "Unbekannter Kunde";
  };

  const getEmployeeShort = (employee) => {
    if (!employee) return "Nicht zugewiesen";
    const name = `${employee.firstName || ""} ${employee.lastName || ""}`.trim();
    return name || "Nicht zugewiesen";
  };

  monthEntries.forEach((entry) => {
    const item = state.items.find((i) => i.id === entry.itemId);
    const unit = String(getItemUnitName(item) || "").trim().toLowerCase();
    if (!["h", "std", "stunde", "stunden", "stunde/n"].includes(unit)) return;

    const date = String(entry?.date || "").slice(0, 10);
    const customer = state.customers.find((c) => c.id === String(entry.customerId || ""));
    const customerName = getCustomerShort(customer);

    const employee = state.employees.find((e) => e.id === String(entry.employeeId || ""));
    const employeeName = getEmployeeShort(employee);

    const qty = Number(entry.quantity) || 0;
    const unitPrice = entry?.unitPrice != null ? (Number(entry.unitPrice) || 0) : (Number(resolveEntryUnitPrice(item)) || 0);
    const lineTotal = qty * unitPrice;

    if (!byDay.has(date)) byDay.set(date, new Map());
    const dayMap = byDay.get(date);

    const key = `${customerName}|||${employeeName}|||${unitPrice}`;
    const current = dayMap.get(key) || { customerName, employeeName, qty: 0, unitPrice, amount: 0 };
    current.qty += qty;
    current.amount += lineTotal;
    dayMap.set(key, current);
  });

  const shortWeekdays = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];

  const body = Array.from({ length: lastDate }, (_, offset) => {
    const day = offset + 1;
    const isoDate = `${year}-${String(monthIndex + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    const dayMap = byDay.get(isoDate) || new Map();
    const rows = [...dayMap.values()].sort((a, b) => a.customerName.localeCompare(b.customerName, "de-CH") || a.employeeName.localeCompare(b.employeeName, "de-CH"));
    const totalHours = rows.reduce((sum, row) => sum + (Number(row.qty) || 0), 0);
    const totalAmount = rows.reduce((sum, row) => sum + (Number(row.amount) || 0), 0);

    const weekdayIndex = new Date(year, monthIndex, day).getDay();
    const weekdayShort = shortWeekdays[weekdayIndex] || "";
    const isWeekend = weekdayIndex === 0 || weekdayIndex === 6;

    const rowsHtml = (() => {
      if (!rows.length) return "";
      let html = "";
      let currentCustomer = "";
      let customerHours = 0;
      let customerAmount = 0;
      const flushCustomerTotal = () => {
        if (!currentCustomer) return;
        html += `
          <li class="hours-calendar-row hours-money-row hours-customer-total">
            <span class="col-customer">Kundentotal</span>
            <span class="col-employee"></span>
            <strong class="col-hours">${escapeHtml(String((Math.round((customerHours || 0) * 100) / 100).toLocaleString("de-CH")))} h</strong>
            <strong class="col-rate"></strong>
            <strong class="col-total">${escapeHtml(formatCurrency(customerAmount))}</strong>
          </li>
        `;
      };
      rows.forEach((row, index) => {
        const customerChanged = currentCustomer && currentCustomer !== row.customerName;
        if (customerChanged) {
          flushCustomerTotal();
          html += `<li class="hours-customer-divider" aria-hidden="true"></li>`;
          customerHours = 0;
          customerAmount = 0;
        }
        if (!currentCustomer || currentCustomer !== row.customerName) {
          currentCustomer = row.customerName;
        }
        customerHours += Number(row.qty) || 0;
        customerAmount += Number(row.amount) || 0;
        html += `
          <li class="hours-calendar-row hours-money-row">
            <span class="col-customer">${escapeHtml(row.customerName)}</span>
            <span class="col-employee">${escapeHtml(row.employeeName)}</span>
            <strong class="col-hours">${escapeHtml(String((Math.round((row.qty || 0) * 100) / 100).toLocaleString("de-CH")))} h</strong>
            <strong class="col-rate">${escapeHtml(formatCurrency(row.unitPrice))}</strong>
            <strong class="col-total">${escapeHtml(formatCurrency(row.amount))}</strong>
          </li>
        `;
        if (index === rows.length - 1) {
          flushCustomerTotal();
        }
      });
      return html;
    })();

    const detailsHtml = rows.length
      ? `
        <div class="hours-calendar-cell-columns hours-money-columns"><span>Kunde</span><span>Mitarbeiter</span><span>Stunden</span><span>Ansatz</span><span>Total</span></div>
        <ul class="hours-calendar-list">${rowsHtml}</ul>
        <div class="hours-calendar-total hours-money-total hours-day-total">
          <span>Tagestotal</span>
          <strong>${escapeHtml(String((Math.round((totalHours || 0) * 100) / 100).toLocaleString("de-CH")))} h</strong>
          <strong>${escapeHtml(formatCurrency(totalAmount))}</strong>
        </div>
      `
      : `<div class="hours-day-empty">Keine Erfassung</div>`;

    const classes = ["hours-day-card"];
    if (rows.length) classes.push("has-data"); else classes.push("no-data");
    if (isWeekend) classes.push("is-weekend");

    return `
      <article class="${classes.join(" ")}">
        <div class="hours-day-meta">
          <span class="day-weekday">${escapeHtml(weekdayShort)}</span>
          <span class="day-no">${day}</span>
          <span class="day-date">${escapeHtml(formatDateCH(isoDate))}</span>
        </div>
        <div class="hours-day-content">
          ${detailsHtml}
        </div>
      </article>
    `;
  }).join("");

  reportCalendar.innerHTML = `
    <h3>Stunden ${escapeHtml(formatMonthCH(month))}</h3>
    <div class="hours-day-list">${body}</div>
  `;
}
function renderProductReportOverview() {
  if (!reportCalendarProducts || !reportMonth) return;
  const month = String(reportMonth.value || "").trim();
  if (!month) {
    reportCalendarProducts.innerHTML = "<small>Bitte Monat wählen.</small>";
    return;
  }

  const monthMatch = month.match(/^(\d{4})-(\d{2})$/);
  if (!monthMatch) {
    reportCalendarProducts.innerHTML = "<small>Ungültiger Monat.</small>";
    return;
  }

  const year = Number(monthMatch[1]);
  const monthIndex = Number(monthMatch[2]) - 1;
  const lastDate = new Date(year, monthIndex + 1, 0).getDate();

  const byDay = new Map();
  const monthEntries = state.entries.filter((entry) => String(entry?.date || "").startsWith(`${month}-`));

  const getCustomerShort = (customer) => {
    if (!customer) return "Unbekannter Kunde";
    const name = `${customer.firstName || ""} ${customer.lastName || ""}`.trim();
    if (customer.company && name) return `${customer.company} - ${name}`;
    return customer.company || name || "Unbekannter Kunde";
  };

  monthEntries.forEach((entry) => {
    const item = state.items.find((i) => i.id === entry.itemId);

    const normalizeUnitToken = (value) => String(value || "")
      .trim()
      .toLowerCase()
      .replaceAll("ä", "ae")
      .replaceAll("ö", "oe")
      .replaceAll("ü", "ue")
      .replaceAll(".", "")
      .replaceAll("/", "")
      .replaceAll(" ", "");

    const unitTokens = [
      item?.unitId,
      item?.unit,
      entry?.unitId,
      entry?.unit,
      getItemUnitName(item)
    ].map(normalizeUnitToken).filter(Boolean);

    const isPieceUnit = unitTokens.some((token) => ["unit-stk", "stk", "stueck", "stuck"].includes(token));
    if (!isPieceUnit) return;

    const date = String(entry?.date || "").slice(0, 10);
    const customer = state.customers.find((c) => c.id === String(entry.customerId || ""));
    const customerName = getCustomerShort(customer);
    const itemName = item?.name ? String(item.name) : "Unbekanntes Produkt";

    const qty = Number(entry.quantity) || 0;
    const unitPrice = entry?.unitPrice != null ? (Number(entry.unitPrice) || 0) : (Number(resolveEntryUnitPrice(item)) || 0);
    const amount = qty * unitPrice;

    if (!byDay.has(date)) byDay.set(date, new Map());
    const dayMap = byDay.get(date);

    const key = `${customerName}|||${itemName}|||${unitPrice}`;
    const current = dayMap.get(key) || { customerName, itemName, qty: 0, unitPrice, amount: 0 };
    current.qty += qty;
    current.amount += amount;
    dayMap.set(key, current);
  });

  const shortWeekdays = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];

  const body = Array.from({ length: lastDate }, (_, offset) => {
    const day = offset + 1;
    const isoDate = `${year}-${String(monthIndex + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    const dayMap = byDay.get(isoDate) || new Map();
    const rows = [...dayMap.values()].sort((a, b) => a.customerName.localeCompare(b.customerName, "de-CH") || a.itemName.localeCompare(b.itemName, "de-CH"));
    const totalQty = rows.reduce((sum, row) => sum + (Number(row.qty) || 0), 0);
    const totalAmount = rows.reduce((sum, row) => sum + (Number(row.amount) || 0), 0);

    const weekdayIndex = new Date(year, monthIndex, day).getDay();
    const weekdayShort = shortWeekdays[weekdayIndex] || "";
    const isWeekend = weekdayIndex === 0 || weekdayIndex === 6;

    const rowsHtml = (() => {
      if (!rows.length) return "";
      let html = "";
      let currentCustomer = "";
      let customerQty = 0;
      let customerAmount = 0;
      const flushCustomerTotal = () => {
        if (!currentCustomer) return;
        html += `
          <li class="hours-calendar-row hours-money-row products-row hours-customer-total">
            <span class="col-customer">Kundentotal</span>
            <span class="col-employee"></span>
            <strong class="col-hours">${escapeHtml(String((Math.round((customerQty || 0) * 100) / 100).toLocaleString("de-CH")))} Stk</strong>
            <strong class="col-rate"></strong>
            <strong class="col-total">${escapeHtml(formatCurrency(customerAmount))}</strong>
          </li>
        `;
      };
      rows.forEach((row, index) => {
        const customerChanged = currentCustomer && currentCustomer !== row.customerName;
        if (customerChanged) {
          flushCustomerTotal();
          html += `<li class="hours-customer-divider" aria-hidden="true"></li>`;
          customerQty = 0;
          customerAmount = 0;
        }
        if (!currentCustomer || currentCustomer !== row.customerName) {
          currentCustomer = row.customerName;
        }
        customerQty += Number(row.qty) || 0;
        customerAmount += Number(row.amount) || 0;
        html += `
          <li class="hours-calendar-row hours-money-row products-row">
            <span class="col-customer">${escapeHtml(row.customerName)}</span>
            <span class="col-employee">${escapeHtml(row.itemName)}</span>
            <strong class="col-hours">${escapeHtml(String((Math.round((row.qty || 0) * 100) / 100).toLocaleString("de-CH")))} Stk</strong>
            <strong class="col-rate">${escapeHtml(formatCurrency(row.unitPrice))}</strong>
            <strong class="col-total">${escapeHtml(formatCurrency(row.amount))}</strong>
          </li>
        `;
        if (index === rows.length - 1) {
          flushCustomerTotal();
        }
      });
      return html;
    })();

    const detailsHtml = rows.length
      ? `
        <div class="hours-calendar-cell-columns hours-money-columns products-columns"><span>Kunde</span><span>Produkt</span><span>Stk</span><span>Ansatz</span><span>Total</span></div>
        <ul class="hours-calendar-list">${rowsHtml}</ul>
        <div class="hours-calendar-total hours-money-total products-total hours-day-total">
          <span>Tagestotal</span>
          <strong>${escapeHtml(String((Math.round((totalQty || 0) * 100) / 100).toLocaleString("de-CH")))} Stk</strong>
          <strong>${escapeHtml(formatCurrency(totalAmount))}</strong>
        </div>
      `
      : `<div class="hours-day-empty">Keine Erfassung</div>`;

    const classes = ["hours-day-card"];
    if (rows.length) classes.push("has-data"); else classes.push("no-data");
    if (isWeekend) classes.push("is-weekend");

    return `
      <article class="${classes.join(" ")}">
        <div class="hours-day-meta">
          <span class="day-weekday">${escapeHtml(weekdayShort)}</span>
          <span class="day-no">${day}</span>
          <span class="day-date">${escapeHtml(formatDateCH(isoDate))}</span>
        </div>
        <div class="hours-day-content">
          ${detailsHtml}
        </div>
      </article>
    `;
  }).join("");

  reportCalendarProducts.innerHTML = `
    <h3>Produkte ${escapeHtml(formatMonthCH(month))}</h3>
    <div class="hours-day-list">${body}</div>
  `;
}
function formatDateCH(isoDate) {
  if (!isoDate) return "";
  const parts = String(isoDate).split("-");
  if (parts.length !== 3) return String(isoDate);
  const [year, month, day] = parts;
  if (!year || !month || !day) return String(isoDate);
  return `${day}.${month}.${year}`;
}

function formatWeekdayCH(isoDate) {
  const parsed = new Date(String(isoDate));
  if (Number.isNaN(parsed.getTime())) return "";
  return new Intl.DateTimeFormat("de-CH", { weekday: "long" }).format(parsed);
}
function generateId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(36).slice(2, 11)}`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function renderCustomerSuggestions(queryValue) {
  const query = normalizeSearchText(queryValue);
  const filtered = state.customers.filter((c) => {
    if (c.deleted) return false;
    if (!query) return true;
    const haystack = normalizeSearchText([
      c.company,
      c.firstName,
      c.lastName,
      c.street,
      c.zip,
      c.city,
      c.phone,
      c.email
    ]
      .join(" "));
    return haystack.includes(query);
  });

  customerSuggestionMap = new Map();
  filtered.forEach((c) => {
    const display = getCustomerDisplay(c);
    customerSuggestionMap.set(display, c.id);
    customerSuggestionMap.set(normalizeSearchText(display), c.id);
  });

  customerSuggestions.innerHTML = filtered
    .slice(0, 30)
    .map((c) => `<option value="${escapeHtml(getCustomerDisplay(c))}"></option>`)
    .join("");
}

function renderEmployeeSuggestions(queryValue) {
  const query = normalizeSearchText(queryValue);
  const filtered = state.employees.filter((e) => {
    if (e.deleted) return false;
    if (!query) return true;
    const haystack = normalizeSearchText([
      e.firstName,
      e.lastName,
      e.street,
      e.zip,
      e.city,
      e.phone,
      e.email
    ].join(" "));
    return haystack.includes(query);
  });

  employeeSuggestionMap = new Map();
  filtered.forEach((e) => {
    const display = getEmployeeDisplay(e);
    employeeSuggestionMap.set(display, e.id);
    employeeSuggestionMap.set(normalizeSearchText(display), e.id);
  });

  employeeSuggestions.innerHTML = filtered
    .slice(0, 30)
    .map((e) => `<option value="${escapeHtml(getEmployeeDisplay(e))}"></option>`)
    .join("");
}


function renderItemSuggestions(queryValue) {
  const query = normalizeSearchText(queryValue);
  const filtered = state.items.filter((item) => {
    if (item.deleted) return false;
    if (!query) return true;
    const haystack = normalizeSearchText([
      item.name,
      getItemTypeName(item),
      getItemUnitName(item),
      String(item.price),
      String(item.costPrice ?? "")
    ].join(" "));
    return haystack.includes(query);
  });

  itemSuggestionMap = new Map();
  filtered.forEach((item) => {
    const display = getItemDisplay(item);
    itemSuggestionMap.set(display, item.id);
    itemSuggestionMap.set(normalizeSearchText(display), item.id);
  });

  itemSuggestions.innerHTML = filtered
    .slice(0, 30)
    .map((item) => `<option value="${escapeHtml(getItemDisplay(item))}"></option>`)
    .join("");
}

function getItemIdFromSuggestion(inputValue) {
  const normalizedInput = (inputValue || "").trim();
  if (!normalizedInput) return "";

  if (itemSuggestionMap.has(normalizedInput)) {
    return itemSuggestionMap.get(normalizedInput) || "";
  }

  const normalized = normalizeSearchText(normalizedInput);
  if (itemSuggestionMap.has(normalized)) {
    return itemSuggestionMap.get(normalized) || "";
  }

  for (const [label, id] of itemSuggestionMap.entries()) {
    if (normalizeSearchText(label) === normalized) return id;
  }

  return "";
}

function getItemDisplay(item) {
  const unit = getItemUnitName(item);
  return `${item.name} (${formatCurrency(item.price)} / ${unit})`;
}
function getCustomerIdFromSuggestion(inputValue) {
  const normalizedInput = (inputValue || "").trim();
  if (!normalizedInput) return "";

  if (customerSuggestionMap.has(normalizedInput)) {
    return customerSuggestionMap.get(normalizedInput) || "";
  }

  const normalized = normalizeSearchText(normalizedInput);
  if (customerSuggestionMap.has(normalized)) {
    return customerSuggestionMap.get(normalized) || "";
  }

  for (const [label, id] of customerSuggestionMap.entries()) {
    if (normalizeSearchText(label) === normalized) return id;
  }

  return "";
}

function getEmployeeIdFromSuggestion(inputValue) {
  const normalizedInput = (inputValue || "").trim();
  if (!normalizedInput) return "";

  if (employeeSuggestionMap.has(normalizedInput)) {
    return employeeSuggestionMap.get(normalizedInput) || "";
  }

  const normalized = normalizeSearchText(normalizedInput);
  if (employeeSuggestionMap.has(normalized)) {
    return employeeSuggestionMap.get(normalized) || "";
  }

  for (const [label, id] of employeeSuggestionMap.entries()) {
    if (normalizeSearchText(label) === normalized) return id;
  }

  return "";
}

function getCustomerDisplay(customer) {
  const lead = `${customer.company ? `${customer.company} - ` : ""}${customer.firstName} ${customer.lastName}`.trim();
  const area = [customer.zip, customer.city].filter(Boolean).join(" ");
  return area ? `${lead} (${area})` : lead;
}

function getEmployeeDisplay(employee) {
  const lead = `${employee.firstName || ""} ${employee.lastName || ""}`.trim();
  const area = [employee.zip, employee.city].filter(Boolean).join(" ");
  return area ? `${lead} (${area})` : lead;
}

function normalizeSearchText(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

function normalizeCatalogName(value) {
  return String(value || "").trim();
}

function normalizeNamedCatalog(values, fallbackPrefix) {
  const list = [];
  const seen = new Set();
  for (const raw of values || []) {
    const name = normalizeCatalogName(raw?.name || raw?.label || raw);
    if (!name) continue;
    const key = normalizeSearchText(name);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    list.push({ id: String(raw?.id || generateId()), name, deleted: raw?.deleted === true || String(raw?.deleted || "").toLowerCase() === "true" });
  }
  if (!list.length && fallbackPrefix) {
    list.push({ id: generateId(), name: fallbackPrefix, deleted: false });
  }
  return list;
}

function findCatalogByName(list, name) {
  const key = normalizeSearchText(name);
  return (list || []).find((item) => normalizeSearchText(item.name) === key) || null;
}

function ensureCatalogContains(list, name) {
  const normalized = normalizeCatalogName(name);
  let found = findCatalogByName(list, normalized);
  if (!found) {
    found = { id: generateId(), name: normalized, deleted: false };
    list.push(found);
  }
  if (found.deleted) found.deleted = false;
  return found;
}

function findOrCreateCatalogByName(list, name) {
  return ensureCatalogContains(list, name);
}

function getCatalogNameById(list, id) {
  return (list || []).find((item) => item.id === id)?.name || "";
}

function getItemTypeName(item) {
  if (!item) return "";
  if (item.typeId) {
    const resolved = getCatalogNameById(state.productTypes, item.typeId);
    if (resolved) return resolved;
  }
  return normalizeCatalogName(item.type);
}

function getItemUnitName(item) {
  if (!item) return "";
  if (item.unitId) {
    const resolved = getCatalogNameById(state.units, item.unitId);
    if (resolved) return resolved;
  }
  return normalizeCatalogName(item.unit);
}

function setSelectedCustomer(customerId) {
  const customer = state.customers.find((c) => !c.deleted && c.id === customerId);
  if (!customer) {
    clearSelectedCustomer();
    return;
  }
  entryCustomer.value = customer.id;
  entryCustomerDisplay.value = getCustomerDisplay(customer);
  refreshRequiredFieldStates();
}

function setSelectedEmployee(employeeId) {
  const employee = state.employees.find((e) => !e.deleted && e.id === employeeId);
  if (!employee) {
    clearSelectedEmployee();
    return;
  }
  entryEmployee.value = employee.id;
  entryEmployeeDisplay.value = getEmployeeDisplay(employee);
  refreshRequiredFieldStates();
}


function setSelectedItem(itemId) {
  const item = state.items.find((i) => i.id === itemId);
  if (!item) {
    clearSelectedItem();
    return;
  }
  entryItem.value = item.id;
  entryItemDisplay.value = getItemDisplay(item);
  updateEntryQuantityConstraints();
  refreshRequiredFieldStates();
}

function clearSelectedItem() {
  entryItem.value = "";
  entryItemDisplay.value = "";
  updateEntryQuantityConstraints();
  refreshRequiredFieldStates();
}
function clearSelectedCustomer() {
  entryCustomer.value = "";
  entryCustomerDisplay.value = "";
  refreshRequiredFieldStates();
}

function clearSelectedEmployee() {
  entryEmployee.value = "";
  entryEmployeeDisplay.value = "";
  refreshRequiredFieldStates();
}

function ensureCleaningServiceExists() {
  enforceFixedUnitsCatalog();
  ensureCatalogContains(state.productTypes, "Dienstleistung");
  ensureCatalogContains(state.productTypes, "Produkt");

  state.items.forEach((item) => {
    item.typeId = item.typeId || findOrCreateCatalogByName(state.productTypes, item.type || "Dienstleistung").id;
    if (item.unitId !== "unit-h" && item.unitId !== "unit-stk") item.unitId = "unit-h";
    item.type = getItemTypeName(item);
    item.unit = item.unitId === "unit-stk" ? "Stk" : "h";
    item.costPrice = parseAmount(item.costPrice ?? item.purchasePrice ?? item.einstandspreis);
  });
}

function normalizeCurrency(value) {
  const code = String(value || "").trim().toUpperCase();
  return /^[A-Z]{3}$/.test(code) ? code : "CHF";
}

function normalizeMonths(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 12;
  return Math.min(24, Math.max(1, Math.round(n)));
}

function pruneOldEntriesForExport() {
  const plan = listExportCleanupPlan();
  const entryIdsToRemove = new Set(plan.entries.map((entry) => entry.id));
  const customerIdsToRemove = new Set(plan.customers.map((customer) => customer.id));
  const employeeIdsToRemove = new Set(plan.employees.map((employee) => employee.id));
  const itemIdsToRemove = new Set(plan.items.map((item) => item.id));
  const unitIdsToRemove = new Set(plan.units.map((unit) => unit.id));
  const typeIdsToRemove = new Set(plan.productTypes.map((type) => type.id));

  const hasChanges = entryIdsToRemove.size || customerIdsToRemove.size || employeeIdsToRemove.size || itemIdsToRemove.size || unitIdsToRemove.size || typeIdsToRemove.size;
  if (!hasChanges) {
    return { entries: [], customers: [], employees: [], items: [], units: [], productTypes: [] };
  }

  if (entryIdsToRemove.size) {
    state.entries = state.entries.filter((entry) => !entryIdsToRemove.has(entry.id));
  }
  if (customerIdsToRemove.size) {
    state.customers = state.customers.filter((customer) => !customerIdsToRemove.has(customer.id));
  }
  if (employeeIdsToRemove.size) {
    state.employees = state.employees.filter((employee) => !employeeIdsToRemove.has(employee.id));
  }
  if (itemIdsToRemove.size) {
    state.items = state.items.filter((item) => !itemIdsToRemove.has(item.id));
  }
  if (unitIdsToRemove.size) {
    state.units = state.units.filter((unit) => !unitIdsToRemove.has(unit.id));
  }
  if (typeIdsToRemove.size) {
    state.productTypes = state.productTypes.filter((type) => !typeIdsToRemove.has(type.id));
  }

  saveState();
  renderAll();

  return plan;
}

function listExportCleanupPlan() {
  const entries = listOldEntriesForExport();
  const entryIdsToRemove = new Set(entries.map((entry) => entry.id));
  const remainingEntries = state.entries.filter((entry) => !entryIdsToRemove.has(entry.id));

  const usedCustomerIds = new Set(remainingEntries.map((entry) => String(entry.customerId || "")).filter(Boolean));
  const usedEmployeeIds = new Set(remainingEntries.map((entry) => String(entry.employeeId || "")).filter(Boolean));
  const usedItemIds = new Set(remainingEntries.map((entry) => String(entry.itemId || "")).filter(Boolean));
  const customers = state.customers.filter((customer) => customer.deleted && !usedCustomerIds.has(String(customer.id || "")));
  const employees = state.employees.filter((employee) => employee.deleted && !usedEmployeeIds.has(String(employee.id || "")));
  const items = state.items.filter((item) => item.deleted && !usedItemIds.has(String(item.id || "")));
  const itemIdsToRemove = new Set(items.map((item) => item.id));

  const remainingItems = state.items.filter((item) => !itemIdsToRemove.has(item.id));
  const usedUnitIds = new Set(remainingItems.map((item) => String(item.unitId || "")).filter(Boolean));
  const usedTypeIds = new Set(remainingItems.map((item) => String(item.typeId || "")).filter(Boolean));

  const units = state.units.filter((unit) => unit.deleted && !usedUnitIds.has(String(unit.id || "")));
  const productTypes = state.productTypes.filter((type) => type.deleted && !usedTypeIds.has(String(type.id || "")));

  return { entries, customers, employees, items, units, productTypes };
}

function listOldEntriesForExport() {
  const months = normalizeMonths(state?.settings?.retentionMonths);
  const cutoff = new Date();
  cutoff.setDate(1);
  cutoff.setMonth(cutoff.getMonth() - months);
  return state.entries.filter((entry) => {
    const raw = String(entry?.date || "").trim();
    if (!raw) return false;
    const parsed = new Date(raw);
    if (Number.isNaN(parsed.getTime())) return false;
    const entryMonth = new Date(parsed.getFullYear(), parsed.getMonth(), 1);
    return entryMonth < cutoff;
  });
}

function renderExportCleanupPreview() {
  renderExportCleanupResult(listExportCleanupPlan());
}

function renderExportCleanupResult(plan) {
  const cleanupPlan = plan || { entries: [], customers: [], employees: [], items: [], units: [], productTypes: [] };
  const removedEntries = Array.isArray(cleanupPlan.entries) ? cleanupPlan.entries : [];
  const removedCustomers = Array.isArray(cleanupPlan.customers) ? cleanupPlan.customers : [];
  const removedEmployees = Array.isArray(cleanupPlan.employees) ? cleanupPlan.employees : [];
  const removedItems = Array.isArray(cleanupPlan.items) ? cleanupPlan.items : [];
  const removedUnits = Array.isArray(cleanupPlan.units) ? cleanupPlan.units : [];
  const removedTypes = Array.isArray(cleanupPlan.productTypes) ? cleanupPlan.productTypes : [];
  const months = normalizeMonths(state?.settings?.retentionMonths);

  if (!removedEntries.length && !removedCustomers.length && !removedEmployees.length && !removedItems.length && !removedUnits.length && !removedTypes.length) {
    exportCleanupResult.hidden = true;
    exportCleanupList.innerHTML = "";
    return;
  }

  const sections = [];

  if (removedEntries.length) {
    const entriesHtml = removedEntries
      .slice()
      .sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")))
      .map((entry) => {
        const customer = state.customers.find((c) => c.id === entry.customerId);
        const item = state.items.find((i) => i.id === entry.itemId);
        const customerName = customer ? getCustomerDisplay(customer) : "Unbekannter Kunde";
        const itemLabel = item ? item.name : "Unbekannte Leistung";
        const unit = getItemUnitName(item);
        const qty = Number(entry.quantity) || 0;
        return `
          <article class="cleanup-result-row">
            <strong>${escapeHtml(formatDateCH(entry.date || ""))}</strong><br>
            ${escapeHtml(customerName)}<br>
            ${escapeHtml(itemLabel)} | ${escapeHtml(String(qty))} ${escapeHtml(unit)}
          </article>
        `;
      })
      .join("");
    sections.push(`<details class="cleanup-details"><summary>Erfasste Leistungen, die beim Export gelöscht werden da sie älter als ${months} Monate sind (Einstellung Datenbereinigung) (${removedEntries.length})</summary>${entriesHtml}</details>`);
  }

  if (removedCustomers.length) {
    const customerHtml = removedCustomers
      .map((customer) => `<article class="cleanup-result-row">${escapeHtml(getCustomerDisplay(customer) || "Unbekannter Kunde")}</article>`)
      .join("");
    sections.push(`<details class="cleanup-details"><summary>Gelöschte Kunden ohne Abhängigkeiten (${removedCustomers.length})</summary>${customerHtml}</details>`);
  }

  if (removedEmployees.length) {
    const employeeHtml = removedEmployees
      .map((employee) => `<article class="cleanup-result-row">${escapeHtml(getEmployeeDisplay(employee) || "Unbekannter Mitarbeiter")}</article>`)
      .join("");
    sections.push(`<details class="cleanup-details"><summary>Gelöschte Mitarbeiter ohne Abhängigkeiten (${removedEmployees.length})</summary>${employeeHtml}</details>`);
  }

  if (removedItems.length) {
    const itemHtml = removedItems
      .map((item) => `<article class="cleanup-result-row">${escapeHtml(item.name || "Unbekanntes Produkt")}</article>`)
      .join("");
    sections.push(`<details class="cleanup-details"><summary>Gelöschte Produkte ohne Abhängigkeiten (${removedItems.length})</summary>${itemHtml}</details>`);
  }

  if (removedUnits.length) {
    const unitHtml = removedUnits
      .map((unit) => `<article class="cleanup-result-row">${escapeHtml(unit.name || "Unbekannte Einheit")}</article>`)
      .join("");
    sections.push(`<details class="cleanup-details"><summary>Gelöschte Einheiten ohne Abhängigkeiten (${removedUnits.length})</summary>${unitHtml}</details>`);
  }

  if (removedTypes.length) {
    const typeHtml = removedTypes
      .map((type) => `<article class="cleanup-result-row">${escapeHtml(type.name || "Unbekannter Produkttyp")}</article>`)
      .join("");
    sections.push(`<details class="cleanup-details"><summary>Gelöschte Produkttypen ohne Abhängigkeiten (${removedTypes.length})</summary>${typeHtml}</details>`);
  }

  exportCleanupResult.hidden = false;
  exportCleanupList.innerHTML = sections.join("");
}


































































































































































































































