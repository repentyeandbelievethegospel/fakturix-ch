const STORAGE_KEY = "fakturix_ch_erfassung_data_v1";
const LEGACY_STORAGE_KEY = "putzpro_mobile_data_v1";

const state = loadState();
let editingCustomerId = null;
let editingEmployeeId = null;
let editingItemId = null;
let editingEntryId = null;
let customerSuggestionMap = new Map();
let employeeSuggestionMap = new Map();
ensureCleaningServiceExists();

const tabs = [...document.querySelectorAll(".tabs-main .tab")];
const panels = [...document.querySelectorAll(".panel")];

const customerForm = document.getElementById("customerForm");
const employeeForm = document.getElementById("employeeForm");
const itemForm = document.getElementById("itemForm");
const entryForm = document.getElementById("entryForm");
const settingsForm = document.getElementById("settingsForm");

const customerSubmitBtn = document.getElementById("customerSubmitBtn");
const customerCancelBtn = document.getElementById("customerCancelBtn");
const employeeSubmitBtn = document.getElementById("employeeSubmitBtn");
const employeeCancelBtn = document.getElementById("employeeCancelBtn");
const itemSubmitBtn = document.getElementById("itemSubmitBtn");
const itemCancelBtn = document.getElementById("itemCancelBtn");
const entrySubmitBtn = document.getElementById("entrySubmitBtn");
const entryCancelBtn = document.getElementById("entryCancelBtn");

const customerList = document.getElementById("customerList");
const customerSearch = document.getElementById("customerSearch");
const customerSearchClear = document.getElementById("customerSearchClear");
const employeeList = document.getElementById("employeeList");
const employeeSearch = document.getElementById("employeeSearch");
const employeeSearchClear = document.getElementById("employeeSearchClear");
const itemList = document.getElementById("itemList");
const itemSearch = document.getElementById("itemSearch");
const itemSearchClear = document.getElementById("itemSearchClear");
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
const entryItem = document.getElementById("entryItem");
const entryDate = document.getElementById("entryDate");
const entryQuantityInput = entryForm.querySelector('input[name="quantity"]');

const exportBtn = document.getElementById("exportBtn");
const exportStatus = document.getElementById("exportStatus");
const exportCleanupResult = document.getElementById("exportCleanupResult");
const exportCleanupList = document.getElementById("exportCleanupList");
const importBtn = document.getElementById("importBtn");
const importFile = document.getElementById("importFile");
const importStatus = document.getElementById("importStatus");
const retentionMonths = document.getElementById("retentionMonths");
const settingsStatus = document.getElementById("settingsStatus");
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
wireSettings();
wireRequiredFieldStates();
renderAll();

function loadState() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY) || localStorage.getItem(LEGACY_STORAGE_KEY);
    if (!saved) return { customers: [], employees: [], items: [], entries: [], settings: { retentionMonths: 12 } };
    const parsed = JSON.parse(saved);
    return {
      customers: Array.isArray(parsed.customers) ? parsed.customers : [],
      employees: Array.isArray(parsed.employees) ? parsed.employees : [],
      items: Array.isArray(parsed.items) ? parsed.items : [],
      entries: Array.isArray(parsed.entries) ? parsed.entries : [],
      settings: {
        retentionMonths: normalizeMonths(parsed?.settings?.retentionMonths)
      }
    };
  } catch {
    return { customers: [], employees: [], items: [], entries: [], settings: { retentionMonths: 12 } };
  }
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
    });
  });
}

function wireSubTabs() {
  initSubTabGroup(".panel#stammdaten");
  initSubTabGroup(".panel#wartung");
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
      email: data.email?.trim() || ""
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
      iban: data.iban?.trim() || "",
      hourlyWage: parseAmount(data.hourlyWage)
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

    const item = {
      id: editingItemId || generateId(),
      name: data.name.trim(),
      type: data.type,
      unit: data.unit.trim(),
      price: Number(data.price)
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

  entryCancelBtn?.addEventListener("click", () => {
    resetEntryForm();
  });

  entryDate.addEventListener("change", () => {
    renderEntries();
  });

  entryDate.addEventListener("input", () => {
    renderEntries();
  });

  entryItem.addEventListener("change", () => {
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

  entryForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const data = formToObject(entryForm);
    const isEditingEntry = Boolean(editingEntryId);
    if (!data.customerId) {
      alert("Bitte zuerst einen Kunden auswählen.");
      return;
    }
    if (!data.employeeId) {
      alert("Bitte zuerst einen Mitarbeiter auswählen.");
      return;
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
    if (item?.unit === "Stück" && !Number.isInteger(quantity)) {
      alert("Bei Einheit 'Stück' sind nur ganze Zahlen erlaubt.");
      entryForm.quantity.focus();
      return;
    }

    const existingEntry = editingEntryId ? state.entries.find((e) => e.id === editingEntryId) : null;
    const resolvedUnitPrice = resolveEntryUnitPrice(item);
    const unitPrice = existingEntry && existingEntry.itemId === data.itemId && existingEntry.unitPrice != null
      ? existingEntry.unitPrice
      : resolvedUnitPrice;

    const entry = {
      id: editingEntryId || generateId(),
      customerId: data.customerId,
      employeeId: data.employeeId,
      itemId: data.itemId,
      date: data.date,
      quantity: item?.unit === "Stück" ? Math.trunc(quantity) : quantity,
      note: data.note?.trim() || "",
      unitPrice
    };

    if (editingEntryId) {
      const index = state.entries.findIndex((e) => e.id === editingEntryId);
      if (index >= 0) state.entries[index] = entry;
      else state.entries.push(entry);
    } else {
      state.entries.push(entry);
    }

    saveState();
    entryForm.quantity.value = "";
    entryForm.note.value = "";
    entryItem.value = "";
    clearSelectedCustomer();
    entryCustomerSearch.value = "";
    clearSelectedEmployee();
    entryEmployeeSearch.value = "";
    renderCustomerSuggestions("");
    renderEmployeeSuggestions("");
    resetEntryForm();
    updateEntryQuantityConstraints();
    refreshRequiredFieldStates();
    renderEntries();
    showFlashMessage(isEditingEntry ? "Leistung aktualisiert." : "Leistung gespeichert.");
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
    const removedEntries = pruneOldEntriesForExport();
    const removedCount = removedEntries.length;
      const payload = {
        app: "Fakturix CH Leistungserfassung",
        exportedAt: new Date().toISOString(),
        customers: state.customers,
        employees: state.employees,
        items: state.items,
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
                exportStatus.textContent = `Export über Teilen-Menü erfolgreich. Alte Erfassungen gelöscht: ${removedCount}.`;
              renderExportCleanupPreview();
              return;
            } catch (shareError) {
              console.warn("Share fehlgeschlagen, wechsle auf Download-Fallback.", shareError);
            }
          }
        }

        downloadBlob(blob, fileName);
        exportStatus.textContent = `Datei "${fileName}" wurde heruntergeladen oder zum Speichern angeboten. Alte Erfassungen gelöscht: ${removedCount}.`;
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
    state.settings.retentionMonths = normalizeMonths(retentionMonths.value);
    saveState();
    settingsStatus.textContent = `Gespeichert: Erfassungen werden ${state.settings.retentionMonths} Monate aufbewahrt.`;
    renderExportCleanupPreview();
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
        state.settings = normalized.settings;
        ensureCleaningServiceExists();

      resetCustomerForm();
      resetEmployeeForm();
      resetItemForm();
      saveState();
      renderAll();

      importStatus.textContent = `Import erfolgreich: ${state.customers.length} Kunden, ${state.employees.length} Mitarbeiter, ${state.items.length} Einträge, ${state.entries.length} Erfassungen.`;
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

  const confirmed = confirm("Kunde löschen? Zugehörige Erfassungen werden ebenfalls gelöscht.");
  if (!confirmed) return;

  state.customers = state.customers.filter((c) => c.id !== customerId);
  state.entries = state.entries.filter((e) => e.customerId !== customerId);

  if (editingCustomerId === customerId) {
    resetCustomerForm();
  }

  saveState();
  renderAll();
}

function deleteEmployee(employeeId) {
  const employee = state.employees.find((e) => e.id === employeeId);
  if (!employee) return;

  const confirmed = confirm("Mitarbeiter löschen?");
  if (!confirmed) return;

  state.employees = state.employees.filter((e) => e.id !== employeeId);
  state.entries = state.entries.filter((entry) => entry.employeeId !== employeeId);

  if (editingEmployeeId === employeeId) {
    resetEmployeeForm();
  }

  saveState();
  renderAll();
}

function editItem(itemId) {
  const item = state.items.find((i) => i.id === itemId);
  if (!item) return;

  itemForm.name.value = item.name || "";
  itemForm.type.value = item.type || "Dienstleistung";
  itemForm.unit.value = item.unit === "Stück" ? "Stück" : "Stunde/n";
  itemForm.price.value = String(item.price ?? "");

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

  const confirmed = confirm("Dienstleistung/Produkt löschen? Zugehörige Erfassungen werden ebenfalls gelöscht.");
  if (!confirmed) return;

  state.items = state.items.filter((i) => i.id !== itemId);
  state.entries = state.entries.filter((e) => e.itemId !== itemId);
  ensureCleaningServiceExists();

  if (editingItemId === itemId) {
    resetItemForm();
  }

  saveState();
  renderAll();
}

function editEntry(entryId) {
  const entry = state.entries.find((e) => e.id === entryId);
  if (!entry) return;
  setSelectedCustomer(entry.customerId);
  setSelectedEmployee(entry.employeeId);
  entryCustomerSearch.value = entryCustomerDisplay.value || "";
  entryEmployeeSearch.value = entryEmployeeDisplay.value || "";
  entryItem.value = String(entry.itemId || "");
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
  itemForm.unit.value = "Stunde/n";
  itemSubmitBtn.textContent = "Eintrag speichern";
  itemCancelBtn.hidden = true;
  refreshRequiredFieldStates();
}

function resetEntryForm() {
  editingEntryId = null;
  entrySubmitBtn.textContent = "Erfassung speichern";
  if (entryCancelBtn) entryCancelBtn.hidden = true;
  entryItem.value = "";
  entryForm.quantity.value = "";
  entryForm.note.value = "";
  clearSelectedCustomer();
  entryCustomerSearch.value = "";
  clearSelectedEmployee();
  entryEmployeeSearch.value = "";
  renderCustomerSuggestions("");
  renderEmployeeSuggestions("");
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

  const customers = importedCustomers.map((c) => ({
    id: c?.id || generateId(),
    company: String(c?.company || ""),
    firstName: String(c?.firstName || ""),
    lastName: String(c?.lastName || ""),
    street: String(c?.street || ""),
    zip: String(c?.zip || ""),
    city: String(c?.city || ""),
    phone: String(c?.phone || ""),
    email: String(c?.email || "")
  }));

  const items = importedItems.map((i) => ({
    id: i?.id || generateId(),
    name: String(i?.name || ""),
    type: i?.type === "Produkt" ? "Produkt" : "Dienstleistung",
    unit: i?.unit === "Stück" ? "Stück" : "Stunde/n",
    price: Number(i?.price) || 0
  }));

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
    hourlyWage: parseAmount(e?.hourlyWage)
  }));

  const customerIds = new Set(customers.map((c) => c.id));
  const employeeIds = new Set(employees.map((e) => e.id));
  const itemIds = new Set(items.map((i) => i.id));
  const today = new Date().toISOString().slice(0, 10);

  const entries = importedEntries
    .map((e) => ({
      id: e?.id || generateId(),
      customerId: String(e?.customerId || ""),
      employeeId: String(e?.employeeId || ""),
      itemId: String(e?.itemId || ""),
      date: String(e?.date || today),
      quantity: Number(e?.quantity) > 0 ? Number(e.quantity) : 1,
      note: String(e?.note || ""),
      unitPrice: parseAmount(e?.unitPrice)
    }))
    .filter((e) => {
      const employeeOk = !e.employeeId || employeeIds.has(e.employeeId);
      return customerIds.has(e.customerId) && itemIds.has(e.itemId) && employeeOk;
    });

  const settings = {
    retentionMonths: normalizeMonths(parsed?.settings?.retentionMonths)
  };

  return { customers, employees, items, entries, settings };
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
  renderEmployees();
  renderItems();
  renderEntrySelects();
  renderEntries();
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
}

function refreshRequiredFieldStates() {
  const requiredFields = document.querySelectorAll("form input[required], form select[required], form textarea[required]");
  requiredFields.forEach((field) => updateRequiredFieldState(field));
  updateEntrySelectionDisplayStates();
  syncEntrySubmitButtonState();
}

function updateRequiredFieldState(field) {
  const rawValue = String(field?.value ?? "").trim();
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

function updateEntrySelectionDisplayStates() {
  setDisplayFieldState(entryCustomerDisplay, String(entryCustomer.value || "").trim().length > 0);
  setDisplayFieldState(entryEmployeeDisplay, String(entryEmployee.value || "").trim().length > 0);
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
  return isFormReadyForSubmit(entryForm);
}

function syncEntrySubmitButtonState() {
  if (entrySubmitBtn) entrySubmitBtn.disabled = !isFormReadyForSubmit(entryForm);
  if (customerSubmitBtn) customerSubmitBtn.disabled = !isFormReadyForSubmit(customerForm);
  if (employeeSubmitBtn) employeeSubmitBtn.disabled = !isFormReadyForSubmit(employeeForm);
  if (itemSubmitBtn) itemSubmitBtn.disabled = !isFormReadyForSubmit(itemForm);
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
  retentionMonths.value = String(normalizeMonths(state?.settings?.retentionMonths));
}

function renderCustomers() {
  if (!state.customers.length) {
    customerList.innerHTML = "<small>Noch keine Kunden erfasst.</small>";
    return;
  }

  const query = normalizeSearchText(customerSearch.value);
  const visibleCustomers = state.customers.filter((c) => {
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
    customerList.innerHTML = "<small>Keine Kunden passend zur Suche gefunden.</small>";
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

function renderEmployees() {
  if (!state.employees.length) {
    employeeList.innerHTML = "<small>Noch keine Mitarbeiter erfasst.</small>";
    return;
  }

  const query = normalizeSearchText(employeeSearch.value);
  const visibleEmployees = state.employees.filter((e) => {
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
    employeeList.innerHTML = "<small>Keine Mitarbeiter passend zur Suche gefunden.</small>";
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

function renderItems() {
  if (!state.items.length) {
    itemList.innerHTML = "<small>Noch keine Dienstleistungen/Produkte erfasst.</small>";
    return;
  }

  const query = normalizeSearchText(itemSearch.value);
  const visibleItems = state.items.filter((i) => {
    if (!query) return true;
    const haystack = normalizeSearchText([
      i.name,
      i.type,
      i.unit,
      String(i.price)
    ].join(" "));
    return haystack.includes(query);
  });

  if (!visibleItems.length) {
    itemList.innerHTML = "<small>Keine Dienstleistungen/Produkte passend zur Suche gefunden.</small>";
    return;
  }

  itemList.innerHTML = visibleItems
    .map(
      (i) => `
      <article class="card">
        <strong>${escapeHtml(i.name)}</strong>
        <small>${escapeHtml(i.type)} | ${formatCurrency(i.price)} / ${escapeHtml(i.unit)}</small>
        <div class="card-actions">
          <button type="button" class="secondary" data-action="edit" data-id="${escapeHtml(i.id)}">Bearbeiten</button>
          <button type="button" class="danger" data-action="delete" data-id="${escapeHtml(i.id)}">Löschen</button>
        </div>
      </article>
    `
    )
    .join("");
}

function renderEntrySelects() {
  const selectedCustomerExists = state.customers.some((c) => c.id === entryCustomer.value);
  if (!selectedCustomerExists) {
    clearSelectedCustomer();
  }
  const selectedEmployeeExists = state.employees.some((e) => e.id === entryEmployee.value);
  if (!selectedEmployeeExists) {
    clearSelectedEmployee();
  }

  const previousItemId = entryItem.value;
  entryItem.innerHTML = state.items.length
    ? `<option value="">Bitte auswählen</option>${state.items
      .map((i) => `<option value="${escapeHtml(i.id)}">${escapeHtml(i.name)} (${formatCurrency(i.price)})</option>`)
      .join("")}`
    : "<option value=''>Zuerst Dienstleistungen/Produkte erfassen</option>";

  entryItem.disabled = !state.items.length;
  if (state.items.length) {
    if (state.items.some((item) => item.id === previousItemId)) {
      entryItem.value = previousItemId;
    } else {
      entryItem.value = "";
    }
  }
  updateEntryQuantityConstraints();

  renderCustomerSuggestions(entryCustomerSearch.value);
  renderEmployeeSuggestions(entryEmployeeSearch.value);
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
  return selectedItem?.unit === "Stück";
}

function resetEntryCustomerSelection() {
  entryCustomerSearch.value = "";
  clearSelectedCustomer();
  entryEmployeeSearch.value = "";
  clearSelectedEmployee();
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
          const itemName = item ? item.name : "Unbekannter Eintrag";
          const employeeName = employee ? `${employee.firstName || ""} ${employee.lastName || ""}`.trim() : "Unbekannter Mitarbeiter";
          const unitPrice = e.unitPrice != null ? e.unitPrice : resolveEntryUnitPrice(item);
          const total = unitPrice * e.quantity;
          return `
            <article class="card">
              <p class="entry-main">${escapeHtml(itemName)} | ${e.quantity} ${escapeHtml(item?.unit || "")}</p>
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
  return new Intl.NumberFormat("de-CH", {
    style: "currency",
    currency: "CHF"
  }).format(value || 0);
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

function setSelectedCustomer(customerId) {
  const customer = state.customers.find((c) => c.id === customerId);
  if (!customer) {
    clearSelectedCustomer();
    return;
  }
  entryCustomer.value = customer.id;
  entryCustomerDisplay.value = getCustomerDisplay(customer);
  refreshRequiredFieldStates();
}

function setSelectedEmployee(employeeId) {
  const employee = state.employees.find((e) => e.id === employeeId);
  if (!employee) {
    clearSelectedEmployee();
    return;
  }
  entryEmployee.value = employee.id;
  entryEmployeeDisplay.value = getEmployeeDisplay(employee);
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
  const exists = state.items.some((item) => normalizeSearchText(item.name) === normalizeSearchText("Reinigungsarbeiten"));
  if (exists) return;
  state.items.push({
    id: generateId(),
    name: "Reinigungsarbeiten",
    type: "Dienstleistung",
    unit: "Stunde/n",
    price: 35
  });
  saveState();
}

function normalizeMonths(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 12;
  return Math.min(12, Math.max(1, Math.round(n)));
}

function pruneOldEntriesForExport() {
  const removedEntries = listOldEntriesForExport();
  if (!removedEntries.length) return [];
  const idsToRemove = new Set(removedEntries.map((entry) => entry.id));
  state.entries = state.entries.filter((entry) => !idsToRemove.has(entry.id));
  if (removedEntries.length > 0) {
    saveState();
    renderEntries();
  }
  return removedEntries;
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
  renderExportCleanupResult(listOldEntriesForExport());
}

function renderExportCleanupResult(entries) {
  const removed = Array.isArray(entries) ? entries : [];
  if (!removed.length) {
    exportCleanupResult.hidden = true;
    exportCleanupList.innerHTML = "";
    return;
  }
  exportCleanupResult.hidden = false;
  exportCleanupList.innerHTML = removed
    .slice()
    .sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")))
    .map((entry) => {
      const customer = state.customers.find((c) => c.id === entry.customerId);
      const item = state.items.find((i) => i.id === entry.itemId);
      const customerName = customer ? getCustomerDisplay(customer) : "Unbekannter Kunde";
      const itemLabel = item ? item.name : "Unbekannte Leistung";
      const unit = item?.unit || "";
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
}


















































