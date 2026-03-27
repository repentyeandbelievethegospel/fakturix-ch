const STORAGE_KEY = "fakturix_ch_abrechnung_state_v2";
const LEGACY_STORAGE_KEY = "putzpro_desktop_state_v2";

const state = loadState();
let currentInvoiceHtml = "";

const tabs = [...document.querySelectorAll(".tab[data-tab]")];
const tabPanels = [...document.querySelectorAll(".tab-panel")];
const importSubTabs = [...document.querySelectorAll(".subtab[data-import-tab]")];
const importSubPanels = [...document.querySelectorAll(".import-subpanel")];

const importFiles = document.getElementById("importFiles");
const importBtn = document.getElementById("importBtn");
const camtImportFile = document.getElementById("camtImportFile");
const camtImportBtn = document.getElementById("camtImportBtn");
const resetDataBtn = document.getElementById("resetDataBtn");
const resetConfirmCheckbox = document.getElementById("resetConfirmCheckbox");
const importStatus = document.getElementById("importStatus");
const importSuccessDialog = document.getElementById("importSuccessDialog");
const importSuccessSummary = document.getElementById("importSuccessSummary");
const importSuccessDetails = document.getElementById("importSuccessDetails");
const importSuccessCloseBtn = document.getElementById("importSuccessCloseBtn");
const camtImportStatus = document.getElementById("camtImportStatus");
const camtImportResult = document.getElementById("camtImportResult");
const camtImportResultList = document.getElementById("camtImportResultList");
const camtImportDialog = document.getElementById("camtImportDialog");
const camtImportDialogList = document.getElementById("camtImportDialogList");
const camtImportDialogCloseBtn = document.getElementById("camtImportDialogCloseBtn");

const companyForm = document.getElementById("companyForm");
const companyLogoInput = document.getElementById("companyLogoInput");
const companyLogoPreview = document.getElementById("companyLogoPreview");
const companyLogoPreviewText = document.getElementById("companyLogoPreviewText");
const companyStatus = document.getElementById("companyStatus");
const headerLogo = document.getElementById("headerLogo");

const billingForm = document.getElementById("billingForm");
const billingCustomer = document.getElementById("billingCustomer");
const billingMonth = document.getElementById("billingMonth");
const billingPaymentSlip = document.getElementById("billingPaymentSlip");
const billingStatus = document.getElementById("billingStatus");
const billingSuccessDialog = document.getElementById("billingSuccessDialog");
const billingSuccessSummary = document.getElementById("billingSuccessSummary");
const billingSuccessTitle = document.getElementById("billingSuccessTitle");
const billingSuccessIcon = document.getElementById("billingSuccessIcon");
const billingSuccessDetails = document.getElementById("billingSuccessDetails");
const billingSuccessCloseBtn = document.getElementById("billingSuccessCloseBtn");
const invoicePreview = document.getElementById("invoicePreview");
const invoiceList = document.getElementById("invoiceList");
const invoiceFilterMonth = document.getElementById("invoiceFilterMonth");
const invoiceFilterCustomer = document.getElementById("invoiceFilterCustomer");
const invoiceFilterCustomerList = document.getElementById("invoiceFilterCustomerList");
const invoiceFilterStatus = document.getElementById("invoiceFilterStatus");
const invoiceFilterSentStatus = document.getElementById("invoiceFilterSentStatus");
const invoiceFilterClearBtn = document.getElementById("invoiceFilterClearBtn");
const invoiceDeleteAllBtn = document.getElementById("invoiceDeleteAllBtn");
const invoiceDeleteAllDialog = document.getElementById("invoiceDeleteAllDialog");
const invoiceDeleteAllList = document.getElementById("invoiceDeleteAllList");
const invoiceDeleteAllConfirmCheckbox = document.getElementById("invoiceDeleteAllConfirmCheckbox");
const invoiceDeleteAllCancelBtn = document.getElementById("invoiceDeleteAllCancelBtn");
const invoiceDeleteAllContinueBtn = document.getElementById("invoiceDeleteAllContinueBtn");
const hoursMonth = document.getElementById("hoursMonth");
const hoursSummary = document.getElementById("hoursSummary");
const hoursByCustomer = document.getElementById("hoursByCustomer");
const hoursByEmployee = document.getElementById("hoursByEmployee");
const hoursByDays = document.getElementById("hoursByDays");
const hoursByProductDays = document.getElementById("hoursByProductDays");
const hoursByProductsEmployee = document.getElementById("hoursByProductsEmployee");
const hoursByProductsCustomer = document.getElementById("hoursByProductsCustomer");
const hoursByProductsProfit = document.getElementById("hoursByProductsProfit");
const hoursByProfit = document.getElementById("hoursByProfit");
const hoursByVat = document.getElementById("hoursByVat");
const hoursSubTabs = [...document.querySelectorAll(".subtab[data-hours-tab]")];
const hoursCustomerPanel = document.getElementById("hours-customer-panel");
const hoursEmployeePanel = document.getElementById("hours-employee-panel");
const hoursDaysPanel = document.getElementById("hours-days-panel");
const hoursProductDaysPanel = document.getElementById("hours-product-days-panel");
const hoursProductsEmployeePanel = document.getElementById("hours-products-employee-panel");
const hoursProductsCustomerPanel = document.getElementById("hours-products-customer-panel");
const hoursProductsProfitPanel = document.getElementById("hours-products-profit-panel");
const hoursProfitPanel = document.getElementById("hours-profit-panel");
const hoursVatPanel = document.getElementById("hours-vat-panel");
const backupExportBtn = document.getElementById("backupExportBtn");
const backupImportFile = document.getElementById("backupImportFile");
const backupImportBtn = document.getElementById("backupImportBtn");
const invoiceCleanupMonths = document.getElementById("invoiceCleanupMonths");
const invoiceCleanupBtn = document.getElementById("invoiceCleanupBtn");
const invoiceCleanupConfirmCheckbox = document.getElementById("invoiceCleanupConfirmCheckbox");
const invoiceCleanupConfirmValue = document.getElementById("invoiceCleanupConfirmValue");
const backupStatus = document.getElementById("backupStatus");
const cleanupResult = document.getElementById("cleanupResult");
const cleanupResultList = document.getElementById("cleanupResultList");
const backupNowBtn = document.getElementById("backupNowBtn");
const backupNowStatus = document.getElementById("backupNowStatus");
const maintenanceActionDialog = document.getElementById("maintenanceActionDialog");
const maintenanceActionIcon = document.getElementById("maintenanceActionIcon");
const maintenanceActionTitle = document.getElementById("maintenanceActionTitle");
const maintenanceActionSummary = document.getElementById("maintenanceActionSummary");
const maintenanceActionDetails = document.getElementById("maintenanceActionDetails");
const maintenanceActionCloseBtn = document.getElementById("maintenanceActionCloseBtn");
let invoiceFilterCustomerNameToId = new Map();
let activeHoursTab = "mitarbeiter";
let hoursCustomerFilterText = "";
let hoursEmployeeFilterText = "";
let hoursProductsEmployeeFilterText = "";
let hoursProductsCustomerFilterText = "";
let camtPendingOverpayActions = new Map();
let camtLastResultFileName = "";
let camtLastResultRows = [];
let companyLogoPreviewObjectUrl = "";
let pendingDeleteAllInvoiceIds = [];

wireImport();
wireCamtImport();
wireImportSubTabs();
wireCompany();
wireCompanyValidation();
wireBilling();
wireInvoiceList();
wireHoursReport();
wireTabs();
wireBackup();
wireResetConfirmation();
wireImportBackupValidation();
renderAll();


function wireTabs() {
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      activateTab(tab.dataset.tab || "");
    });
  });
}

function wireImportSubTabs() {
  if (!importSubTabs.length) return;
  importSubTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      activateImportSubTab(tab.dataset.importTab || "import-entries-panel");
    });
  });
  activateImportSubTab("import-entries-panel");
}

function activateImportSubTab(targetPanelId) {
  importSubTabs.forEach((tab) => {
    const active = tab.dataset.importTab === targetPanelId;
    tab.classList.toggle("active", active);
    tab.setAttribute("aria-selected", String(active));
  });
  importSubPanels.forEach((panel) => {
    const active = panel.id === targetPanelId;
    panel.classList.toggle("active", active);
    panel.hidden = !active;
  });
}

function activateTab(target) {
  tabs.forEach((t) => {
    const active = t.dataset.tab === target;
    t.classList.toggle("active", active);
    t.setAttribute("aria-selected", String(active));
  });
  tabPanels.forEach((p) => {
    const active = p.id === target;
    p.classList.toggle("active", active);
    p.hidden = !active;
  });
}
function baseState() {
  return {
    company: {
      logoDataUrl: "",
      name: "",
      street: "",
      zip: "",
      city: "",
      website: "",
      email: "",
      mobile: "",
      iban: "",
      currency: "CHF",
      vatNo: "",
      vatRate: 8.1,
      paymentTermDays: 30,
      invoiceText: "Vielen Dank für Ihren Auftrag.",
      appendixText: ""
    },
    customers: [],
    employees: [],
    items: [],
    entries: [],
    invoices: []
  };
}

function wireBackup() {
  const syncBackupImportButton = () => {
    backupImportBtn.disabled = !backupImportFile.files?.[0];
  };
  const syncCleanupButton = () => {
    invoiceCleanupBtn.disabled = !invoiceCleanupConfirmCheckbox?.checked;
  };
  const syncCleanupConfirmText = () => {
    if (!invoiceCleanupConfirmValue) return;
    const months = Math.min(24, Math.max(1, Math.round(toNumber(invoiceCleanupMonths?.value, 12))));
    invoiceCleanupConfirmValue.textContent = `${months} ${months === 1 ? "Monat" : "Monate"}`;
  };
  backupImportFile.addEventListener("change", syncBackupImportButton);
  invoiceCleanupConfirmCheckbox?.addEventListener("change", syncCleanupButton);
  invoiceCleanupConfirmCheckbox?.addEventListener("input", syncCleanupButton);
  invoiceCleanupConfirmCheckbox?.addEventListener("click", syncCleanupButton);
  invoiceCleanupMonths?.addEventListener("change", () => {
    syncCleanupConfirmText();
    if (invoiceCleanupConfirmCheckbox) invoiceCleanupConfirmCheckbox.checked = false;
    syncCleanupButton();
    refreshImportBackupValidationStates();
  });
  syncBackupImportButton();
  syncCleanupConfirmText();
  syncCleanupButton();

  if (maintenanceActionCloseBtn) {
    maintenanceActionCloseBtn.addEventListener("click", closeMaintenanceActionDialog);
  }
  if (maintenanceActionDialog) {
    maintenanceActionDialog.addEventListener("click", (event) => {
      if (event.target === maintenanceActionDialog) closeMaintenanceActionDialog();
    });
  }

  backupExportBtn.addEventListener("click", async () => {
    const result = await exportDesktopBackup();
    if (result.canceled) {
      backupStatus.textContent = "Backup-Speichern wurde abgebrochen.";
      showMaintenanceActionDialog({ title: "Backup", summary: "Backup-Speichern wurde abgebrochen.", kind: "warning" });
      return;
    }
    if (result.savedWithDialog) {
      backupStatus.textContent = `Backup gespeichert: ${result.fileName}`;
      showMaintenanceActionDialog({ title: "Backup erstellt", summary: `Backup gespeichert: ${result.fileName}`, details: [result.fileName], kind: "success" });
      return;
    }
    backupStatus.textContent = `Backup exportiert: ${result.fileName}`;
    showMaintenanceActionDialog({ title: "Backup erstellt", summary: `Backup exportiert: ${result.fileName}`, details: [result.fileName], kind: "success" });
  });

  backupNowBtn.addEventListener("click", async () => {
    const result = await exportDesktopBackup();
    if (result.canceled) {
      backupNowStatus.textContent = "Backup-Speichern wurde abgebrochen.";
      showMaintenanceActionDialog({ title: "Backup", summary: "Backup-Speichern wurde abgebrochen.", kind: "warning" });
      return;
    }
    backupNowStatus.textContent = result.savedWithDialog
      ? `Backup gespeichert: ${result.fileName}`
      : `Backup exportiert: ${result.fileName}`;
    showMaintenanceActionDialog({
      title: "Backup erstellt",
      summary: backupNowStatus.textContent,
      details: [result.fileName],
      kind: "success"
    });
  });

  backupImportBtn.addEventListener("click", async () => {
    try {
      const file = backupImportFile.files?.[0];
      if (!file) {
        backupStatus.textContent = "Bitte zuerst eine Backup-Datei auswählen.";
        showMaintenanceActionDialog({ title: "Restore", summary: "Bitte zuerst eine Backup-Datei auswählen.", kind: "warning" });
        return;
      }
      if (!confirm("Backup importieren und alle aktuellen Desktop-Daten ersetzen?")) return;
      if (!confirm("Sind Sie wirklich sicher?")) return;

      const parsed = JSON.parse(await file.text());
      const normalized = normalizeBackup(parsed);

      state.company = normalized.company;
      state.customers = normalized.customers;
      state.employees = normalized.employees;
      state.items = normalized.items;
      state.entries = normalized.entries;
        state.invoices = normalized.invoices;

        saveState();
        renderAll();
        setPreviewHtml("");
        backupImportFile.value = "";
        syncBackupImportButton();
        refreshImportBackupValidationStates();
        backupStatus.textContent = "Backup erfolgreich importiert.";
        showMaintenanceActionDialog({ title: "Restore", summary: "Backup erfolgreich importiert.", details: [file.name], kind: "success" });
      } catch {
        backupStatus.textContent = "Backup-Import fehlgeschlagen. Bitte Datei prüfen.";
        showMaintenanceActionDialog({ title: "Restore", summary: "Backup-Import fehlgeschlagen. Bitte Datei prüfen.", kind: "error" });
      }
    });

      window.__handleInvoiceCleanupClick = () => {
    if (!invoiceCleanupConfirmCheckbox?.checked) {
      backupStatus.textContent = "Bitte zuerst die Bestätigung aktivieren.";
      showMaintenanceActionDialog({ title: "Datenbereinigung", summary: "Bitte zuerst die Bestätigung aktivieren.", kind: "warning" });
      return;
    }

    const months = Math.min(24, Math.max(1, Math.round(toNumber(invoiceCleanupMonths.value, 6))));
    const invoiceCandidates = getInvoicesOlderThanMonths(months);
    const entryCandidates = getEntriesOlderThanMonths(months);

    if (!invoiceCandidates.length && !entryCandidates.length) {
      backupStatus.textContent = `Keine Erfassungen/Rechnungen älter als ${months} Monate gefunden.`;
      showMaintenanceActionDialog({ title: "Datenbereinigung", summary: `Keine Erfassungen/Rechnungen älter als ${months} Monate gefunden.`, kind: "warning" });
      renderCleanupResult([], []);
      return;
    }

    const preview = invoiceCandidates
      .slice(0, 10)
      .map((invoice) => formatInvoiceCleanupLine(invoice))
      .join("\n");
    const moreText = invoiceCandidates.length > 10 ? `\n... und ${invoiceCandidates.length - 10} weitere` : "";
    const previewBlock = invoiceCandidates.length
      ? `\n\nBetroffene Rechnungen:\n${preview}${moreText}`
      : "";

    if (
      !confirm(
        `Es werden gelöscht:\n- Rechnungen: ${invoiceCandidates.length}\n- Erfassungen: ${entryCandidates.length}${previewBlock}`
      )
    ) return;
    if (!confirm("Sind Sie wirklich sicher?")) return;

    const cleanupResultData = pruneDataOlderThanMonths(months);
    saveState();
    renderAll();
    setPreviewHtml("");

    backupStatus.textContent = `Gelöscht: ${cleanupResultData.removedInvoices.length} Rechnungen, ${cleanupResultData.removedEntries.length} Erfassungen.`;
    showMaintenanceActionDialog({
      title: "Datenbereinigung",
      summary: `Gelöscht: ${cleanupResultData.removedInvoices.length} Rechnungen, ${cleanupResultData.removedEntries.length} Erfassungen.`,
      details: [
        `${cleanupResultData.removedInvoices.length} Rechnung(en) entfernt`,
        `${cleanupResultData.removedEntries.length} Erfassung(en) entfernt`
      ],
      kind: "success"
    });

    renderCleanupResult(cleanupResultData.removedInvoices, cleanupResultData.removedEntries);
    if (invoiceCleanupConfirmCheckbox) invoiceCleanupConfirmCheckbox.checked = false;
    syncCleanupButton();
  };

  invoiceCleanupBtn.onclick = window.__handleInvoiceCleanupClick;
}

async function exportDesktopBackup() {
  const payload = {
    app: "Fakturix CH Abrechnung",
    version: 1,
    exportedAt: new Date().toISOString(),
    data: state
  };
  const dateTag = new Date().toISOString().slice(0, 10);
  const fileName = `fakturix-ch-abrechnung-backup-${dateTag}.json`;
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const saveResult = await saveBlobWithDialog(blob, fileName);
  if (saveResult === "unsupported" || saveResult === "error") {
    downloadBlob(blob, fileName);
  }
  return { fileName, savedWithDialog: saveResult === "saved", canceled: saveResult === "canceled" };
}

async function saveBlobWithDialog(blob, fileName) {
  if (typeof window.showSaveFilePicker !== "function") return "unsupported";
  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: fileName,
      types: [
        {
          description: "JSON-Datei",
          accept: { "application/json": [".json"] }
        }
      ]
    });
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
    return "saved";
  } catch (error) {
    if (error && error.name === "AbortError") return "canceled";
    console.error(error);
    return "error";
  }
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY) || localStorage.getItem(LEGACY_STORAGE_KEY);
    if (!raw) return baseState();
    const parsed = JSON.parse(raw);
    return {
      company: parsed.company && typeof parsed.company === "object" ? { ...baseState().company, ...parsed.company } : baseState().company,
      customers: Array.isArray(parsed.customers) ? parsed.customers : [],
      employees: Array.isArray(parsed.employees) ? parsed.employees : [],
      items: (Array.isArray(parsed.items) ? parsed.items : []).map((item) => normalizeItemRecord(item)),
      entries: (Array.isArray(parsed.entries) ? parsed.entries : []).map((entry) => normalizeEntryRecord(entry)),
      invoices: Array.isArray(parsed.invoices) ? parsed.invoices.map((inv) => ({ ...inv, paymentSlipType: normalizePaymentSlipType(inv?.paymentSlipType), sent: Boolean(inv?.sent), sentAt: String(inv?.sentAt || "") })) : []
    };
  } catch {
    return baseState();
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function wireImport() {
  const syncImportButton = () => {
    importBtn.disabled = !(importFiles.files && importFiles.files.length);
  };
  importFiles.addEventListener("change", syncImportButton);
  syncImportButton();

  if (importSuccessCloseBtn) {
    importSuccessCloseBtn.addEventListener("click", closeImportSuccessDialog);
  }
  if (importSuccessDialog) {
    importSuccessDialog.addEventListener("click", (event) => {
      if (event.target === importSuccessDialog) closeImportSuccessDialog();
    });
  }
  importBtn.addEventListener("click", async () => {
    try {
      const files = [...(importFiles.files || [])];
      if (!files.length) {
        importStatus.textContent = "Bitte mindestens eine Export-Datei auswählen.";
        return;
      }

      const imports = [];
      const parseErrors = [];
      for (const file of files) {
        try {
          const rawText = await file.text();
          const cleanedText = String(rawText ?? "").replace(/^\uFEFF/, "").trim();
          if (!cleanedText) throw new Error("Datei ist leer.");
          const parsed = JSON.parse(cleanedText);
          const normalized = parsed && typeof parsed === "object"
            ? (parsed.data && typeof parsed.data === "object" ? parsed.data : parsed)
            : null;
          if (!normalized) throw new Error("Dateiinhalt ist kein gueltiges JSON-Objekt.");
          imports.push(normalized);
        } catch (error) {
          const reason = error?.message ? String(error.message) : "ungueltiges JSON";
          parseErrors.push(`${file.name}: ${reason}`);
        }
      }

      if (parseErrors.length) {
        importStatus.textContent = `Import fehlgeschlagen: ${parseErrors[0]}`;
        return;
      }

      const importedSummary = summarizeImportedEntities(imports);
      mergeImports(imports);
      saveState();
      renderAll();
      importStatus.textContent = `Import erfolgreich. Neu importiert: ${importedSummary.customers} Kunden, ${importedSummary.employees} Mitarbeiter, ${importedSummary.items} Leistungen/Produkte, ${importedSummary.entries} Erfassungen. Gesamt: ${state.customers.length} Kunden, ${state.employees.length} Mitarbeiter, ${state.items.length} Leistungen/Produkte, ${state.entries.length} Erfassungen.`;
      showImportSuccessDialog(importedSummary, files);
      importFiles.value = "";
      syncImportButton();

      refreshImportBackupValidationStates();
          } catch (error) {
      const reason = error?.message ? String(error.message) : "Bitte JSON-Export pruefen.";
      importStatus.textContent = `Import fehlgeschlagen: ${reason}`;
    }
  });

  resetDataBtn.addEventListener("click", () => {
    if (!confirm("Alle Desktop-Daten löschen?")) return;
    if (!confirm("Sind Sie wirklich sicher?")) return;
    Object.assign(state, baseState());
    saveState();
    renderAll();
    importStatus.textContent = "Alle Daten wurden gelöscht.";
    showMaintenanceActionDialog({ title: "Applikation zurücksetzen", summary: "Alle Daten wurden gelöscht.", kind: "success" });
    billingStatus.textContent = "";
    setPreviewHtml("");
    resetConfirmCheckbox.checked = false;
    resetDataBtn.disabled = true;
    refreshImportBackupValidationStates();
  });
}

function closeImportSuccessDialog() {
  if (!importSuccessDialog) return;
  importSuccessDialog.hidden = true;
}

function showImportSuccessDialog(importedSummary, files) {
  if (!importSuccessDialog || !importSuccessSummary || !importSuccessDetails) return;
  const fileCount = Array.isArray(files) ? files.length : 0;
  importSuccessSummary.textContent = `${fileCount} Datei(en) erfolgreich importiert.`;
const details = [
    `Neu importiert: ${importedSummary.customers} Kunden`,
    `Neu importiert: ${importedSummary.employees} Mitarbeiter`,
    `Neu importiert: ${importedSummary.items} Leistungen/Produkte`,
    `Neu importiert: ${importedSummary.entries} Erfassungen`,
    `Gesamt: ${state.customers.length} Kunden`,
    `Gesamt: ${state.employees.length} Mitarbeiter`,
    `Gesamt: ${state.items.length} Leistungen/Produkte`,
    `Gesamt: ${state.entries.length} Erfassungen`
  ];
  importSuccessDetails.innerHTML = details
    .map((line) => `<article class="cleanup-result-row">${escapeHtml(line)}</article>`)
    .join("");
  importSuccessDialog.hidden = false;
}
function closeBillingSuccessDialog() {
  if (!billingSuccessDialog) return;
  billingSuccessDialog.hidden = true;
}

function showBillingSuccessDialog(month, createdCount, alreadyExisting, missingData, customerId) {
  if (!billingSuccessDialog || !billingSuccessSummary || !billingSuccessDetails) return;
  const mode = customerId === "all" ? "Alle Kunden" : "Einzelkunde";
  billingSuccessSummary.textContent = `Monat ${formatMonthCH(month)} | ${mode}`;

  const createdAny = createdCount > 0;
  const noDataOnly = !createdAny && missingData > 0 && alreadyExisting === 0;
  const onlyExistingOrMissing = !createdAny && (alreadyExisting > 0 || missingData > 0);
  if (billingSuccessTitle) {
    billingSuccessTitle.textContent = "Ergebnis Rechnungslauf";
  }
  if (billingSuccessIcon) {
    if (noDataOnly) {
      billingSuccessIcon.textContent = "✖";
    } else if (onlyExistingOrMissing) {
      billingSuccessIcon.textContent = "!";
    } else {
      billingSuccessIcon.textContent = "✔";
    }
    billingSuccessIcon.classList.toggle("warning", onlyExistingOrMissing && !noDataOnly);
    billingSuccessIcon.classList.toggle("error", noDataOnly);
  }
  if (noDataOnly) {
    billingSuccessSummary.textContent += " | Keine erfassten Leistungen für den gewählten Monat vorhanden.";
  }
const details = [
    `Erstellt: ${createdCount} Rechnung(en)`,
    `Bereits vorhanden: ${alreadyExisting}`
  ];
  billingSuccessDetails.innerHTML = details
    .map((line) => `<article class="cleanup-result-row">${escapeHtml(line)}</article>`)
    .join("");
  billingSuccessDialog.hidden = false;
}
function closeMaintenanceActionDialog() {
  if (!maintenanceActionDialog) return;
  maintenanceActionDialog.hidden = true;
}

function showMaintenanceActionDialog({ title, summary = "", details = [], kind = "success" }) {
  if (!maintenanceActionDialog || !maintenanceActionIcon || !maintenanceActionTitle || !maintenanceActionSummary || !maintenanceActionDetails) return;
  const type = ["success", "warning", "error"].includes(kind) ? kind : "success";
  maintenanceActionTitle.textContent = String(title || "Wartung");
  maintenanceActionSummary.textContent = String(summary || "");
  maintenanceActionIcon.classList.toggle("warning", type === "warning");
  maintenanceActionIcon.classList.toggle("error", type === "error");
  maintenanceActionIcon.textContent = type === "error" ? "✖" : type === "warning" ? "!" : "✔";
  const rows = Array.isArray(details) ? details.filter(Boolean) : [];
  maintenanceActionDetails.innerHTML = rows.length
    ? rows.map((line) => `<article class="cleanup-result-row">${escapeHtml(String(line))}</article>`).join("")
    : "";
  maintenanceActionDialog.hidden = false;
}
function wireCamtImport() {
  if (!camtImportBtn || !camtImportFile || !camtImportStatus) return;
  const syncCamtImportButton = () => {
    camtImportBtn.disabled = !(camtImportFile.files && camtImportFile.files.length);
    refreshImportBackupValidationStates();
  };
  camtImportFile.addEventListener("change", syncCamtImportButton);
  syncCamtImportButton();
  if (camtImportDialogCloseBtn) {
    camtImportDialogCloseBtn.addEventListener("click", closeCamtImportDialog);
  }
  if (camtImportDialog) {
    camtImportDialog.addEventListener("click", (event) => {
      if (event.target === camtImportDialog) closeCamtImportDialog();
    });
  }
  if (camtImportResultList) {
    camtImportResultList.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-action='confirm-overpay'][data-id]");
      if (!button) return;
      const actionId = String(button.dataset.id || "");
      const action = camtPendingOverpayActions.get(actionId);
      if (!action) return;
      const invoice = state.invoices.find((i) => i.id === action.invoiceId);
      if (!invoice) {
        camtImportStatus.textContent = "Rechnung für diese Aktion wurde nicht gefunden.";
        camtPendingOverpayActions.delete(actionId);
      } else if (invoice.paid) {
        camtImportStatus.textContent = `Rechnung ${invoice.invoiceNo || "(ohne Nr.)"} ist bereits bezahlt.`;
        camtPendingOverpayActions.delete(actionId);
      } else {
        markInvoicePaidFromBooking(invoice, action.booking, action.importMeta);
        saveState();
        renderInvoiceList();
        if (currentInvoiceHtml) setPreviewHtml(currentInvoiceHtml);
        camtImportStatus.textContent = `Rechnung ${invoice.invoiceNo || "(ohne Nr.)"} wurde trotz Mehrbetrag als bezahlt markiert.`;
        camtPendingOverpayActions.delete(actionId);
      }
      camtLastResultRows = camtLastResultRows.map((row) =>
        row.actionId === actionId
          ? {
              ...row,
              matched: true,
              resultType: "success",
              pendingAction: false,
              matchType: "Referenz + Mehrbetrag (bestätigt)",
              reason: "",
              invoiceNo: row.invoiceNo || invoice?.invoiceNo || "(ohne Nr.)"
            }
          : row
      );
      renderCamtImportResult(camtLastResultFileName, camtLastResultRows);
    });
  }
  camtImportBtn.addEventListener("click", async () => {
    const file = camtImportFile.files?.[0];
    if (!file) {
      camtImportStatus.textContent = "Bitte zuerst einen Kontoauszug auswählen.";
      return;
    }
    try {
      const text = await file.text();
      const bookings = parseCamtBookings(text);
      if (!bookings.length) {
        camtImportStatus.textContent = `Keine Buchungen in ${file.name} gefunden oder CAMT-Format nicht erkannt.`;
        camtPendingOverpayActions = new Map();
        camtLastResultFileName = "";
        camtLastResultRows = [];
        renderCamtImportResult("", []);
        return;
      }

      const openInvoices = state.invoices.filter((i) => !i.paid);
      const paidInvoices = state.invoices.filter((i) => i.paid);
      if (!openInvoices.length) {
        let questionOnly = 0;
        let errorOnly = 0;
        const noOpenResults = bookings.map((booking) => {
          const paidMatch = findMatchingInvoiceForBooking(booking, paidInvoices, new Set()) ||
            (findReferenceOnlyInvoice(booking, paidInvoices, new Set()) ? { invoice: findReferenceOnlyInvoice(booking, paidInvoices, new Set()), matchType: "Referenz" } : null) ||
            (findAmountOnlyInvoice(booking, paidInvoices, new Set()) ? { invoice: findAmountOnlyInvoice(booking, paidInvoices, new Set()), matchType: "Betrag" } : null);
          if (paidMatch?.invoice) {
            questionOnly += 1;
            return {
              booking,
              matched: false,
              resultType: "question",
              reason: `Bereits bezahlt (${paidMatch.invoice.invoiceNo || "(ohne Nr.)"})`
            };
          }
          errorOnly += 1;
          return { booking, matched: false, resultType: "error", reason: "Keine offenen Rechnungen" };
        });
        camtImportStatus.textContent = `Kontoauszug importiert: ${bookings.length} Buchung(en), 0 automatisch bezahlt, ${questionOnly} zu prüfen, ${errorOnly} ohne Treffer.`;
        camtPendingOverpayActions = new Map();
        camtLastResultFileName = file.name;
        camtLastResultRows = noOpenResults;
        renderCamtImportResult(file.name, noOpenResults);
        showCamtImportDialog();
        return;
      }

      const matchedInvoiceIds = new Set();
      const matchedInvoiceNos = [];
      let successCount = 0;
      let questionCount = 0;
      let errorCount = 0;
      const bookingResults = [];
      camtPendingOverpayActions = new Map();

      const importMeta = { fileName: file.name, importedAt: new Date().toISOString() };
      bookings.forEach((booking) => {
        const strictMatch = findMatchingInvoiceForBooking(booking, openInvoices, matchedInvoiceIds);
        if (strictMatch) {
          const { invoice, matchType } = strictMatch;
          markInvoicePaidFromBooking(invoice, booking, importMeta);
          matchedInvoiceIds.add(invoice.id);
          matchedInvoiceNos.push(invoice.invoiceNo || "(ohne Nr.)");
          successCount += 1;
          bookingResults.push({ booking, matched: true, resultType: "success", invoiceNo: invoice.invoiceNo || "(ohne Nr.)", matchType });
          return;
        }

        const refOpenInvoice = findReferenceOnlyInvoice(booking, openInvoices, matchedInvoiceIds);
        const amountOpenInvoice = findAmountOnlyInvoice(booking, openInvoices, matchedInvoiceIds);

        if (refOpenInvoice) {
          const bookingAmount = roundAmount(booking.amount);
          const invoiceAmount = roundAmount(refOpenInvoice.grandTotal);
          const diff = roundAmount(bookingAmount - invoiceAmount);
          const actionId = generateId();
          camtPendingOverpayActions.set(actionId, { invoiceId: refOpenInvoice.id, booking, importMeta });
          questionCount += 1;
          if (bookingAmount > invoiceAmount) {
            bookingResults.push({
              booking,
              matched: false,
              resultType: "question",
              invoiceNo: refOpenInvoice.invoiceNo || "(ohne Nr.)",
              reason: `Referenz passt, Betrag höher um ${formatCurrency(diff)}`,
              pendingAction: true,
              actionId
            });
            return;
          }
          bookingResults.push({
            booking,
            matched: false,
            resultType: "question",
            invoiceNo: refOpenInvoice.invoiceNo || "(ohne Nr.)",
            reason: `Referenz passt, Betrag ${diff < 0 ? `tiefer um ${formatCurrency(Math.abs(diff))}` : "abweichend"}`,
            pendingAction: true,
            actionId
          });
          return;
        }

        if (amountOpenInvoice) {
          const actionId = generateId();
          camtPendingOverpayActions.set(actionId, { invoiceId: amountOpenInvoice.id, booking, importMeta });
          questionCount += 1;
          bookingResults.push({
            booking,
            matched: false,
            resultType: "question",
            invoiceNo: amountOpenInvoice.invoiceNo || "(ohne Nr.)",
            reason: "Betrag passt, Referenz/Mitteilung passt nicht",
            pendingAction: true,
            actionId
          });
          return;
        }

        const paidStrict = findMatchingInvoiceForBooking(booking, paidInvoices, new Set());
        const paidRef = findReferenceOnlyInvoice(booking, paidInvoices, new Set());
        const paidAmount = findAmountOnlyInvoice(booking, paidInvoices, new Set());
        const paidInvoice = paidStrict?.invoice || paidRef || paidAmount;
        if (paidInvoice) {
          questionCount += 1;
          bookingResults.push({
            booking,
            matched: false,
            resultType: "question",
            reason: `Bereits bezahlt (${paidInvoice.invoiceNo || "(ohne Nr.)"})`
          });
          return;
        }

        errorCount += 1;
        bookingResults.push({ booking, matched: false, resultType: "error", reason: "Kein Treffer" });
      });

      if (matchedInvoiceIds.size > 0) {
        saveState();
        renderInvoiceList();
        if (currentInvoiceHtml) setPreviewHtml(currentInvoiceHtml);
      }

      const preview = matchedInvoiceNos.slice(0, 8).join(", ");
      const more = matchedInvoiceNos.length > 8 ? ` (+${matchedInvoiceNos.length - 8} weitere)` : "";
      camtImportStatus.textContent =
        `Kontoauszug importiert: ${bookings.length} Buchung(en), ${successCount} automatisch bezahlt, ${questionCount} zu prüfen, ${errorCount} ohne Treffer.` +
        (preview ? ` Treffer: ${preview}${more}` : "");
      camtLastResultFileName = file.name;
      camtLastResultRows = bookingResults;
      renderCamtImportResult(file.name, bookingResults);
      showCamtImportDialog();
    } catch {
      camtImportStatus.textContent = "Kontoauszug-Import fehlgeschlagen. Bitte Datei prüfen.";
      camtPendingOverpayActions = new Map();
      camtLastResultFileName = "";
      camtLastResultRows = [];
      renderCamtImportResult("", []);
    }
  });
}

function renderCamtImportResult(fileName, results) {
  if (!camtImportResult || !camtImportResultList) return;
  const safe = Array.isArray(results) ? results : [];
  if (!safe.length) {
    camtImportResult.hidden = true;
    camtImportResultList.innerHTML = "";
    return;
  }
  camtImportResult.hidden = false;
  const renderRow = (r) => {
      const booking = r.booking || {};
      const amount = formatCurrency(toNumber(booking.amount, 0));
      const date = booking.date ? formatDateCH(booking.date.slice(0, 10)) : "-";
      const reference = booking.reference ? escapeHtml(String(booking.reference)) : "(ohne Referenz)";
      const alreadyPaid = /^Bereits bezahlt/i.test(String(r.reason || ""));
      if (r.resultType === "success") {
        return `
          <article class="cleanup-result-row">
            <span class="import-icon icon-success" aria-hidden="true">✔</span>
            <strong>Treffer: ${escapeHtml(r.invoiceNo || "(ohne Nr.)")}</strong><br>
            Datum: ${escapeHtml(date)} | Betrag: ${escapeHtml(amount)}<br>
            Match: ${escapeHtml(r.matchType || "Unbekannt")}<br>
            Referenz: ${reference}<br>
            <em><span class="import-note-icon" aria-hidden="true">ℹ</span> Im Rechnungsjournal als bezahlt markiert.</em>
          </article>
        `;
      }
      if (r.pendingAction && r.actionId) {
        return `
          <article class="cleanup-result-row">
            <span class="import-icon icon-question" aria-hidden="true">?</span>
            <strong>Prüfung: ${escapeHtml(r.invoiceNo || "(ohne Nr.)")}</strong><br>
            Datum: ${escapeHtml(date)} | Betrag: ${escapeHtml(amount)}<br>
            Grund: ${escapeHtml(r.reason || "Mehrbetrag")}<br>
            Referenz: ${reference}<br>
            <div class="actions" style="margin-top:0.4rem;">
              <button type="button" class="primary" data-action="confirm-overpay" data-id="${escapeHtml(r.actionId)}">Trotzdem als bezahlt markieren</button>
            </div>
          </article>
        `;
      }
      if (r.resultType === "question") {
        return `
          <article class="cleanup-result-row">
            <span class="import-icon ${alreadyPaid ? "icon-success" : "icon-question"}" aria-hidden="true">${alreadyPaid ? "✔" : "?"}</span>
            <strong>${alreadyPaid ? "Bereits bezahlt" : "Prüfen"}</strong><br>
            Datum: ${escapeHtml(date)} | Betrag: ${escapeHtml(amount)}<br>
            Grund: ${escapeHtml(r.reason || "Teiltreffer")}<br>
            Referenz: ${reference}
          </article>
        `;
      }
      return `
        <article class="cleanup-result-row">
          <span class="import-icon icon-error" aria-hidden="true">✖</span>
          <strong>Kein Treffer</strong><br>
          Datum: ${escapeHtml(date)} | Betrag: ${escapeHtml(amount)}<br>
          Grund: ${escapeHtml(r.reason || "Kein Treffer")}<br>
          Referenz: ${reference}
        </article>
      `;
    };
  const reviewRows = safe
    .filter((r) => r.resultType === "question" && !/^Bereits bezahlt/i.test(String(r.reason || "")))
    .map(renderRow)
    .join("");
  const successRows = safe
    .filter((r) => r.resultType === "success")
    .map(renderRow)
    .join("");
  const errorRows = safe
    .filter((r) => r.resultType !== "success" && r.resultType !== "question" && !/^Bereits bezahlt/i.test(String(r.reason || "")))
    .map(renderRow)
    .join("");
  const alreadyPaidRows = safe
    .filter((r) => /^Bereits bezahlt/i.test(String(r.reason || "")))
    .map(renderRow)
    .join("");
  const paidSection = alreadyPaidRows
    ? `
      <details class="cleanup-result-row" style="margin-top:0.6rem;">
        <summary><strong>Bereits bezahlt (${safe.filter((r) => /^Bereits bezahlt/i.test(String(r.reason || ""))).length}) anzeigen</strong></summary>
        <div class="cleanup-result-list" style="margin-top:0.6rem;">
          ${alreadyPaidRows}
        </div>
      </details>
    `
    : "";
  camtImportResultList.innerHTML = `
    <small>Datei: ${escapeHtml(fileName || "-")} | Buchungen: ${safe.length}</small>
    ${reviewRows}
    ${successRows}
    ${errorRows}
    ${paidSection}
  `;
}

function closeCamtImportDialog() {
  if (!camtImportDialog) return;
  camtImportDialog.hidden = true;
}

function showCamtImportDialog() {
  if (!camtImportDialog || !camtImportDialogList || !camtImportResultList || !camtImportResult) return;
  if (camtImportResult.hidden) return;
  camtImportDialogList.innerHTML = camtImportResultList.innerHTML;
  camtImportDialog.hidden = false;
}
function parseCamtBookings(xmlText) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(String(xmlText || ""), "application/xml");
  if (doc.querySelector("parsererror")) return [];

  const entryNodes = [...doc.getElementsByTagNameNS("*", "Ntry")];
  const bookings = [];

  entryNodes.forEach((entry) => {
    const entryAmount = toNumber(getFirstDescendant(entry, ["Amt"])?.textContent, 0);
    const cdtDbt = (getFirstDescendant(entry, ["CdtDbtInd"])?.textContent || "").trim().toUpperCase();
    const sign = cdtDbt === "DBIT" ? -1 : 1;
    const bookingDate =
      (getFirstDescendant(entry, ["BookgDt", "Dt"])?.textContent || "").trim() ||
      (getFirstDescendant(entry, ["ValDt", "Dt"])?.textContent || "").trim();

    const txNodes = [...entry.getElementsByTagNameNS("*", "TxDtls")];
    if (!txNodes.length) {
      const ref = collectBookingReference(entry);
      const amount = sign * entryAmount;
      if (Number.isFinite(amount) && amount > 0) {
        bookings.push({ amount, date: bookingDate, reference: ref });
      }
      return;
    }

    txNodes.forEach((tx) => {
      const txAmountNode = getFirstDescendant(tx, ["AmtDtls", "TxAmt", "Amt"]) || getFirstDescendant(tx, ["Amt"]);
      const txAmount = toNumber(txAmountNode?.textContent, entryAmount);
      const amount = sign * txAmount;
      if (!(Number.isFinite(amount) && amount > 0)) return;
      const ref = collectBookingReference(tx) || collectBookingReference(entry);
      const txDate =
        (getFirstDescendant(tx, ["BookgDt", "Dt"])?.textContent || "").trim() ||
        (getFirstDescendant(tx, ["ValDt", "Dt"])?.textContent || "").trim() ||
        bookingDate;
      bookings.push({ amount, date: txDate, reference: ref });
    });
  });

  return bookings;
}

function getFirstDescendant(root, pathParts) {
  let current = root;
  for (const tag of pathParts) {
    if (!current) return null;
    current = current.getElementsByTagNameNS("*", tag)?.[0] || null;
  }
  return current;
}

function findMatchingInvoiceForBooking(booking, openInvoices, matchedInvoiceIds) {
  const normalizedRef = normalizeLoose(booking.reference);
  const bookingAmount = roundAmount(booking.amount);

  const candidates = openInvoices.filter((invoice) => !matchedInvoiceIds.has(invoice.id));
  const amountCandidates = candidates.filter((invoice) => roundAmount(invoice.grandTotal) === bookingAmount);

  if (normalizedRef) {
    const byRefAndAmount = amountCandidates.find((invoice) => referenceMatchesInvoiceNo(normalizedRef, invoice.invoiceNo));
    if (byRefAndAmount) return { invoice: byRefAndAmount, matchType: "Referenz+Betrag" };
  }
  return null;
}

function findReferenceOnlyInvoice(booking, invoices, excludedIds) {
  const normalizedRef = normalizeLoose(booking.reference);
  if (!normalizedRef) return null;
  const excluded = excludedIds || new Set();
  return (invoices || []).find((invoice) => !excluded.has(invoice.id) && referenceMatchesInvoiceNo(normalizedRef, invoice.invoiceNo)) || null;
}

function findAmountOnlyInvoice(booking, invoices, excludedIds) {
  const amount = roundAmount(booking.amount);
  const bookingDate = parseIsoDateLoose(booking.date);
  const excluded = excludedIds || new Set();
  const candidates = (invoices || [])
    .filter((invoice) => !excluded.has(invoice.id))
    .filter((invoice) => roundAmount(invoice.grandTotal) === amount)
    .sort((a, b) => {
      const da = distanceToDate(bookingDate, parseIsoDateLoose(a.createdAt));
      const db = distanceToDate(bookingDate, parseIsoDateLoose(b.createdAt));
      return da - db;
    });
  return candidates[0] || null;
}

function collectBookingReference(rootNode) {
  const values = [];
  const pushNodeValues = (tagName) => {
    const nodes = [...rootNode.getElementsByTagNameNS("*", tagName)];
    nodes.forEach((n) => {
      const text = String(n?.textContent || "").trim();
      if (text) values.push(text);
    });
  };
  pushNodeValues("Ref");
  pushNodeValues("Ustrd");
  pushNodeValues("EndToEndId");
  pushNodeValues("TxId");
  pushNodeValues("InstrId");
  pushNodeValues("AddtlTxInf");
  pushNodeValues("AddtlNtryInf");
  return values.join(" ");
}

function referenceMatchesInvoiceNo(normalizedRef, invoiceNo) {
  const invoiceToken = normalizeLoose(invoiceNo);
  if (!invoiceToken) return false;
  const refCanonical = canonicalizeId(normalizedRef);
  const invoiceCanonical = canonicalizeId(invoiceToken);
  if (normalizedRef.includes(invoiceToken)) return true;
  if (invoiceToken.includes(normalizedRef)) return true;
  if (refCanonical.includes(invoiceCanonical)) return true;
  if (invoiceCanonical.includes(refCanonical)) return true;
  return false;
}

function canonicalizeId(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "")
    .replace(/\d+/g, (digits) => String(Number(digits)));
}

function normalizeLoose(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

function roundAmount(value) {
  return Math.round(toNumber(value, 0) * 100) / 100;
}

function parseIsoDateLoose(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  const d = new Date(raw);
  if (!Number.isNaN(d.getTime())) return d;
  const onlyDate = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (onlyDate) return new Date(`${onlyDate[1]}-${onlyDate[2]}-${onlyDate[3]}T00:00:00`);
  return null;
}

function distanceToDate(a, b) {
  if (!a || !b) return Number.MAX_SAFE_INTEGER;
  return Math.abs(a.getTime() - b.getTime());
}

function markInvoicePaidFromBooking(invoice, booking, importMeta = null) {
  invoice.paid = true;
  invoice.paidAt = booking?.date || new Date().toISOString().slice(0, 10);
  invoice.paidMethod = "automatisch";
  invoice.paidImportFile = String(importMeta?.fileName || "").trim();
  invoice.paidImportedAt = String(importMeta?.importedAt || "").trim();
  invoice.paidStatusSetAt = String(importMeta?.importedAt || new Date().toISOString()).trim();
}

function summarizeImportedEntities(imports) {
  const customerIds = new Set();
  const employeeIds = new Set();
  const itemIds = new Set();
  const entryIds = new Set();

  for (const pack of imports || []) {
    for (const c of pack?.customers || []) if (c?.id) customerIds.add(c.id);
    for (const e of pack?.employees || []) if (e?.id) employeeIds.add(e.id);
    for (const i of pack?.items || []) if (i?.id) itemIds.add(i.id);
    for (const e of pack?.entries || []) if (e?.id) entryIds.add(e.id);
  }

  return {
    customers: customerIds.size,
    employees: employeeIds.size,
    items: itemIds.size,
    entries: entryIds.size
  };
}

function normalizeItemRecord(item) {
  if (!item || typeof item !== "object") return { id: generateId(), name: "", price: 0, costPrice: 0, deleted: false, typeId: "", unitId: "", type: "", unit: "" };
  return {
    ...item,
    id: String(item.id || generateId()),
    name: String(item.name || "").trim(),
    price: toNumber(item.price, 0),
    costPrice: toNumber(item.costPrice, 0),
    deleted: Boolean(item.deleted),
    typeId: String(item.typeId || ""),
    unitId: String(item.unitId || ""),
    type: String(item.type || "").trim(),
    unit: String(item.unit || "").trim()
  };
}

function normalizeEntryRecord(entry) {
  if (!entry || typeof entry !== "object") return { id: generateId(), customerId: "", employeeId: "", itemId: "", date: "", quantity: 0, note: "", unitPrice: 0, costPrice: 0 };
  return {
    ...entry,
    id: String(entry.id || generateId()),
    customerId: String(entry.customerId || ""),
    employeeId: String(entry.employeeId || ""),
    itemId: String(entry.itemId || ""),
    date: String(entry.date || ""),
    quantity: toNumber(entry.quantity, 0),
    note: String(entry.note || ""),
    unitPrice: toNumber(entry.unitPrice, 0),
    costPrice: toNumber(entry.costPrice, 0)
  };
}


function getFilteredInvoices() {
  const filterMonth = String(invoiceFilterMonth?.value || "").trim();
  const filterCustomer = resolveInvoiceFilterCustomerId();
  const filterStatus = String(invoiceFilterStatus?.value || "all").trim();
  const filterSentStatus = String(invoiceFilterSentStatus?.value || "all").trim();
  return state.invoices
    .filter((i) => {
      const monthOk = !filterMonth || i.month === filterMonth;
      const customerOk = !filterCustomer || i.customerId === filterCustomer;
      const due = isOnOrBeforeToday(i.dueDate);
      const statusOk =
        filterStatus === "all" ||
        (filterStatus === "paid" && i.paid) ||
        (filterStatus === "open" && !i.paid) ||
        (filterStatus === "due" && !i.paid && due);
      const sentStatusOk =
        filterSentStatus === "all" ||
        (filterSentStatus === "sent" && Boolean(i.sent)) ||
        (filterSentStatus === "unsent" && !Boolean(i.sent));
      return monthOk && customerOk && statusOk && sentStatusOk;
    })
    .slice()
    .sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || "")));
}

function closeInvoiceDeleteAllDialog() {
  if (!invoiceDeleteAllDialog) return;
  invoiceDeleteAllDialog.hidden = true;
  pendingDeleteAllInvoiceIds = [];
  if (invoiceDeleteAllConfirmCheckbox) invoiceDeleteAllConfirmCheckbox.checked = false;
  if (invoiceDeleteAllContinueBtn) invoiceDeleteAllContinueBtn.disabled = true;
}
function renderInvoiceList() {
  if (!state.invoices.length) {
    invoiceList.innerHTML = "<small>Noch keine Rechnungen erstellt.</small>";
    return;
  }
  const visibleInvoices = getFilteredInvoices();

  if (!visibleInvoices.length) {
    invoiceList.innerHTML = "<small>Keine Rechnungen für den gewählten Filter vorhanden.</small>";
    return;
  }

    invoiceList.innerHTML = visibleInvoices
      .map((inv) => {
      const customer = state.customers.find((c) => c.id === inv.customerId);
      const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
      const isUnpaidDue = !inv.paid && isOnOrBeforeToday(inv.dueDate);
      const paidClass = inv.paid ? "paid" : isUnpaidDue ? "open-overdue" : "open-upcoming";
      const paidText = inv.paid ? "Bezahlt" : isUnpaidDue ? "Fällig" : "Offen";
      const isAutoPaid = String(inv.paidMethod || "").trim().toLowerCase() === "automatisch";
      const paidMetaHtml = inv.paid
        ? [
            `<small>Bezahlt am: ${escapeHtml(formatDateCH(inv.paidAt || inv.createdAt || ""))}</small>`,
            `<small>Status gesetzt: ${escapeHtml(formatPaidMethod(inv.paidMethod))}</small>`,
            `<small>Status gesetzt am: ${escapeHtml(formatDateTimeCH(inv.paidStatusSetAt || inv.paidImportedAt || inv.paidAt || inv.createdAt || ""))}</small>`,
            isAutoPaid && inv.paidImportedAt ? `<small>Importiert am: ${escapeHtml(formatDateTimeCH(inv.paidImportedAt))}</small>` : "",
            isAutoPaid && inv.paidImportFile ? `<small>Import-Datei: ${escapeHtml(inv.paidImportFile)}</small>` : ""
          ].filter(Boolean).join("")
        : isUnpaidDue
          ? `<small>Überfällig seit ${escapeHtml(formatDateCH(inv.dueDate))} (${getOverdueDays(inv.dueDate)} Tage)</small>`
          : `<small>Fällig am ${escapeHtml(formatDateCH(inv.dueDate))}</small>`;
      const sent = Boolean(inv.sent);
      const sentClass = sent ? "sent" : "unsent";
      const sentText = sent ? "Versendet" : "Nicht versendet";
      const sentMetaHtml = sent
        ? `<small>Versendet am: ${escapeHtml(formatDateCH(inv.sentAt || inv.createdAt || ""))}</small>`
        : `<small>Noch nicht versendet</small>`;
      return `
        <article class="invoice-row">
          <div class="top">
            <div>
              <strong>${escapeHtml(inv.invoiceNo)}</strong><br>
              <small>${escapeHtml(customerName)} | Monat ${escapeHtml(formatMonthCH(inv.month))}</small><br>
              <small>Fällig: ${escapeHtml(formatDateCH(inv.dueDate))} | Total: ${formatCurrency(inv.grandTotal)}</small>
            </div>
            <div class="status-cards">
              <section class="status-card">
                <div class="status-card-title">Versandstatus</div>
                <span class="pill ${sentClass}">
                  <strong>${sentText}</strong>
                  ${sentMetaHtml}
                </span>
              </section>
              <section class="status-card">
                <div class="status-card-title">Zahlungsstatus</div>
                <span class="pill ${paidClass}">
                  <strong>${paidText}</strong>
                  ${paidMetaHtml}
                </span>
              </section>
            </div>
          </div>
            <div class="actions">
              <div class="primary-actions">
                <button type="button" data-action="view" data-id="${escapeHtml(inv.id)}">Vorschau</button>
                <button type="button" data-action="print" data-id="${escapeHtml(inv.id)}">Drucken</button>
                <button type="button" data-action="save" data-id="${escapeHtml(inv.id)}">Speichern</button>
                <button type="button" class="danger" data-action="delete" data-id="${escapeHtml(inv.id)}">Löschen</button>
              </div>
              <div class="status-action-pair">
                <button type="button" class="${inv.sent ? "" : "primary"}" data-action="toggle-sent" data-id="${escapeHtml(inv.id)}">
                  ${inv.sent ? "Als nicht versendet markieren" : "Als versendet markieren"}
                </button>
                <button type="button" class="${inv.paid ? "" : "primary"}" data-action="toggle-paid" data-id="${escapeHtml(inv.id)}">
                  ${inv.paid ? "Als offen markieren" : "Als bezahlt markieren"}
                </button>
              </div>
            </div>
        </article>
      `;
      })
        .join("");
}

function resolveInvoiceFilterCustomerId() {
  const typed = String(invoiceFilterCustomer.value || "").trim();
  if (!typed) return "";
  return invoiceFilterCustomerNameToId.get(typed) || "";
}

function mergeImports(imports) {
  const customerMap = new Map((state.customers || []).map((c) => [c.id, c]));
  const employeeMap = new Map((state.employees || []).map((e) => [e.id, e]));
  const itemMap = new Map((state.items || []).map((i) => [i.id, i]));
  const entryMap = new Map((state.entries || []).map((e) => [e.id, e]));

  for (const pack of imports || []) {
    if (!pack || typeof pack !== "object") continue;
    for (const c of pack.customers || []) {
      if (!c?.id) continue;
      customerMap.set(c.id, c);
    }
    for (const e of pack.employees || []) {
      if (!e?.id) continue;
      employeeMap.set(e.id, {
        ...e,
        iban: String(e?.iban || "").trim(),
        hourlyWage: toNumber(e?.hourlyWage, 0)
      });
    }
    for (const i of pack.items || []) {
      const normalized = normalizeItemRecord(i);
      if (!normalized.id) continue;
      itemMap.set(normalized.id, normalized);
    }
    for (const e of pack.entries || []) {
      const normalized = normalizeEntryRecord(e);
      if (!normalized.id) continue;
      entryMap.set(normalized.id, normalized);
    }
  }

  state.customers = [...customerMap.values()];
  state.employees = [...employeeMap.values()];
  state.items = [...itemMap.values()];
  state.entries = [...entryMap.values()];
}

function normalizeBackup(parsed) {
  const fallback = baseState();
  const data = parsed?.data && typeof parsed.data === "object" ? parsed.data : parsed;
  return {
    company: data?.company && typeof data.company === "object" ? { ...fallback.company, ...data.company } : fallback.company,
    customers: Array.isArray(data?.customers) ? data.customers : [],
    employees: Array.isArray(data?.employees) ? data.employees : [],
    items: (Array.isArray(data?.items) ? data.items : []).map((item) => normalizeItemRecord(item)),
    entries: (Array.isArray(data?.entries) ? data.entries : []).map((entry) => normalizeEntryRecord(entry)),
    invoices: Array.isArray(data?.invoices) ? data.invoices.map((inv) => ({ ...inv, paymentSlipType: normalizePaymentSlipType(inv?.paymentSlipType), sent: Boolean(inv?.sent), sentAt: String(inv?.sentAt || "") })) : []
  };
}
function wireResetConfirmation() {
  const syncState = () => {
    resetDataBtn.disabled = !resetConfirmCheckbox.checked;
  };
  resetConfirmCheckbox.addEventListener("change", syncState);
  syncState();
}

function wireCompany() {
  companyForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    refreshCompanyValidationStates();
    if (!companyForm.checkValidity()) {
      companyForm.reportValidity();
      return;
    }
    const data = Object.fromEntries(new FormData(companyForm).entries());

    if (companyLogoInput.files?.[0]) {
      state.company.logoDataUrl = await fileToDataUrl(companyLogoInput.files[0]);
    }

    state.company.name = String(data.name || "").trim();
    state.company.street = String(data.street || "").trim();
    state.company.zip = String(data.zip || "").trim();
    state.company.city = String(data.city || "").trim();
    state.company.website = String(data.website || "").trim();
    state.company.email = String(data.email || "").trim();
    state.company.mobile = String(data.mobile || "").trim();
    state.company.iban = String(data.iban || "").trim().replace(/\s+/g, " ");
    state.company.currency = normalizeCurrencyCode(data.currency);
    state.company.vatNo = String(data.vatNo || "").trim();
    state.company.vatRate = toNumber(data.vatRate, 0);
    state.company.paymentTermDays = Math.max(0, Math.round(toNumber(data.paymentTermDays, 30)));
    state.company.invoiceText = String(data.invoiceText || "").trim();
    state.company.appendixText = String(data.appendixText || "").trim();

    saveState();
    companyStatus.textContent = "Firmendaten gespeichert.";
    refreshCompanyValidationStates();
  });
}

function wireCompanyValidation() {
  if (!companyForm) return;
  const requiredFields = companyForm.querySelectorAll("input[required], select[required], textarea[required]");
  requiredFields.forEach((field) => {
    const refresh = () => updateCompanyFieldState(field);
    field.addEventListener("input", refresh);
    field.addEventListener("change", refresh);
    field.addEventListener("blur", refresh);
  });
  if (companyForm.currency) {
    companyForm.currency.addEventListener("input", () => {
      const cursor = companyForm.currency.selectionStart ?? companyForm.currency.value.length;
      companyForm.currency.value = String(companyForm.currency.value || "").toUpperCase().slice(0, 3);
      companyForm.currency.setSelectionRange(cursor, cursor);
    });
  }
  if (companyLogoInput) {
    companyLogoInput.addEventListener("change", async () => {
      companyLogoInput.setCustomValidity("");
      if (companyLogoInput.files?.[0]) {
        const objectUrl = URL.createObjectURL(companyLogoInput.files[0]);
        setCompanyLogoPreview(objectUrl, "Neues Logo ausgewählt.");
      } else {
        setCompanyLogoPreview(String(state.company?.logoDataUrl || "").trim(), state.company?.logoDataUrl ? "Gespeichertes Logo." : "Noch kein Logo hochgeladen.");
      }
      updateCompanyLogoOptionalState();
    });
  }
  refreshCompanyValidationStates();
}

function wireImportBackupValidation() {
  const fields = [importFiles, camtImportFile, backupImportFile, invoiceCleanupMonths, invoiceCleanupConfirmCheckbox, resetConfirmCheckbox].filter(Boolean);
  fields.forEach((field) => {
    const refresh = () => updateSimpleFieldValidationState(field);
    field.addEventListener("change", refresh);
    field.addEventListener("input", refresh);
    refresh();
  });
  refreshImportBackupValidationStates();
}

function refreshImportBackupValidationStates() {
  const fields = [importFiles, camtImportFile, backupImportFile, invoiceCleanupMonths, invoiceCleanupConfirmCheckbox, resetConfirmCheckbox].filter(Boolean);
  fields.forEach((field) => updateSimpleFieldValidationState(field));
}

function updateSimpleFieldValidationState(field) {
  if (!field) return;
  const required = Boolean(field.required);
  let hasValue = true;
  if (field.type === "file") {
    hasValue = Boolean(field.files?.length);
  } else if (field.type === "checkbox") {
    hasValue = Boolean(field.checked);
  } else {
    hasValue = String(field.value ?? "").trim().length > 0;
  }

  if (required) {
    const isValid = hasValue && field.checkValidity();
    field.classList.remove("optional-filled");
    field.classList.toggle("required-empty", !isValid);
    field.classList.toggle("required-filled", isValid);
    const label = field.closest("label");
    if (!label) return;
    label.classList.remove("optional-indicator", "optional-ok");
    label.classList.add("required-indicator");
    label.classList.toggle("required-empty", !isValid);
    label.classList.toggle("required-filled", isValid);
    return;
  }

  field.classList.remove("required-empty", "required-filled");
  field.classList.add("optional-filled");
  const label = field.closest("label");
  if (!label) return;
  label.classList.remove("required-indicator", "required-empty", "required-filled");
  label.classList.add("optional-indicator", "optional-ok");
}

function refreshCompanyValidationStates() {
  if (!companyForm) return;
  const requiredFields = companyForm.querySelectorAll("input[required], select[required], textarea[required]");
  requiredFields.forEach((field) => updateCompanyFieldState(field));
  const optionalFields = companyForm.querySelectorAll("input:not([required]), select:not([required]), textarea:not([required])");
  optionalFields.forEach((field) => updateCompanyOptionalFieldState(field));
  updateCompanyLogoOptionalState();
}

function updateCompanyFieldState(field) {
  const value = String(field?.value ?? "").trim();
  const hasValue = value.length > 0;
  const isValid = hasValue && field.checkValidity();
  field.classList.toggle("required-empty", !isValid);
  field.classList.toggle("required-filled", isValid);
  const label = field.closest("label");
  if (!label) return;
  label.classList.add("required-indicator");
  label.classList.remove("optional-indicator", "optional-ok");
  label.classList.toggle("required-empty", !isValid);
  label.classList.toggle("required-filled", isValid);
}

function updateCompanyOptionalFieldState(field) {
  field.classList.remove("required-empty", "required-filled");
  field.classList.add("optional-filled");
  if (field?.id === "companyLogoInput") return;
  const label = field.closest("label");
  if (!label) return;
  label.classList.remove("required-indicator", "required-empty", "required-filled");
  label.classList.add("optional-indicator", "optional-ok");
}

function updateCompanyLogoOptionalState() {
  const logoLabel = companyForm?.querySelector("label[data-logo-field='true']");
  if (!logoLabel) return;
  const hasSelectedFile = Boolean(companyLogoInput?.files?.[0]);
  const hasStoredLogo = Boolean(String(state.company?.logoDataUrl || "").trim());
  const isOk = hasSelectedFile || hasStoredLogo;
  if (companyLogoInput) {
    companyLogoInput.classList.remove("required-empty", "required-filled");
    companyLogoInput.classList.add("optional-filled");
  }
  logoLabel.classList.remove("required-indicator", "required-empty", "required-filled");
  logoLabel.classList.remove("optional-indicator", "optional-ok");
  if (isOk) {
    logoLabel.classList.add("optional-indicator", "optional-ok");
  }
}

function setCompanyLogoPreview(src, text) {
  if (companyLogoPreviewObjectUrl) {
    URL.revokeObjectURL(companyLogoPreviewObjectUrl);
    companyLogoPreviewObjectUrl = "";
  }
  const safeSrc = String(src || "").trim();
  if (safeSrc.startsWith("blob:")) {
    companyLogoPreviewObjectUrl = safeSrc;
  }
  if (companyLogoPreview) {
    if (safeSrc) {
      companyLogoPreview.src = safeSrc;
      companyLogoPreview.hidden = false;
    } else {
      companyLogoPreview.hidden = true;
      companyLogoPreview.removeAttribute("src");
    }
  }
  if (companyLogoPreviewText) {
    companyLogoPreviewText.textContent = String(text || "");
  }
}

function wireBilling() {
  if (billingSuccessCloseBtn) {
    billingSuccessCloseBtn.addEventListener("click", closeBillingSuccessDialog);
  }
  if (billingSuccessDialog) {
    billingSuccessDialog.addEventListener("click", (event) => {
      if (event.target === billingSuccessDialog) closeBillingSuccessDialog();
    });
  }
  if (billingMonth && invoiceFilterMonth) {
    billingMonth.addEventListener("change", () => {
      invoiceFilterMonth.value = billingMonth.value || "";
      renderInvoiceList();
    });
  }
  billingForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const customerId = billingCustomer.value;
    const month = billingMonth.value;
    const paymentSlipType = normalizePaymentSlipType(billingPaymentSlip?.value);
    if (!customerId || !month) return;
    const customerIds = customerId === "all" ? state.customers.map((c) => c.id) : [customerId];
    const createdInvoices = [];
    let alreadyExisting = 0;
    let missingData = 0;

    customerIds.forEach((id) => {
      const existing = state.invoices.find((i) => i.customerId === id && i.month === month);
      if (existing) {
        alreadyExisting += 1;
        return;
      }
      const invoice = buildInvoice(id, month, paymentSlipType);
      if (!invoice) {
        missingData += 1;
        return;
      }
      state.invoices.push(invoice);
      createdInvoices.push(invoice);
    });

    if (!createdInvoices.length && !alreadyExisting && !missingData) {
      billingStatus.textContent = "Keine Kunden verfügbar.";
      setPreviewHtml("");
      return;
    }

    saveState();
    renderInvoiceList();

    if (createdInvoices.length) {
      setPreviewHtml(renderInvoiceHtml(createdInvoices[0]));
    } else {
      setPreviewHtml("");
    }

    if (customerId === "all") {
      billingStatus.textContent = `Monat ${formatMonthCH(month)}: ${createdInvoices.length} Rechnung(en) erstellt, ${alreadyExisting} bereits vorhanden, ${missingData} ohne Daten.`;
    } else if (createdInvoices.length === 1) {
      billingStatus.textContent = `Rechnung erstellt (${createdInvoices[0].invoiceNo}).`;
    } else if (alreadyExisting) {
      const existing = state.invoices.find((i) => i.customerId === customerId && i.month === month);
      if (existing) {
        setPreviewHtml(renderInvoiceHtml(existing));
        billingStatus.textContent = `Rechnung existiert bereits (${existing.invoiceNo}) und wurde geladen.`;
      } else {
        billingStatus.textContent = "Rechnung existiert bereits.";
      }
    } else {
      billingStatus.textContent = "Keine Daten für die gewählte Kombination vorhanden.";
    }

    if (createdInvoices.length || alreadyExisting || missingData) {
      showBillingSuccessDialog(month, createdInvoices.length, alreadyExisting, missingData, customerId);
    }
  });
}

function wireInvoiceList() {
  invoiceFilterMonth.addEventListener("change", () => {
    renderInvoiceList();
  });
  invoiceFilterCustomer.addEventListener("input", () => {
    renderInvoiceList();
  });
  invoiceFilterCustomer.addEventListener("change", () => {
    renderInvoiceList();
  });
  invoiceFilterStatus.addEventListener("change", () => {
    renderInvoiceList();
  });
  invoiceFilterSentStatus?.addEventListener("change", () => {
    renderInvoiceList();
  });
  invoiceFilterClearBtn.addEventListener("click", () => {
    invoiceFilterMonth.value = "";
    invoiceFilterCustomer.value = "";
    invoiceFilterStatus.value = "all";
    if (invoiceFilterSentStatus) invoiceFilterSentStatus.value = "all";
    renderInvoiceList();
  });

  if (invoiceDeleteAllBtn) {
    invoiceDeleteAllBtn.addEventListener("click", () => {
      const visible = getFilteredInvoices();
      if (!visible.length) {
        billingStatus.textContent = "Keine Rechnungen für den gewählten Filter vorhanden.";
        return;
      }
      pendingDeleteAllInvoiceIds = visible.map((inv) => inv.id);
      if (invoiceDeleteAllList) {
        invoiceDeleteAllList.innerHTML = visible
          .map((inv) => {
            const customer = state.customers.find((c) => c.id === inv.customerId);
            const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
            return `<article class="cleanup-result-row"><strong>${escapeHtml(inv.invoiceNo || "(ohne Nr.)")}</strong><br>${escapeHtml(customerName)} | ${escapeHtml(formatMonthCH(inv.month || ""))} | ${escapeHtml(formatCurrency(inv.grandTotal))}</article>`;
          })
          .join("");
      }
      if (invoiceDeleteAllConfirmCheckbox) invoiceDeleteAllConfirmCheckbox.checked = false;
      if (invoiceDeleteAllContinueBtn) invoiceDeleteAllContinueBtn.disabled = true;
      if (invoiceDeleteAllDialog) invoiceDeleteAllDialog.hidden = false;
    });
  }

  invoiceDeleteAllConfirmCheckbox?.addEventListener("change", () => {
    if (invoiceDeleteAllContinueBtn) {
      invoiceDeleteAllContinueBtn.disabled = !invoiceDeleteAllConfirmCheckbox.checked;
    }
  });

  invoiceDeleteAllCancelBtn?.addEventListener("click", () => {
    closeInvoiceDeleteAllDialog();
  });

  invoiceDeleteAllContinueBtn?.addEventListener("click", () => {
    const affected = state.invoices.filter((inv) => pendingDeleteAllInvoiceIds.includes(inv.id));
    if (!affected.length) {
      closeInvoiceDeleteAllDialog();
      billingStatus.textContent = "Keine Rechnungen mehr zum Löschen vorhanden.";
      return;
    }
    const listText = affected
      .map((inv) => {
        const customer = state.customers.find((c) => c.id === inv.customerId);
        const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
        return `- ${inv.invoiceNo || "(ohne Nr.)"} | ${customerName} | ${formatMonthCH(inv.month || "")}`;
      })
      .join("\n");
    if (!confirm(`Sind Sie wirklich sicher?\n\nBetroffen:\n${listText}`)) return;

    state.invoices = state.invoices.filter((inv) => !pendingDeleteAllInvoiceIds.includes(inv.id));
    saveState();
    renderInvoiceList();
    setPreviewHtml("");
    closeInvoiceDeleteAllDialog();
    billingStatus.textContent = `${affected.length} Rechnung(en) gelöscht.`;
  });
  invoiceList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action][data-id]");
    if (!button) return;
    const invoiceId = button.dataset.id;
    const action = button.dataset.action;
    const invoice = state.invoices.find((i) => i.id === invoiceId);
    if (!invoice) return;

    if (action === "view") {
      const html = renderInvoiceHtml(invoice);
      setPreviewHtml(html);
      openPrintWindow(html, `Rechnung-${invoice.invoiceNo || "ohne-nummer"}`, { autoPrint: false, closeAfterPrint: false });
      billingStatus.textContent = `Vorschau in neuem Tab geöffnet (${invoice.invoiceNo}).`;
      return;
    }

    if (action === "toggle-paid") {
      const question = invoice.paid
        ? `Rechnung ${invoice.invoiceNo} wirklich als offen markieren?`
        : `Rechnung ${invoice.invoiceNo} wirklich als bezahlt markieren?`;
      if (!confirm(question)) return;
      if (invoice.paid) {
        invoice.paid = false;
        invoice.paidAt = "";
        invoice.paidMethod = "";
        invoice.paidImportFile = "";
        invoice.paidImportedAt = "";
        invoice.paidStatusSetAt = "";
      } else {
        invoice.paid = true;
        invoice.paidAt = new Date().toISOString().slice(0, 10);
        invoice.paidMethod = "manuell";
        invoice.paidImportFile = "";
        invoice.paidImportedAt = "";
        invoice.paidStatusSetAt = new Date().toISOString();
      }
      saveState();
      renderInvoiceList();
      return;
    }

        if (action === "toggle-sent") {
      const question = invoice.sent
        ? `Rechnung ${invoice.invoiceNo} wirklich als nicht versendet markieren?`
        : `Rechnung ${invoice.invoiceNo} wirklich als versendet markieren?`;
      if (!confirm(question)) return;
      if (invoice.sent) {
        invoice.sent = false;
        invoice.sentAt = "";
      } else {
        invoice.sent = true;
        invoice.sentAt = new Date().toISOString().slice(0, 10);
      }
      saveState();
      renderInvoiceList();
      return;
    }

    if (action === "delete") {
      if (!confirm(`Rechnung ${invoice.invoiceNo} wirklich löschen?`)) return;
      state.invoices = state.invoices.filter((i) => i.id !== invoice.id);
      saveState();
      renderInvoiceList();
      if (currentInvoiceHtml && invoicePreview.innerHTML.includes(invoice.invoiceNo)) {
        setPreviewHtml("");
      }
      return;
    }

    if (action === "print") {
        const html = renderInvoiceHtml(invoice);
        setPreviewHtml(html);
        openPrintWindow(html, `Rechnung-${invoice.invoiceNo || "ohne-nummer"}`);
        return;
    }

    if (action === "save") {
        const html = renderInvoiceHtml(invoice);
        setPreviewHtml(html);
        saveInvoiceAsPdf(html, `Rechnung-${invoice.invoiceNo || "ohne-nummer"}`);
        return;
    }
  });
}

function wireHoursReport() {
  if (!hoursMonth) return;
  hoursMonth.addEventListener("change", () => {
    renderHoursReport();
  });
  hoursSubTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      activateHoursTab(tab.dataset.hoursTab || "mitarbeiter");
    });
  });
  if (hoursByEmployee) {
    hoursByEmployee.addEventListener("input", (event) => {
      const input = event.target.closest("#hoursEmployeeFilter");
      if (!input) return;
      hoursEmployeeFilterText = String(input.value || "");
      renderHoursReport();
    });
    hoursByEmployee.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-action]");
      if (!button) return;
      const action = String(button.dataset.action || "");
      const context = getFilteredHoursContext();

      if (["save-employee-hours-pdf", "preview-employee-hours-pdf", "export-employee-hours-csv"].includes(action)) {
        if (!context.month) {
          billingStatus.textContent = "Bitte zuerst einen Monat auswählen.";
          return;
        }
        if (!context.filteredEntries.length) {
          billingStatus.textContent = "Keine Stundenpositionen für den ausgewählten Monat vorhanden.";
          return;
        }
        if (action === "save-employee-hours-pdf") {
          const html = renderEmployeeHoursReportHtml(context.month, context.filteredEntries);
          saveInvoiceAsPdf(html, `Mitarbeiter-Stundenauswertung-${context.month}`);
          return;
        }
        if (action === "preview-employee-hours-pdf") {
          const html = renderEmployeeHoursReportHtml(context.month, context.filteredEntries);
          openPrintWindow(html, `Mitarbeiter-Stundenauswertung-${context.month}`);
          return;
        }
        if (action === "export-employee-hours-csv") {
          const csv = buildEmployeeHoursCsv(context.month, context.filteredEntries);
          const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
          downloadBlob(blob, `Mitarbeiter-Stundenauswertung-${context.month}.csv`);
          return;
        }
      }

      if (!["print-employee", "save-employee"].includes(action)) return;
      if (!context.month) {
        billingStatus.textContent = "Bitte zuerst einen Monat auswählen.";
        return;
      }
      const employeeIdRaw = String(button.dataset.id || "");
      const employeeId = employeeIdRaw === "__unassigned__" ? "" : employeeIdRaw;
      const employeeEntries = context.filteredEntries.filter((entry) => String(entry.employeeId || "") === employeeId);
      if (!employeeEntries.length) {
        billingStatus.textContent = "Keine Daten für den ausgewählten Mitarbeiter vorhanden.";
        return;
      }
      const html = renderEmployeeSettlementHtml(employeeId, context.month, employeeEntries);
      const employee = state.employees.find((e) => e.id === employeeId);
      const employeeName = employee ? formatEmployeeName(employee) : "Mitarbeiter";
      const docTitle = `Mitarbeiter-Abrechnung-${employeeName}-${context.month || ""}`;
      if (action === "print-employee") {
        openPrintWindow(html, docTitle);
        return;
      }
      if (action === "save-employee") {
        saveInvoiceAsPdf(html, docTitle);
      }
    });
  }
  if (hoursByCustomer) {
    hoursByCustomer.addEventListener("input", (event) => {
      const input = event.target.closest("#hoursCustomerFilter");
      if (!input) return;
      hoursCustomerFilterText = String(input.value || "");
      renderHoursReport();
    });
    hoursByCustomer.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-action]");
      if (!button) return;
      const action = String(button.dataset.action || "");
      const context = getFilteredHoursContext();
      if (["save-customer-hours-pdf", "preview-customer-hours-pdf", "export-customer-hours-csv"].includes(action)) {
        if (!context.month) {
          billingStatus.textContent = "Bitte zuerst einen Monat auswählen.";
          return;
        }
        if (!context.filteredEntries.length) {
          billingStatus.textContent = "Keine Stundenpositionen für den ausgewählten Monat vorhanden.";
          return;
        }
        if (action === "save-customer-hours-pdf") {
          const html = renderCustomerHoursReportHtml(context.month, context.filteredEntries);
          saveInvoiceAsPdf(html, `Kunden-Stundenauswertung-${context.month}`);
          return;
        }
        if (action === "preview-customer-hours-pdf") {
          const html = renderCustomerHoursReportHtml(context.month, context.filteredEntries);
          openPrintWindow(html, `Kunden-Stundenauswertung-${context.month}`);
          return;
        }
        if (action === "export-customer-hours-csv") {
          const csv = buildCustomerHoursCsv(context.month, context.filteredEntries);
          const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
          downloadBlob(blob, `Kunden-Stundenauswertung-${context.month}.csv`);
          return;
        }
      }

      if (!["print-customer", "save-customer"].includes(action)) return;
      if (!context.month) {
        billingStatus.textContent = "Bitte zuerst einen Monat auswählen.";
        return;
      }
      const customerIdRaw = String(button.dataset.id || "");
      const customerId = customerIdRaw === "__unknown_customer__" ? "" : customerIdRaw;
      const customerEntries = context.filteredEntries.filter((entry) => String(entry.customerId || "") === customerId);
      if (!customerEntries.length) {
        billingStatus.textContent = "Keine Daten für den ausgewählten Kunden vorhanden.";
        return;
      }
      const html = renderCustomerHoursReportHtml(context.month, customerEntries);
      const customer = state.customers.find((c) => c.id === customerId);
      const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
      const docTitle = `Kunden-Stundenauswertung-${customerName}-${context.month || ""}`;
      if (action === "print-customer") {
        openPrintWindow(html, docTitle);
        return;
      }
      if (action === "save-customer") {
        saveInvoiceAsPdf(html, docTitle);
      }
    });
  }
  if (hoursByProductsEmployee) {
    hoursByProductsEmployee.addEventListener("input", (event) => {
      const input = event.target.closest("#hoursProductsEmployeeFilter");
      if (!input) return;
      hoursProductsEmployeeFilterText = String(input.value || "");
      renderHoursReport();
    });
    hoursByProductsEmployee.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-action]");
      if (!button) return;
      const action = String(button.dataset.action || "");
      const context = getFilteredProductContext();
      if (!context.month) {
        billingStatus.textContent = "Bitte zuerst einen Monat auswählen.";
        return;
      }
      if (!context.filteredEntries.length) {
        billingStatus.textContent = "Keine Produktpositionen (Stk) für den ausgewählten Monat vorhanden.";
        return;
      }
      if (action === "save-product-employee-report-pdf") {
        const html = renderProductEmployeeReportHtml(context.month, context.filteredEntries);
        saveInvoiceAsPdf(html, `Produkte-Mitarbeiter-Auswertung-${context.month}`);
        return;
      }
      if (action === "preview-product-employee-report-pdf") {
        const html = renderProductEmployeeReportHtml(context.month, context.filteredEntries);
        openPrintWindow(html, `Produkte-Mitarbeiter-Auswertung-${context.month}`);
        return;
      }
      if (action === "export-product-employee-report-csv") {
        const csv = buildProductEmployeeReportCsv(context.month, context.filteredEntries);
        const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        downloadBlob(blob, `Produkte-Mitarbeiter-Auswertung-${context.month}.csv`);
      }
    });
  }
  if (hoursByProductsCustomer) {
    hoursByProductsCustomer.addEventListener("input", (event) => {
      const input = event.target.closest("#hoursProductsCustomerFilter");
      if (!input) return;
      hoursProductsCustomerFilterText = String(input.value || "");
      renderHoursReport();
    });
    hoursByProductsCustomer.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-action]");
      if (!button) return;
      const action = String(button.dataset.action || "");
      const context = getFilteredProductContext();
      if (!context.month) {
        billingStatus.textContent = "Bitte zuerst einen Monat auswählen.";
        return;
      }
      if (!context.filteredEntries.length) {
        billingStatus.textContent = "Keine Produktpositionen (Stk) für den ausgewählten Monat vorhanden.";
        return;
      }
      if (action === "save-product-customer-report-pdf") {
        const html = renderProductCustomerReportHtml(context.month, context.filteredEntries);
        saveInvoiceAsPdf(html, `Produkte-Kunden-Auswertung-${context.month}`);
        return;
      }
      if (action === "preview-product-customer-report-pdf") {
        const html = renderProductCustomerReportHtml(context.month, context.filteredEntries);
        openPrintWindow(html, `Produkte-Kunden-Auswertung-${context.month}`);
        return;
      }
      if (action === "export-product-customer-report-csv") {
        const csv = buildProductCustomerReportCsv(context.month, context.filteredEntries);
        const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        downloadBlob(blob, `Produkte-Kunden-Auswertung-${context.month}.csv`);
      }
    });
  }

  if (hoursByProfit) {
    hoursByProfit.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-action]");
      if (!button) return;
      const action = String(button.dataset.action || "");
      const context = getFilteredHoursContext();
      if (!context.month) {
        billingStatus.textContent = "Bitte zuerst einen Monat auswählen.";
        return;
      }
      if (!context.filteredEntries.length) {
        billingStatus.textContent = "Keine Stundenpositionen für den ausgewählten Monat vorhanden.";
        return;
      }
      if (action === "save-hours-profit-report-pdf") {
        const html = renderHoursProfitReportHtml(context.month, context.filteredEntries);
        saveInvoiceAsPdf(html, `Stunden-Gewinn-Auswertung-${context.month}`, { orientation: "landscape" });
        return;
      }
      if (action === "preview-hours-profit-report-pdf") {
        const html = renderHoursProfitReportHtml(context.month, context.filteredEntries);
        openPrintWindow(html, `Stunden-Gewinn-Auswertung-${context.month}`, { orientation: "landscape" });
        return;
      }
      if (action === "export-hours-profit-report-csv") {
        const csv = buildHoursProfitCsv(context.month, context.filteredEntries);
        const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        downloadBlob(blob, `Stunden-Gewinn-Auswertung-${context.month}.csv`);
      }
    });
  }
  if (hoursByProductsProfit) {
    hoursByProductsProfit.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-action]");
      if (!button) return;
      const action = String(button.dataset.action || "");
      const context = getFilteredProductContext();
      if (!context.month) {
        billingStatus.textContent = "Bitte zuerst einen Monat auswählen.";
        return;
      }
      if (!context.filteredEntries.length) {
        billingStatus.textContent = "Keine Produktpositionen (Stk) für den ausgewählten Monat vorhanden.";
        return;
      }
      if (action === "save-products-profit-report-pdf") {
        const html = renderProductsProfitReportHtml(context.month, context.filteredEntries);
        saveInvoiceAsPdf(html, `Produkte-Gewinn-Auswertung-${context.month}`, { orientation: "landscape" });
        return;
      }
      if (action === "preview-products-profit-report-pdf") {
        const html = renderProductsProfitReportHtml(context.month, context.filteredEntries);
        openPrintWindow(html, `Produkte-Gewinn-Auswertung-${context.month}`, { orientation: "landscape" });
        return;
      }
      if (action === "export-products-profit-report-csv") {
        const csv = buildProductsProfitCsv(context.month, context.filteredEntries);
        const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        downloadBlob(blob, `Produkte-Gewinn-Auswertung-${context.month}.csv`);
      }
    });
  }
  if (hoursByVat) {
    hoursByVat.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-action]");
      if (!button) return;
      const action = String(button.dataset.action || "");
      const month = String(hoursMonth?.value || "");
      const monthEntries = state.entries.filter((entry) => String(entry.date || "").startsWith(month));
      if (!month) {
        billingStatus.textContent = "Bitte zuerst einen Monat auswählen.";
        return;
      }
      if (!monthEntries.length) {
        billingStatus.textContent = "Keine Erfassungen für den ausgewählten Monat vorhanden.";
        return;
      }
      const vatRate = toNumber(state.company?.vatRate, 0);
      if (action === "save-vat-report-pdf") {
        const html = renderVatReportHtml(month, monthEntries, vatRate);
        saveInvoiceAsPdf(html, `Mehrwertsteuer-Auswertung-${month}`);
        return;
      }
      if (action === "preview-vat-report-pdf") {
        const html = renderVatReportHtml(month, monthEntries, vatRate);
        openPrintWindow(html, `Mehrwertsteuer-Auswertung-${month}`);
        return;
      }
      if (action === "export-vat-report-csv") {
        const csv = buildVatCsv(month, monthEntries, vatRate);
        const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        downloadBlob(blob, `Mehrwertsteuer-Auswertung-${month}.csv`);
      }
    });
  }
  activateHoursTab(activeHoursTab);
}

function activateHoursTab(target) {
  activeHoursTab = ["mitarbeiter", "kunden", "tage", "stunden-gewinn", "mehrwertsteuer", "produkte-mitarbeiter", "produkte-kunden", "produkte-gewinn", "produkte-tage"].includes(target) ? target : "mitarbeiter";
  hoursSubTabs.forEach((tab) => {
    const active = tab.dataset.hoursTab === activeHoursTab;
    tab.classList.toggle("active", active);
    tab.setAttribute("aria-selected", String(active));
  });
  if (hoursCustomerPanel) {
    const active = activeHoursTab === "kunden";
    hoursCustomerPanel.classList.toggle("active", active);
    hoursCustomerPanel.hidden = !active;
  }
  if (hoursEmployeePanel) {
    const active = activeHoursTab === "mitarbeiter";
    hoursEmployeePanel.classList.toggle("active", active);
    hoursEmployeePanel.hidden = !active;
  }
  if (hoursProductsEmployeePanel) {
    const active = activeHoursTab === "produkte-mitarbeiter";
    hoursProductsEmployeePanel.classList.toggle("active", active);
    hoursProductsEmployeePanel.hidden = !active;
  }
  if (hoursProductsCustomerPanel) {
    const active = activeHoursTab === "produkte-kunden";
    hoursProductsCustomerPanel.classList.toggle("active", active);
    hoursProductsCustomerPanel.hidden = !active;
  }
  if (hoursProductsProfitPanel) {
    const active = activeHoursTab === "produkte-gewinn";
    hoursProductsProfitPanel.classList.toggle("active", active);
    hoursProductsProfitPanel.hidden = !active;
  }
  if (hoursDaysPanel) {
    const active = activeHoursTab === "tage";
    hoursDaysPanel.classList.toggle("active", active);
    hoursDaysPanel.hidden = !active;
  }
  if (hoursProfitPanel) {
    const active = activeHoursTab === "stunden-gewinn";
    hoursProfitPanel.classList.toggle("active", active);
    hoursProfitPanel.hidden = !active;
  }
  if (hoursVatPanel) {
    const active = activeHoursTab === "mehrwertsteuer";
    hoursVatPanel.classList.toggle("active", active);
    hoursVatPanel.hidden = !active;
  }
  if (hoursProductDaysPanel) {
    const active = activeHoursTab === "produkte-tage";
    hoursProductDaysPanel.classList.toggle("active", active);
    hoursProductDaysPanel.hidden = !active;
  }
}

function renderAll() {
  renderHeaderLogo();
  renderCompanyForm();
  renderBillingCustomers();
  renderInvoiceFilterCustomers();
  renderDefaultMonth();
  renderDefaultInvoiceFilterMonth();
  renderDefaultHoursMonth();
  renderHoursReport();
  renderInvoiceList();
  setPreviewHtml(currentInvoiceHtml);
}

function renderHeaderLogo() {
  if (!headerLogo) return;
  headerLogo.hidden = true;
  headerLogo.removeAttribute("src");
}

function renderCompanyForm() {
  companyForm.name.value = state.company.name || "";
  companyForm.street.value = state.company.street || "";
  companyForm.zip.value = state.company.zip || "";
  companyForm.city.value = state.company.city || "";
  companyForm.website.value = state.company.website || "";
  companyForm.email.value = state.company.email || "";
  companyForm.mobile.value = state.company.mobile || "";
  companyForm.iban.value = state.company.iban || "";
  companyForm.currency.value = normalizeCurrencyCode(state.company.currency);
  companyForm.vatNo.value = state.company.vatNo || "";
  companyForm.vatRate.value = String(state.company.vatRate ?? 8.1);
  companyForm.paymentTermDays.value = String(state.company.paymentTermDays ?? 30);
  companyForm.invoiceText.value = state.company.invoiceText || "";
  if (companyForm.appendixText) companyForm.appendixText.value = state.company.appendixText || "";
  setCompanyLogoPreview(String(state.company.logoDataUrl || "").trim(), state.company.logoDataUrl ? "Gespeichertes Logo." : "Noch kein Logo hochgeladen.");
  refreshCompanyValidationStates();
}

function renderBillingCustomers() {
  if (!state.customers.length) {
    billingCustomer.innerHTML = "<option value=''>Zuerst Exporte importieren</option>";
    billingCustomer.disabled = true;
    return;
  }
  billingCustomer.disabled = false;
  const selected = billingCustomer.value || "all";
  billingCustomer.innerHTML = [`<option value="all">Alle Kunden</option>`]
    .concat(
      state.customers.map((c) => `<option value="${escapeHtml(c.id)}">${escapeHtml(formatCustomerName(c))}</option>`)
    )
    .join("");
  billingCustomer.value = state.customers.some((c) => c.id === selected) || selected === "all" ? selected : "all";
}

function renderInvoiceFilterCustomers() {
  const selected = String(invoiceFilterCustomer.value || "").trim();
  const sortedCustomers = state.customers
    .slice()
    .sort((a, b) => formatCustomerName(a).localeCompare(formatCustomerName(b), "de-CH"));
  invoiceFilterCustomerNameToId = new Map(
    sortedCustomers.map((c) => [formatCustomerName(c), c.id])
  );
  invoiceFilterCustomerList.innerHTML = sortedCustomers
    .map((c) => `<option value="${escapeHtml(formatCustomerName(c))}"></option>`)
    .join("");
  if (selected && !invoiceFilterCustomerNameToId.has(selected)) {
    invoiceFilterCustomer.value = "";
  }
}

function renderDefaultMonth() {
  if (!billingMonth.value) {
    const now = new Date();
    now.setMonth(now.getMonth() - 1);
    billingMonth.value = now.toISOString().slice(0, 7);
  }
}

function renderDefaultInvoiceFilterMonth() {
  if (!invoiceFilterMonth.value) {
    const now = new Date();
    now.setMonth(now.getMonth() - 1);
    invoiceFilterMonth.value = now.toISOString().slice(0, 7);
  }
}

function renderDefaultHoursMonth() {
  if (hoursMonth && !hoursMonth.value) {
    hoursMonth.value = new Date().toISOString().slice(0, 7);
  }
}

function renderHoursReport() {
  if (!hoursSummary || !hoursByCustomer || !hoursByEmployee || !hoursByProductsEmployee || !hoursByProductsCustomer || !hoursByProductsProfit || !hoursByProfit || !hoursByDays || !hoursByProductDays || !hoursByVat) return;
  const activeElement = document.activeElement;
  const keepCustomerFilterFocus = activeElement?.id === "hoursCustomerFilter";
  const keepEmployeeFilterFocus = activeElement?.id === "hoursEmployeeFilter";
  const keepProductsEmployeeFilterFocus = activeElement?.id === "hoursProductsEmployeeFilter";
  const keepProductsCustomerFilterFocus = activeElement?.id === "hoursProductsCustomerFilter";
  const customerFilterSelectionStart = keepCustomerFilterFocus ? activeElement.selectionStart : null;
  const customerFilterSelectionEnd = keepCustomerFilterFocus ? activeElement.selectionEnd : null;
  const employeeFilterSelectionStart = keepEmployeeFilterFocus ? activeElement.selectionStart : null;
  const employeeFilterSelectionEnd = keepEmployeeFilterFocus ? activeElement.selectionEnd : null;
  const productsEmployeeFilterSelectionStart = keepProductsEmployeeFilterFocus ? activeElement.selectionStart : null;
  const productsEmployeeFilterSelectionEnd = keepProductsEmployeeFilterFocus ? activeElement.selectionEnd : null;
  const productsCustomerFilterSelectionStart = keepProductsCustomerFilterFocus ? activeElement.selectionStart : null;
  const productsCustomerFilterSelectionEnd = keepProductsCustomerFilterFocus ? activeElement.selectionEnd : null;

  const hoursContext = getFilteredHoursContext();
  const { month, filteredEntries } = hoursContext;
  const productContext = getFilteredProductContext();
  const productEntries = productContext.filteredEntries;

  if (!month) {
    hoursSummary.innerHTML = "<small>Bitte Monat wählen.</small>";
    hoursByCustomer.innerHTML = "";
    hoursByEmployee.innerHTML = "";
    hoursByProductsEmployee.innerHTML = "";
    hoursByProductsCustomer.innerHTML = "";
    hoursByProductsProfit.innerHTML = "";
    hoursByProfit.innerHTML = "";
    hoursByVat.innerHTML = "";
    hoursByDays.innerHTML = "";
    hoursByProductDays.innerHTML = "";
    return;
  }

  const monthEntries = state.entries.filter((entry) => String(entry.date || "").startsWith(month));
  const monthTotalAmount = monthEntries.reduce((sum, entry) => sum + toNumber(entry.unitPrice, 0) * toNumber(entry.quantity, 0), 0);
  const monthTotalProfit = monthEntries.reduce((sum, entry) => {
    const qty = toNumber(entry.quantity, 0);
    const revenue = qty * toNumber(entry.unitPrice, 0);
    const item = state.items.find((i) => i.id === String(entry.itemId || ""));
    let cost = 0;
    if (isHourlyItem(item)) {
      const employee = state.employees.find((e) => e.id === String(entry.employeeId || ""));
      cost = qty * toNumber(employee?.hourlyWage, 0);
    } else if (isPieceItem(item)) {
      const costPrice = toNumber(entry.costPrice, toNumber(item?.costPrice, 0));
      cost = qty * costPrice;
    }
    return sum + (revenue - cost);
  }, 0);
  const monthTotalCost = monthTotalAmount - monthTotalProfit;
  const vatRate = toNumber(state.company?.vatRate, 0);
  const monthVatAmount = monthTotalAmount * (vatRate / 100);
  const monthTotalInclVat = monthTotalAmount + monthVatAmount;
  const totalHoursQty = filteredEntries.reduce((sum, entry) => sum + toNumber(entry.quantity, 0), 0);
  const totalProductsQty = productEntries.reduce((sum, entry) => sum + toNumber(entry.quantity, 0), 0);

  hoursSummary.innerHTML = `
    <article class="hours-kpi">
      <small>Monat</small>
      <strong>${escapeHtml(formatMonthCH(month))}</strong>
    </article>
    <article class="hours-kpi">
      <small>Erfassungen total</small>
      <strong>${monthEntries.length}</strong>
    </article>
    <article class="hours-kpi">
      <small>Stunden total</small>
      <strong>${formatNumberCH(totalHoursQty)}</strong>
    </article>
    <article class="hours-kpi">
      <small>Stk total</small>
      <strong>${formatNumberCH(totalProductsQty)}</strong>
    </article>
    <article class="hours-kpi">
      <small>Umsatz total (alle)</small>
      <strong>${formatCurrency(monthTotalAmount)}</strong>
    </article>
    <article class="hours-kpi">
      <small>MwSt-Satz</small>
      <strong>${formatNumberCH(vatRate)}%</strong>
    </article>
    <article class="hours-kpi">
      <small>MwSt-Betrag (alle)</small>
      <strong>${formatCurrency(monthVatAmount)}</strong>
    </article>
    <article class="hours-kpi">
      <small>Umsatz inkl. MwSt</small>
      <strong>${formatCurrency(monthTotalInclVat)}</strong>
    </article>
    <article class="hours-kpi">
      <small>Kosten total (alle)</small>
      <strong>${formatCurrency(monthTotalCost)}</strong>
    </article>
    <article class="hours-kpi">
      <small>Gewinn total (alle)</small>
      <strong>${formatCurrency(monthTotalProfit)}</strong>
    </article>
    <div class="hours-summary-alert" role="note" aria-label="Hinweis"><span class="hours-summary-alert-icon" aria-hidden="true">!</span><span class="hours-summary-alert-text">Hinweis: Vom Gewinn müssen natürlich noch sämtliche andere anfallenden Kosten abgezogen werden (Sozialabgaben Löhne, Versicherungskosten etc.).</span></div>
  `;

  const emptyHoursMessage = `<small>Keine Stundenpositionen für ${escapeHtml(formatMonthCH(month))} vorhanden.</small>`;
  const emptyProductsMessage = `<small>Keine Produktpositionen (Stk) für ${escapeHtml(formatMonthCH(month))} vorhanden.</small>`;
  const emptyVatMessage = `<small>Keine Erfassungen für ${escapeHtml(formatMonthCH(month))} vorhanden.</small>`;

  const customerFilter = String(hoursCustomerFilterText || "");
  hoursByCustomer.innerHTML = filteredEntries.length
    ? renderCustomerSummaryWithDetails(filteredEntries, customerFilter)
    : emptyHoursMessage;

  const employeeFilter = String(hoursEmployeeFilterText || "");
  hoursByEmployee.innerHTML = filteredEntries.length
    ? renderEmployeeSummaryTable(filteredEntries, employeeFilter) + renderEmployeeHoursDetails(filteredEntries, employeeFilter)
    : emptyHoursMessage;

  const productsEmployeeFilter = String(hoursProductsEmployeeFilterText || "");
  hoursByProductsEmployee.innerHTML = productEntries.length
    ? renderProductsByEmployeeTable(productEntries, productsEmployeeFilter)
    : emptyProductsMessage;

  const productsCustomerFilter = String(hoursProductsCustomerFilterText || "");
  hoursByProductsCustomer.innerHTML = productEntries.length
    ? renderProductsByCustomerTable(productEntries, productsCustomerFilter)
    : emptyProductsMessage;

  hoursByProductsProfit.innerHTML = productEntries.length
    ? renderReportActions("preview-products-profit-report-pdf", "save-products-profit-report-pdf", "export-products-profit-report-csv") + renderProductsProfitDetails(productEntries)
    : emptyProductsMessage;

  hoursByProfit.innerHTML = filteredEntries.length
    ? renderReportActions("preview-hours-profit-report-pdf", "save-hours-profit-report-pdf", "export-hours-profit-report-csv") + renderHoursProfitDetails(filteredEntries)
    : emptyHoursMessage;

  hoursByVat.innerHTML = monthEntries.length
    ? renderReportActions("preview-vat-report-pdf", "save-vat-report-pdf", "export-vat-report-csv") + renderVatBreakdown(monthEntries, vatRate)
    : emptyVatMessage;

  hoursByDays.innerHTML = filteredEntries.length
    ? renderHoursCalendarOverview(month, filteredEntries)
    : emptyHoursMessage;

  hoursByProductDays.innerHTML = productEntries.length
    ? renderProductCalendarOverview(month, productEntries)
    : emptyProductsMessage;

  if (keepCustomerFilterFocus && activeHoursTab === "kunden") {
    const filterInput = hoursByCustomer.querySelector("#hoursCustomerFilter");
    if (filterInput) {
      filterInput.focus();
      const maxPos = filterInput.value.length;
      const startPos = Math.max(0, Math.min(maxPos, Number(customerFilterSelectionStart ?? maxPos)));
      const endPos = Math.max(0, Math.min(maxPos, Number(customerFilterSelectionEnd ?? maxPos)));
      if (typeof filterInput.setSelectionRange === "function") filterInput.setSelectionRange(startPos, endPos);
    }
  }

  if (keepEmployeeFilterFocus && activeHoursTab === "mitarbeiter") {
    const filterInput = hoursByEmployee.querySelector("#hoursEmployeeFilter");
    if (filterInput) {
      filterInput.focus();
      const maxPos = filterInput.value.length;
      const startPos = Math.max(0, Math.min(maxPos, Number(employeeFilterSelectionStart ?? maxPos)));
      const endPos = Math.max(0, Math.min(maxPos, Number(employeeFilterSelectionEnd ?? maxPos)));
      if (typeof filterInput.setSelectionRange === "function") filterInput.setSelectionRange(startPos, endPos);
    }
  }

  if (keepProductsEmployeeFilterFocus && activeHoursTab === "produkte-mitarbeiter") {
    const filterInput = hoursByProductsEmployee.querySelector("#hoursProductsEmployeeFilter");
    if (filterInput) {
      filterInput.focus();
      const maxPos = filterInput.value.length;
      const startPos = Math.max(0, Math.min(maxPos, Number(productsEmployeeFilterSelectionStart ?? maxPos)));
      const endPos = Math.max(0, Math.min(maxPos, Number(productsEmployeeFilterSelectionEnd ?? maxPos)));
      if (typeof filterInput.setSelectionRange === "function") filterInput.setSelectionRange(startPos, endPos);
    }
  }

  if (keepProductsCustomerFilterFocus && activeHoursTab === "produkte-kunden") {
    const filterInput = hoursByProductsCustomer.querySelector("#hoursProductsCustomerFilter");
    if (filterInput) {
      filterInput.focus();
      const maxPos = filterInput.value.length;
      const startPos = Math.max(0, Math.min(maxPos, Number(productsCustomerFilterSelectionStart ?? maxPos)));
      const endPos = Math.max(0, Math.min(maxPos, Number(productsCustomerFilterSelectionEnd ?? maxPos)));
      if (typeof filterInput.setSelectionRange === "function") filterInput.setSelectionRange(startPos, endPos);
    }
  }
}


function renderReportActions(previewAction, saveAction, csvAction) {
  return `
    <div class="hours-actions">
      <button type="button" data-action="${escapeHtml(previewAction)}">Drucken</button>
      <button type="button" data-action="${escapeHtml(saveAction)}">Speichern</button>
      <button type="button" data-action="${escapeHtml(csvAction)}">CSV (Excel)</button>
    </div>
  `;
}
function renderVatBreakdown(entries, vatRate) {
  const byCustomer = new Map();
  entries.forEach((entry) => {
    const customerId = String(entry.customerId || "");
    const customer = state.customers.find((c) => c.id === customerId);
    const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
    if (!byCustomer.has(customerId)) {
      byCustomer.set(customerId, { customerName, net: 0 });
    }
    const row = byCustomer.get(customerId);
    row.net += toNumber(entry.unitPrice, 0) * toNumber(entry.quantity, 0);
  });

  const rows = [...byCustomer.values()].sort((a, b) => b.net - a.net);
  const body = rows
    .map((row) => {
      const vat = row.net * (vatRate / 100);
      const gross = row.net + vat;
      return `
        <tr>
          <td>${escapeHtml(row.customerName)}</td>
          <td class="num">${formatCurrency(row.net)}</td>
          <td class="num">${formatNumberCH(vatRate)}%</td>
          <td class="num">${formatCurrency(vat)}</td>
          <td class="num">${formatCurrency(gross)}</td>
        </tr>
      `;
    })
    .join("");

  const totalNet = rows.reduce((sum, row) => sum + row.net, 0);
  const totalVat = totalNet * (vatRate / 100);
  const totalGross = totalNet + totalVat;

  return `
    <h3>Mehrwertsteuer nach Kunde</h3>
    <table class="hours-table">
      <thead>
        <tr>
          <th>Kunde</th>
          <th class="num">Netto</th>
          <th class="num">MwSt-Satz</th>
          <th class="num">MwSt-Betrag</th>
          <th class="num">Brutto</th>
        </tr>
      </thead>
      <tbody>${body}</tbody>
      <tfoot>
        <tr>
          <th>Total</th>
          <th class="num">${formatCurrency(totalNet)}</th>
          <th class="num">${formatNumberCH(vatRate)}%</th>
          <th class="num">${formatCurrency(totalVat)}</th>
          <th class="num">${formatCurrency(totalGross)}</th>
        </tr>
      </tfoot>
    </table>
  `;
}
function renderHoursCalendarOverview(month, entries) {
  if (!month) return "<small>Bitte Monat auswählen.</small>";

  const monthMatch = month.match(/^(\d{4})-(\d{2})$/);
  if (!monthMatch) return "<small>Ungültiger Monat.</small>";

  const year = Number(monthMatch[1]);
  const monthIndex = Number(monthMatch[2]) - 1;
  if (!Number.isFinite(year) || !Number.isFinite(monthIndex) || monthIndex < 0 || monthIndex > 11) {
    return "<small>Ungültiger Monat.</small>";
  }

  const lastDate = new Date(year, monthIndex + 1, 0).getDate();

  const byDay = new Map();
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

  entries.forEach((entry) => {
    const date = String(entry?.date || "").slice(0, 10);
    if (!date.startsWith(`${month}-`)) return;

    const customer = state.customers.find((c) => c.id === String(entry.customerId || ""));
    const customerName = getCustomerShort(customer);

    const employee = state.employees.find((e) => e.id === String(entry.employeeId || ""));
    const employeeName = getEmployeeShort(employee);

    const qty = toNumber(entry.quantity, 0);
    const unitPrice = toNumber(entry.unitPrice, 0);
    const amount = unitPrice * qty;

    if (!byDay.has(date)) byDay.set(date, new Map());
    const dayMap = byDay.get(date);

    const key = `${customerName}|||${employeeName}|||${unitPrice}`;
    const current = dayMap.get(key) || { customerName, employeeName, qty: 0, unitPrice, amount: 0 };
    current.qty += qty;
    current.amount += amount;
    dayMap.set(key, current);
  });

  const shortWeekdays = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];

  const body = Array.from({ length: lastDate }, (_, offset) => {
    const day = offset + 1;
    const isoDate = `${year}-${String(monthIndex + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    const dayMap = byDay.get(isoDate) || new Map();
    const rows = [...dayMap.values()].sort((a, b) => b.amount - a.amount || a.customerName.localeCompare(b.customerName, "de-CH"));
    const totalHours = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
    const totalAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  const totalProfit = rows.reduce((sum, row) => sum + toNumber(row.profit, 0), 0);

    const weekdayIndex = new Date(year, monthIndex, day).getDay();
    const weekdayShort = shortWeekdays[weekdayIndex] || "";
    const isWeekend = weekdayIndex === 0 || weekdayIndex === 6;

    const rowsHtml = rows
      .map((row) => `
        <li class="hours-calendar-row hours-money-row">
          <span class="col-customer">${escapeHtml(row.customerName)}</span>
          <span class="col-employee">${escapeHtml(row.employeeName)}</span>
          <strong class="col-hours">${formatNumberCH(row.qty)} h</strong>
          <strong class="col-rate">${escapeHtml(formatCurrency(row.unitPrice))}</strong>
          <strong class="col-total">${escapeHtml(formatCurrency(row.amount))}</strong>
        </li>
      `)
      .join("");

    const detailsHtml = rows.length
      ? `
        <div class="hours-calendar-cell-columns hours-money-columns"><span>Kunde</span><span>Mitarbeiter</span><span>Stunden</span><span>Ansatz</span><span>Total</span></div>
        <ul class="hours-calendar-list">${rowsHtml}</ul>
        <div class="hours-calendar-total hours-money-total">
          <span>Total</span>
          <strong>${formatNumberCH(totalHours)} h</strong>
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

  return `
    <h3>Kalenderübersicht ${escapeHtml(formatMonthCH(month))}</h3>
    <div class="hours-day-list">${body}</div>
  `;
}
function renderProductCalendarOverview(month, entries) {
  if (!month) return "<small>Bitte Monat auswählen.</small>";

  const monthMatch = month.match(/^(\d{4})-(\d{2})$/);
  if (!monthMatch) return "<small>Ungültiger Monat.</small>";

  const year = Number(monthMatch[1]);
  const monthIndex = Number(monthMatch[2]) - 1;
  if (!Number.isFinite(year) || !Number.isFinite(monthIndex) || monthIndex < 0 || monthIndex > 11) {
    return "<small>Ungültiger Monat.</small>";
  }

  const lastDate = new Date(year, monthIndex + 1, 0).getDate();
  const byDay = new Map();

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

  entries.forEach((entry) => {
    const date = String(entry?.date || "").slice(0, 10);
    if (!date.startsWith(`${month}-`)) return;

    const customer = state.customers.find((c) => c.id === String(entry.customerId || ""));
    const customerName = getCustomerShort(customer);
    const employee = state.employees.find((e) => e.id === String(entry.employeeId || ""));
    const employeeName = getEmployeeShort(employee);
    const item = state.items.find((i) => i.id === String(entry.itemId || ""));
    const productName = item?.name || "Unbekanntes Produkt";

    const qty = toNumber(entry.quantity, 0);
    const unitPrice = toNumber(entry.unitPrice, 0);
    const amount = qty * unitPrice;

    if (!byDay.has(date)) byDay.set(date, new Map());
    const dayMap = byDay.get(date);
    const key = `${customerName}|||${employeeName}|||${productName}|||${unitPrice}`;
    const current = dayMap.get(key) || { customerName, employeeName, productName, qty: 0, unitPrice, amount: 0 };
    current.qty += qty;
    current.amount += amount;
    dayMap.set(key, current);
  });

  const shortWeekdays = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];
  const body = Array.from({ length: lastDate }, (_, offset) => {
    const day = offset + 1;
    const isoDate = `${year}-${String(monthIndex + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    const dayMap = byDay.get(isoDate) || new Map();
    const rows = [...dayMap.values()].sort((a, b) => b.amount - a.amount || a.customerName.localeCompare(b.customerName, "de-CH"));
    const totalQty = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
    const totalAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  const totalProfit = rows.reduce((sum, row) => sum + toNumber(row.profit, 0), 0);

    const weekdayIndex = new Date(year, monthIndex, day).getDay();
    const weekdayShort = shortWeekdays[weekdayIndex] || "";
    const isWeekend = weekdayIndex === 0 || weekdayIndex === 6;

    const rowsHtml = rows
      .map((row) => `
        <li class="hours-calendar-row hours-money-row products-calendar-row">
          <span class="col-customer">${escapeHtml(row.customerName)}</span>
          <span class="col-employee">${escapeHtml(row.employeeName)}</span>
          <span class="col-product">${escapeHtml(row.productName)}</span>
          <strong class="col-hours">${formatNumberCH(row.qty)} Stk</strong>
          <strong class="col-rate">${escapeHtml(formatCurrency(row.unitPrice))}</strong>
          <strong class="col-total">${escapeHtml(formatCurrency(row.amount))}</strong>
        </li>
      `)
      .join("");

    const detailsHtml = rows.length
      ? `
        <div class="hours-calendar-cell-columns hours-money-columns products-calendar-columns"><span>Kunde</span><span>Mitarbeiter</span><span>Produkt</span><span>Stk</span><span>Ansatz</span><span>Total</span></div>
        <ul class="hours-calendar-list">${rowsHtml}</ul>
        <div class="hours-calendar-total hours-money-total products-calendar-total">
          <span>Total</span>
          <strong>${formatNumberCH(totalQty)} Stk</strong>
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

  return `
    <h3>Produkte Kalenderübersicht ${escapeHtml(formatMonthCH(month))}</h3>
    <div class="hours-day-list">${body}</div>
  `;
}
function isHourlyItem(item) {
  if (!item || typeof item !== "object") return false;
  const unitId = String(item.unitId || "").trim().toLowerCase();
  const unit = String(item.unit || "").trim().toLowerCase();
  if (unitId === "unit-h") return true;
  if (["h", "std", "stunde", "stunden", "stunde/n", "hour", "hours"].includes(unit)) return true;
  return unit.includes("stund");
}
function getFilteredHoursContext() {
  const month = String(hoursMonth?.value || "").trim();
  const monthEntries = state.entries.filter((entry) => String(entry.date || "").startsWith(month));
  const filteredEntries = monthEntries.filter((entry) => {
    const item = state.items.find((i) => i.id === entry.itemId);
    return isHourlyItem(item);
  });
  return { month, filteredEntries };
}

function isPieceItem(item) {
  if (!item || typeof item !== "object") return false;
  const unitId = String(item.unitId || "").trim().toLowerCase();
  const unit = String(item.unit || "").trim().toLowerCase();
  if (unitId === "unit-stk") return true;
  return ["stk", "stk.", "stück", "stueck", "piece", "pieces", "pcs"].includes(unit);
}

function getFilteredProductContext() {
  const month = String(hoursMonth?.value || "").trim();
  const monthEntries = state.entries.filter((entry) => String(entry.date || "").startsWith(month));
  const filteredEntries = monthEntries.filter((entry) => {
    const item = state.items.find((i) => i.id === entry.itemId);
    return isPieceItem(item);
  });
  return { month, filteredEntries };
}

function buildProductReportRows(entries) {
  const grouped = new Map();
  entries.forEach((entry) => {
    const itemId = String(entry.itemId || "");
    const item = state.items.find((i) => i.id === itemId);
    const productName = item?.name || "Unbekanntes Produkt";
    const productType = item?.type || item?.typeId || "-";
    if (!grouped.has(itemId)) {
      grouped.set(itemId, { itemId, productName, productType, count: 0, qty: 0, amount: 0, details: [] });
    }
    const row = grouped.get(itemId);
    const qty = toNumber(entry.quantity, 0);
    const unitPrice = toNumber(entry.unitPrice, toNumber(item?.price, 0));
    const total = qty * unitPrice;
    const customer = state.customers.find((c) => c.id === String(entry.customerId || ""));
    const employee = state.employees.find((e) => e.id === String(entry.employeeId || ""));
    row.count += 1;
    row.qty += qty;
    row.amount += total;
    row.costAmount += costTotal;
    row.profit += profit;
    row.details.push({
      date: String(entry.date || ""),
      customerName: customer ? formatCustomerName(customer) : "Unbekannter Kunde",
      employeeName: employee ? formatEmployeeName(employee) : "Nicht zugewiesen",
      qty,
      unitPrice,
      total
    });
  });

  return [...grouped.values()]
    .map((row) => ({
      ...row,
      details: row.details.sort((a, b) => String(a.date).localeCompare(String(b.date)))
    }))
    .sort((a, b) => b.qty - a.qty || b.amount - a.amount || String(a.productName).localeCompare(String(b.productName), "de-CH"));
}

function renderProductsSummaryTable(entries, searchTerm = "") {
  const rows = buildProductReportRows(entries);
  const visibleRows = rows.filter((row) => matchesSearch(row.productName, searchTerm));
  const body = visibleRows.map((row) => `
    <tr>
      <td>${escapeHtml(row.productName)}</td>
      <td>${escapeHtml(String(row.productType || "-"))}</td>
      <td class="num">${row.count}</td>
      <td class="num">${formatNumberCH(row.qty)}</td>
      <td class="num">${formatCurrency(row.amount)}</td>
    </tr>
  `).join("");

  return `
    <h3>Nach Produkt (Einheit Stk)</h3>
    <div class="hours-actions">
      <button type="button" data-action="preview-product-report-pdf">Drucken</button>
      <button type="button" data-action="save-product-report-pdf">Speichern</button>
      <button type="button" data-action="export-product-report-csv">CSV (Excel)</button>
    </div>
    <label>Produkt suchen
      <input id="hoursProductsFilter" type="search" placeholder="Produktname..." value="${escapeHtml(searchTerm)}" autocomplete="off" />
    </label>
    <table class="hours-table">
      <thead>
        <tr>
          <th>Produkt</th>
          <th>Typ</th>
          <th class="num">Erfassungen</th>
          <th class="num">Stk</th>
          <th class="num">Betrag</th>
        </tr>
      </thead>
      <tbody>${body || `<tr><td colspan="5"><small>Keine Produkte passend zur Suche gefunden.</small></td></tr>`}</tbody>
    </table>
  `;
}

function renderProductReportHtml(month, entries) {
  const company = state.company || {};
  const rows = buildProductReportRows(entries);
  const totalQty = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
  const totalAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  const totalProfit = rows.reduce((sum, row) => sum + toNumber(row.profit, 0), 0);

  const sectionsHtml = rows.map((row) => {
    const detailRows = row.details.map((detail) => `
      <tr>
        <td>${escapeHtml(formatWeekdayShortCH(detail.date))}</td>
        <td>${escapeHtml(formatDateCH(detail.date))}</td>
        <td>${escapeHtml(detail.customerName)}</td>
        <td>${escapeHtml(detail.employeeName)}</td>
        <td class="num">${formatNumberCH(detail.qty)}</td>
        <td class="num">${formatAmountCH(detail.unitPrice)}</td>
        <td class="num">${formatAmountCH(detail.total)}</td>
      </tr>
    `).join("");

    return `
      <section class="invoice-text-box">
        <h3>${escapeHtml(row.productName)} <small>(${formatNumberCH(row.qty)} Stk | ${formatCurrency(row.amount)})</small></h3>
        <table class="invoice-lines hours-table">
          <thead>
            <tr>
              <th>Tag</th>
              <th>Datum</th>
              <th>Kunde</th>
              <th>Mitarbeiter</th>
              <th class="num">Stk</th>
              <th class="num">Ansatz</th>
              <th class="num">Total</th>
            </tr>
          </thead>
          <tbody>${detailRows || `<tr><td colspan="6"><small>Keine Details vorhanden.</small></td></tr>`}</tbody>
        </table>
      </section>
    `;
  }).join("");

  return `
    <section class="invoice employee-settlement">
      <div class="invoice-top">
        <div>
          ${company.logoDataUrl ? `<img class="logo" src="${company.logoDataUrl}" alt="Logo" />` : ""}
          <p>
            <strong>${escapeHtml(company.name || "")}</strong><br>
            ${escapeHtml(company.street || "")}<br>
            ${escapeHtml(company.zip || "")} ${escapeHtml(company.city || "")}
          </p>
        </div>
        <div class="invoice-address-box">
          <strong>Produktauswertung (Stk)</strong><br>
          Monat: ${escapeHtml(formatMonthCH(month))}
        </div>
      </div>
      <hr>
      <p><strong>Monat:</strong> ${escapeHtml(formatMonthCH(month))}</p>
      <p><strong>Währung:</strong> ${escapeHtml(normalizeCurrencyCode(company.currency))}</p>
      <p><strong>Total Stück:</strong> ${formatNumberCH(totalQty)} | <strong>Total Betrag:</strong> ${formatCurrency(totalAmount)} | <strong>Total Gewinn:</strong> ${formatCurrency(totalProfit)}</p>
      <p><small>Hinweis: Alle Beträge ohne MwSt.</small></p>
      ${sectionsHtml || "<small>Keine Produktpositionen vorhanden.</small>"}
    </section>
  `;
}

function buildProductReportCsv(month, entries) {
  const rows = buildProductReportRows(entries);
  const lines = [];
  lines.push(["Monat", "\u200E" + formatMonthLongCH(month)].map(csvEscape).join(";"));
  lines.push(["Produkt", "Typ", "Wochentag", "Datum", "Kunde", "Mitarbeiter", "Stk", "Ansatz", "Total"].map(csvEscape).join(";"));
  const totalMonthQty = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
  const totalMonthAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  rows.forEach((row) => {
    row.details.forEach((detail) => {
      lines.push([
        row.productName,
        row.productType,
        formatWeekdayLongCH(detail.date),
        formatDateCH(detail.date),
        detail.customerName,
        detail.employeeName,
        formatNumberCH(detail.qty),
        formatAmountCH(detail.unitPrice),
        formatAmountCH(detail.total)
      ].map(csvEscape).join(";"));
    });
    lines.push([
      `${row.productName} Total`,
      "",
      "",
      "",
      "",
      "",
      formatNumberCH(row.qty),
      "",
      formatAmountCH(row.amount)
    ].map(csvEscape).join(";"));
  });
  lines.push(["Total Monat", "", "", "", "", "", formatNumberCH(totalMonthQty), "", formatAmountCH(totalMonthAmount)].map(csvEscape).join(";"));
  return "\uFEFF" + lines.join("\r\n");
}
function buildProductByEmployeeRows(entries) {
  const grouped = new Map();
  entries.forEach((entry) => {
    const employeeId = String(entry.employeeId || "");
    const employee = state.employees.find((e) => e.id === employeeId);
    const employeeName = employee ? formatEmployeeName(employee) : "Nicht zugewiesen";
    if (!grouped.has(employeeId)) {
      grouped.set(employeeId, { employeeId, employeeName, count: 0, qty: 0, amount: 0, costAmount: 0, profit: 0, details: [] });
    }
    const row = grouped.get(employeeId);
    const qty = toNumber(entry.quantity, 0);
    const unitPrice = toNumber(entry.unitPrice, 0);
    const total = qty * unitPrice;
    const customer = state.customers.find((c) => c.id === String(entry.customerId || ""));
    const item = state.items.find((i) => i.id === String(entry.itemId || ""));
    const costPrice = toNumber(entry.costPrice, toNumber(item?.costPrice, 0));
    const costTotal = qty * costPrice;
    const profit = total - costTotal;
    row.count += 1;
    row.qty += qty;
    row.amount += total;
    row.costAmount += costTotal;
    row.profit += profit;
    row.details.push({
      date: String(entry.date || ""),
      customerName: customer ? formatCustomerName(customer) : "Unbekannter Kunde",
      productName: item?.name || "Unbekanntes Produkt",
      qty,
      unitPrice,
      total
    });
  });
  return [...grouped.values()]
    .map((row) => ({ ...row, details: row.details.sort((a, b) => String(a.date).localeCompare(String(b.date))) }))
    .sort((a, b) => b.qty - a.qty || b.amount - a.amount || String(a.employeeName).localeCompare(String(b.employeeName), "de-CH"));
}

function buildProductByCustomerRows(entries) {
  const grouped = new Map();
  entries.forEach((entry) => {
    const customerId = String(entry.customerId || "");
    const customer = state.customers.find((c) => c.id === customerId);
    const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
    if (!grouped.has(customerId)) {
      grouped.set(customerId, { customerId, customerName, count: 0, qty: 0, amount: 0, costAmount: 0, profit: 0, details: [] });
    }
    const row = grouped.get(customerId);
    const qty = toNumber(entry.quantity, 0);
    const unitPrice = toNumber(entry.unitPrice, 0);
    const total = qty * unitPrice;
    const employee = state.employees.find((e) => e.id === String(entry.employeeId || ""));
    const item = state.items.find((i) => i.id === String(entry.itemId || ""));
    const costPrice = toNumber(entry.costPrice, toNumber(item?.costPrice, 0));
    const costTotal = qty * costPrice;
    const profit = total - costTotal;
    row.count += 1;
    row.qty += qty;
    row.amount += total;
    row.costAmount += costTotal;
    row.profit += profit;
    row.details.push({
      date: String(entry.date || ""),
      employeeName: employee ? formatEmployeeName(employee) : "Nicht zugewiesen",
      productName: item?.name || "Unbekanntes Produkt",
      qty,
      unitPrice,
      total
    });
  });
  return [...grouped.values()]
    .map((row) => ({ ...row, details: row.details.sort((a, b) => String(a.date).localeCompare(String(b.date))) }))
    .sort((a, b) => b.qty - a.qty || b.amount - a.amount || String(a.customerName).localeCompare(String(b.customerName), "de-CH"));
}

function renderProductsByEmployeeTable(entries, searchTerm = "") {
  const rows = buildProductByEmployeeRows(entries);
  const visibleRows = rows.filter((row) => matchesSearch(row.employeeName, searchTerm));
  const body = visibleRows.map((row) => `
    <tr>
      <td>${escapeHtml(row.employeeName)}</td>
      <td class="num">${row.count}</td>
      <td class="num">${formatNumberCH(row.qty)}</td>
      <td class="num">${formatCurrency(row.amount)}</td>
      <td class="num">${formatCurrency(row.profit)}</td>
    </tr>
  `).join("");
  const detailsHtml = visibleRows.map((row) => {
    const detailRows = row.details.map((detail) => `
      <tr>
        <td>${escapeHtml(formatWeekdayShortCH(detail.date))}</td>
        <td>${escapeHtml(formatDateCH(detail.date))}</td>
        <td>${escapeHtml(detail.customerName)}</td>
        <td>${escapeHtml(detail.productName)}</td>
        <td class="num">${formatNumberCH(detail.qty)}</td>
        <td class="num">${formatAmountCH(detail.total)}</td>
      </tr>
    `).join("");
    return `
      <section class="hours-detail-block">
        <h4>${escapeHtml(row.employeeName)} <small>(${formatNumberCH(row.qty)} Stk | ${formatCurrency(row.amount)} | Gewinn ${formatCurrency(row.profit)})</small></h4>
        <table class="hours-table">
          <thead>
            <tr>
              <th>Tag</th>
              <th>Datum</th>
              <th>Kunde</th>
              <th>Produkt</th>
              <th class="num">Stk</th>
              <th class="num">Total</th>
            </tr>
          </thead>
          <tbody>${detailRows || `<tr><td colspan="6"><small>Keine Details vorhanden.</small></td></tr>`}</tbody>
        </table>
      </section>
    `;
  }).join("");
  return `
    <h3>Produkte nach Mitarbeiter (Stk)</h3>
    <div class="hours-actions">
      <button type="button" data-action="preview-product-employee-report-pdf">Drucken</button>
      <button type="button" data-action="save-product-employee-report-pdf">Speichern</button>
      <button type="button" data-action="export-product-employee-report-csv">CSV (Excel)</button>
    </div>
    <label>Mitarbeiter suchen
      <input id="hoursProductsEmployeeFilter" type="search" placeholder="Name..." value="${escapeHtml(searchTerm)}" autocomplete="off" />
    </label>
    <table class="hours-table">
      <thead>
        <tr>
          <th>Mitarbeiter</th>
          <th class="num">Erfassungen</th>
          <th class="num">Stk</th>
          <th class="num">Betrag</th>
          <th class="num">Gewinn</th>
        </tr>
      </thead>
      <tbody>${body || `<tr><td colspan="5"><small>Keine Mitarbeiter passend zur Suche gefunden.</small></td></tr>`}</tbody>
    </table>
    <div class="hours-detail-wrap">
      <h3>Details je Mitarbeiter</h3>
      ${detailsHtml || "<small>Keine Detaildaten passend zur Suche gefunden.</small>"}
    </div>
  `;
}

function renderProductsByCustomerTable(entries, searchTerm = "") {
  const rows = buildProductByCustomerRows(entries);
  const visibleRows = rows.filter((row) => matchesSearch(row.customerName, searchTerm));
  const body = visibleRows.map((row) => `
    <tr>
      <td>${escapeHtml(row.customerName)}</td>
      <td class="num">${row.count}</td>
      <td class="num">${formatNumberCH(row.qty)}</td>
      <td class="num">${formatCurrency(row.amount)}</td>
      <td class="num">${formatCurrency(row.profit)}</td>
    </tr>
  `).join("");
  const detailsHtml = visibleRows.map((row) => {
    const detailRows = row.details.map((detail) => `
      <tr>
        <td>${escapeHtml(formatWeekdayShortCH(detail.date))}</td>
        <td>${escapeHtml(formatDateCH(detail.date))}</td>
        <td>${escapeHtml(detail.employeeName)}</td>
        <td>${escapeHtml(detail.productName)}</td>
        <td class="num">${formatNumberCH(detail.qty)}</td>
        <td class="num">${formatAmountCH(detail.total)}</td>
      </tr>
    `).join("");
    return `
      <section class="hours-detail-block">
        <h4>${escapeHtml(row.customerName)} <small>(${formatNumberCH(row.qty)} Stk | ${formatCurrency(row.amount)} | Gewinn ${formatCurrency(row.profit)})</small></h4>
        <table class="hours-table">
          <thead>
            <tr>
              <th>Tag</th>
              <th>Datum</th>
              <th>Mitarbeiter</th>
              <th>Produkt</th>
              <th class="num">Stk</th>
              <th class="num">Total</th>
            </tr>
          </thead>
          <tbody>${detailRows || `<tr><td colspan="6"><small>Keine Details vorhanden.</small></td></tr>`}</tbody>
        </table>
      </section>
    `;
  }).join("");
  return `
    <h3>Produkte nach Kunden (Stk)</h3>
    <div class="hours-actions">
      <button type="button" data-action="preview-product-customer-report-pdf">Drucken</button>
      <button type="button" data-action="save-product-customer-report-pdf">Speichern</button>
      <button type="button" data-action="export-product-customer-report-csv">CSV (Excel)</button>
    </div>
    <label>Kunde suchen
      <input id="hoursProductsCustomerFilter" type="search" placeholder="Name..." value="${escapeHtml(searchTerm)}" autocomplete="off" />
    </label>
    <table class="hours-table">
      <thead>
        <tr>
          <th>Kunde</th>
          <th class="num">Erfassungen</th>
          <th class="num">Stk</th>
          <th class="num">Betrag</th>
          <th class="num">Gewinn</th>
        </tr>
      </thead>
      <tbody>${body || `<tr><td colspan="5"><small>Keine Kunden passend zur Suche gefunden.</small></td></tr>`}</tbody>
    </table>
    <div class="hours-detail-wrap">
      <h3>Details je Kunde</h3>
      ${detailsHtml || "<small>Keine Detaildaten passend zur Suche gefunden.</small>"}
    </div>
  `;
}

function renderHoursProfitDetails(entries) {
  const rows = (entries || [])
    .map((entry) => {
      const customer = state.customers.find((c) => c.id === String(entry.customerId || ""));
      const employee = state.employees.find((e) => e.id === String(entry.employeeId || ""));
      const item = state.items.find((i) => i.id === String(entry.itemId || ""));
      const qty = toNumber(entry.quantity, 0);
      const unitPrice = toNumber(entry.unitPrice, 0);
      const hourlyWage = toNumber(employee?.hourlyWage, 0);
      const amount = qty * unitPrice;
      const wageAmount = qty * hourlyWage;
      const profit = amount - wageAmount;
      return {
        date: String(entry.date || ""),
        customerName: customer ? formatCustomerName(customer) : "Unbekannter Kunde",
        employeeName: employee ? formatEmployeeName(employee) : "Nicht zugewiesen",
        itemName: item?.name || "Unbekannte Leistung",
        qty,
        unitPrice,
        hourlyWage,
        amount,
        wageAmount,
        profit
      };
    })
    .sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.customerName).localeCompare(String(b.customerName), "de-CH"));

  const totalQty = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
  const totalAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  const totalWage = rows.reduce((sum, row) => sum + toNumber(row.wageAmount, 0), 0);
  const totalProfit = rows.reduce((sum, row) => sum + toNumber(row.profit, 0), 0);

  const body = rows.map((row) => `
    <tr>
      <td>${escapeHtml(formatWeekdayShortCH(row.date))}</td>
      <td>${escapeHtml(formatDateCH(row.date))}</td>
      <td>${escapeHtml(row.customerName)}</td>
      <td>${escapeHtml(row.employeeName)}</td>
      <td>${escapeHtml(row.itemName)}</td>
      <td class="num">${formatNumberCH(row.qty)}</td>
      <td class="num">${formatAmountCH(row.unitPrice)}</td>
      <td class="num">${formatAmountCH(row.hourlyWage)}</td>
      <td class="num">${formatAmountCH(row.amount)}</td>
      <td class="num">${formatAmountCH(row.wageAmount)}</td>
      <td class="num">${formatAmountCH(row.profit)}</td>
    </tr>
  `).join("");

  return `
    <h3>Stunden Gewinn (alle Erfassungen)</h3>
    <table class="hours-table hours-profit-table">
      <thead>
        <tr>
          <th>Tag</th>
          <th>Datum</th>
          <th>Kunde</th>
          <th>Mitarbeiter</th>
          <th>Leistung</th>
          <th class="num">Std</th>
          <th class="num">Ansatz</th>
          <th class="num">Stundenlohn</th>
          <th class="num">Umsatz</th>
          <th class="num">Lohn total</th>
          <th class="num">Gewinn</th>
        </tr>
      </thead>
      <tbody>${body || `<tr><td colspan="11"><small>Keine Daten vorhanden.</small></td></tr>`}</tbody>
      <tfoot>
        <tr>
          <th colspan="5">Total</th>
          <th class="num">${formatNumberCH(totalQty)}</th>
          <th></th>
          <th></th>
          <th class="num">${formatAmountCH(totalAmount)}</th>
          <th class="num">${formatAmountCH(totalWage)}</th>
          <th class="num">${formatAmountCH(totalProfit)}</th>
        </tr>
      </tfoot>
    </table>
  `;
}
function renderProductsProfitDetails(entries) {
  const rows = (entries || [])
    .map((entry) => {
      const customer = state.customers.find((c) => c.id === String(entry.customerId || ""));
      const employee = state.employees.find((e) => e.id === String(entry.employeeId || ""));
      const item = state.items.find((i) => i.id === String(entry.itemId || ""));
      const qty = toNumber(entry.quantity, 0);
      const unitPrice = toNumber(entry.unitPrice, 0);
      const costPrice = toNumber(entry.costPrice, toNumber(item?.costPrice, 0));
      const amount = qty * unitPrice;
      const costAmount = qty * costPrice;
      const profit = amount - costAmount;
      return {
        date: String(entry.date || ""),
        customerName: customer ? formatCustomerName(customer) : "Unbekannter Kunde",
        employeeName: employee ? formatEmployeeName(employee) : "Nicht zugewiesen",
        productName: item?.name || "Unbekanntes Produkt",
        qty,
        unitPrice,
        costPrice,
        amount,
        costAmount,
        profit
      };
    })
    .sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.customerName).localeCompare(String(b.customerName), "de-CH"));

  const totalQty = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
  const totalAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  const totalCost = rows.reduce((sum, row) => sum + toNumber(row.costAmount, 0), 0);
  const totalProfit = rows.reduce((sum, row) => sum + toNumber(row.profit, 0), 0);

  const body = rows.map((row) => `
    <tr>
      <td>${escapeHtml(formatWeekdayShortCH(row.date))}</td>
      <td>${escapeHtml(formatDateCH(row.date))}</td>
      <td>${escapeHtml(row.customerName)}</td>
      <td>${escapeHtml(row.employeeName)}</td>
      <td>${escapeHtml(row.productName)}</td>
      <td class="num">${formatNumberCH(row.qty)}</td>
      <td class="num">${formatAmountCH(row.unitPrice)}</td>
      <td class="num">${formatAmountCH(row.costPrice)}</td>
      <td class="num">${formatAmountCH(row.amount)}</td>
      <td class="num">${formatAmountCH(row.profit)}</td>
    </tr>
  `).join("");

  return `
    <h3>Produkte Gewinn (alle Erfassungen)</h3>
    <table class="hours-table products-profit-table">
      <thead>
        <tr>
          <th>Tag</th>
          <th>Datum</th>
          <th>Kunde</th>
          <th>Mitarbeiter</th>
          <th>Produkt</th>
          <th class="num">Stk</th>
          <th class="num">Preis</th>
          <th class="num">Einstand</th>
          <th class="num">Umsatz</th>
          <th class="num">Gewinn</th>
        </tr>
      </thead>
      <tbody>${body || `<tr><td colspan="10"><small>Keine Daten vorhanden.</small></td></tr>`}</tbody>
      <tfoot>
        <tr>
          <th colspan="5">Total</th>
          <th class="num">${formatNumberCH(totalQty)}</th>
          <th></th>
          <th class="num">${formatAmountCH(totalCost)}</th>
          <th class="num">${formatAmountCH(totalAmount)}</th>
          <th class="num">${formatAmountCH(totalProfit)}</th>
        </tr>
      </tfoot>
    </table>
  `;
}
function renderProductEmployeeReportHtml(month, entries) {
  const company = state.company || {};
  const rows = buildProductByEmployeeRows(entries);
  const totalQty = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
  const totalAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  const totalProfit = rows.reduce((sum, row) => sum + toNumber(row.profit, 0), 0);
  const body = rows.map((row) => `
    <tr>
      <td>${escapeHtml(row.employeeName)}</td>
      <td class="num">${row.count}</td>
      <td class="num">${formatNumberCH(row.qty)}</td>
      <td class="num">${formatAmountCH(row.amount)}</td>
      <td class="num">${formatAmountCH(row.profit)}</td>
    </tr>
  `).join("");
  return `
    <section class="invoice employee-settlement">
      <div class="invoice-top">
        <div>
          ${company.logoDataUrl ? `<img class="logo" src="${company.logoDataUrl}" alt="Logo" />` : ""}
          <p><strong>${escapeHtml(company.name || "")}</strong><br>${escapeHtml(company.street || "")}<br>${escapeHtml(company.zip || "")} ${escapeHtml(company.city || "")}</p>
        </div>
        <div class="invoice-address-box"><strong>Produkte nach Mitarbeiter</strong><br>Monat: ${escapeHtml(formatMonthCH(month))}</div>
      </div>
      <hr>
      <table class="hours-table">
        <thead><tr><th>Mitarbeiter</th><th class="num">Erfassungen</th><th class="num">Stk</th><th class="num">Betrag</th><th class="num">Gewinn</th></tr></thead>
        <tbody>${body || `<tr><td colspan="5"><small>Keine Daten vorhanden.</small></td></tr>`}</tbody>
      </table>
      <p><strong>Total Stück:</strong> ${formatNumberCH(totalQty)} | <strong>Total Betrag:</strong> ${formatCurrency(totalAmount)} | <strong>Total Gewinn:</strong> ${formatCurrency(totalProfit)}</p>
      <p><small>Hinweis: Alle Beträge ohne MwSt.</small></p>
    </section>
  `;
}

function renderProductCustomerReportHtml(month, entries) {
  const company = state.company || {};
  const rows = buildProductByCustomerRows(entries);
  const totalQty = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
  const totalAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  const totalProfit = rows.reduce((sum, row) => sum + toNumber(row.profit, 0), 0);
  const body = rows.map((row) => `
    <tr>
      <td>${escapeHtml(row.customerName)}</td>
      <td class="num">${row.count}</td>
      <td class="num">${formatNumberCH(row.qty)}</td>
      <td class="num">${formatAmountCH(row.amount)}</td>
      <td class="num">${formatAmountCH(row.profit)}</td>
    </tr>
  `).join("");
  return `
    <section class="invoice employee-settlement">
      <div class="invoice-top">
        <div>
          ${company.logoDataUrl ? `<img class="logo" src="${company.logoDataUrl}" alt="Logo" />` : ""}
          <p><strong>${escapeHtml(company.name || "")}</strong><br>${escapeHtml(company.street || "")}<br>${escapeHtml(company.zip || "")} ${escapeHtml(company.city || "")}</p>
        </div>
        <div class="invoice-address-box"><strong>Produkte nach Kunden</strong><br>Monat: ${escapeHtml(formatMonthCH(month))}</div>
      </div>
      <hr>
      <table class="hours-table">
        <thead><tr><th>Kunde</th><th class="num">Erfassungen</th><th class="num">Stk</th><th class="num">Betrag</th><th class="num">Gewinn</th></tr></thead>
        <tbody>${body || `<tr><td colspan="5"><small>Keine Daten vorhanden.</small></td></tr>`}</tbody>
      </table>
      <p><strong>Total Stück:</strong> ${formatNumberCH(totalQty)} | <strong>Total Betrag:</strong> ${formatCurrency(totalAmount)} | <strong>Total Gewinn:</strong> ${formatCurrency(totalProfit)}</p>
      <p><small>Hinweis: Alle Beträge ohne MwSt.</small></p>
    </section>
  `;
}

function buildProductEmployeeReportCsv(month, entries) {
  const rows = buildProductByEmployeeRows(entries);
  const lines = [];
  lines.push(["Monat", "\u200E" + formatMonthLongCH(month)].map(csvEscape).join(";"));
  lines.push(["Mitarbeiter", "Positionen", "Stk", "Betrag", "Gewinn"].map(csvEscape).join(";"));
  rows.forEach((row) => lines.push([row.employeeName, row.count, formatNumberCH(row.qty), formatAmountCH(row.amount), formatAmountCH(row.profit)].map(csvEscape).join(";")));
  const totalQty = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
  const totalAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  const totalProfit = rows.reduce((sum, row) => sum + toNumber(row.profit, 0), 0);
  lines.push(["Total", "", formatNumberCH(totalQty), formatAmountCH(totalAmount), formatAmountCH(totalProfit)].map(csvEscape).join(";"));
  return "\uFEFF" + lines.join("\r\n");
}

function buildProductCustomerReportCsv(month, entries) {
  const rows = buildProductByCustomerRows(entries);
  const lines = [];
  lines.push(["Monat", "\u200E" + formatMonthLongCH(month)].map(csvEscape).join(";"));
  lines.push(["Kunde", "Positionen", "Stk", "Betrag", "Gewinn"].map(csvEscape).join(";"));
  rows.forEach((row) => lines.push([row.customerName, row.count, formatNumberCH(row.qty), formatAmountCH(row.amount), formatAmountCH(row.profit)].map(csvEscape).join(";")));
  const totalQty = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
  const totalAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  const totalProfit = rows.reduce((sum, row) => sum + toNumber(row.profit, 0), 0);
  lines.push(["Total", "", formatNumberCH(totalQty), formatAmountCH(totalAmount), formatAmountCH(totalProfit)].map(csvEscape).join(";"));
  return "\uFEFF" + lines.join("\r\n");
}


function buildHoursProfitRows(entries) {
  return (entries || [])
    .map((entry) => {
      const customer = state.customers.find((c) => c.id === String(entry.customerId || ""));
      const employee = state.employees.find((e) => e.id === String(entry.employeeId || ""));
      const item = state.items.find((i) => i.id === String(entry.itemId || ""));
      const qty = toNumber(entry.quantity, 0);
      const unitPrice = toNumber(entry.unitPrice, 0);
      const hourlyWage = toNumber(employee?.hourlyWage, 0);
      const amount = qty * unitPrice;
      const wageAmount = qty * hourlyWage;
      const profit = amount - wageAmount;
      return {
        date: String(entry.date || ""),
        customerName: customer ? formatCustomerName(customer) : "Unbekannter Kunde",
        employeeName: employee ? formatEmployeeName(employee) : "Nicht zugewiesen",
        itemName: item?.name || "Unbekannte Leistung",
        qty,
        unitPrice,
        hourlyWage,
        amount,
        wageAmount,
        profit
      };
    })
    .sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.customerName).localeCompare(String(b.customerName), "de-CH"));
}

function buildProductsProfitRows(entries) {
  return (entries || [])
    .map((entry) => {
      const customer = state.customers.find((c) => c.id === String(entry.customerId || ""));
      const employee = state.employees.find((e) => e.id === String(entry.employeeId || ""));
      const item = state.items.find((i) => i.id === String(entry.itemId || ""));
      const qty = toNumber(entry.quantity, 0);
      const unitPrice = toNumber(entry.unitPrice, 0);
      const costPrice = toNumber(entry.costPrice, toNumber(item?.costPrice, 0));
      const amount = qty * unitPrice;
      const costAmount = qty * costPrice;
      const profit = amount - costAmount;
      return {
        date: String(entry.date || ""),
        customerName: customer ? formatCustomerName(customer) : "Unbekannter Kunde",
        employeeName: employee ? formatEmployeeName(employee) : "Nicht zugewiesen",
        productName: item?.name || "Unbekanntes Produkt",
        qty,
        unitPrice,
        costPrice,
        amount,
        costAmount,
        profit
      };
    })
    .sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.customerName).localeCompare(String(b.customerName), "de-CH"));
}

function buildVatRows(entries, vatRate) {
  const grouped = new Map();
  (entries || []).forEach((entry) => {
    const customerId = String(entry.customerId || "");
    const customer = state.customers.find((c) => c.id === customerId);
    const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
    if (!grouped.has(customerId)) grouped.set(customerId, { customerName, net: 0 });
    const row = grouped.get(customerId);
    row.net += toNumber(entry.unitPrice, 0) * toNumber(entry.quantity, 0);
  });
  return [...grouped.values()]
    .map((row) => {
      const vatAmount = row.net * (vatRate / 100);
      const gross = row.net + vatAmount;
      return { customerName: row.customerName, net: row.net, vatRate, vatAmount, gross };
    })
    .sort((a, b) => b.net - a.net);
}

function renderHoursProfitReportHtml(month, entries) {
  const company = state.company || {};
  return `
    <section class="invoice employee-settlement landscape-profit-report">
      <div class="invoice-top">
        <div>
          ${company.logoDataUrl ? `<img class="logo" src="${company.logoDataUrl}" alt="Logo" />` : ""}
          <p><strong>${escapeHtml(company.name || "")}</strong><br>${escapeHtml(company.street || "")}<br>${escapeHtml(company.zip || "")} ${escapeHtml(company.city || "")}</p>
        </div>
        <div class="invoice-address-box"><strong>Stunden Gewinn</strong><br>Monat: ${escapeHtml(formatMonthCH(month))}</div>
      </div>
      <hr>
      ${renderHoursProfitDetails(entries)}
      <p><small>Hinweis: Alle Beträge ohne MwSt.</small></p>
    </section>
  `;
}

function renderProductsProfitReportHtml(month, entries) {
  const company = state.company || {};
  return `
    <section class="invoice employee-settlement landscape-profit-report">
      <div class="invoice-top">
        <div>
          ${company.logoDataUrl ? `<img class="logo" src="${company.logoDataUrl}" alt="Logo" />` : ""}
          <p><strong>${escapeHtml(company.name || "")}</strong><br>${escapeHtml(company.street || "")}<br>${escapeHtml(company.zip || "")} ${escapeHtml(company.city || "")}</p>
        </div>
        <div class="invoice-address-box"><strong>Produkte Gewinn</strong><br>Monat: ${escapeHtml(formatMonthCH(month))}</div>
      </div>
      <hr>
      ${renderProductsProfitDetails(entries)}
      <p><small>Hinweis: Alle Beträge ohne MwSt.</small></p>
    </section>
  `;
}

function renderVatReportHtml(month, entries, vatRate) {
  const company = state.company || {};
  return `
    <section class="invoice employee-settlement">
      <div class="invoice-top">
        <div>
          ${company.logoDataUrl ? `<img class="logo" src="${company.logoDataUrl}" alt="Logo" />` : ""}
          <p><strong>${escapeHtml(company.name || "")}</strong><br>${escapeHtml(company.street || "")}<br>${escapeHtml(company.zip || "")} ${escapeHtml(company.city || "")}</p>
        </div>
        <div class="invoice-address-box"><strong>Mehrwertsteuer</strong><br>Monat: ${escapeHtml(formatMonthCH(month))}</div>
      </div>
      <hr>
      ${renderVatBreakdown(entries, vatRate)}
    </section>
  `;
}

function buildHoursProfitCsv(month, entries) {
  const rows = buildHoursProfitRows(entries);
  const lines = [];
  lines.push(["Monat", "\u200E" + formatMonthLongCH(month)].map(csvEscape).join(";"));
  lines.push(["Tag", "Datum", "Kunde", "Mitarbeiter", "Leistung", "Std", "Ansatz", "Stundenlohn", "Umsatz", "Lohn total", "Gewinn"].map(csvEscape).join(";"));
  rows.forEach((row) => lines.push([
    formatWeekdayShortCH(row.date),
    formatDateCH(row.date),
    row.customerName,
    row.employeeName,
    row.itemName,
    formatNumberCH(row.qty),
    formatAmountCH(row.unitPrice),
    formatAmountCH(row.hourlyWage),
    formatAmountCH(row.amount),
    formatAmountCH(row.wageAmount),
    formatAmountCH(row.profit)
  ].map(csvEscape).join(";")));
  const totalQty = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
  const totalAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  const totalWage = rows.reduce((sum, row) => sum + toNumber(row.wageAmount, 0), 0);
  const totalProfit = rows.reduce((sum, row) => sum + toNumber(row.profit, 0), 0);
  lines.push(["Total", "", "", "", "", formatNumberCH(totalQty), "", "", formatAmountCH(totalAmount), formatAmountCH(totalWage), formatAmountCH(totalProfit)].map(csvEscape).join(";"));
  return "\uFEFF" + lines.join("\r\n");
}

function buildProductsProfitCsv(month, entries) {
  const rows = buildProductsProfitRows(entries);
  const lines = [];
  lines.push(["Monat", "\u200E" + formatMonthLongCH(month)].map(csvEscape).join(";"));
  lines.push(["Tag", "Datum", "Kunde", "Mitarbeiter", "Produkt", "Stk", "Preis", "Einstand", "Umsatz", "Kosten", "Gewinn"].map(csvEscape).join(";"));
  rows.forEach((row) => lines.push([
    formatWeekdayShortCH(row.date),
    formatDateCH(row.date),
    row.customerName,
    row.employeeName,
    row.productName,
    formatNumberCH(row.qty),
    formatAmountCH(row.unitPrice),
    formatAmountCH(row.costPrice),
    formatAmountCH(row.amount),
    formatAmountCH(row.costAmount),
    formatAmountCH(row.profit)
  ].map(csvEscape).join(";")));
  const totalQty = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
  const totalAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  const totalCost = rows.reduce((sum, row) => sum + toNumber(row.costAmount, 0), 0);
  const totalProfit = rows.reduce((sum, row) => sum + toNumber(row.profit, 0), 0);
  lines.push(["Total", "", "", "", "", formatNumberCH(totalQty), "", "", formatAmountCH(totalAmount), formatAmountCH(totalCost), formatAmountCH(totalProfit)].map(csvEscape).join(";"));
  return "\uFEFF" + lines.join("\r\n");
}

function buildVatCsv(month, entries, vatRate) {
  const rows = buildVatRows(entries, vatRate);
  const lines = [];
  lines.push(["Monat", "\u200E" + formatMonthLongCH(month)].map(csvEscape).join(";"));
  lines.push(["Kunde", "Netto", "MwSt-Satz", "MwSt-Betrag", "Brutto"].map(csvEscape).join(";"));
  rows.forEach((row) => lines.push([
    row.customerName,
    formatAmountCH(row.net),
    `${formatNumberCH(row.vatRate)}%`,
    formatAmountCH(row.vatAmount),
    formatAmountCH(row.gross)
  ].map(csvEscape).join(";")));
  const totalNet = rows.reduce((sum, row) => sum + toNumber(row.net, 0), 0);
  const totalVat = rows.reduce((sum, row) => sum + toNumber(row.vatAmount, 0), 0);
  const totalGross = rows.reduce((sum, row) => sum + toNumber(row.gross, 0), 0);
  lines.push(["Total", formatAmountCH(totalNet), `${formatNumberCH(vatRate)}%`, formatAmountCH(totalVat), formatAmountCH(totalGross)].map(csvEscape).join(";"));
  return "\uFEFF" + lines.join("\r\n");
}
function summarizeEntries(entries, keyField, resolveLabel) {
  const grouped = new Map();
  entries.forEach((entry) => {
    const key = String(entry?.[keyField] || "");
    const label = resolveLabel(entry);
    if (!grouped.has(key)) {
      grouped.set(key, { label, count: 0, qty: 0, amount: 0 });
    }
    const row = grouped.get(key);
    row.count += 1;
    row.qty += toNumber(entry.quantity, 0);
    row.amount += toNumber(entry.unitPrice, 0) * toNumber(entry.quantity, 0);
  });
  return [...grouped.values()].sort((a, b) => b.qty - a.qty || b.amount - a.amount);
}

function renderHoursTable(title, rows) {
  const body = rows
    .map(
      (row) => `
        <tr>
          <td>${escapeHtml(row.label || row.customerName || row.employeeName || "Unbekannt")}</td>
          <td class="num">${row.count}</td>
          <td class="num">${formatNumberCH(row.qty)}</td>
          <td class="num">${formatCurrency(row.amount)}</td>
        </tr>
      `
    )
    .join("");

  return `
    <h3>${escapeHtml(title)}</h3>
    <table class="hours-table">
      <thead>
        <tr>
          <th>Name</th>
          <th class="num">Erfassungen</th>
          <th class="num">Menge</th>
          <th class="num">Betrag</th>
        </tr>
      </thead>
      <tbody>${body}</tbody>
    </table>
  `;
}

function renderCustomerSummaryWithDetails(entries, searchTerm = "") {
  const grouped = new Map();
  entries.forEach((entry) => {
    const customerId = String(entry.customerId || "");
    const customer = state.customers.find((c) => c.id === customerId);
    const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
    if (!grouped.has(customerId)) {
      grouped.set(customerId, {
        customerId,
        customerName,
        count: 0,
        qty: 0,
        amount: 0,
        details: new Map()
      });
    }
    const row = grouped.get(customerId);
    const qty = toNumber(entry.quantity, 0);
    const amount = toNumber(entry.unitPrice, 0) * qty;
    row.count += 1;
    row.qty += qty;
    row.amount += amount;

    const detailKey = `${entry.date || ""}|${entry.employeeId || ""}`;
    const existingDetail = row.details.get(detailKey) || {
      date: String(entry.date || ""),
      employeeId: String(entry.employeeId || ""),
      qty: 0,
      amount: 0
    };
    existingDetail.qty += qty;
    existingDetail.amount += amount;
    row.details.set(detailKey, existingDetail);
  });

  const rows = [...grouped.values()].sort((a, b) => b.qty - a.qty || b.amount - a.amount);
  const visibleRows = rows.filter((row) => matchesSearch(row.customerName, searchTerm));
  const body = visibleRows
    .map(
      (row) => `
        <tr>
          <td>${escapeHtml(row.customerName)}</td>
          <td class="num">${row.count}</td>
          <td class="num">${formatNumberCH(row.qty)}</td>
          <td class="num">${formatCurrency(row.amount)}</td>
          <td class="num hours-action-cell">
            <div class="hours-actions">
              <button type="button" data-action="print-customer" data-id="${escapeHtml(row.customerId || "__unknown_customer__")}">Drucken</button>
              <button type="button" data-action="save-customer" data-id="${escapeHtml(row.customerId || "__unknown_customer__")}">Speichern</button>
            </div>
          </td>
        </tr>
      `
    )
    .join("");

  const detailsHtml = visibleRows
    .map((row) => {
      const detailRows = [...row.details.values()]
        .sort((a, b) => String(a.date).localeCompare(String(b.date)))
        .map((detail) => {
          const employee = state.employees.find((e) => e.id === detail.employeeId);
          const employeeName = employee ? formatEmployeeName(employee) : "Nicht zugewiesen";
          return `
            <tr>
              <td>${escapeHtml(formatWeekdayLongCH(detail.date))}</td>
              <td>${escapeHtml(formatDateCH(detail.date))}</td>
              <td>${escapeHtml(employeeName)}</td>
              <td class="num">${formatNumberCH(detail.qty)}</td>
              <td class="num">${formatCurrency(detail.amount)}</td>
            </tr>
          `;
        })
        .join("");
      return `
        <section class="hours-detail-block">
          <h4>${escapeHtml(row.customerName)} <small>(${formatNumberCH(row.qty)} Std. | ${formatCurrency(row.amount)} | Gewinn ${formatCurrency(row.profit)})</small></h4>
          <table class="hours-table">
            <thead>
              <tr>
                <th>Wochentag</th>
                <th>Datum</th>
                <th>Mitarbeiter</th>
                <th class="num">Stunden</th>
                <th class="num">Betrag</th>
              </tr>
            </thead>
            <tbody>${detailRows}</tbody>
          </table>
        </section>
      `;
    })
    .join("");
  return `
    <h3>Nach Kunde</h3>
    <div class="hours-actions">
      <button type="button" data-action="preview-customer-hours-pdf">Drucken</button>
      <button type="button" data-action="save-customer-hours-pdf">Speichern</button>
      <button type="button" data-action="export-customer-hours-csv">CSV (Excel)</button>
    </div>
    <label>Kunde suchen
      <input id="hoursCustomerFilter" type="search" placeholder="Name..." value="${escapeHtml(searchTerm)}" autocomplete="off" />
    </label>
    <table class="hours-table">
      <thead>
        <tr>
          <th>Name</th>
          <th class="num">Erfassungen</th>
          <th class="num">Stunden</th>
          <th class="num">Betrag</th>
          <th class="num">Abrechnung</th>
        </tr>
      </thead>
      <tbody>${body || `<tr><td colspan="5"><small>Keine Kunden passend zur Suche gefunden.</small></td></tr>`}</tbody>
    </table>
    <div class="hours-detail-wrap">
      <h3>Details je Kunde</h3>
      ${detailsHtml || "<small>Keine Detaildaten passend zur Suche gefunden.</small>"}
    </div>
  `;
}

function buildCustomerHoursReportRows(entries) {
  const grouped = new Map();
  entries.forEach((entry) => {
    const customerId = String(entry.customerId || "");
    const customer = state.customers.find((c) => c.id === customerId);
    const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
    if (!grouped.has(customerId)) {
      grouped.set(customerId, {
        customerId,
        customerName,
        qty: 0,
        amount: 0,
        details: new Map()
      });
    }
    const row = grouped.get(customerId);
    const qty = toNumber(entry.quantity, 0);
    const amount = toNumber(entry.unitPrice, 0) * qty;
    row.qty += qty;
    row.amount += amount;

    const detailKey = `${entry.date || ""}|${entry.employeeId || ""}`;
    const detail = row.details.get(detailKey) || {
      date: String(entry.date || ""),
      employeeId: String(entry.employeeId || ""),
      qty: 0,
      amount: 0
    };
    detail.qty += qty;
    detail.amount += amount;
    row.details.set(detailKey, detail);
  });
  return [...grouped.values()].sort((a, b) => String(a.customerName).localeCompare(String(b.customerName)));
}

function renderCustomerHoursReportHtml(month, entries) {
  const company = state.company || {};
  const rows = buildCustomerHoursReportRows(entries);
  const totalHours = rows.reduce((sum, row) => sum + row.qty, 0);
  const totalAmount = rows.reduce((sum, row) => sum + row.amount, 0);

  const sectionsHtml = rows
    .map((row) => {
      const detailRows = [...row.details.values()]
        .sort((a, b) => String(a.date).localeCompare(String(b.date)))
        .map((detail) => {
          const employee = state.employees.find((e) => e.id === detail.employeeId);
          const employeeName = employee ? formatEmployeeName(employee) : "Nicht zugewiesen";
          return `
            <tr>
              <td>${escapeHtml(formatWeekdayShortCH(detail.date))}</td>
              <td>${escapeHtml(formatDateCH(detail.date))}</td>
              <td>${escapeHtml(employeeName)}</td>
              <td class="num">${formatNumberCH(detail.qty)}</td>
              <td class="num">${formatAmountCH(detail.amount)}</td>
            </tr>
          `;
        })
        .join("");
      return `
        <section class="invoice-text-box">
          <h3>${escapeHtml(row.customerName)} <small>(${formatNumberCH(row.qty)} Std. | ${formatCurrency(row.amount)})</small></h3>
          <table class="invoice-lines hours-table">
            <colgroup>
              <col style="width: 6%;" />
              <col style="width: 14%;" />
              <col style="width: 44%;" />
              <col style="width: 12%;" />
              <col style="width: 24%;" />
            </colgroup>
            <thead>
              <tr>
                <th>Tag</th>
                <th>Datum</th>
                <th>Mitarbeiter</th>
                <th class="num">Stunden</th>
                <th class="num">Total</th>
              </tr>
            </thead>
            <tbody>${detailRows || `<tr><td colspan="5"><small>Keine Details vorhanden.</small></td></tr>`}</tbody>
          </table>
        </section>
      `;
    })
    .join("");

  return `
    <section class="invoice employee-settlement">
      <div class="invoice-top">
        <div>
          ${company.logoDataUrl ? `<img class="logo" src="${company.logoDataUrl}" alt="Logo" />` : ""}
          <p>
            <strong>${escapeHtml(company.name || "")}</strong><br>
            ${escapeHtml(company.street || "")}<br>
            ${escapeHtml(company.zip || "")} ${escapeHtml(company.city || "")}
          </p>
        </div>
        <div class="invoice-address-box">
          <strong>Stundenauswertung Kunden</strong><br>
          Monat: ${escapeHtml(formatMonthCH(month))}
        </div>
      </div>
      <hr>
      <p><strong>Monat:</strong> ${escapeHtml(formatMonthCH(month))}</p>
      <p><strong>Währung:</strong> ${escapeHtml(normalizeCurrencyCode(company.currency))}</p>
      <p><strong>Total Stunden:</strong> ${formatNumberCH(totalHours)} | <strong>Total Betrag:</strong> ${formatCurrency(totalAmount)}</p>
        <p><small>Hinweis: Alle Beträge ohne MwSt.</small></p>
      ${sectionsHtml || "<small>Keine Stundenpositionen vorhanden.</small>"}
    </section>
  `;
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (!/[";\r\n]/.test(text)) return text;
  return `"${text.replace(/"/g, '""')}"`;
}

function buildCustomerHoursCsv(month, entries) {
  const rows = buildCustomerHoursReportRows(entries);
  const lines = [];
  lines.push(["Monat", "\u200E" + formatMonthLongCH(month)].map(csvEscape).join(";"));
  lines.push(["Kunde", "Wochentag", "Datum", "Mitarbeiter", "Stunden", "Betrag CHF"].map(csvEscape).join(";"));
  const totalMonthHours = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
  const totalMonthAmount = rows.reduce((sum, row) => sum + toNumber(row.amount, 0), 0);
  rows.forEach((row) => {
const details = [...row.details.values()].sort((a, b) => String(a.date).localeCompare(String(b.date)));
    details.forEach((detail) => {
      const employee = state.employees.find((e) => e.id === detail.employeeId);
      const employeeName = employee ? formatEmployeeName(employee) : "Nicht zugewiesen";
      lines.push([
        row.customerName,
        formatWeekdayLongCH(detail.date),
        formatDateCH(detail.date),
        employeeName,
        formatNumberCH(detail.qty),
        formatAmountCH(detail.amount)
      ].map(csvEscape).join(";"));
    });
    lines.push([
      `${row.customerName} Total`,
      "",
      "",
      "",
      formatNumberCH(row.qty),
      formatAmountCH(row.amount)
    ].map(csvEscape).join(";"));
  });
  lines.push(["Total Monat", "", "", "", formatNumberCH(totalMonthHours), formatAmountCH(totalMonthAmount)].map(csvEscape).join(";"));
  return "\uFEFF" + lines.join("\r\n");
}
function renderEmployeeSummaryTable(entries, searchTerm = "") {
  const grouped = new Map();
  entries.forEach((entry) => {
    const employeeId = String(entry.employeeId || "");
    const employee = state.employees.find((e) => e.id === employeeId);
    const label = employee ? formatEmployeeName(employee) : "Nicht zugewiesen";
    const wage = toNumber(employee?.hourlyWage, 0);
    if (!grouped.has(employeeId)) {
      grouped.set(employeeId, { employeeId, label, count: 0, qty: 0, wage, totalWage: 0 });
    }
    const row = grouped.get(employeeId);
    const qty = toNumber(entry.quantity, 0);
    row.count += 1;
    row.qty += qty;
    row.totalWage += qty * row.wage;
  });

  const rows = [...grouped.values()].sort((a, b) => b.qty - a.qty || b.totalWage - a.totalWage);
  const visibleRows = rows.filter((row) => matchesSearch(row.label, searchTerm));
  const body = visibleRows
    .map(
      (row) => `
        <tr>
          <td>${escapeHtml(row.label)}</td>
          <td class="num">${row.count}</td>
          <td class="num">${formatNumberCH(row.qty)}</td>
          <td class="num">${formatCurrency(row.wage)}</td>
          <td class="num">${formatCurrency(row.totalWage)}</td>
          <td class="num hours-action-cell">
            <div class="hours-actions">
              <button type="button" data-action="print-employee" data-id="${escapeHtml(row.employeeId || "__unassigned__")}">Drucken</button>
              <button type="button" data-action="save-employee" data-id="${escapeHtml(row.employeeId || "__unassigned__")}">Speichern</button>
            </div>
          </td>
        </tr>
      `
    )
    .join("");

  return `
    <h3>Nach Mitarbeiter</h3>
    <div class="hours-actions">
      <button type="button" data-action="preview-employee-hours-pdf">Drucken</button>
      <button type="button" data-action="save-employee-hours-pdf">Speichern</button>
      <button type="button" data-action="export-employee-hours-csv">CSV (Excel)</button>
    </div>
    <label>Mitarbeiter suchen
      <input id="hoursEmployeeFilter" type="search" placeholder="Name..." value="${escapeHtml(searchTerm)}" autocomplete="off" />
    </label>
    <table class="hours-table">
      <thead>
        <tr>
          <th>Name</th>
          <th class="num">Erfassungen</th>
          <th class="num">Stunden</th>
          <th class="num">Vereinbarter Stundenlohn</th>
          <th class="num">Total Lohn vor Abzügen</th>
          <th class="num">Abrechnung</th>
        </tr>
      </thead>
      <tbody>${body || `<tr><td colspan="6"><small>Keine Mitarbeiter passend zur Suche gefunden.</small></td></tr>`}</tbody>
    </table>
  `;
}

function renderEmployeeHoursDetails(entries, searchTerm = "") {
  const byEmployee = new Map();
  entries.forEach((entry) => {
    const employeeId = String(entry.employeeId || "");
    const customerId = String(entry.customerId || "");
    const date = String(entry.date || "");
    const key = `${employeeId}|${customerId}|${date}`;
    if (!byEmployee.has(employeeId)) byEmployee.set(employeeId, new Map());
    const rowMap = byEmployee.get(employeeId);
    const row = rowMap.get(key) || { employeeId, customerId, date, qty: 0 };
    row.qty += toNumber(entry.quantity, 0);
    rowMap.set(key, row);
  });

  const employeeSections = [...byEmployee.entries()]
    .map(([employeeId, rowMap]) => {
      const employee = state.employees.find((e) => e.id === employeeId);
      const employeeName = employee ? formatEmployeeName(employee) : "Nicht zugewiesen";
      if (!matchesSearch(employeeName, searchTerm)) return "";
      const hourlyWage = toNumber(employee?.hourlyWage, 0);
      const rows = [...rowMap.values()]
        .sort((a, b) => String(a.date).localeCompare(String(b.date)))
        .map((row) => {
          const customer = state.customers.find((c) => c.id === row.customerId);
          const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
          const wageTotal = row.qty * hourlyWage;
          return `
            <tr>
              <td>${escapeHtml(formatWeekdayLongCH(row.date))}</td>
              <td>${escapeHtml(formatDateCH(row.date))}</td>
              <td>${escapeHtml(customerName)}</td>
              <td class="num">${formatNumberCH(row.qty)}</td>
              <td class="num">${formatCurrency(hourlyWage)}</td>
              <td class="num">${formatCurrency(wageTotal)}</td>
            </tr>
          `;
        })
        .join("");

      return `
        <section class="hours-detail-block">
          <h4>${escapeHtml(employeeName)}</h4>
          <table class="hours-table">
            <thead>
              <tr>
                <th>Wochentag</th>
                <th>Datum</th>
                <th>Kunde</th>
                <th class="num">Stunden</th>
                <th class="num">Vereinbarter Stundenlohn</th>
                <th class="num">Total Lohn vor Abzügen</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </section>
      `;
    })
    .join("");
  return `
    <div class="hours-detail-wrap">
      <h3>Details je Mitarbeiter</h3>
      ${employeeSections || "<small>Keine Detaildaten passend zur Suche gefunden.</small>"}
    </div>
  `;
}

function buildEmployeeHoursReportRows(entries) {
  const grouped = new Map();
  entries.forEach((entry) => {
    const employeeId = String(entry.employeeId || "");
    const employee = state.employees.find((e) => e.id === employeeId);
    const employeeName = employee ? formatEmployeeName(employee) : "Nicht zugewiesen";
    const hourlyWage = toNumber(employee?.hourlyWage, 0);
    if (!grouped.has(employeeId)) {
      grouped.set(employeeId, {
        employeeId,
        employeeName,
        hourlyWage,
        qty: 0,
        totalWage: 0,
        details: new Map()
      });
    }
    const row = grouped.get(employeeId);
    const qty = toNumber(entry.quantity, 0);
    row.qty += qty;
    row.totalWage += qty * row.hourlyWage;

    const detailKey = `${entry.date || ""}|${entry.customerId || ""}`;
    const detail = row.details.get(detailKey) || {
      date: String(entry.date || ""),
      customerId: String(entry.customerId || ""),
      qty: 0
    };
    detail.qty += qty;
    row.details.set(detailKey, detail);
  });
  return [...grouped.values()].sort((a, b) => String(a.employeeName).localeCompare(String(b.employeeName)));
}

function renderEmployeeHoursReportHtml(month, entries) {
  const company = state.company || {};
  const rows = buildEmployeeHoursReportRows(entries);
  const totalHours = rows.reduce((sum, row) => sum + row.qty, 0);
  const totalWage = rows.reduce((sum, row) => sum + row.totalWage, 0);

  const sectionsHtml = rows
    .map((row) => {
      const detailRows = [...row.details.values()]
        .sort((a, b) => String(a.date).localeCompare(String(b.date)))
        .map((detail) => {
          const customer = state.customers.find((c) => c.id === detail.customerId);
          const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
          const detailTotal = detail.qty * row.hourlyWage;
          return `
            <tr>
              <td>${escapeHtml(formatWeekdayShortCH(detail.date))}</td>
              <td>${escapeHtml(formatDateCH(detail.date))}</td>
              <td>${escapeHtml(customerName)}</td>
              <td class="num">${formatNumberCH(detail.qty)}</td>
              <td class="num">${formatAmountCH(row.hourlyWage)}</td>
              <td class="num">${formatAmountCH(detailTotal)}</td>
            </tr>
          `;
        })
        .join("");
      return `
        <section class="invoice-text-box">
          <h3>${escapeHtml(row.employeeName)} <small>(${formatNumberCH(row.qty)} Std. | ${formatCurrency(row.totalWage)})</small></h3>
          <table class="invoice-lines hours-table">
            <colgroup>
              <col style="width: 6%;" />
              <col style="width: 14%;" />
              <col style="width: 34%;" />
              <col style="width: 12%;" />
              <col style="width: 17%;" />
              <col style="width: 17%;" />
            </colgroup>
            <thead>
              <tr>
                <th>Tag</th>
                <th>Datum</th>
                <th>Kunde</th>
                <th class="num">Stunden</th>
                <th class="num">Stundenlohn</th>
                <th class="num">Total Lohn</th>
              </tr>
            </thead>
            <tbody>${detailRows || `<tr><td colspan="6"><small>Keine Details vorhanden.</small></td></tr>`}</tbody>
          </table>
        </section>
      `;
    })
    .join("");

  return `
    <section class="invoice employee-settlement">
      <div class="invoice-top">
        <div>
          ${company.logoDataUrl ? `<img class="logo" src="${company.logoDataUrl}" alt="Logo" />` : ""}
          <p>
            <strong>${escapeHtml(company.name || "")}</strong><br>
            ${escapeHtml(company.street || "")}<br>
            ${escapeHtml(company.zip || "")} ${escapeHtml(company.city || "")}
          </p>
        </div>
        <div class="invoice-address-box">
          <strong>Stundenauswertung Mitarbeiter</strong><br>
          Monat: ${escapeHtml(formatMonthCH(month))}
        </div>
      </div>
      <hr>
      <p><strong>Monat:</strong> ${escapeHtml(formatMonthCH(month))}</p>
      <p><strong>Währung:</strong> ${escapeHtml(normalizeCurrencyCode(company.currency))}</p>
      <p><strong>Total Stunden:</strong> ${formatNumberCH(totalHours)} | <strong>Total Lohn:</strong> ${formatCurrency(totalWage)}</p>
        <p><small>Hinweis: Alle Beträge ohne MwSt.</small></p>
      ${sectionsHtml || "<small>Keine Stundenpositionen vorhanden.</small>"}
    </section>
  `;
}

function buildEmployeeHoursCsv(month, entries) {
  const rows = buildEmployeeHoursReportRows(entries);
  const lines = [];
  lines.push(["Monat", "\u200E" + formatMonthLongCH(month)].map(csvEscape).join(";"));
  lines.push(["Mitarbeiter", "Wochentag", "Datum", "Kunde", "Stunden", "Stundenlohn", "Total Lohn"].map(csvEscape).join(";"));
  const totalMonthHours = rows.reduce((sum, row) => sum + toNumber(row.qty, 0), 0);
  const totalMonthWage = rows.reduce((sum, row) => sum + toNumber(row.totalWage, 0), 0);
  rows.forEach((row) => {
const details = [...row.details.values()].sort((a, b) => String(a.date).localeCompare(String(b.date)));
    details.forEach((detail) => {
      const customer = state.customers.find((c) => c.id === detail.customerId);
      const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
      const detailTotal = detail.qty * row.hourlyWage;
      lines.push([
        row.employeeName,
        formatWeekdayLongCH(detail.date),
        formatDateCH(detail.date),
        customerName,
        formatNumberCH(detail.qty),
        formatAmountCH(row.hourlyWage),
        formatAmountCH(detailTotal)
      ].map(csvEscape).join(";"));
    });
    lines.push([
      `${row.employeeName} Total`,
      "",
      "",
      "",
      formatNumberCH(row.qty),
      "",
      formatAmountCH(row.totalWage)
    ].map(csvEscape).join(";"));
  });
  lines.push(["Total Monat", "", "", "", formatNumberCH(totalMonthHours), "", formatAmountCH(totalMonthWage)].map(csvEscape).join(";"));
  return "\uFEFF" + lines.join("\r\n");
}
function renderEmployeeSettlementHtml(employeeId, month, entries) {
  const company = state.company || {};
  const employee = state.employees.find((e) => e.id === employeeId) || {};
  const employeeName = formatEmployeeName(employee);
  const hourlyWage = toNumber(employee.hourlyWage, 0);
  const employeeIban = formatIbanDisplay(employee.iban);

  const grouped = new Map();
  entries.forEach((entry) => {
    const customerId = String(entry.customerId || "");
    const date = String(entry.date || "");
    const key = `${date}|${customerId}`;
    const current = grouped.get(key) || { date, customerId, qty: 0 };
    current.qty += toNumber(entry.quantity, 0);
    grouped.set(key, current);
  });

  const rows = [...grouped.values()].sort((a, b) => String(a.date).localeCompare(String(b.date)));
  const totalHours = rows.reduce((sum, row) => sum + row.qty, 0);
  const totalWage = totalHours * hourlyWage;

  const rowsHtml = rows
    .map((row) => {
      const customer = state.customers.find((c) => c.id === row.customerId);
      const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
      const rowTotal = row.qty * hourlyWage;
      return `
        <tr>
          <td>${escapeHtml(formatWeekdayShortCH(row.date))}</td>
          <td>${escapeHtml(formatDateCH(row.date))}</td>
          <td>${escapeHtml(customerName)}</td>
          <td class="num">${formatNumberCH(row.qty)}</td>
          <td class="num">${formatCurrencyAligned(hourlyWage)}</td>
          <td class="num">${formatCurrencyAligned(rowTotal)}</td>
        </tr>
      `;
    })
    .join("");

  const createdAt = new Date().toISOString().slice(0, 10);

  return `
    <section class="invoice employee-settlement">
      <div class="invoice-top">
        <div>
          ${company.logoDataUrl ? `<img class="logo" src="${company.logoDataUrl}" alt="Logo" />` : ""}
          <p>
            <strong>${escapeHtml(company.name || "")}</strong><br>
            ${escapeHtml(company.street || "")}<br>
            ${escapeHtml(company.zip || "")} ${escapeHtml(company.city || "")}<br>
            ${escapeHtml(company.email || "")}
          </p>
        </div>
      </div>

      <div class="address-window">
        ${escapeHtml(employeeName)}<br>
        ${escapeHtml(employee.street || "")}<br>
        ${escapeHtml(employee.zip || "")} ${escapeHtml(employee.city || "")}
      </div>

      <h2>Mitarbeiter-Abrechnung ${escapeHtml(formatMonthCH(month))}</h2>
      <p class="meta">
        Erstellt am: ${escapeHtml(formatDateCH(createdAt))}
        | Zeitraum: ${escapeHtml(formatMonthCH(month))}
      </p>

      <table>
        <thead>
          <tr>
            <th>Tag</th>
            <th>Datum</th>
            <th>Kunde</th>
            <th class="num">Stunden</th>
            <th class="num">Stundenlohn</th>
            <th class="num">Total</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>

      <div class="totals">
        <p><span>Stunden total</span><span>${formatNumberCH(totalHours)}</span></p>
        <p><span>Vereinbarter Stundenlohn</span><span>${formatCurrency(hourlyWage)}</span></p>
        <p class="grand"><span>Total Lohn vor Abzügen</span><span>${formatCurrency(totalWage)}</span></p>
      </div>

      <section class="invoice-text-block">
        <p>Bankverbindung von ${escapeHtml(employeeName)}: ${employeeIban ? escapeHtml(employeeIban) : "Keine IBAN hinterlegt"}</p>
      </section>
    </section>
  `;
}

function getInvoicesOlderThanMonths(months) {
  const safeMonths = Math.min(24, Math.max(1, Math.round(toNumber(months, 12))));
  const now = new Date();
  const cutoff = new Date(now.getFullYear(), now.getMonth(), 1);
  cutoff.setMonth(cutoff.getMonth() - safeMonths);
  return state.invoices
    .filter((invoice) => {
      const periodDate = getInvoicePeriodDate(invoice);
      if (!periodDate) return false;
      return periodDate < cutoff;
    })
    .slice()
    .sort((a, b) => String(a.month || "").localeCompare(String(b.month || "")));
}

function pruneDataOlderThanMonths(months) {
  const removedInvoices = getInvoicesOlderThanMonths(months);
  const removedEntries = getEntriesOlderThanMonths(months);

  if (removedInvoices.length) {
    const invoiceIds = new Set(removedInvoices.map((invoice) => invoice.id));
    state.invoices = state.invoices.filter((invoice) => !invoiceIds.has(invoice.id));
  }
  if (removedEntries.length) {
    const entryIds = new Set(removedEntries.map((entry) => entry.id));
    state.entries = state.entries.filter((entry) => !entryIds.has(entry.id));
  }

  return { removedInvoices, removedEntries };
}

function getEntriesOlderThanMonths(months) {
  const safeMonths = Math.min(24, Math.max(1, Math.round(toNumber(months, 12))));
  const now = new Date();
  const cutoff = new Date(now.getFullYear(), now.getMonth(), 1);
  cutoff.setMonth(cutoff.getMonth() - safeMonths);
  return state.entries
    .filter((entry) => {
      const periodDate = getEntryPeriodDate(entry);
      if (!periodDate) return false;
      return periodDate < cutoff;
    })
    .slice()
    .sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));
}

function getInvoicePeriodDate(invoice) {
  const monthRaw = String(invoice?.month || "").trim();
  if (/^\d{4}-\d{2}$/.test(monthRaw)) {
    const [yearText, monthText] = monthRaw.split("-");
    const year = Number(yearText);
    const month = Number(monthText);
    if (Number.isFinite(year) && Number.isFinite(month) && month >= 1 && month <= 12) {
      return new Date(year, month - 1, 1);
    }
  }
  const createdRaw = String(invoice?.createdAt || "").trim();
  if (createdRaw) {
    const createdDate = new Date(createdRaw);
    if (!Number.isNaN(createdDate.getTime())) {
      return new Date(createdDate.getFullYear(), createdDate.getMonth(), 1);
    }
  }
  return null;
}

function getEntryPeriodDate(entry) {
  const dateRaw = String(entry?.date || "").trim();
  if (!dateRaw) return null;
  const parsed = new Date(dateRaw);
  if (Number.isNaN(parsed.getTime())) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), 1);
}

function formatInvoiceCleanupLine(invoice) {
  const customer = state.customers.find((c) => c.id === invoice.customerId);
  const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
  const month = invoice.month ? formatMonthCH(invoice.month) : "-";
  return `${invoice.invoiceNo || "(ohne Nr.)"} | ${customerName} | ${month}`;
}

function renderCleanupResult(invoices, entries) {
  const safeInvoices = Array.isArray(invoices) ? invoices : [];
  const safeEntries = Array.isArray(entries) ? entries : [];
  if (!safeInvoices.length && !safeEntries.length) {
    cleanupResult.hidden = true;
    cleanupResultList.innerHTML = "";
    return;
  }
  cleanupResult.hidden = false;
  const invoicesHtml = safeInvoices.length
    ? `
      <section>
        <strong>Rechnungen (${safeInvoices.length})</strong>
        <div class="cleanup-result-list">
          ${safeInvoices
            .map((invoice) => {
              const customer = state.customers.find((c) => c.id === invoice.customerId);
              const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
              const month = invoice.month ? formatMonthCH(invoice.month) : "-";
              const total = formatCurrency(toNumber(invoice.grandTotal, 0));
              return `
                <article class="cleanup-result-row">
                  <strong>${escapeHtml(invoice.invoiceNo || "(ohne Nr.)")}</strong><br>
                  Kunde: ${escapeHtml(customerName)}<br>
                  Monat: ${escapeHtml(month)} | Total: ${escapeHtml(total)}
                </article>
              `;
            })
            .join("")}
        </div>
      </section>
    `
    : "";

  const entriesHtml = safeEntries.length
    ? `
      <section>
        <strong>Erfassungen (${safeEntries.length})</strong>
        <div class="cleanup-result-list">
          ${safeEntries
            .map((entry) => {
              const customer = state.customers.find((c) => c.id === entry.customerId);
              const item = state.items.find((i) => i.id === entry.itemId);
              const customerName = customer ? formatCustomerName(customer) : "Unbekannter Kunde";
              const itemName = item?.name || "Unbekannte Leistung";
              const quantity = toNumber(entry.quantity, 0);
              const unit = item?.unit || "";
              return `
                <article class="cleanup-result-row">
                  <strong>${escapeHtml(formatDateCH(entry.date || ""))}</strong><br>
                  Kunde: ${escapeHtml(customerName)}<br>
                  Leistung: ${escapeHtml(itemName)} | Menge: ${escapeHtml(`${quantity} ${unit}`.trim())}
                </article>
              `;
            })
            .join("")}
        </div>
      </section>
    `
    : "";

  cleanupResultList.innerHTML = `${invoicesHtml}${entriesHtml}`;
}
function buildInvoice(customerId, month, paymentSlipType = "swiss_qr") {
  const customer = state.customers.find((c) => c.id === customerId);
  if (!customer) return null;

  const monthEntries = state.entries.filter((e) => e.customerId === customerId && String(e.date || "").startsWith(month));
  if (!monthEntries.length) return null;

  const rows = monthEntries
    .slice()
    .sort((a, b) => String(a.date).localeCompare(String(b.date)))
    .map((e) => {
      const item = state.items.find((i) => i.id === e.itemId);
      const unit = item?.unit || "";
      const label = item?.name || "Unbekannter Eintrag";
      const unitPrice = e.unitPrice != null ? toNumber(e.unitPrice, 0) : toNumber(item?.price, 0);
      const quantity = toNumber(e.quantity, 0);
      const total = unitPrice * quantity;
      return {
        date: e.date,
        label,
        unit,
        quantity,
        unitPrice,
        total,
        note: e.note || ""
      };
    });

  const subtotal = rows.reduce((acc, r) => acc + r.total, 0);
  const vatRate = toNumber(state.company.vatRate, 0);
  const vatAmount = subtotal * (vatRate / 100);
  const grandTotal = subtotal + vatAmount;
  const createdAt = new Date().toISOString().slice(0, 10);
  const dueDate = addDays(createdAt, Math.max(0, Math.round(toNumber(state.company.paymentTermDays, 30))));
  const invoiceNo = buildInvoiceNo(customer, month);
  const qrPayload = buildSwissQrPayload({
    iban: state.company.iban,
    currency: normalizeQrCurrencyCode(state.company.currency),
    creditorName: state.company.name,
    creditorStreet: state.company.street,
    creditorZip: state.company.zip,
    creditorCity: state.company.city,
    debtorName: formatCustomerName(customer),
    debtorStreet: customer.street || "",
    debtorZip: customer.zip || "",
    debtorCity: customer.city || "",
    amount: grandTotal,
    message: `Rechnung ${invoiceNo}`
  });

  return {
    id: generateId(),
    invoiceNo,
    month,
    customerId,
    rows,
    subtotal,
    vatRate,
    vatAmount,
    grandTotal,
    createdAt,
    dueDate,
    paid: false,
    paidAt: "",
    paidMethod: "",
    paidImportFile: "",
    paidImportedAt: "",
    paidStatusSetAt: "",
    sent: false,
    sentAt: "",
    paymentSlipType: normalizePaymentSlipType(paymentSlipType),
    qrPayload
  };
}

function renderInvoiceHtml(invoice) {
  const c = state.customers.find((x) => x.id === invoice.customerId) || {};
  const company = state.company || {};
  const createdAt = invoice.createdAt || new Date().toISOString().slice(0, 10);
  const paymentSlipType = normalizePaymentSlipType(invoice.paymentSlipType);
  const currencyCode = normalizeCurrencyCode(company.currency);

  const rowsHtml = (invoice.rows || [])
    .map(
      (r) => `
          <tr>
            <td>${escapeHtml(formatWeekdayShortCH(r.date))}</td>
            <td>${escapeHtml(formatDateCH(r.date))}</td>
            <td>${escapeHtml(r.label)}${r.note ? `<br><small>${escapeHtml(r.note)}</small>` : ""}</td>
            <td class="num">${r.quantity}</td>
            <td>${escapeHtml(r.unit)}</td>
            <td class="num">${formatAmountCH(r.unitPrice)}</td>
            <td class="nowrap">${escapeHtml(currencyCode)}</td>
            <td class="num">${formatAmountCH(r.total)}</td>
          </tr>
        `
    )
    .join("");

  const customerPersonName = `${c.firstName || ""} ${c.lastName || ""}`.trim();
  const customerAddressNameLines = [];
  if (c.company) customerAddressNameLines.push(escapeHtml(c.company));
  if (customerPersonName) customerAddressNameLines.push(escapeHtml(customerPersonName));
  if (!customerAddressNameLines.length) customerAddressNameLines.push(escapeHtml(formatCustomerName(c)));

  let paymentSlipSection = "";
  if (paymentSlipType === "swiss_qr") {
    const swissQrBillSvg = renderSwissQrBillSvg(invoice, c, company);
    paymentSlipSection = `<section class="qr-page">${swissQrBillSvg}</section>`;
  } else if (paymentSlipType === "appendix") {
    const appendixText = String(company.appendixText || "").trim();
    paymentSlipSection = `
      <section class="qr-page">
        <div class="invoice-text-block">
          <p>${escapeHtml(appendixText || "Individuelles Beiblatt")}</p>
        </div>
      </section>
    `;
  }

  return `
    <section class="invoice">
      <div class="invoice-top">
        <div>
          ${company.logoDataUrl ? `<img class="logo" src="${company.logoDataUrl}" alt="Logo" />` : ""}
          <p>
            <strong>${escapeHtml(company.name || "")}</strong><br>
            ${escapeHtml(company.street || "")}<br>
            ${escapeHtml(company.zip || "")} ${escapeHtml(company.city || "")}<br>
            ${escapeHtml(company.email || "")}
            ${company.vatNo ? `<br>MwSt-Nr: ${escapeHtml(company.vatNo)}` : ""}
          </p>
        </div>
      </div>

      <div class="address-window">
        ${customerAddressNameLines.join("<br>")}<br>
        ${escapeHtml(c.street || "")}<br>
        ${escapeHtml(c.zip || "")} ${escapeHtml(c.city || "")}
      </div>

      <h2>Rechnung ${escapeHtml(invoice.invoiceNo)}</h2>
      <p class="meta">
        Erstellt am: ${escapeHtml(formatDateCH(createdAt))}
        |
        Leistungsmonat: ${escapeHtml(formatMonthCH(invoice.month))}
        | Fällig am: ${escapeHtml(formatDateCH(invoice.dueDate))}
      </p>

      <table>
        <colgroup>
          <col style="width: 5%;" />
          <col style="width: 11%;" />
          <col style="width: 33%;" />
          <col style="width: 9%;" />
          <col style="width: 12%;" />
          <col style="width: 10%;" />
          <col style="width: 8%;" />
          <col style="width: 12%;" />
        </colgroup>
        <thead>
          <tr>
            <th>Tag</th>
            <th>Datum</th>
            <th>Leistung</th>
            <th class="num">Menge</th>
            <th>Einheit</th>
            <th class="num">Ansatz</th>
            <th class="nowrap">Währung</th>
            <th class="num">Total</th>
          </tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>

      <div class="totals">
        <p><span>Zwischentotal</span><span>${formatCurrency(invoice.subtotal)}</span></p>
        <p><span>MwSt ${toNumber(invoice.vatRate, 0).toFixed(2)}%</span><span>${formatCurrency(invoice.vatAmount)}</span></p>
        <p class="grand"><span>Total</span><span>${formatCurrency(invoice.grandTotal)}</span></p>
      </div>

      <section class="invoice-text-block">
        <p>${escapeHtml(company.invoiceText || "")}</p>
      </section>

      ${paymentSlipSection}
    </section>
  `;
}
function buildSwissQrPayload({
  iban,
  currency,
  creditorName,
  creditorStreet,
  creditorZip,
  creditorCity,
  debtorName,
  debtorStreet,
  debtorZip,
  debtorCity,
  amount,
  message
}) {
  return [
    "SPC",
    "0200",
    "1",
    String(iban || ""),
    "S",
    String(creditorName || ""),
    String(creditorStreet || ""),
    "",
    String(creditorZip || ""),
    String(creditorCity || ""),
    "CH",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    `${toNumber(amount, 0).toFixed(2)}`,
    normalizeQrCurrencyCode(currency),
    "S",
    String(debtorName || ""),
    String(debtorStreet || ""),
    "",
    String(debtorZip || ""),
    String(debtorCity || ""),
    "CH",
    "NON",
    "",
    String(message || "Rechnung Fakturix CH"),
    "EPD"
  ].join("\n");
}

function normalizePaymentSlipType(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (raw === "appendix") return "appendix";
  if (raw === "none") return "none";
  return "swiss_qr";
}
function normalizeInvoiceNoToken(value) {
  const raw = String(value || "")
    .trim()
    .toUpperCase()
    .replaceAll("Ä", "AE")
    .replaceAll("Ö", "OE")
    .replaceAll("Ü", "UE")
    .replaceAll("ß", "SS");
  const ascii = raw
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9-]/g, "");
  return ascii || "K";
}

function buildInvoiceNo(customer, month) {
  const baseName = String(customer.lastName || customer.firstName || "K");
  const customerShort = normalizeInvoiceNoToken(baseName).slice(0, 3) || "K";
  const sequence = String(state.invoices.length + 1).padStart(3, "0");
  return `${month.replace("-", "")}-${customerShort}-${sequence}`;
}

function formatCustomerName(customer) {
  const lead = `${customer.company ? `${customer.company} - ` : ""}${customer.firstName || ""} ${customer.lastName || ""}`.trim();
  return lead || "Kunde";
}

function formatEmployeeName(employee) {
  const lead = `${employee.firstName || ""} ${employee.lastName || ""}`.trim();
  return lead || "Mitarbeiter";
}

function addDays(isoDate, days) {
  const date = new Date(`${isoDate}T00:00:00`);
  date.setDate(date.getDate() + days);
  return date.toISOString().slice(0, 10);
}

function formatMonthCH(month) {
  const [year, m] = String(month).split("-");
  if (!year || !m) return String(month);
  return `${m}.${year}`;
}

function formatMonthLongCH(month) {
  const [year, m] = String(month).split("-");
  if (!year || !m) return String(month);
  const date = new Date(`${year}-${m}-01T00:00:00`);
  if (Number.isNaN(date.getTime())) return String(month);
  return new Intl.DateTimeFormat("de-CH", { month: "long", year: "numeric" }).format(date);
}

function formatDateCH(isoDate) {
  const [year, m, day] = String(isoDate || "").split("-");
  if (!year || !m || !day) return String(isoDate || "");
  return `${day}.${m}.${year}`;
}

function formatDateTimeCH(isoDateTime) {
  const value = String(isoDateTime || "").trim();
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat("de-CH", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  }).format(date);
}

function formatWeekdayShortCH(isoDate) {
  const date = new Date(String(isoDate || ""));
  if (Number.isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat("de-CH", { weekday: "short" }).format(date).replace(".", "");
}

function formatWeekdayLongCH(isoDate) {
  const date = new Date(String(isoDate || ""));
  if (Number.isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat("de-CH", { weekday: "long" }).format(date);
}

function formatCurrency(value, currencyCode) {
  const currency = normalizeCurrencyCode(currencyCode || state.company?.currency);
  try {
    return new Intl.NumberFormat("de-CH", { style: "currency", currency }).format(toNumber(value, 0));
  } catch {
    return `${currency} ${formatAmountCH(value)}`;
  }
}

function formatAmountCH(value) {
  return new Intl.NumberFormat("de-CH", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(toNumber(value, 0));
}

function formatNumberCH(value) {
  return new Intl.NumberFormat("de-CH", { minimumFractionDigits: 0, maximumFractionDigits: 2 }).format(toNumber(value, 0));
}

function formatCurrencyAligned(value, currencyCode) {
  const currency = normalizeCurrencyCode(currencyCode || state.company?.currency);
  const amount = formatAmountCH(value);
  return `<span class="money"><span class="cur">${escapeHtml(currency)}</span><span class="amt">${escapeHtml(amount)}</span></span>`;
}

function normalizeCurrencyCode(value) {
  const code = String(value || "").trim().toUpperCase();
  return /^[A-Z]{3}$/.test(code) ? code : "CHF";
}

function normalizeQrCurrencyCode(value) {
  const code = normalizeCurrencyCode(value);
  return code === "CHF" || code === "EUR" ? code : "CHF";
}

function formatPaidMethod(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (raw === "automatisch") return "automatisch";
  if (raw === "manuell") return "manuell";
  return "nicht erfasst";
}

function isOnOrBeforeToday(isoDate) {
  const value = String(isoDate || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const today = new Date().toISOString().slice(0, 10);
  return value <= today;
}

function getOverdueDays(isoDate) {
  const value = String(isoDate || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return 0;
  const dueDate = new Date(`${value}T00:00:00`);
  if (Number.isNaN(dueDate.getTime())) return 0;
  const today = new Date();
  const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const diffMs = todayStart.getTime() - dueDate.getTime();
  return Math.max(0, Math.floor(diffMs / 86400000));
}

function formatIbanDisplay(value) {
  const cleaned = String(value || "").replace(/\s+/g, "").toUpperCase();
  if (!cleaned) return "";
  return cleaned.replace(/(.{4})/g, "$1 ").trim();
}

function matchesSearch(text, searchTerm) {
  const q = String(searchTerm || "").trim().toLocaleLowerCase("de-CH");
  if (!q) return true;
  return String(text || "").toLocaleLowerCase("de-CH").includes(q);
}


function toNumber(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function generateId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(36).slice(2, 11)}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function openPrintWindow(invoiceHtml, suggestedTitle = "Rechnung", options = {}) {
  const win = window.open("", "_blank");
  if (!win) {
    billingStatus.textContent = "Popup wurde blockiert. Bitte Popups erlauben und erneut drucken.";
    return;
  }
  const title = sanitizeDocumentTitle(suggestedTitle);
  const autoPrint = options?.autoPrint !== false;
  const closeAfterPrint = options?.closeAfterPrint !== false;
  const orientation = options?.orientation === "landscape" ? "landscape" : "portrait";

  win.document.open();
  win.document.write(buildPrintDocumentHtml(invoiceHtml, title, autoPrint, closeAfterPrint, orientation));
  win.document.close();
}

async function saveInvoiceAsPdf(invoiceHtml, suggestedTitle = "Rechnung", options = {}) {
  const title = sanitizeDocumentTitle(suggestedTitle);
  const orientation = options?.orientation === "landscape" ? "landscape" : "portrait";
  const JsPdfClass = globalThis?.jspdf?.jsPDF;
  if (!JsPdfClass || typeof globalThis.html2canvas !== "function") {
    billingStatus.textContent = "PDF-Library nicht verfügbar. Fallback auf Druckdialog.";
    openPrintWindow(invoiceHtml, title, { closeAfterPrint: false, orientation });
    return;
  }

  const temp = document.createElement("div");
  temp.style.position = "fixed";
  temp.style.left = "-20000px";
  temp.style.top = "0";
  temp.style.width = orientation === "landscape" ? "1123px" : "794px";
  temp.style.background = "#ffffff";
  temp.innerHTML = `<article class="invoice-preview print-only-preview">${invoiceHtml}</article>`;
  document.body.appendChild(temp);

  billingStatus.textContent = "PDF wird erstellt...";
  try {
    const canvas = await globalThis.html2canvas(temp, {
      scale: 2,
      useCORS: true,
      backgroundColor: "#ffffff"
    });

    const pdf = new JsPdfClass({ unit: "mm", format: "a4", orientation });
    const margin = 8;
    const pageWidth = orientation === "landscape" ? 297 : 210;
    const pageHeight = orientation === "landscape" ? 210 : 297;
    const contentWidthMm = pageWidth - margin * 2;
    const contentHeightMm = pageHeight - margin * 2;
    const pxPerMm = canvas.width / contentWidthMm;
    const pageSlicePx = Math.max(1, Math.floor(contentHeightMm * pxPerMm));

    let offsetY = 0;
    let firstPage = true;
    while (offsetY < canvas.height) {
      const sliceHeight = Math.min(pageSlicePx, canvas.height - offsetY);
      const sliceCanvas = document.createElement("canvas");
      sliceCanvas.width = canvas.width;
      sliceCanvas.height = sliceHeight;
      const sliceCtx = sliceCanvas.getContext("2d");
      if (!sliceCtx) throw new Error("Canvas context unavailable");
      sliceCtx.drawImage(
        canvas,
        0,
        offsetY,
        canvas.width,
        sliceHeight,
        0,
        0,
        canvas.width,
        sliceHeight
      );
      const imgData = sliceCanvas.toDataURL("image/png");
      const imgHeightMm = sliceHeight / pxPerMm;
      if (!firstPage) pdf.addPage();
      pdf.addImage(imgData, "PNG", margin, margin, contentWidthMm, imgHeightMm, undefined, "FAST");
      firstPage = false;
      offsetY += sliceHeight;
    }
    const fileName = `${title}.pdf`;
    const pdfBlob = pdf.output("blob");
    const saveResult = await savePdfBlobWithDialog(pdfBlob, fileName);
    if (saveResult === "unsupported" || saveResult === "error") {
      downloadBlob(pdfBlob, fileName);
      billingStatus.textContent = `PDF heruntergeladen: ${fileName}`;
    } else if (saveResult === "canceled") {
      billingStatus.textContent = "Speichern wurde abgebrochen.";
    } else {
      billingStatus.textContent = `PDF gespeichert: ${fileName}`;
    }
  } catch (error) {
    console.error(error);
    billingStatus.textContent = "PDF-Erstellung fehlgeschlagen. Fallback auf Druckdialog.";
    openPrintWindow(invoiceHtml, title, { closeAfterPrint: false, orientation });
  } finally {
    temp.remove();
  }
}

function buildPrintDocumentHtml(invoiceHtml, title, autoPrint, closeAfterPrint = true, orientation = "portrait") {
  const printScript = autoPrint
    ? `
        <script>
          window.onload = () => {
            window.print();
            ${closeAfterPrint ? "window.onafterprint = () => window.close();" : ""}
          };
        <\/script>
      `
    : "";
  const orientationCss = orientation === "landscape"
    ? `<style>@page { size: A4 landscape; margin: 8mm; }</style>`
    : "";
  return `
    <!doctype html>
    <html lang="de-CH">
      <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>${escapeHtml(title)}</title>
        <base href="${window.location.href}" />
        <link rel="stylesheet" href="styles.css" />
        ${orientationCss}
      </head>
      <body class="print-only-body">
        <article class="invoice-preview print-only-preview">${invoiceHtml}</article>
        ${printScript}
      </body>
    </html>
  `;
}

function sanitizeDocumentTitle(value) {
  const raw = String(value || "").trim();
  const cleaned = raw
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned || "Dokument";
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

function setPreviewHtml(html) {
  currentInvoiceHtml = html || "";
  invoicePreview.innerHTML = currentInvoiceHtml;
}

function renderSwissQrBillSvg(invoice, customer, company) {
  try {
    const SwissQRBillClass = globalThis?.SwissQRBill?.svg?.SwissQRBill;
    if (!SwissQRBillClass) {
      return `
        <div class="qr-fallback">
          Swiss QR Bill Library nicht geladen.
        </div>
      `;
    }

    const iban = String(company.iban || "").replace(/\s+/g, "");
    const qrIban = isQrIban(iban);
    const data = {
      creditor: {
        account: iban,
        name: String(company.name || ""),
        address: String(company.street || ""),
        zip: String(company.zip || ""),
        city: String(company.city || ""),
        country: "CH"
      },
      amount: toNumber(invoice.grandTotal, 0),
      currency: normalizeQrCurrencyCode(invoice.currency || company.currency),
      reference: qrIban ? buildQrrReference(invoice.invoiceNo) : undefined,
      debtor: {
        name: String(formatCustomerName(customer) || ""),
        address: String(customer.street || ""),
        zip: String(customer.zip || ""),
        city: String(customer.city || ""),
        country: "CH"
      },
      additionalInformation: `Rechnung ${invoice.invoiceNo}`
    };

    const validationErrors = validateSwissQrData(data, qrIban);
    if (validationErrors.length) {
      const items = validationErrors.map((e) => `<li>${escapeHtml(e)}</li>`).join("");
      return `
        <div class="qr-fallback">
          <strong>QR-Bill Validierung fehlgeschlagen:</strong>
          <ul>${items}</ul>
        </div>
      `;
    }

    const qrBill = new SwissQRBillClass(data, {
      language: "DE",
      scissors: true,
      outlines: true
    });

    if (typeof qrBill.toString === "function") {
      return `<div class="qr-lib-wrap">${qrBill.toString()}</div>`;
    }

    if (qrBill.element) {
      return `<div class="qr-lib-wrap">${qrBill.element.outerHTML}</div>`;
    }

    return `<div class="qr-fallback">Swiss QR Bill konnte nicht erzeugt werden.</div>`;
  } catch (error) {
    console.error(error);
    return `<div class="qr-fallback">Fehler bei QR-Bill-Erzeugung.</div>`;
  }
}

function isQrIban(iban) {
  const cleaned = String(iban || "").replace(/\s+/g, "");
  if (cleaned.length < 9) return false;
  const iid = Number(cleaned.slice(4, 9));
  return Number.isFinite(iid) && iid >= 30000 && iid <= 31999;
}

function buildQrrReference(seed) {
  const digits = String(seed || "").replace(/\D/g, "");
  const base26 = digits.padStart(26, "0").slice(-26);
  const check = qrrChecksum(base26);
  return `${base26}${check}`;
}

function qrrChecksum(base26) {
  const table = [0, 9, 4, 6, 8, 2, 7, 1, 3, 5];
  let carry = 0;
  for (const ch of String(base26)) {
    carry = table[(carry + Number(ch)) % 10];
  }
  return String((10 - carry) % 10);
}

function validateSwissQrData(data, qrIban) {
  const errors = [];

  const iban = String(data?.creditor?.account || "");
  if (!iban) {
    errors.push("IBAN fehlt.");
  } else {
    if (!/^[A-Z]{2}[0-9A-Z]{13,32}$/.test(iban)) {
      errors.push("IBAN-Format ist ungültig.");
    }
    if (!isValidIban(iban)) {
      errors.push("IBAN-Prüfziffer ist ungültig.");
    }
    const country = iban.slice(0, 2);
    if (country !== "CH" && country !== "LI") {
      errors.push("Nur CH/LI-IBAN ist erlaubt.");
    }
  }

  if (!String(data?.creditor?.name || "").trim()) errors.push("Kreditor Name fehlt.");
  if (!String(data?.creditor?.address || "").trim()) errors.push("Kreditor Adresse fehlt.");
  if (!String(data?.creditor?.zip || "").trim()) errors.push("Kreditor PLZ fehlt.");
  if (!String(data?.creditor?.city || "").trim()) errors.push("Kreditor Ort fehlt.");

  if (!String(data?.debtor?.name || "").trim()) errors.push("Debitor Name fehlt.");
  if (!String(data?.debtor?.address || "").trim()) errors.push("Debitor Adresse fehlt.");
  if (!String(data?.debtor?.zip || "").trim()) errors.push("Debitor PLZ fehlt.");
  if (!String(data?.debtor?.city || "").trim()) errors.push("Debitor Ort fehlt.");

  const amount = Number(data?.amount);
  if (!Number.isFinite(amount) || amount < 0) errors.push("Betrag ist ungültig.");
  if (String(data?.currency || "") !== "CHF" && String(data?.currency || "") !== "EUR") {
    errors.push("Währung muss CHF oder EUR sein.");
  }

  const info = String(data?.additionalInformation || "");
  if (info.length > 140) errors.push("Mitteilung darf max. 140 Zeichen haben.");

  const reference = String(data?.reference || "");
  if (qrIban) {
    if (!reference) {
      errors.push("Bei QR-IBAN ist eine QRR-Referenz zwingend.");
    } else if (!/^\d{27}$/.test(reference)) {
      errors.push("QRR-Referenz muss 27-stellig numerisch sein.");
    } else {
      const base = reference.slice(0, 26);
      const check = reference.slice(26);
      if (qrrChecksum(base) !== check) {
        errors.push("QRR-Referenz Prüfziffer ist ungültig.");
      }
    }
  } else if (reference) {
    errors.push("Bei normaler IBAN darf keine QR-Referenz verwendet werden.");
  }

  return errors;
}

function isValidIban(ibanRaw) {
  const iban = String(ibanRaw || "").replace(/\s+/g, "").toUpperCase();
  if (iban.length < 15 || iban.length > 34) return false;
  const rearranged = iban.slice(4) + iban.slice(0, 4);
  let numeric = "";
  for (const ch of rearranged) {
    if (/[A-Z]/.test(ch)) numeric += String(ch.charCodeAt(0) - 55);
    else if (/[0-9]/.test(ch)) numeric += ch;
    else return false;
  }
  let remainder = 0;
  for (const d of numeric) {
    remainder = (remainder * 10 + Number(d)) % 97;
  }
  return remainder === 1;
}





































































































































































































































































