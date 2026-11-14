
import {
  validateLicense,
  readExcel,
  getSheet,
  addRecord,
  updateRecord,
  saveExcel
} from "../excel.js";

// ==========================================
// CONFIGURATION
// ==========================================

const cfg = window.MODULE_CONFIG;
const $ = id => document.getElementById(id);

const session = JSON.parse(
  sessionStorage.getItem("accsystem_session") || "null"
);

if (!session) {
  window.top.location.href = "../login.html";
  throw new Error("No active session");
}

if (!cfg) {
  throw new Error("Module configuration not found.");
}

const access = session.Access || [];
const admin = session.Role === "Administrator";

const allowed = perm =>
  admin ||
  access.some(x =>
    String(x.Module).trim() === cfg.permission &&
    (
      x[perm] === true ||
      String(x[perm]).toLowerCase() === "true" ||
      String(x[perm]) === "1"
    )
  );

// ==========================================
// PERMISSION CHECK
// ==========================================

if (!allowed("CanView")) {
  document.body.innerHTML = "<h2>Access denied</h2>";
  throw new Error("No view permission");
}

// ==========================================
// VARIABLES AND HELPERS
// ==========================================

let data = [];

const esc = v =>
  String(v ?? "").replace(/[&<>"']/g, c => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  }[c]));

const n = v => Number(v) || 0;

const now = () => new Date().toISOString();

const id = (prefix, arr, key) =>
  prefix +
  String(
    Math.max(
      0,
      ...arr.map(x =>
        parseInt(String(x[key] || "").replace(/\D/g, "")) || 0
      )
    ) + 1
  ).padStart(5, "0");

const sheet = s => getSheet(s) || [];

// ==========================================
// MESSAGE
// ==========================================

function message(t, bad = false) {
  $("msg").textContent = t;
  $("msg").className = "msg " + (bad ? "bad" : "good");
}

// ==========================================
// SELECT OPTIONS
// ==========================================

function options(field) {
  if (!field.options) return "";

  let vals = [];

  if (field.options === "customers") vals = sheet("Customers");
  else if (field.options === "vendors") vals = sheet("Vendors");
  else if (field.options === "items") vals = sheet("Items");
  else if (field.options === "banks") vals = sheet("Banks");
  else if (field.options === "cash") vals = sheet("CashAccounts");
  else if (field.options === "coa") vals = sheet("ChartOfAccounts");

  return vals.map(x => {
    const value = x[field.valueKey || "CustomerID"];

    const label =
      x[field.labelKey || "CustomerName"] ||
      x.AccountName ||
      x.ItemName ||
      x.BankName ||
      x.CashAccountName ||
      value;

    return `<option value="${esc(value)}">${esc(label)}</option>`;
  }).join("");
}

// ==========================================
// INPUT GENERATOR
// ==========================================

function input(f) {
  const v = f.value || "";

  if (f.type === "select") {
    return `
      <select name="${f.name}" ${f.required ? "required" : ""}>
        ${
          f.options
            ? options(f)
            : f.choices.map(x =>
                `<option>${esc(x)}</option>`
              ).join("")
        }
      </select>
    `;
  }

  return `
    <input
      name="${f.name}"
      type="${f.type || "text"}"
      value="${esc(v)}"
      ${f.required ? "required" : ""}
      ${f.step ? `step="${f.step}"` : ""}
      ${f.min !== undefined ? `min="${f.min}"` : ""}
    >
  `;
}

// ==========================================
// RENDER
// ==========================================

function render() {
  $("title").textContent = cfg.title;
  $("desc").textContent = cfg.desc;

  $("formFields").innerHTML = cfg.fields
    .map(f => `
      <label>
        ${esc(f.label)}
        ${input(f)}
      </label>
    `).join("");

  $("save").disabled = !allowed("CanCreate");

  $("tableHead").innerHTML =
    cfg.columns.map(c => `<th>${esc(c.label)}</th>`).join("") +
    "<th>Action</th>";

  $("tableBody").innerHTML = data.map((r, i) => `
    <tr>
      ${cfg.columns.map(c => `<td>${esc(r[c.key])}</td>`).join("")}
      <td>
        ${
          allowed("CanEdit")
            ? `<button data-edit="${i}">Edit</button>`
            : ""
        }
        ${
          allowed("CanDelete")
            ? `<button class="danger" data-del="${i}">Deactivate</button>`
            : ""
        }
      </td>
    </tr>
  `).join("") ||
  `<tr>
    <td colspan="${cfg.columns.length + 1}" class="empty">
      No records found.
    </td>
  </tr>`;
}

// ==========================================
// LOAD DATABASE WITH DIAGNOSTICS
// ==========================================

async function load() {
  try {
    console.log("STEP 1: Validating license...");

    const lic = await validateLicense();

    console.log("LICENSE RESULT:", lic);

    if (lic === false || lic?.valid === false) {
      throw new Error(lic?.message || "License invalid");
    }

    console.log("STEP 2: Reading Excel database...");

    const result = await readExcel();

    console.log("WORKBOOK SHEETS:", result);

    console.log("STEP 3: Checking database status...");

    const records = getSheet(cfg.sheet);

    console.log("CURRENT MODULE SHEET:", cfg.sheet);
    console.log("RECORD COUNT:", records.length);

    data = records;

    render();

    console.log("STEP 4: Database loaded successfully.");

    message("Database loaded successfully.");

  } catch (e) {
    console.error("DATABASE LOAD ERROR:", e);

    message("Database error: " + e.message, true);
  }
}

// ==========================================
// SAVE RECORD
// ==========================================

$("form").addEventListener("submit", async e => {
  e.preventDefault();

  try {
    if (!allowed("CanCreate")) {
      throw new Error("No create permission");
    }

    console.log("SAVE: Checking database...");

    const fd = new FormData(e.target);
    const rec = {};

    cfg.fields.forEach(f => {
      rec[f.name] = fd.get(f.name);
    });

    cfg.fields.forEach(f => {
      if (f.type === "number") {
        rec[f.name] = n(rec[f.name]);
      }
    });

    const t = now();

    if (cfg.beforeSave) {
      await cfg.beforeSave(rec);
    }

    rec[cfg.key] = id(cfg.prefix, data, cfg.key);

    const docNo = {
      Invoices: "InvoiceNo",
      Bills: "BillNo",
      CustomerPayments: "PaymentNo",
      VendorPayments: "PaymentNo"
    }[cfg.sheet];

    if (docNo && !rec[docNo]) {
      rec[docNo] = id(cfg.prefix, data, docNo);
    }

    rec.CreatedAt = t;
    rec.UpdatedAt = t;

    if (cfg.defaults) {
      const defaultValues =
        typeof cfg.defaults === "function"
          ? cfg.defaults(rec)
          : cfg.defaults;

      Object.assign(rec, defaultValues);
    }

    console.log("SAVE: Record prepared:", rec);

    console.log("SAVE: Adding record to sheet:", cfg.sheet);

    addRecord(cfg.sheet, rec);

    if (cfg.afterSave) {
      await cfg.afterSave(rec);
    }

    console.log("SAVE: Uploading Excel database...");

    await saveExcel();

    console.log("SAVE: Successful.");

    data = getSheet(cfg.sheet);

    e.target.reset();

    render();

    message("Record saved successfully.");

  } catch (err) {
    console.error("SAVE ERROR:", err);

    message(err.message, true);
  }
});

// ==========================================
// EDIT AND DEACTIVATE
// ==========================================

$("tableBody").addEventListener("click", async e => {
  const del = e.target.dataset.del;
  const edit = e.target.dataset.edit;

  // DEACTIVATE
  if (del !== undefined) {
    const r = data[Number(del)];

    if (!confirm("Deactivate this record?")) return;

    try {
      if (!allowed("CanDelete")) {
        throw new Error("No delete permission");
      }

      updateRecord(
        cfg.sheet,
        cfg.key,
        r[cfg.key],
        {
          IsActive: false,
          UpdatedAt: now()
        }
      );

      await saveExcel();

      data = getSheet(cfg.sheet);

      render();

      message("Record deactivated.");

    } catch (err) {
      console.error("DEACTIVATE ERROR:", err);
      message(err.message, true);
    }
  }

  // EDIT
  else if (edit !== undefined) {
    const r = data[Number(edit)];

    if (!allowed("CanEdit")) return;

    const changes = {};

    for (const f of cfg.fields) {
      const v = prompt(
        "Edit " + f.label,
        r[f.name] ?? ""
      );

      if (v === null) return;

      changes[f.name] =
        f.type === "number" ? n(v) : v;
    }

    try {
      updateRecord(
        cfg.sheet,
        cfg.key,
        r[cfg.key],
        {
          ...changes,
          UpdatedAt: now()
        }
      );

      await saveExcel();

      data = getSheet(cfg.sheet);

      render();

      message("Record updated.");

    } catch (err) {
      console.error("EDIT ERROR:", err);
      message(err.message, true);
    }
  }
});

// ==========================================
// INITIALIZE
// ==========================================

load();
