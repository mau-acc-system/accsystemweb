// ==========================================
// ACC SYSTEM - EXCEL DATABASE ENGINE
// ==========================================

// DATABASE CONFIGURATION
const CONFIG = {
  owner: "mau-acc-system",
  repo: "accsystemdb",
  branch: "main",
  file: "Excel-Database.xlsx"
};

const LICENSE_FILE = "./license.tk";

const API_URL =
  `https://api.github.com/repos/${CONFIG.owner}/${CONFIG.repo}/contents/${encodeURIComponent(CONFIG.file)}`;

let githubToken = null;
let workbook = null;
let fileSHA = null;

// ==========================================
// LICENSE DECODER
// ==========================================

function decodeLicense(encoded) {
  const hex = encoded.trim();

  if (
    !hex ||
    hex.length % 2 !== 0 ||
    !/^[0-9a-f]+$/i.test(hex)
  ) {
    throw new Error("Invalid license file.");
  }

  const bytes = new Uint8Array(
    hex.match(/.{2}/g).map(byte => parseInt(byte, 16))
  );

  return new TextDecoder().decode(bytes);
}

export async function loadLicense() {
  const response = await fetch(LICENSE_FILE, {
    cache: "no-store"
  });

  if (!response.ok) {
    throw new Error("License file not found.");
  }

  const encoded = await response.text();
  const token = decodeLicense(encoded);

  if (!token) {
    throw new Error("Empty license.");
  }

  githubToken = token;

  return true;
}

export function setGitHubToken(token) {
  githubToken = token?.trim() || null;
}

export function clearGitHubToken() {
  githubToken = null;
  workbook = null;
  fileSHA = null;
}

export function hasLicense() {
  return Boolean(githubToken);
}

// ==========================================
// LICENSE VALIDATION
// ==========================================

export async function validateLicense() {
  try {
    await loadLicense();

    const response = await githubRequest(API_URL);

    if (!response.ok) {
      clearGitHubToken();

      return {
        valid: false,
        message: "Invalid license or access denied."
      };
    }

    const data = await response.json();
    fileSHA = data.sha;

    return {
      valid: true,
      message: "License verified."
    };

  } catch (error) {
    clearGitHubToken();

    return {
      valid: false,
      message: error.message
    };
  }
}

// ==========================================
// GITHUB API
// ==========================================

async function githubRequest(url, options = {}) {
  if (!githubToken) {
    throw new Error("License required.");
  }

  return fetch(url, {
    ...options,
    headers: {
      "Authorization": `Bearer ${githubToken}`,
      "Accept": "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
      ...options.headers
    }
  });
}

// ==========================================
// READ EXCEL DATABASE
// ==========================================

export async function readExcel() {
  if (!githubToken) {
    await loadLicense();
  }

  if (typeof XLSX === "undefined") {
    throw new Error("SheetJS library is not loaded.");
  }

  const response = await githubRequest(API_URL);

  if (!response.ok) {
    throw new Error("Unable to access Excel database.");
  }

  const data = await response.json();

  if (!data.content) {
    throw new Error("Excel database content unavailable.");
  }

  const binary = atob(data.content.replace(/\s/g, ""));
  const bytes = new Uint8Array(binary.length);

  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }

  workbook = XLSX.read(bytes, {
    type: "array"
  });

  fileSHA = data.sha;

  return getSheetNames();
}

// ==========================================
// SHEET FUNCTIONS
// ==========================================

export function getSheetNames() {
  if (!workbook) {
    throw new Error("Database not loaded.");
  }

  return workbook.SheetNames;
}

export function getSheet(sheetName) {
  if (!workbook) {
    throw new Error("Database not loaded.");
  }

  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }

  return XLSX.utils.sheet_to_json(sheet, {
    defval: ""
  });
}

// ==========================================
// ADD RECORD
// ==========================================

export function addRecord(sheetName, record) {
  const sheet = getWorksheet(sheetName);
  const headers = getHeaders(sheet);

  const invalidFields = Object.keys(record).filter(
    key => !headers.includes(key)
  );

  if (invalidFields.length) {
    throw new Error(
      `Unknown column(s): ${invalidFields.join(", ")}`
    );
  }

  const newRecord = {};

  headers.forEach(header => {
    newRecord[header] = record[header] ?? "";
  });

  XLSX.utils.sheet_add_json(sheet, [newRecord], {
    header: headers,
    skipHeader: true,
    origin: -1
  });

  return true;
}

// ==========================================
// UPDATE RECORD
// ==========================================

export function updateRecord(
  sheetName,
  column,
  value,
  updates
) {
  const sheet = getWorksheet(sheetName);
  const headers = getHeaders(sheet);
  const records = getSheet(sheetName);

  const index = records.findIndex(
    row => String(row[column]) === String(value)
  );

  if (index === -1) {
    throw new Error("Record not found.");
  }

  Object.assign(records[index], updates);

  rebuildSheet(sheetName, headers, records);

  return true;
}

// ==========================================
// DELETE RECORD
// ==========================================

export function deleteRecord(
  sheetName,
  column,
  value
) {
  const sheet = getWorksheet(sheetName);
  const headers = getHeaders(sheet);
  const records = getSheet(sheetName);

  const filtered = records.filter(
    row => String(row[column]) !== String(value)
  );

  if (filtered.length === records.length) {
    throw new Error("Record not found.");
  }

  rebuildSheet(sheetName, headers, filtered);

  return true;
}

// ==========================================
// SAVE EXCEL DATABASE
// ==========================================

export async function saveExcel(
  message = "Update Excel database"
) {
  if (!githubToken) {
    throw new Error("License required.");
  }

  if (!workbook) {
    throw new Error("Database not loaded.");
  }

  const content = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "base64"
  });

  const response = await githubRequest(API_URL, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      message,
      content,
      sha: fileSHA,
      branch: CONFIG.branch
    })
  });

  if (!response.ok) {
    if (
      response.status === 409 ||
      response.status === 422
    ) {
      throw new Error(
        "Database conflict. Reload before saving again."
      );
    }

    throw new Error("Failed to save Excel database.");
  }

  const result = await response.json();

  fileSHA = result.content.sha;

  return true;
}

// ==========================================
// INTERNAL HELPERS
// ==========================================

function getWorksheet(sheetName) {
  if (!workbook) {
    throw new Error("Database not loaded.");
  }

  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }

  return sheet;
}

function getHeaders(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  return rows[0] || [];
}

function rebuildSheet(sheetName, headers, records) {
  const newSheet = XLSX.utils.json_to_sheet(records, {
    header: headers
  });

  workbook.Sheets[sheetName] = newSheet;
}

// ==========================================
// DATABASE STATUS
// ==========================================

export function isDatabaseLoaded() {
  return Boolean(workbook);
}
