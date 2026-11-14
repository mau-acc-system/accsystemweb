/* =========================================
   ACCOUNTING SYSTEM
   EXCEL DATABASE ENGINE
   ========================================= */

const CONFIG = {
    owner: "mau-acc-system",
    repo: "accsystemdb",
    branch: "main",
    file: "Excel-Database.xlsx"
};

const API_URL = `https://api.github.com/repos/${CONFIG.owner}/${CONFIG.repo}/contents/${encodeURIComponent(CONFIG.file)}`;

let githubToken = null;
let workbook = null;
let fileSHA = null;

// =========================================
// 1. CONFIGURATION
// =========================================

export function setGitHubToken(token) {
    if (!token || !token.trim()) {
        throw new Error("GitHub token is required.");
    }

    githubToken = token.trim();
}

// =========================================
// 2. API REQUEST
// =========================================

async function githubRequest(url, options = {}) {

    if (!githubToken) {
        throw new Error("GitHub token not configured.");
    }

    const response = await fetch(url, {
        ...options,
        headers: {
            "Accept": "application/vnd.github+json",
            "Authorization": `Bearer ${githubToken}`,
            "X-GitHub-Api-Version": "2022-11-28",
            ...options.headers
        }
    });

    const data = await response.json();

    if (!response.ok) {
        throw new Error(
            data.message || `GitHub error: ${response.status}`
        );
    }

    return data;
}

// =========================================
// 3. READ EXCEL DATABASE
// =========================================

export async function readExcel() {

    const data = await githubRequest(
        `${API_URL}?ref=${encodeURIComponent(CONFIG.branch)}`
    );

    if (!data.content) {
        throw new Error("Excel file content not found.");
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

    return workbook;
}

// =========================================
// 4. GET SHEET DATA
// =========================================

export function getSheet(sheetName) {

    if (!workbook) {
        throw new Error("Read Excel database first.");
    }

    const sheet = workbook.Sheets[sheetName];

    if (!sheet) {
        throw new Error(`Sheet '${sheetName}' not found.`);
    }

    return XLSX.utils.sheet_to_json(sheet, {
        defval: ""
    });
}

// =========================================
// 5. ADD RECORD
// =========================================

export function addRecord(sheetName, record) {

    if (!workbook) {
        throw new Error("Read Excel database first.");
    }

    const sheet = workbook.Sheets[sheetName];

    if (!sheet) {
        throw new Error(`Sheet '${sheetName}' not found.`);
    }

    XLSX.utils.sheet_add_json(sheet, [record], {
        skipHeader: true,
        origin: -1
    });

    return true;
}

// =========================================
// 6. UPDATE RECORD
// =========================================

export function updateRecord(
    sheetName,
    column,
    value,
    updates
) {

    if (!workbook) {
        throw new Error("Read Excel database first.");
    }

    const records = getSheet(sheetName);

    const index = records.findIndex(
        row => String(row[column]) === String(value)
    );

    if (index === -1) {
        throw new Error("Record not found.");
    }

    const updated = {
        ...records[index],
        ...updates
    };

    records[index] = updated;

    workbook.Sheets[sheetName] =
        XLSX.utils.json_to_sheet(records);

    return true;
}

// =========================================
// 7. DELETE RECORD
// =========================================

export function deleteRecord(
    sheetName,
    column,
    value
) {

    if (!workbook) {
        throw new Error("Read Excel database first.");
    }

    const records = getSheet(sheetName);

    const filtered = records.filter(
        row => String(row[column]) !== String(value)
    );

    if (filtered.length === records.length) {
        throw new Error("Record not found.");
    }

    workbook.Sheets[sheetName] =
        XLSX.utils.json_to_sheet(filtered);

    return true;
}

// =========================================
// 8. SAVE EXCEL TO GITHUB
// =========================================

export async function saveExcel(message = "Update Excel database") {

    if (!workbook || !fileSHA) {
        throw new Error("Load Excel database before saving.");
    }

    const base64 = XLSX.write(workbook, {
        bookType: "xlsx",
        type: "base64"
    });

    const result = await githubRequest(API_URL, {
        method: "PUT",

        headers: {
            "Content-Type": "application/json"
        },

        body: JSON.stringify({
            message,
            content: base64,
            sha: fileSHA,
            branch: CONFIG.branch
        })
    });

    fileSHA = result.content.sha;

    return {
        success: true,
        message: "Excel database successfully updated!",
        sha: fileSHA
    };
}

// =========================================
// 9. GET AVAILABLE SHEETS
// =========================================

export function getSheetNames() {

    if (!workbook) {
        throw new Error("Read Excel database first.");
    }

    return workbook.SheetNames;
}

// =========================================
// 10. DATABASE STATUS
// =========================================

export function isDatabaseLoaded() {
    return workbook !== null;
}
