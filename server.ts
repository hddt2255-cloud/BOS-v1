import express from "express";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const SHEET_NAMES = {
  MASTER_RKAS: "Database Master RKAS",
  BELANJA: "Database Belanja",
  SI: "SI",
  SURAT_TRANSFER: "Surat Transfer",
  REKAP_SERAPAN: "Rekap Serapan",
  REKENING_KORAN: "Rekening Koran",
  MONTHS: [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ],
  USERS: "Role Pengguna",
  DISTRIBUSI: "Realisasi Distribusi"
};

const ORDERED_SHEETS = [
  SHEET_NAMES.MASTER_RKAS,
  SHEET_NAMES.BELANJA,
  SHEET_NAMES.SI,
  SURAT_TRANSFER_NAME(),
  SHEET_NAMES.REKAP_SERAPAN,
  SHEET_NAMES.REKENING_KORAN,
  "Januari", "Februari", "Maret", "April", "Mei", "Juni",
  "Juli", "Agustus", "September", "Oktober", "November", "Desember",
  SHEET_NAMES.USERS,
  SHEET_NAMES.DISTRIBUSI
];

function SURAT_TRANSFER_NAME() {
  return SHEET_NAMES.SURAT_TRANSFER;
}

const SHEET_HEADERS: Record<string, string[]> = {
  "Database Master RKAS": ["No. Urut", "Kode Rekening", "Kode Kegiatan", "Uraian Kegiatan", "Jumlah", "Belanja Operasi", "Belanja Modal", "Level", "Status Aksi", "Keterangan", "NPSN"],
  "Database Belanja": ["NO", "BELANJA", "NOMOR REKENING", "NAMA BANK", "ATAS NAMA", "NPWP", "TOTAL", "SUMBER DANA", "TGL TRANSAKSI SI", "BULAN", "MASTER NO", "SUB ITEMS JSON", "TGL INPUT", "ID", "NPSN"],
  "SI": ["NO", "BELANJA", "NOMOR REKENING", "NAMA BANK", "ATAS NAMA", "NPWP", "TOTAL", "SUMBER DANA", "TGL TRANSAKSI SI", "BULAN", "NPSN"],
  "Surat Transfer": ["NO", "TANGGAL SURAT", "NOMOR SURAT", "TGL TRANSAKSI SI", "NOMOR REKENING SEKOLAH", "NAMA REKENING / SEKOLAH", "UNIT KERJA", "ALAMAT", "KEPALA SEKOLAH", "BENDAHARA", "NO HP", "EMAIL", "BANK TUJUAN", "JUMLAH TOTAL", "TERBILANG", "NPSN"],
  "Rekap Serapan": ["No. Urut", "Kode Rekening", "Kode Kegiatan", "Uraian Master RKAS", "Pagu Target (Rp)", "Realisasi Belanja (Rp)", "Sisa Pagu (Rp)", "Status", "Keterangan Catatan", "NPSN"],
  "Rekening Koran": ["NO", "TGL", "DEBIT", "KREDIT", "TOTAL", "JENIS TRANSAKSI", "KETERANGAN", "ID TRANSAKSI", "NPSN"],
  "PERBULAN_DEFAULT": ["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Volume", "Satuan", "Tarif Harga", "Jumlah", "Level", "NPSN"],
  "Role Pengguna": ["NO", "ROLE", "USER / NPSN", "PASSWORD", "NAMA SEKOLAH", "KEPALA SEKOLAH", "BENDAHARA SEKOLAH", "ALAMAT", "KABUPATEN", "PROVINSI", "TAHUN ANGGARAN", "SUMBER DANA", "NO REKENING", "NAMA BANK", "EMAIL", "NO HP", "SALDO AWAL", "UPDATED AT", "EXTRA JSON"],
  "Realisasi Distribusi": ["NO", "TANGGAL", "MASTER NO", "KODE REKENING", "KODE KEGIATAN", "NAMA BARANG", "JUMLAH KELUAR", "PENERIMA", "KETERANGAN", "ID", "NPSN"]
};

const DATA_DIR = path.join(__dirname, "data");
const SPREADSHEET_DB_FILE = path.join(DATA_DIR, "sibos_spreadsheet_20sheets.json");
const SHEETS_META_FILE = path.join(DATA_DIR, "sibos_sheets_meta.json");

type SheetStore = Record<string, any[][]>;

interface SheetsMeta {
  spreadsheetId: string;
  spreadsheetUrl: string;
  spreadsheetTitle: string;
  updatedAt: string;
}

let spreadsheetStore: SheetStore = {};
let sheetsMeta: SheetsMeta = {
  spreadsheetId: "",
  spreadsheetUrl: "",
  spreadsheetTitle: "SIBOS-V1",
  updatedAt: ""
};

function loadSpreadsheetFromDisk(): SheetStore {
  try {
    if (!fs.existsSync(DATA_DIR)) {
      fs.mkdirSync(DATA_DIR, { recursive: true });
    }
    if (fs.existsSync(SPREADSHEET_DB_FILE)) {
      const raw = fs.readFileSync(SPREADSHEET_DB_FILE, "utf-8");
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === "object") {
        return parsed;
      }
    }
  } catch (err) {
    console.error("Failed to load spreadsheet store:", err);
  }
  return {};
}

function saveSpreadsheetToDisk(store: SheetStore) {
  try {
    if (!fs.existsSync(DATA_DIR)) {
      fs.mkdirSync(DATA_DIR, { recursive: true });
    }
    fs.writeFileSync(SPREADSHEET_DB_FILE, JSON.stringify(store, null, 2), "utf-8");
  } catch (err) {
    console.error("Failed to save spreadsheet store:", err);
  }
}

function loadSheetsMetaFromDisk(): SheetsMeta {
  try {
    if (!fs.existsSync(DATA_DIR)) {
      fs.mkdirSync(DATA_DIR, { recursive: true });
    }
    if (fs.existsSync(SHEETS_META_FILE)) {
      const raw = fs.readFileSync(SHEETS_META_FILE, "utf-8");
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === "object") {
        return {
          spreadsheetId: String(parsed.spreadsheetId || ""),
          spreadsheetUrl: String(parsed.spreadsheetUrl || ""),
          spreadsheetTitle: String(parsed.spreadsheetTitle || "SIBOS-V1"),
          updatedAt: String(parsed.updatedAt || "")
        };
      }
    }
  } catch (err) {
    console.error("Failed to load sheets meta:", err);
  }
  return {
    spreadsheetId: "",
    spreadsheetUrl: "",
    spreadsheetTitle: "SIBOS-V1",
    updatedAt: ""
  };
}

function saveSheetsMetaToDisk(meta: SheetsMeta) {
  try {
    if (!fs.existsSync(DATA_DIR)) {
      fs.mkdirSync(DATA_DIR, { recursive: true });
    }
    sheetsMeta = meta;
    fs.writeFileSync(SHEETS_META_FILE, JSON.stringify(meta, null, 2), "utf-8");
  } catch (err) {
    console.error("Failed to save sheets meta:", err);
  }
}

spreadsheetStore = loadSpreadsheetFromDisk();
sheetsMeta = loadSheetsMetaFromDisk();

/**
 * Helper Google Sheets API v4 untuk membuat dan menyinkronkan Spreadsheet "SIBOS-V1" (20 Sheet)
 */
async function pushAllSheetsToGoogleSpreadsheet(accessToken: string, spreadsheetId: string) {
  if (!accessToken || !spreadsheetId) return;
  ensureAllSheets();

  // 1. Bersihkan nilai lama agar baris sisa tidak tertinggal
  const rangesToClear = ORDERED_SHEETS.map((name) => `'${name.replace(/'/g, "''")}'!A1:Z5000`);
  await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchClear`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ ranges: rangesToClear })
  });

  // 2. Tulis seluruh 20 sheet sekaligus menggunakan values:batchUpdate
  const dataPayload = ORDERED_SHEETS.map((sheetName) => {
    const headers = SHEET_HEADERS[sheetName] || SHEET_HEADERS["PERBULAN_DEFAULT"];
    const rows = spreadsheetStore[sheetName] && spreadsheetStore[sheetName].length > 0
      ? spreadsheetStore[sheetName]
      : [headers];
    return {
      range: `'${sheetName.replace(/'/g, "''")}'!A1`,
      majorDimension: "ROWS",
      values: rows
    };
  });

  const updateRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        valueInputOption: "USER_ENTERED",
        data: dataPayload
      })
    }
  );

  if (!updateRes.ok) {
    const errText = await updateRes.text();
    throw new Error(`Gagal menulis ke Google Sheets SIBOS-V1: ${errText}`);
  }
}

async function pullAllSheetsFromGoogleSpreadsheet(accessToken: string, spreadsheetId: string): Promise<boolean> {
  if (!accessToken || !spreadsheetId) return false;
  ensureAllSheets();

  const rangesQuery = ORDERED_SHEETS.map(
    (name) => `ranges=${encodeURIComponent(`'${name.replace(/'/g, "''")}'!A1:Z5000`)}`
  ).join("&");

  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchGet?${rangesQuery}&valueRenderOption=UNFORMATTED_VALUE`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    }
  );

  if (!res.ok) {
    return false;
  }

  const data = await res.json();
  const valueRanges = Array.isArray(data.valueRanges) ? data.valueRanges : [];
  if (valueRanges.length === 0) return false;

  let updated = false;
  for (let i = 0; i < ORDERED_SHEETS.length; i++) {
    const sheetName = ORDERED_SHEETS[i];
    const headers = SHEET_HEADERS[sheetName] || SHEET_HEADERS["PERBULAN_DEFAULT"];
    const vr = valueRanges[i];
    if (vr && Array.isArray(vr.values) && vr.values.length > 0) {
      spreadsheetStore[sheetName] = vr.values;
      updated = true;
    } else if (!spreadsheetStore[sheetName] || spreadsheetStore[sheetName].length === 0) {
      spreadsheetStore[sheetName] = [headers];
    }
  }

  if (updated) {
    saveSpreadsheetToDisk(spreadsheetStore);
  }
  return updated;
}

async function createOrEnsureSibosGoogleSpreadsheet(accessToken: string, requestedSpreadsheetId?: string): Promise<SheetsMeta> {
  ensureAllSheets();
  const targetId = String(requestedSpreadsheetId || sheetsMeta.spreadsheetId || "").trim();

  if (targetId) {
    // Cek apakah spreadsheetId valid & bisa diakses
    const checkRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${targetId}`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (checkRes.ok) {
      const existingDoc = await checkRes.json();
      const existingSheetTitles = new Set<string>(
        (existingDoc.sheets || []).map((s: any) => String(s?.properties?.title || ""))
      );

      // Tambahkan sheet yang belum ada dari 20 ORDERED_SHEETS
      const addSheetRequests: any[] = [];
      ORDERED_SHEETS.forEach((sName, idx) => {
        if (!existingSheetTitles.has(sName)) {
          addSheetRequests.push({
            addSheet: {
              properties: {
                title: sName,
                index: idx,
                gridProperties: {
                  frozenRowCount: 1
                }
              }
            }
          });
        }
      });

      if (addSheetRequests.length > 0) {
        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${targetId}:batchUpdate`, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({ requests: addSheetRequests })
        });
      }

      // Tarik data yang sudah ada di Google Spreadsheet terlebih dahulu (mis. Password di Role Pengguna)
      await pullAllSheetsFromGoogleSpreadsheet(accessToken, targetId);
      // Pastikan seluruh header & struktur 20 sheet tersinkronisasi
      await pushAllSheetsToGoogleSpreadsheet(accessToken, targetId);

      const meta: SheetsMeta = {
        spreadsheetId: targetId,
        spreadsheetUrl: existingDoc.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${targetId}/edit`,
        spreadsheetTitle: existingDoc.properties?.title || "SIBOS-V1",
        updatedAt: new Date().toISOString()
      };
      saveSheetsMetaToDisk(meta);
      return meta;
    }
  }

  // Buat Spreadsheet baru bernama "SIBOS-V1" lengkap dengan 20 Sheet & format header biru tua
  const sheetsDefinition = ORDERED_SHEETS.map((sheetName, idx) => {
    const headers = SHEET_HEADERS[sheetName] || SHEET_HEADERS["PERBULAN_DEFAULT"];
    return {
      properties: {
        sheetId: idx + 1,
        title: sheetName,
        index: idx,
        gridProperties: {
          rowCount: 1000,
          columnCount: Math.max(headers.length + 2, 20),
          frozenRowCount: 1
        }
      }
    };
  });

  const createRes = await fetch("https://sheets.googleapis.com/v4/spreadsheets", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      properties: {
        title: "SIBOS-V1",
        locale: "id_ID"
      },
      sheets: sheetsDefinition
    })
  });

  if (!createRes.ok) {
    const errText = await createRes.text();
    throw new Error(`Gagal membuat Google Spreadsheet SIBOS-V1: ${errText}`);
  }

  const createdDoc = await createRes.json();
  const newSpreadsheetId = String(createdDoc.spreadsheetId || "");
  const newSpreadsheetUrl = String(
    createdDoc.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${newSpreadsheetId}/edit`
  );

  // Format baris pertama (Header) setiap sheet menjadi biru tua (#1e3a8a) dengan teks putih tebal
  const styleRequests = ORDERED_SHEETS.map((sheetName, idx) => {
    const headers = SHEET_HEADERS[sheetName] || SHEET_HEADERS["PERBULAN_DEFAULT"];
    return {
      repeatCell: {
        range: {
          sheetId: idx + 1,
          startRowIndex: 0,
          endRowIndex: 1,
          startColumnIndex: 0,
          endColumnIndex: headers.length
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 30 / 255, green: 58 / 255, blue: 138 / 255 },
            textFormat: {
              foregroundColor: { red: 1, green: 1, blue: 1 },
              bold: true
            }
          }
        },
        fields: "userEnteredFormat(backgroundColor,textFormat)"
      }
    };
  });

  await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${newSpreadsheetId}:batchUpdate`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ requests: styleRequests })
  });

  // Tulis seluruh 20 sheet beserta data awal ke SIBOS-V1
  await pushAllSheetsToGoogleSpreadsheet(accessToken, newSpreadsheetId);

  const meta: SheetsMeta = {
    spreadsheetId: newSpreadsheetId,
    spreadsheetUrl: newSpreadsheetUrl,
    spreadsheetTitle: "SIBOS-V1",
    updatedAt: new Date().toISOString()
  };
  saveSheetsMetaToDisk(meta);
  return meta;
}

function normalizeKodeKegiatanGs(rawVal: any, level: string, parentStdPrefix: string, isPerBulan: boolean): string {
  if (rawVal === null || rawVal === undefined || rawVal === "") return "";
  if (rawVal instanceof Date) {
    const d = rawVal.getDate();
    const m = rawVal.getMonth() + 1;
    const y = rawVal.getFullYear();
    const dd = d < 10 ? "0" + d : "" + d;
    const mm = m < 10 ? "0" + m : "" + m;

    if (level === "kegiatan" || (parentStdPrefix && !isPerBulan)) {
      const pfx = parentStdPrefix || (y >= 2000 && y <= 2099 ? ("0" + (y - 2000)).slice(-2) : "03");
      return pfx + "." + mm + "." + dd + ".";
    }
    return mm + "." + dd + ".";
  }

  const s = String(rawVal).trim();
  const matchDate = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (matchDate) {
    const p1 = ("0" + matchDate[1]).slice(-2);
    const p2 = ("0" + matchDate[2]).slice(-2);
    const yr = parseInt(matchDate[3], 10);
    if (level === "kegiatan" || parentStdPrefix) {
      const pfx2 = parentStdPrefix || (yr >= 2000 ? ("0" + (yr - 2000)).slice(-2) : "03");
      return pfx2 + "." + p2 + "." + p1 + ".";
    }
    return p2 + "." + p1 + ".";
  }

  return s;
}

function angkaToTerbilangGs(n: any): string {
  const nominal = Math.abs(Math.floor(Number(n) || 0));
  if (nominal === 0) return "Nol Rupiah";
  const satuan = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"];

  function bilang(val: number): string {
    if (val < 12) return satuan[val];
    if (val < 20) return bilang(val - 10) + " Belas";
    if (val < 100) return bilang(Math.floor(val / 10)) + " Puluh " + bilang(val % 10);
    if (val < 200) return "Seratus " + bilang(val - 100);
    if (val < 1000) return bilang(Math.floor(val / 100)) + " Ratus " + bilang(val % 100);
    if (val < 2000) return "Seribu " + bilang(val - 1000);
    if (val < 1000000) return bilang(Math.floor(val / 1000)) + " Ribu " + bilang(val % 1000);
    if (val < 1000000000) return bilang(Math.floor(val / 1000000)) + " Juta " + bilang(val % 1000000);
    if (val < 1000000000000) return bilang(Math.floor(val / 1000000000)) + " Milyar " + bilang(val % 1000000000);
    return bilang(Math.floor(val / 1000000000000)) + " Triliun " + bilang(val % 1000000000000);
  }

  const hasil = bilang(nominal).replace(/\s+/g, " ").trim();
  return hasil ? hasil + " Rupiah" : "Nol Rupiah";
}

function parseIndoDateToTsGs(dateStr: any): number {
  if (!dateStr || typeof dateStr !== "string" || !dateStr.trim()) return 0;
  const parts = dateStr.trim().split("/");
  if (parts.length === 3) {
    const d = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    const y = parseInt(parts[2], 10);
    if (!isNaN(d) && !isNaN(m) && !isNaN(y)) {
      return new Date(y, m, d).getTime();
    }
  }
  return 0;
}

/**
 * Inisialisasi & Pengurutan 20 Sheet tanpa pernah menulis password di dalam kode skrip
 */
function ensureAllSheets() {
  let modified = false;

  for (let i = 0; i < ORDERED_SHEETS.length; i++) {
    const sheetName = ORDERED_SHEETS[i];
    const headers = SHEET_HEADERS[sheetName] || SHEET_HEADERS["PERBULAN_DEFAULT"];

    if (!Array.isArray(spreadsheetStore[sheetName]) || spreadsheetStore[sheetName].length === 0) {
      spreadsheetStore[sheetName] = [headers];
      modified = true;
    }

    // Role Pengguna baris pertama: Kolom PASSWORD (kolom ke-4 / index 3) dikosongkan ("")
    // agar tidak ada password yang ditulis di dalam skrip manapun
    if (sheetName === SHEET_NAMES.USERS && spreadsheetStore[sheetName].length === 1) {
      spreadsheetStore[sheetName].push([
        1,
        "Admin",
        "20231556",
        "",
        "SDIT ANNISA",
        "",
        "",
        "",
        "Kota Bekasi",
        "Jawa Barat",
        "2026",
        "BOS Reguler",
        "",
        "Bank BJB Kantor Cabang Jatiasih",
        "",
        "",
        0,
        new Date().toISOString(),
        "{}"
      ]);
      modified = true;
    }
  }

  if (modified) {
    saveSpreadsheetToDisk(spreadsheetStore);
  }
}

/**
 * Verifikasi Password Admin langsung dari Sheet "Role Pengguna" (Kolom PASSWORD)
 * Tanpa ada password yang di-hardcode di skrip
 */
function apiVerifyAdminLogin(inputPass: any) {
  ensureAllSheets();
  const sheet = spreadsheetStore[SHEET_NAMES.USERS];
  if (!sheet || sheet.length <= 1) {
    return {
      success: false,
      message: "Sheet 'Role Pengguna' masih kosong. Silakan isi password langsung di Spreadsheet SIBOS-V1 pada Sheet 'Role Pengguna' (Kolom PASSWORD)."
    };
  }

  const pClean = String(inputPass || "").trim();
  if (!pClean) {
    return { success: false, message: "Password wajib diisi!" };
  }

  const rows = sheet.slice(1);
  let hasConfiguredPassword = false;

  for (let i = 0; i < rows.length; i++) {
    const rRole = String(rows[i][1] || "").trim();
    const rUser = String(rows[i][2] || "").trim();
    const rPass = String(rows[i][3] || "").trim();

    // Gunakan kolom PASSWORD (kolom D), atau fallback ke kolom C jika kolom D kosong
    const candidatePass = rPass !== "" ? rPass : rUser;
    if (!candidatePass) continue;

    hasConfiguredPassword = true;

    if (pClean === candidatePass) {
      return {
        success: true,
        role: rRole || "Admin"
      };
    }
  }

  if (!hasConfiguredPassword) {
    return {
      success: false,
      message: "Kolom PASSWORD di Sheet 'Role Pengguna' belum diisi. Silakan tulis password langsung di Spreadsheet SIBOS-V1 pada Sheet 'Role Pengguna'."
    };
  }

  return {
    success: false,
    message: "Password Admin salah! Pastikan sesuai dengan password di Sheet 'Role Pengguna' pada Spreadsheet SIBOS-V1."
  };
}

/**
 * Simpan perubahan User & Password Admin langsung ke Sheet "Role Pengguna"
 */
function apiUpdateAdminCredentials(newUser: any, newPassword: any) {
  ensureAllSheets();
  const sheet = spreadsheetStore[SHEET_NAMES.USERS];
  const uClean = String(newUser || "").trim();
  const pClean = String(newPassword || "").trim();
  if (!pClean) {
    return { success: false, message: "Password baru tidak boleh kosong." };
  }

  if (sheet.length > 1) {
    if (uClean) sheet[1][2] = uClean;
    sheet[1][3] = pClean;
    if (!String(sheet[1][1] || "").trim()) {
      sheet[1][1] = "Admin";
    }
  } else {
    sheet.push([
      1, "Admin", uClean || "20231556", pClean, "SDIT ANNISA",
      "", "", "", "Kota Bekasi", "Jawa Barat", "2026",
      "BOS Reguler", "", "Bank BJB Kantor Cabang Jatiasih", "", "", 0, new Date().toISOString(), "{}"
    ]);
  }

  saveSpreadsheetToDisk(spreadsheetStore);
  return { success: true };
}

function apiGetAccounts() {
  ensureAllSheets();
  const sheet = spreadsheetStore[SHEET_NAMES.USERS];
  if (!sheet || sheet.length <= 1) return [];

  const data = sheet.slice(1);
  const accounts: any[] = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const role = String(row[1] || "").trim();
    let userOrNpsn = String(row[2] || "").trim();
    let namaSekolah = String(row[4] || "").trim();
    if (namaSekolah === "PANEL SUPER ADMIN SIBOS" || namaSekolah === "SD NEGERI") namaSekolah = "SDIT ANNISA";
    if (userOrNpsn === "20253978" || userOrNpsn.toLowerCase() === "superadmin") userOrNpsn = "20231556";
    let kepalaSekolah = String(row[5] || "").trim();
    if (kepalaSekolah === "Administrator") kepalaSekolah = "";
    let extraCfg: any = {};
    try { extraCfg = JSON.parse(String(row[18] || "{}")); } catch (e) { extraCfg = {}; }
    let resolvedNpsn = extraCfg.schoolNpsn || (/^\d{6,10}$/.test(userOrNpsn) ? userOrNpsn : "");
    if (resolvedNpsn === "20253978") resolvedNpsn = "20231556";
    if (!userOrNpsn && !namaSekolah && !resolvedNpsn) continue;
    accounts.push({
      npsn: resolvedNpsn || userOrNpsn || "20231556",
      user: userOrNpsn,
      namaSekolah: namaSekolah || "SDIT ANNISA",
      kepalaSekolah: kepalaSekolah || "",
      role: role || "Admin"
    });
  }
  return accounts;
}

function removeSchoolDataByNpsn(targetNpsn: string) {
  if (!targetNpsn) return;
  for (let i = 0; i < ORDERED_SHEETS.length; i++) {
    const sName = ORDERED_SHEETS[i];
    if (sName === SHEET_NAMES.USERS) continue;
    const sh = spreadsheetStore[sName];
    if (!sh || sh.length <= 1) continue;
    const header = sh[0];
    const colCount = header.length;
    const kept = [header];
    for (let r = 1; r < sh.length; r++) {
      if (String(sh[r][colCount - 1] || "").trim() !== targetNpsn) {
        kept.push(sh[r]);
      }
    }
    spreadsheetStore[sName] = kept;
  }
}

function apiSaveAccounts(accountsList: any[], deletedNpsn?: string | null) {
  ensureAllSheets();
  const sheet = spreadsheetStore[SHEET_NAMES.USERS];
  if (sheet && sheet.length > 1 && Array.isArray(accountsList) && accountsList.length > 0) {
    const acc = accountsList[0];
    if (acc) {
      if (acc.namaSekolah) sheet[1][4] = String(acc.namaSekolah);
      if (acc.kepalaSekolah !== undefined) sheet[1][5] = String(acc.kepalaSekolah);
    }
  }
  if (deletedNpsn) {
    removeSchoolDataByNpsn(String(deletedNpsn).trim());
  }
  saveSpreadsheetToDisk(spreadsheetStore);
  return { success: true };
}

function apiGetSchoolData(npsn: any) {
  let cleanNpsn = String(npsn || "").trim();
  if (cleanNpsn === "20253978" || cleanNpsn.toLowerCase() === "superadmin") cleanNpsn = "20231556";

  ensureAllSheets();

  const userSheet = spreadsheetStore[SHEET_NAMES.USERS];
  const headerInfo: any = {
    npsn: cleanNpsn || "20231556",
    namaSekolah: "SDIT ANNISA",
    alamat: "",
    kabupaten: "Kota Bekasi",
    provinsi: "Jawa Barat",
    tahunAnggaran: "2026",
    sumberDana: "BOS Reguler",
    totalPenerimaan: 0,
    kepalaSekolah: "",
    bendaharaSekolah: "",
    noRekening: "",
    namaBank: "Bank BJB Kantor Cabang Jatiasih",
    email: "",
    noHp: ""
  };
  let saldoAwal = 0;
  let updatedAt = "";
  let extraConfig: any = {};

  if (userSheet && userSheet.length > 1) {
    const uRows = userSheet.slice(1);
    let targetRow: any[] | null = null;
    for (let u = 0; u < uRows.length; u++) {
      const rUser = String(uRows[u][2] || "").trim();
      let rExt: any = {};
      try { rExt = JSON.parse(String(uRows[u][18] || "{}")); } catch (e) {}
      if (rUser === cleanNpsn || String(rExt.schoolNpsn || "").trim() === cleanNpsn) {
        targetRow = uRows[u];
        extraConfig = rExt;
        break;
      }
    }
    if (!targetRow && uRows.length > 0) {
      targetRow = uRows[0];
      try { extraConfig = JSON.parse(String(uRows[0][18] || "{}")); } catch (e) {}
    }

    if (targetRow) {
      let namaSek = String(targetRow[4] || "").trim();
      if (!namaSek || namaSek === "PANEL SUPER ADMIN SIBOS" || namaSek === "SD NEGERI") namaSek = "SDIT ANNISA";
      headerInfo.namaSekolah = namaSek;
      headerInfo.kepalaSekolah = String(targetRow[5] || "") === "Administrator" ? "" : String(targetRow[5] || "");
      headerInfo.bendaharaSekolah = String(targetRow[6] || "") === "-" ? "" : String(targetRow[6] || "");
      headerInfo.alamat = String(targetRow[7] || "") === "-" ? "" : String(targetRow[7] || "");
      headerInfo.kabupaten = String(targetRow[8] || "Kota Bekasi");
      headerInfo.provinsi = String(targetRow[9] || "Jawa Barat");
      headerInfo.tahunAnggaran = String(targetRow[10] || "2026");
      headerInfo.sumberDana = String(targetRow[11] || "BOS Reguler");
      headerInfo.noRekening = String(targetRow[12] || "");
      headerInfo.namaBank = String(targetRow[13] || "Bank BJB Kantor Cabang Jatiasih");
      headerInfo.email = String(targetRow[14] || "") === "-" ? "" : String(targetRow[14] || "");
      headerInfo.noHp = String(targetRow[15] || "") === "-" ? "" : String(targetRow[15] || "");
      saldoAwal = Number(targetRow[16] || 0);
      updatedAt = String(targetRow[17] || "");
    }
  }

  const manualPurchasedMap: Record<string, any> = extraConfig.manualPurchasedMap || {};
  const manualKeteranganTextMap: Record<string, any> = extraConfig.manualKeteranganTextMap || {};
  const customLogoUrl = extraConfig.customLogoUrl || "";
  const customKopSuratUrl = extraConfig.customKopSuratUrl || "";

  const masterSheet = spreadsheetStore[SHEET_NAMES.MASTER_RKAS];
  const masterRkas: any[] = [];
  let currentStdPrefix = "";

  if (masterSheet && masterSheet.length > 1) {
    const mRows = masterSheet.slice(1);
    for (let m = 0; m < mRows.length; m++) {
      const mr = mRows[m];
      const uraian = String(mr[3] || "").trim();
      if (!uraian && !mr[1] && !mr[2]) continue;
      const rowNpsn = String(mr[10] || "").trim();
      if (cleanNpsn && rowNpsn && rowNpsn !== cleanNpsn) continue;

      const no = Number(mr[0]) || (masterRkas.length + 1);
      const level = String(mr[7] || "item").trim();
      const kodeKeg = normalizeKodeKegiatanGs(mr[2], level, currentStdPrefix, false);
      const stdMatch = kodeKeg.match(/^(\d{2})\./);
      if (stdMatch) currentStdPrefix = stdMatch[1];

      const statusAksi = String(mr[8] || "").trim();
      const ket = String(mr[9] || "").trim();

      if (statusAksi === "Sudah Dibelanjakan" || statusAksi === "TRUE" || statusAksi === "true") {
        manualPurchasedMap[no] = true;
      }
      if (ket && !manualKeteranganTextMap[no]) {
        manualKeteranganTextMap[no] = ket;
      }

      masterRkas.push({
        no,
        kodeRekening: String(mr[1] || "").trim(),
        kodeKegiatan: kodeKeg,
        uraian,
        jumlah: Number(mr[4] || 0),
        operasi: Number(mr[5] || 0),
        modal: Number(mr[6] || 0),
        level
      });
    }
  }

  const belanjaSheet = spreadsheetStore[SHEET_NAMES.BELANJA];
  const belanjaBankList: any[] = [];
  if (belanjaSheet && belanjaSheet.length > 1) {
    const bRows = belanjaSheet.slice(1);
    for (let b = 0; b < bRows.length; b++) {
      const br = bRows[b];
      const namaBelanja = String(br[1] || "").trim();
      if (!namaBelanja) continue;
      const bNpsn = String(br[14] || "").trim();
      if (cleanNpsn && bNpsn && bNpsn !== cleanNpsn) continue;

      let subItems: any = undefined;
      try {
        if (br[11] && String(br[11]).trim() !== "" && String(br[11]).trim() !== "[]") {
          subItems = JSON.parse(String(br[11]));
        }
      } catch (e) {}

      belanjaBankList.push({
        id: br[13] ? (isNaN(Number(br[13])) ? String(br[13]) : Number(br[13])) : (Date.now() + b),
        belanja: namaBelanja,
        noRek: String(br[2] || "-"),
        namaBank: String(br[3] || "-"),
        atasNama: String(br[4] || "-"),
        npwp: String(br[5] || "-"),
        total: Number(br[6] || 0),
        sumberDana: String(br[7] || "BOS REGULER"),
        tglTransaksiSI: String(br[8] || "-"),
        bulan: String(br[9] || "Januari"),
        masterNo: br[10] !== "" && br[10] !== null ? br[10] : null,
        subItems,
        tglInput: String(br[12] || "")
      });
    }
  }

  const danaMasukList: any[] = Array.isArray(extraConfig.danaMasukList) ? extraConfig.danaMasukList : [];
  if (danaMasukList.length === 0) {
    const rkSheet = spreadsheetStore[SHEET_NAMES.REKENING_KORAN];
    if (rkSheet && rkSheet.length > 1) {
      const rkRows = rkSheet.slice(1);
      for (let k = 0; k < rkRows.length; k++) {
        const rkr = rkRows[k];
        const rkrNpsn = String(rkr[8] || "").trim();
        if (cleanNpsn && rkrNpsn && rkrNpsn !== cleanNpsn) continue;
        const jenis = String(rkr[5] || "").trim();
        if (jenis === "SALDO_AWAL" && saldoAwal === 0) {
          saldoAwal = Number(rkr[3] || rkr[4] || 0);
        } else if (jenis === "DANA_MASUK" && Number(rkr[3] || 0) > 0) {
          danaMasukList.push({
            id: String(rkr[7] || ("dm-" + (k + 1))),
            tgl: String(rkr[1] || ""),
            kredit: Number(rkr[3] || 0),
            keterangan: String(rkr[6] || "Dana Masuk")
          });
        }
      }
    }
  }

  const perBulanData: Record<string, any[]> = {};
  for (let mIdx = 0; mIdx < SHEET_NAMES.MONTHS.length; mIdx++) {
    const mName = SHEET_NAMES.MONTHS[mIdx];
    perBulanData[mName] = [];
    const mSheet = spreadsheetStore[mName];
    let mStdPrefix = "";
    if (mSheet && mSheet.length > 1) {
      const pRows = mSheet.slice(1);
      for (let p = 0; p < pRows.length; p++) {
        const pr = pRows[p];
        const pUraian = String(pr[3] || "").trim();
        if (!pUraian && !pr[1] && !pr[2]) continue;
        const pNpsn = String(pr[9] || "").trim();
        if (cleanNpsn && pNpsn && pNpsn !== cleanNpsn) continue;

        const pLevel = String(pr[8] || "item");
        const pKodeProg = normalizeKodeKegiatanGs(pr[2], pLevel, mStdPrefix, true);
        const pStdMatch = pKodeProg.match(/^(\d{2})\./);
        if (pStdMatch) mStdPrefix = pStdMatch[1];

        perBulanData[mName].push({
          no: Number(pr[0]) || (perBulanData[mName].length + 1),
          kodeRekening: String(pr[1] || "").trim(),
          kodeProgram: pKodeProg,
          uraian: pUraian,
          volume: pr[4] !== null && pr[4] !== undefined && pr[4] !== "" ? String(pr[4]) : "",
          satuan: String(pr[5] || ""),
          tarif: pr[6] !== null && pr[6] !== undefined && pr[6] !== "" ? Number(pr[6]) : "",
          jumlah: Number(pr[7] || 0),
          level: pLevel
        });
      }
    }
  }

  const distSheet = spreadsheetStore[SHEET_NAMES.DISTRIBUSI];
  const distribusiList: any[] = [];
  if (distSheet && distSheet.length > 1) {
    const dRows = distSheet.slice(1);
    for (let d = 0; d < dRows.length; d++) {
      const dr = dRows[d];
      const namaBrg = String(dr[5] || "").trim();
      if (!namaBrg) continue;
      const dNpsn = String(dr[10] || "").trim();
      if (cleanNpsn && dNpsn && dNpsn !== cleanNpsn) continue;

      let tglDist = dr[1];
      if (tglDist instanceof Date) {
        const yyyy = tglDist.getFullYear();
        const mm = ("0" + (tglDist.getMonth() + 1)).slice(-2);
        const dd = ("0" + tglDist.getDate()).slice(-2);
        tglDist = yyyy + "-" + mm + "-" + dd;
      } else {
        tglDist = String(tglDist || "").trim();
      }

      distribusiList.push({
        id: String(dr[9] || ("dist-" + (d + 1))),
        tanggal: tglDist,
        masterNo: dr[2] !== null && dr[2] !== undefined ? String(dr[2]) : "",
        kodeRekening: String(dr[3] || "").trim(),
        kodeKegiatan: normalizeKodeKegiatanGs(dr[4], "item", "", false),
        namaBarang: namaBrg,
        jumlahKeluar: Number(dr[6] || 0),
        penerima: String(dr[7] || "").trim(),
        keterangan: String(dr[8] || "").trim()
      });
    }
  }

  return {
    updatedAt,
    headerInfo,
    masterRkas,
    perBulanData,
    belanjaBankList,
    danaMasukList,
    distribusiList,
    saldoAwal,
    manualPurchasedMap,
    manualKeteranganTextMap,
    customLogoUrl,
    customKopSuratUrl
  };
}

function replaceRowsForNpsn(sheetName: string, totalCols: number, targetNpsn: string, newRows: any[][]) {
  const headers = SHEET_HEADERS[sheetName] || SHEET_HEADERS["PERBULAN_DEFAULT"];
  const current = spreadsheetStore[sheetName] || [headers];
  const kept = [current[0] || headers];

  for (let r = 1; r < current.length; r++) {
    const rowNpsn = String(current[r][totalCols - 1] || "").trim();
    if (rowNpsn && rowNpsn !== targetNpsn) {
      kept.push(current[r]);
    }
  }

  const combined = kept.concat(newRows || []);
  spreadsheetStore[sheetName] = combined;
}

function apiSaveSchoolData(npsn: any, payload: any) {
  let cleanNpsn = String(npsn || (payload && payload.headerInfo && payload.headerInfo.npsn) || "20231556").trim();
  if (cleanNpsn === "20253978" || cleanNpsn.toLowerCase() === "superadmin") cleanNpsn = "20231556";

  ensureAllSheets();

  const nowIso = (payload && payload.updatedAt) || new Date().toISOString();
  const hInfo = (payload && payload.headerInfo) || {};
  const masterRkas = Array.isArray(payload && payload.masterRkas) ? payload.masterRkas : [];
  const belanjaList = Array.isArray(payload && payload.belanjaBankList) ? payload.belanjaBankList : [];
  const danaMasukList = Array.isArray(payload && payload.danaMasukList) ? payload.danaMasukList : [];
  const distribusiList = Array.isArray(payload && payload.distribusiList) ? payload.distribusiList : [];
  const perBulanData = (payload && payload.perBulanData) || {};
  const saldoAwal = Number((payload && payload.saldoAwal) || 0);
  const manualPurchasedMap = (payload && payload.manualPurchasedMap) || {};
  const manualKeteranganTextMap = (payload && payload.manualKeteranganTextMap) || {};
  const customLogoUrl = String((payload && payload.customLogoUrl) || "");
  const customKopSuratUrl = String((payload && payload.customKopSuratUrl) || "");

  const realizedByNo: Record<string, number> = {};
  const realizedByUraian: Record<string, number> = {};
  for (let bIdx = 0; bIdx < belanjaList.length; bIdx++) {
    const bItem = belanjaList[bIdx];
    const bTotal = Number(bItem.total || 0);
    if (bItem.masterNo) realizedByNo[String(bItem.masterNo)] = (realizedByNo[String(bItem.masterNo)] || 0) + bTotal;
    if (bItem.belanja) {
      const uKey = String(bItem.belanja).trim().toLowerCase();
      realizedByUraian[uKey] = (realizedByUraian[uKey] || 0) + bTotal;
    }
    if (Array.isArray(bItem.subItems)) {
      for (let sIdx = 0; sIdx < bItem.subItems.length; sIdx++) {
        const sub = bItem.subItems[sIdx];
        if (sub && sub.no) realizedByNo[String(sub.no)] = (realizedByNo[String(sub.no)] || 0) + Number(sub.jumlah || 0);
      }
    }
  }

  // 1. Sheet 1: Database Master RKAS
  const masterRows: any[][] = [];
  let mStdPrefix = "";
  for (let i = 0; i < masterRkas.length; i++) {
    const m = masterRkas[i];
    const mNo = m.no !== undefined ? m.no : (i + 1);
    const mLevel = String(m.level || "item");
    const cleanKodeKeg = normalizeKodeKegiatanGs(m.kodeKegiatan, mLevel, mStdPrefix, false);
    const stdMatch = cleanKodeKeg.match(/^(\d{2})\./);
    if (stdMatch) mStdPrefix = stdMatch[1];

    const isBought = !!manualPurchasedMap[mNo] || (realizedByNo[String(mNo)] > 0) || (realizedByUraian[String(m.uraian || "").trim().toLowerCase()] > 0);
    masterRows.push([
      mNo, String(m.kodeRekening || ""), cleanKodeKeg, String(m.uraian || ""),
      Number(m.jumlah || 0), Number(m.operasi || 0), Number(m.modal || 0), mLevel,
      isBought ? "Sudah Dibelanjakan" : "Belum Dibelanjakan", String(manualKeteranganTextMap[mNo] || ""), cleanNpsn
    ]);
  }
  replaceRowsForNpsn(SHEET_NAMES.MASTER_RKAS, 11, cleanNpsn, masterRows);

  // 2. Sheet 2: Database Belanja
  const belanjaRows: any[][] = [];
  for (let j = 0; j < belanjaList.length; j++) {
    const b = belanjaList[j];
    belanjaRows.push([
      j + 1, String(b.belanja || ""), String(b.noRek || "-"), String(b.namaBank || "-"),
      String(b.atasNama || "-"), String(b.npwp || "-"), Number(b.total || 0),
      String(b.sumberDana || "BOS REGULER"), String(b.tglTransaksiSI || "-"), String(b.bulan || "Januari"),
      b.masterNo !== null && b.masterNo !== undefined ? String(b.masterNo) : "",
      Array.isArray(b.subItems) && b.subItems.length > 0 ? JSON.stringify(b.subItems) : "",
      String(b.tglInput || ""), String(b.id || (j + 1)), cleanNpsn
    ]);
  }
  replaceRowsForNpsn(SHEET_NAMES.BELANJA, 15, cleanNpsn, belanjaRows);

  // 3. Sheet 3: SI
  const siRows: any[][] = [];
  for (let k = 0; k < belanjaList.length; k++) {
    const sb = belanjaList[k];
    siRows.push([
      k + 1, String(sb.belanja || ""), String(sb.noRek || "-"), String(sb.namaBank || "-"),
      String(sb.atasNama || "-"), String(sb.npwp || "-"), Number(sb.total || 0),
      String(sb.sumberDana || "BOS REGULER"), String(sb.tglTransaksiSI || "-"), String(sb.bulan || "Januari"), cleanNpsn
    ]);
  }
  replaceRowsForNpsn(SHEET_NAMES.SI, 11, cleanNpsn, siRows);

  // 4. Sheet 4: Surat Transfer
  const suratRows: any[][] = [];
  const totalSemuaBelanja = belanjaList.reduce((acc: number, it: any) => acc + Number(it.total || 0), 0);
  const byTglMap: Record<string, number> = {};
  for (let st = 0; st < belanjaList.length; st++) {
    const tglKey = String(belanjaList[st].tglTransaksiSI || "Semua").trim();
    byTglMap[tglKey] = (byTglMap[tglKey] || 0) + Number(belanjaList[st].total || 0);
  }
  const tglKeys = Object.keys(byTglMap);
  if (tglKeys.length === 0) {
    suratRows.push([
      1, (hInfo.kabupaten || "Bekasi") + ", " + new Date().toLocaleDateString("id-ID"),
      "421.2/    /" + (hInfo.namaSekolah || "SEKOLAH") + "/" + (hInfo.tahunAnggaran || "2026"),
      "Semua", String(hInfo.noRekening || ""), String(hInfo.namaSekolah || "SDIT ANNISA"),
      String(hInfo.namaSekolah || "SDIT ANNISA"), String(hInfo.alamat || ""),
      String(hInfo.kepalaSekolah || ""), String(hInfo.bendaharaSekolah || ""),
      String(hInfo.noHp || ""), String(hInfo.email || ""),
      String(hInfo.namaBank || "Bank BJB Kantor Cabang Jatiasih"),
      totalSemuaBelanja, angkaToTerbilangGs(totalSemuaBelanja), cleanNpsn
    ]);
  } else {
    for (let tk = 0; tk < tglKeys.length; tk++) {
      const tglStr = tglKeys[tk];
      const nomTanggal = byTglMap[tglStr];
      suratRows.push([
        tk + 1, (hInfo.kabupaten || "Bekasi") + ", " + tglStr,
        "421.2/00" + (tk + 1) + "/SI-BOS/" + (hInfo.tahunAnggaran || "2026"),
        tglStr, String(hInfo.noRekening || ""), String(hInfo.namaSekolah || "SDIT ANNISA"),
        String(hInfo.namaSekolah || "SDIT ANNISA"), String(hInfo.alamat || ""),
        String(hInfo.kepalaSekolah || ""), String(hInfo.bendaharaSekolah || ""),
        String(hInfo.noHp || ""), String(hInfo.email || ""),
        String(hInfo.namaBank || "Bank BJB Kantor Cabang Jatiasih"),
        nomTanggal, angkaToTerbilangGs(nomTanggal), cleanNpsn
      ]);
    }
  }
  replaceRowsForNpsn(SHEET_NAMES.SURAT_TRANSFER, 16, cleanNpsn, suratRows);

  // 5. Sheet 5: Rekap Serapan
  const rekapRows: any[][] = [];
  let rekapStdPrefix = "";
  for (let rk = 0; rk < masterRkas.length; rk++) {
    const rm = masterRkas[rk];
    const rmLevel = String(rm.level || "item");
    const rmKodeKeg = normalizeKodeKegiatanGs(rm.kodeKegiatan, rmLevel, rekapStdPrefix, false);
    const rmStdMatch = rmKodeKeg.match(/^(\d{2})\./);
    if (rmStdMatch) rekapStdPrefix = rmStdMatch[1];

    const relBank = (realizedByNo[String(rm.no)] || 0) + (realizedByUraian[String(rm.uraian || "").trim().toLowerCase()] || 0);
    const isManual = !!manualPurchasedMap[rm.no];
    const paguVal = Number(rm.jumlah || 0);
    const relFinal = isManual ? (relBank > 0 ? relBank : paguVal) : relBank;
    const sisaVal = Math.max(0, paguVal - relFinal);
    const isBought = isManual || relBank > 0;
    rekapRows.push([
      rm.no !== undefined ? rm.no : (rk + 1), String(rm.kodeRekening || ""), rmKodeKeg,
      String(rm.uraian || ""), paguVal, relFinal, sisaVal,
      isBought ? (isManual && relBank === 0 ? "Dibelanjakan (Manual)" : "Dibelanjakan") : "Belum",
      String(manualKeteranganTextMap[rm.no] || "-"), cleanNpsn
    ]);
  }
  replaceRowsForNpsn(SHEET_NAMES.REKAP_SERAPAN, 10, cleanNpsn, rekapRows);

  // 6. Sheet 6: Rekening Koran
  const rekKoranRows: any[][] = [[1, "", "", saldoAwal, saldoAwal, "SALDO_AWAL", "Saldo Awal Rekening", "row-saldo-awal", cleanNpsn]];
  let runningBal = saldoAwal;
  const dateGroupMap: Record<string, { tgl: string; debit: number; kredit: number; ket: string[] }> = {};
  for (let dmIdx = 0; dmIdx < danaMasukList.length; dmIdx++) {
    const dm = danaMasukList[dmIdx];
    const tglDmStr = String(dm.tgl || "").trim();
    if (!tglDmStr) continue;
    if (!dateGroupMap[tglDmStr]) dateGroupMap[tglDmStr] = { tgl: tglDmStr, debit: 0, kredit: 0, ket: [] };
    dateGroupMap[tglDmStr].kredit += Number(dm.kredit || 0);
    if (dm.keterangan) dateGroupMap[tglDmStr].ket.push(dm.keterangan);
  }
  for (let blIdx = 0; blIdx < belanjaList.length; blIdx++) {
    const bl = belanjaList[blIdx];
    const tglBlStr = String(bl.tglTransaksiSI || "01/01/2026").trim();
    if (!tglBlStr) continue;
    if (!dateGroupMap[tglBlStr]) dateGroupMap[tglBlStr] = { tgl: tglBlStr, debit: 0, kredit: 0, ket: [] };
    dateGroupMap[tglBlStr].debit += Number(bl.total || 0);
  }
  const sortedDates = Object.keys(dateGroupMap).map((k) => dateGroupMap[k]);
  sortedDates.sort((a, b) => parseIndoDateToTsGs(a.tgl) - parseIndoDateToTsGs(b.tgl));
  for (let sd = 0; sd < sortedDates.length; sd++) {
    const grp = sortedDates[sd];
    runningBal = runningBal + grp.kredit - grp.debit;
    const jenisTx = (grp.kredit > 0 && grp.debit === 0) ? "DANA_MASUK" : (grp.debit > 0 && grp.kredit === 0 ? "BELANJA_SI" : "MUTASI_GABUNGAN");
    rekKoranRows.push([
      sd + 2, grp.tgl, grp.debit > 0 ? grp.debit : "", grp.kredit > 0 ? grp.kredit : "",
      runningBal, jenisTx, grp.ket.length > 0 ? grp.ket.join("; ") : "Mutasi Transaksi Bank (SI)",
      "date-group-" + grp.tgl, cleanNpsn
    ]);
  }
  replaceRowsForNpsn(SHEET_NAMES.REKENING_KORAN, 9, cleanNpsn, rekKoranRows);

  // 7. Sheet 7 s/d 18: Perbulan
  for (let mIdx = 0; mIdx < SHEET_NAMES.MONTHS.length; mIdx++) {
    const monthName = SHEET_NAMES.MONTHS[mIdx];
    const monthArr = Array.isArray(perBulanData[monthName]) ? perBulanData[monthName] : [];
    const monthRows: any[][] = [];
    let mMonthStdPrefix = "";
    for (let rIdx = 0; rIdx < monthArr.length; rIdx++) {
      const mrItem = monthArr[rIdx];
      const mrLevel = String(mrItem.level || "item");
      const cleanKodeProg = normalizeKodeKegiatanGs(mrItem.kodeProgram || mrItem.kodeKegiatan || "", mrLevel, mMonthStdPrefix, true);
      const mpStdMatch = cleanKodeProg.match(/^(\d{2})\./);
      if (mpStdMatch) mMonthStdPrefix = mpStdMatch[1];

      monthRows.push([
        mrItem.no !== undefined ? mrItem.no : (rIdx + 1), String(mrItem.kodeRekening || ""),
        cleanKodeProg, String(mrItem.uraian || ""),
        mrItem.volume !== undefined && mrItem.volume !== null ? String(mrItem.volume) : "",
        String(mrItem.satuan || ""),
        mrItem.tarif !== undefined && mrItem.tarif !== null && mrItem.tarif !== "" ? Number(mrItem.tarif) : "",
        Number(mrItem.jumlah || 0), mrLevel, cleanNpsn
      ]);
    }
    replaceRowsForNpsn(monthName, 10, cleanNpsn, monthRows);
  }

  // 8. Sheet 20: Realisasi Distribusi
  const distRows: any[][] = [];
  for (let dIdx = 0; dIdx < distribusiList.length; dIdx++) {
    const dItem = distribusiList[dIdx];
    distRows.push([
      dIdx + 1,
      String(dItem.tanggal || ""),
      dItem.masterNo !== undefined && dItem.masterNo !== null ? String(dItem.masterNo) : "",
      String(dItem.kodeRekening || ""),
      normalizeKodeKegiatanGs(dItem.kodeKegiatan || "", "item", "", false),
      String(dItem.namaBarang || ""),
      Number(dItem.jumlahKeluar || 0),
      String(dItem.penerima || ""),
      String(dItem.keterangan || ""),
      String(dItem.id || (dIdx + 1)),
      cleanNpsn
    ]);
  }
  replaceRowsForNpsn(SHEET_NAMES.DISTRIBUSI, 11, cleanNpsn, distRows);

  // 9. Sheet 19: Role Pengguna
  const userSheet = spreadsheetStore[SHEET_NAMES.USERS];
  let foundUser = false;
  const extraJsonStr = JSON.stringify({
    schoolNpsn: cleanNpsn,
    manualPurchasedMap,
    manualKeteranganTextMap,
    danaMasukList,
    customLogoUrl,
    customKopSuratUrl
  });

  if (userSheet && userSheet.length > 1) {
    let targetIdx = -1;
    for (let u = 1; u < userSheet.length; u++) {
      const rUserOrNpsn = String(userSheet[u][2] || "").trim();
      let rExtra: any = {};
      try { rExtra = JSON.parse(String(userSheet[u][18] || "{}")); } catch (e) { rExtra = {}; }
      if (rUserOrNpsn === cleanNpsn || String(rExtra.schoolNpsn || "").trim() === cleanNpsn) {
        targetIdx = u;
        break;
      }
    }
    if (targetIdx === -1 && userSheet.length > 1) {
      targetIdx = 1;
    }
    if (targetIdx !== -1) {
      // JANGAN menimpa kolom ROLE (index 1), USER / NPSN (index 2), dan PASSWORD (index 3)
      // agar User & Password Admin yang tersimpan di Sheet Role Pengguna tetap utuh.
      userSheet[targetIdx][4] = String(hInfo.namaSekolah || userSheet[targetIdx][4] || "SDIT ANNISA");
      userSheet[targetIdx][5] = String(hInfo.kepalaSekolah || userSheet[targetIdx][5] || "");
      userSheet[targetIdx][6] = String(hInfo.bendaharaSekolah || "");
      userSheet[targetIdx][7] = String(hInfo.alamat || "");
      userSheet[targetIdx][8] = String(hInfo.kabupaten || "Kota Bekasi");
      userSheet[targetIdx][9] = String(hInfo.provinsi || "Jawa Barat");
      userSheet[targetIdx][10] = String(hInfo.tahunAnggaran || "2026");
      userSheet[targetIdx][11] = String(hInfo.sumberDana || "BOS Reguler");
      userSheet[targetIdx][12] = String(hInfo.noRekening || "");
      userSheet[targetIdx][13] = String(hInfo.namaBank || "Bank BJB Kantor Cabang Jatiasih");
      userSheet[targetIdx][14] = String(hInfo.email || "");
      userSheet[targetIdx][15] = String(hInfo.noHp || "");
      userSheet[targetIdx][16] = Number(saldoAwal || 0);
      userSheet[targetIdx][17] = nowIso;
      userSheet[targetIdx][18] = extraJsonStr;
      foundUser = true;
    }
  }

  if (!foundUser) {
    // Kolom PASSWORD (index 3) dikosongkan ("") — tidak ada password yang ditulis di dalam skrip
    userSheet.push([
      userSheet.length, "Admin", cleanNpsn || "20231556", "",
      String(hInfo.namaSekolah || "SDIT ANNISA"), String(hInfo.kepalaSekolah || ""),
      String(hInfo.bendaharaSekolah || ""), String(hInfo.alamat || ""),
      String(hInfo.kabupaten || "Kota Bekasi"), String(hInfo.provinsi || "Jawa Barat"),
      String(hInfo.tahunAnggaran || "2026"), String(hInfo.sumberDana || "BOS Reguler"),
      String(hInfo.noRekening || ""), String(hInfo.namaBank || "Bank BJB Kantor Cabang Jatiasih"),
      String(hInfo.email || ""), String(hInfo.noHp || ""), Number(saldoAwal || 0), nowIso, extraJsonStr
    ]);
  }

  saveSpreadsheetToDisk(spreadsheetStore);
  return { success: true, updatedAt: nowIso, npsn: cleanNpsn };
}

function apiGetLastUpdatedAt(npsn?: any) {
  ensureAllSheets();
  const userSheet = spreadsheetStore[SHEET_NAMES.USERS];
  if (!userSheet || userSheet.length <= 1) return "";
  return String(userSheet[1][17] || "");
}

function extractBearerToken(req: express.Request): string {
  const authHeader = String(req.headers.authorization || "");
  if (authHeader.startsWith("Bearer ")) {
    return authHeader.slice(7).trim();
  }
  return "";
}

async function startServer() {
  ensureAllSheets();

  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: "50mb" }));

  // Status Spreadsheet SIBOS-V1
  app.get("/api/sheets/status", (_req, res) => {
    res.json({
      ok: true,
      spreadsheetId: sheetsMeta.spreadsheetId,
      spreadsheetUrl: sheetsMeta.spreadsheetUrl,
      spreadsheetTitle: sheetsMeta.spreadsheetTitle || "SIBOS-V1",
      updatedAt: sheetsMeta.updatedAt
    });
  });

  // Buat atau Hubungkan Spreadsheet "SIBOS-V1" (20 Sheet) di Google Sheets pengguna
  app.post("/api/sheets/create-or-connect", async (req, res) => {
    try {
      const accessToken = extractBearerToken(req);
      if (!accessToken) {
        return res.status(401).json({ ok: false, error: "Token Google Sheets tidak ditemukan. Silakan Sign in with Google terlebih dahulu." });
      }
      const { spreadsheetId } = req.body || {};
      const meta = await createOrEnsureSibosGoogleSpreadsheet(accessToken, spreadsheetId);
      return res.json({
        ok: true,
        ...meta,
        schoolData: apiGetSchoolData("20231556")
      });
    } catch (err: any) {
      console.error("Create/Connect SIBOS-V1 error:", err);
      return res.status(500).json({ ok: false, error: err?.message || "Gagal membuat/menghubungkan Spreadsheet SIBOS-V1." });
    }
  });

  // Muat (Pull) seluruh 20 Sheet dari Google Spreadsheet SIBOS-V1
  app.post("/api/sheets/pull", async (req, res) => {
    try {
      const accessToken = extractBearerToken(req);
      if (!accessToken) {
        return res.status(401).json({ ok: false, error: "Silakan Sign in with Google terlebih dahulu." });
      }
      if (!sheetsMeta.spreadsheetId) {
        return res.status(400).json({ ok: false, error: "Spreadsheet SIBOS-V1 belum dibuat." });
      }
      await pullAllSheetsFromGoogleSpreadsheet(accessToken, sheetsMeta.spreadsheetId);
      const npsn = req.body?.npsn || "20231556";
      return res.json({
        ok: true,
        schoolData: apiGetSchoolData(npsn)
      });
    } catch (err: any) {
      console.error("Pull SIBOS-V1 error:", err);
      return res.status(500).json({ ok: false, error: err?.message || "Gagal memuat data dari Spreadsheet SIBOS-V1." });
    }
  });

  // Simpan (Push) seluruh 20 Sheet ke Google Spreadsheet SIBOS-V1
  app.post("/api/sheets/push", async (req, res) => {
    try {
      const accessToken = extractBearerToken(req);
      if (!accessToken) {
        return res.status(401).json({ ok: false, error: "Silakan Sign in with Google terlebih dahulu." });
      }
      if (!sheetsMeta.spreadsheetId) {
        const meta = await createOrEnsureSibosGoogleSpreadsheet(accessToken);
        return res.json({ ok: true, ...meta });
      }
      await pushAllSheetsToGoogleSpreadsheet(accessToken, sheetsMeta.spreadsheetId);
      saveSheetsMetaToDisk({
        ...sheetsMeta,
        updatedAt: new Date().toISOString()
      });
      return res.json({
        ok: true,
        ...sheetsMeta
      });
    } catch (err: any) {
      console.error("Push SIBOS-V1 error:", err);
      return res.status(500).json({ ok: false, error: err?.message || "Gagal menyimpan ke Spreadsheet SIBOS-V1." });
    }
  });

  // Endpoint RPC untuk menjalankan seluruh fungsi Kode.gs pada 20 Sheet Spreadsheet Cloud & Google Sheets SIBOS-V1
  app.post("/api/gas", async (req, res) => {
    try {
      const { fnName, args = [], syncGoogleSheets = false } = req.body || {};
      const accessToken = extractBearerToken(req);

      // Jika sudah terhubung ke Google Spreadsheet SIBOS-V1 dan ada token, tarik data terbaru saat membaca/login
      if (
        accessToken &&
        sheetsMeta.spreadsheetId &&
        (fnName === "apiVerifyAdminLogin" || fnName === "apiGetSchoolData" || fnName === "apiGetAccounts")
      ) {
        try {
          await pullAllSheetsFromGoogleSpreadsheet(accessToken, sheetsMeta.spreadsheetId);
        } catch (pullErr) {
          console.warn("Non-fatal pull warning:", pullErr);
        }
      }

      switch (fnName) {
        case "apiVerifyAdminLogin":
          return res.json({ ok: true, result: apiVerifyAdminLogin(args[0]) });
        case "apiUpdateAdminCredentials": {
          const r = apiUpdateAdminCredentials(args[0], args[1]);
          if (accessToken && sheetsMeta.spreadsheetId && syncGoogleSheets) {
            await pushAllSheetsToGoogleSpreadsheet(accessToken, sheetsMeta.spreadsheetId);
          }
          return res.json({ ok: true, result: r });
        }
        case "apiGetAccounts":
          return res.json({ ok: true, result: apiGetAccounts() });
        case "apiSaveAccounts": {
          const r = apiSaveAccounts(args[0], args[1]);
          if (accessToken && sheetsMeta.spreadsheetId && syncGoogleSheets) {
            await pushAllSheetsToGoogleSpreadsheet(accessToken, sheetsMeta.spreadsheetId);
          }
          return res.json({ ok: true, result: r });
        }
        case "apiGetSchoolData":
          return res.json({ ok: true, result: apiGetSchoolData(args[0]) });
        case "apiSaveSchoolData": {
          const r = apiSaveSchoolData(args[0], args[1]);
          if (accessToken && sheetsMeta.spreadsheetId && syncGoogleSheets) {
            await pushAllSheetsToGoogleSpreadsheet(accessToken, sheetsMeta.spreadsheetId);
          }
          return res.json({ ok: true, result: r });
        }
        case "apiGetLastUpdatedAt":
          return res.json({ ok: true, result: apiGetLastUpdatedAt(args[0]) });
        case "ensureAllSheets":
          ensureAllSheets();
          return res.json({ ok: true, result: { success: true } });
        default:
          return res.status(400).json({ ok: false, error: `Unknown function: ${fnName}` });
      }
    } catch (err: any) {
      console.error("GAS RPC Error:", err);
      return res.status(500).json({ ok: false, error: err?.message || "Internal server error" });
    }
  });

  if (process.env.NODE_ENV !== "production") {
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(__dirname, "dist");
    app.use(express.static(distPath));
    app.get("*all", (_req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`SiBOS V1 Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
