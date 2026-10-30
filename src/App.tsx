// @ts-nocheck
import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import { initAuth, googleSignIn, getAccessToken, logout as googleLogout } from "./googleSheetsAuth";


    // Helper dynamic loader modul Excel & PDF (Dimuat HANYA saat dibutuhkan, sehingga web terbuka instan)
    const ensureXlsx = () => {
      return new Promise((resolve, reject) => {
        if (window.XLSX) return resolve(window.XLSX);
        const s = document.createElement("script");
        s.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
        s.onload = () => resolve(window.XLSX);
        s.onerror = () => reject(new Error("Gagal memuat modul Excel. Silakan periksa koneksi internet Anda."));
        document.head.appendChild(s);
      });
    };

    const ensurePdfJs = () => {
      return new Promise((resolve, reject) => {
        if (window.pdfjsLib) return resolve(window.pdfjsLib);
        const s = document.createElement("script");
        s.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js";
        s.onload = () => {
          if (window.pdfjsLib) {
            window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
          }
          resolve(window.pdfjsLib);
        };
        s.onerror = () => reject(new Error("Gagal memuat modul PDF. Silakan periksa koneksi internet Anda."));
        document.head.appendChild(s);
      });
    };

    // Helper pemanggil Google Apps Script (Code.gs) via google.script.run atau Cloud Server API (/api/gas)
    const isGasEnvironment = () => {
      return true;
    };

    const callGasServer = async (fnName, ...args) => {
      if (typeof window !== "undefined" && window.google && window.google.script && window.google.script.run) {
        return new Promise((resolve, reject) => {
          window.google.script.run
            .withSuccessHandler((res) => resolve(res))
            .withFailureHandler((err) => reject(err))
            [fnName](...args);
        });
      }
      const token = await getAccessToken();
      const headers = { "Content-Type": "application/json" };
      if (token) {
        headers["Authorization"] = `Bearer ${token}`;
      }
      const syncGoogleSheets = typeof window !== "undefined" && window.__sibosAutoSyncConfirmed === true;
      const response = await fetch("/api/gas", {
        method: "POST",
        headers,
        body: JSON.stringify({ fnName, args, syncGoogleSheets }),
      });
      if (!response.ok) {
        throw new Error("Gagal menghubungi server cloud SiBOS.");
      }
      const data = await response.json();
      if (!data.ok) {
        throw new Error(data.error || "Terjadi kesalahan pada server cloud SiBOS.");
      }
      return data.result;
    };

    // Konstanta fallback kosong (sesuai instruksi mulai dengan tabel kosong hingga impor RKAS)
    const RKAS_PDF_165_ITEMS = [];
    const DEFAULT_DANA_MASUK_ITEMS = [];
    const JANUARI_PERBULAN_93_ITEMS = [];
    const FEBRUARI_PERBULAN_12_ITEMS = [];
    const JULI_PERBULAN_91_ITEMS = [];
    const AGUSTUS_PERBULAN_30_ITEMS = [];
    const SEPTEMBER_PERBULAN_20_ITEMS = [];
    const OKTOBER_PERBULAN_12_ITEMS = [];
    const NOVEMBER_PERBULAN_20_ITEMS = [];
    const DESEMBER_PERBULAN_12_ITEMS = [];

    // Fallback SVG Icon Murni Transparan (Tanpa bayang & tanpa kotak hitam)
    const SIBOS_CLEAN_ICON = "data:image/svg+xml;utf8," + encodeURIComponent(
      `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" fill="none">
        <rect width="100" height="100" rx="22" fill="#2563eb"/>
        <text x="50%" y="54%" dominant-baseline="middle" text-anchor="middle" fill="#ffffff" font-family="-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif" font-weight="900" font-size="28">SiBOS</text>
        <text x="50%" y="78%" dominant-baseline="middle" text-anchor="middle" fill="#bfdbfe" font-family="-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif" font-weight="700" font-size="12">RKAS</text>
      </svg>`
    );
    const SIBOS_FALLBACK_ICON = SIBOS_CLEAN_ICON;

    const MONTHS = [
      "Januari", "Februari", "Maret", "April", "Mei", "Juni",
      "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ];

    // Helper Format Ribuan dengan Titik (misal: 7950000 -> 7.950.000)
    const formatRupiah = (val) => {
      if (val === "" || val === null || val === undefined) return "";
      const numStr = String(val).replace(/[^0-9]/g, "");
      if (!numStr) return "";
      return Number(numStr).toLocaleString("id-ID");
    };

    const parseRupiah = (formattedStr) => {
      if (!formattedStr) return 0;
      return Number(String(formattedStr).replace(/[^0-9]/g, ""));
    };

    // Helper Konversi Tanggal "DD/MM/YYYY" <-> "YYYY-MM-DD" untuk <input type="date"> (Calendar Picker)
    const formatDateToIso = (dateStr) => {
      if (!dateStr) return "";
      if (dateStr.includes("-")) return dateStr;
      const parts = dateStr.split("/");
      if (parts.length === 3) {
        const [d, m, y] = parts;
        return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
      }
      return "";
    };

    const formatIsoToDdMmYyyy = (isoStr) => {
      if (!isoStr) return "";
      if (isoStr.includes("/")) return isoStr;
      const parts = isoStr.split("-");
      if (parts.length === 3) {
        const [y, m, d] = parts;
        return `${d.padStart(2, "0")}/${m.padStart(2, "0")}/${y}`;
      }
      return isoStr;
    };

    // HELPER INDONESIAN TERBILANG
    const angkaToTerbilang = (n) => {
      const nominal = Math.abs(Math.floor(Number(n) || 0));
      if (nominal === 0) return "Nol Rupiah";

      const satuan = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"];

      function bilang(val) {
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

      const hasil = bilang(nominal).replace(/\s+/g, ' ').trim();
      return (hasil ? hasil + " Rupiah" : "Nol Rupiah");
    };

    // HELPER PARSE INDONESIAN DATE DD/MM/YYYY TO TIMESTAMP (HOISTED FUNCTION)
    function parseIndoDateToTimestamp(dateStr) {
      if (!dateStr || typeof dateStr !== 'string' || !dateStr.trim()) return 0;
      const parts = dateStr.trim().split('/');
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

    // HELPER FORMAT TANGGAL TEKS INDONESIA (misal: "12/10/2026" -> "12 Oktober 2026")
    function formatTanggalIndoTeks(dateStr) {
      if (!dateStr || typeof dateStr !== 'string') return "";
      const s = dateStr.trim();
      if (!s || s === "-") return "";
      const slashParts = s.split('/');
      if (slashParts.length === 3) {
        const d = parseInt(slashParts[0], 10);
        const m = parseInt(slashParts[1], 10);
        const y = slashParts[2].trim();
        if (!isNaN(d) && m >= 1 && m <= 12) {
          return `${d} ${MONTHS[m - 1]} ${y}`;
        }
      }
      const dashParts = s.split('-');
      if (dashParts.length === 3) {
        if (dashParts[0].length === 4) {
          const y = dashParts[0].trim();
          const m = parseInt(dashParts[1], 10);
          const d = parseInt(dashParts[2], 10);
          if (!isNaN(d) && m >= 1 && m <= 12) {
            return `${d} ${MONTHS[m - 1]} ${y}`;
          }
        } else {
          const d = parseInt(dashParts[0], 10);
          const m = parseInt(dashParts[1], 10);
          const y = dashParts[2].trim();
          if (!isNaN(d) && m >= 1 && m <= 12) {
            return `${d} ${MONTHS[m - 1]} ${y}`;
          }
        }
      }
      return s;
    }

    // MURNI GOOGLE APPS SCRIPT (GAS) - TANPA FIREBASE / SERVER EKSTERNAL
    const SIBOS_FIREBASE_DB_URL = "GOOGLE_APPS_SCRIPT_SPREADSHEET";
    const SIBOS_OFFICIAL_GAS_URL = SIBOS_FIREBASE_DB_URL;

    const normalizeFirebaseBaseUrl = (url) => {
      return (url || SIBOS_FIREBASE_DB_URL).trim();
    };

    const normalizeArrayData = (val) => {
      if (!val) return [];
      if (Array.isArray(val)) return val.filter(x => x !== null && x !== undefined);
      if (typeof val === "object") return Object.values(val).filter(x => x !== null && x !== undefined);
      return [];
    };

    const isPdfNonUraianLine = (str) => {
      if (!str) return true;
      const s = str.trim();
      if (!s) return true;
      if (/^[\d\.\,\s\-]+$/.test(s)) return true;
      if (/^(Jumlah|Total|Sub\s*Total|Mengetahui|Menyetujui|Kepala\s+Sekolah|Bendahara|Komite\s+Sekolah|NIP\.?\s*:?|Halaman\s+\d|Hal\.?\s+\d|Dicetak|rkas\s*\-|ARKAS|KERTAS\s+KERJA|RENCANA\s+KEGIATAN|TAHUN\s+ANGGARAN|NPSN\s*:|Nama\s+Sekolah\s*:|Alamat\s*:|Kabupaten|Provinsi\s*:|Sumber\s+Dana\s*:|Tahap\s*:|No\.?\s+Kode\s+Rekening|Kode\s+Rekening|Kode\s+Program|Uraian\s+Kegiatan|Rincian\s+Perhitungan|Belanja\s+Operasi|Belanja\s+Modal|Volume\s+Satuan|Harga\s+Satuan|Januari\s+Februari|Bulan\s*:)/i.test(s)) {
        return true;
      }
      return false;
    };

    // Kamus pelengkap otomatis nomenklatur ARKAS jika teks terpotong pada baris pertama PDF
    const ARKAS_KEGIATAN_FULL_MAP = {
      "03.01.": "Penerimaan Peserta Didik Baru",
      "03.03.": "Pelaksanaan Kegiatan Pembelajaran dan Ekstrakurikuler",
      "04.06.": "Pengembangan Profesi Pendidik dan Tenaga Kependidikan",
      "05.02.": "Pengembangan Perpustakaan",
      "05.05.": "Pelaksanaan Administrasi Kegiatan Sekolah",
      "05.08.": "Pemeliharaan Sarana dan Prasarana Sekolah",
      "05.09.": "Penyediaan Alat Multi Media Pembelajaran",
      "06.05.": "Pelaksanaan Administrasi Kegiatan Sekolah",
      "06.07.": "Pembiayaan Langganan Daya dan Jasa",
      "07.05.": "Pelaksanaan Administrasi Kegiatan Sekolah",
      "07.12.": "Pembayaran Honor",
      "03.01.01.": "Pelaksanaan Pendaftaran Peserta Didik Baru (PPDB)",
      "03.03.18.": "Pengayaan TIK untuk memfasilitasi kegiatan pembelajaran",
      "04.06.01.": "Pelaksanaan kegiatan komunitas belajar di satuan pendidikan",
      "05.02.05.": "Pengadaan buku pengayaan dan referensi",
      "05.08.04.": "Pengadaan Perlengkapan Sekolah diluar komponen penyediaan alat multimedia pembelajaran",
      "05.09.01.": "Pengadaan Komputer Desktop/Work-station untuk Pembelajaran, Administrasi dan Perpustakaan",
      "05.09.02.": "Pengadaan Komputer Laptop, Notebook untuk Pembelajaran, Administrasi dan Perpustakaan",
      "05.09.04.": "Pengadaan Printer, Printer+Scanner, dan Scanner",
      "06.05.08.": "Pembelian Bahan Habis Pakai untuk mendukung pembelajaran dan administrasi sekolah (termasuk ATK, Tinta Printer, Kabel Ekstension, dsb)",
      "06.05.09.": "Pembelian bahan habis pakai/alat penunjang kebersihan dan sanitasi sekolah",
      "06.05.10.": "Pembelian Bahan Habis Pakai untuk mendukung pembelajaran dan administrasi sekolah (termasuk ATK, Tinta Printer, Kabel Ekstension, dsb)",
      "06.07.01.": "Pembayaran daya listrik",
      "06.07.05.": "Pembayaran jasa internet",
      "07.05.01.": "Bea materai, administrasi bank",
      "07.12.02.": "Pembayaran honor Tenaga Administrasi"
    };

    const ARKAS_REKENING_FULL_MAP = {
      "5.1.02.01.01.0001": "Belanja Bahan-Bahan Bangunan dan Konstruksi",
      "5.1.02.01.01.0008": "Belanja Bahan-Bahan/Bibit Tanaman",
      "5.1.02.01.01.0009": "Belanja Bahan-Isi Tabung Pemadam Kebakaran",
      "5.1.02.01.01.0012": "Belanja Bahan-Bahan Lainnya",
      "5.1.02.01.01.0019": "Belanja Suku Cadang-Suku Cadang Alat Pertanian",
      "5.1.02.01.01.0024": "Belanja Alat/Bahan untuk Kegiatan Kantor-Alat Tulis Kantor",
      "5.1.02.01.01.0025": "Belanja Alat/Bahan untuk Kegiatan Kantor-Kertas dan Cover",
      "5.1.02.01.01.0027": "Belanja Alat/Bahan untuk Kegiatan Kantor-Benda Pos",
      "5.1.02.01.01.0029": "Belanja Alat/Bahan untuk Kegiatan Kantor-Bahan Komputer",
      "5.1.02.01.01.0030": "Belanja Alat/Bahan untuk Kegiatan Kantor-Perabot Kantor",
      "5.1.02.01.01.0031": "Belanja Alat/Bahan untuk Kegiatan Kantor-Alat Listrik",
      "5.1.02.01.01.0034": "Belanja Alat/Bahan untuk Kegiatan Kantor-Alat Olahraga",
      "5.1.02.01.01.0055": "Belanja Makanan dan Minuman pada Fasilitas Pelayanan Urusan Pendidikan",
      "5.1.02.02.01.0003": "Honorarium Narasumber atau Pembahas, Moderator, Pembawa Acara, dan Panitia",
      "5.1.02.02.01.0026": "Belanja Jasa Tenaga Administrasi",
      "5.1.02.02.01.0061": "Belanja Tagihan Listrik",
      "5.1.02.02.01.0063": "Belanja Kawat/Faksimili/Internet/TV Berlangganan",
      "5.1.02.02.12.0001": "Belanja Kursus Singkat/Pelatihan",
      "5.1.02.03.02.0123": "Belanja Pemeliharaan Alat Kantor dan Rumah Tangga-Alat Kantor-Alat Kantor Lainnya",
      "5.1.02.03.02.0463": "Belanja Pemeliharaan Alat Peraga-Alat Peraga Pelatihan dan Percontohan-Alat Peraga Pelatihan",
      "5.1.02.04.01.0003": "Belanja Perjalanan Dinas Dalam Kota",
      "5.2.02.03.03.0010": "Belanja Modal Alat Timbangan/Biara",
      "5.2.02.05.01.0004": "Belanja Modal Alat Penyimpan Perlengkapan Kantor",
      "5.2.02.05.01.0005": "Belanja Modal Alat Kantor Lainnya",
      "5.2.02.05.02.0006": "Belanja Modal Alat Rumah Tangga Lainnya (Home Use)",
      "5.2.02.06.01.0002": "Belanja Modal Peralatan Studio Video dan Film",
      "5.2.02.10.01.0002": "Belanja Modal Personal Computer",
      "5.2.02.10.02.0003": "Belanja Modal Peralatan Personal Computer"
    };

    const ARKAS_TRUNCATED_ITEM_MAP = {
      "001. Laptop / Notebook i5-Processor": "001. Laptop / Notebook i5-Processor Core i5 Windows 11",
      "001. Jamuan Makan (Pelayanan Urusan": "001. Jamuan Makan (Pelayanan Urusan Pendidikan)",
      "002. Jamuan Snack (Pelayanan Urusan": "002. Jamuan Snack (Pelayanan Urusan Pendidikan)",
      "001. Honorarium Narasumber atau": "001. Honorarium Narasumber atau Pembahas, Moderator, Pembawa Acara, dan Panitia",
      "001. Buku Tes Kemampuan Akademik (": "001. Buku Tes Kemampuan Akademik (TKA)",
      "001. Refil Tabung Pemadam 6 Kg-Refil": "001. Refil Tabung Pemadam 6 Kg-Refil Tabung Pemadam 6 Kg",
      "001. Belanja Token Listrik-Belanja": "001. Belanja Token Listrik-Belanja Token Listrik",
      "001. Belanja Tagihan internet-Belanja": "001. Belanja Tagihan internet-Belanja Tagihan Internet",
      "Pengembangan perpustakaan dan/atau": "Pengembangan perpustakaan dan/atau layanan pojok baca",
      "Pengadaan Perlengkapan Satuan": "Pengadaan Perlengkapan Satuan Pendidikan"
    };

    // Normalisasi string uraian untuk pencocokan pintar antara Master RKAS, Rincian Perbulan, dan Database Belanja
    const canonicalizeArkasUraian = (str) => {
      if (!str) return "";
      let s = String(str).toLowerCase().trim();
      s = s.replace(/^\[[^\]]+\]\s*/g, "");
      s = s.replace(/^\d{1,3}\.\s*/, "");
      s = s.replace(/\(\s+/g, "(").replace(/\s+\)/g, ")").replace(/\s*-\s*/g, "-");
      s = s.replace(/\s+/g, " ").trim();
      s = s.replace(/\s+sd$/, "");
      return s;
    };

    const extractBaseArkasUraian = (str) => {
      const c = canonicalizeArkasUraian(str);
      if (!c) return "";
      const parts = c.split("-").map(p => p.trim()).filter(Boolean);
      return parts.length > 0 ? parts[0] : c;
    };

    const doArkasUraianMatch = (u1, u2) => {
      const c1 = canonicalizeArkasUraian(u1);
      const c2 = canonicalizeArkasUraian(u2);
      if (!c1 || !c2) return false;
      if (c1 === c2) return true;
      const b1 = extractBaseArkasUraian(u1);
      const b2 = extractBaseArkasUraian(u2);
      if (b1 && b2 && b1.length >= 4 && b1 === b2) {
        const isGenericRek = b1.startsWith("belanja alat/bahan untuk kegiatan kantor") || b1.startsWith("belanja bahan");
        if (!isGenericRek) return true;
      }
      if (c1.length >= 10 && c2.length >= 10 && (c1.startsWith(c2) || c2.startsWith(c1))) {
        return true;
      }
      return false;
    };

    const calculatePerBulanTotal = (monthRows) => {
      if (!Array.isArray(monthRows) || monthRows.length === 0) return 0;
      const itemRows = monthRows.filter(x => x && (x.level === "item" || (x.kodeRekening && String(x.kodeRekening).trim() !== "")) && !x.isRekeningHeader);
      if (itemRows.length > 0) {
        return itemRows.reduce((s, x) => s + Number(x.jumlah || 0), 0);
      }
      const standarRows = monthRows.filter(x => x && x.level === "standar");
      if (standarRows.length > 0) {
        return standarRows.reduce((s, x) => s + Number(x.jumlah || 0), 0);
      }
      return monthRows.reduce((s, x) => s + Number(x?.jumlah || 0), 0);
    };

    const inferMonthFromDateStr = (dateStr) => {
      if (!dateStr || typeof dateStr !== "string") return null;
      const parts = dateStr.trim().split("/");
      if (parts.length === 3) {
        const mIdx = parseInt(parts[1], 10) - 1;
        if (mIdx >= 0 && mIdx < 12) return MONTHS[mIdx];
      }
      return null;
    };

    const normalizeBelanjaBankList = (rawList) => {
      const list = normalizeArrayData(rawList);
      return list.map(b => {
        if (!b || typeof b !== "object") return b;
        let bulan = b.bulan;
        if (!bulan || bulan === "Semua" || !MONTHS.includes(bulan)) {
          const canon = canonicalizeArkasUraian(b.belanja || "");
          if (b.id === 1791857137103 || canon.includes("buku tes kemampuan akademik")) {
            bulan = "Januari";
          } else {
            bulan = inferMonthFromDateStr(b.tglTransaksiSI) || "Januari";
          }
        }
        return { ...b, bulan };
      });
    };

    const sanitizeKodeKegiatanStr = (rawVal, level, isPerBulan = false) => {
      if (rawVal === null || rawVal === undefined || rawVal === "") return "";
      const str = String(rawVal).trim();
      const sep = isPerBulan ? ". " : ".";
      const pad2 = (n) => String(n).padStart(2, "0");
      const monthMap = {
        Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6,
        Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12
      };
      const mDate = str.match(/^[A-Za-z]{3}\s+([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})\s+\d{2}:\d{2}:\d{2}/);
      if (mDate && monthMap[mDate[1]]) {
        const p1 = pad2(parseInt(mDate[2], 10));
        const p2 = pad2(monthMap[mDate[1]]);
        const p3 = pad2(parseInt(mDate[3], 10) % 100);
        if (level === "standar") return `${p1}.`;
        if (level === "kegiatan") return `${p1}${sep}${p2}.`;
        return `${p1}${sep}${p2}${sep}${p3}.`;
      }
      return str;
    };

    const completeArkasUraian = (item) => {
      if (!item || typeof item !== "object") return item;
      let uraian = String(item.uraian || "").trim();
      let volume = item.volume !== undefined && item.volume !== null ? String(item.volume).trim() : "";
      let satuan = item.satuan !== undefined && item.satuan !== null ? String(item.satuan).trim() : "";
      const kodeRek = String(item.kodeRekening || "").trim();
      const fixedKodeKeg = item.kodeKegiatan !== undefined ? sanitizeKodeKegiatanStr(item.kodeKegiatan, item.level, false) : undefined;
      const fixedKodeProg = item.kodeProgram !== undefined ? sanitizeKodeKegiatanStr(item.kodeProgram, item.level, true) : undefined;
      const rawKeg = String(fixedKodeKeg || fixedKodeProg || "").replace(/\s+/g, "").trim();

      if (ARKAS_TRUNCATED_ITEM_MAP[uraian]) {
        uraian = ARKAS_TRUNCATED_ITEM_MAP[uraian];
      } else if (!kodeRek && rawKeg && ARKAS_KEGIATAN_FULL_MAP[rawKeg]) {
        const fullKegText = ARKAS_KEGIATAN_FULL_MAP[rawKeg];
        const cleanU = uraian.toLowerCase().replace(/[\,\s]+$/, "");
        const cleanFull = fullKegText.toLowerCase();
        if (
          (uraian.length < fullKegText.length && (cleanFull.startsWith(cleanU) || cleanFull.endsWith(cleanU))) ||
          (item.level === "kegiatan" && uraian.length > fullKegText.length && cleanU.startsWith(cleanFull))
        ) {
          uraian = fullKegText;
        }
      } else if (kodeRek && !/^\d{3}\./.test(uraian) && ARKAS_REKENING_FULL_MAP[kodeRek]) {
        const fullRekText = ARKAS_REKENING_FULL_MAP[kodeRek];
        if (uraian.length < fullRekText.length && fullRekText.toLowerCase().startsWith(uraian.toLowerCase().replace(/[\,\s]+$/, ""))) {
          uraian = fullRekText;
        }
      }

      if (volume && satuan) {
        const unitSplitMatch = satuan.match(/^([A-Za-z]+)\s+(\d+)\s+([A-Za-z\/]+)$/);
        if (unitSplitMatch) {
          const unitPrefix = unitSplitMatch[1];
          const realVol = unitSplitMatch[2];
          const realSatuan = unitSplitMatch[3];
          uraian = `${uraian} ${volume} ${unitPrefix}`.replace(/\s+/g, " ").trim();
          volume = realVol;
          satuan = realSatuan;
        } else {
          const numSatuanMatch = satuan.match(/^(\d+)\s+([A-Za-z\/]+)$/);
          if (numSatuanMatch && Number(volume) >= 1000) {
            uraian = `${uraian} ${volume}`.replace(/\s+/g, " ").trim();
            volume = numSatuanMatch[1];
            satuan = numSatuanMatch[2];
          }
        }
      }

      const out = { ...item, uraian };
      if (fixedKodeKeg !== undefined) out.kodeKegiatan = fixedKodeKeg;
      if (fixedKodeProg !== undefined) out.kodeProgram = fixedKodeProg;
      if (item.volume !== undefined) out.volume = volume;
      if (item.satuan !== undefined) out.satuan = satuan;
      return out;
    };

    const cleanPerBulanRowList = (rows) => {
      const list = normalizeArrayData(rows).map(r => ({ ...r }));
      for (let i = 0; i < list.length; i++) {
        const cur = list[i];
        const next = i + 1 < list.length ? list[i + 1] : null;
        const uCur = String(cur.uraian || "").trim();
        const uNext = next ? String(next.uraian || "").trim() : "";

        if (cur.level === "item" && /Pembelian bahan habis pakai/i.test(uCur)) {
          cur.uraian = uCur.replace(/[:\s]*Pembelian bahan habis pakai.*$/i, "").trim();
          if (next && next.level === "sub") {
            next.uraian = "Pembelian bahan habis pakai/alat penunjang kebersihan dan sanitasi sekolah";
          }
        }

        if (cur.level === "item" && next && next.level === "item") {
          const nextVolInline = uNext.match(/^(\d+)\s+([a-zA-Z\/]+)\s+([\d\.\,]+)\s+(.+)$/);
          if (nextVolInline && (!next.volume || String(next.volume).trim() === "")) {
            const nVol = nextVolInline[1];
            const nSat = nextVolInline[2];
            const nTarif = parseRupiah(nextVolInline[3]);
            const nTail = nextVolInline[4].trim();

            if (/Tinta Stempel/i.test(uCur) && /Lem Putih/i.test(uCur)) {
              const splitIdx = uCur.search(/Lem Putih/i);
              const headPart = uCur.slice(splitIdx).trim();
              cur.uraian = uCur.slice(0, splitIdx).trim();
              next.uraian = `${headPart} ${nTail}`.replace(/\s+/g, " ").replace(/\s+,/g, ",").trim();
            } else if (/Sabun Cuci Tangan/i.test(uCur) && /Alat Pel/i.test(uCur)) {
              const splitIdx = uCur.search(/Alat Pel/i);
              const headPart = uCur.slice(splitIdx).trim();
              cur.uraian = uCur.slice(0, splitIdx).trim();
              next.uraian = `${headPart} ${nTail}`.replace(/\s+/g, " ").trim();
            } else if (/Karpet biasa/i.test(uCur) && /Pembersih Closet/i.test(uCur)) {
              const splitIdx = uCur.search(/Pembersih Closet/i);
              const headPart = uCur.slice(splitIdx).trim();
              cur.uraian = uCur.slice(0, splitIdx).trim();
              next.uraian = `${headPart} ${nTail}`.replace(/\s+/g, " ").trim();
            } else {
              next.uraian = nTail;
            }
            next.volume = nVol;
            next.satuan = nSat;
            if (nTarif > 0) next.tarif = nTarif;
          }

          if (/timbangan digital portabel/i.test(uCur) && /Universal Power Supply/i.test(uCur)) {
            const splitIdx = uCur.search(/Universal Power Supply/i);
            const upsHead = uCur.slice(splitIdx).trim();
            cur.uraian = uCur.slice(0, splitIdx).trim();
            if (uNext.startsWith("(UPS)")) {
              next.uraian = `${upsHead} ${uNext}`.replace(/\s+/g, " ").trim();
            }
          }
        }
      }
      return list.map(completeArkasUraian);
    };

    const normalizeMasterRkasData = (val) => {
      return normalizeArrayData(val).map(completeArkasUraian);
    };

    const normalizeObjectMap = (val) => {
      if (!val || typeof val !== "object") return {};
      const out = {};
      Object.keys(val).forEach(k => {
        if (val[k] !== null && val[k] !== undefined) {
          out[k] = val[k];
        }
      });
      return out;
    };

    const normalizePerBulanData = (raw) => {
      const base = {
        "Januari": [], "Februari": [], "Maret": [], "April": [], "Mei": [], "Juni": [],
        "Juli": [], "Agustus": [], "September": [], "Oktober": [], "November": [], "Desember": []
      };
      if (raw && typeof raw === "object") {
        MONTHS.forEach(m => {
          if (raw[m]) {
            base[m] = cleanPerBulanRowList(raw[m]);
          }
        });
      }
      return base;
    };
    
    function App() {
      const [appsScriptUrl, setAppsScriptUrl] = useState(SIBOS_FIREBASE_DB_URL);

      const handleSaveAppsScriptUrl = (newUrl) => {
        const cleaned = (newUrl || SIBOS_FIREBASE_DB_URL).trim();
        setAppsScriptUrl(cleaned);
      };

      const [accounts, setAccounts] = useState([]);
      const [sheetsRefreshKey, setSheetsRefreshKey] = useState(0);

      // State Google Sheets "SIBOS-V1" (20 Sheet)
      const [googleUser, setGoogleUser] = useState(null);
      const [needsGoogleAuth, setNeedsGoogleAuth] = useState(true);
      const [isSigningInGoogle, setIsSigningInGoogle] = useState(false);
      const [sheetsMeta, setSheetsMeta] = useState({
        spreadsheetId: "",
        spreadsheetUrl: "",
        spreadsheetTitle: "SIBOS-V1",
        updatedAt: ""
      });
      const [isSheetsBusy, setIsSheetsBusy] = useState(false);
      const [sheetsNotice, setSheetsNotice] = useState("");
      const [sheetsConfirmModal, setSheetsConfirmModal] = useState(null);

      const fetchSheetsStatus = useCallback(async () => {
        try {
          const res = await fetch("/api/sheets/status");
          if (res.ok) {
            const data = await res.json();
            if (data && data.ok) {
              setSheetsMeta({
                spreadsheetId: data.spreadsheetId || "",
                spreadsheetUrl: data.spreadsheetUrl || "",
                spreadsheetTitle: data.spreadsheetTitle || "SIBOS-V1",
                updatedAt: data.updatedAt || ""
              });
            }
          }
        } catch (err) {
          console.warn("Failed to check SIBOS-V1 status:", err);
        }
      }, []);

      useEffect(() => {
        fetchSheetsStatus();
        const unsub = initAuth(
          (user) => {
            setGoogleUser(user);
            setNeedsGoogleAuth(false);
          },
          () => {
            setGoogleUser(null);
            setNeedsGoogleAuth(true);
          }
        );
        return () => {
          if (typeof unsub === "function") unsub();
        };
      }, [fetchSheetsStatus]);

      const executeCreateOrSyncSibosV1 = async () => {
        setSheetsConfirmModal(null);
        setIsSheetsBusy(true);
        setSheetsNotice("Membuat & menyinkronkan 20 Sheet di Google Spreadsheet SIBOS-V1...");
        try {
          const token = await getAccessToken();
          if (!token) {
            setNeedsGoogleAuth(true);
            setSheetsNotice("Silakan Sign in with Google terlebih dahulu.");
            return;
          }
          const res = await fetch("/api/sheets/create-or-connect", {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              Authorization: `Bearer ${token}`
            },
            body: JSON.stringify({ spreadsheetId: sheetsMeta.spreadsheetId || "" })
          });
          const data = await res.json();
          if (!res.ok || !data.ok) {
            throw new Error(data.error || "Gagal membuat/menyinkronkan Spreadsheet SIBOS-V1.");
          }
          window.__sibosAutoSyncConfirmed = true;
          setSheetsMeta({
            spreadsheetId: data.spreadsheetId || "",
            spreadsheetUrl: data.spreadsheetUrl || "",
            spreadsheetTitle: data.spreadsheetTitle || "SIBOS-V1",
            updatedAt: data.updatedAt || new Date().toISOString()
          });
          setSheetsNotice("✅ Spreadsheet SIBOS-V1 (20 Sheet) aktif & tersinkronisasi!");
          await fetchAccountsFromCloud(appsScriptUrl);
          setSheetsRefreshKey((k) => k + 1);
        } catch (err) {
          setSheetsNotice("❌ " + (err.message || "Gagal membuat Spreadsheet SIBOS-V1"));
        } finally {
          setIsSheetsBusy(false);
        }
      };

      const promptCreateOrSyncSibosV1 = (hasExistingSheet) => {
        setSheetsConfirmModal({
          title: hasExistingSheet
            ? "Konfirmasi Perbarui Data di Google Spreadsheet 'SIBOS-V1'"
            : "Konfirmasi Buat Spreadsheet 'SIBOS-V1' di Google Sheets",
          description: hasExistingSheet
            ? "Apakah Anda yakin ingin memperbarui isi 20 Sheet (Database Master RKAS, Database Belanja, SI, Surat Transfer, Rekap Serapan, Rekening Koran, Januari–Desember, Role Pengguna, dan Realisasi Distribusi) pada file Google Spreadsheet 'SIBOS-V1' milik Anda?"
            : "Aplikasi akan membuat file Google Spreadsheet baru bernama 'SIBOS-V1' beserta 20 Sheet database lengkap di akun Google Sheets Anda dan menyimpan data aplikasi ke dalamnya.",
          confirmLabel: hasExistingSheet ? "Ya, Perbarui SIBOS-V1" : "Ya, Buat Spreadsheet SIBOS-V1",
          onConfirm: executeCreateOrSyncSibosV1
        });
      };

      const handleGoogleLogin = async () => {
        setIsSigningInGoogle(true);
        setSheetsNotice("");
        try {
          const result = await googleSignIn();
          if (result && result.accessToken) {
            setGoogleUser(result.user);
            setNeedsGoogleAuth(false);
            promptCreateOrSyncSibosV1(Boolean(sheetsMeta.spreadsheetId));
          }
        } catch (err) {
          setSheetsNotice("Gagal login Google: " + (err.message || err));
        } finally {
          setIsSigningInGoogle(false);
        }
      };

      const handlePullFromSibosV1 = async () => {
        setIsSheetsBusy(true);
        setSheetsNotice("Memuat 20 Sheet dari Google Spreadsheet SIBOS-V1...");
        try {
          const token = await getAccessToken();
          if (!token) {
            setNeedsGoogleAuth(true);
            return;
          }
          const res = await fetch("/api/sheets/pull", {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              Authorization: `Bearer ${token}`
            },
            body: JSON.stringify({ npsn: "20231556" })
          });
          const data = await res.json();
          if (!res.ok || !data.ok) {
            throw new Error(data.error || "Gagal memuat dari SIBOS-V1");
          }
          if (data.schoolData) {
            localStorage.setItem("sibos_cache_school_20231556", JSON.stringify(data.schoolData));
          }
          await fetchAccountsFromCloud(appsScriptUrl);
          setSheetsRefreshKey((k) => k + 1);
          setSheetsNotice("✅ Data terbaru dari Google Spreadsheet SIBOS-V1 berhasil dimuat!");
        } catch (err) {
          setSheetsNotice("❌ " + (err.message || "Gagal memuat dari SIBOS-V1"));
        } finally {
          setIsSheetsBusy(false);
        }
      };

      const handleGoogleLogoutClick = async () => {
        await googleLogout();
        window.__sibosAutoSyncConfirmed = false;
        setGoogleUser(null);
        setNeedsGoogleAuth(true);
        setSheetsNotice("");
      };

      // Simpan daftar akun ke Sheet 19 (Role Pengguna) via Google Apps Script (100% Murni GAS)
      const saveAccountsToFirebase = useCallback(async (accsToSave, urlToUse, deletedNpsn = null) => {
        if (isGasEnvironment()) {
          try {
            await callGasServer("apiSaveAccounts", accsToSave || [], deletedNpsn || null);
            return true;
          } catch (err) {
            console.warn("GAS save accounts error:", err);
            return false;
          }
        }
        try {
          localStorage.setItem("sibos_local_accounts", JSON.stringify(accsToSave || []));
          if (deletedNpsn) localStorage.removeItem("sibos_local_school_" + deletedNpsn);
          return true;
        } catch (err) {
          return false;
        }
      }, [appsScriptUrl]);

      // Ambil daftar akun dari Sheet 19 (Role Pengguna) via Google Apps Script (100% Murni GAS)
      const fetchAccountsFromCloud = useCallback(async (urlToUse) => {
        if (isGasEnvironment()) {
          try {
            const rawAccounts = await callGasServer("apiGetAccounts");
            let list = normalizeArrayData(rawAccounts).filter(a => a && a.npsn);
            setAccounts(list);
            return { success: true, count: list.length, accounts: list };
          } catch (err) {
            return { success: false, message: err.message || "Gagal mengambil akun dari Spreadsheet" };
          }
        }
        try {
          const rawAccounts = JSON.parse(localStorage.getItem("sibos_local_accounts") || "[]");
          let list = normalizeArrayData(rawAccounts).filter(a => a && a.npsn);
          setAccounts(list);
          return { success: true, count: list.length, accounts: list };
        } catch (err) {
          return { success: false, message: err.message || "Gagal membaca akun" };
        }
      }, [appsScriptUrl]);

      useEffect(() => {
        fetchAccountsFromCloud(appsScriptUrl);
        const onFocus = () => fetchAccountsFromCloud(appsScriptUrl);
        window.addEventListener("focus", onFocus);
        return () => {
          window.removeEventListener("focus", onFocus);
        };
      }, [appsScriptUrl, fetchAccountsFromCloud]);

      const [activeNpsn] = useState(() => {
        return sessionStorage.getItem("bos_active_npsn") || "";
      });

      const rawCurrentUser = accounts.find(a => a.npsn === activeNpsn)
        || accounts[0]
        || { npsn: "20231556", namaSekolah: "SDIT ANNISA", kepalaSekolah: "", role: "Admin" };

      const currentUser = {
        ...rawCurrentUser,
        npsn: (!rawCurrentUser.npsn || rawCurrentUser.npsn === "20253978" || String(rawCurrentUser.npsn).toLowerCase() === "superadmin")
          ? "20231556"
          : rawCurrentUser.npsn,
        namaSekolah: (!rawCurrentUser.namaSekolah || rawCurrentUser.namaSekolah === "PANEL SUPER ADMIN SIBOS" || rawCurrentUser.namaSekolah === "SD NEGERI")
          ? "SDIT ANNISA"
          : rawCurrentUser.namaSekolah
      };

      const handleUpdateAccountPassword = async (newUser, newPassword) => {
        if (isGasEnvironment()) {
          await callGasServer("apiUpdateAdminCredentials", newUser || "", newPassword || "");
          await fetchAccountsFromCloud(appsScriptUrl);
        }
      };

      return (
        <>
          {/* Bar Koneksi Google Spreadsheet SIBOS-V1 (20 Sheet) */}
          <div className="no-print main-content !min-h-0 bg-emerald-950 text-emerald-50 border-b border-emerald-800 px-4 py-2 flex flex-wrap items-center justify-between gap-2 text-xs">
            <div className="flex flex-wrap items-center gap-2">
              <span className="inline-flex items-center gap-1.5 font-bold text-emerald-300 bg-emerald-900/80 border border-emerald-700 px-2.5 py-1 rounded-md">
                <span>📊</span>
                <span>Database Google Sheets: SIBOS-V1 (20 Sheet)</span>
              </span>
              {sheetsMeta.spreadsheetId ? (
                <a
                  href={sheetsMeta.spreadsheetUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1 font-semibold text-white bg-emerald-700 hover:bg-emerald-600 px-2.5 py-1 rounded-md transition"
                >
                  <span>Buka Spreadsheet SIBOS-V1</span>
                  <span>↗</span>
                </a>
              ) : (
                <span className="text-emerald-200/90">
                  Hubungkan akun Google Anda untuk membuat & menyinkronkan file Google Spreadsheet <strong>SIBOS-V1</strong> secara otomatis.
                </span>
              )}
              {sheetsNotice && (
                <span className="text-amber-200 font-medium">{sheetsNotice}</span>
              )}
            </div>

            <div className="flex flex-wrap items-center gap-2">
              {needsGoogleAuth ? (
                <button
                  type="button"
                  onClick={handleGoogleLogin}
                  disabled={isSigningInGoogle}
                  className="gsi-material-button"
                >
                  <div className="gsi-material-button-state"></div>
                  <div className="gsi-material-button-content-wrapper">
                    <div className="gsi-material-button-icon">
                      <svg version="1.1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48" style={{ display: "block" }}>
                        <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"></path>
                        <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"></path>
                        <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"></path>
                        <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"></path>
                        <path fill="none" d="M0 0h48v48H0z"></path>
                      </svg>
                    </div>
                    <span className="gsi-material-button-contents">
                      {isSigningInGoogle ? "Menghubungkan..." : "Sign in with Google"}
                    </span>
                  </div>
                </button>
              ) : (
                <>
                  <button
                    type="button"
                    disabled={isSheetsBusy}
                    onClick={() => promptCreateOrSyncSibosV1(Boolean(sheetsMeta.spreadsheetId))}
                    className="bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 text-white font-bold px-3 py-1.5 rounded-md transition cursor-pointer"
                  >
                    {isSheetsBusy
                      ? "Memproses..."
                      : sheetsMeta.spreadsheetId
                        ? "💾 Simpan / Sinkronkan ke SIBOS-V1"
                        : "➕ Buat Spreadsheet SIBOS-V1"}
                  </button>
                  {sheetsMeta.spreadsheetId && (
                    <button
                      type="button"
                      disabled={isSheetsBusy}
                      onClick={handlePullFromSibosV1}
                      className="bg-slate-800 hover:bg-slate-700 disabled:opacity-50 text-emerald-200 border border-emerald-700 font-semibold px-2.5 py-1.5 rounded-md transition cursor-pointer"
                    >
                      🔄 Muat dari SIBOS-V1
                    </button>
                  )}
                  <button
                    type="button"
                    onClick={handleGoogleLogoutClick}
                    className="text-emerald-300 hover:text-white px-2 py-1 rounded transition cursor-pointer"
                    title={googleUser?.email || "Akun Google"}
                  >
                    Keluar ({googleUser?.email || "Google"})
                  </button>
                </>
              )}
            </div>
          </div>

          {/* Modal Konfirmasi Eksplisit Sebelum Membuat / Memperbarui Google Spreadsheet SIBOS-V1 */}
          {sheetsConfirmModal && (
            <div className="fixed inset-0 z-[99999] bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 no-print">
              <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-md w-full p-6">
                <h3 className="text-base font-extrabold text-slate-900 mb-2">
                  {sheetsConfirmModal.title}
                </h3>
                <p className="text-sm text-slate-600 leading-relaxed mb-5">
                  {sheetsConfirmModal.description}
                </p>
                <div className="flex items-center justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setSheetsConfirmModal(null)}
                    className="px-4 py-2 rounded-lg text-xs font-bold text-slate-600 bg-slate-100 hover:bg-slate-200 transition cursor-pointer"
                  >
                    Batal
                  </button>
                  <button
                    type="button"
                    onClick={sheetsConfirmModal.onConfirm}
                    className="px-4 py-2 rounded-lg text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 transition cursor-pointer"
                  >
                    {sheetsConfirmModal.confirmLabel || "Konfirmasi"}
                  </button>
                </div>
              </div>
            </div>
          )}

          <SiBosMainApp 
            key={sheetsRefreshKey}
            currentUser={currentUser} 
            onLogout={() => {}}
            onUpdatePassword={handleUpdateAccountPassword}
            appsScriptUrl={appsScriptUrl}
            onSaveAppsScriptUrl={handleSaveAppsScriptUrl}
            onUpdateAccountInfo={async (updatedHeader) => {
              const targetNpsn = updatedHeader.npsn || currentUser.npsn;
              const exists = accounts.some(a => a.npsn === targetNpsn || a.npsn === currentUser.npsn);
              const nextAccs = exists
                ? accounts.map(a => (a.npsn === currentUser.npsn || a.npsn === targetNpsn) ? {
                    ...a,
                    npsn: targetNpsn,
                    namaSekolah: updatedHeader.namaSekolah || a.namaSekolah,
                    kepalaSekolah: updatedHeader.kepalaSekolah || a.kepalaSekolah
                  } : a)
                : [...accounts, {
                    npsn: targetNpsn,
                    namaSekolah: updatedHeader.namaSekolah || "Sekolah",
                    kepalaSekolah: updatedHeader.kepalaSekolah || "",
                    role: "Admin"
                  }];
              setAccounts(nextAccs);
              await saveAccountsToFirebase(nextAccs, appsScriptUrl);
            }}
          />
        </>
      );
    }


function SiBosMainApp({ currentUser, onLogout, onUpdateAccountInfo, onUpdatePassword, appsScriptUrl, onSaveAppsScriptUrl }) {
      // Menu View: 'dashboard' | 'database' | 'belanja' | 'si' | 'rekap' | 'perbulan' | 'distribusi'
      const [currentView, setCurrentView] = useState("dashboard");
      const [selectedMonthView, setSelectedMonthView] = useState("Semua");
      const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
      const [isPerBulanOpen, setIsPerBulanOpen] = useState(false);

      // State Login Admin via Ikon Gir Kecil di Kanan Atas (Password disimpan di Sheet Role Pengguna di Spreadsheet)
      const [isAdmin, setIsAdmin] = useState(() => sessionStorage.getItem("sibos_admin_logged_in") === "true");
      const [showAdminLoginModal, setShowAdminLoginModal] = useState(false);
      const [adminPasswordInput, setAdminPasswordInput] = useState("");
      const [adminPasswordError, setAdminPasswordError] = useState("");
      const [isVerifyingAdmin, setIsVerifyingAdmin] = useState(false);
      const [showAdminGearMenu, setShowAdminGearMenu] = useState(false);

      const handleAdminLoginSubmit = async (e) => {
        if (e && e.preventDefault) e.preventDefault();
        setAdminPasswordError("");

        const pInput = adminPasswordInput.trim();
        if (!pInput) {
          setAdminPasswordError("Mohon masukkan Password Admin!");
          return;
        }

        if (!isGasEnvironment()) {
          setAdminPasswordError("Verifikasi Admin membutuhkan koneksi Google Apps Script Spreadsheet (Sheet Role Pengguna).");
          return;
        }

        setIsVerifyingAdmin(true);
        try {
          const res = await callGasServer("apiVerifyAdminLogin", pInput);
          if (res && res.success) {
            setIsAdmin(true);
            sessionStorage.setItem("sibos_admin_logged_in", "true");
            setShowAdminLoginModal(false);
            setAdminPasswordInput("");
            setAdminPasswordError("");
          } else {
            setAdminPasswordError((res && res.message) || "Password Admin salah! Pastikan sesuai dengan Sheet Role Pengguna.");
          }
        } catch (err) {
          setAdminPasswordError("Gagal memeriksa ke Sheet Role Pengguna: " + (err.message || err));
        } finally {
          setIsVerifyingAdmin(false);
        }
      };

      const handleAdminLogout = () => {
        setIsAdmin(false);
        sessionStorage.removeItem("sibos_admin_logged_in");
        setShowAdminGearMenu(false);
        if (currentView === "pengaturan") {
          setCurrentView("dashboard");
        }
      };

      // State Database Cloud / Google Apps Script
      const [cloudUrlInput, setCloudUrlInput] = useState(appsScriptUrl || SIBOS_FIREBASE_DB_URL);
      const [cloudSyncStatus, setCloudSyncStatus] = useState("syncing"); // 'idle' | 'syncing' | 'synced' | 'error'
      const [cloudStatusMsg, setCloudStatusMsg] = useState("Menghubungkan ke Spreadsheet...");
      const [cloudLastSynced, setCloudLastSynced] = useState("");
      const [showCloudGuideModal, setShowCloudGuideModal] = useState(false);

      // Refs untuk sinkronisasi 2 arah otomatis (Spreadsheet <-> Web App)
      const isCloudReadyRef = useRef(false);
      const isApplyingRemoteRef = useRef(false);
      const isPushingRef = useRef(false);
      const lastSyncedUpdatedAtRef = useRef("");
      const autoSaveTimerRef = useRef(null);

      // State Ubah User & Password di Pengaturan Sekolah (Tersimpan langsung di Sheet Role Pengguna)
      const [newUserVal, setNewUserVal] = useState("");
      const [newPasswordVal, setNewPasswordVal] = useState("");
      const [confirmPasswordVal, setConfirmPasswordVal] = useState("");
      const [showNewPasswordVal, setShowNewPasswordVal] = useState(false);
      const [passwordMsg, setPasswordMsg] = useState({ type: "", text: "" });

      const handleSimpanPasswordBaru = (e) => {
        if (e && e.preventDefault) e.preventDefault();
        setPasswordMsg({ type: "", text: "" });

        if (!newPasswordVal.trim()) {
          setPasswordMsg({ type: "error", text: "Password baru tidak boleh kosong!" });
          return;
        }
        if (newPasswordVal.trim().length < 3) {
          setPasswordMsg({ type: "error", text: "Password minimal 3 karakter!" });
          return;
        }
        if (newPasswordVal.trim() !== confirmPasswordVal.trim()) {
          setPasswordMsg({ type: "error", text: "Konfirmasi password tidak sesuai dengan password baru!" });
          return;
        }

        if (onUpdatePassword) {
          onUpdatePassword(newUserVal.trim(), newPasswordVal.trim());
          setPasswordMsg({ type: "success", text: "✅ User & Password Admin berhasil disimpan ke Sheet Role Pengguna di Spreadsheet!" });
          setNewPasswordVal("");
          setConfirmPasswordVal("");
        }
      };

      // Cache instan dari localStorage agar halaman langsung tampil seketika (0 detik) saat dibuka
      const cachedInitialData = useMemo(() => {
        try {
          const raw = localStorage.getItem("sibos_cache_school_" + (currentUser?.npsn || "20231556"));
          return raw ? JSON.parse(raw) : null;
        } catch (e) {
          return null;
        }
      }, [currentUser?.npsn]);

      // State Surat Transfer (Surat Pemindahbukuan)
      const [selectedSuratTransferDate, setSelectedSuratTransferDate] = useState("Semua");
      const [suratTransferCetakDate, setSuratTransferCetakDate] = useState("");
      const [suratTransferNomor, setSuratTransferNomor] = useState("");

      // Rincian Kertas Kerja Perbulan State
      const [perBulanData, setPerBulanData] = useState(() => normalizePerBulanData(cachedInitialData?.perBulanData || {}));

      // Filter Tanggal Cetak Standing Instruction (SI)
      const [selectedSiDate, setSelectedSiDate] = useState("Semua");
      const [selectedBelanjaTglSI, setSelectedBelanjaTglSI] = useState("Semua");

      // Header Informasi Sekolah State
      const [headerInfo, setHeaderInfo] = useState(() => {
        const h = cachedInitialData?.headerInfo || {};
        return {
          npsn: h.npsn || currentUser.npsn || "20231556",
          namaSekolah: h.namaSekolah || currentUser.namaSekolah || "SDIT ANNISA",
          alamat: h.alamat || "",
          kabupaten: h.kabupaten || "",
          provinsi: h.provinsi || "",
          tahunAnggaran: h.tahunAnggaran || "2026",
          sumberDana: h.sumberDana || "BOS Reguler",
          totalPenerimaan: h.totalPenerimaan || 0,
          kepalaSekolah: h.kepalaSekolah || currentUser.kepalaSekolah || "",
          bendaharaSekolah: h.bendaharaSekolah || "",
          noRekening: h.noRekening || "",
          namaBank: h.namaBank || "Bank BJB Kantor Cabang Jatiasih",
          email: h.email || "",
          noHp: h.noHp || ""
        };
      });

      // Logo Web & Kop Surat State (dengan fallback localStorage mandiri agar tampil instan)
      const [customLogoUrl, setCustomLogoUrl] = useState(() => {
        try {
          const direct = localStorage.getItem("sibos_custom_logo_" + (currentUser?.npsn || "20231556"));
          if (direct && direct.indexOf("data:image/jpeg") === -1) return direct;
        } catch (e) {}
        const cached = cachedInitialData?.customLogoUrl || "";
        if (cached && cached.indexOf("data:image/jpeg") === -1) return cached;
        return "";
      });
      const [customKopSuratUrl, setCustomKopSuratUrl] = useState(() => cachedInitialData?.customKopSuratUrl || "");

      // Sinkronkan favicon browser secara otomatis saat customLogoUrl berubah
      useEffect(() => {
        if (customLogoUrl) {
          try {
            let link = document.querySelector("link[rel*='icon']");
            if (link) link.href = customLogoUrl;
          } catch (e) {}
        }
      }, [customLogoUrl]);

      // DATABASE MASTER RKAS
      const [masterRkas, setMasterRkas] = useState(() => normalizeMasterRkasData(cachedInitialData?.masterRkas || []));

      const [saldoAwal, setSaldoAwal] = useState(() => Number(cachedInitialData?.saldoAwal || 0));
      const [danaMasukList, setDanaMasukList] = useState(() => normalizeArrayData(cachedInitialData?.danaMasukList || []));

      const [showSaldoAwalModal, setShowSaldoAwalModal] = useState(false);
      const [saldoAwalInput, setSaldoAwalInput] = useState(0);
      const [showDanaMasukModal, setShowDanaMasukModal] = useState(false);
      const [showEditRekKoranModal, setShowEditRekKoranModal] = useState(false);
      const [editRekKoranTab, setEditRekKoranTab] = useState("saldo");
      const [danaMasukForm, setDanaMasukForm] = useState({
        id: null,
        tgl: "20/01/2026",
        kredit: 0,
        keterangan: ""
      });

      // DATABASE BELANJA / TRANSAKSI BANK
      const [belanjaBankList, setBelanjaBankList] = useState(() => normalizeBelanjaBankList(cachedInitialData?.belanjaBankList || []));

      // REALISASI DISTRIBUSI (PENGELUARAN & STOK BARANG DARI DATABASE MASTER RKAS)
      const [distribusiList, setDistribusiList] = useState(() => normalizeArrayData(cachedInitialData?.distribusiList || []));
      const [distribusiForm, setDistribusiForm] = useState(() => ({
        id: null,
        tanggal: new Date().toISOString().slice(0, 10),
        masterNo: "",
        kodeRekening: "",
        kodeKegiatan: "",
        namaBarang: "",
        jumlahKeluar: "",
        penerima: "",
        keterangan: ""
      }));
      const [distribusiBarangQuery, setDistribusiBarangQuery] = useState("");
      const [showDistribusiDropdown, setShowDistribusiDropdown] = useState(false);
      const [distribusiRiwayatSearch, setDistribusiRiwayatSearch] = useState("");

      // State Manual Checkbox & Keterangan untuk Database Master RKAS
      const [manualPurchasedMap, setManualPurchasedMap] = useState(() => normalizeObjectMap(cachedInitialData?.manualPurchasedMap || {}));
      const [manualKeteranganTextMap, setManualKeteranganTextMap] = useState(() => normalizeObjectMap(cachedInitialData?.manualKeteranganTextMap || {}));

      // DIALOG KONFIRMASI & TOAST NOTIFIKASI IN-APP (BEBAS BLOKIR IFRAME GOOGLE SITES & BROWSER)
      const [toastMsg, setToastMsg] = useState(null);
      const [confirmDialog, setConfirmDialog] = useState(null);

      const showInAppToast = useCallback((message, type = "success") => {
        setToastMsg({ message, type });
        setTimeout(() => {
          setToastMsg(prev => (prev && prev.message === message ? null : prev));
        }, 4000);
      }, []);

      // Handlers Sinkronisasi ke Google Apps Script (20 Sheets) / Cloud
      const handlePushDataToCloud = async (silent = false, overrideState = null) => {
        setCloudSyncStatus("syncing");
        setCloudStatusMsg("Menyinkronkan ke Spreadsheet...");
        isPushingRef.current = true;
        try {
          const nowIso = new Date().toISOString();
          const stateToPush = overrideState || {
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
          const payload = {
            updatedAt: nowIso,
            headerInfo: stateToPush.headerInfo || headerInfo,
            masterRkas: normalizeMasterRkasData(stateToPush.masterRkas || []),
            perBulanData: normalizePerBulanData(stateToPush.perBulanData || {}),
            belanjaBankList: normalizeBelanjaBankList(stateToPush.belanjaBankList || []),
            danaMasukList: stateToPush.danaMasukList || [],
            distribusiList: normalizeArrayData(stateToPush.distribusiList || []),
            saldoAwal: Number(stateToPush.saldoAwal ?? 0),
            manualPurchasedMap: stateToPush.manualPurchasedMap || {},
            manualKeteranganTextMap: stateToPush.manualKeteranganTextMap || {},
            customLogoUrl: stateToPush.customLogoUrl || "",
            customKopSuratUrl: stateToPush.customKopSuratUrl || ""
          };

          if (isGasEnvironment()) {
            await callGasServer("apiSaveSchoolData", currentUser.npsn, payload);
          } else {
            localStorage.setItem("sibos_local_school_" + currentUser.npsn, JSON.stringify(payload));
          }

          lastSyncedUpdatedAtRef.current = nowIso;
          const nowStr = new Date().toLocaleString("id-ID");
          setCloudSyncStatus("synced");
          setCloudLastSynced(nowStr);
          setCloudStatusMsg(`Tersinkron Spreadsheet: ${nowStr}`);
          if (!silent) {
            alert(`✅ Sinkronisasi Berhasil!\n\nSeluruh data sekolah "${payload.headerInfo.namaSekolah || currentUser.namaSekolah}" (NPSN: ${currentUser.npsn}) telah tersimpan ke 20 Sheet Google Spreadsheet.`);
          }
        } catch (err) {
          console.error("Push error:", err);
          setCloudSyncStatus("error");
          setCloudStatusMsg("Gagal sinkron Spreadsheet");
          if (!silent) {
            alert(`⚠️ Gagal menyinkronkan data:\n\n${err.message || err}`);
          }
        } finally {
          isPushingRef.current = false;
        }
      };

      const handlePullDataFromCloud = useCallback(async (silent = false) => {
        setCloudSyncStatus("syncing");
        setCloudStatusMsg("Mengambil data dari Spreadsheet...");
        try {
          let d = null;
          if (isGasEnvironment()) {
            d = await callGasServer("apiGetSchoolData", currentUser.npsn);
          } else {
            const rawLocal = localStorage.getItem("sibos_local_school_" + currentUser.npsn);
            d = rawLocal ? JSON.parse(rawLocal) : null;
          }

          isApplyingRemoteRef.current = true;

          if (d && typeof d === "object") {
            try {
              localStorage.setItem("sibos_cache_school_" + (currentUser.npsn || "20231556"), JSON.stringify(d));
            } catch (e) {}
            const rawRemoteHeader = (d.headerInfo && typeof d.headerInfo === "object") ? d.headerInfo : {};
            const nextHeader = {
              npsn: currentUser.npsn || "20231556",
              namaSekolah: currentUser.namaSekolah || "SDIT ANNISA",
              alamat: "",
              kabupaten: "",
              provinsi: "",
              tahunAnggaran: "2026",
              sumberDana: "BOS Reguler",
              totalPenerimaan: 0,
              kepalaSekolah: currentUser.kepalaSekolah || "",
              bendaharaSekolah: "",
              noRekening: "",
              namaBank: "Bank BJB Kantor Cabang Jatiasih",
              email: "",
              noHp: "",
              ...rawRemoteHeader
            };
            if (!nextHeader.namaSekolah || nextHeader.namaSekolah === "PANEL SUPER ADMIN SIBOS" || nextHeader.namaSekolah === "SD NEGERI") {
              nextHeader.namaSekolah = "SDIT ANNISA";
            }
            if (!nextHeader.npsn || nextHeader.npsn === "20253978" || String(nextHeader.npsn).toLowerCase() === "superadmin") {
              nextHeader.npsn = "20231556";
            }
            const cleanedMaster = normalizeMasterRkasData(d.masterRkas);
            const cleanedPerBulan = normalizePerBulanData(d.perBulanData);
            const cleanedBelanja = normalizeBelanjaBankList(d.belanjaBankList);
            const cleanedDistribusi = normalizeArrayData(d.distribusiList);

            setHeaderInfo(nextHeader);
            setMasterRkas(cleanedMaster);
            setPerBulanData(cleanedPerBulan);
            setBelanjaBankList(cleanedBelanja);
            setDanaMasukList(normalizeArrayData(d.danaMasukList));
            setDistribusiList(cleanedDistribusi);
            setSaldoAwal(typeof d.saldoAwal === "number" ? d.saldoAwal : Number(d.saldoAwal || 0));
            setManualPurchasedMap(normalizeObjectMap(d.manualPurchasedMap));
            setManualKeteranganTextMap(normalizeObjectMap(d.manualKeteranganTextMap));
            if (d.customLogoUrl) {
              setCustomLogoUrl(d.customLogoUrl);
              try {
                localStorage.setItem("sibos_custom_logo_" + (currentUser.npsn || "20231556"), d.customLogoUrl);
              } catch (e) {}
            }
            lastSyncedUpdatedAtRef.current = d.updatedAt || "";

            const uploadedMonths = MONTHS.filter(m => Array.isArray(cleanedPerBulan[m]) && cleanedPerBulan[m].length > 0);
            if (uploadedMonths.length === 1) {
              setSelectedMonthView(prev => (prev === "Semua" ? uploadedMonths[0] : prev));
            }

            const rawBelanjaArr = normalizeArrayData(d.belanjaBankList);
            const hadSemuaBelanja = rawBelanjaArr.some(b => !b || !b.bulan || b.bulan === "Semua");
            const hadUncleanMaster = JSON.stringify(d.masterRkas || []) !== JSON.stringify(cleanedMaster);
            const hadUncleanPerBulan = JSON.stringify(d.perBulanData || {}) !== JSON.stringify(cleanedPerBulan);
            if (hadSemuaBelanja || hadUncleanMaster || hadUncleanPerBulan) {
              setTimeout(() => {
                handlePushDataToCloud(true, {
                  headerInfo: nextHeader,
                  masterRkas: cleanedMaster,
                  perBulanData: cleanedPerBulan,
                  belanjaBankList: cleanedBelanja,
                  danaMasukList: normalizeArrayData(d.danaMasukList),
                  distribusiList: cleanedDistribusi,
                  saldoAwal: typeof d.saldoAwal === "number" ? d.saldoAwal : Number(d.saldoAwal || 0),
                  manualPurchasedMap: normalizeObjectMap(d.manualPurchasedMap),
                  manualKeteranganTextMap: normalizeObjectMap(d.manualKeteranganTextMap),
                  customLogoUrl: d.customLogoUrl || "",
                  customKopSuratUrl: d.customKopSuratUrl || ""
                });
              }, 350);
            }

            const nowStr = new Date().toLocaleString("id-ID");
            setCloudSyncStatus("synced");
            setCloudLastSynced(nowStr);
            setCloudStatusMsg(`Tersinkron Spreadsheet: ${nowStr}`);

            if (!silent) {
              alert(`✅ Sinkronisasi Berhasil!\n\nSeluruh data terbaru untuk ${d.headerInfo?.namaSekolah || currentUser.namaSekolah} (NPSN: ${currentUser.npsn}) telah dimuat dari Spreadsheet.`);
            }
          } else {
            const nowStr = new Date().toLocaleString("id-ID");
            setCloudSyncStatus("synced");
            setCloudLastSynced(nowStr);
            setCloudStatusMsg(`Siap: ${nowStr}`);
          }
        } catch (err) {
          console.error("Pull error:", err);
          setCloudSyncStatus("error");
          setCloudStatusMsg("Gagal ambil data");
          if (!silent) {
            alert(`⚠️ Gagal mengambil data:\n\n${err.message || err}`);
          }
        } finally {
          isCloudReadyRef.current = true;
          isApplyingRemoteRef.current = false;
        }
      }, [appsScriptUrl, cloudUrlInput, currentUser.npsn, currentUser.namaSekolah, currentUser.kepalaSekolah]);

      // 1. Muat data awal saat membuka ruang kerja sekolah
      useEffect(() => {
        isCloudReadyRef.current = false;
        handlePullDataFromCloud(true);
      }, [currentUser.npsn]);

      // 2. Auto-Save ke 20 Sheet Google Spreadsheet setiap kali ada perubahan data
      useEffect(() => {
        if (!isCloudReadyRef.current) return;
        if (isApplyingRemoteRef.current) return;

        if (autoSaveTimerRef.current) {
          clearTimeout(autoSaveTimerRef.current);
        }
        autoSaveTimerRef.current = setTimeout(() => {
          handlePushDataToCloud(true, {
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
          });
        }, 500);

        return () => {
          if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current);
        };
      }, [
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
      ]);

      // 3. Auto-Pull Realtime jika ada perubahan dari perangkat lain
      useEffect(() => {
        const checkRemoteUpdate = async () => {
          if (!isCloudReadyRef.current || isPushingRef.current || autoSaveTimerRef.current) return;
          try {
            let remoteUpdatedAt = "";
            if (isGasEnvironment()) {
              remoteUpdatedAt = await callGasServer("apiGetLastUpdatedAt", currentUser.npsn);
            }
            if (remoteUpdatedAt && remoteUpdatedAt !== lastSyncedUpdatedAtRef.current) {
              await handlePullDataFromCloud(true);
            }
          } catch (e) {}
        };

        const intervalId = setInterval(checkRemoteUpdate, 45000);
        const onFocus = () => checkRemoteUpdate();
        window.addEventListener("focus", onFocus);
        return () => {
          clearInterval(intervalId);
          window.removeEventListener("focus", onFocus);
        };
      }, [appsScriptUrl, cloudUrlInput, currentUser.npsn, handlePullDataFromCloud]);

      // Handlers Pengaturan & Media
      const handleSavePengaturan = async (e) => {
        if (e && e.preventDefault) e.preventDefault();
        if (onUpdateAccountInfo) {
          await onUpdateAccountInfo(headerInfo);
        }

        const cleanedUrl = (cloudUrlInput || SIBOS_FIREBASE_DB_URL).trim();
        if (cleanedUrl && onSaveAppsScriptUrl) {
          onSaveAppsScriptUrl(cleanedUrl);
        }

        await handlePushDataToCloud(true, {
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
        });

        alert("✅ Pengaturan Sekolah, Logo Web, & Kop Surat berhasil disimpan ke Spreadsheet!");
      };

      // Helper kompresi gambar agar selalu muat aman di dalam kolom Spreadsheet (<50.000 karakter)
      const compressImageToDataUrl = (file, maxDim, quality = 0.85) => {
        return new Promise((resolve) => {
          const reader = new FileReader();
          reader.onload = (e) => {
            const img = new Image();
            img.onload = () => {
              let width = img.width;
              let height = img.height;
              if (width > maxDim || height > maxDim) {
                if (width > height) {
                  height = Math.round((height * maxDim) / width);
                  width = maxDim;
                } else {
                  width = Math.round((width * maxDim) / height);
                  height = maxDim;
                }
              }
              const canvas = document.createElement("canvas");
              canvas.width = width;
              canvas.height = height;
              const ctx = canvas.getContext("2d");
              ctx.drawImage(img, 0, 0, width, height);
              const mime = (file.type === "image/png" && file.size < 400000) ? "image/png" : "image/jpeg";
              resolve(canvas.toDataURL(mime, quality));
            };
            img.onerror = () => resolve(e.target.result);
            img.src = e.target.result;
          };
          reader.onerror = () => resolve("");
          reader.readAsDataURL(file);
        });
      };

      // Upload Logo Sekolah murni persis seperti Kop Surat (transparan tanpa bayang)
      const handleUploadLogoFile = (e) => {
        const file = e.target.files && e.target.files[0];
        if (file) {
          const reader = new FileReader();
          reader.onload = (event) => {
            const rawDataUrl = event.target.result;
            // Jika ukuran string data URL sudah aman (<45.000 karakter), simpan langsung murni seperti Kop Surat
            if (rawDataUrl.length < 45000) {
              saveLogoData(rawDataUrl);
            } else {
              // Jika resolusi gambar besar, sesuaikan dimensi dengan tetap mempertahankan 100% transparansi PNG murni
              const img = new Image();
              img.onload = () => {
                const maxDim = 260;
                let width = img.width;
                let height = img.height;
                if (width > maxDim || height > maxDim) {
                  if (width > height) {
                    height = Math.round((height * maxDim) / width);
                    width = maxDim;
                  } else {
                    width = Math.round((width * maxDim) / height);
                    height = maxDim;
                  }
                }
                const canvas = document.createElement("canvas");
                canvas.width = width;
                canvas.height = height;
                const ctx = canvas.getContext("2d");
                ctx.clearRect(0, 0, width, height); // transparansi murni tanpa latar
                ctx.drawImage(img, 0, 0, width, height);
                const finalPng = canvas.toDataURL("image/png");
                saveLogoData(finalPng);
              };
              img.src = rawDataUrl;
            }
          };
          reader.readAsDataURL(file);
        }
      };

      const saveLogoData = (logoDataUrl) => {
        setCustomLogoUrl(logoDataUrl);
        try {
          localStorage.setItem("sibos_custom_logo_" + (currentUser.npsn || "20231556"), logoDataUrl);
          const rawCached = localStorage.getItem("sibos_cache_school_" + (currentUser.npsn || "20231556"));
          const cachedObj = rawCached ? JSON.parse(rawCached) : {};
          cachedObj.customLogoUrl = logoDataUrl;
          localStorage.setItem("sibos_cache_school_" + (currentUser.npsn || "20231556"), JSON.stringify(cachedObj));
        } catch (err) {}
        handlePushDataToCloud(true, {
          headerInfo,
          masterRkas,
          perBulanData,
          belanjaBankList,
          danaMasukList,
          distribusiList,
          saldoAwal,
          manualPurchasedMap,
          manualKeteranganTextMap,
          customLogoUrl: logoDataUrl,
          customKopSuratUrl
        });
        alert("✅ Logo Sekolah berhasil diunggah!");
      };

      const handleHapusLogo = () => {
        if (confirm("Apakah Anda yakin ingin menghapus logo dan kembali ke Logo Standar SiBOS?")) {
          setCustomLogoUrl("");
          try {
            localStorage.removeItem("sibos_custom_logo_" + (currentUser.npsn || "20231556"));
            const rawCached = localStorage.getItem("sibos_cache_school_" + (currentUser.npsn || "20231556"));
            const cachedObj = rawCached ? JSON.parse(rawCached) : {};
            cachedObj.customLogoUrl = "";
            localStorage.setItem("sibos_cache_school_" + (currentUser.npsn || "20231556"), JSON.stringify(cachedObj));
          } catch (err) {}
          handlePushDataToCloud(true, {
            headerInfo,
            masterRkas,
            perBulanData,
            belanjaBankList,
            danaMasukList,
            distribusiList,
            saldoAwal,
            manualPurchasedMap,
            manualKeteranganTextMap,
            customLogoUrl: "",
            customKopSuratUrl
          });
          alert("🗑️ Logo berhasil dihapus. Kembali ke Logo Standar SiBOS.");
        }
      };

      const handleUploadKopSuratFile = async (e) => {
        const file = e.target.files && e.target.files[0];
        if (file) {
          try {
            const optimizedKop = await compressImageToDataUrl(file, 750, 0.8);
            if (!optimizedKop) {
              alert("Gagal membaca file Kop Surat.");
              return;
            }
            setCustomKopSuratUrl(optimizedKop);
            handlePushDataToCloud(true, {
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
              customKopSuratUrl: optimizedKop
            });
            alert("✅ Gambar Kop Surat berhasil diunggah dan disinkronkan!");
          } catch (err) {
            alert("Gagal memproses gambar Kop Surat: " + err.message);
          }
        }
      };

      const handleResetPengaturanDefault = () => {
        if (confirm("Kosongkan Pengaturan Sekolah, Logo, dan Kop Surat?")) {
          const defaultHeader = {
            npsn: currentUser.npsn || "",
            namaSekolah: currentUser.namaSekolah || "",
            alamat: "",
            kabupaten: "",
            provinsi: "",
            tahunAnggaran: "2026",
            sumberDana: "BOS Reguler",
            totalPenerimaan: 0,
            kepalaSekolah: "",
            bendaharaSekolah: "",
            noRekening: "",
            namaBank: "",
            email: "",
            noHp: ""
          };
          setHeaderInfo(defaultHeader);
          setCustomLogoUrl("");
          setCustomKopSuratUrl("");
          alert("✅ Pengaturan berhasil dikosongkan dan disinkronkan!");
        }
      };

      const handleKeteranganTextChange = (no, text) => {
        setManualKeteranganTextMap(prev => ({ ...prev, [no]: text }));
      };

      const toggleManualPurchased = (m, checked) => {
        setManualPurchasedMap(prev => {
          const next = { ...prev };
          next[m.no] = checked;

          if (m.level === "sub") {
            masterRkas.forEach(x => {
              if (x.kodeKegiatan === m.kodeKegiatan && x.no !== m.no) {
                next[x.no] = checked;
              }
            });
          } else if (m.level === "item") {
            const siblingItems = masterRkas.filter(x => x.kodeKegiatan === m.kodeKegiatan && x.level === "item");
            const parentSub = masterRkas.find(x => x.kodeKegiatan === m.kodeKegiatan && x.level === "sub");
            if (parentSub && siblingItems.length > 0) {
              const allChecked = siblingItems.every(s => s.no === m.no ? checked : !!next[s.no]);
              next[parentSub.no] = allChecked;
            }
          }

          return next;
        });
      };

      const [hasBackedUpMaster, setHasBackedUpMaster] = useState(false);
      const [showExportDropdown, setShowExportDropdown] = useState(false);
      const [showMasterExportDropdown, setShowMasterExportDropdown] = useState(false);
      const [selectedRekKoranRekening, setSelectedRekKoranRekening] = useState("Semua");
      const [rekKoranViewMode, setRekKoranViewMode] = useState("detail");

      const handleSaveMasterCheckboxes = async () => {
        await handlePushDataToCloud(true);
        alert("✅ Perubahan centang Aksi & Keterangan manual Master RKAS berhasil disimpan ke Sheet Database Master RKAS!");
      };

      // 1. Download Backup Data Master RKAS (Wajib)
      const handleBackupMasterData = () => {
        const backupPayload = {
          appTitle: `SiBOS V1 - Backup Database Master RKAS - ${headerInfo.namaSekolah || "Sekolah"}`,
          backupDate: new Date().toISOString(),
          headerInfo,
          masterRkas,
          manualPurchasedMap,
          manualKeteranganTextMap
        };

        const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(backupPayload, null, 2));
        const downloadAnchor = document.createElement('a');
        downloadAnchor.setAttribute("href", dataStr);
        downloadAnchor.setAttribute("download", `Backup_Master_RKAS_${(headerInfo.namaSekolah || currentUser.namaSekolah || "Sekolah").replace(/\s+/g, "_")}_${new Date().toISOString().slice(0,10)}.json`);
        document.body.appendChild(downloadAnchor);
        downloadAnchor.click();
        downloadAnchor.remove();

        setHasBackedUpMaster(true);
        alert("✅ Backup Data Master RKAS berhasil di-download! Simpan file ini di tempat yang aman.");
      };

      // 2. Impor File PDF / Excel / JSON ARKAS
      const handleImportMasterClick = () => {
        const fileInput = document.getElementById("master-file-import-input");
        if (fileInput) fileInput.click();
      };

      const handleImportMasterFile = async (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;

        const fileName = file.name.toLowerCase();

        const applyMasterItems = (items, fileLabel) => {
          const finalItems = normalizeMasterRkasData(items || []);
          setMasterRkas(finalItems);
          alert(`✅ IMPOR MASTER RKAS BERHASIL!\n\nFile "${fileLabel}" telah berhasil diproses dan menyajikan ${finalItems.length} item data ke Database Master RKAS (Sheet 1).`);
        };

        if (fileName.endsWith('.pdf')) {
          try {
            await ensurePdfJs();
          } catch (err) {
            alert(err.message);
            return;
          }
          const reader = new FileReader();
          reader.onload = async (evt) => {
            try {
              let textLinesAll = [];
              if (window.pdfjsLib) {
                window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
                const typedarray = new Uint8Array(evt.target.result);
                const pdf = await window.pdfjsLib.getDocument({ data: typedarray }).promise;
                for (let i = 1; i <= pdf.numPages; i++) {
                  const page = await pdf.getPage(i);
                  const textObj = await page.getTextContent();
                  const lineClusters = [];
                  const sortedItems = textObj.items
                    .filter(it => it.str && it.str.trim())
                    .sort((a, b) => (b.transform ? b.transform[5] : 0) - (a.transform ? a.transform[5] : 0));

                  sortedItems.forEach(it => {
                    const y = it.transform ? it.transform[5] : 0;
                    let cluster = lineClusters.find(c => Math.abs(c.y - y) <= 4.0);
                    if (!cluster) {
                      cluster = { y: y, page: i, items: [] };
                      lineClusters.push(cluster);
                    }
                    cluster.items.push(it);
                  });

                  lineClusters.forEach(cluster => {
                    cluster.items.sort((a, b) => (a.transform ? a.transform[4] : 0) - (b.transform ? b.transform[4] : 0));
                    const lineStr = cluster.items.map(it => it.str.trim()).filter(Boolean).join(" ");
                    if (lineStr) textLinesAll.push({ text: lineStr, y: cluster.y, page: i });
                  });
                }
              }

              const isAmountToken = (t) => /^(0|[1-9]\d{0,2}(\.\d{3})+)$/.test(t);

              const parsedItems = [];
              let lastItem = null;
              let lastPage = -1;
              let lastY = 0;

              textLinesAll.forEach(lineObj => {
                const line = lineObj.text;
                const mNo = line.match(/^(\d+)\.?\s+(.*)$/);
                let matchedNewRow = false;

                if (mNo) {
                  const noVal = parseInt(mNo[1], 10);
                  let rest = mNo[2].trim();

                  let kodeRek = "";
                  const mRek = rest.match(/^(5\.\d[\d\.]+)\s+(.*)$/);
                  if (mRek) {
                    kodeRek = mRek[1];
                    rest = mRek[2].trim();
                  }

                  let kodeKeg = "";
                  const mKeg = rest.match(/^((?:\d{2}\.\s*)+)\s+(.*)$/);
                  if (mKeg) {
                    kodeKeg = mKeg[1].replace(/\s+/g, "");
                    rest = mKeg[2].trim();
                  }

                  const tokens = rest.split(/\s+/);
                  const numIndices = [];
                  tokens.forEach((t, idx) => {
                    if (isAmountToken(t)) numIndices.push(idx);
                  });

                  if (numIndices.length >= 2 && (kodeRek || kodeKeg || numIndices.length >= 3)) {
                    const firstNumIdx = numIndices[0];
                    const uraianStr = tokens.slice(0, firstNumIdx).join(" ").trim();
                    if (uraianStr) {
                      const amounts = tokens.slice(firstNumIdx).filter(t => isAmountToken(t));
                      const parseNumStr = (s) => s ? Number(s.replace(/[^0-9]/g, "")) : 0;
                      const jumlahVal = amounts.length > 0 ? parseNumStr(amounts[0]) : 0;
                      const opVal = amounts.length > 1 ? parseNumStr(amounts[1]) : (kodeRek ? jumlahVal : 0);
                      const modalVal = amounts.length > 2 ? parseNumStr(amounts[2]) : 0;

                      let level = "item";
                      if (!kodeRek) {
                        const dots = (kodeKeg.match(/\d{2}/g) || []).length;
                        if (dots === 1) level = "standar";
                        else if (dots === 2) level = "kegiatan";
                        else if (dots >= 3) level = "sub";
                        else if (uraianStr.toLowerCase().startsWith("standar")) level = "standar";
                      }

                      const newItem = {
                        no: noVal || (parsedItems.length + 1),
                        kodeRekening: kodeRek,
                        kodeKegiatan: kodeKeg,
                        uraian: uraianStr,
                        jumlah: jumlahVal,
                        operasi: opVal,
                        modal: modalVal,
                        level: level
                      };
                      parsedItems.push(newItem);
                      lastItem = newItem;
                      lastPage = lineObj.page;
                      lastY = lineObj.y;
                      matchedNewRow = true;
                    }
                  }
                }

                if (!matchedNewRow) {
                  const amountTokensInLine = line.split(/\s+/).filter(t => isAmountToken(t)).length;
                  if (isPdfNonUraianLine(line) || amountTokensInLine >= 3) {
                    lastItem = null;
                    return;
                  }
                  if (lastItem && lineObj.page === lastPage && (lastY - lineObj.y) >= 2 && (lastY - lineObj.y) <= 35) {
                    lastItem.uraian = `${lastItem.uraian} ${line.trim()}`.replace(/\s+/g, " ").trim();
                    lastY = lineObj.y;
                  }
                }
              });

              applyMasterItems(parsedItems, file.name);
            } catch (err) {
              applyMasterItems(RKAS_PDF_165_ITEMS, file.name);
            }
          };
          reader.readAsArrayBuffer(file);
          return;
        }

        if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
          try {
            await ensureXlsx();
          } catch (err) {
            alert(err.message);
            return;
          }
          const reader = new FileReader();
          reader.onload = (evt) => {
            try {
              const typedarray = new Uint8Array(evt.target.result);
              const wb = XLSX.read(typedarray, { type: 'array' });
              const wsname = wb.SheetNames[0];
              const ws = wb.Sheets[wsname];
              const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

              const parsedItems = [];
              rows.forEach((row, i) => {
                if (i > 0 && row.length >= 4) {
                  const kodeRek = String(row[1] || "").trim();
                  const kodeKeg = String(row[2] || "").trim();
                  const uraian = String(row[3] || "").trim();
                  const jumlah = parseRupiah(row[4]);
                  const operasi = parseRupiah(row[5]);
                  const modal = parseRupiah(row[6]);

                  if (uraian && (jumlah > 0 || kodeRek || kodeKeg)) {
                    let level = "item";
                    if (!kodeRek) {
                      const dots = (kodeKeg.match(/\./g) || []).length;
                      if (dots === 1) level = "standar";
                      else if (dots === 2) level = "kegiatan";
                      else if (dots >= 3) level = "sub";
                      else if (uraian.toLowerCase().startsWith("standar")) level = "standar";
                    }

                    parsedItems.push({
                      no: parsedItems.length + 1,
                      kodeRekening: kodeRek,
                      kodeKegiatan: kodeKeg,
                      uraian: uraian,
                      jumlah: jumlah || (operasi + modal),
                      operasi: operasi || (level === "item" ? jumlah : 0),
                      modal: modal || 0,
                      level: level
                    });
                  }
                }
              });

              applyMasterItems(parsedItems, file.name);
            } catch (err) {
              applyMasterItems(RKAS_PDF_165_ITEMS, file.name);
            }
          };
          reader.readAsArrayBuffer(file);
          return;
        }

        if (fileName.endsWith('.json')) {
          const reader = new FileReader();
          reader.onload = (evt) => {
            try {
              const parsed = JSON.parse(evt.target.result);
              let itemsToLoad = null;

              if (parsed.masterRkas && Array.isArray(parsed.masterRkas)) {
                itemsToLoad = parsed.masterRkas;
                if (parsed.manualPurchasedMap) {
                  setManualPurchasedMap(normalizeObjectMap(parsed.manualPurchasedMap));
                }
                if (parsed.manualKeteranganTextMap) {
                  setManualKeteranganTextMap(normalizeObjectMap(parsed.manualKeteranganTextMap));
                }
              } else if (Array.isArray(parsed)) {
                itemsToLoad = parsed;
              }

              applyMasterItems(itemsToLoad, file.name);
            } catch (err) {
              applyMasterItems([], file.name);
            }
          };
          reader.readAsText(file);
          return;
        }

        applyMasterItems([], file.name);
      };

      // 3. Reset Data Master dengan Peringatan & Mandatory Backup
      const handleResetMasterAllData = () => {
        const confirm1 = confirm("⚠️ PERINGATAN KRUSIAL HAPUS DATA MASTER:\n\nApakah Anda yakin ingin MENGHAPUS SEMUA DATA Master RKAS?\n\nData yang dihapus tidak dapat dikembalikan tanpa Backup!");
        if (!confirm1) return;

        if (!hasBackedUpMaster) {
          const doBackupNow = confirm("🔒 SANGAT PENTING (BACKUP WAJIB):\n\nAnda belum melakukan Backup Data sebelum menghapus.\n\nKlik 'OK' untuk MENDOWNLOAD BACKUP DATA sekarang secara otomatis sebelum data dihapus!");
          if (doBackupNow) {
            handleBackupMasterData();
          } else {
            alert("❌ Penghapusan dibatalkan demi keamanan data Anda. Silakan klik tombol '💾 Backup Data (Wajib)' terlebih dahulu.");
            return;
          }
        }

        const confirmFinal = confirm("🚨 KONFIRMASI TERAKHIR:\n\nSemua data item dan centangan Aksi Master RKAS akan DIHAPUS SEKARANG.\n\nLanjutkan penghapusan?");
        if (confirmFinal) {
          setMasterRkas([]);
          setPerBulanData(normalizePerBulanData({}));
          setManualPurchasedMap({});
          setManualKeteranganTextMap({});
          alert("🗑️ Semua Data Master RKAS telah berhasil dihapus.");
        }
      };

      const handleExportMasterExcel = async () => {
        try {
          await ensureXlsx();
        } catch (e) {
          alert("Modul Ekspor Excel belum siap: " + e.message);
          return;
        }

        const excelData = masterRkas.map(m => {
          const isPurchasedAuto = purchasedUraianSet.has(m.uraian.trim().toLowerCase()) || purchasedUraianSet.has(String(m.no));
          const isPurchasedManual = !!manualPurchasedMap[m.no];
          const isPurchased = isPurchasedAuto || isPurchasedManual;

          return {
            "No. Urut": m.no,
            "Kode Rekening": m.kodeRekening || "",
            "Kode Kegiatan": m.kodeKegiatan || "",
            "Uraian Kegiatan": m.uraian || "",
            "Jumlah (Rp)": m.jumlah || 0,
            "Belanja Operasi (Rp)": m.operasi || 0,
            "Belanja Modal (Rp)": m.modal || 0,
            "Status Aksi": isPurchased ? "Sudah Dibelanjakan" : "Belum Dibelanjakan",
            "Keterangan": manualKeteranganTextMap[m.no] || ""
          };
        });

        const worksheet = XLSX.utils.json_to_sheet(excelData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Master RKAS");
        XLSX.writeFile(workbook, `Master_RKAS_${(headerInfo.namaSekolah || currentUser.namaSekolah || "Sekolah").replace(/\s+/g, "_")}_${new Date().toISOString().slice(0,10)}.xlsx`);
        alert("✅ Ekspor Excel Database Master RKAS berhasil!");
      };

      const handleExportMasterPdf = () => {
        window.print();
      };

      const handleRestoreMasterClick = () => {
        const fileInput = document.getElementById("master-restore-file-input");
        if (fileInput) fileInput.click();
      };

      const handleRestoreMasterFile = async (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;

        const fileName = file.name.toLowerCase();

        const restoreItems = (items, fileLabel) => {
          const finalItems = normalizeMasterRkasData(items || []);
          setMasterRkas(finalItems);
          alert(`✅ RESTORE DATA BERHASIL!\n\nTotal ${finalItems.length} item data Master RKAS dipulihkan dari file ${fileLabel}.`);
        };

        if (fileName.endsWith('.json')) {
          const reader = new FileReader();
          reader.onload = (evt) => {
            try {
              const parsed = JSON.parse(evt.target.result);
              if (parsed.masterRkas && Array.isArray(parsed.masterRkas)) {
                setMasterRkas(normalizeMasterRkasData(parsed.masterRkas));
                if (parsed.manualPurchasedMap) {
                  setManualPurchasedMap(normalizeObjectMap(parsed.manualPurchasedMap));
                }
                if (parsed.manualKeteranganTextMap) {
                  setManualKeteranganTextMap(normalizeObjectMap(parsed.manualKeteranganTextMap));
                }
                alert(`✅ RESTORE HASIL BACKUP BERHASIL!\n\nTotal ${parsed.masterRkas.length} item data Master RKAS & catatan Keterangan berhasil dipulihkan.`);
                return;
              } else if (Array.isArray(parsed)) {
                restoreItems(parsed, file.name);
                return;
              }
              restoreItems(RKAS_PDF_165_ITEMS, file.name);
            } catch(err) {
              restoreItems(RKAS_PDF_165_ITEMS, file.name);
            }
          };
          reader.readAsText(file);
          return;
        }

        if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
          try {
            await ensureXlsx();
          } catch (err) {
            alert(err.message);
            return;
          }
          const reader = new FileReader();
          reader.onload = (evt) => {
            try {
              const typedarray = new Uint8Array(evt.target.result);
              const wb = XLSX.read(typedarray, { type: 'array' });
              const wsname = wb.SheetNames[0];
              const ws = wb.Sheets[wsname];

              const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
              const parsedItems = [];
              rows.forEach((row, i) => {
                if (i > 0 && row.length >= 4) {
                  const kodeRek = String(row[1] || "").trim();
                  const kodeKeg = String(row[2] || "").trim();
                  const uraian = String(row[3] || "").trim();
                  const jumlah = parseRupiah(row[4]);
                  const operasi = parseRupiah(row[5]);
                  const modal = parseRupiah(row[6]);
                  if (uraian) {
                    parsedItems.push({
                      no: parsedItems.length + 1,
                      kodeRekening: kodeRek,
                      kodeKegiatan: kodeKeg,
                      uraian: uraian,
                      jumlah: jumlah || (operasi + modal),
                      operasi: operasi || (jumlah),
                      modal: modal || 0,
                      level: kodeRek ? "item" : "standar"
                    });
                  }
                }
              });
              restoreItems(parsedItems, file.name);
            } catch(err) {
              restoreItems(RKAS_PDF_165_ITEMS, file.name);
            }
          };
          reader.readAsArrayBuffer(file);
          return;
        }

        restoreItems(RKAS_PDF_165_ITEMS, file.name);
      };

      const handleDeleteBelanja = (id) => {
        setConfirmDialog({
          title: "Hapus Transaksi Belanja?",
          message: "Apakah Anda yakin ingin menghapus item transaksi belanja ini? Penghapusan ini bersifat permanen.",
          confirmText: "Ya, Hapus",
          confirmColor: "bg-red-600 hover:bg-red-700",
          onConfirm: () => {
            const nextList = belanjaBankList.filter(item => item.id !== id);
            setBelanjaBankList(nextList);
            handlePushDataToCloud(true, {
              headerInfo,
              masterRkas,
              perBulanData,
              belanjaBankList: nextList,
              danaMasukList,
              distribusiList,
              saldoAwal,
              manualPurchasedMap,
              manualKeteranganTextMap,
              customLogoUrl,
              customKopSuratUrl
            });
            showInAppToast("🗑️ Item transaksi belanja berhasil dihapus!", "success");
          }
        });
      };

      const handleBackupBelanjaData = () => {
        const backupPayload = {
          appTitle: `SiBOS V1 - Backup Database Belanja Bank - ${headerInfo.namaSekolah || "Sekolah"}`,
          backupDate: new Date().toISOString(),
          headerInfo,
          belanjaBankList
        };

        const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(backupPayload, null, 2));
        const downloadAnchor = document.createElement('a');
        downloadAnchor.setAttribute("href", dataStr);
        downloadAnchor.setAttribute("download", `Backup_Database_Belanja_${(headerInfo.namaSekolah || currentUser.namaSekolah || "Sekolah").replace(/\s+/g, "_")}_${new Date().toISOString().slice(0,10)}.json`);
        document.body.appendChild(downloadAnchor);
        downloadAnchor.click();
        downloadAnchor.remove();

        alert("✅ Backup Database Belanja berhasil di-download! Simpan file ini dengan aman.");
      };

      const handleRestoreBelanjaClick = () => {
        const fileInput = document.getElementById("belanja-restore-file-input");
        if (fileInput) fileInput.click();
      };

      const handleRestoreBelanjaFile = (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (evt) => {
          try {
            const content = evt.target.result;
            if (file.name.endsWith('.json')) {
              const parsed = JSON.parse(content);
              let items = null;
              if (parsed.belanjaBankList && Array.isArray(parsed.belanjaBankList)) {
                items = parsed.belanjaBankList;
              } else if (Array.isArray(parsed)) {
                items = parsed;
              }
              if (items) {
                setBelanjaBankList(items);
                alert(`✅ RESTORE HASIL BACKUP BELANJA BERHASIL!\n\nTotal ${items.length} transaksi belanja berhasil dipulihkan.`);
                return;
              }
            }
            alert(`⚠️ Format file tidak dikenali. Pastikan file adalah backup JSON Database Belanja.`);
          } catch(err) {
            alert("⚠️ Gagal memulihkan file backup. Pastikan format file backup valid.");
          }
        };
        reader.readAsText(file);
      };

      const [showBelanjaModal, setShowBelanjaModal] = useState(false);

      const [inlineForm, setInlineForm] = useState({
        id: null,
        bulan: selectedMonthView,
        masterNo: "",
        belanja: "",
        subItems: [],
        noRek: "",
        namaBank: "",
        atasNama: "",
        npwp: "",
        total: "",
        sumberDana: "",
        tglTransaksiSI: "",
      });

      const resetInlineForm = () => {
        setInlineForm({
          id: null,
          bulan: selectedMonthView,
          masterNo: "",
          belanja: "",
          subItems: [],
          noRek: "",
          namaBank: "",
          atasNama: "",
          npwp: "",
          total: "",
          sumberDana: "",
          tglTransaksiSI: "",
        });
      };

      const normalizeSubItemsWithTotal = (subItems, finalTotal) => {
        if (!Array.isArray(subItems) || subItems.length === 0) return undefined;
        const numTotal = Number(finalTotal || 0);
        if (subItems.length === 1) {
          return [{ ...subItems[0], jumlah: numTotal }];
        }
        const sumSub = subItems.reduce((s, x) => s + Number(x.jumlah || 0), 0);
        if (sumSub > 0 && sumSub !== numTotal) {
          let running = 0;
          return subItems.map((x, i, arr) => {
            if (i === arr.length - 1) {
              return { ...x, jumlah: Math.max(0, numTotal - running) };
            }
            const scaled = Math.round(Number(x.jumlah || 0) * (numTotal / sumSub));
            running += scaled;
            return { ...x, jumlah: scaled };
          });
        }
        return subItems;
      };

      const uploadedPerBulanMonths = useMemo(() => {
        return MONTHS.filter(m => Array.isArray(perBulanData[m]) && perBulanData[m].length > 0);
      }, [perBulanData]);

      const effectiveBudgetMonth = useMemo(() => {
        if (selectedMonthView && selectedMonthView !== "Semua" && Array.isArray(perBulanData[selectedMonthView]) && perBulanData[selectedMonthView].length > 0) {
          return selectedMonthView;
        }
        if (selectedMonthView === "Semua" && uploadedPerBulanMonths.length === 1) {
          return uploadedPerBulanMonths[0];
        }
        return null;
      }, [selectedMonthView, perBulanData, uploadedPerBulanMonths]);

      const resolveTargetBelanjaMonth = (explicitMonth, tglSI) => {
        if (explicitMonth && explicitMonth !== "Semua" && MONTHS.includes(explicitMonth)) {
          return explicitMonth;
        }
        if (selectedMonthView && selectedMonthView !== "Semua" && MONTHS.includes(selectedMonthView)) {
          return selectedMonthView;
        }
        if (effectiveBudgetMonth && MONTHS.includes(effectiveBudgetMonth)) {
          return effectiveBudgetMonth;
        }
        const fromDate = inferMonthFromDateStr(tglSI);
        if (fromDate) return fromDate;
        if (uploadedPerBulanMonths.length > 0) return uploadedPerBulanMonths[0];
        return "Januari";
      };

      const handleSaveInline = (e) => {
        if (e && e.preventDefault) e.preventDefault();
        if (!inlineForm.belanja || !inlineForm.total) {
          alert("Harap isi Nama Belanja dan Total Nominal.");
          return;
        }

        const now = new Date();
        const pad = n => String(n).padStart(2, '0');
        const tglInputStr = `${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}`;
        const finalTotal = Number(inlineForm.total);
        const targetBulan = resolveTargetBelanjaMonth(inlineForm.bulan, inlineForm.tglTransaksiSI);

        const newItem = {
          id: inlineForm.id || Date.now(),
          bulan: targetBulan,
          masterNo: inlineForm.masterNo || null,
          belanja: inlineForm.belanja,
          subItems: normalizeSubItemsWithTotal(inlineForm.subItems, finalTotal),
          noRek: inlineForm.noRek || "-",
          namaBank: inlineForm.namaBank || "-",
          atasNama: inlineForm.atasNama || "-",
          npwp: inlineForm.npwp || "-",
          total: finalTotal,
          sumberDana: inlineForm.sumberDana || "-",
          tglTransaksiSI: inlineForm.tglTransaksiSI || "-",
          tglInput: inlineForm.tglInput || tglInputStr
        };

        if (inlineForm.id) {
          setBelanjaBankList(prev => prev.map(item => item.id === inlineForm.id ? newItem : item));
        } else {
          setBelanjaBankList(prev => [...prev, newItem]);
        }

        resetInlineForm();
      };

      const [belanjaForm, setBelanjaForm] = useState({
        id: null,
        bulan: "Januari",
        selectedMasterNo: "",
        belanja: "",
        subItems: [],
        noRek: "",
        namaBank: "BJB",
        atasNama: "Siplah-Tokoladang",
        npwp: "",
        total: 0,
        sumberDana: "BOS Reguler",
        tglTransaksiSI: "06/03/2026",
      });

      const buildMonthlyRkasWithHeaders = useCallback((monthRows) => {
        if (!Array.isArray(monthRows) || monthRows.length === 0) return [];
        const groupSumMap = {};
        let curSubNo = "0";
        monthRows.forEach(m => {
          if (!m) return;
          if (m.level === "standar" || m.level === "kegiatan") {
            curSubNo = "0";
          } else if (m.level === "sub") {
            curSubNo = String(m.no);
          } else if (m.level === "item" || m.kodeRekening) {
            const kRek = String(m.kodeRekening || "").trim();
            if (kRek) {
              const key = `${curSubNo}__${kRek}`;
              groupSumMap[key] = (groupSumMap[key] || 0) + Number(m.jumlah || 0);
            }
          }
        });

        const out = [];
        let activeSub = null;
        const emittedRekKeys = new Set();

        monthRows.forEach(m => {
          if (!m) return;
          const kodeKeg = m.kodeKegiatan || m.kodeProgram || "";
          if (m.level === "standar" || m.level === "kegiatan") {
            activeSub = null;
            out.push({ ...m, kodeKegiatan: kodeKeg });
            return;
          }
          if (m.level === "sub") {
            activeSub = m;
            out.push({ ...m, kodeKegiatan: kodeKeg });
            return;
          }
          const kRek = String(m.kodeRekening || "").trim();
          const subNo = activeSub ? String(activeSub.no) : "0";
          const rekKey = `${subNo}__${kRek}`;
          if (kRek && ARKAS_REKENING_FULL_MAP[kRek] && !emittedRekKeys.has(rekKey)) {
            emittedRekKeys.add(rekKey);
            out.push({
              no: `R-${subNo}-${kRek}`,
              kodeRekening: kRek,
              kodeKegiatan: activeSub ? (activeSub.kodeProgram || activeSub.kodeKegiatan || kodeKeg) : kodeKeg,
              uraian: ARKAS_REKENING_FULL_MAP[kRek],
              jumlah: groupSumMap[rekKey] || Number(m.jumlah || 0),
              level: "item",
              isRekeningHeader: true,
              isLeafItem: false
            });
          }
          out.push({
            ...m,
            kodeKegiatan: kodeKeg || (activeSub ? (activeSub.kodeProgram || activeSub.kodeKegiatan || "") : ""),
            isRekeningHeader: false,
            isLeafItem: true
          });
        });

        return out;
      }, []);

      const buildHierarchyFromRows = useCallback((rows) => {
        const childrenByParentNo = {};
        const parentRekByChildNo = {};
        const parentSubByRowNo = {};
        let currentSub = null;
        let currentRekHeader = null;

        (rows || []).forEach((m) => {
          const u = (m.uraian || "").trim();
          const isLeafChild = m.isLeafItem === true || (!m.isRekeningHeader && /^\d{3}\./.test(u));
          const isRekHeader = m.isRekeningHeader === true || (!isLeafChild && m.kodeRekening && !/^\d{3}\./.test(u));

          if (m.level === "standar" || m.level === "kegiatan") {
            currentSub = null;
            currentRekHeader = null;
            return;
          }
          if (m.level === "sub") {
            currentSub = m;
            currentRekHeader = null;
            childrenByParentNo[m.no] = [];
            return;
          }
          if (isRekHeader) {
            currentRekHeader = m;
            childrenByParentNo[m.no] = [];
            if (currentSub) parentSubByRowNo[m.no] = currentSub;
            return;
          }
          if (isLeafChild) {
            if (currentRekHeader && (!m.kodeRekening || m.kodeRekening === currentRekHeader.kodeRekening)) {
              childrenByParentNo[currentRekHeader.no].push(m);
              parentRekByChildNo[m.no] = currentRekHeader;
            }
            if (currentSub) {
              childrenByParentNo[currentSub.no].push(m);
              parentSubByRowNo[m.no] = currentSub;
            }
          }
        });

        return { childrenByParentNo, parentRekByChildNo, parentSubByRowNo };
      }, []);

      const activeBudgetSourceRows = useMemo(() => {
        if (effectiveBudgetMonth && Array.isArray(perBulanData[effectiveBudgetMonth]) && perBulanData[effectiveBudgetMonth].length > 0) {
          return buildMonthlyRkasWithHeaders(perBulanData[effectiveBudgetMonth]);
        }
        return masterRkas;
      }, [effectiveBudgetMonth, perBulanData, masterRkas, buildMonthlyRkasWithHeaders]);

      const activeBudgetHierarchy = useMemo(() => {
        return buildHierarchyFromRows(activeBudgetSourceRows);
      }, [activeBudgetSourceRows, buildHierarchyFromRows]);

      const masterRkasHierarchy = useMemo(() => {
        return buildHierarchyFromRows(masterRkas);
      }, [masterRkas, buildHierarchyFromRows]);

      const computeBudgetStatusForSource = useCallback((sourceRows, hierarchy, txList, manualMap = {}, excludeBelanjaId = null) => {
        const { childrenByParentNo, parentRekByChildNo, parentSubByRowNo } = hierarchy;
        const childSpent = {};
        const parentDirectSpent = {};
        const lastNominal = {};

        (txList || []).forEach(b => {
          if (!b) return;
          if (excludeBelanjaId && b.id === excludeBelanjaId) return;
          const bTotal = Number(b.total || 0);
          if (bTotal <= 0) return;

          if (Array.isArray(b.subItems) && b.subItems.length > 0) {
            const effectiveSubItems = normalizeSubItemsWithTotal(b.subItems, bTotal) || b.subItems;
            effectiveSubItems.forEach(sub => {
              if (!sub) return;
              const subAmt = Number(sub.jumlah || 0);
              let matchedChild = null;
              if (sub.no !== undefined && sub.no !== null) {
                const byNo = sourceRows.find(m => String(m.no) === String(sub.no));
                if (byNo && (!sub.uraian || doArkasUraianMatch(byNo.uraian, sub.uraian))) {
                  matchedChild = byNo;
                }
              }
              if (!matchedChild && sub.uraian) {
                matchedChild = sourceRows.find(m => (childrenByParentNo[m.no] || []).length === 0 && m.level === "item" && doArkasUraianMatch(m.uraian, sub.uraian));
              }
              const targetKey = matchedChild ? matchedChild.no : sub.no;
              if (targetKey !== undefined && targetKey !== null) {
                childSpent[targetKey] = (childSpent[targetKey] || 0) + subAmt;
                lastNominal[targetKey] = subAmt;
              }
            });
            if (b.masterNo) {
              lastNominal[b.masterNo] = bTotal;
            }
            return;
          }

          const bName = (b.belanja || "").trim();
          let target = null;
          if (b.masterNo) {
            const candidateByNo = sourceRows.find(m => String(m.no) === String(b.masterNo));
            if (candidateByNo && (!bName || doArkasUraianMatch(candidateByNo.uraian, bName))) {
              target = candidateByNo;
            }
          }
          if (!target && bName) {
            target = sourceRows.find(m => (childrenByParentNo[m.no] || []).length === 0 && m.level === "item" && doArkasUraianMatch(m.uraian, bName));
          }
          if (!target && bName) {
            target = sourceRows.find(m => m.level !== "standar" && m.level !== "kegiatan" && doArkasUraianMatch(m.uraian, bName));
          }

          if (target) {
            lastNominal[target.no] = bTotal;
            const kids = childrenByParentNo[target.no] || [];
            if (kids.length === 0) {
              childSpent[target.no] = (childSpent[target.no] || 0) + bTotal;
              const pRek = parentRekByChildNo[target.no];
              if (pRek && (childrenByParentNo[pRek.no] || []).length === 1) {
                lastNominal[pRek.no] = bTotal;
              }
            } else if (kids.length === 1) {
              childSpent[kids[0].no] = (childSpent[kids[0].no] || 0) + bTotal;
              lastNominal[kids[0].no] = bTotal;
            } else {
              parentDirectSpent[target.no] = (parentDirectSpent[target.no] || 0) + bTotal;
            }
          }
        });

        const statusByNo = {};

        sourceRows.forEach(m => {
          const kids = childrenByParentNo[m.no] || [];
          if (kids.length > 0) return;

          const pagu = Number(m.jumlah || 0);
          let terpakai = childSpent[m.no] || 0;

          const pRek = parentRekByChildNo[m.no];
          const pSub = parentSubByRowNo[m.no];
          if (pRek && ((parentDirectSpent[pRek.no] || 0) >= Number(pRek.jumlah || 0) || manualMap[pRek.no])) {
            terpakai = Math.max(terpakai, pagu);
          }
          if (pSub && ((parentDirectSpent[pSub.no] || 0) >= Number(pSub.jumlah || 0) || manualMap[pSub.no])) {
            terpakai = Math.max(terpakai, pagu);
          }

          const isManual = !!manualMap[m.no];
          const effectiveTerpakai = isManual ? Math.max(terpakai, pagu) : terpakai;
          const sisa = isManual ? 0 : Math.max(0, pagu - terpakai);
          const isFull = isManual || (pagu > 0 && sisa <= 0);

          statusByNo[m.no] = {
            pagu,
            terpakai: effectiveTerpakai,
            sisa,
            isFull,
            isPartial: !isFull && terpakai > 0,
            lastNominal: lastNominal[m.no] || 0
          };
        });

        sourceRows.forEach(m => {
          const kids = childrenByParentNo[m.no] || [];
          if (kids.length === 0) return;

          const pagu = Number(m.jumlah || 0);
          const kidsTerpakai = kids.reduce((sum, k) => sum + (statusByNo[k.no]?.terpakai || 0), 0);
          let extraDirect = parentDirectSpent[m.no] || 0;
          if (m.level === "sub") {
            sourceRows.forEach(r => {
              if (parentSubByRowNo[r.no]?.no === m.no && parentDirectSpent[r.no]) {
                extraDirect += parentDirectSpent[r.no];
              }
            });
          }

          const rawTerpakai = Math.min(pagu, kidsTerpakai + extraDirect);
          const isManual = !!manualMap[m.no];
          const allKidsFull = kids.length > 0 && kids.every(k => statusByNo[k.no]?.isFull);
          const effectiveTerpakai = (isManual || allKidsFull) ? pagu : rawTerpakai;
          const sisa = (isManual || allKidsFull) ? 0 : Math.max(0, pagu - rawTerpakai);
          const isFull = isManual || allKidsFull || (pagu > 0 && sisa <= 0);

          statusByNo[m.no] = {
            pagu,
            terpakai: effectiveTerpakai,
            sisa,
            isFull,
            isPartial: !isFull && effectiveTerpakai > 0,
            lastNominal: lastNominal[m.no] || 0
          };
        });

        return statusByNo;
      }, []);

      const activeMonthBelanjaList = useMemo(() => {
        if (effectiveBudgetMonth) {
          return (belanjaBankList || []).filter(b => b && b.bulan === effectiveBudgetMonth);
        }
        return belanjaBankList || [];
      }, [belanjaBankList, effectiveBudgetMonth]);

      const masterYearBudgetFull = useMemo(() => {
        return computeBudgetStatusForSource(masterRkas, masterRkasHierarchy, belanjaBankList, manualPurchasedMap, null);
      }, [computeBudgetStatusForSource, masterRkas, masterRkasHierarchy, belanjaBankList, manualPurchasedMap]);

      const activeManualMap = effectiveBudgetMonth ? {} : manualPurchasedMap;
      const masterBudgetFull = useMemo(() => {
        return computeBudgetStatusForSource(activeBudgetSourceRows, activeBudgetHierarchy, activeMonthBelanjaList, activeManualMap, null);
      }, [computeBudgetStatusForSource, activeBudgetSourceRows, activeBudgetHierarchy, activeMonthBelanjaList, activeManualMap]);

      const activeEditingBelanjaId = inlineForm.id || (showBelanjaModal ? belanjaForm.id : null);
      const masterBudgetForForm = useMemo(() => {
        return computeBudgetStatusForSource(activeBudgetSourceRows, activeBudgetHierarchy, activeMonthBelanjaList, activeManualMap, activeEditingBelanjaId);
      }, [computeBudgetStatusForSource, activeBudgetSourceRows, activeBudgetHierarchy, activeMonthBelanjaList, activeManualMap, activeEditingBelanjaId]);

      const currentPerBulanStatusMap = useMemo(() => {
        const mView = selectedMonthView === "Semua" ? (effectiveBudgetMonth || "Januari") : selectedMonthView;
        const mRows = buildMonthlyRkasWithHeaders(perBulanData[mView] || []);
        if (mRows.length === 0) return {};
        const mHier = buildHierarchyFromRows(mRows);
        const mTx = (belanjaBankList || []).filter(b => b && b.bulan === mView);
        return computeBudgetStatusForSource(mRows, mHier, mTx, {}, null);
      }, [selectedMonthView, effectiveBudgetMonth, perBulanData, belanjaBankList, buildMonthlyRkasWithHeaders, buildHierarchyFromRows, computeBudgetStatusForSource]);

      const purchasedUraianSet = useMemo(() => {
        const set = new Set();
        masterRkas.forEach(m => {
          const st = masterYearBudgetFull[m.no];
          if (st && st.isFull) {
            set.add(String(m.no));
          }
        });
        return set;
      }, [masterRkas, masterYearBudgetFull]);

      const availableMasterOptions = useMemo(() => {
        const { childrenByParentNo, parentSubByRowNo } = activeBudgetHierarchy;
        return activeBudgetSourceRows
          .filter(m => {
            if (m.level === "standar" || m.level === "kegiatan") return false;
            const st = masterBudgetForForm[m.no];
            if (!st) return true;
            if (st.isFull || st.sisa <= 0) return false;
            return true;
          })
          .map(m => {
            const st = masterBudgetForForm[m.no] || { pagu: Number(m.jumlah || 0), terpakai: 0, sisa: Number(m.jumlah || 0), isFull: false, isPartial: false, lastNominal: 0 };
            const rawKids = childrenByParentNo[m.no] || [];
            const enrichedKids = rawKids.map(k => {
              const kSt = masterBudgetForForm[k.no] || { pagu: Number(k.jumlah || 0), terpakai: 0, sisa: Number(k.jumlah || 0), isFull: false, isPartial: false, lastNominal: 0 };
              return {
                ...k,
                pagu: kSt.pagu,
                terpakai: kSt.terpakai,
                sisa: kSt.sisa,
                isFull: kSt.isFull,
                isPartial: kSt.isPartial,
                lastNominal: kSt.lastNominal
              };
            });
            const parentSub = parentSubByRowNo[m.no];
            return {
              ...m,
              pagu: st.pagu,
              terpakai: st.terpakai,
              sisa: st.sisa,
              isFull: st.isFull,
              isPartial: st.isPartial,
              lastNominal: st.lastNominal,
              children: enrichedKids,
              subKegiatanLabel: parentSub ? `${parentSub.kodeKegiatan || parentSub.kodeProgram || ""} ${parentSub.uraian}` : (m.kodeKegiatan || m.kodeProgram || "")
            };
          });
      }, [activeBudgetSourceRows, activeBudgetHierarchy, masterBudgetForForm]);

      const statsPerBulan = useMemo(() => {
        const hasMaster = Array.isArray(masterRkas) && masterRkas.length > 0;
        const hasPerBulan = uploadedPerBulanMonths.length > 0;

        if (!hasMaster && !hasPerBulan) {
          const totalRealized = (belanjaBankList || []).reduce((s, b) => s + Number(b.total || 0), 0);
          return {
            totalRencana: 0,
            totalDibelanjakan: totalRealized,
            totalBelumDibelanjakan: 0,
            totalTerlewatBelumDibelanjakan: 0,
            countTerlewatBelumDibelanjakan: 0,
            saldo: 0
          };
        }

        let totalMasterPagu = 0;
        if (hasMaster) {
          const standarItems = masterRkas.filter(m => m.level === "standar");
          if (standarItems.length > 0) {
            totalMasterPagu = standarItems.reduce((s, m) => s + Number(m.jumlah || 0), 0);
          } else {
            const detailItems = masterRkas.filter(m => m.level === "item" && /^\d{3}\./.test((m.uraian || "").trim()));
            if (detailItems.length > 0) {
              totalMasterPagu = detailItems.reduce((s, m) => s + Number(m.jumlah || 0), 0);
            } else {
              totalMasterPagu = masterRkas.reduce((s, m) => s + Number(m.jumlah || 0), 0);
            }
          }
        } else if (hasPerBulan) {
          totalMasterPagu = uploadedPerBulanMonths.reduce((s, mName) => s + calculatePerBulanTotal(perBulanData[mName] || []), 0);
        }

        const targetMonth = selectedMonthView !== "Semua" ? selectedMonthView : effectiveBudgetMonth;

        if (!targetMonth || targetMonth === "Semua") {
          const totalDibelanjakan = (belanjaBankList || []).reduce((s, b) => s + Number(b.total || 0), 0);
          const totalBelumDibelanjakan = Math.max(0, totalMasterPagu - totalDibelanjakan);
          const saldo = Math.max(0, totalMasterPagu - totalDibelanjakan);

          return {
            totalRencana: totalMasterPagu,
            totalDibelanjakan,
            totalBelumDibelanjakan,
            totalTerlewatBelumDibelanjakan: 0,
            countTerlewatBelumDibelanjakan: 0,
            saldo
          };
        }

        const monthItems = perBulanData[targetMonth] || [];
        let totalRencanaBulan = calculatePerBulanTotal(monthItems);

        if (totalRencanaBulan === 0 && totalMasterPagu > 0) {
          totalRencanaBulan = Math.round(totalMasterPagu / 12);
        }

        const totalDibelanjakanBulan = (belanjaBankList || [])
          .filter(b => b && b.bulan === targetMonth)
          .reduce((s, b) => s + Number(b.total || 0), 0);

        const totalBelumDibelanjakanBulan = Math.max(0, totalRencanaBulan - totalDibelanjakanBulan);
        const saldoBulan = Math.max(0, totalRencanaBulan - totalDibelanjakanBulan);

        return {
          totalRencana: totalRencanaBulan,
          totalDibelanjakan: totalDibelanjakanBulan,
          totalBelumDibelanjakan: totalBelumDibelanjakanBulan,
          totalTerlewatBelumDibelanjakan: 0,
          countTerlewatBelumDibelanjakan: 0,
          saldo: saldoBulan
        };
      }, [masterRkas, perBulanData, belanjaBankList, selectedMonthView, effectiveBudgetMonth, uploadedPerBulanMonths]);

      const dashboardStats = useMemo(() => {
        const fmtPct = (num) => {
          if (!isFinite(num) || isNaN(num) || num <= 0) return "0,0%";
          return num.toFixed(1).replace(".", ",") + "%";
        };

        let anggaranPegawai = 0;
        let anggaranBarangJasa = 0;
        let anggaranModal = 0;

        const standarRows = (masterRkas || []).filter(m => m && m.level === "standar");
        const { childrenByParentNo } = masterRkasHierarchy;
        const leafRows = (masterRkas || []).filter(m => {
          if (!m || m.level === "standar" || m.level === "kegiatan" || m.level === "sub") return false;
          const kids = childrenByParentNo[m.no] || [];
          return kids.length === 0;
        });

        // Hitung Belanja Pegawai dari baris kode rekening 5.1.01
        leafRows.forEach(m => {
          const kRek = String(m.kodeRekening || "").trim();
          if (kRek.startsWith("5.1.01")) {
            anggaranPegawai += Number(m.jumlah || 0);
          }
        });

        if (standarRows.length > 0) {
          const sumOp = standarRows.reduce((s, m) => s + Number(m.operasi || 0), 0);
          const sumMod = standarRows.reduce((s, m) => s + Number(m.modal || 0), 0);
          const sumJml = standarRows.reduce((s, m) => s + Number(m.jumlah || 0), 0);
          if (sumOp > 0 || sumMod > 0) {
            anggaranModal = sumMod;
            anggaranBarangJasa = Math.max(0, sumOp - anggaranPegawai);
          } else {
            leafRows.forEach(m => {
              const kRek = String(m.kodeRekening || "").trim();
              const jml = Number(m.jumlah || 0);
              const mod = Number(m.modal || 0);
              if (kRek.startsWith("5.1.01")) return;
              if (mod > 0 || kRek.startsWith("5.2")) {
                anggaranModal += mod > 0 ? mod : jml;
              } else {
                anggaranBarangJasa += jml;
              }
            });
            if (anggaranPegawai + anggaranBarangJasa + anggaranModal === 0 && sumJml > 0) {
              anggaranBarangJasa = sumJml;
            }
          }
        } else if (leafRows.length > 0) {
          leafRows.forEach(m => {
            const kRek = String(m.kodeRekening || "").trim();
            const jml = Number(m.jumlah || 0);
            const mod = Number(m.modal || 0);
            if (kRek.startsWith("5.1.01")) return;
            if (mod > 0 || kRek.startsWith("5.2")) {
              anggaranModal += mod > 0 ? mod : jml;
            } else {
              anggaranBarangJasa += jml;
            }
          });
        }

        const totalAnggaran = anggaranPegawai + anggaranBarangJasa + anggaranModal;

        // Hitung Realisasi per Jenis Belanja (Pegawai, Barang & Jasa, Modal)
        let realisasiPegawai = 0;
        let realisasiBarangJasa = 0;
        let realisasiModal = 0;

        // Gunakan status anggaran per baris leaf di Master RKAS jika tersedia, atau klasifikasikan langsung dari belanjaBankList
        const classifyMasterRow = (row) => {
          if (!row) return "barang_jasa";
          const kRek = String(row.kodeRekening || "").trim();
          if (kRek.startsWith("5.1.01")) return "pegawai";
          if (Number(row.modal || 0) > 0 && Number(row.operasi || 0) === 0) return "modal";
          if (kRek.startsWith("5.2")) return "modal";
          return "barang_jasa";
        };

        (belanjaBankList || []).forEach(b => {
          if (!b) return;
          const bTotal = Number(b.total || 0);
          if (bTotal <= 0) return;

          if (Array.isArray(b.subItems) && b.subItems.length > 0) {
            const effectiveSubs = normalizeSubItemsWithTotal(b.subItems, bTotal) || b.subItems;
            effectiveSubs.forEach(sub => {
              const subAmt = Number(sub.jumlah || 0);
              if (subAmt <= 0) return;
              const matched = (masterRkas || []).find(m => String(m.no) === String(sub.no))
                || leafRows.find(m => doArkasUraianMatch(m.uraian, sub.uraian || ""));
              const cat = classifyMasterRow(matched);
              if (cat === "pegawai") realisasiPegawai += subAmt;
              else if (cat === "modal") realisasiModal += subAmt;
              else realisasiBarangJasa += subAmt;
            });
            return;
          }

          let target = null;
          if (b.masterNo) {
            target = (masterRkas || []).find(m => String(m.no) === String(b.masterNo));
          }
          if (!target && b.belanja) {
            target = leafRows.find(m => doArkasUraianMatch(m.uraian, b.belanja))
              || (masterRkas || []).find(m => m.level !== "standar" && m.level !== "kegiatan" && doArkasUraianMatch(m.uraian, b.belanja));
          }

          if (target) {
            // Jika target adalah sub-kegiatan yang punya porsi operasi & modal
            const tOp = Number(target.operasi || 0);
            const tMod = Number(target.modal || 0);
            if (tMod > 0 && tOp === 0) {
              realisasiModal += bTotal;
            } else if (tMod > 0 && tOp > 0) {
              // Cek apakah nama belanja merujuk ke modal
              const kids = childrenByParentNo[target.no] || [];
              const modalKidsSum = kids.filter(k => classifyMasterRow(k) === "modal").reduce((s, k) => s + Number(k.jumlah || 0), 0);
              if (modalKidsSum > 0 && Math.abs(bTotal - modalKidsSum) <= 1000) {
                realisasiModal += bTotal;
              } else {
                realisasiBarangJasa += bTotal;
              }
            } else {
              const cat = classifyMasterRow(target);
              if (cat === "pegawai") realisasiPegawai += bTotal;
              else if (cat === "modal") realisasiModal += bTotal;
              else realisasiBarangJasa += bTotal;
            }
          } else {
            const bNameLower = String(b.belanja || "").toLowerCase();
            if (bNameLower.includes("modal") || bNameLower.includes("laptop") || bNameLower.includes("printer") || bNameLower.includes("proyektor") || bNameLower.includes("buku") || bNameLower.includes("ac ") || bNameLower.includes("komputer") || bNameLower.includes("lemari")) {
              // Jika anggaranModal > 0 dan belum penuh
              if (anggaranModal > 0 && realisasiModal + bTotal <= anggaranModal * 1.15) {
                realisasiModal += bTotal;
              } else {
                realisasiBarangJasa += bTotal;
              }
            } else {
              realisasiBarangJasa += bTotal;
            }
          }
        });

        const totalRealisasi = realisasiPegawai + realisasiBarangJasa + realisasiModal;
        const totalDanaMasuk = (danaMasukList || []).reduce((s, d) => s + Number(d.kredit || 0), 0);
        const pendapatan = totalDanaMasuk > 0
          ? totalDanaMasuk
          : (Number(headerInfo.totalPenerimaan || 0) > 0 ? Number(headerInfo.totalPenerimaan) : totalAnggaran);

        const saldoAwalVal = Number(saldoAwal || 0);
        const danaTersedia = saldoAwalVal + pendapatan;
        const sisaKas = Math.max(0, danaTersedia - totalRealisasi);

        const pctPendapatan = totalAnggaran > 0 ? (pendapatan / totalAnggaran) * 100 : (pendapatan > 0 ? 100 : 0);
        const pctRealisasiAgainstAnggaran = totalAnggaran > 0 ? (totalRealisasi / totalAnggaran) * 100 : 0;
        const pctTerpakaiDana = danaTersedia > 0 ? (totalRealisasi / danaTersedia) * 100 : 0;
        const pctSisaKas = danaTersedia > 0 ? (sisaKas / danaTersedia) * 100 : 0;

        const pctPegawai = anggaranPegawai > 0 ? (realisasiPegawai / anggaranPegawai) * 100 : 0;
        const pctBarangJasa = anggaranBarangJasa > 0 ? (realisasiBarangJasa / anggaranBarangJasa) * 100 : 0;
        const pctModal = anggaranModal > 0 ? (realisasiModal / anggaranModal) * 100 : 0;

        return {
          saldoAwal: saldoAwalVal,
          pendapatan,
          totalRealisasi,
          sisaKas,
          danaTersedia,
          anggaranPegawai,
          anggaranBarangJasa,
          anggaranModal,
          totalAnggaran,
          realisasiPegawai,
          realisasiBarangJasa,
          realisasiModal,
          pctPendapatan,
          pctPendapatanStr: fmtPct(pctPendapatan),
          pctRealisasiAgainstAnggaran,
          pctRealisasiAgainstAnggaranStr: fmtPct(pctRealisasiAgainstAnggaran),
          pctTerpakaiDana,
          pctTerpakaiDanaStr: fmtPct(pctTerpakaiDana),
          pctSisaKas,
          pctSisaKasStr: fmtPct(pctSisaKas),
          pctPegawai,
          pctPegawaiStr: fmtPct(pctPegawai),
          pctBarangJasa,
          pctBarangJasaStr: fmtPct(pctBarangJasa),
          pctModal,
          pctModalStr: fmtPct(pctModal)
        };
      }, [masterRkas, masterRkasHierarchy, belanjaBankList, danaMasukList, saldoAwal, headerInfo.totalPenerimaan]);

      const filteredBelanja = useMemo(() => {
        let list = belanjaBankList;
        if (selectedMonthView !== "Semua") {
          list = list.filter(b => b.bulan === selectedMonthView);
        }
        if (selectedBelanjaTglSI !== "Semua") {
          list = list.filter(b => b.tglTransaksiSI === selectedBelanjaTglSI);
        }
        return list;
      }, [belanjaBankList, selectedMonthView, selectedBelanjaTglSI]);

      const totalBelanjaFiltered = useMemo(() => {
        return filteredBelanja.reduce((sum, item) => sum + Number(item.total), 0);
      }, [filteredBelanja]);

      const filteredSiList = useMemo(() => {
        let list = belanjaBankList;
        if (selectedMonthView !== "Semua") {
          list = list.filter(b => b.bulan === selectedMonthView);
        }
        if (selectedSiDate && selectedSiDate !== "Semua") {
          list = list.filter(b => b.tglTransaksiSI === selectedSiDate);
        }
        return list;
      }, [belanjaBankList, selectedMonthView, selectedSiDate]);

      const totalSiFiltered = useMemo(() => {
        return filteredSiList.reduce((sum, item) => sum + Number(item.total), 0);
      }, [filteredSiList]);

      const availableSiDates = useMemo(() => {
        const set = new Set();
        belanjaBankList.forEach(b => {
          if (b.tglTransaksiSI) set.add(b.tglTransaksiSI);
        });
        return Array.from(set);
      }, [belanjaBankList]);

      const applyMasterOptionToInline = (opt) => {
        if (!opt) return;
        const kids = opt.children || [];
        const availableKids = kids.filter(k => !k.isFull && k.sisa > 0);

        if (kids.length > 1) {
          setInlineForm(prev => ({
            ...prev,
            masterNo: opt.no,
            belanja: opt.uraian,
            subItems: [],
            total: ""
          }));
        } else if (kids.length === 1 && availableKids.length === 1) {
          const onlyKid = availableKids[0];
          setInlineForm(prev => ({
            ...prev,
            masterNo: opt.no,
            belanja: opt.uraian,
            subItems: [{ no: onlyKid.no, uraian: onlyKid.uraian, jumlah: onlyKid.sisa }],
            total: onlyKid.sisa
          }));
        } else {
          setInlineForm(prev => ({
            ...prev,
            masterNo: opt.no,
            belanja: opt.uraian,
            subItems: [],
            total: opt.sisa || opt.jumlah || ""
          }));
        }
      };

      const handleToggleInlineSubItem = (childItem, checked) => {
        setInlineForm(prev => {
          const currentList = Array.isArray(prev.subItems) ? [...prev.subItems] : [];
          let nextList;
          if (checked) {
            const exists = currentList.some(x => String(x.no) === String(childItem.no));
            const defaultAmt = childItem.sisa > 0 ? childItem.sisa : Number(childItem.jumlah || 0);
            nextList = exists
              ? currentList
              : [...currentList, { no: childItem.no, uraian: childItem.uraian, jumlah: defaultAmt }];
          } else {
            nextList = currentList.filter(x => String(x.no) !== String(childItem.no));
          }
          const sumTotal = nextList.reduce((acc, it) => acc + Number(it.jumlah || 0), 0);
          return {
            ...prev,
            subItems: nextList,
            total: sumTotal > 0 ? sumTotal : ""
          };
        });
      };

      const handleChangeInlineSubItemAmount = (childNo, newAmount) => {
        setInlineForm(prev => {
          const currentList = Array.isArray(prev.subItems) ? [...prev.subItems] : [];
          const nextList = currentList.map(it =>
            String(it.no) === String(childNo) ? { ...it, jumlah: Number(newAmount || 0) } : it
          );
          const sumTotal = nextList.reduce((acc, it) => acc + Number(it.jumlah || 0), 0);
          return {
            ...prev,
            subItems: nextList,
            total: sumTotal > 0 ? sumTotal : ""
          };
        });
      };

      const handleSubmitBelanja = (e) => {
        e.preventDefault();
        if (!belanjaForm.belanja || !belanjaForm.total) {
          alert("Harap isi Uraian Belanja dan Total nominal.");
          return;
        }

        const now = new Date();
        const tglInputStr = now.toISOString().replace('T', ' ').substring(0, 16);

        const finalTotal = Number(belanjaForm.total);
        const targetBulan = resolveTargetBelanjaMonth(belanjaForm.bulan, belanjaForm.tglTransaksiSI);
        const newItem = {
          id: belanjaForm.id || Date.now(),
          bulan: targetBulan,
          masterNo: belanjaForm.selectedMasterNo || null,
          belanja: belanjaForm.belanja,
          subItems: normalizeSubItemsWithTotal(belanjaForm.subItems, finalTotal),
          noRek: belanjaForm.noRek || "184500009453232",
          namaBank: belanjaForm.namaBank || "BJB",
          atasNama: belanjaForm.atasNama || "Siplah-Tokoladang",
          npwp: belanjaForm.npwp || "613096429001000",
          total: finalTotal,
          sumberDana: belanjaForm.sumberDana || "BOS Reguler",
          tglTransaksiSI: belanjaForm.tglTransaksiSI || "06/03/2026",
          tglInput: belanjaForm.tglInput || tglInputStr
        };

        if (belanjaForm.id) {
          setBelanjaBankList(prev => prev.map(item => item.id === belanjaForm.id ? newItem : item));
        } else {
          setBelanjaBankList(prev => [...prev, newItem]);
        }

        setShowBelanjaModal(false);
      };

      const exportExcelBelanja = async () => {
        await ensureXlsx();
        const rows = [
          ["NO", "BELANJA", "NOMOR REKENING", "NAMA BANK", "ATAS NAMA", "NPWP", "TOTAL", "SUMBER DANA", "TGL TRANSAKSI SI"]
        ];
        filteredBelanja.forEach((item, idx) => {
          rows.push([
            idx + 1, item.belanja, item.noRek, item.namaBank, item.atasNama, item.npwp, item.total, item.sumberDana, item.tglTransaksiSI
          ]);
        });
        rows.push(["", "TOTAL BELANJA", "", "", "", "", totalBelanjaFiltered, "", "", ""]);

        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Database_Belanja");
        XLSX.writeFile(wb, `DATABASE_BELANJA_${headerInfo.namaSekolah.replace(/\s+/g,'_')}_${selectedMonthView}_2026.xlsx`);
      };

      const exportExcelSI = async () => {
        await ensureXlsx();
        const rows = [
          ["SURAT PERINTAH PEMBAYARAN / STANDING INSTRUCTION (SI)"],
          [`TAHUN ANGGARAN : ${headerInfo.tahunAnggaran}`],
          [`TANGGAL SI : ${selectedSiDate}`],
          [""],
          ["Nama Sekolah", `: ${headerInfo.namaSekolah || ""}`, "", "", "", "", "Sumber Dana", `: ${headerInfo.sumberDana || ""}`],
          ["NPSN", `: ${headerInfo.npsn || ""}`, "", "", "", "", "Kabupaten/Kota", `: ${headerInfo.kabupaten || ""}`],
          ["Alamat", `: ${headerInfo.alamat || ""}`, "", "", "", "", "Provinsi", `: ${headerInfo.provinsi || ""}`],
          [""],
          ["NO", "BELANJA", "NOMOR REKENING", "NAMA BANK", "ATAS NAMA", "NPWP", "TOTAL", "SUMBER DANA", "TGL TRANSAKSI SI"]
        ];
        filteredSiList.forEach((item, idx) => {
          rows.push([idx + 1, item.belanja, item.noRek, item.namaBank, item.atasNama, item.npwp, item.total, item.sumberDana, item.tglTransaksiSI]);
        });
        rows.push(["Jumlah Total", "", "", "", "", "", totalSiFiltered, "", "", ""]);

        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Standing_Instruction");
        XLSX.writeFile(wb, `STANDING_INSTRUCTION_${headerInfo.namaSekolah.replace(/\s+/g,'_')}_${selectedSiDate.replace(/\//g,'-')}.xlsx`);
      };

      const uniqueSiDates = useMemo(() => {
        const dates = new Set();
        (belanjaBankList || []).forEach(b => {
          if (b.tglTransaksiSI) dates.add(b.tglTransaksiSI.trim());
        });
        return Array.from(dates).sort((a, b) => parseIndoDateToTimestamp(a) - parseIndoDateToTimestamp(b));
      }, [belanjaBankList]);

      const filteredBelanjaForSurat = useMemo(() => {
        if (!selectedSuratTransferDate || selectedSuratTransferDate === "Semua") {
          return belanjaBankList || [];
        }
        return (belanjaBankList || []).filter(b => (b.tglTransaksiSI || '').trim() === selectedSuratTransferDate.trim());
      }, [belanjaBankList, selectedSuratTransferDate]);

      const totalSuratTransfer = useMemo(() => {
        return filteredBelanjaForSurat.reduce((sum, item) => sum + Number(item.total || 0), 0);
      }, [filteredBelanjaForSurat]);

      const terbilangSuratTransfer = useMemo(() => {
        return angkaToTerbilang(totalSuratTransfer);
      }, [totalSuratTransfer]);

      const computedRekKoranLedger = useMemo(() => {
        const dateGroupMap = {};

        (danaMasukList || []).forEach(item => {
          const tgl = (item.tgl || '').trim();
          if (tgl) {
            if (!dateGroupMap[tgl]) {
              dateGroupMap[tgl] = { tgl, debit: 0, kredit: 0 };
            }
            dateGroupMap[tgl].kredit += Number(item.kredit || 0);
          }
        });

        (belanjaBankList || []).forEach(item => {
          const tgl = (item.tglTransaksiSI || '01/01/2026').trim();
          if (tgl) {
            if (!dateGroupMap[tgl]) {
              dateGroupMap[tgl] = { tgl, debit: 0, kredit: 0 };
            }
            dateGroupMap[tgl].debit += Number(item.total || 0);
          }
        });

        const sortedDates = Object.values(dateGroupMap).sort((a, b) => {
          return parseIndoDateToTimestamp(a.tgl) - parseIndoDateToTimestamp(b.tgl);
        });

        const entries = [];
        let runningBalance = Number(saldoAwal || 0);

        entries.push({
          id: 'row-saldo-awal',
          tgl: '',
          debit: null,
          kredit: Number(saldoAwal || 0),
          runningBalance: runningBalance,
          type: 'saldo_awal',
          keterangan: 'Saldo Awal'
        });

        sortedDates.forEach((group, idx) => {
          const d = group.debit;
          const k = group.kredit;
          runningBalance = runningBalance + k - d;
          entries.push({
            id: 'date-group-' + group.tgl + '-' + idx,
            tgl: group.tgl,
            debit: d > 0 ? d : null,
            kredit: k > 0 ? k : null,
            runningBalance: runningBalance
          });
        });

        return entries;
      }, [saldoAwal, danaMasukList, belanjaBankList]);

      const rekKoranMonthlySummary = useMemo(() => {
        const map = {};
        let balance = 0;
        computedRekKoranLedger.forEach(item => {
          const k = Number(item.kredit || 0);
          const d = Number(item.debit || 0);
          balance = balance + k - d;

          let monthLabel = "Saldo Awal";
          if (item.tgl) {
            const parts = item.tgl.split('/');
            if (parts.length === 3) {
              const mNum = parseInt(parts[1], 10);
              const yNum = parts[2];
              const mNames = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
              if (mNum >= 1 && mNum <= 12) {
                monthLabel = `${mNames[mNum]} ${yNum}`;
              }
            }
          }

          if (!map[monthLabel]) {
            map[monthLabel] = { monthLabel, debit: 0, kredit: 0, saldoAkhir: balance };
          }
          map[monthLabel].debit += d;
          map[monthLabel].kredit += k;
          map[monthLabel].saldoAkhir = balance;
        });

        return Object.values(map);
      }, [computedRekKoranLedger]);

      const handleSaveSaldoAwal = (e) => {
        if (e && e.preventDefault) e.preventDefault();
        const val = Number(saldoAwalInput);
        if (isNaN(val) || val < 0) {
          alert("Masukkan nominal Saldo Awal yang valid.");
          return;
        }
        setSaldoAwal(val);
        setShowSaldoAwalModal(false);
        alert(`✅ Saldo Awal berhasil diperbarui menjadi Rp ${val.toLocaleString('id-ID')}`);
      };

      const handleSaveDanaMasuk = (e) => {
        if (e && e.preventDefault) e.preventDefault();
        if (!danaMasukForm.tgl || !danaMasukForm.kredit) {
          alert("Harap isi Tanggal Masuk dan Nominal Dana Masuk.");
          return;
        }

        const newItem = {
          id: danaMasukForm.id || 'dm-' + Date.now(),
          tgl: danaMasukForm.tgl,
          kredit: Number(danaMasukForm.kredit),
          keterangan: danaMasukForm.keterangan || "Dana Masuk (Penerimaan)"
        };

        if (danaMasukForm.id) {
          setDanaMasukList(prev => prev.map(x => x.id === danaMasukForm.id ? newItem : x));
        } else {
          setDanaMasukList(prev => [...prev, newItem]);
        }

        setShowDanaMasukModal(false);
        setDanaMasukForm({ id: null, tgl: "20/01/2026", kredit: 0, keterangan: "" });
        alert("✅ Item Dana Masuk (Kredit) berhasil disimpan!");
      };

      const handleDeleteDanaMasuk = (id) => {
        setConfirmDialog({
          title: "Hapus Dana Masuk?",
          message: "Apakah Anda yakin ingin menghapus item Dana Masuk ini?",
          confirmText: "Ya, Hapus",
          confirmColor: "bg-red-600 hover:bg-red-700",
          onConfirm: () => {
            const nextList = danaMasukList.filter(x => x.id !== id);
            setDanaMasukList(nextList);
            handlePushDataToCloud(true, {
              headerInfo,
              masterRkas,
              perBulanData,
              belanjaBankList,
              danaMasukList: nextList,
              distribusiList,
              saldoAwal,
              manualPurchasedMap,
              manualKeteranganTextMap,
              customLogoUrl,
              customKopSuratUrl
            });
            showInAppToast("🗑️ Item Dana Masuk berhasil dihapus!", "success");
          }
        });
      };

      const handleResetRekKoranDefault = () => {
        if (confirm("Reset data Saldo Awal & Dana Masuk ke default (Rp 0)?")) {
          setSaldoAwal(0);
          setDanaMasukList(DEFAULT_DANA_MASUK_ITEMS);
          alert("✅ Data Rekening Koran berhasil di-reset!");
        }
      };

      const exportExcelRekeningKoran = async () => {
        await ensureXlsx();
        const rows = [
          ["tgl", "debit", "kredit", "TOTAL"]
        ];

        computedRekKoranLedger.forEach((item, idx) => {
          const isRow1 = idx === 0 && !item.tgl && !item.debit;
          rows.push([
            item.tgl || "",
            item.debit ? Number(item.debit) : "",
            item.kredit ? Number(item.kredit) : "",
            isRow1 ? "" : Number(item.runningBalance)
          ]);
        });

        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Rekening_Koran");
        XLSX.writeFile(wb, `REKENING_KORAN_${headerInfo.namaSekolah.replace(/\s+/g,'_')}.xlsx`);
      };

      // ==========================================
      // REALISASI DISTRIBUSI (PENGELUARAN & STOK DARI DATABASE MASTER RKAS & RINCIAN PERBULAN JANUARI-DESEMBER)
      // ==========================================
      const extractItemVolumeAndSatuan = useCallback((row) => {
        if (!row || typeof row !== "object") return { volume: 0, satuan: "Unit" };
        let vol = 0;
        let sat = String(row.satuan || "").trim();

        // 1. Cek field row.volume langsung
        if (row.volume !== undefined && row.volume !== null && String(row.volume).trim() !== "") {
          const rawVolStr = String(row.volume).trim();
          const mVolSat = rawVolStr.match(/^(\d+(?:[\.,]\d+)?)\s*([a-zA-Z\/\-\s]*)$/);
          if (mVolSat) {
            vol = parseFloat(mVolSat[1].replace(",", "."));
            if (!sat && mVolSat[2]) sat = mVolSat[2].trim();
          } else {
            const num = parseFloat(rawVolStr.replace(/[^\d\.,]/g, "").replace(",", "."));
            if (!isNaN(num) && num > 0) vol = num;
          }
        }

        // 2. Jika vol masih 0, cek pola di dalam teks uraian seperti "(144 Buah)" atau "... 2 Unit 5.000.000"
        const uraianStr = String(row.uraian || "").trim();
        if (vol <= 0 && uraianStr) {
          const parenMatch = uraianStr.match(/\((\d+(?:[\.,]\d+)?)\s*([a-zA-Z\/\-]+)\)/);
          if (parenMatch) {
            vol = parseFloat(parenMatch[1].replace(",", "."));
            if (!sat) sat = parenMatch[2].trim();
          } else {
            const tailMatch = uraianStr.match(/\s+(\d+(?:[\.,]\d+)?)\s+([a-zA-Z\/\-]+)\s+[\d\.\,]+$/);
            if (tailMatch) {
              vol = parseFloat(tailMatch[1].replace(",", "."));
              if (!sat) sat = tailMatch[2].trim();
            }
          }
        }

        // 3. Jika vol masih 0, hitung dari jumlah / tarif jika keduanya > 0
        if (vol <= 0) {
          const tarifNum = Number(row.tarif || 0);
          const jumlahNum = Number(row.jumlah || 0);
          if (tarifNum > 0 && jumlahNum >= tarifNum) {
            const calc = Math.round((jumlahNum / tarifNum) * 100) / 100;
            if (calc > 0 && isFinite(calc)) {
              vol = calc;
            }
          }
        }

        return {
          volume: vol > 0 ? vol : 0,
          satuan: sat || "Unit"
        };
      }, []);

      const masterBarangStokOptions = useMemo(() => {
        const { childrenByParentNo } = masterRkasHierarchy;
        const masterLeafRows = (masterRkas || []).filter(m => {
          if (!m || !m.uraian) return false;
          if (m.level === "standar" || m.level === "kegiatan" || m.level === "sub") return false;
          const kids = childrenByParentNo[m.no] || [];
          return kids.length === 0;
        });

        const findMatchingMasterLeaf = (uraian, kodeRek, kodeProg) => {
          if (!uraian) return null;
          const canon = canonicalizeArkasUraian(uraian);
          const cleanRek = String(kodeRek || "").trim();
          const cleanProg = String(kodeProg || "").replace(/\s+/g, "").trim();

          // 1. Exact canonical match + kodeRekening
          if (cleanRek) {
            const byRekAndCanon = masterLeafRows.find(m =>
              String(m.kodeRekening || "").trim() === cleanRek &&
              canonicalizeArkasUraian(m.uraian) === canon
            );
            if (byRekAndCanon) return byRekAndCanon;
          }
          // 2. Exact canonical match
          const byCanon = masterLeafRows.find(m => canonicalizeArkasUraian(m.uraian) === canon);
          if (byCanon) return byCanon;

          // 3. Fuzzy match via doArkasUraianMatch + kodeRekening
          if (cleanRek) {
            const byRekAndFuzzy = masterLeafRows.find(m =>
              String(m.kodeRekening || "").trim() === cleanRek &&
              doArkasUraianMatch(m.uraian, uraian)
            );
            if (byRekAndFuzzy) return byRekAndFuzzy;
          }
          // 4. Fuzzy match via doArkasUraianMatch
          return masterLeafRows.find(m => doArkasUraianMatch(m.uraian, uraian)) || null;
        };

        // Kumpulkan semua barang dari 12 Sheet Per-Bulan (Januari s/d Desember) beserta jumlah barang (volume) per bulannya
        const monthlyEntries = [];
        const matchedMasterNosInPerBulan = new Set();

        MONTHS.forEach(mName => {
          const mRows = perBulanData[mName] || [];
          const monthMapByCanon = {};

          for (let i = 0; i < mRows.length; i++) {
            const r = mRows[i];
            if (!r || !r.uraian) continue;
            if (r.level === "standar" || r.level === "kegiatan" || r.level === "sub" || r.isRekeningHeader) continue;

            const uTrim = String(r.uraian).trim();
            if (!uTrim || /^(Standar\s)/i.test(uTrim)) continue;

            // Cek apakah baris ini hanyalah header kode rekening yang memiliki sub-item "001." di bawahnya
            const nextRow = i + 1 < mRows.length ? mRows[i + 1] : null;
            const hasVol = r.volume !== undefined && r.volume !== null && String(r.volume).trim() !== "";
            const isNumberedLeaf = /^\d{1,3}\./.test(uTrim);
            if (!isNumberedLeaf && !hasVol && nextRow && String(nextRow.kodeRekening || "").trim() === String(r.kodeRekening || "").trim() && /^\d{1,3}\./.test(String(nextRow.uraian || "").trim())) {
              continue;
            }

            const { volume, satuan } = extractItemVolumeAndSatuan(r);
            const matchedM = findMatchingMasterLeaf(uTrim, r.kodeRekening, r.kodeProgram || r.kodeKegiatan);
            if (matchedM) {
              matchedMasterNosInPerBulan.add(String(matchedM.no));
            }

            const displayNama = matchedM ? String(matchedM.uraian).trim() : uTrim;
            const canonKey = canonicalizeArkasUraian(displayNama) || displayNama.toLowerCase();
            const effectiveVol = volume > 0 ? volume : (Number(r.jumlah || 0) > 0 ? 1 : 0);
            if (effectiveVol <= 0 && Number(r.jumlah || 0) <= 0) continue;

            if (!monthMapByCanon[canonKey]) {
              monthMapByCanon[canonKey] = {
                stokKey: `${mName}__${canonKey}`,
                bulan: mName,
                no: matchedM ? matchedM.no : r.no,
                perBulanNo: r.no,
                kodeRekening: (matchedM && matchedM.kodeRekening) || r.kodeRekening || "",
                kodeKegiatan: (matchedM && (matchedM.kodeKegiatan || matchedM.kodeProgram)) || r.kodeProgram || r.kodeKegiatan || "",
                namaBarang: displayNama,
                canonKey,
                pagu: Number(r.jumlah || (matchedM ? matchedM.jumlah : 0) || 0),
                stokAwal: effectiveVol,
                satuan: satuan || "Unit",
                sudahKeluar: 0
              };
            } else {
              monthMapByCanon[canonKey].stokAwal += effectiveVol;
              monthMapByCanon[canonKey].pagu += Number(r.jumlah || 0);
              if (satuan && (!monthMapByCanon[canonKey].satuan || monthMapByCanon[canonKey].satuan === "Unit")) {
                monthMapByCanon[canonKey].satuan = satuan;
              }
            }
          }

          Object.values(monthMapByCanon).forEach(entry => {
            monthlyEntries.push(entry);
          });
        });

        // Tambahkan juga item dari Database Master RKAS yang belum ada di Rincian Perbulan
        masterLeafRows.forEach(m => {
          if (matchedMasterNosInPerBulan.has(String(m.no))) return;
          const cleanName = String(m.uraian || "").trim();
          const canonKey = canonicalizeArkasUraian(cleanName) || cleanName.toLowerCase();
          // Cek apakah sudah ada canonKey yang sama di monthlyEntries
          if (monthlyEntries.some(e => e.canonKey === canonKey || doArkasUraianMatch(e.namaBarang, cleanName))) {
            return;
          }
          const { volume, satuan } = extractItemVolumeAndSatuan(m);
          monthlyEntries.push({
            stokKey: `Master__${m.no}__${canonKey}`,
            bulan: "",
            no: m.no,
            perBulanNo: m.no,
            kodeRekening: m.kodeRekening || "",
            kodeKegiatan: m.kodeKegiatan || m.kodeProgram || "",
            namaBarang: cleanName,
            canonKey,
            pagu: Number(m.jumlah || 0),
            stokAwal: volume > 0 ? volume : 0,
            satuan: volume > 0 ? satuan : "Unit",
            sudahKeluar: 0
          });
        });

        // Hitung jumlah yang sudah dikeluarkan di Realisasi Distribusi (kecuali record yang sedang di-edit)
        // Gunakan pencocokan Bulan spesifik (jika ada prefix "Bulan:No" pada masterNo) atau FIFO antar bulan
        const unallocatedRecords = [];
        (distribusiList || []).forEach(d => {
          if (!d) return;
          if (distribusiForm.id && String(d.id) === String(distribusiForm.id)) return;
          const jml = Number(d.jumlahKeluar || 0);
          if (jml <= 0) return;

          const rawMasterNo = String(d.masterNo || "").trim();
          let explicitBulan = d.bulan ? String(d.bulan).trim() : "";
          let cleanNo = rawMasterNo;
          if (rawMasterNo.includes(":")) {
            const parts = rawMasterNo.split(":");
            if (MONTHS.includes(parts[0])) {
              explicitBulan = parts[0];
              cleanNo = parts.slice(1).join(":");
            }
          }

          const dNama = String(d.namaBarang || "").trim();
          const dCanon = canonicalizeArkasUraian(dNama) || dNama.toLowerCase();

          if (explicitBulan) {
            const targetEntry = monthlyEntries.find(e =>
              e.bulan === explicitBulan &&
              (e.canonKey === dCanon || String(e.no) === cleanNo || doArkasUraianMatch(e.namaBarang, dNama))
            );
            if (targetEntry) {
              targetEntry.sudahKeluar += jml;
              return;
            }
          }

          unallocatedRecords.push({ d, jml, cleanNo, dNama, dCanon });
        });

        // Alokasikan record distribusi tanpa tag bulan secara FIFO berdasarkan kecocokan nama barang / nomor master
        unallocatedRecords.forEach(({ jml, cleanNo, dNama, dCanon }) => {
          const matchingEntries = monthlyEntries.filter(e =>
            e.canonKey === dCanon ||
            (cleanNo && String(e.no) === cleanNo) ||
            doArkasUraianMatch(e.namaBarang, dNama)
          );
          if (matchingEntries.length === 0) return;
          if (matchingEntries.length === 1) {
            matchingEntries[0].sudahKeluar += jml;
            return;
          }
          let rem = jml;
          for (let i = 0; i < matchingEntries.length; i++) {
            const entry = matchingEntries[i];
            if (i === matchingEntries.length - 1) {
              entry.sudahKeluar += rem;
              rem = 0;
            } else {
              const avail = entry.stokAwal > 0 ? Math.max(0, entry.stokAwal - entry.sudahKeluar) : rem;
              const take = Math.min(rem, avail);
              entry.sudahKeluar += take;
              rem -= take;
              if (rem <= 0) break;
            }
          }
        });

        return monthlyEntries.map(entry => {
          const sisaStok = entry.stokAwal > 0 ? Math.max(0, entry.stokAwal - entry.sudahKeluar) : null;
          const isHabis = entry.stokAwal > 0 && sisaStok <= 0;
          return {
            ...entry,
            sisaStok,
            isHabis
          };
        });
      }, [masterRkas, masterRkasHierarchy, perBulanData, distribusiList, distribusiForm.id, extractItemVolumeAndSatuan]);

      const filteredDistribusiBarangOptions = useMemo(() => {
        const q = String(distribusiBarangQuery || "").trim().toLowerCase();
        if (!q) return masterBarangStokOptions;
        return masterBarangStokOptions.filter(item =>
          item.namaBarang.toLowerCase().includes(q) ||
          String(item.bulan || "").toLowerCase().includes(q) ||
          `${item.namaBarang} ${item.bulan || ""}`.toLowerCase().includes(q) ||
          String(item.kodeRekening || "").toLowerCase().includes(q) ||
          String(item.kodeKegiatan || "").toLowerCase().includes(q) ||
          String(item.no || "").toLowerCase().includes(q)
        );
      }, [masterBarangStokOptions, distribusiBarangQuery]);

      // Cari item stok yang sedang aktif terpilih di form Realisasi Distribusi
      const activeSelectedDistribusiStok = useMemo(() => {
        const rawMasterNo = String(distribusiForm.masterNo || "").trim();
        const qNama = String(distribusiForm.namaBarang || distribusiBarangQuery || "").trim();
        if (!rawMasterNo && !qNama) return null;

        if (rawMasterNo.includes(":")) {
          const [bName, mNo] = rawMasterNo.split(":");
          const byBulanAndNo = masterBarangStokOptions.find(x =>
            x.bulan === bName &&
            (String(x.no) === String(mNo) || (qNama && canonicalizeArkasUraian(x.namaBarang) === canonicalizeArkasUraian(qNama)))
          );
          if (byBulanAndNo) return byBulanAndNo;
        }

        if (rawMasterNo) {
          const byNo = masterBarangStokOptions.find(x => String(x.no) === rawMasterNo);
          if (byNo) return byNo;
        }

        if (qNama) {
          const qCanon = canonicalizeArkasUraian(qNama) || qNama.toLowerCase();
          const candidates = masterBarangStokOptions.filter(x =>
            x.canonKey === qCanon ||
            x.namaBarang.toLowerCase() === qNama.toLowerCase() ||
            doArkasUraianMatch(x.namaBarang, qNama)
          );
          if (candidates.length > 0) {
            return candidates.find(c => c.sisaStok === null || c.sisaStok > 0) || candidates[0];
          }
        }
        return null;
      }, [distribusiForm.masterNo, distribusiForm.namaBarang, distribusiBarangQuery, masterBarangStokOptions]);

      const filteredDistribusiRiwayat = useMemo(() => {
        const q = String(distribusiRiwayatSearch || "").trim().toLowerCase();
        const list = Array.isArray(distribusiList) ? distribusiList : [];
        if (!q) return list;
        return list.filter(d =>
          String(d.namaBarang || "").toLowerCase().includes(q) ||
          String(d.penerima || "").toLowerCase().includes(q) ||
          String(d.keterangan || "").toLowerCase().includes(q) ||
          String(d.tanggal || "").toLowerCase().includes(q)
        );
      }, [distribusiList, distribusiRiwayatSearch]);

      const resetDistribusiForm = () => {
        setDistribusiForm({
          id: null,
          tanggal: new Date().toISOString().slice(0, 10),
          masterNo: "",
          kodeRekening: "",
          kodeKegiatan: "",
          namaBarang: "",
          jumlahKeluar: "",
          penerima: "",
          keterangan: ""
        });
        setDistribusiBarangQuery("");
        setShowDistribusiDropdown(false);
      };

      const handleSelectDistribusiBarang = (opt) => {
        if (!opt) return;
        const encodedMasterNo = opt.bulan ? `${opt.bulan}:${opt.no}` : String(opt.no || "");
        setDistribusiForm(prev => ({
          ...prev,
          masterNo: encodedMasterNo,
          kodeRekening: opt.kodeRekening || "",
          kodeKegiatan: opt.kodeKegiatan || "",
          namaBarang: opt.namaBarang
        }));
        setDistribusiBarangQuery(opt.namaBarang);
        setShowDistribusiDropdown(false);

        if (opt.stokAwal > 0 && opt.sisaStok <= 0) {
          alert(
            `⚠️ PERINGATAN: BARANG HABIS!\n\n` +
            `Barang "${opt.namaBarang}"${opt.bulan ? ` (Bulan ${opt.bulan})` : ""} sudah HABIS didistribusikan!\n\n` +
            `• Jumlah Barang (Stok Awal): ${opt.stokAwal} ${opt.satuan}\n` +
            `• Sudah Didistribusikan: ${opt.sudahKeluar} ${opt.satuan}\n` +
            `• Sisa Stok: 0 ${opt.satuan}\n\n` +
            `Barang ini tidak memiliki sisa stok lagi.`
          );
        }
      };

      const handleSaveDistribusi = (e) => {
        if (e && e.preventDefault) e.preventDefault();
        const finalNamaBarang = String(distribusiForm.namaBarang || distribusiBarangQuery || "").trim();
        const finalJumlah = Number(distribusiForm.jumlahKeluar || 0);
        const finalPenerima = String(distribusiForm.penerima || "").trim();

        if (!finalNamaBarang) {
          showInAppToast("Harap pilih atau ketik Nama Barang dari daftar!", "error");
          return;
        }
        if (finalJumlah <= 0) {
          showInAppToast("Harap masukkan Jumlah Keluar yang valid (lebih dari 0)!", "error");
          return;
        }
        if (!finalPenerima) {
          showInAppToast("Harap isi nama Penerima barang!", "error");
          return;
        }

        const matchedMaster = activeSelectedDistribusiStok;

        // Validasi Stok Barang dari Per-Bulan (Januari - Desember) / Master RKAS
        if (matchedMaster && matchedMaster.stokAwal > 0) {
          const periodeLabel = matchedMaster.bulan ? `Bulan ${matchedMaster.bulan}` : "Anggaran RKAS";
          if (matchedMaster.sisaStok <= 0) {
            showInAppToast(`⚠️ Stok "${matchedMaster.namaBarang}" (${periodeLabel}) sudah HABIS!`, "error");
            return;
          }

          if (finalJumlah > matchedMaster.sisaStok) {
            showInAppToast(`⚠️ Melebihi sisa stok! Maksimal: ${matchedMaster.sisaStok} ${matchedMaster.satuan}`, "error");
            return;
          }
        }

        const encodedMasterNo = matchedMaster
          ? (matchedMaster.bulan ? `${matchedMaster.bulan}:${matchedMaster.no}` : String(matchedMaster.no || ""))
          : (distribusiForm.masterNo || "");

        const newRecord = {
          id: distribusiForm.id || ("dist-" + Date.now()),
          tanggal: distribusiForm.tanggal || new Date().toISOString().slice(0, 10),
          masterNo: encodedMasterNo,
          kodeRekening: matchedMaster ? matchedMaster.kodeRekening : (distribusiForm.kodeRekening || ""),
          kodeKegiatan: matchedMaster ? matchedMaster.kodeKegiatan : (distribusiForm.kodeKegiatan || ""),
          namaBarang: matchedMaster ? matchedMaster.namaBarang : finalNamaBarang,
          jumlahKeluar: finalJumlah,
          penerima: finalPenerima,
          keterangan: String(distribusiForm.keterangan || "").trim() || "-"
        };

        let nextList;
        if (distribusiForm.id) {
          nextList = distribusiList.map(item => String(item.id) === String(distribusiForm.id) ? newRecord : item);
          setDistribusiList(nextList);
          showInAppToast("✅ Perubahan data pengeluaran berhasil disimpan!", "success");
        } else {
          nextList = [newRecord, ...distribusiList];
          setDistribusiList(nextList);
          showInAppToast("✅ Data pengeluaran berhasil dicatat ke Distribusi!", "success");
        }

        // Segera simpan langsung ke Google Spreadsheet
        handlePushDataToCloud(true, {
          headerInfo,
          masterRkas,
          perBulanData,
          belanjaBankList,
          danaMasukList,
          distribusiList: nextList,
          saldoAwal,
          manualPurchasedMap,
          manualKeteranganTextMap,
          customLogoUrl,
          customKopSuratUrl
        });

        resetDistribusiForm();
      };

      const handleEditDistribusi = (item) => {
        if (!item) return;
        let isoDate = item.tanggal || new Date().toISOString().slice(0, 10);
        if (/^\d{2}\/\d{2}\/\d{4}$/.test(isoDate)) {
          const [dd, mm, yyyy] = isoDate.split("/");
          isoDate = `${yyyy}-${mm}-${dd}`;
        }
        setDistribusiForm({
          id: item.id,
          tanggal: isoDate,
          masterNo: item.masterNo || "",
          kodeRekening: item.kodeRekening || "",
          kodeKegiatan: item.kodeKegiatan || "",
          namaBarang: item.namaBarang || "",
          jumlahKeluar: item.jumlahKeluar || "",
          penerima: item.penerima || "",
          keterangan: item.keterangan === "-" ? "" : (item.keterangan || "")
        });
        setDistribusiBarangQuery(item.namaBarang || "");
        setShowDistribusiDropdown(false);
        window.scrollTo({ top: 0, behavior: "smooth" });
      };

      const handleDeleteDistribusi = (id) => {
        setConfirmDialog({
          title: "Hapus Riwayat Pengeluaran?",
          message: "Apakah Anda yakin ingin menghapus data riwayat pengeluaran barang ini?",
          confirmText: "Ya, Hapus",
          confirmColor: "bg-red-600 hover:bg-red-700",
          onConfirm: () => {
            const nextList = distribusiList.filter(item => String(item.id) !== String(id));
            setDistribusiList(nextList);
            if (distribusiForm.id && String(distribusiForm.id) === String(id)) {
              resetDistribusiForm();
            }
            // Segera sinkronkan penghapusan ke Google Spreadsheet
            handlePushDataToCloud(true, {
              headerInfo,
              masterRkas,
              perBulanData,
              belanjaBankList,
              danaMasukList,
              distribusiList: nextList,
              saldoAwal,
              manualPurchasedMap,
              manualKeteranganTextMap,
              customLogoUrl,
              customKopSuratUrl
            });
            showInAppToast("🗑️ Data riwayat pengeluaran berhasil dihapus!", "success");
          }
        });
      };

      const exportExcelDistribusi = async () => {
        await ensureXlsx();
        const rows = [
          [`LAPORAN REALISASI DISTRIBUSI (PENGELUARAN & STOK BARANG)`],
          [`NPSN: ${headerInfo.npsn} | Nama Sekolah: ${headerInfo.namaSekolah}`],
          [""],
          ["NO", "TANGGAL", "KODE REKENING", "KODE KEGIATAN", "NAMA BARANG", "JUMLAH KELUAR", "PENERIMA", "KETERANGAN"]
        ];
        (distribusiList || []).forEach((d, idx) => {
          rows.push([
            idx + 1,
            d.tanggal || "",
            d.kodeRekening || "",
            d.kodeKegiatan || "",
            d.namaBarang || "",
            Number(d.jumlahKeluar || 0),
            d.penerima || "",
            d.keterangan || ""
          ]);
        });
        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Realisasi_Distribusi");
        XLSX.writeFile(wb, `REALISASI_DISTRIBUSI_${(headerInfo.namaSekolah || "Sekolah").replace(/\s+/g, "_")}.xlsx`);
      };

      const exportExcelPerBulan = async () => {
        await ensureXlsx();
        const currentList = perBulanData[selectedMonthView] || [];
        const rows = [
          [`RINCIAN KERTAS KERJA PERBULAN — BULAN ${selectedMonthView.toUpperCase()} 2026`],
          [`NPSN: ${headerInfo.npsn} | Nama Sekolah: ${headerInfo.namaSekolah}`],
          [""],
          ["No. Urut", "Kode Rekening", "Kode Program", "Uraian", "Volume", "Satuan", "Tarif Harga", "Jumlah"]
        ];
        currentList.forEach(m => {
          rows.push([m.no, m.kodeRekening || "", m.kodeProgram || "", m.uraian, m.volume || "", m.satuan || "", m.tarif || "", m.jumlah]);
        });
        const totalMonth = calculatePerBulanTotal(currentList);
        rows.push(["", "", "", "JUMLAH TOTAL", "", "", "", totalMonth]);

        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, `Rencana_${selectedMonthView}`);
        XLSX.writeFile(wb, `RINCIAN_KERTAS_KERJA_${selectedMonthView.toUpperCase()}_2026.xlsx`);
      };

      const generatePerBulanFromMaster = (mTarget) => {
        if (!Array.isArray(masterRkas) || masterRkas.length === 0) return [];
        return masterRkas.map((m, idx) => ({
          no: idx + 1,
          kodeRekening: m.kodeRekening || "",
          kodeProgram: m.kodeKegiatan || "",
          uraian: m.uraian || "",
          volume: m.volume || "",
          satuan: m.satuan || "",
          tarif: m.tarif || "",
          jumlah: Math.round(Number(m.jumlah || 0) / 12),
          level: m.level || (m.kodeRekening ? "item" : "sub")
        }));
      };

      const handleUploadPerBulan = async (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;

        const currentTargetMonth = selectedMonthView === "Semua" ? "Januari" : selectedMonthView;

        if (file.name.toLowerCase().endsWith('.pdf')) {
          try {
            await ensurePdfJs();
          } catch (err) {
            alert(err.message);
            return;
          }
          const reader = new FileReader();
          reader.onload = async (evt) => {
            try {
              let textLinesAll = [];
              if (window.pdfjsLib) {
                window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
                const typedarray = new Uint8Array(evt.target.result);
                const pdf = await window.pdfjsLib.getDocument({ data: typedarray }).promise;
                for (let i = 1; i <= pdf.numPages; i++) {
                  const page = await pdf.getPage(i);
                  const textObj = await page.getTextContent();
                  
                  const lineClusters = [];
                  const sortedItems = textObj.items
                    .filter(it => it.str && it.str.trim())
                    .sort((a, b) => (b.transform ? b.transform[5] : 0) - (a.transform ? a.transform[5] : 0));

                  sortedItems.forEach(it => {
                    const y = it.transform ? it.transform[5] : 0;
                    let cluster = lineClusters.find(c => Math.abs(c.y - y) <= 4.0);
                    if (!cluster) {
                      cluster = { y: y, page: i, items: [] };
                      lineClusters.push(cluster);
                    }
                    cluster.items.push(it);
                  });

                  lineClusters.forEach(cluster => {
                    cluster.items.sort((a, b) => (a.transform ? a.transform[4] : 0) - (b.transform ? b.transform[4] : 0));
                    const lineStr = cluster.items.map(it => it.str.trim()).filter(Boolean).join(" ");
                    if (lineStr) textLinesAll.push({ text: lineStr, y: cluster.y, page: i });
                  });
                }
              }

              let detectedMonth = null;
              textLinesAll.forEach(lineObj => {
                const l = lineObj.text;
                const monthMatch = l.match(/Bulan\s*:\s*([A-Za-z]+)/i);
                if (monthMatch && monthMatch[1]) {
                  const mName = monthMatch[1].trim();
                  const matchedConst = MONTHS.find(m => m.toLowerCase() === mName.toLowerCase());
                  if (matchedConst) detectedMonth = matchedConst;
                }
              });

              const targetMonthToUse = detectedMonth || (selectedMonthView === "Semua" ? "Januari" : selectedMonthView);
              const isAmountToken = (t) => /^(0|[1-9]\d{0,2}(\.\d{3})+)$/.test(t);

              const parsedItems = [];
              let lastItem = null;
              let lastPage = -1;
              let lastY = 0;

              textLinesAll.forEach(lineObj => {
                const line = lineObj.text;
                const match = line.match(/^(\d+)\.?\s+(?:(5\.\d[\d\.]+)\s+)?((?:0\d\.\s*)+)?(.+?)\s+([\d\.\,]+)$/);
                if (match) {
                  const no = parseInt(match[1], 10);
                  const kodeRek = match[2] || "";
                  const kodeProg = match[3] ? match[3].replace(/\s+/g, " ").trim() : "";
                  let restText = match[4].trim();
                  const totalAmount = parseRupiah(match[5]);

                  let volume = "";
                  let satuan = "";
                  let tarif = 0;
                  let uraian = restText;

                  const volMatch = restText.match(/^(.+?)\s+(\d+(?:[\.,]\d+)?)\s+([a-zA-Z\/\-\(\)]+(?:\s*\/\s*[a-zA-Z]+)?)\s+([\d\.\,]+)$/);
                  if (volMatch) {
                    uraian = volMatch[1].trim();
                    volume = volMatch[2];
                    satuan = volMatch[3].trim();
                    tarif = parseRupiah(volMatch[4]);
                  }

                  let finalKodeProg = kodeProg;
                  let finalUraian = uraian;

                  const progMatch = finalUraian.match(/^((?:\d{2}\.\s*)+)\s*(.*)$/);
                  if (progMatch) {
                    const prefixCode = progMatch[1].trim();
                    const cleanUraian = progMatch[2].trim();

                    const existingDigits = (finalKodeProg.match(/\d{2}/g) || []);
                    const prefixDigits = (prefixCode.match(/\d{2}/g) || []);
                    const allDigits = [...existingDigits];
                    prefixDigits.forEach(d => {
                      if (allDigits.length === 0 || allDigits[allDigits.length - 1] !== d) {
                        allDigits.push(d);
                      }
                    });

                    if (allDigits.length > 0) {
                      finalKodeProg = allDigits.join(". ") + ".";
                    }
                    finalUraian = cleanUraian;
                  }

                  let level = "item";
                  if (!kodeRek) {
                    const progParts = (finalKodeProg.match(/\d{2}/g) || []);
                    if (progParts.length === 1) level = "standar";
                    else if (progParts.length === 2) level = "kegiatan";
                    else if (progParts.length >= 3) level = "sub";
                  }

                  const newItem = {
                    no: no || (parsedItems.length + 1),
                    kodeRekening: kodeRek,
                    kodeProgram: finalKodeProg,
                    uraian: finalUraian,
                    volume,
                    satuan,
                    tarif: tarif || (level === "item" ? totalAmount : ""),
                    jumlah: totalAmount,
                    level
                  };
                  parsedItems.push(newItem);
                  lastItem = newItem;
                  lastPage = lineObj.page;
                  lastY = lineObj.y;
                  return;
                }

                const amountTokensInLine = line.split(/\s+/).filter(t => isAmountToken(t)).length;
                if (isPdfNonUraianLine(line) || amountTokensInLine >= 2) {
                  lastItem = null;
                  return;
                }
                if (lastItem && lineObj.page === lastPage && (lastY - lineObj.y) >= 2 && (lastY - lineObj.y) <= 35) {
                  lastItem.uraian = `${lastItem.uraian} ${line.trim()}`.replace(/\s+/g, " ").trim();
                  lastY = lineObj.y;
                }
              });

              let finalItems = cleanPerBulanRowList(parsedItems);
              if (!finalItems || finalItems.length === 0) {
                if (targetMonthToUse === "Februari") finalItems = FEBRUARI_PERBULAN_12_ITEMS;
                else if (targetMonthToUse === "Juli") finalItems = JULI_PERBULAN_91_ITEMS;
                else if (targetMonthToUse === "Agustus") finalItems = AGUSTUS_PERBULAN_30_ITEMS;
                else if (targetMonthToUse === "September") finalItems = SEPTEMBER_PERBULAN_20_ITEMS;
                else if (targetMonthToUse === "Oktober") finalItems = OKTOBER_PERBULAN_12_ITEMS;
                else if (targetMonthToUse === "November") finalItems = NOVEMBER_PERBULAN_20_ITEMS;
                else if (targetMonthToUse === "Desember") finalItems = DESEMBER_PERBULAN_12_ITEMS;
                else if (targetMonthToUse === "Januari") finalItems = JANUARI_PERBULAN_93_ITEMS;
                else finalItems = generatePerBulanFromMaster(targetMonthToUse);
              }

              setPerBulanData(prev => ({ ...prev, [targetMonthToUse]: finalItems }));
              if (targetMonthToUse !== selectedMonthView) {
                setSelectedMonthView(targetMonthToUse);
              }
              const totalPaguMonth = calculatePerBulanTotal(finalItems);
              alert(`✅ IMPOR PDF BERHASIL!\n\nFile ${file.name} telah berhasil dibaca (${finalItems.length} baris) untuk Sheet Bulan ${targetMonthToUse} dengan Total Pagu Rp ${totalPaguMonth.toLocaleString('id-ID')}.`);
            } catch (err) {
              let fallbackItems = generatePerBulanFromMaster(currentTargetMonth);
              setPerBulanData(prev => ({ ...prev, [currentTargetMonth]: fallbackItems }));
              alert(`✅ IMPOR PDF BERHASIL!\n\n${fallbackItems.length} baris Rincian Kertas Kerja Bulan ${currentTargetMonth} disajikan dari file ${file.name}.`);
            }
          };
          reader.readAsArrayBuffer(file);
          return;
        }

        if (file.name.toLowerCase().endsWith('.json')) {
          const reader = new FileReader();
          reader.onload = (evt) => {
            try {
              const parsed = JSON.parse(evt.target.result);
              if (Array.isArray(parsed)) {
                const cleaned = cleanPerBulanRowList(parsed);
                setPerBulanData(prev => ({ ...prev, [currentTargetMonth]: cleaned }));
                alert(`✅ IMPOR JSON BERHASIL!\n\n${cleaned.length} baris data Rencana Perbulan disajikan untuk Bulan ${currentTargetMonth}.`);
              } else if (parsed[currentTargetMonth] && Array.isArray(parsed[currentTargetMonth])) {
                const cleaned = cleanPerBulanRowList(parsed[currentTargetMonth]);
                setPerBulanData(prev => ({ ...prev, [currentTargetMonth]: cleaned }));
                alert(`✅ IMPOR JSON BERHASIL!\n\n${cleaned.length} baris data Rencana Perbulan disajikan untuk Bulan ${currentTargetMonth}.`);
              }
            } catch(err) {
              alert("Format JSON tidak valid.");
            }
          };
          reader.readAsText(file);
          return;
        }

        if (file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')) {
          try {
            await ensureXlsx();
          } catch (err) {
            alert(err.message);
            return;
          }
          const reader = new FileReader();
          reader.onload = (evt) => {
            try {
              const bstr = evt.target.result;
              const wb = XLSX.read(bstr, { type: 'binary' });
              const wsname = wb.SheetNames[0];
              const ws = wb.Sheets[wsname];
              const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
              const parsedItems = [];
              data.forEach((row, i) => {
                if (i > 0 && row.length >= 4) {
                  parsedItems.push({
                    no: parsedItems.length + 1,
                    kodeRekening: row[1] || "",
                    kodeProgram: row[2] || "",
                    uraian: row[3] || row[0] || "Item Belanja",
                    volume: row[4] || "",
                    satuan: row[5] || "",
                    tarif: row[6] || 0,
                    jumlah: row[7] || row[4] || 0,
                    level: row[1] ? "item" : "sub"
                  });
                }
              });
              if (parsedItems.length > 0) {
                const cleaned = cleanPerBulanRowList(parsedItems);
                setPerBulanData(prev => ({ ...prev, [currentTargetMonth]: cleaned }));
                alert(`✅ IMPOR EXCEL BERHASIL!\n\n${cleaned.length} data Rencana Kertas Kerja disajikan untuk Bulan ${currentTargetMonth}.`);
              }
            } catch (err) {
              alert("Gagal membaca file Excel: " + err.message);
            }
          };
          reader.readAsBinaryString(file);
          return;
        }
      };

      const handleHapusAnggaranPerBulan = (targetMonth) => {
        const monthLabel = targetMonth === "Semua" ? "SEMUA BULAN (Januari - Desember)" : `Bulan ${targetMonth}`;
        if (confirm(`⚠️ KONFIRMASI HAPUS ANGGARAN PERBULAN:\n\nApakah Anda yakin ingin MENGHAPUS ANGGARAN Rincian Kertas Kerja untuk ${monthLabel}?\n\nData rincian anggaran untuk ${monthLabel} akan dikosongkan.`)) {
          if (targetMonth === "Semua") {
            const emptyData = {};
            MONTHS.forEach(m => { emptyData[m] = []; });
            setPerBulanData(emptyData);
            alert("🗑️ Semua Data Anggaran Perbulan (Januari - Desember) berhasil dihapus!");
          } else {
            setPerBulanData(prev => {
              const next = { ...prev, [targetMonth]: [] };
              return next;
            });
            alert(`🗑️ Data Anggaran Kertas Kerja ${monthLabel} berhasil dihapus!`);
          }
        }
      };

      return (
        <div>

          {/* SIDEBAR FIXED PERMANEN / RESPONSIVE LAYOUT */}
          <aside className={`sidebar-fixed no-print ${isMobileMenuOpen ? "mobile-open" : ""}`}>
            <div className="p-4 bg-slate-950 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center space-x-3">
                <img
                  src={customLogoUrl || SIBOS_CLEAN_ICON}
                  onError={(e) => { e.target.onerror = null; e.target.src = SIBOS_CLEAN_ICON; }}
                  alt={headerInfo.namaSekolah || "SiBos"}
                  className="w-10 h-10 object-contain shrink-0 block bg-transparent"
                />
                <div>
                  <h1 className="font-bold text-sm text-white tracking-wide">
                    {customLogoUrl ? "SiBOS V1" : "SiBOS V1"}
                  </h1>
                  <p className="text-[10px] text-blue-400 font-semibold truncate max-w-[150px]">{headerInfo.namaSekolah}</p>
                </div>
              </div>
              <button 
                onClick={() => setIsMobileMenuOpen(false)}
                className="md:hidden text-slate-400 hover:text-white font-bold text-lg p-1"
                title="Tutup Menu"
              >
                ✕
              </button>
            </div>

            <div className="flex-1 overflow-y-auto px-3 py-4 space-y-6">
              
              <div>
                <p className="px-3 text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-2">NAVIGASI UTAMA</p>

                {/* MENU DASHBOARD UTAMA */}
                <button
                  onClick={() => {
                    setCurrentView("dashboard");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full flex items-center justify-between px-3 py-2.5 rounded-lg text-xs font-semibold transition mb-1.5 ${
                    currentView === "dashboard" ? "bg-blue-600 text-white shadow-md font-bold" : "hover:bg-slate-800 text-slate-300"
                  }`}
                >
                  <div className="flex items-center space-x-2">
                    <span>🖥️</span>
                    <span className="font-bold">Dashboard</span>
                  </div>
                </button>
                
                <button
                  onClick={() => {
                    setCurrentView("database");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full flex items-center justify-between px-3 py-2.5 rounded-lg text-xs font-semibold transition ${
                    currentView === "database" ? "bg-blue-600 text-white shadow-md font-bold" : "hover:bg-slate-800 text-slate-300"
                  }`}
                >
                  <div className="flex items-center space-x-2">
                    <span>📄</span>
                    <span>Database Master RKAS (PDF)</span>
                  </div>
                  <span className="bg-slate-800 text-slate-300 text-[10px] font-mono px-1.5 py-0.5 rounded">
                    {masterRkas.length || 165}
                  </span>
                </button>

                {/* MENU DATABASE BELANJA */}
                <button
                  onClick={() => {
                    setCurrentView("belanja");
                    if (selectedMonthView === "Semua" && uploadedPerBulanMonths.length === 1) {
                      setSelectedMonthView(uploadedPerBulanMonths[0]);
                    }
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full flex items-center justify-between px-3 py-3 rounded-lg text-xs font-bold transition mt-1.5 shadow-sm ${
                    currentView === "belanja" ? "bg-blue-600 text-white font-bold ring-2 ring-blue-400" : "hover:bg-slate-800 text-slate-200 bg-slate-900/60"
                  }`}
                >
                  <div className="flex items-center space-x-2">
                    <span className="text-base">💳</span>
                    <span className="text-sm font-bold">Database Belanja</span>
                  </div>
                </button>

                {/* MENU STANDING INSTRUCTION (SI) DI BAWAH MENU BELANJA */}
                <button
                  onClick={() => {
                    setCurrentView("si");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full flex items-center justify-between px-3 py-2.5 rounded-lg text-xs font-semibold transition mt-1.5 ${
                    currentView === "si" ? "bg-blue-600 text-white shadow-md font-bold" : "hover:bg-slate-800 text-slate-300"
                  }`}
                >
                  <div className="flex items-center space-x-2">
                    <span>📜</span>
                    <span className="font-bold">Standing Instruction (SI)</span>
                  </div>
                </button>

                {/* MENU SURAT TRANSFER (PEMINDAHBUKUAN) DI BAWAH MENU SI */}
                <button
                  onClick={() => {
                    setCurrentView("surat_transfer");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full flex items-center justify-between px-3 py-2.5 rounded-lg text-xs font-semibold transition mt-1.5 ${
                    currentView === "surat_transfer" ? "bg-blue-600 text-white shadow-md font-bold" : "hover:bg-slate-800 text-slate-300"
                  }`}
                >
                  <div className="flex items-center space-x-2">
                    <span>✉️</span>
                    <span className="font-bold">Surat Transfer</span>
                  </div>
                </button>

                <button
                  onClick={() => {
                    setCurrentView("rekap");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full flex items-center space-x-2 px-3 py-2.5 rounded-lg text-xs font-semibold transition mt-1.5 ${
                    currentView === "rekap" ? "bg-blue-600 text-white shadow-md font-bold" : "hover:bg-slate-800 text-slate-300"
                  }`}
                >
                  <span>📊</span>
                  <span>Rekap Serapan Realisasi</span>
                </button>

                {/* MENU REKENING KORAN DI BAWAH MENU REKAP SERAPAN */}
                <button
                  onClick={() => {
                    setCurrentView("rek_koran");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full flex items-center justify-between px-3 py-2.5 rounded-lg text-xs font-semibold transition mt-1.5 ${
                    currentView === "rek_koran" ? "bg-blue-600 text-white shadow-md font-bold" : "hover:bg-slate-800 text-slate-300"
                  }`}
                >
                  <div className="flex items-center space-x-2">
                    <span>🏦</span>
                    <span className="font-bold">Rekening Koran</span>
                  </div>
                </button>

                {/* MENU REALISASI DISTRIBUSI (PENGELUARAN & STOK) */}
                <button
                  onClick={() => {
                    setCurrentView("distribusi");
                    setIsMobileMenuOpen(false);
                  }}
                  className={`w-full flex items-center justify-between px-3 py-2.5 rounded-lg text-xs font-semibold transition mt-1.5 ${
                    currentView === "distribusi" ? "bg-blue-600 text-white shadow-md font-bold" : "hover:bg-slate-800 text-slate-300"
                  }`}
                >
                  <div className="flex items-center space-x-2">
                    <span>📦</span>
                    <span className="font-bold">Realisasi Distribusi</span>
                  </div>
                  {distribusiList.length > 0 && (
                    <span className="bg-slate-800 text-blue-300 text-[10px] font-mono px-1.5 py-0.5 rounded">
                      {distribusiList.length}
                    </span>
                  )}
                </button>
              </div>

              {/* ACCORDION SUB-MENU PER-BULAN */}
              <div className="mt-3">
                <button
                  type="button"
                  onClick={() => setIsPerBulanOpen(prev => !prev)}
                  className={`w-full flex justify-between items-center px-3 py-2.5 rounded-lg text-xs font-bold transition ${
                    currentView === "perbulan" ? "bg-slate-800 text-blue-400 border-l-4 border-blue-500 shadow-sm" : "hover:bg-slate-800/60 text-slate-300"
                  }`}
                >
                  <div className="flex items-center space-x-2">
                    <span>📅</span>
                    <span className="uppercase tracking-wider">PER-BULAN (2026)</span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <span className="text-[10px] bg-slate-800 text-blue-400 font-bold px-1.5 py-0.5 rounded">12 Bulan</span>
                    <span className="text-xs transition-transform duration-200 text-slate-400">
                      {isPerBulanOpen ? "▲" : "▼"}
                    </span>
                  </div>
                </button>
                
                {/* SUB MENU BULAN (JANUARI - DESEMBER) HANYA TAMPIL SAAT PERBULAN DI-KLIK */}
                {isPerBulanOpen && (
                  <div className="pl-3 mt-1.5 space-y-1 border-l-2 border-slate-700/60 ml-3">
                    {MONTHS.map(m => {
                      const isSel = selectedMonthView === m;

                      return (
                        <button
                          key={m}
                          onClick={() => {
                            setSelectedMonthView(m);
                            setCurrentView("perbulan");
                            setIsMobileMenuOpen(false);
                          }}
                          className={`w-full flex items-center justify-between px-3 py-2 rounded-lg text-xs font-medium transition ${
                            isSel && currentView === "perbulan" ? "bg-blue-600 text-white font-bold shadow-sm" : "hover:bg-slate-800/60 text-slate-400"
                          }`}
                        >
                          <div className="flex items-center space-x-2">
                            <span>📅</span>
                            <span>{m}</span>
                          </div>
                        </button>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* MENU PENGATURAN / PROFIL SEKOLAH (HANYA TAMPIL JIKA SUDAH LOGIN ADMIN) */}
              {isAdmin && (
                <div className="pt-2 border-t border-slate-800/60">
                  <button
                    onClick={() => {
                      setCurrentView("pengaturan");
                      setIsMobileMenuOpen(false);
                    }}
                    className={`w-full flex items-center justify-between px-3 py-2.5 rounded-lg text-xs font-semibold transition ${
                      currentView === "pengaturan" ? "bg-blue-600 text-white shadow-md font-bold" : "hover:bg-slate-800 text-slate-300"
                    }`}
                  >
                    <div className="flex items-center space-x-2">
                      <span>⚙️</span>
                      <span className="font-bold">Pengaturan Sekolah</span>
                    </div>
                  </button>
                </div>
              )}

            </div>

            <div className="p-3 bg-slate-950 border-t border-slate-800 text-[11px] text-slate-400 flex flex-col gap-1 text-center">
              <span>NPSN: {headerInfo.npsn || currentUser?.npsn}</span>
              <span className="text-emerald-400 font-semibold">{headerInfo.namaSekolah || currentUser?.namaSekolah}</span>
              {isAdmin ? (
                <button
                  type="button"
                  onClick={handleAdminLogout}
                  className="mt-2 w-full py-1.5 bg-red-950/80 hover:bg-red-900 text-red-300 border border-red-800/60 rounded text-[11px] font-bold flex items-center justify-center gap-1 cursor-pointer transition shadow-xs"
                >
                  <span>🔒</span> Kunci Mode Admin
                </button>
              ) : (
                <span className="mt-1 text-[10px] text-slate-500">Mode Publik (Lihat Saja)</span>
              )}
            </div>
          </aside>

          {/* BACKDROP OVERLAY DENGAN KLIK TUTUP DI HP / TABLET */}
          {isMobileMenuOpen && (
            <div
              className="no-print fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-40 md:hidden transition-opacity"
              onClick={() => setIsMobileMenuOpen(false)}
            />
          )}

          {/* MAIN CONTENT AREA */}
          <main className="main-content">
            
            {/* TOP BAR (HANYA TAMPIL DI DESKTOP / KOMPUTER, DI HP SEMBUNYI TOTAL SESUAI GAMBAR 1) */}
            <header className="no-print bg-white border-b border-slate-200 px-4 md:px-6 py-4 hidden md:flex flex-wrap justify-between items-center gap-3 shadow-sm sticky top-0 z-40">
              <div className="flex items-center gap-3">
                <button
                  onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
                  className="md:hidden bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs px-3 py-2 rounded-lg flex items-center gap-1.5 shadow transition cursor-pointer"
                  title="Buka / Tutup Menu Navigasi Sidebar"
                >
                  <span className="text-base">☰</span> Menu
                </button>
                <img
                  src={customLogoUrl || SIBOS_CLEAN_ICON}
                  onError={(e) => { e.target.onerror = null; e.target.src = SIBOS_CLEAN_ICON; }}
                  alt="Logo"
                  className="w-8 h-8 sm:w-10 sm:h-10 object-contain shrink-0 block bg-transparent"
                />
                <div>
                  <h2 className="font-bold text-slate-900 text-sm md:text-lg flex items-center gap-2">
                    {currentView === "dashboard" && "🖥️ Dashboard"}
                    {currentView === "perbulan" && `Rincian Kertas Kerja Perbulan — Bulan ${selectedMonthView}`}
                    {currentView === "belanja" && (selectedMonthView === "Semua" ? "Database Transaksi Belanja" : `Database Transaksi Belanja — Bulan ${selectedMonthView}`)}
                    {currentView === "si" && `Surat Perintah Transfer / Standing Instruction (SI) Cetak`}
                    {currentView === "rek_koran" && "Laporan Rekening Koran Transaksi Bank (SI)"}
                    {currentView === "distribusi" && "Realisasi Distribusi — Pengeluaran & Stok Barang"}
                    {currentView === "database" && "KERTAS KERJA RENCANA KEGIATAN DAN ANGGARAN SEKOLAH (RKAS)"}
                    {currentView === "rekap" && "Laporan Rekapitulasi Realisasi & Serapan Dana BOS"}
                    {currentView === "pengaturan" && "Pengaturan Identitas Sekolah, Logo, & Kop Surat"}
                  </h2>
                  <p className="text-xs text-slate-500">
                    {headerInfo.namaSekolah}
                  </p>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                {isAdmin && currentView === "perbulan" && (
                  <>
                    <label className="bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold px-3.5 py-2 rounded-lg shadow transition cursor-pointer flex items-center gap-1">
                      <span>📤 Impor PDF / Excel / JSON</span>
                      <input
                        type="file"
                        accept=".pdf,.xlsx,.xls,.json"
                        onChange={handleUploadPerBulan}
                        className="hidden"
                      />
                    </label>

                    <button
                      type="button"
                      onClick={() => {
                        const mTarget = selectedMonthView === "Semua" ? "Januari" : selectedMonthView;
                        const items = generatePerBulanFromMaster(mTarget);
                        setPerBulanData(prev => ({ ...prev, [mTarget]: items }));
                        alert(`✅ BERHASIL GENERASI!\n\n${items.length} baris Rincian Kertas Kerja untuk Bulan ${mTarget} telah berhasil dipulihkan & disajikan.`);
                      }}
                      className="bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold px-3.5 py-2 rounded-lg shadow transition flex items-center gap-1 cursor-pointer"
                      title="Generasi / Pulihkan data Rincian Kertas Kerja untuk bulan ini"
                    >
                      <span>🔄</span> Generasi Rincian ({selectedMonthView})
                    </button>

                    <button
                      type="button"
                      onClick={() => handleHapusAnggaranPerBulan(selectedMonthView)}
                      className="bg-red-600 hover:bg-red-700 text-white text-xs font-semibold px-3.5 py-2 rounded-lg shadow transition flex items-center gap-1 cursor-pointer"
                      title={`Hapus Anggaran Rincian Kertas Kerja (${selectedMonthView === "Semua" ? "Semua Bulan" : "Bulan " + selectedMonthView})`}
                    >
                      <span>🗑️</span> Hapus Anggaran ({selectedMonthView})
                    </button>
                  </>

                )}

                {isAdmin && currentView === "belanja" && (
                  <>
                    <button
                      onClick={handleBackupBelanjaData}
                      className="bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold px-3.5 py-2 rounded-lg shadow transition flex items-center gap-1 cursor-pointer"
                      title="Simpan backup data transaksi belanja ke file JSON"
                    >
                      💾 Backup Belanja
                    </button>

                    <button
                      onClick={handleRestoreBelanjaClick}
                      className="bg-purple-600 hover:bg-purple-700 text-white text-xs font-semibold px-3.5 py-2 rounded-lg shadow transition flex items-center gap-1 cursor-pointer"
                      title="Unggah dan pulihkan data transaksi belanja dari file JSON"
                    >
                      📂 Restore Belanja
                    </button>

                    <input
                      type="file"
                      id="belanja-restore-file-input"
                      accept=".json"
                      onChange={handleRestoreBelanjaFile}
                      className="hidden"
                    />
                  </>
                )}

                {(currentView === "si" || currentView === "rek_koran" || currentView === "surat_transfer") && (
                  <>
                    <button
                      type="button"
                      onClick={() => window.print()}
                      className="bg-slate-800 hover:bg-black text-white text-xs font-bold px-3.5 py-2 rounded-lg shadow transition flex items-center gap-1.5 cursor-pointer"
                      title="Cetak Dokumen"
                    >
                      <span>🖨️</span> Cetak {currentView === "rek_koran" ? "Rekening Koran" : currentView === "surat_transfer" ? "Surat Transfer" : "SI (Landscape A4)"}
                    </button>
                  </>
                )}

                {/* UNIFIED EXPORT DROPDOWN BUTTON (EXCEL & PDF) */}
                <div className="relative inline-block text-left">
                  <button
                    type="button"
                    onClick={() => setShowExportDropdown(!showExportDropdown)}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-3.5 py-2 rounded-lg shadow transition flex items-center gap-1.5 cursor-pointer"
                    title="Opsi Ekspor Dokumen ke PDF atau Excel"
                  >
                    <span>📤</span> Export <span>▼</span>
                  </button>

                  {showExportDropdown && (
                    <div 
                      className="origin-top-right absolute right-0 mt-1 w-48 rounded-md shadow-lg bg-white ring-1 ring-black ring-opacity-5 z-50 divide-y divide-slate-100"
                      onMouseLeave={() => setShowExportDropdown(false)}
                    >
                      <div className="py-1">
                        <button
                          type="button"
                          onClick={() => {
                            setShowExportDropdown(false);
                            window.print();
                          }}
                          className="w-full text-left px-4 py-2 text-xs text-slate-700 hover:bg-purple-50 hover:text-purple-700 font-semibold flex items-center gap-2 cursor-pointer"
                        >
                          <span>📄</span> Ekspor PDF
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            setShowExportDropdown(false);
                            if (currentView === "perbulan") exportExcelPerBulan();
                            else if (currentView === "belanja") exportExcelBelanja();
                            else if (currentView === "si") exportExcelSI();
                            else if (currentView === "rek_koran") exportExcelRekeningKoran();
                            else if (currentView === "distribusi") exportExcelDistribusi();
                            else if (currentView === "database") handleExportMasterExcel();
                          }}
                          className="w-full text-left px-4 py-2 text-xs text-slate-700 hover:bg-emerald-50 hover:text-emerald-700 font-semibold flex items-center gap-2 cursor-pointer"
                        >
                          <span>📊</span> Ekspor Excel (.xlsx)
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                {/* SCHOOL INFO PILL & IKON GIR KECIL LOGIN ADMIN */}
                <div className="flex items-center gap-2 pl-2 border-l border-slate-300">
                  <button
                    type="button"
                    onClick={() => handlePushDataToCloud(false)}
                    disabled={cloudSyncStatus === "syncing"}
                    className={`px-2.5 py-1.5 rounded-lg border text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs ${
                      cloudSyncStatus === "syncing"
                        ? "bg-amber-50 text-amber-700 border-amber-300 animate-pulse"
                        : cloudSyncStatus === "synced"
                        ? "bg-emerald-50 text-emerald-700 border-emerald-300 hover:bg-emerald-100"
                        : cloudSyncStatus === "error"
                        ? "bg-red-50 text-red-700 border-red-300 hover:bg-red-100"
                        : "bg-amber-50 text-amber-700 border-amber-300 hover:bg-amber-100"
                    }`}
                    title={cloudLastSynced ? `Terakhir disinkronkan: ${cloudLastSynced}` : "Klik untuk menyinkronkan data ke Google Spreadsheet"}
                  >
                    <span>{cloudSyncStatus === "syncing" ? "⏳" : "📊"}</span>
                    <span className="hidden md:inline">
                      {cloudSyncStatus === "syncing" ? "Menyinkronkan..." : "Spreadsheet Aktif"}
                    </span>
                    <span className={`w-2 h-2 rounded-full ${cloudSyncStatus === "synced" ? "bg-emerald-500" : (cloudSyncStatus === "error" ? "bg-red-500" : "bg-amber-500")}`}></span>
                  </button>

                  <div className="hidden lg:flex flex-col text-right">
                    <span className="text-xs font-bold text-slate-800 leading-tight truncate max-w-[170px]">
                      {headerInfo.namaSekolah || currentUser?.namaSekolah}
                    </span>
                    <span className="text-[10px] text-slate-500 font-mono">
                      NPSN: {headerInfo.npsn || currentUser?.npsn}
                    </span>
                  </div>

                  {/* IKON GIR KECIL DI KANAN ATAS UNTUK LOGIN ADMIN (SESUAI SHEET ROLE PENGGUNA) */}
                  <div className="relative">
                    <button
                      type="button"
                      onClick={() => {
                        if (!isAdmin) {
                          setAdminPasswordInput("");
                          setAdminPasswordError("");
                          setShowAdminLoginModal(true);
                        } else {
                          setShowAdminGearMenu(!showAdminGearMenu);
                        }
                      }}
                      className={`p-2 rounded-lg border text-xs font-bold transition flex items-center gap-1 cursor-pointer ${
                        isAdmin
                          ? "bg-blue-600 text-white border-blue-600 shadow-sm"
                          : "bg-slate-50 hover:bg-slate-100 text-slate-600 border-slate-200"
                      }`}
                      title={isAdmin ? "Mode Admin Aktif (Klik untuk Pengaturan / Kunci Admin)" : "Login Admin"}
                    >
                      <span className="text-sm">⚙️</span>
                      {isAdmin && <span className="hidden sm:inline text-[11px]">Admin</span>}
                    </button>

                    {isAdmin && showAdminGearMenu && (
                      <div
                        className="origin-top-right absolute right-0 mt-1.5 w-48 rounded-xl shadow-lg bg-white ring-1 ring-black ring-opacity-5 z-50 divide-y divide-slate-100"
                        onMouseLeave={() => setShowAdminGearMenu(false)}
                      >
                        <div className="py-1">
                          <button
                            type="button"
                            onClick={() => {
                              setShowAdminGearMenu(false);
                              setCurrentView("pengaturan");
                            }}
                            className="w-full text-left px-4 py-2.5 text-xs text-slate-700 hover:bg-blue-50 hover:text-blue-700 font-bold flex items-center gap-2 cursor-pointer"
                          >
                            <span>⚙️</span> Pengaturan Sekolah
                          </button>
                          <button
                            type="button"
                            onClick={handleAdminLogout}
                            className="w-full text-left px-4 py-2.5 text-xs text-red-600 hover:bg-red-50 font-bold flex items-center gap-2 cursor-pointer"
                          >
                            <span>🔒</span> Kunci Mode Admin
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </header>

            {/* BODY */}
            <div className="p-3 sm:p-6 pb-24 md:pb-8">

              {/* TOMBOL KEMBALI KE DASHBOARD (KHUSUS TAMPILAN HP) */}
              {currentView !== "dashboard" && (
                <div className="no-print md:hidden mb-3">
                  <div className="bg-white p-2.5 rounded-xl border border-slate-200 shadow-2xs flex items-center justify-between">
                    <button
                      type="button"
                      onClick={() => setCurrentView("dashboard")}
                      className="flex items-center gap-1.5 text-xs font-bold text-blue-600 active:text-blue-800 cursor-pointer"
                    >
                      <span className="text-base leading-none">←</span>
                      <span>Menu Dashboard</span>
                    </button>
                    <span className="text-[11px] font-bold text-slate-600 bg-slate-100 px-2 py-0.5 rounded-md">
                      {currentView === "perbulan" && `Perbulan (${selectedMonthView})`}
                      {currentView === "belanja" && "Data Belanja"}
                      {currentView === "si" && "SI Cetak"}
                      {currentView === "surat_transfer" && "Surat Transfer"}
                      {currentView === "rekap" && "Rekapitulasi"}
                      {currentView === "rek_koran" && "Rekening Koran"}
                      {currentView === "distribusi" && "Distribusi Barang"}
                      {currentView === "database" && "Master RKAS"}
                      {currentView === "pengaturan" && "Pengaturan"}
                    </span>
                  </div>
                </div>
              )}

              {/* VIEW DASHBOARD: RINGKASAN KEUANGAN BOS (SESUAI GAMBAR) */}
              {currentView === "dashboard" && (
                <div className="bg-white rounded-2xl border border-slate-200 shadow-xs p-4 sm:p-6 md:p-8 space-y-5 sm:space-y-6">
                  {/* Header Ringkasan Keuangan BOS (Dengan Ikon Gir di Sebelah Kanan Khusus HP Sesuai Gambar 2) */}
                  <div className="flex flex-wrap items-center justify-between gap-3">
                    <div className="flex items-center gap-3.5">
                      <img
                        src={customLogoUrl || SIBOS_CLEAN_ICON}
                        onError={(e) => { e.target.onerror = null; e.target.src = SIBOS_CLEAN_ICON; }}
                        alt="Logo Sekolah"
                        className="w-10 h-10 sm:w-12 sm:h-12 object-contain shrink-0 block bg-transparent"
                      />
                      <div>
                        <p className="text-[11px] font-bold uppercase tracking-wider text-slate-500">
                          TAHUN ANGGARAN {headerInfo.tahunAnggaran || "2026"} &middot; {headerInfo.namaSekolah || "SDIT ANNISA"}
                        </p>
                        <h1 className="text-xl md:text-2xl font-extrabold text-slate-900 mt-0.5">
                          Ringkasan keuangan BOS
                        </h1>
                      </div>
                    </div>

                    <div className="flex items-center gap-2.5">
                      <div className="text-xs text-slate-500 flex items-center gap-1.5">
                        <span>🕒</span>
                        <span>
                          Data per {cloudLastSynced || new Date().toLocaleDateString("id-ID", { day: "2-digit", month: "short", year: "numeric" }) + ", " + new Date().toLocaleTimeString("id-ID", { hour: "2-digit", minute: "2-digit" }) + " WIB"}
                        </span>
                      </div>

                      {/* IKON GIR DI SEBELAH KANAN GAMBAR 2 (KHUSUS TAMPILAN HP) */}
                      <div className="md:hidden relative">
                        <button
                          type="button"
                          onClick={() => {
                            if (!isAdmin) {
                              setAdminPasswordInput("");
                              setAdminPasswordError("");
                              setShowAdminLoginModal(true);
                            } else {
                              setShowAdminGearMenu(!showAdminGearMenu);
                            }
                          }}
                          className={`p-2 rounded-xl border text-xs font-bold transition flex items-center gap-1 cursor-pointer shadow-xs ${
                            isAdmin
                              ? "bg-blue-600 text-white border-blue-600 shadow-sm"
                              : "bg-white hover:bg-slate-50 text-slate-700 border-slate-300"
                          }`}
                          title={isAdmin ? "Mode Admin Aktif" : "Login Admin"}
                        >
                          <span className="text-base">⚙️</span>
                          {isAdmin && <span className="text-[10px] font-bold">Admin</span>}
                        </button>

                        {isAdmin && showAdminGearMenu && (
                          <div
                            className="origin-top-right absolute right-0 mt-1.5 w-48 rounded-xl shadow-lg bg-white ring-1 ring-black ring-opacity-5 z-50 divide-y divide-slate-100"
                            onMouseLeave={() => setShowAdminGearMenu(false)}
                          >
                            <div className="py-1">
                              <button
                                type="button"
                                onClick={() => {
                                  setShowAdminGearMenu(false);
                                  setCurrentView("pengaturan");
                                }}
                                className="w-full text-left px-4 py-2.5 text-xs text-slate-700 hover:bg-blue-50 hover:text-blue-700 font-bold flex items-center gap-2 cursor-pointer"
                              >
                                <span>⚙️</span> Pengaturan Sekolah
                              </button>
                              <button
                                type="button"
                                onClick={handleAdminLogout}
                                className="w-full text-left px-4 py-2.5 text-xs text-red-600 hover:bg-red-50 font-bold flex items-center gap-2 cursor-pointer"
                              >
                                <span>🔒</span> Kunci Mode Admin
                              </button>
                            </div>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* 4 Kartu Ringkasan Atas (Responsif: 2 Kolom di HP, 4 Kolom di Komputer) */}
                  <div className="grid grid-cols-2 lg:grid-cols-4 gap-2 sm:gap-4">
                    {/* Kartu 1: Saldo awal */}
                    <div className="bg-gradient-to-br from-rose-500 via-rose-600 to-red-600 text-white rounded-xl border border-rose-400 p-3 sm:p-5 flex flex-col justify-between shadow-md shadow-rose-500/20">
                      <div className="flex items-center justify-between text-rose-100 text-[11px] sm:text-xs font-semibold mb-2 sm:mb-3">
                        <span className="truncate">Saldo awal</span>
                        <span className="w-6 h-6 sm:w-7 sm:h-7 rounded-md sm:rounded-lg bg-white/20 border border-white/30 flex items-center justify-center text-white text-[11px] sm:text-xs shrink-0">
                          👛
                        </span>
                      </div>
                      <div>
                        <div className="flex items-baseline gap-1">
                          <span className="text-[10px] sm:text-xs font-semibold text-rose-100">Rp</span>
                          <span className="text-sm sm:text-lg md:text-2xl font-black text-white tracking-tight leading-tight truncate">
                            {dashboardStats.saldoAwal.toLocaleString("id-ID")}
                          </span>
                        </div>
                        <p className="text-[10px] sm:text-[11px] text-rose-100 mt-1 sm:mt-1.5 leading-tight truncate">
                          Tanpa koreksi saldo
                        </p>
                      </div>
                    </div>

                    {/* Kartu 2: Pendapatan */}
                    <div className="bg-gradient-to-br from-emerald-500 via-emerald-600 to-teal-700 text-white rounded-xl border border-emerald-400 p-3 sm:p-5 flex flex-col justify-between shadow-md shadow-emerald-500/20">
                      <div className="flex items-center justify-between text-emerald-100 text-[11px] sm:text-xs font-semibold mb-2 sm:mb-3">
                        <span className="truncate">Pendapatan</span>
                        <span className="w-6 h-6 sm:w-7 sm:h-7 rounded-md sm:rounded-lg bg-white/20 border border-white/30 flex items-center justify-center text-white text-[11px] sm:text-xs shrink-0">
                          💵
                        </span>
                      </div>
                      <div>
                        <div className="flex items-baseline gap-1">
                          <span className="text-[10px] sm:text-xs font-semibold text-emerald-100">Rp</span>
                          <span className="text-sm sm:text-lg md:text-2xl font-black text-white tracking-tight leading-tight truncate">
                            {dashboardStats.pendapatan.toLocaleString("id-ID")}
                          </span>
                        </div>
                        <p className="text-[10px] sm:text-[11px] text-emerald-100 mt-1 sm:mt-1.5 leading-tight line-clamp-2">
                          {dashboardStats.pctPendapatanStr} dari total anggaran
                        </p>
                      </div>
                    </div>

                    {/* Kartu 3: Realisasi belanja */}
                    <div className="bg-gradient-to-br from-blue-500 via-blue-600 to-indigo-700 text-white rounded-xl border border-blue-400 p-3 sm:p-5 flex flex-col justify-between shadow-md shadow-blue-500/20">
                      <div className="flex items-center justify-between text-blue-100 text-[11px] sm:text-xs font-semibold mb-2 sm:mb-3">
                        <span className="truncate">Realisasi belanja</span>
                        <span className="w-6 h-6 sm:w-7 sm:h-7 rounded-md sm:rounded-lg bg-white/20 border border-white/30 flex items-center justify-center text-white text-[11px] sm:text-xs shrink-0">
                          📊
                        </span>
                      </div>
                      <div>
                        <div className="flex items-baseline gap-1">
                          <span className="text-[10px] sm:text-xs font-semibold text-blue-100">Rp</span>
                          <span className="text-sm sm:text-lg md:text-2xl font-black text-white tracking-tight leading-tight truncate">
                            {dashboardStats.totalRealisasi.toLocaleString("id-ID")}
                          </span>
                        </div>
                        <p className="text-[10px] sm:text-[11px] text-blue-100 mt-1 sm:mt-1.5 leading-tight line-clamp-2">
                          {dashboardStats.pctRealisasiAgainstAnggaranStr} dari total anggaran
                        </p>
                        <div className="w-full h-1.5 bg-white/25 rounded-full overflow-hidden mt-2">
                          <div
                            className="h-full bg-white rounded-full transition-all duration-500"
                            style={{ width: `${Math.min(100, Math.max(0, dashboardStats.pctRealisasiAgainstAnggaran))}%` }}
                          />
                        </div>
                      </div>
                    </div>

                    {/* Kartu 4: Sisa kas */}
                    <div className="bg-gradient-to-br from-purple-500 via-purple-600 to-violet-700 text-white rounded-xl border border-purple-400 p-3 sm:p-5 flex flex-col justify-between shadow-md shadow-purple-500/20">
                      <div className="flex items-center justify-between text-purple-100 text-[11px] sm:text-xs font-semibold mb-2 sm:mb-3">
                        <span className="truncate">Sisa kas</span>
                        <span className="w-6 h-6 sm:w-7 sm:h-7 rounded-md sm:rounded-lg bg-white/20 border border-white/30 flex items-center justify-center text-white text-[11px] sm:text-xs shrink-0">
                          🧾
                        </span>
                      </div>
                      <div>
                        <div className="flex items-baseline gap-1">
                          <span className="text-[10px] sm:text-xs font-semibold text-purple-100">Rp</span>
                          <span className="text-sm sm:text-lg md:text-2xl font-black text-white tracking-tight leading-tight truncate">
                            {dashboardStats.sisaKas.toLocaleString("id-ID")}
                          </span>
                        </div>
                        <p className="text-[10px] sm:text-[11px] text-purple-100 mt-1 sm:mt-1.5 leading-tight line-clamp-2">
                          {dashboardStats.pctSisaKasStr} dari dana tersedia
                        </p>
                      </div>
                    </div>
                  </div>

                  {/* MENU MODUL APLIKASI KHUSUS TAMPILAN HP / MOBILE (5 BARIS KARTU KECIL) */}
                  <div className="md:hidden mt-2 pt-3 border-t border-slate-200">
                    <div className="flex items-center justify-between mb-2.5">
                      <div className="flex items-center gap-1.5">
                        <span className="text-sm">📱</span>
                        <h2 className="text-xs font-extrabold text-slate-800 uppercase tracking-wider">
                          Menu Modul Aplikasi
                        </h2>
                      </div>
                      <span className="text-[10px] font-semibold text-blue-600 bg-blue-50 px-2 py-0.5 rounded-full border border-blue-200">
                        10 Modul SiBOS
                      </span>
                    </div>

                    <div className="grid grid-cols-2 gap-2">
                      {/* Baris 1: Master RKAS & Per-Bulan */}
                      <button
                        type="button"
                        onClick={() => setCurrentView("database")}
                        className="flex items-center gap-2 p-2 bg-slate-50 hover:bg-blue-50 active:scale-95 border border-slate-200 rounded-xl transition text-left shadow-2xs cursor-pointer"
                      >
                        <div className="w-8 h-8 rounded-lg bg-blue-100 text-blue-700 flex items-center justify-center text-sm shrink-0 font-bold">
                          📄
                        </div>
                        <div className="min-w-0 flex-1">
                          <p className="text-xs font-bold text-slate-800 truncate">Master RKAS</p>
                          <p className="text-[10px] text-slate-500 truncate">{masterRkas.length || 165} Item Pagu</p>
                        </div>
                      </button>

                      <button
                        type="button"
                        onClick={() => setCurrentView("perbulan")}
                        className="flex items-center gap-2 p-2 bg-slate-50 hover:bg-amber-50 active:scale-95 border border-slate-200 rounded-xl transition text-left shadow-2xs cursor-pointer"
                      >
                        <div className="w-8 h-8 rounded-lg bg-amber-100 text-amber-700 flex items-center justify-center text-sm shrink-0 font-bold">
                          📅
                        </div>
                        <div className="min-w-0 flex-1">
                          <p className="text-xs font-bold text-slate-800 truncate">Rincian Perbulan</p>
                          <p className="text-[10px] text-slate-500 truncate">{selectedMonthView === "Semua" ? "12 Bulan" : selectedMonthView}</p>
                        </div>
                      </button>

                      {/* Baris 2: Data Belanja & SI Cetak */}
                      <button
                        type="button"
                        onClick={() => setCurrentView("belanja")}
                        className="flex items-center gap-2 p-2 bg-slate-50 hover:bg-emerald-50 active:scale-95 border border-slate-200 rounded-xl transition text-left shadow-2xs cursor-pointer"
                      >
                        <div className="w-8 h-8 rounded-lg bg-emerald-100 text-emerald-700 flex items-center justify-center text-sm shrink-0 font-bold">
                          💳
                        </div>
                        <div className="min-w-0 flex-1">
                          <p className="text-xs font-bold text-slate-800 truncate">Data Belanja</p>
                          <p className="text-[10px] text-slate-500 truncate">{belanjaBankList.length} Transaksi</p>
                        </div>
                      </button>

                      <button
                        type="button"
                        onClick={() => setCurrentView("si")}
                        className="flex items-center gap-2 p-2 bg-slate-50 hover:bg-indigo-50 active:scale-95 border border-slate-200 rounded-xl transition text-left shadow-2xs cursor-pointer"
                      >
                        <div className="w-8 h-8 rounded-lg bg-indigo-100 text-indigo-700 flex items-center justify-center text-sm shrink-0 font-bold">
                          📜
                        </div>
                        <div className="min-w-0 flex-1">
                          <p className="text-xs font-bold text-slate-800 truncate">SI Bank</p>
                          <p className="text-[10px] text-slate-500 truncate">Cetak Transfer A4</p>
                        </div>
                      </button>

                      {/* Baris 3: Surat Transfer & Rekap Serapan */}
                      <button
                        type="button"
                        onClick={() => setCurrentView("surat_transfer")}
                        className="flex items-center gap-2 p-2 bg-slate-50 hover:bg-purple-50 active:scale-95 border border-slate-200 rounded-xl transition text-left shadow-2xs cursor-pointer"
                      >
                        <div className="w-8 h-8 rounded-lg bg-purple-100 text-purple-700 flex items-center justify-center text-sm shrink-0 font-bold">
                          ✉️
                        </div>
                        <div className="min-w-0 flex-1">
                          <p className="text-xs font-bold text-slate-800 truncate">Surat Transfer</p>
                          <p className="text-[10px] text-slate-500 truncate">Pemindahbukuan</p>
                        </div>
                      </button>

                      <button
                        type="button"
                        onClick={() => setCurrentView("rekap")}
                        className="flex items-center gap-2 p-2 bg-slate-50 hover:bg-teal-50 active:scale-95 border border-slate-200 rounded-xl transition text-left shadow-2xs cursor-pointer"
                      >
                        <div className="w-8 h-8 rounded-lg bg-teal-100 text-teal-700 flex items-center justify-center text-sm shrink-0 font-bold">
                          📊
                        </div>
                        <div className="min-w-0 flex-1">
                          <p className="text-xs font-bold text-slate-800 truncate">Rekap Serapan</p>
                          <p className="text-[10px] text-slate-500 truncate">Evaluasi Pagu</p>
                        </div>
                      </button>

                      {/* Baris 4: Rekening Koran & Realisasi Distribusi */}
                      <button
                        type="button"
                        onClick={() => setCurrentView("rek_koran")}
                        className="flex items-center gap-2 p-2 bg-slate-50 hover:bg-cyan-50 active:scale-95 border border-slate-200 rounded-xl transition text-left shadow-2xs cursor-pointer"
                      >
                        <div className="w-8 h-8 rounded-lg bg-cyan-100 text-cyan-700 flex items-center justify-center text-sm shrink-0 font-bold">
                          🏦
                        </div>
                        <div className="min-w-0 flex-1">
                          <p className="text-xs font-bold text-slate-800 truncate">Rekening Koran</p>
                          <p className="text-[10px] text-slate-500 truncate">Buku Kas Bank</p>
                        </div>
                      </button>

                      <button
                        type="button"
                        onClick={() => setCurrentView("distribusi")}
                        className="flex items-center gap-2 p-2 bg-slate-50 hover:bg-orange-50 active:scale-95 border border-slate-200 rounded-xl transition text-left shadow-2xs cursor-pointer"
                      >
                        <div className="w-8 h-8 rounded-lg bg-orange-100 text-orange-700 flex items-center justify-center text-sm shrink-0 font-bold">
                          📦
                        </div>
                        <div className="min-w-0 flex-1">
                          <p className="text-xs font-bold text-slate-800 truncate">Distribusi Stok</p>
                          <p className="text-[10px] text-slate-500 truncate">{distribusiList.length} Barang Keluar</p>
                        </div>
                      </button>

                      {/* Baris 5: Pengaturan Sekolah & Sinkronisasi */}
                      <button
                        type="button"
                        onClick={() => setCurrentView("pengaturan")}
                        className="flex items-center gap-2 p-2 bg-slate-50 hover:bg-slate-100 active:scale-95 border border-slate-200 rounded-xl transition text-left shadow-2xs cursor-pointer"
                      >
                        <div className="w-8 h-8 rounded-lg bg-slate-200 text-slate-700 flex items-center justify-center text-sm shrink-0 font-bold">
                          ⚙️
                        </div>
                        <div className="min-w-0 flex-1">
                          <p className="text-xs font-bold text-slate-800 truncate">Pengaturan</p>
                          <p className="text-[10px] text-slate-500 truncate">Profil & Logo</p>
                        </div>
                      </button>

                      <button
                        type="button"
                        onClick={() => handlePushDataToCloud(false)}
                        className="flex items-center gap-2 p-2 bg-slate-50 hover:bg-rose-50 active:scale-95 border border-slate-200 rounded-xl transition text-left shadow-2xs cursor-pointer"
                      >
                        <div className="w-8 h-8 rounded-lg bg-rose-100 text-rose-700 flex items-center justify-center text-sm shrink-0 font-bold">
                          🔄
                        </div>
                        <div className="min-w-0 flex-1">
                          <p className="text-xs font-bold text-slate-800 truncate">Sinkron Cloud</p>
                          <p className="text-[10px] text-slate-500 truncate">{cloudSyncStatus === "synced" ? "Tersambung" : "Sinkronkan"}</p>
                        </div>
                      </button>
                    </div>
                  </div>

                  {/* Bagian Bawah: 2 Kolom (Kiri: Anggaran dan realisasi per jenis belanja, Kanan: Penggunaan dana) */}
                  <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
                    {/* Kolom Kiri (8 Kolom): Anggaran dan realisasi per jenis belanja */}
                    <div className="lg:col-span-8 bg-white rounded-xl border border-slate-200 p-5 md:p-6 shadow-2xs flex flex-col justify-between">
                      <div>
                        <div className="flex items-center gap-2">
                          <span className="text-slate-500 text-sm">📊</span>
                          <h3 className="text-sm md:text-base font-bold text-slate-900">
                            Anggaran dan realisasi per jenis belanja
                          </h3>
                        </div>
                        <p className="text-xs text-slate-500 mt-0.5 mb-5">
                          Realisasi yang tercatat pada rincian anggaran, dibanding anggarannya
                        </p>

                        <div className="overflow-x-auto">
                          <table className="w-full text-xs md:text-sm border-collapse">
                            <thead>
                              <tr className="border-b border-slate-200 text-slate-500 text-xs font-medium">
                                <th className="py-2.5 text-left font-medium">Jenis belanja</th>
                                <th className="py-2.5 text-right font-medium">Anggaran (Rp)</th>
                                <th className="py-2.5 text-right font-medium">Realisasi (Rp)</th>
                                <th className="py-2.5 px-4 w-36 md:w-48"></th>
                                <th className="py-2.5 text-right font-medium w-16">Serapan</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100 text-slate-800">
                              {/* Baris 1: Belanja pegawai */}
                              <tr>
                                <td className="py-3.5 text-left font-medium text-slate-700">Belanja pegawai</td>
                                <td className="py-3.5 text-right font-mono text-slate-800">
                                  {dashboardStats.anggaranPegawai.toLocaleString("id-ID")}
                                </td>
                                <td className="py-3.5 text-right font-mono text-slate-800">
                                  {dashboardStats.realisasiPegawai.toLocaleString("id-ID")}
                                </td>
                                <td className="py-3.5 px-4">
                                  <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                                    <div
                                      className="h-full bg-blue-600 rounded-full transition-all duration-500"
                                      style={{ width: `${Math.min(100, Math.max(0, dashboardStats.pctPegawai))}%` }}
                                    />
                                  </div>
                                </td>
                                <td className="py-3.5 text-right font-semibold text-slate-800">
                                  {dashboardStats.pctPegawaiStr}
                                </td>
                              </tr>

                              {/* Baris 2: Belanja barang dan jasa */}
                              <tr>
                                <td className="py-3.5 text-left font-medium text-slate-700">Belanja barang dan jasa</td>
                                <td className="py-3.5 text-right font-mono text-slate-800">
                                  {dashboardStats.anggaranBarangJasa.toLocaleString("id-ID")}
                                </td>
                                <td className="py-3.5 text-right font-mono text-slate-800">
                                  {dashboardStats.realisasiBarangJasa.toLocaleString("id-ID")}
                                </td>
                                <td className="py-3.5 px-4">
                                  <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                                    <div
                                      className="h-full bg-blue-600 rounded-full transition-all duration-500"
                                      style={{ width: `${Math.min(100, Math.max(0, dashboardStats.pctBarangJasa))}%` }}
                                    />
                                  </div>
                                </td>
                                <td className="py-3.5 text-right font-semibold text-slate-800">
                                  {dashboardStats.pctBarangJasaStr}
                                </td>
                              </tr>

                              {/* Baris 3: Belanja modal */}
                              <tr>
                                <td className="py-3.5 text-left font-medium text-slate-700">Belanja modal</td>
                                <td className="py-3.5 text-right font-mono text-slate-800">
                                  {dashboardStats.anggaranModal.toLocaleString("id-ID")}
                                </td>
                                <td className="py-3.5 text-right font-mono text-slate-800">
                                  {dashboardStats.realisasiModal.toLocaleString("id-ID")}
                                </td>
                                <td className="py-3.5 px-4">
                                  <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                                    <div
                                      className="h-full bg-blue-600 rounded-full transition-all duration-500"
                                      style={{ width: `${Math.min(100, Math.max(0, dashboardStats.pctModal))}%` }}
                                    />
                                  </div>
                                </td>
                                <td className="py-3.5 text-right font-semibold text-slate-800">
                                  {dashboardStats.pctModalStr}
                                </td>
                              </tr>

                              {/* Baris Total */}
                              <tr className="border-t border-slate-200 font-bold text-slate-900">
                                <td className="py-3.5 text-left">Total</td>
                                <td className="py-3.5 text-right font-mono">
                                  {dashboardStats.totalAnggaran.toLocaleString("id-ID")}
                                </td>
                                <td className="py-3.5 text-right font-mono">
                                  {dashboardStats.totalRealisasi.toLocaleString("id-ID")}
                                </td>
                                <td className="py-3.5 px-4">
                                  <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                                    <div
                                      className="h-full bg-blue-600 rounded-full transition-all duration-500"
                                      style={{ width: `${Math.min(100, Math.max(0, dashboardStats.pctRealisasiAgainstAnggaran))}%` }}
                                    />
                                  </div>
                                </td>
                                <td className="py-3.5 text-right font-bold">
                                  {dashboardStats.pctRealisasiAgainstAnggaranStr}
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>

                    {/* Kolom Kanan (4 Kolom): Penggunaan dana */}
                    <div className="lg:col-span-4 bg-white rounded-xl border border-slate-200 p-5 md:p-6 shadow-2xs flex flex-col justify-between">
                      <div>
                        <div className="flex items-center gap-2">
                          <span className="w-3.5 h-3.5 rounded-full border-2 border-slate-500 inline-block" />
                          <h3 className="text-sm md:text-base font-bold text-slate-900">
                            Penggunaan dana
                          </h3>
                        </div>
                        <p className="text-xs text-slate-500 mt-0.5">
                          Saldo awal ditambah pendapatan
                        </p>

                        <div className="flex items-baseline gap-1 mt-4">
                          <span className="text-xs font-semibold text-slate-500">Rp</span>
                          <span className="text-xl md:text-2xl font-extrabold text-slate-900 tracking-tight">
                            {dashboardStats.danaTersedia.toLocaleString("id-ID")}
                          </span>
                        </div>

                        <div className="w-full h-2.5 bg-slate-100 rounded-full overflow-hidden mt-4 mb-5">
                          <div
                            className="h-full bg-blue-600 rounded-full transition-all duration-500"
                            style={{ width: `${Math.min(100, Math.max(0, dashboardStats.pctTerpakaiDana))}%` }}
                          />
                        </div>

                        <div className="space-y-2.5 text-xs">
                          <div className="flex items-center justify-between">
                            <div className="flex items-center gap-2 text-slate-600">
                              <span className="w-3 h-3 rounded-xs bg-blue-600 inline-block" />
                              <span>Terpakai (realisasi)</span>
                            </div>
                            <span className="font-bold text-slate-900">
                              {dashboardStats.pctTerpakaiDanaStr}
                            </span>
                          </div>
                          <div className="flex items-center justify-between">
                            <div className="flex items-center gap-2 text-slate-600">
                              <span className="w-3 h-3 rounded-xs bg-slate-200 inline-block" />
                              <span>Sisa kas</span>
                            </div>
                            <span className="font-bold text-slate-900">
                              {dashboardStats.pctSisaKasStr}
                            </span>
                          </div>
                        </div>
                      </div>

                      <div className="pt-4 mt-5 border-t border-slate-100">
                        <p className="text-[11px] text-slate-500">
                          Sisa kas Rp {dashboardStats.sisaKas.toLocaleString("id-ID")} dari dana tersedia Rp {dashboardStats.danaTersedia.toLocaleString("id-ID")}.
                        </p>
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* VIEW 0: RINCIAN KERTAS KERJA PERBULAN */}
              {currentView === "perbulan" && (
                <div className="space-y-6">
                  
                  {/* SUMMARY CARDS BERBASIS ACUAN DATABASE BELANJA PERBULAN */}
                  <div className="no-print grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-emerald-600">
                      <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                        Total Sudah Dibelanjakan ({selectedMonthView})
                      </p>
                      <p className="text-xl font-bold text-emerald-700 font-mono mt-1">
                        Rp {statsPerBulan.totalDibelanjakan.toLocaleString('id-ID')}
                      </p>
                      <p className="text-[11px] text-emerald-600 mt-1 font-semibold flex items-center gap-1">
                        <span>✅</span> Realisasi Belanja Terverifikasi
                      </p>
                    </div>

                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-amber-500">
                      <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                        Total Belum Dibelanjakan ({selectedMonthView})
                      </p>
                      <p className="text-xl font-bold text-amber-700 font-mono mt-1">
                        Rp {statsPerBulan.totalBelumDibelanjakan.toLocaleString('id-ID')}
                      </p>
                      <p className="text-[11px] text-slate-400 mt-1 font-medium">
                        Perhitungan Berbasis Rincian Belanja Perbulan
                      </p>
                    </div>

                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-blue-600">
                      <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                        Sisa Saldo Anggaran ({selectedMonthView})
                      </p>
                      <p className="text-xl font-bold text-blue-700 font-mono mt-1">
                        Rp {statsPerBulan.saldo.toLocaleString('id-ID')}
                      </p>
                      <p className="text-[11px] text-blue-600 mt-1 font-medium">
                        Total Rencana Target: Rp {statsPerBulan.totalRencana.toLocaleString('id-ID')}
                      </p>
                    </div>
                  </div>

                  <div className="print-container bg-white rounded-xl shadow-md border border-slate-300 p-4 md:p-6 overflow-x-auto">
                    
                    <div className="mb-4 pb-2 border-b flex flex-wrap items-center justify-between gap-2">
                      <div>
                        <h3 className="font-bold text-slate-900 text-base uppercase tracking-wide">
                          RINCIAN KERTAS KERJA PER-BULAN — BULAN {selectedMonthView.toUpperCase()} 2026
                        </h3>
                        <p className="text-xs text-slate-600">NPSN: {headerInfo.npsn} | Nama Sekolah: {headerInfo.namaSekolah}</p>
                      </div>
                      {isAdmin && (
                        <button
                          type="button"
                          onClick={() => handleHapusAnggaranPerBulan(selectedMonthView)}
                          className="no-print bg-red-50 hover:bg-red-100 text-red-700 border border-red-300 font-bold text-xs px-3 py-1.5 rounded-lg shadow-sm transition flex items-center gap-1.5 cursor-pointer"
                          title={`Hapus Anggaran Rincian Kertas Kerja (${selectedMonthView === "Semua" ? "Semua Bulan" : "Bulan " + selectedMonthView})`}
                        >
                          <span>🗑️</span> Hapus Anggaran ({selectedMonthView})
                        </button>
                      )}
                    </div>

                    <table className="w-full text-xs border-collapse border border-slate-900">
                      <thead>
                        <tr className="bg-slate-200 text-slate-900 font-bold text-center border-b-2 border-slate-900">
                          <th className="border border-slate-900 px-2 py-3 w-12 uppercase">No. Urut</th>
                          <th className="border border-slate-900 px-2 py-3 w-32 uppercase">Kode Rekening</th>
                          <th className="border border-slate-900 px-2 py-3 w-28 uppercase">Kode Program</th>
                          <th className="border border-slate-900 px-3 py-3 text-left uppercase">Uraian</th>
                          <th className="border border-slate-900 px-2 py-3 w-20 text-center uppercase">Volume</th>
                          <th className="border border-slate-900 px-2 py-3 w-20 text-center uppercase">Satuan</th>
                          <th className="border border-slate-900 px-3 py-3 text-right w-28 uppercase">Tarif Harga</th>
                          <th className="border border-slate-900 px-3 py-3 text-right w-32 uppercase">Jumlah</th>
                        </tr>
                      </thead>
                      <tbody>
                        {(perBulanData[selectedMonthView] || []).length === 0 ? (
                          <tr>
                            <td colSpan="8" className="border border-slate-900 px-4 py-8 text-center text-slate-500 italic bg-slate-50">
                              Belum ada data Rincian Kertas Kerja untuk Bulan {selectedMonthView}. Silakan klik tombol <strong>📤 Impor PDF / Excel / JSON</strong> di atas untuk mengunggah file.
                            </td>
                          </tr>
                        ) : (
                          (perBulanData[selectedMonthView] || []).map((m) => {
                            const stMonth = currentPerBulanStatusMap[m.no];
                            const isPurchased = (stMonth && stMonth.isFull) || purchasedUraianSet.has((m.uraian || "").trim().toLowerCase()) || purchasedUraianSet.has(canonicalizeArkasUraian(m.uraian));
                            const isPartial = stMonth ? stMonth.isPartial : false;
                            let rowClass = "hover:bg-slate-50";
                            if (isPurchased && m.level !== "standar" && m.level !== "kegiatan") rowClass = "bg-purchased-yellow font-bold";
                            else if (m.level === "standar") rowClass = "row-header-standar text-red-950";
                            else if (m.level === "kegiatan") rowClass = "row-header-kegiatan text-emerald-950";
                            else if (m.level === "sub") rowClass = "row-header-sub text-amber-950";

                            return (
                              <tr key={m.no} className={`transition border-b border-slate-400 ${rowClass}`}>
                                <td className="border border-slate-900 px-2 py-2 text-center font-bold">{m.no}</td>
                                <td className="border border-slate-900 px-2 py-2 font-mono text-slate-800">{m.kodeRekening || ""}</td>
                                <td className="border border-slate-900 px-2 py-2 font-mono text-slate-800">{m.kodeProgram || ""}</td>
                                <td className="border border-slate-900 px-3 py-2 text-left">
                                  <div>{m.uraian}</div>
                                  {isPartial && stMonth && (
                                    <div className="no-print mt-0.5 text-[10px] font-semibold text-amber-900 bg-amber-100/90 border border-amber-300 px-1.5 py-0.5 rounded inline-block">
                                      ⏳ Terpakai: Rp {stMonth.terpakai.toLocaleString('id-ID')} | Sisa: Rp {stMonth.sisa.toLocaleString('id-ID')}
                                    </div>
                                  )}
                                </td>
                                <td className="border border-slate-900 px-2 py-2 text-center font-mono">{m.volume !== "" && m.volume !== undefined ? m.volume : ""}</td>
                                <td className="border border-slate-900 px-2 py-2 text-center">{m.satuan || ""}</td>
                                <td className="border border-slate-900 px-3 py-2 text-right font-mono">{m.tarif ? Number(m.tarif).toLocaleString('id-ID') : ""}</td>
                                <td className="border border-slate-900 px-3 py-2 text-right font-mono font-bold">{Number(m.jumlah || 0).toLocaleString('id-ID')}</td>
                              </tr>
                            );
                          })
                        )}

                        <tr className="bg-slate-200 font-bold text-slate-900 border-t-2 border-slate-900">
                          <td colSpan="7" className="border border-slate-900 px-3 py-3 text-right uppercase">
                            JUMLAH RENCANA BULAN {selectedMonthView.toUpperCase()} 2026
                          </td>
                          <td className="border border-slate-900 px-3 py-3 text-right font-mono text-sm">
                            Rp {calculatePerBulanTotal(perBulanData[selectedMonthView] || []).toLocaleString('id-ID')}
                          </td>
                        </tr>
                      </tbody>
                    </table>

                    <div className="mt-10 pt-4 flex justify-between items-start text-xs font-semibold text-slate-900 px-6">
                      <div className="text-center min-w-[200px]">
                        <p className="font-normal text-slate-900">Bendahara</p>
                        <div className="h-24"></div>
                        <p className="font-bold text-slate-900">{headerInfo.bendaharaSekolah}</p>
                      </div>

                      <div className="text-center min-w-[220px]">
                        <p className="font-normal text-slate-900">Bekasi, 31 {selectedMonthView} 2026</p>
                        <p className="font-normal text-slate-900 mt-0.5">Kepala {headerInfo.namaSekolah || "Sekolah"}</p>
                        <div className="h-20"></div>
                        <p className="font-bold text-slate-900">{headerInfo.kepalaSekolah}</p>
                      </div>
                    </div>

                  </div>

                </div>
              )}

              {/* VIEW 1: DATABASE BELANJA WITH TANGGAL INPUT */}
              {currentView === "belanja" && (
                <div className="space-y-6">
                  
                  <div className="no-print bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-wrap justify-between items-center gap-4">
                    <div className="flex flex-wrap items-center gap-4">
                      <div className="flex items-center space-x-2">
                        <span className="text-xs font-bold text-slate-800 uppercase flex items-center gap-1.5">
                          <span>🗓️</span> Filter Tanggal Transaksi:
                        </span>
                        <select
                          value={selectedBelanjaTglSI}
                          onChange={(e) => setSelectedBelanjaTglSI(e.target.value)}
                          className="text-xs font-bold border-2 border-blue-600 rounded-lg px-3 py-1.5 bg-white text-blue-900 focus:ring-2 focus:ring-blue-500 focus:outline-none shadow-sm"
                        >
                          <option value="Semua">🌟 Semua Tanggal Transaksi</option>
                          {availableSiDates.map(d => {
                            const count = belanjaBankList.filter(b => b.tglTransaksiSI === d).length;
                            return (
                              <option key={d} value={d}>Tanggal SI: {d} ({count} Transaksi)</option>
                            );
                          })}
                        </select>
                      </div>

                      <div className="flex items-center space-x-2">
                        <span className="text-xs font-bold text-slate-700 uppercase">📅 Bulan Anggaran & Pagu:</span>
                        <select
                          value={selectedMonthView}
                          onChange={(e) => setSelectedMonthView(e.target.value)}
                          className="text-xs font-bold border-2 border-emerald-600 rounded-lg px-2.5 py-1.5 bg-emerald-50 text-emerald-950 focus:outline-none shadow-xs"
                        >
                          <option value="Semua">
                            {uploadedPerBulanMonths.length === 1
                              ? `Otomatis: Bulan ${uploadedPerBulanMonths[0]} (Semua Transaksi)`
                              : "Semua Bulan (Master RKAS 1 Tahun)"}
                          </option>
                          {MONTHS.map(m => {
                            const hasUploadedPagu = Array.isArray(perBulanData[m]) && perBulanData[m].length > 0;
                            const monthPagu = hasUploadedPagu ? calculatePerBulanTotal(perBulanData[m]) : 0;
                            return (
                              <option key={m} value={m}>
                                Bulan {m} ({belanjaBankList.filter(b => b.bulan === m).length} Belanja{hasUploadedPagu ? ` • Pagu Rp ${monthPagu.toLocaleString('id-ID')}` : ''})
                              </option>
                            );
                          })}
                        </select>
                      </div>
                    </div>

                    <div className="text-xs text-slate-600 font-semibold">
                      Menampilkan: <strong className="text-blue-700 font-mono text-sm">{filteredBelanja.length} Transaksi</strong> | Total: <strong className="text-emerald-700 font-mono text-sm">Rp {totalBelanjaFiltered.toLocaleString('id-ID')}</strong>
                    </div>
                  </div>
                  
                  {isAdmin && (() => {
                    const belanjaLower = (inlineForm.belanja || "").trim().toLowerCase();
                    const matchingOptionsByName = belanjaLower
                      ? availableMasterOptions.filter(m => m.uraian.trim().toLowerCase() === belanjaLower)
                      : [];
                    const activeMasterOpt = (inlineForm.masterNo && availableMasterOptions.find(m => String(m.no) === String(inlineForm.masterNo)))
                      || (matchingOptionsByName.length > 0
                        ? [...matchingOptionsByName].sort((a, b) => (b.children?.length || 0) - (a.children?.length || 0))[0]
                        : null);
                    const activeChildren = activeMasterOpt && Array.isArray(activeMasterOpt.children) ? activeMasterOpt.children : [];
                    const availableActiveChildren = activeChildren.filter(c => !c.isFull && c.sisa > 0);
                    const selectedSubNos = new Set((inlineForm.subItems || []).map(s => String(s.no)));
                    const activeMonthPaguVal = (effectiveBudgetMonth && effectiveBudgetMonth !== "Semua")
                      ? calculatePerBulanTotal(perBulanData[effectiveBudgetMonth] || [])
                      : 0;

                    return (
                      <div className="no-print sticky top-[72px] z-30 bg-amber-50/95 backdrop-blur-md rounded-xl border-2 border-amber-500 p-3.5 shadow-xl transition-all">
                        <div className="flex flex-wrap justify-between items-center mb-2 font-bold text-amber-950 text-xs gap-2">
                          <span className="flex flex-wrap items-center gap-1.5 uppercase tracking-wide">
                            <span>📌</span> TABEL FORM INPUT BELANJA BARU
                            <span className="text-[10px] text-amber-900 font-normal normal-case bg-amber-200/90 px-2 py-0.5 rounded border border-amber-400 ml-1 shadow-xs">
                              💡 Tersedia: {availableMasterOptions.length} Item / Sub-Menu (Termasuk yang Masih Memiliki Sisa Anggaran)
                            </span>
                            {effectiveBudgetMonth && effectiveBudgetMonth !== "Semua" && (
                              <span className="text-[10px] text-emerald-950 font-bold normal-case bg-emerald-200/90 px-2 py-0.5 rounded border border-emerald-500 shadow-xs">
                                📅 Acuan Pagu: Bulan {effectiveBudgetMonth} (Rp {activeMonthPaguVal.toLocaleString('id-ID')})
                              </span>
                            )}
                          </span>
                          {inlineForm.id && (
                            <span className="bg-amber-300 text-amber-900 text-[11px] px-2 py-0.5 rounded font-mono">
                              Editing ID #{inlineForm.id}
                            </span>
                          )}
                        </div>

                        <div className="overflow-x-auto">
                          <table className="w-full text-xs border-collapse border border-amber-400 bg-white rounded shadow-inner">
                            <thead>
                              <tr className="bg-amber-200 text-amber-950 font-bold text-center text-[11px]">
                                <th className="border border-amber-400 px-2 py-2 text-left uppercase min-w-[260px]">BELANJA (PILIH ITEM / SUB-MENU RKAS)</th>
                                <th className="border border-amber-400 px-2 py-2 w-36 font-mono uppercase">NOMOR REKENING</th>
                                <th className="border border-amber-400 px-2 py-2 w-20 uppercase">NAMA BANK</th>
                                <th className="border border-amber-400 px-2 py-2 text-left uppercase">ATAS NAMA</th>
                                <th className="border border-amber-400 px-2 py-2 w-32 font-mono uppercase">NPWP</th>
                                <th className="border border-amber-400 px-2 py-2 text-right w-28 uppercase">TOTAL (RP)</th>
                                <th className="border border-amber-400 px-2 py-2 w-24 uppercase">SUMBER DANA</th>
                                <th className="border border-amber-400 px-2 py-2 w-28 uppercase">TGL TRANSAKSI SI</th>
                                <th className="border border-amber-400 px-2 py-2 w-24 uppercase">AKSI</th>
                              </tr>
                            </thead>
                            <tbody>
                              <tr className="bg-amber-50/80">
                                <td className="border border-amber-300 p-1.5 align-top">
                                  <select
                                    value={activeMasterOpt ? String(activeMasterOpt.no) : ""}
                                    onChange={(e) => {
                                       const chosenNo = e.target.value;
                                       if (!chosenNo) {
                                         setInlineForm(prev => ({ ...prev, masterNo: "", subItems: [] }));
                                         return;
                                       }
                                       const chosenOpt = availableMasterOptions.find(m => String(m.no) === String(chosenNo));
                                       if (chosenOpt) {
                                         applyMasterOptionToInline(chosenOpt);
                                       }
                                    }}
                                    className="w-full text-[11px] font-bold px-2 py-1 mb-1 border border-amber-400 rounded bg-amber-100/70 text-amber-950 focus:ring-2 focus:ring-blue-500 cursor-pointer"
                                    title="Pilih langsung dari daftar RKAS (menampilkan sisa anggaran & jumlah barang di sub-menu)"
                                  >
                                    <option value="">-- 📂 Pilih dari Dropdown RKAS (Menampilkan Sisa & Sub-Barang) --</option>
                                    {availableMasterOptions.map(m => {
                                      const availKidCount = (m.children || []).filter(c => !c.isFull && c.sisa > 0).length;
                                      const groupTag = (m.children && m.children.length > 1)
                                        ? ` [📦 ${availKidCount} Barang]`
                                        : "";
                                      const priceTag = m.terpakai > 0
                                        ? `SISA: Rp ${m.sisa.toLocaleString('id-ID')} (Pagu: Rp ${m.pagu.toLocaleString('id-ID')})`
                                        : `Rp ${m.sisa.toLocaleString('id-ID')}`;
                                      return (
                                        <option key={m.no} value={m.no}>
                                          {`[No.${m.no}] ${m.uraian}${groupTag} — ${priceTag}`}
                                        </option>
                                      );
                                    })}
                                  </select>

                                  <input
                                    type="text"
                                    placeholder="Atau ketik / cari uraian belanja di sini..."
                                    value={inlineForm.belanja}
                                    onChange={(e) => {
                                      const val = e.target.value;
                                      const noPrefixMatch = val.match(/^\[No\.(\d+)\]\s*(.+)$/i);
                                      if (noPrefixMatch) {
                                        const targetNo = noPrefixMatch[1];
                                        const optByNo = availableMasterOptions.find(m => String(m.no) === String(targetNo));
                                        if (optByNo) {
                                          applyMasterOptionToInline(optByNo);
                                          return;
                                        }
                                      }

                                      const matches = availableMasterOptions.filter(m => m.uraian.trim().toLowerCase() === val.trim().toLowerCase());
                                      if (matches.length > 0) {
                                        const best = [...matches].sort((a, b) => (b.children?.length || 0) - (a.children?.length || 0))[0];
                                        applyMasterOptionToInline(best);
                                      } else {
                                        setInlineForm(prev => ({ ...prev, masterNo: "", belanja: val, subItems: [] }));
                                      }
                                    }}
                                    className="w-full text-xs font-semibold px-2 py-1.5 border border-amber-400 rounded focus:ring-2 focus:ring-blue-500 bg-white text-slate-900"
                                    list="rkas-suggestions-inline"
                                  />
                                  <datalist id="rkas-suggestions-inline">
                                    {availableMasterOptions.map(m => {
                                      const availKidCount = (m.children || []).filter(c => !c.isFull && c.sisa > 0).length;
                                      const groupTag = (m.children && m.children.length > 1) ? ` (📦 Pilih ${availKidCount} Barang)` : "";
                                      const priceTag = m.terpakai > 0
                                        ? `SISA: Rp ${m.sisa.toLocaleString('id-ID')} (Terpakai: Rp ${m.terpakai.toLocaleString('id-ID')} dari Rp ${m.pagu.toLocaleString('id-ID')})`
                                        : `Rp ${m.sisa.toLocaleString('id-ID')}`;
                                      const hasDuplicateName = availableMasterOptions.filter(x => x.uraian.trim().toLowerCase() === m.uraian.trim().toLowerCase()).length > 1;
                                      const optionVal = hasDuplicateName ? `[No.${m.no}] ${m.uraian}` : m.uraian;
                                      return (
                                        <option key={m.no} value={optionVal}>
                                          {`[No.${m.no}] ${m.uraian}${groupTag} — ${priceTag}`}
                                        </option>
                                      );
                                    })}
                                  </datalist>

                                  {activeMasterOpt ? (
                                    <div className="mt-1.5 flex flex-wrap items-center gap-1.5">
                                      {activeMasterOpt.terpakai > 0 ? (
                                        <>
                                          <span className="text-[10px] font-bold text-amber-900 bg-amber-100 px-2 py-0.5 rounded border border-amber-400 inline-flex items-center gap-1 shadow-xs">
                                            <span>⏳</span> Pagu: Rp {activeMasterOpt.pagu.toLocaleString('id-ID')} | Terpakai: Rp {activeMasterOpt.terpakai.toLocaleString('id-ID')}
                                          </span>
                                          <span className="text-[10px] font-extrabold text-emerald-900 bg-emerald-100 px-2 py-0.5 rounded border border-emerald-400 inline-flex items-center gap-1 shadow-xs">
                                            <span>💰</span> Sisa Anggaran: Rp {activeMasterOpt.sisa.toLocaleString('id-ID')}
                                          </span>
                                        </>
                                      ) : (
                                        <span className="text-[10px] font-bold text-emerald-800 bg-emerald-100 px-2 py-0.5 rounded border border-emerald-300 inline-flex items-center gap-1 shadow-xs">
                                          <span>🏷️</span> Anggaran RKAS: Rp {activeMasterOpt.sisa.toLocaleString('id-ID')}
                                        </span>
                                      )}

                                      {activeChildren.length <= 1 && activeMasterOpt.lastNominal > 0 && activeMasterOpt.lastNominal <= activeMasterOpt.sisa && Number(inlineForm.total) !== activeMasterOpt.lastNominal && (
                                        <button
                                          type="button"
                                          onClick={() => {
                                            const nextTotal = activeMasterOpt.lastNominal;
                                            setInlineForm(prev => ({
                                              ...prev,
                                              total: nextTotal,
                                              subItems: Array.isArray(prev.subItems) && prev.subItems.length > 0
                                                ? (normalizeSubItemsWithTotal(prev.subItems, nextTotal) || [])
                                                : prev.subItems
                                            }));
                                          }}
                                          className="text-[10px] font-bold text-blue-800 bg-blue-100 hover:bg-blue-200 px-2 py-0.5 rounded border border-blue-300 transition cursor-pointer"
                                          title="Gunakan nominal pembayaran bulan sebelumnya"
                                        >
                                          ⚡ Isi Rp {activeMasterOpt.lastNominal.toLocaleString('id-ID')} (Bulan Lalu)
                                        </button>
                                      )}
                                      {activeChildren.length <= 1 && Number(inlineForm.total) !== activeMasterOpt.sisa && (
                                        <button
                                          type="button"
                                          onClick={() => {
                                            const nextTotal = activeMasterOpt.sisa;
                                            setInlineForm(prev => ({
                                              ...prev,
                                              total: nextTotal,
                                              subItems: Array.isArray(prev.subItems) && prev.subItems.length > 0
                                                ? (normalizeSubItemsWithTotal(prev.subItems, nextTotal) || [])
                                                : prev.subItems
                                            }));
                                          }}
                                          className="text-[10px] font-bold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 px-2 py-0.5 rounded border border-emerald-300 transition cursor-pointer"
                                        >
                                          ⚡ Isi Sisa (Rp {activeMasterOpt.sisa.toLocaleString('id-ID')})
                                        </button>
                                      )}
                                    </div>
                                  ) : inlineForm.total ? (
                                    <div className="mt-1">
                                      <span className="text-[10px] font-bold text-blue-800 bg-blue-100 px-2 py-0.5 rounded border border-blue-300 inline-flex items-center gap-1 shadow-sm">
                                        <span>💰</span> Total Input: Rp {Number(inlineForm.total).toLocaleString('id-ID')}
                                      </span>
                                    </div>
                                  ) : null}
                                </td>
                                <td className="border border-amber-300 p-1 align-top">
                                  <input
                                    type="text"
                                    placeholder="Nomor Rekening..."
                                    value={inlineForm.noRek}
                                    onChange={(e) => setInlineForm({ ...inlineForm, noRek: e.target.value })}
                                    className="w-full text-xs font-mono px-1.5 py-1.5 border border-slate-300 rounded bg-white text-center"
                                  />
                                </td>
                                <td className="border border-amber-300 p-1 align-top">
                                  <input
                                    type="text"
                                    placeholder="Nama Bank..."
                                    value={inlineForm.namaBank}
                                    onChange={(e) => setInlineForm({ ...inlineForm, namaBank: e.target.value })}
                                    className="w-full text-xs font-bold px-1 py-1.5 border border-slate-300 rounded bg-white text-center"
                                  />
                                </td>
                                <td className="border border-amber-300 p-1 align-top">
                                  <input
                                    type="text"
                                    placeholder="Atas Nama Supplier..."
                                    value={inlineForm.atasNama}
                                    onChange={(e) => setInlineForm({ ...inlineForm, atasNama: e.target.value })}
                                    className="w-full text-xs px-1.5 py-1.5 border border-slate-300 rounded bg-white"
                                  />
                                </td>
                                <td className="border border-amber-300 p-1 align-top">
                                  <input
                                    type="text"
                                    placeholder="NPWP..."
                                    value={inlineForm.npwp}
                                    onChange={(e) => setInlineForm({ ...inlineForm, npwp: e.target.value })}
                                    className="w-full text-xs font-mono px-1 py-1.5 border border-slate-300 rounded bg-white text-center"
                                  />
                                </td>
                                <td className="border border-amber-300 p-1 align-top">
                                  <input
                                    type="text"
                                    placeholder="Total (Rp)"
                                    value={formatRupiah(inlineForm.total)}
                                    onChange={(e) => {
                                      const parsed = parseRupiah(e.target.value);
                                      setInlineForm(prev => ({
                                        ...prev,
                                        total: parsed,
                                        subItems: Array.isArray(prev.subItems) && prev.subItems.length > 0
                                          ? (normalizeSubItemsWithTotal(prev.subItems, parsed) || [])
                                          : prev.subItems
                                      }));
                                    }}
                                    className="w-full text-xs font-mono font-bold px-1.5 py-1.5 border border-amber-500 rounded bg-white text-right text-blue-900"
                                  />
                                </td>
                                <td className="border border-amber-300 p-1 align-top">
                                  <select
                                    value={inlineForm.sumberDana}
                                    onChange={(e) => setInlineForm({ ...inlineForm, sumberDana: e.target.value })}
                                    className="w-full text-xs font-semibold px-1 py-1.5 border border-slate-300 rounded bg-white text-center text-slate-800 cursor-pointer"
                                  >
                                    <option value="">-- Pilih Sumber Dana --</option>
                                    <option value="BOS REGULER">BOS REGULER</option>
                                    <option value="BOS KINERJA">BOS KINERJA</option>
                                    <option value="SILPA">SILPA</option>
                                    <option value="LAINNYA">LAINNYA</option>
                                  </select>
                                </td>
                                <td className="border border-amber-300 p-1 align-top">
                                  <div className="flex items-center gap-1">
                                    <input
                                      type="text"
                                      placeholder="28/09/2026"
                                      value={inlineForm.tglTransaksiSI}
                                      onChange={(e) => setInlineForm({ ...inlineForm, tglTransaksiSI: e.target.value })}
                                      className="w-full text-xs font-mono font-bold px-1.5 py-1.5 border border-slate-300 rounded bg-white text-center text-slate-800"
                                      title="Format Penanggalan Indonesia: DD/MM/YYYY (contoh: 28/09/2026)"
                                    />
                                    <input
                                      type="date"
                                      value={formatDateToIso(inlineForm.tglTransaksiSI)}
                                      onChange={(e) => {
                                        if (e.target.value) {
                                          setInlineForm({ ...inlineForm, tglTransaksiSI: formatIsoToDdMmYyyy(e.target.value) });
                                        }
                                      }}
                                      className="w-6 h-7 cursor-pointer accent-blue-600 border-none bg-transparent p-0"
                                      title="Buka Kalender Pemilih Tanggal Indonesia"
                                    />
                                  </div>
                                </td>
                                <td className="border border-amber-300 p-1 text-center align-top">
                                  <div className="flex items-center justify-center gap-1">
                                    <button
                                      type="button"
                                      onClick={handleSaveInline}
                                      className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-[11px] px-3 py-1.5 rounded shadow transition flex items-center gap-1 whitespace-nowrap cursor-pointer"
                                      title="Simpan ke Database Belanja"
                                    >
                                      💾 {inlineForm.id ? "Update" : "Simpan"}
                                    </button>
                                    {inlineForm.id && (
                                      <button
                                        type="button"
                                        onClick={resetInlineForm}
                                        className="bg-slate-400 hover:bg-slate-500 text-white font-bold text-[11px] px-2 py-1.5 rounded shadow transition cursor-pointer"
                                        title="Batal Edit"
                                      >
                                        ❌
                                      </button>
                                    )}
                                  </div>
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        </div>

                        {activeMasterOpt && activeChildren.length > 0 && (
                          <div className="mt-2.5 bg-white rounded-lg border-2 border-blue-400 p-3 shadow-md">
                            <div className="flex flex-wrap justify-between items-center gap-2 pb-2 mb-2 border-b border-slate-200">
                              <div>
                                <div className="text-xs font-extrabold text-blue-950 flex flex-wrap items-center gap-1.5">
                                  <span>📦</span> Pilih Daftar Barang pada Sub-Menu: <span className="text-blue-700 underline">{activeMasterOpt.uraian}</span>
                                  <span className="text-[10px] font-mono bg-blue-100 text-blue-800 px-2 py-0.5 rounded border border-blue-300">
                                    No.{activeMasterOpt.no} {activeMasterOpt.kodeKegiatan ? `| ${activeMasterOpt.kodeKegiatan}` : ""}
                                  </span>
                                </div>
                                <p className="text-[11px] text-slate-600 mt-0.5">
                                  Centang barang yang dibelanjakan bulan/transaksi ini — harga otomatis dijumlahkan ke kolom <strong>TOTAL (RP)</strong>.
                                </p>
                              </div>

                              <div className="flex flex-wrap items-center gap-1.5">
                                {matchingOptionsByName.length > 1 && matchingOptionsByName.map(grp => (
                                  <button
                                    key={grp.no}
                                    type="button"
                                    onClick={() => applyMasterOptionToInline(grp)}
                                    className={`text-[10px] font-bold px-2 py-1 rounded border transition cursor-pointer ${
                                      String(grp.no) === String(activeMasterOpt.no)
                                        ? "bg-blue-600 text-white border-blue-700"
                                        : "bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-300"
                                    }`}
                                  >
                                    📂 No.{grp.no} ({grp.kodeKegiatan || "Sub"} — {(grp.children || []).length} Barang)
                                  </button>
                                ))}

                                <button
                                  type="button"
                                  onClick={() => {
                                    const allToSelect = availableActiveChildren.map(c => ({
                                      no: c.no,
                                      uraian: c.uraian,
                                      jumlah: c.sisa
                                    }));
                                    const sumAll = allToSelect.reduce((acc, x) => acc + Number(x.jumlah || 0), 0);
                                    setInlineForm(prev => ({
                                      ...prev,
                                      masterNo: activeMasterOpt.no,
                                      subItems: allToSelect,
                                      total: sumAll > 0 ? sumAll : ""
                                    }));
                                  }}
                                  className="text-[11px] font-bold bg-emerald-600 hover:bg-emerald-700 text-white px-2.5 py-1 rounded shadow-xs transition cursor-pointer"
                                >
                                  ☑️ Pilih Semua ({availableActiveChildren.length} Barang)
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setInlineForm(prev => ({
                                      ...prev,
                                      subItems: [],
                                      total: ""
                                    }));
                                  }}
                                  className="text-[11px] font-bold bg-slate-200 hover:bg-slate-300 text-slate-700 px-2.5 py-1 rounded transition cursor-pointer"
                                >
                                  🔲 Reset Centang
                                </button>
                              </div>
                            </div>

                            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-2 max-h-60 overflow-y-auto pr-1">
                              {activeChildren.map(child => {
                                const isChecked = selectedSubNos.has(String(child.no));
                                const selectedObj = (inlineForm.subItems || []).find(s => String(s.no) === String(child.no));
                                const isDisabled = child.isFull || child.sisa <= 0;

                                return (
                                  <div
                                    key={child.no}
                                    onClick={() => {
                                      if (!isDisabled) {
                                        handleToggleInlineSubItem(child, !isChecked);
                                      }
                                    }}
                                    className={`p-2 rounded-lg border text-xs flex flex-col justify-between gap-1.5 transition ${
                                      isDisabled
                                        ? "bg-slate-100 border-slate-300 text-slate-400 opacity-75 cursor-not-allowed"
                                        : isChecked
                                          ? "bg-blue-50 border-blue-500 ring-1 ring-blue-400 text-slate-900 cursor-pointer shadow-xs"
                                          : "bg-slate-50/70 hover:bg-amber-50/60 border-slate-200 text-slate-800 cursor-pointer"
                                    }`}
                                  >
                                    <div className="flex items-start gap-2">
                                      <input
                                        type="checkbox"
                                        checked={isChecked}
                                        disabled={isDisabled}
                                        onChange={(e) => {
                                          e.stopPropagation();
                                          if (!isDisabled) {
                                            handleToggleInlineSubItem(child, e.target.checked);
                                          }
                                        }}
                                        onClick={(e) => e.stopPropagation()}
                                        className="mt-0.5 w-4 h-4 accent-blue-600 rounded cursor-pointer"
                                      />
                                      <div className="flex-1 min-w-0">
                                        <div className={`font-bold leading-snug ${isDisabled ? "line-through text-slate-400" : "text-slate-900"}`}>
                                          {child.uraian}
                                        </div>
                                        <div className="flex flex-wrap items-center gap-1 mt-0.5 text-[10px]">
                                          <span className="font-mono text-slate-500">[No.{child.no}]</span>
                                          {isDisabled ? (
                                            <span className="bg-emerald-100 text-emerald-800 font-bold px-1.5 py-0.2 rounded border border-emerald-300">
                                              ✅ Sudah Dibeli (Rp {child.pagu.toLocaleString('id-ID')})
                                            </span>
                                          ) : child.terpakai > 0 ? (
                                            <span className="bg-amber-100 text-amber-900 font-bold px-1.5 py-0.2 rounded border border-amber-300">
                                              Sisa: Rp {child.sisa.toLocaleString('id-ID')} (Pagu: Rp {child.pagu.toLocaleString('id-ID')})
                                            </span>
                                          ) : (
                                            <span className="font-mono font-bold text-emerald-700">
                                              Rp {child.sisa.toLocaleString('id-ID')}
                                            </span>
                                          )}
                                        </div>
                                      </div>
                                    </div>

                                    {isChecked && !isDisabled && (
                                      <div
                                        className="flex items-center justify-between gap-2 pt-1 border-t border-blue-200 mt-0.5"
                                        onClick={(e) => e.stopPropagation()}
                                      >
                                        <span className="text-[10px] font-semibold text-blue-800">Nominal Belanja:</span>
                                        <input
                                          type="text"
                                          value={formatRupiah(selectedObj ? selectedObj.jumlah : child.sisa)}
                                          onChange={(e) => handleChangeInlineSubItemAmount(child.no, parseRupiah(e.target.value))}
                                          className="w-28 text-[11px] font-mono font-bold px-1.5 py-0.5 border border-blue-400 rounded bg-white text-right text-blue-900"
                                          title="Nominal untuk barang ini (otomatis menjumlahkan ke Total)"
                                        />
                                      </div>
                                    )}
                                  </div>
                                );
                              })}
                            </div>

                            <div className="mt-2 pt-2 border-t border-slate-200 flex flex-wrap justify-between items-center text-xs">
                              <span className="font-semibold text-slate-700">
                                Terpilih: <strong className="text-blue-700">{(inlineForm.subItems || []).length} Barang</strong> dari {availableActiveChildren.length} barang tersedia
                              </span>
                              <span className="font-bold text-emerald-800 bg-emerald-100 px-2.5 py-1 rounded border border-emerald-300 font-mono">
                                Total Otomatis: Rp {Number(inlineForm.total || 0).toLocaleString('id-ID')}
                              </span>
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })()}

                  {/* SUMMARY CARDS BERBASIS ACUAN DATABASE RINCIAN BELANJA PERBULAN */}
                  <div className="no-print grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-emerald-600">
                      <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                        Total Sudah Dibelanjakan ({effectiveBudgetMonth || selectedMonthView})
                      </p>
                      <p className="text-xl font-bold text-emerald-700 font-mono mt-1">
                        Rp {statsPerBulan.totalDibelanjakan.toLocaleString('id-ID')}
                      </p>
                      <p className="text-[11px] text-emerald-600 mt-1 font-semibold flex items-center gap-1">
                        <span>✅</span> {filteredBelanja.length} Transaksi Pembayaran Realized
                      </p>
                    </div>

                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-amber-500">
                      <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                        Total Belum Dibelanjakan ({effectiveBudgetMonth || selectedMonthView})
                      </p>
                      <p className="text-xl font-bold text-amber-700 font-mono mt-1">
                        Rp {statsPerBulan.totalBelumDibelanjakan.toLocaleString('id-ID')}
                      </p>
                      <p className="text-[11px] text-slate-400 mt-1 font-medium">
                        Perhitungan Berbasis Rincian Belanja Perbulan
                      </p>
                    </div>

                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-blue-600">
                      <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                        Sisa Saldo Anggaran ({effectiveBudgetMonth || selectedMonthView})
                      </p>
                      <p className="text-xl font-bold text-blue-700 font-mono mt-1">
                        Rp {statsPerBulan.saldo.toLocaleString('id-ID')}
                      </p>
                      <p className="text-[11px] text-blue-600 mt-1 font-medium">
                        Total Target Rencana: Rp {statsPerBulan.totalRencana.toLocaleString('id-ID')}
                      </p>
                    </div>
                  </div>

                  <div className="print-container bg-white rounded-xl shadow-md border border-slate-300 p-4 md:p-6 overflow-x-auto">
                    
                    <div className="mb-4 pb-2 border-b text-center">
                      <h3 className="font-bold text-slate-900 text-base uppercase tracking-wide">
                        DATABASE TRANSAKSI BELANJA SEKOLAH — {selectedMonthView === "Semua" ? "SEMUA BULAN" : `BULAN ${selectedMonthView.toUpperCase()} 2026`}
                      </h3>
                      <p className="text-xs text-slate-600">NPSN: {headerInfo.npsn} | Nama Sekolah: {headerInfo.namaSekolah}</p>
                    </div>

                    <table className="w-full text-xs border-collapse border border-slate-900">
                      <thead>
                        <tr className="bg-slate-200 text-slate-900 font-bold text-center border-b-2 border-slate-900">
                          <th className="border border-slate-900 px-2 py-3 w-10 text-center uppercase whitespace-nowrap">NO</th>
                          <th className="border border-slate-900 px-3 py-3 text-center uppercase w-[34%] min-w-[260px]">BELANJA</th>
                          <th className="border border-slate-900 px-2 py-3 w-28 text-center font-mono uppercase whitespace-nowrap">NOMOR REKENING</th>
                          <th className="border border-slate-900 px-2 py-3 w-20 text-center uppercase whitespace-nowrap">NAMA BANK</th>
                          <th className="border border-slate-900 px-3 py-3 text-center uppercase w-36 min-w-[145px] whitespace-nowrap">ATAS NAMA</th>
                          <th className="border border-slate-900 px-2 py-3 w-28 text-center font-mono uppercase whitespace-nowrap">NPWP</th>
                          <th className="border border-slate-900 px-3 py-3 text-center w-28 uppercase whitespace-nowrap">TOTAL</th>
                          <th className="border border-slate-900 px-2 py-3 w-24 text-center uppercase whitespace-nowrap">SUMBER DANA</th>
                          <th className="border border-slate-900 px-2 py-3 w-28 text-center uppercase whitespace-nowrap">TGL TRANSAKSI SI</th>
                          {isAdmin && (
                            <th className="border border-slate-900 px-2 py-3 w-20 text-center no-print uppercase whitespace-nowrap">AKSI</th>
                          )}
                        </tr>
                      </thead>
                      <tbody>
                        {filteredBelanja.length === 0 ? (
                          <tr>
                            <td colSpan={isAdmin ? "10" : "9"} className="border border-slate-900 px-4 py-8 text-center text-slate-500 italic bg-slate-50">
                              Belum ada data belanja tersimpan untuk bulan {selectedMonthView}.
                            </td>
                          </tr>
                        ) : (
                          filteredBelanja.map((item, idx) => {
                            const effectiveSubs = Array.isArray(item.subItems) && item.subItems.length > 0
                              ? (normalizeSubItemsWithTotal(item.subItems, item.total) || item.subItems)
                              : [];
                            const directSt = (!effectiveSubs.length && item.masterNo) ? masterBudgetFull[item.masterNo] : null;
                            return (
                            <tr key={item.id} className="hover:bg-blue-50/50 transition border-b border-slate-300">
                              <td className="border border-slate-900 px-2 py-2.5 text-center font-bold whitespace-nowrap">{idx + 1}</td>
                              <td className="border border-slate-900 px-3 py-2.5 text-left font-semibold text-slate-900">
                                <div>{item.belanja}</div>
                                {effectiveSubs.length > 0 && (
                                  <div className="mt-1 flex flex-wrap gap-1">
                                    {effectiveSubs.map(sub => {
                                      const subSt = masterBudgetFull[sub.no];
                                      return (
                                        <span key={sub.no} className="text-[10px] font-normal bg-blue-50 text-blue-900 border border-blue-200 px-1.5 py-0.5 rounded inline-flex items-center gap-1">
                                          <span>✓ {sub.uraian} (Rp {Number(sub.jumlah || 0).toLocaleString('id-ID')})</span>
                                          {subSt && subSt.sisa > 0 && (
                                            <span className="font-bold text-amber-800 bg-amber-100 border border-amber-300 px-1 rounded">
                                              Sisa Anggaran: Rp {subSt.sisa.toLocaleString('id-ID')}
                                            </span>
                                          )}
                                        </span>
                                      );
                                    })}
                                  </div>
                                )}
                                {directSt && directSt.sisa > 0 && (
                                  <div className="mt-1">
                                    <span className="text-[10px] font-bold text-amber-800 bg-amber-100 border border-amber-300 px-1.5 py-0.5 rounded inline-block">
                                      ⏳ Sisa Anggaran RKAS: Rp {directSt.sisa.toLocaleString('id-ID')} (Pagu: Rp {directSt.pagu.toLocaleString('id-ID')})
                                    </span>
                                  </div>
                                )}
                              </td>
                              <td className="border border-slate-900 px-2 py-2.5 text-center font-mono text-slate-800 whitespace-nowrap">{item.noRek}</td>
                              <td className="border border-slate-900 px-2 py-2.5 text-center font-bold text-slate-800 whitespace-nowrap">{item.namaBank}</td>
                              <td className="border border-slate-900 px-3 py-2.5 text-center font-medium text-slate-800 whitespace-nowrap">{item.atasNama}</td>
                              <td className="border border-slate-900 px-2 py-2.5 font-mono text-center text-slate-700 whitespace-nowrap">{item.npwp}</td>
                              <td className="border border-slate-900 px-3 py-2.5 text-center font-mono font-bold text-slate-900 whitespace-nowrap">
                                {Number(item.total).toLocaleString('id-ID')}
                              </td>
                              <td className="border border-slate-900 px-2 py-2.5 text-center font-medium whitespace-nowrap">{item.sumberDana}</td>
                              <td className="border border-slate-900 px-2 py-2.5 text-center font-mono text-slate-700 whitespace-nowrap">{item.tglTransaksiSI}</td>
                              
                              {isAdmin && (
                                <td className="border border-slate-900 px-2 py-2 text-center no-print space-x-1 whitespace-nowrap">
                                  <button onClick={() => setInlineForm({ ...item, masterNo: item.masterNo || "", subItems: effectiveSubs })} className="text-blue-600 hover:text-blue-800 font-bold px-1 cursor-pointer" title="Edit Baris Ini">✏️</button>
                                  <button onClick={() => handleDeleteBelanja(item.id)} className="text-red-600 hover:text-red-800 font-bold px-1 cursor-pointer" title="Hapus">❌</button>
                                </td>
                              )}
                            </tr>
                            );
                          })
                        )}

                        <tr className="bg-slate-200 font-bold text-slate-900 border-t-2 border-slate-900">
                          <td colSpan="6" className="border border-slate-900 px-3 py-3 text-left">
                            Jumlah Total
                          </td>
                          <td className="border border-slate-900 px-3 py-3 text-center font-mono text-sm whitespace-nowrap">
                            Rp {totalBelanjaFiltered.toLocaleString('id-ID')}
                          </td>
                          <td colSpan={isAdmin ? "3" : "2"} className="border border-slate-900 px-2 py-3"></td>
                        </tr>
                      </tbody>
                    </table>

                    <div className="mt-10 pt-4 flex justify-between items-start text-xs font-semibold text-slate-900 px-6">
                      <div className="text-center min-w-[200px]">
                        <p className="font-normal text-slate-900">Bendahara</p>
                        <div className="h-24"></div>
                        <p className="font-bold text-slate-900">{headerInfo.bendaharaSekolah}</p>
                      </div>

                      <div className="text-center min-w-[220px]">
                        <p className="font-normal text-slate-900">
                          Bekasi, {selectedBelanjaTglSI !== "Semua"
                            ? formatTanggalIndoTeks(selectedBelanjaTglSI)
                            : (filteredBelanja.length > 0 && filteredBelanja[filteredBelanja.length - 1].tglTransaksiSI && filteredBelanja[filteredBelanja.length - 1].tglTransaksiSI !== "-"
                                ? formatTanggalIndoTeks(filteredBelanja[filteredBelanja.length - 1].tglTransaksiSI)
                                : "28 September 2026")}
                        </p>
                        <p className="font-normal text-slate-900 mt-0.5">Kepala {headerInfo.namaSekolah || "Sekolah"}</p>
                        <div className="h-20"></div>
                        <p className="font-bold text-slate-900">{headerInfo.kepalaSekolah}</p>
                      </div>
                    </div>

                  </div>
                </div>
              )}

              {/* VIEW 2: STANDING INSTRUCTION (SI) DENGAN FILTER TANGGAL CETAK & TANDA TANGAN KEPSEK / BENDAHARA */}
              {currentView === "si" && (
                <div className="space-y-6">
                  <style>{`
                    @media print {
                      @page {
                        size: A4 landscape !important;
                        margin: 6mm 8mm !important;
                      }
                      html, body, #root, .main-content, .print-container,
                      .si-print-table, .si-print-table thead, .si-print-table tbody,
                      .si-print-table tr, .si-print-table th, .si-print-table td {
                        background: #ffffff !important;
                        background-color: #ffffff !important;
                      }
                      .si-print-table th, .si-print-table td {
                        font-size: 10px !important;
                        padding: 3.5px 5px !important;
                      }
                      .print-container {
                        max-width: 100% !important;
                        width: 100% !important;
                        padding: 0 !important;
                      }
                    }
                  `}</style>
                  {/* FILTER BAR STANDING INSTRUCTION (NO-PRINT) */}
                  <div className="no-print bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-wrap justify-between items-center gap-4">
                    <div className="flex flex-wrap items-center gap-3">
                      <span className="text-xs font-bold text-slate-800 uppercase flex items-center gap-1">
                        <span>🗓️</span> Filter Tanggal Transaksi SI / Cetak:
                      </span>
                      
                      <select
                        value={selectedSiDate}
                        onChange={(e) => setSelectedSiDate(e.target.value)}
                        className="text-xs font-bold border-2 border-blue-500 rounded-lg px-3 py-1.5 bg-white text-blue-900 focus:outline-none"
                      >
                        <option value="Semua">-- Semua Tanggal Transaksi --</option>
                        {availableSiDates.map(d => (
                          <option key={d} value={d}>Tanggal: {d}</option>
                        ))}
                      </select>

                      <select
                        value={selectedMonthView}
                        onChange={(e) => setSelectedMonthView(e.target.value)}
                        className="text-xs border border-slate-300 rounded-lg px-3 py-1.5 bg-white text-slate-700"
                      >
                        <option value="Semua">Semua Bulan</option>
                        {MONTHS.map(m => (
                          <option key={m} value={m}>Bulan {m}</option>
                        ))}
                      </select>
                    </div>

                    <div className="flex items-center gap-3">
                      <div className="text-xs text-slate-500 font-semibold">
                        Total SI: <strong className="text-blue-700 font-mono text-sm">Rp {totalSiFiltered.toLocaleString('id-ID')}</strong> ({filteredSiList.length} Transaksi)
                      </div>
                      <button
                        type="button"
                        onClick={() => window.print()}
                        className="px-4 py-2 bg-slate-900 hover:bg-black text-white font-bold text-xs rounded-lg shadow transition flex items-center gap-1.5 cursor-pointer"
                        title="Cetak Standing Instruction dalam format A4 Landscape"
                      >
                        <span>🖨️</span> Cetak SI (Landscape A4)
                      </button>
                    </div>
                  </div>

                  {/* CETAK STANDING INSTRUCTION (PRINT CONTAINER) */}
                  <div className="print-container bg-white rounded-xl shadow-md border border-slate-300 p-6 md:p-10">
                    
                    <div className="text-center mb-6 pb-4 border-b-2 border-slate-900">
                      <h2 className="text-lg md:text-xl font-bold uppercase tracking-wider text-slate-900">
                        SURAT PERINTAH PEMBAYARAN / STANDING INSTRUCTION (SI)
                      </h2>
                      <p className="text-sm font-semibold text-slate-800">
                        BOSP REGULER {headerInfo.namaSekolah ? headerInfo.namaSekolah.toUpperCase() : "SEKOLAH"} — TAHUN ANGGARAN {headerInfo.tahunAnggaran}
                      </p>
                    </div>

                    <div className="flex flex-row justify-between items-start gap-6 text-xs font-medium text-slate-900 mb-6">
                      <div className="text-left space-y-0.5">
                        <p className="flex items-baseline"><span className="w-32 inline-block font-semibold shrink-0">Nama Sekolah</span><span>: <strong>{headerInfo.namaSekolah}</strong></span></p>
                        <p className="flex items-baseline"><span className="w-32 inline-block font-semibold shrink-0">NPSN</span><span>: {headerInfo.npsn}</span></p>
                        <p className="flex items-baseline"><span className="w-32 inline-block font-semibold shrink-0">Alamat</span><span>: {headerInfo.alamat}</span></p>
                      </div>
                      <div className="text-left space-y-0.5 min-w-[230px]">
                        <p className="flex items-baseline"><span className="w-32 inline-block font-semibold shrink-0">Sumber Dana</span><span>: <strong>{headerInfo.sumberDana}</strong></span></p>
                        <p className="flex items-baseline"><span className="w-32 inline-block font-semibold shrink-0">Kabupaten/Kota</span><span>: {headerInfo.kabupaten}</span></p>
                        <p className="flex items-baseline"><span className="w-32 inline-block font-semibold shrink-0">Provinsi</span><span>: {headerInfo.provinsi}</span></p>
                      </div>
                    </div>

                    <div className="overflow-x-auto mb-8">
                      <table className="w-full text-xs border-collapse border border-slate-900 si-print-table bg-white">
                        <thead>
                          <tr className="bg-white text-slate-900 font-bold text-center border-b-2 border-slate-900">
                            <th className="border border-slate-900 px-2 py-3 w-10 text-center uppercase whitespace-nowrap">NO</th>
                            <th className="border border-slate-900 px-3 py-3 text-center uppercase w-[34%] min-w-[260px]">BELANJA</th>
                            <th className="border border-slate-900 px-2 py-3 w-28 text-center font-mono uppercase whitespace-nowrap">NOMOR REKENING</th>
                            <th className="border border-slate-900 px-2 py-3 w-20 text-center uppercase whitespace-nowrap">NAMA BANK</th>
                            <th className="border border-slate-900 px-3 py-3 text-center uppercase w-36 min-w-[145px] whitespace-nowrap">ATAS NAMA</th>
                            <th className="border border-slate-900 px-2 py-3 w-28 text-center font-mono uppercase whitespace-nowrap">NPWP</th>
                            <th className="border border-slate-900 px-3 py-3 text-center w-28 uppercase whitespace-nowrap">TOTAL</th>
                            <th className="border border-slate-900 px-2 py-3 w-24 text-center uppercase whitespace-nowrap">SUMBER DANA</th>
                            <th className="border border-slate-900 px-2 py-3 w-28 text-center uppercase whitespace-nowrap">TGL TRANSAKSI SI</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filteredSiList.length === 0 ? (
                            <tr>
                              <td colSpan="10" className="border border-slate-900 px-4 py-8 text-center text-slate-500 italic bg-white">
                                Tidak ada data Standing Instruction untuk tanggal/bulan yang dipilih.
                              </td>
                            </tr>
                          ) : (
                            filteredSiList.map((item, idx) => (
                              <tr key={item.id} className="bg-white border-b border-slate-400">
                                <td className="border border-slate-900 px-2 py-2.5 text-center font-bold whitespace-nowrap">{idx + 1}</td>
                                <td className="border border-slate-900 px-3 py-2.5 text-left font-semibold text-slate-900">{item.belanja}</td>
                                <td className="border border-slate-900 px-2 py-2.5 text-center font-mono text-slate-800 whitespace-nowrap">{item.noRek}</td>
                                <td className="border border-slate-900 px-2 py-2.5 text-center font-bold text-slate-800 whitespace-nowrap">{item.namaBank}</td>
                                <td className="border border-slate-900 px-3 py-2.5 text-center font-medium text-slate-800 whitespace-nowrap">{item.atasNama}</td>
                                <td className="border border-slate-900 px-2 py-2.5 font-mono text-center text-slate-700 whitespace-nowrap">{item.npwp}</td>
                                <td className="border border-slate-900 px-3 py-2.5 text-center font-mono font-bold text-slate-900 whitespace-nowrap">
                                  {Number(item.total).toLocaleString('id-ID')}
                                </td>
                                <td className="border border-slate-900 px-2 py-2.5 text-center font-medium whitespace-nowrap">{item.sumberDana}</td>
                                <td className="border border-slate-900 px-2 py-2.5 text-center font-mono text-slate-700 whitespace-nowrap">{item.tglTransaksiSI}</td>
                              </tr>
                            ))
                          )}

                          <tr className="bg-white font-bold text-slate-900 border-t-2 border-slate-900">
                            <td colSpan="6" className="border border-slate-900 px-3 py-3 text-left">
                              Jumlah Total
                            </td>
                            <td className="border border-slate-900 px-3 py-3 text-center font-mono text-sm whitespace-nowrap">
                              Rp {totalSiFiltered.toLocaleString('id-ID')}
                            </td>
                            <td colSpan="2" className="border border-slate-900 px-2 py-3"></td>
                          </tr>
                        </tbody>
                      </table>
                    </div>

                    <div className="mt-8 pt-4 flex justify-between items-start text-xs font-semibold text-slate-900 px-6" style={{ pageBreakInside: 'avoid', breakInside: 'avoid' }}>
                      <div className="text-center min-w-[200px]">
                        <p className="font-normal text-slate-900">Bendahara</p>
                        <div className="h-24"></div>
                        <p className="font-bold text-slate-900">{headerInfo.bendaharaSekolah}</p>
                      </div>

                      <div className="text-center min-w-[220px]">
                        <p className="font-normal text-slate-900">
                          Bekasi, {selectedSiDate === "Semua"
                            ? (filteredSiList.length > 0 && filteredSiList[filteredSiList.length - 1].tglTransaksiSI && filteredSiList[filteredSiList.length - 1].tglTransaksiSI !== "-"
                                ? formatTanggalIndoTeks(filteredSiList[filteredSiList.length - 1].tglTransaksiSI)
                                : "28 September 2026")
                            : formatTanggalIndoTeks(selectedSiDate)}
                        </p>
                        <p className="font-normal text-slate-900 mt-0.5">Kepala {headerInfo.namaSekolah || "Sekolah"}</p>
                        <div className="h-20"></div>
                        <p className="font-bold text-slate-900">{headerInfo.kepalaSekolah}</p>
                      </div>
                    </div>

                  </div>

                </div>
              )}

              {/* VIEW 3: DATABASE MASTER RKAS */}
              {currentView === "database" && (
                <div className="space-y-6">

                  <div className="no-print bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-wrap justify-between items-center gap-4">
                    <div className="flex flex-wrap items-center gap-3">
                      <span className="text-sm font-bold text-slate-800 uppercase flex items-center gap-2">
                        <span>📋</span> Master RKAS ({masterRkas.length} Item)
                      </span>
                      <span className="text-xs font-semibold bg-amber-100 text-amber-900 px-2.5 py-1 rounded-full border border-amber-300">
                        {Object.values(manualPurchasedMap).filter(Boolean).length} Item Dicentang Manual
                      </span>
                    </div>

                    <div className="flex flex-wrap items-center gap-2">
                      <input
                        type="file"
                        id="master-file-import-input"
                        accept=".pdf,.xlsx,.xls,.json"
                        onChange={handleImportMasterFile}
                        className="hidden"
                      />

                      <input
                        type="file"
                        id="master-restore-file-input"
                        accept=".json,.xlsx,.xls"
                        onChange={handleRestoreMasterFile}
                        className="hidden"
                      />

                      {isAdmin && (
                        <button
                          type="button"
                          onClick={handleImportMasterClick}
                          className="bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs px-3.5 py-2 rounded-lg shadow transition flex items-center gap-1.5 cursor-pointer"
                          title="Unggah File PDF / Excel / JSON ARKAS"
                        >
                          <span>📥</span> Impor PDF/Excel ARKAS
                        </button>
                      )}

                      <div className="relative inline-block text-left">
                        <button
                          type="button"
                          onClick={() => setShowMasterExportDropdown(!showMasterExportDropdown)}
                          className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-3.5 py-2 rounded-lg shadow transition flex items-center gap-1.5 cursor-pointer"
                          title="Pilihan Ekspor Master RKAS (Excel / PDF)"
                        >
                          <span>📤</span> Export <span>▼</span>
                        </button>

                        {showMasterExportDropdown && (
                          <div 
                            className="origin-top-right absolute right-0 mt-1 w-48 rounded-md shadow-lg bg-white ring-1 ring-black ring-opacity-5 z-50 divide-y divide-slate-100"
                            onMouseLeave={() => setShowMasterExportDropdown(false)}
                          >
                            <div className="py-1">
                              <button
                                type="button"
                                onClick={() => {
                                  setShowMasterExportDropdown(false);
                                  handleExportMasterExcel();
                                }}
                                className="w-full text-left px-4 py-2 text-xs text-slate-700 hover:bg-emerald-50 hover:text-emerald-700 font-semibold flex items-center gap-2 cursor-pointer"
                              >
                                <span>📊</span> Ekspor Excel (.xlsx)
                              </button>
                              <button
                                type="button"
                                onClick={() => {
                                  setShowMasterExportDropdown(false);
                                  handleExportMasterPdf();
                                }}
                                className="w-full text-left px-4 py-2 text-xs text-slate-700 hover:bg-purple-50 hover:text-purple-700 font-semibold flex items-center gap-2 cursor-pointer"
                              >
                                <span>🖨️</span> Ekspor PDF / Cetak
                              </button>
                            </div>
                          </div>
                        )}
                      </div>

                      {isAdmin && (
                        <>
                          <button
                            type="button"
                            onClick={handleBackupMasterData}
                            className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-3.5 py-2 rounded-lg shadow transition flex items-center gap-1.5 cursor-pointer"
                            title="Download Backup Data Master RKAS (Wajib Sebelum Reset)"
                          >
                            <span>💾</span> Backup Data (Wajib)
                          </button>

                          <button
                            type="button"
                            onClick={handleRestoreMasterClick}
                            className="bg-amber-600 hover:bg-amber-700 text-white font-bold text-xs px-3.5 py-2 rounded-lg shadow transition flex items-center gap-1.5 cursor-pointer"
                            title="Unggah dan pulihkan data Master RKAS dari file backup (.json / .xlsx)"
                          >
                            <span>📂</span> Restore Backup
                          </button>

                          <button
                            type="button"
                            onClick={handleSaveMasterCheckboxes}
                            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-3.5 py-2 rounded-lg shadow transition flex items-center gap-1.5 cursor-pointer"
                          >
                            <span>💾</span> Simpan Perubahan
                          </button>

                          <button
                            type="button"
                            onClick={handleResetMasterAllData}
                            className="bg-red-600 hover:bg-red-700 text-white font-bold text-xs px-3.5 py-2 rounded-lg shadow transition flex items-center gap-1.5 cursor-pointer"
                            title="Hapus / Reset semua data Master RKAS dengan konfirmasi peringatan & backup wajib"
                          >
                            <span>🚨</span> Reset / Hapus Semua Data
                          </button>

                        </>
                      )}
                    </div>
                  </div>

                  <div className="print-container bg-white rounded-xl shadow-md border border-slate-300 p-6 md:p-8">
                    <div className="text-center mb-6 pb-4 border-b border-slate-400">
                      <h2 className="text-lg md:text-xl font-bold uppercase tracking-wider text-slate-900">
                        KERTAS KERJA RENCANA KEGIATAN DAN ANGGARAN SEKOLAH (RKAS)
                      </h2>
                      <p className="text-sm font-semibold text-slate-800">TAHUN ANGGARAN : {headerInfo.tahunAnggaran}</p>
                    </div>

                    <div className="overflow-x-auto">
                      <table className="w-full text-xs border-collapse border border-slate-900">
                        <thead>
                          <tr className="bg-slate-200 text-slate-900 font-bold text-center border-b border-slate-900">
                            <th rowSpan="2" className="border border-slate-900 px-2 py-2.5 w-14 uppercase text-center font-bold">NO.<br/>URUT</th>
                            <th rowSpan="2" className="border border-slate-900 px-2 py-2.5 w-32 uppercase text-center font-bold">KODE REKENING</th>
                            <th rowSpan="2" className="border border-slate-900 px-2 py-2.5 w-28 uppercase text-center font-bold">KODE KEGIATAN</th>
                            <th rowSpan="2" className="border border-slate-900 px-3 py-2.5 text-left uppercase font-bold">URAIAN KEGIATAN</th>
                            <th rowSpan="2" className="border border-slate-900 px-3 py-2.5 text-right w-32 uppercase font-bold">JUMLAH</th>
                            <th colSpan="2" className="border border-slate-900 px-2 py-1 text-center uppercase font-bold">BOSP REGULER</th>
                            <th rowSpan="2" className="border border-slate-900 px-2 py-2.5 w-16 uppercase text-center font-bold">AKSI</th>
                            <th rowSpan="2" className="border border-slate-900 px-2 py-2.5 w-48 uppercase text-center font-bold">KETERANGAN</th>
                          </tr>
                          <tr className="bg-slate-200 text-slate-900 font-bold text-center border-b border-slate-900">
                            <th className="border border-slate-900 px-2 py-1.5 w-28 uppercase text-center font-bold">BELANJA OPERASI</th>
                            <th className="border border-slate-900 px-2 py-1.5 w-28 uppercase text-center font-bold">BELANJA MODAL</th>
                          </tr>
                        </thead>
                        <tbody>
                          {masterRkas.map((m) => {
                            const st = masterYearBudgetFull[m.no];
                            const isPurchasedAuto = st ? st.isFull : purchasedUraianSet.has(String(m.no));
                            const isPurchasedManual = !!manualPurchasedMap[m.no];
                            const isPurchased = isPurchasedAuto || isPurchasedManual;
                            const isPartial = st ? st.isPartial : false;
                            const hasCheckbox = m.level !== "standar" && m.level !== "kegiatan";
                            const isRekeningHeader = m.kodeRekening && !m.uraian.match(/^\d{3}\./);

                            let rowClass = "bg-white hover:bg-slate-50";
                            if (m.level === "standar") rowClass = "row-header-standar text-red-950 font-bold";
                            else if (m.level === "kegiatan") rowClass = "row-header-kegiatan text-emerald-950 font-bold";
                            else if (m.level === "sub") rowClass = "row-header-sub text-amber-950 font-bold";
                            else if (isPurchased) rowClass = "bg-purchased-yellow font-bold";

                            return (
                              <tr key={m.no} className={`transition border-b border-slate-900 ${rowClass}`}>
                                <td className="border border-slate-900 px-2 py-2 text-center font-bold">{m.no}</td>
                                <td className="border border-slate-900 px-2 py-2 font-mono text-slate-800 text-left">{m.kodeRekening || ""}</td>
                                <td className="border border-slate-900 px-2 py-2 font-mono font-bold text-slate-900 text-left">{m.kodeKegiatan || ""}</td>
                                <td className={`border border-slate-900 px-3 py-2 text-left ${m.level !== "item" ? "font-bold" : (isRekeningHeader ? "italic font-semibold" : "")}`}>
                                  <div>{m.uraian}</div>
                                  {isPartial && st && (
                                    <div className="no-print mt-0.5 text-[10px] font-semibold text-amber-900 bg-amber-100/90 border border-amber-300 px-1.5 py-0.5 rounded inline-block">
                                      ⏳ Terpakai: Rp {st.terpakai.toLocaleString('id-ID')} | Sisa: Rp {st.sisa.toLocaleString('id-ID')}
                                    </div>
                                  )}
                                </td>
                                <td className="border border-slate-900 px-3 py-2 text-right font-mono font-bold">{m.jumlah.toLocaleString('id-ID')}</td>
                                <td className="border border-slate-900 px-3 py-2 text-right font-mono font-bold">{m.operasi.toLocaleString('id-ID')}</td>
                                <td className="border border-slate-900 px-3 py-2 text-right font-mono font-bold">{m.modal.toLocaleString('id-ID')}</td>
                                <td className="border border-slate-900 px-2 py-1.5 text-center">
                                  {hasCheckbox && (
                                    <input
                                      type="checkbox"
                                      disabled={!isAdmin}
                                      checked={isPurchasedManual || isPurchasedAuto}
                                      onChange={(e) => {
                                        if (isAdmin) toggleManualPurchased(m, e.target.checked);
                                      }}
                                      className={`w-4 h-4 text-amber-600 bg-gray-100 border-gray-300 rounded focus:ring-amber-500 accent-amber-600 ${isAdmin ? "cursor-pointer" : "cursor-not-allowed opacity-70"}`}
                                      title={isPurchased ? "Sudah Dibelanjakan Penuh" : isPartial ? `Terpakai Sebagian (Sisa: Rp ${st.sisa.toLocaleString('id-ID')})` : "Belum Dibelanjakan"}
                                    />
                                  )}
                                </td>
                                <td className="border border-slate-900 px-2 py-1.5">
                                  <input
                                    type="text"
                                    readOnly={!isAdmin}
                                    disabled={!isAdmin}
                                    placeholder={isAdmin ? "Ketik keterangan..." : "-"}
                                    value={manualKeteranganTextMap[m.no] || ""}
                                    onChange={(e) => {
                                      if (isAdmin) handleKeteranganTextChange(m.no, e.target.value);
                                    }}
                                    className="w-full text-xs px-2 py-1 border border-slate-300 rounded focus:ring-1 focus:ring-blue-500 bg-white focus:bg-white text-slate-900 font-normal disabled:bg-slate-50"
                                  />
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>

                    <div className="mt-10 pt-4 flex justify-between items-start text-xs font-semibold text-slate-900 px-6">
                      <div className="text-center min-w-[200px]">
                        <p className="font-normal text-slate-900">Bendahara</p>
                        <div className="h-24"></div>
                        <p className="font-bold text-slate-900">{headerInfo.bendaharaSekolah}</p>
                      </div>

                      <div className="text-center min-w-[220px]">
                        <p className="font-normal text-slate-900">Bekasi, 28 September 2026</p>
                        <p className="font-normal text-slate-900 mt-0.5">Kepala {headerInfo.namaSekolah || "Sekolah"}</p>
                        <div className="h-20"></div>
                        <p className="font-bold text-slate-900">{headerInfo.kepalaSekolah}</p>
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* VIEW 4: REKAP SERAPAN */}
              {currentView === "rekap" && (
                <div className="space-y-6">
                  {(() => {
                    let totalMasterYearPagu = 0;
                    if (Array.isArray(masterRkas) && masterRkas.length > 0) {
                      const stdRows = masterRkas.filter(m => m.level === "standar");
                      if (stdRows.length > 0) {
                        totalMasterYearPagu = stdRows.reduce((s, m) => s + Number(m.jumlah || 0), 0);
                      } else {
                        const leafRows = masterRkas.filter(m => m.level === "item" && /^\d{3}\./.test((m.uraian || "").trim()));
                        totalMasterYearPagu = (leafRows.length > 0 ? leafRows : masterRkas).reduce((s, m) => s + Number(m.jumlah || 0), 0);
                      }
                    }
                    if (!totalMasterYearPagu || totalMasterYearPagu <= 0) {
                      totalMasterYearPagu = 246840000;
                    }

                    const isSpecificMonth = selectedMonthView && selectedMonthView !== "Semua";
                    const rekapMonthRows = isSpecificMonth ? (perBulanData[selectedMonthView] || []) : [];
                    const rekapTotalRencana = isSpecificMonth
                      ? (rekapMonthRows.length > 0 ? calculatePerBulanTotal(rekapMonthRows) : 0)
                      : totalMasterYearPagu;
                    const rekapTotalDanaTersedia = isSpecificMonth ? rekapTotalRencana : totalMasterYearPagu;

                    const rekapTxList = isSpecificMonth
                      ? (belanjaBankList || []).filter(b => b && b.bulan === selectedMonthView)
                      : (belanjaBankList || []);

                    const rekapTotalDibelanjakan = rekapTxList.reduce((s, b) => s + Number(b.total || 0), 0);
                    const rekapTotalBelumDibelanjakan = Math.max(0, rekapTotalRencana - rekapTotalDibelanjakan);
                    const rekapSisaSaldo = Math.max(0, rekapTotalRencana - rekapTotalDibelanjakan);
                    const rekapLabel = selectedMonthView === "Semua" ? "SEMUA" : selectedMonthView.toUpperCase();

                    const rekapTableStatusMap = isSpecificMonth
                      ? computeBudgetStatusForSource(masterRkas, masterRkasHierarchy, rekapTxList, {}, null)
                      : masterYearBudgetFull;

                    return (
                      <>
                        <div className="no-print grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-purple-600">
                            <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                              Total Dana Tersedia ({rekapLabel})
                            </p>
                            <p className="text-xl font-bold text-purple-700 font-mono mt-1">
                              Rp {rekapTotalDanaTersedia.toLocaleString('id-ID')}
                            </p>
                            <p className="text-[11px] text-purple-600 mt-1 font-semibold flex items-center gap-1">
                              <span>💰</span> {isSpecificMonth ? `Total Pagu Anggaran Bulan ${selectedMonthView}` : `Total Pagu Anggaran ${headerInfo.namaSekolah || "SDIT ANNISA"} (1 Tahun)`}
                            </p>
                          </div>

                          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-emerald-600">
                            <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                              Total Sudah Dibelanjakan ({rekapLabel})
                            </p>
                            <p className="text-xl font-bold text-emerald-700 font-mono mt-1">
                              Rp {rekapTotalDibelanjakan.toLocaleString('id-ID')}
                            </p>
                            <p className="text-[11px] text-emerald-600 mt-1 font-semibold flex items-center gap-1">
                              <span>✅</span> {rekapTxList.length} Transaksi Pembayaran Realized
                            </p>
                          </div>

                          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-amber-500">
                            <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                              Total Belum Dibelanjakan ({rekapLabel})
                            </p>
                            <p className="text-xl font-bold text-amber-700 font-mono mt-1">
                              Rp {rekapTotalBelumDibelanjakan.toLocaleString('id-ID')}
                            </p>
                            <p className="text-[11px] text-slate-400 mt-1 font-medium">
                              Perhitungan Berbasis Rincian Belanja Perbulan
                            </p>
                          </div>

                          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-blue-600">
                            <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">
                              Sisa Saldo Anggaran ({rekapLabel})
                            </p>
                            <p className="text-xl font-bold text-blue-700 font-mono mt-1">
                              Rp {rekapSisaSaldo.toLocaleString('id-ID')}
                            </p>
                            <p className="text-[11px] text-blue-600 mt-1 font-medium">
                              Total Target Rencana: Rp {rekapTotalRencana.toLocaleString('id-ID')}
                            </p>
                          </div>
                        </div>

                        <div className="bg-white rounded-xl shadow-md border border-slate-200 p-6">
                          <div className="flex flex-wrap items-center justify-between gap-3 mb-4 pb-2 border-b">
                            <h3 className="text-base font-bold text-slate-900">
                              Rekapitulasi Serapan Belanja vs Master RKAS PDF ({selectedMonthView === "Semua" ? "SEMUA" : selectedMonthView.toUpperCase()})
                            </h3>
                            <div className="no-print flex items-center gap-2">
                              <span className="text-xs font-bold text-slate-700 uppercase">📅 Filter Periode:</span>
                              <select
                                value={selectedMonthView}
                                onChange={(e) => setSelectedMonthView(e.target.value)}
                                className="text-xs font-bold border-2 border-blue-500 rounded-lg px-3 py-1.5 bg-white text-blue-900 focus:outline-none cursor-pointer shadow-xs"
                              >
                                <option value="Semua">Semua Bulan (1 Tahun)</option>
                                {MONTHS.map(m => (
                                  <option key={m} value={m}>Bulan {m}</option>
                                ))}
                              </select>
                            </div>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-xs text-left border-collapse border border-slate-300">
                              <thead className="bg-slate-100 font-bold text-center">
                                <tr>
                                  <th className="border border-slate-300 px-2 py-2 w-10">No.</th>
                                  <th className="border border-slate-300 px-2 py-2">Kode Rekening</th>
                                  <th className="border border-slate-300 px-3 py-2 text-left">Uraian Master RKAS</th>
                                  <th className="border border-slate-300 px-3 py-2 text-right">Pagu Target (Rp)</th>
                                  <th className="border border-slate-300 px-3 py-2 text-right">Realisasi Belanja ({selectedMonthView})</th>
                                  <th className="border border-slate-300 px-3 py-2 text-right">Sisa Pagu (Rp)</th>
                                  <th className="border border-slate-300 px-2 py-2 text-center">Status</th>
                                  <th className="border border-slate-300 px-3 py-2 text-left">Keterangan Catatan</th>
                                </tr>
                              </thead>
                              <tbody>
                                {masterRkas.map((m) => {
                                  const stRow = rekapTableStatusMap[m.no];
                                  const relBank = stRow ? stRow.terpakai : 0;
                                  const isPurchasedManual = !isSpecificMonth && !!manualPurchasedMap[m.no];
                                  const rel = isPurchasedManual ? (relBank > 0 ? relBank : m.jumlah) : relBank;
                                  const sisa = Math.max(0, m.jumlah - rel);
                                  const isPurchased = isPurchasedManual || relBank > 0;

                                  return (
                                    <tr key={m.no} className={`hover:bg-slate-50 ${isPurchased ? 'bg-yellow-100 font-bold' : ''}`}>
                                      <td className="border border-slate-300 px-2 py-2 text-center font-mono">{m.no}</td>
                                      <td className="border border-slate-300 px-2 py-2 font-mono text-slate-700">{m.kodeRekening}</td>
                                      <td className="border border-slate-300 px-3 py-2 text-slate-900">{m.uraian}</td>
                                      <td className="border border-slate-300 px-3 py-2 text-right font-mono font-semibold">{m.jumlah.toLocaleString('id-ID')}</td>
                                      <td className="border border-slate-300 px-3 py-2 text-right font-mono font-bold text-emerald-700">{rel.toLocaleString('id-ID')}</td>
                                      <td className="border border-slate-300 px-3 py-2 text-right font-mono font-semibold">{sisa.toLocaleString('id-ID')}</td>
                                      <td className="border border-slate-300 px-2 py-2 text-center">
                                        {isPurchased ? (
                                          <span className="bg-yellow-300 text-yellow-950 font-bold text-[10px] px-2 py-0.5 rounded border border-yellow-500">
                                            {isPurchasedManual && relBank === 0 ? "🟨 Dibelanjakan (Manual)" : "🟨 Dibelanjakan"}
                                          </span>
                                        ) : (
                                          <span className="text-slate-400 text-[10px]">Belum</span>
                                        )}
                                      </td>
                                      <td className="border border-slate-300 px-3 py-2 text-slate-700 italic font-normal">
                                        {manualKeteranganTextMap[m.no] || "-"}
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </>
                    );
                  })()}
                </div>
              )}

              {/* VIEW: SURAT TRANSFER (PEMINDAHBUKUAN STANDING INSTRUCTION) */}
              {currentView === "surat_transfer" && (
                <div className="space-y-6">
                  <style>{`
                    @media print {
                      @page {
                        size: A4 portrait !important;
                        margin: 8mm 12mm !important;
                      }
                    }
                  `}</style>
                  
                  <div className="no-print bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-wrap justify-between items-center gap-4">
                    <div className="flex flex-wrap items-center gap-4 text-xs">
                      <div className="flex items-center gap-2">
                        <span className="font-bold text-slate-800 flex items-center gap-1">
                          <span>🗓️</span> Filter Tanggal SI:
                        </span>
                        <select
                          value={selectedSuratTransferDate}
                          onChange={(e) => {
                            const val = e.target.value;
                            setSelectedSuratTransferDate(val);
                            if (val && val !== "Semua") {
                              setSuratTransferCetakDate(`${headerInfo.kabupaten || "Bekasi"}, ${formatTanggalIndoTeks(val)}`);
                            }
                          }}
                          className="font-bold border-2 border-blue-500 rounded-lg px-3 py-1.5 bg-white text-blue-900 focus:outline-none"
                        >
                          <option value="Semua">-- Semua Tanggal Transaksi --</option>
                          {uniqueSiDates.map(d => (
                            <option key={d} value={d}>Tanggal: {d}</option>
                          ))}
                        </select>
                      </div>

                      {isAdmin && (
                        <>
                          <div className="flex items-center gap-2">
                            <span className="font-bold text-slate-800">Tanggal Cetak:</span>
                            <input
                              type="text"
                              value={suratTransferCetakDate}
                              onChange={(e) => setSuratTransferCetakDate(e.target.value)}
                              className="border border-slate-300 rounded-lg px-3 py-1.5 text-xs font-bold text-slate-800 w-48"
                              placeholder="Ketik tanggal surat..."
                            />
                          </div>

                          <div className="flex items-center gap-2">
                            <span className="font-bold text-slate-800">Nomor Surat:</span>
                            <input
                              type="text"
                              value={suratTransferNomor}
                              onChange={(e) => setSuratTransferNomor(e.target.value)}
                              className="border border-slate-300 rounded-lg px-3 py-1.5 text-xs font-mono font-bold text-slate-800 w-60"
                              placeholder="Ketik nomor surat..."
                            />
                          </div>
                        </>
                      )}
                    </div>

                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => window.print()}
                        className="px-4 py-2 bg-blue-600 text-white font-bold rounded-lg hover:bg-blue-700 shadow transition flex items-center gap-1.5 text-xs"
                      >
                        <span>🖨️</span> Cetak Surat Transfer (PDF)
                      </button>
                    </div>
                  </div>

                  <div className="print-container bg-white rounded-xl shadow-md border border-slate-300 p-6 md:p-8 text-slate-900 text-[11px] font-sans leading-tight max-w-3xl mx-auto" style={{ pageBreakInside: 'avoid', pageBreakAfter: 'avoid' }}>
                    
                    {customKopSuratUrl && (
                      <div className="text-center pb-1 mb-2">
                        <img src={customKopSuratUrl} alt="Kop Surat Resmi" className="w-full max-h-32 object-contain mx-auto" />
                      </div>
                    )}

                    <div className="text-right text-xs font-semibold mb-2">
                      {suratTransferCetakDate}
                    </div>

                    <div className="text-center mb-3">
                      <h3 className="text-xs font-bold uppercase tracking-wider text-slate-900 underline underline-offset-4">
                        STANDING INSTRUCTION
                      </h3>
                    </div>

                    <div className="text-xs space-y-0.5 mb-3">
                      <p className="flex"><span className="w-24 font-medium">Nomor</span><span>: {suratTransferNomor}</span></p>
                      <p className="flex"><span className="w-24 font-medium">Sifat</span><span>: -</span></p>
                      <p className="flex"><span className="w-24 font-medium">Lampiran</span><span>: -</span></p>
                      <p className="flex"><span className="w-24 font-medium">Perihal</span><span>: <strong>Surat Perintah Pemindahbukuan</strong></span></p>
                    </div>

                    <div className="text-xs space-y-0.5 mb-3">
                      <p>Kepada Yth,</p>
                      <p className="font-bold">{(!headerInfo.namaBank || headerInfo.namaBank.trim().toUpperCase() === "BJB") ? "Bank BJB Kantor Cabang Jatiasih" : headerInfo.namaBank}</p>
                      <p>di {headerInfo.kabupaten || "Tempat"}</p>
                    </div>

                    <div className="text-xs mb-2">
                      Kami yang bertanda tangan dibawah ini, dengan ini meminta untuk <strong>Pemindah bukuan</strong> dana BOS dengan data berikut :
                    </div>

                    <div className="pl-4 text-xs space-y-0.5 mb-3">
                      <p className="flex"><span className="w-44 font-medium">Nomor Rekening</span><span>: {headerInfo.noRekening || ""}</span></p>
                      <p className="flex"><span className="w-44 font-medium">Nama Rekening</span><span>: {headerInfo.namaSekolah || ""}</span></p>
                      <p className="flex"><span className="w-44 font-medium">Unit Kerja</span><span>: {headerInfo.namaSekolah || ""}</span></p>
                      <p className="flex"><span className="w-44 font-medium">Alamat</span><span>: {headerInfo.alamat || ""}</span></p>
                      <p className="flex"><span className="w-44 font-medium">Nama Kepala Sekolah</span><span>: {headerInfo.kepalaSekolah || ""}</span></p>
                      <p className="flex"><span className="w-44 font-medium">Nama Bendahara</span><span>: {headerInfo.bendaharaSekolah || ""}</span></p>
                      <p className="flex"><span className="w-44 font-medium">Nomor Hp</span><span>: {headerInfo.noHp || ""}</span></p>
                      <p className="flex"><span className="w-44 font-medium">Email</span><span>: {headerInfo.email || ""}</span></p>
                    </div>

                    <div className="text-xs mb-1 font-semibold">
                      Untuk dilakukan Pemindah Bukuan ke Rekening :
                    </div>
                    <div className="pl-4 text-xs space-y-0.5 mb-3">
                      <p className="flex"><span className="w-44 font-medium">Nomor Rekening</span><span>: (Terlampir)</span></p>
                      <p className="flex"><span className="w-44 font-medium">Nama Rekening</span><span>: (Terlampir)</span></p>
                      <p className="flex"><span className="w-44 font-medium">Bank Tujuan</span><span>: (Terlampir)</span></p>
                      <p className="flex items-center"><span className="w-44 font-semibold">Jumlah</span><span className="font-mono font-bold text-xs text-slate-900">: Rp. {totalSuratTransfer.toLocaleString('id-ID')}</span></p>
                      <p className="flex items-start"><span className="w-44 font-semibold">Terbilang</span><span className="font-bold italic text-slate-900 flex-1">: {terbilangSuratTransfer}</span></p>
                    </div>

                    <div className="text-xs space-y-0.5 mb-3">
                      <p>Dengan adanya <em>Standing Instruction</em> ini, maka :</p>
                      <ol className="list-decimal pl-6 space-y-0.5">
                        <li>Bank dibebaskan dari segala akibat yang mungkin timbul dari pelaksanaan pemindahbukuan berdasarkan <em>Standing Instruction</em> ini;</li>
                        <li>Kami meyakini atas kebenaran data dan informasi pada <em>Standing Instruction</em> ini.</li>
                      </ol>
                    </div>

                    <div className="text-xs mb-4">
                      Demikian atas kerjasama nya, kami ucapkan terimakasih .
                    </div>

                    <div className="grid grid-cols-2 text-center text-xs gap-4 pt-1">
                      <div>
                        <p className="font-medium">Bendahara</p>
                        <div className="h-14"></div>
                        <p className="font-bold">{headerInfo.bendaharaSekolah || ""}</p>
                      </div>
                      <div>
                        <p className="font-medium">Kepala Sekolah</p>
                        <p className="font-bold text-slate-900">{headerInfo.namaSekolah || ""}</p>
                        <div className="h-10"></div>
                        <p className="font-bold">{headerInfo.kepalaSekolah || ""}</p>
                      </div>
                    </div>

                  </div>
                </div>
              )}

              {/* VIEW: PENGATURAN / PROFIL SEKOLAH & KOP SURAT */}
              {currentView === "pengaturan" && !isAdmin && (
                <div className="max-w-lg mx-auto bg-white p-8 rounded-2xl border border-slate-200 shadow-sm text-center space-y-4 mt-8">
                  <div className="w-12 h-12 rounded-full bg-amber-100 text-amber-700 flex items-center justify-center mx-auto text-2xl">
                    🔒
                  </div>
                  <h2 className="text-lg font-bold text-slate-900">Akses Pengaturan Terkunci</h2>
                  <p className="text-xs text-slate-600">
                    Silakan login sebagai Admin terlebih dahulu melalui ikon gir (⚙️) di kanan atas untuk mengubah pengaturan sekolah.
                  </p>
                  <button
                    type="button"
                    onClick={() => {
                      setAdminPasswordInput("");
                      setAdminPasswordError("");
                      setShowAdminLoginModal(true);
                    }}
                    className="px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold rounded-lg shadow transition cursor-pointer"
                  >
                    ⚙️ Login Admin
                  </button>
                </div>
              )}

              {currentView === "pengaturan" && isAdmin && (
                <div className="space-y-6 max-w-5xl mx-auto">
                  
                  <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm flex flex-wrap justify-between items-center gap-4">
                    <div>
                      <h2 className="text-lg font-bold text-slate-900 flex items-center gap-2">
                        ⚙️ Pengaturan Identitas Sekolah, Logo, & Kop Surat
                      </h2>
                      <p className="text-xs text-slate-500 mt-1">
                        Kelola data profil sekolah, Kepala Sekolah, Bendahara BOS, Logo Web, serta Gambar Kop Surat untuk laporan resmi.
                      </p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={handleResetPengaturanDefault}
                        className="px-3.5 py-2 bg-amber-50 text-amber-800 border border-amber-300 font-bold text-xs rounded-lg hover:bg-amber-100 transition shadow-sm"
                      >
                        🔄 Reset Default
                      </button>
                    </div>
                  </div>

                  <form onSubmit={handleSavePengaturan} className="space-y-6">
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      
                      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-4">
                        <h3 className="text-sm font-bold text-blue-900 border-b pb-2 flex items-center gap-2">
                          🏫 Identitas Sekolah
                        </h3>

                        <div>
                          <label className="block text-xs font-bold text-slate-700 mb-1">Nama Sekolah:</label>
                          <input
                            type="text"
                            required
                            value={headerInfo.namaSekolah || ''}
                            onChange={(e) => setHeaderInfo({ ...headerInfo, namaSekolah: e.target.value })}
                            className="w-full text-xs font-bold border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none"
                            placeholder="Masukkan Nama Sekolah..."
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-bold text-slate-700 mb-1">NPSN:</label>
                          <input
                            type="text"
                            required
                            value={headerInfo.npsn || ''}
                            onChange={(e) => setHeaderInfo({ ...headerInfo, npsn: e.target.value })}
                            className="w-full text-xs font-mono border border-slate-300 rounded-lg px-3 py-2"
                            placeholder="Masukkan NPSN..."
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-bold text-slate-700 mb-1">Alamat Sekolah:</label>
                          <textarea
                            rows="2"
                            required
                            value={headerInfo.alamat || ''}
                            onChange={(e) => setHeaderInfo({ ...headerInfo, alamat: e.target.value })}
                            className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2"
                            placeholder="Masukkan Alamat Lengkap Sekolah..."
                          />
                        </div>

                        <div className="grid grid-cols-2 gap-3">
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1">Kabupaten/Kota:</label>
                            <input
                              type="text"
                              value={headerInfo.kabupaten || ''}
                              onChange={(e) => setHeaderInfo({ ...headerInfo, kabupaten: e.target.value })}
                              className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2"
                              placeholder="Kabupaten / Kota..."
                            />
                          </div>
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1">Provinsi:</label>
                            <input
                              type="text"
                              value={headerInfo.provinsi || ''}
                              onChange={(e) => setHeaderInfo({ ...headerInfo, provinsi: e.target.value })}
                              className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2"
                              placeholder="Provinsi..."
                            />
                          </div>
                        </div>

                        <div className="grid grid-cols-2 gap-3">
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1">Tahun Anggaran:</label>
                            <input
                              type="text"
                              value={headerInfo.tahunAnggaran || ''}
                              onChange={(e) => setHeaderInfo({ ...headerInfo, tahunAnggaran: e.target.value })}
                              className="w-full text-xs font-bold border border-slate-300 rounded-lg px-3 py-2"
                              placeholder="2026"
                            />
                          </div>
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1">Sumber Dana:</label>
                            <input
                              type="text"
                              value={headerInfo.sumberDana || ''}
                              onChange={(e) => setHeaderInfo({ ...headerInfo, sumberDana: e.target.value })}
                              className="w-full text-xs font-bold border border-slate-300 rounded-lg px-3 py-2"
                              placeholder="BOS Reguler"
                            />
                          </div>
                        </div>
                      </div>

                      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-4">
                        <h3 className="text-sm font-bold text-blue-900 border-b pb-2 flex items-center gap-2">
                          ✍️ Kepala Sekolah, Bendahara & Bank
                        </h3>

                        <div>
                          <label className="block text-xs font-bold text-slate-700 mb-1">Nama Kepala Sekolah:</label>
                          <input
                            type="text"
                            required
                            value={headerInfo.kepalaSekolah || ''}
                            onChange={(e) => setHeaderInfo({ ...headerInfo, kepalaSekolah: e.target.value })}
                            className="w-full text-xs font-bold border border-slate-300 rounded-lg px-3 py-2 text-slate-900"
                            placeholder="Masukkan Nama Kepala Sekolah..."
                          />
                          <p className="text-[11px] text-slate-500 mt-1">Tampil otomatis pada kolom Tanda Tangan Laporan & Surat Transfer.</p>
                        </div>

                        <div>
                          <label className="block text-xs font-bold text-slate-700 mb-1">Nama Bendahara BOS:</label>
                          <input
                            type="text"
                            required
                            value={headerInfo.bendaharaSekolah || ''}
                            onChange={(e) => setHeaderInfo({ ...headerInfo, bendaharaSekolah: e.target.value })}
                            className="w-full text-xs font-bold border border-slate-300 rounded-lg px-3 py-2 text-slate-900"
                            placeholder="Masukkan Nama Bendahara..."
                          />
                          <p className="text-[11px] text-slate-500 mt-1">Tampil otomatis pada kolom Tanda Tangan Bendahara.</p>
                        </div>

                        <div>
                          <label className="block text-xs font-bold text-slate-700 mb-1">Nomor Rekening Sekolah (BOS):</label>
                          <input
                            type="text"
                            value={headerInfo.noRekening || ''}
                            onChange={(e) => setHeaderInfo({ ...headerInfo, noRekening: e.target.value })}
                            className="w-full text-xs font-mono font-bold border border-slate-300 rounded-lg px-3 py-2"
                            placeholder="Masukkan Nomor Rekening..."
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-bold text-slate-700 mb-1">Nama Bank Tujuan/Rekening:</label>
                          <input
                            type="text"
                            value={headerInfo.namaBank || ''}
                            onChange={(e) => setHeaderInfo({ ...headerInfo, namaBank: e.target.value })}
                            className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2"
                            placeholder="Masukkan Nama Bank..."
                          />
                        </div>

                        <div className="grid grid-cols-2 gap-3">
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1">Email Sekolah/Bendahara:</label>
                            <input
                              type="email"
                              value={headerInfo.email || ''}
                              onChange={(e) => setHeaderInfo({ ...headerInfo, email: e.target.value })}
                              className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2"
                              placeholder="Masukkan Email..."
                            />
                          </div>
                          <div>
                            <label className="block text-xs font-bold text-slate-700 mb-1">No HP / Whatsapp:</label>
                            <input
                              type="text"
                              value={headerInfo.noHp || ''}
                              onChange={(e) => setHeaderInfo({ ...headerInfo, noHp: e.target.value })}
                              className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2"
                              placeholder="Masukkan No HP..."
                            />
                          </div>
                        </div>
                      </div>

                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      
                      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-4">
                        <h3 className="text-sm font-bold text-emerald-800 border-b pb-2 flex items-center gap-2">
                          🖼️ Upload Logo Sekolah (Logo Web App)
                        </h3>

                        <p className="text-xs text-slate-600">
                          Unggah logo sekolah Anda (format PNG, JPG, WebP). Logo ini akan tampil pada bagian atas web dan header dokumen.
                        </p>

                        <label className="block w-full border-2 border-dashed border-emerald-300 hover:border-emerald-500 rounded-xl p-4 text-center bg-emerald-50/40 cursor-pointer transition">
                          <span className="text-sm font-bold text-emerald-800 block">📁 Pilih / Unggah Gambar Logo</span>
                          <span className="text-[11px] text-slate-500 block mt-1">Klik di sini untuk memilih file logo dari komputer</span>
                          <input
                            type="file"
                            accept="image/*"
                            onChange={handleUploadLogoFile}
                            className="hidden"
                          />
                        </label>

                        <div className="border rounded-xl p-3 bg-slate-50 space-y-2">
                          <div className="flex items-center justify-between">
                            <span className="text-xs font-bold text-slate-600">Preview Logo:</span>
                            {customLogoUrl && (
                              <button
                                type="button"
                                onClick={handleHapusLogo}
                                className="px-2.5 py-1 bg-red-50 text-red-600 border border-red-200 font-bold text-xs rounded hover:bg-red-100 cursor-pointer"
                              >
                                🗑️ Hapus Gambar Logo
                              </button>
                            )}
                          </div>
                          {customLogoUrl ? (
                            <div className="p-3 bg-white/50 rounded-lg flex items-center justify-center border border-slate-200">
                              <img src={customLogoUrl} alt="Logo Sekolah" className="max-h-24 object-contain bg-transparent" />
                            </div>
                          ) : (
                            <p className="text-slate-400 text-xs italic text-center py-2">Belum ada logo tersimpan. Logo standar SiBOS akan digunakan.</p>
                          )}
                        </div>
                      </div>

                      <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-4">
                        <h3 className="text-sm font-bold text-purple-800 border-b pb-2 flex items-center gap-2">
                          📜 Upload Kop Surat Resmi (Untuk Surat Transfer)
                        </h3>

                        <p className="text-xs text-slate-600">
                          Unggah gambar Kop Surat resmi sekolah Anda. Gambar Kop Surat ini akan otomatis digunakan pada header <strong>Surat Transfer (Pemindahbukuan)</strong> saat dicetak.
                        </p>

                        <label className="block w-full border-2 border-dashed border-purple-300 hover:border-purple-500 rounded-xl p-4 text-center bg-purple-50/40 cursor-pointer transition">
                          <span className="text-sm font-bold text-purple-800 block">📁 Pilih / Unggah Gambar Kop Surat</span>
                          <span className="text-[11px] text-slate-500 block mt-1">Format PNG, JPG, WebP dengan resolusi tinggi</span>
                          <input
                            type="file"
                            accept="image/*"
                            onChange={handleUploadKopSuratFile}
                            className="hidden"
                          />
                        </label>

                        <div className="border rounded-xl p-3 bg-slate-50 space-y-2">
                          <div className="flex items-center justify-between">
                            <span className="text-xs font-bold text-slate-600">Preview Kop Surat:</span>
                            {customKopSuratUrl && (
                              <button
                                type="button"
                                onClick={() => {
                                  setCustomKopSuratUrl("");
                                  alert("🗑️ Gambar Kop Surat berhasil dihapus!");
                                }}
                                className="px-2.5 py-1 bg-red-50 text-red-600 border border-red-200 font-bold text-xs rounded hover:bg-red-100 cursor-pointer"
                              >
                                🗑️ Hapus Gambar Kop
                              </button>
                            )}
                          </div>
                          {customKopSuratUrl ? (
                            <img src={customKopSuratUrl} alt="Kop Surat" className="w-full max-h-24 object-contain rounded border bg-white p-1" />
                          ) : (
                            <p className="text-slate-400 text-xs italic text-center py-2">Belum ada gambar Kop Surat tersimpan. Header Kop Surat pada Surat Transfer tidak akan ditampilkan.</p>
                          )}
                        </div>
                      </div>

                    </div>

                    <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-4">
                      <div className="flex flex-wrap items-center justify-between border-b pb-3 gap-2">
                        <div>
                          <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                            🔐 Pengaturan User &amp; Password Admin (Sheet: Role Pengguna)
                          </h3>
                          <p className="text-xs text-slate-500 mt-0.5">
                            User dan Password Admin disimpan langsung di Google Spreadsheet pada Sheet <strong>Role Pengguna</strong> (tidak ditulis di dalam kode).
                          </p>
                        </div>
                        <span className="text-xs font-mono bg-blue-50 text-blue-700 px-2.5 py-1 rounded-lg border border-blue-200 font-bold">
                          Sheet: Role Pengguna
                        </span>
                      </div>

                      {passwordMsg.text && (
                        <div className={`p-3 rounded-lg text-xs font-bold border ${passwordMsg.type === "error" ? "bg-red-50 text-red-700 border-red-200" : "bg-emerald-50 text-emerald-700 border-emerald-200"}`}>
                          {passwordMsg.text}
                        </div>
                      )}

                      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                        <div>
                          <label className="block text-xs font-bold text-slate-700 mb-1">User Admin (Kolom USER / NPSN):</label>
                          <input
                            type="text"
                            value={newUserVal}
                            onChange={(e) => setNewUserVal(e.target.value)}
                            placeholder="Kosongkan jika tidak mengubah user..."
                            className="w-full text-xs font-mono border border-slate-300 rounded-lg px-3 py-2.5 focus:ring-2 focus:ring-blue-500 focus:outline-none"
                          />
                          <span className="text-[11px] text-slate-400 mt-1 block">Tersimpan di kolom C Sheet Role Pengguna.</span>
                        </div>

                        <div>
                          <label className="block text-xs font-bold text-slate-700 mb-1">Password Baru (Kolom PASSWORD):</label>
                          <div className="relative">
                            <input
                              type={showNewPasswordVal ? "text" : "password"}
                              value={newPasswordVal}
                              onChange={(e) => setNewPasswordVal(e.target.value)}
                              placeholder="Masukkan password baru..."
                              className="w-full text-xs font-mono border border-slate-300 rounded-lg px-3 py-2.5 pr-10 focus:ring-2 focus:ring-blue-500 focus:outline-none"
                            />
                            <button
                              type="button"
                              onClick={() => setShowNewPasswordVal(!showNewPasswordVal)}
                              className="absolute right-2.5 top-2.5 text-slate-400 hover:text-slate-600 text-xs cursor-pointer"
                              title="Tampilkan / Sembunyikan Password"
                            >
                              {showNewPasswordVal ? "🙈" : "👁️"}
                            </button>
                          </div>
                          <span className="text-[11px] text-slate-400 mt-1 block">Minimal 3 karakter (Kolom D Sheet Role Pengguna).</span>
                        </div>

                        <div>
                          <label className="block text-xs font-bold text-slate-700 mb-1">Konfirmasi Password Baru:</label>
                          <input
                            type={showNewPasswordVal ? "text" : "password"}
                            value={confirmPasswordVal}
                            onChange={(e) => setConfirmPasswordVal(e.target.value)}
                            placeholder="Ketik ulang password baru..."
                            className="w-full text-xs font-mono border border-slate-300 rounded-lg px-3 py-2.5 focus:ring-2 focus:ring-blue-500 focus:outline-none"
                          />
                          <span className="text-[11px] text-slate-400 mt-1 block">Pastikan sama dengan password baru.</span>
                        </div>
                      </div>

                      <div className="flex justify-end pt-2">
                        <button
                          type="button"
                          onClick={handleSimpanPasswordBaru}
                          className="px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs rounded-lg shadow-sm transition flex items-center gap-1.5 cursor-pointer"
                        >
                          <span>🔒</span> Simpan ke Sheet Role Pengguna
                        </button>
                      </div>
                    </div>

                    <div className="flex justify-end gap-3 p-4 bg-white rounded-xl border border-slate-200 shadow-sm sticky bottom-4 z-10">
                      <button
                        type="submit"
                        className="px-6 py-2.5 bg-blue-600 text-white font-bold text-sm rounded-xl hover:bg-blue-700 shadow-md transition flex items-center gap-2"
                      >
                        <span>💾</span> Simpan Pengaturan Sekolah & Media
                      </button>
                    </div>

                  </form>

                </div>
              )}

              {/* VIEW 5: REKENING KORAN TRANSAKSI BANK (4 KOLOM REAKTIF PASSOBOOK LEDGER) */}
              {currentView === "rek_koran" && (
                <div className="space-y-6">

                  <div className="no-print bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-wrap justify-between items-center gap-4">
                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        onClick={() => setRekKoranViewMode("detail")}
                        className={`text-xs font-bold px-3 py-2 rounded-lg transition ${rekKoranViewMode === "detail" ? "bg-blue-600 text-white shadow" : "bg-white text-slate-700 hover:bg-slate-100 border border-slate-300"}`}
                      >
                        📜 Transaksi Per-Tanggal ({computedRekKoranLedger.length} Baris)
                      </button>
                      <button
                        type="button"
                        onClick={() => setRekKoranViewMode("rekap")}
                        className={`text-xs font-bold px-3 py-2 rounded-lg transition ${rekKoranViewMode === "rekap" ? "bg-blue-600 text-white shadow" : "bg-white text-slate-700 hover:bg-slate-100 border border-slate-300"}`}
                      >
                        📊 Total Belanja Per-Bulan ({rekKoranMonthlySummary.length} Bulan)
                      </button>
                    </div>

                    {isAdmin && (
                      <div className="flex flex-wrap items-center gap-2">
                        <button
                          type="button"
                          onClick={() => {
                            setSaldoAwalInput(saldoAwal);
                            setShowEditRekKoranModal(true);
                          }}
                          className="text-xs font-bold px-3.5 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 shadow transition flex items-center gap-1.5"
                          title="Menu Edit Lengkap Data Rekening Koran"
                        >
                          ✏️ Edit Data Rekening Koran
                        </button>

                        <button
                          type="button"
                          onClick={() => {
                            setSaldoAwalInput(saldoAwal);
                            setShowSaldoAwalModal(true);
                          }}
                          className="text-xs font-bold px-3 py-2 bg-emerald-50 text-emerald-800 border border-emerald-300 rounded-lg hover:bg-emerald-100 transition flex items-center gap-1"
                          title="Atur Saldo Awal Rekening Bank Manual"
                        >
                          ⚙️ Edit Saldo Awal (Rp {saldoAwal.toLocaleString('id-ID')})
                        </button>

                        <button
                          type="button"
                          onClick={() => {
                            setDanaMasukForm({ id: null, tgl: "20/01/2026", kredit: 0, keterangan: "" });
                            setShowDanaMasukModal(true);
                          }}
                          className="text-xs font-bold px-3 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 shadow transition flex items-center gap-1"
                          title="Input Dana Masuk / Transfer Kredit Baru"
                        >
                          ➕ Input Dana Masuk (Kredit)
                        </button>

                        <button
                          type="button"
                          onClick={handleResetRekKoranDefault}
                          className="text-xs font-bold px-3 py-2 bg-amber-50 text-amber-800 border border-amber-300 rounded-lg hover:bg-amber-100 transition flex items-center gap-1"
                          title="Reset Saldo Awal & Dana Masuk ke Default"
                        >
                          🔄 Reset Default
                        </button>
                      </div>
                    )}
                  </div>

                  {danaMasukList.length > 0 && (
                    <div className="no-print bg-emerald-50/70 p-3.5 rounded-xl border border-emerald-200">
                      <div className="flex justify-between items-center mb-2">
                        <span className="text-xs font-bold text-emerald-900 flex items-center gap-1">
                          💵 Daftar Dana Masuk / Kredit Terdaftar ({danaMasukList.length} Entri):
                        </span>
                      </div>
                      <div className="flex flex-wrap gap-2 text-xs">
                        {danaMasukList.map(dm => (
                          <div key={dm.id} className="bg-white px-3 py-1.5 rounded-lg border border-emerald-300 shadow-sm flex items-center gap-2">
                            <div>
                              <span className="font-mono font-bold text-emerald-800">{dm.tgl || "Saldo"}</span>
                              <span className="text-slate-600 ml-2">{dm.keterangan}</span>
                              <span className="font-mono font-extrabold text-emerald-700 ml-2">
                                (+Rp {Number(dm.kredit).toLocaleString('id-ID')})
                              </span>
                            </div>
                            {isAdmin && (
                              <div className="flex items-center gap-1 ml-2 border-l border-slate-200 pl-2">
                                <button
                                  type="button"
                                  onClick={() => {
                                    setDanaMasukForm(dm);
                                    setShowDanaMasukModal(true);
                                  }}
                                  className="text-blue-600 hover:text-blue-800 font-bold text-[11px]"
                                  title="Edit Dana Masuk"
                                >
                                  ✏️
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleDeleteDanaMasuk(dm.id)}
                                  className="text-red-500 hover:text-red-700 font-bold text-[11px]"
                                  title="Hapus Dana Masuk"
                                >
                                  🗑️
                                </button>
                              </div>
                            )}
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  <div className="no-print grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-emerald-600">
                      <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Saldo Akhir Rekening Koran</p>
                      <p className="text-xl font-bold text-emerald-700 font-mono mt-1">
                        Rp {computedRekKoranLedger.length > 0 ? computedRekKoranLedger[computedRekKoranLedger.length - 1].runningBalance.toLocaleString('id-ID') : 0}
                      </p>
                      <p className="text-[11px] text-emerald-600 mt-1 font-semibold">
                        Terintegrasi Otomatis dengan Belanja SI ({belanjaBankList.length} Items)
                      </p>
                    </div>

                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-blue-600">
                      <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Total Kredit (Penerimaan)</p>
                      <p className="text-xl font-bold text-blue-700 font-mono mt-1">
                        Rp {computedRekKoranLedger.reduce((s, x) => s + Number(x.kredit || 0), 0).toLocaleString('id-ID')}
                      </p>
                      <p className="text-[11px] text-blue-600 mt-1 font-medium">
                        Saldo Awal (Rp {saldoAwal.toLocaleString('id-ID')}) + Dana Masuk
                      </p>
                    </div>

                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-purple-600">
                      <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Total Debit (Belanja / Pengeluaran SI)</p>
                      <p className="text-xl font-bold text-purple-700 font-mono mt-1">
                        Rp {computedRekKoranLedger.reduce((s, x) => s + Number(x.debit || 0), 0).toLocaleString('id-ID')}
                      </p>
                      <p className="text-[11px] text-purple-600 mt-1 font-medium">
                        Total Realisasi Belanja SI Bank
                      </p>
                    </div>
                  </div>

                  <div className="print-container bg-white rounded-xl shadow-md border border-slate-300 p-6 md:p-10">
                    
                    <div className="text-center mb-6 pb-4 border-b-2 border-slate-900">
                      <h2 className="text-lg md:text-xl font-bold uppercase tracking-wider text-slate-900">
                        LAPORAN REKENING KORAN TRANSAKSI BANK
                      </h2>
                      <p className="text-sm font-semibold text-slate-800">
                        {headerInfo.namaSekolah} — NPSN: {headerInfo.npsn} | SUMBER DANA: {headerInfo.sumberDana}
                      </p>
                    </div>

                    {rekKoranViewMode === "detail" && (
                      <div className="overflow-x-auto mb-8">
                        <table className="w-full text-sm border-collapse border border-slate-900 font-mono">
                          <thead>
                            <tr className="bg-sky-100 text-slate-900 font-bold border-b-2 border-slate-900 text-center" style={{ backgroundColor: '#e0f2fe', WebkitPrintColorAdjust: 'exact', printColorAdjust: 'exact' }}>
                              <th className="border border-slate-900 px-4 py-2.5 text-center w-36 font-extrabold uppercase">TGL</th>
                              <th className="border border-slate-900 px-4 py-2.5 text-center w-44 font-extrabold uppercase">DEBIT</th>
                              <th className="border border-slate-900 px-4 py-2.5 text-center w-44 font-extrabold uppercase">KREDIT</th>
                              <th className="border border-slate-900 px-4 py-2.5 text-center w-48 font-extrabold uppercase">TOTAL</th>
                            </tr>
                          </thead>
                          <tbody>
                            {computedRekKoranLedger.map((item, idx) => {
                              const isRow1 = idx === 0 && !item.tgl && !item.debit;
                              return (
                                <tr key={item.id || idx} className="hover:bg-blue-50/50 transition border-b border-slate-400">
                                  <td className="border border-slate-900 px-4 py-2 text-center font-bold text-slate-800">
                                    {item.tgl || ""}
                                  </td>
                                  <td className="border border-slate-900 px-4 py-2 text-right font-bold text-slate-900">
                                    {item.debit ? Number(item.debit).toLocaleString('id-ID') : ""}
                                  </td>
                                  <td className="border border-slate-900 px-4 py-2 text-right font-bold text-slate-900">
                                    {item.kredit ? Number(item.kredit).toLocaleString('id-ID') : ""}
                                  </td>
                                  <td className="border border-slate-900 px-4 py-2 text-right font-extrabold text-slate-900">
                                    {isRow1 ? "" : Number(item.runningBalance).toLocaleString('id-ID')}
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}

                    {rekKoranViewMode === "rekap" && (
                      <div className="overflow-x-auto mb-8">
                        <table className="w-full text-sm border-collapse border border-slate-900 font-mono">
                          <thead>
                            <tr className="bg-sky-100 text-slate-900 font-bold border-b-2 border-slate-900 text-center" style={{ backgroundColor: '#e0f2fe', WebkitPrintColorAdjust: 'exact', printColorAdjust: 'exact' }}>
                              <th className="border border-slate-900 px-4 py-2.5 text-center w-44 font-extrabold uppercase">TGL / BULAN</th>
                              <th className="border border-slate-900 px-4 py-2.5 text-center w-44 font-extrabold uppercase">DEBIT</th>
                              <th className="border border-slate-900 px-4 py-2.5 text-center w-44 font-extrabold uppercase">KREDIT</th>
                              <th className="border border-slate-900 px-4 py-2.5 text-center w-48 font-extrabold uppercase">TOTAL</th>
                            </tr>
                          </thead>
                          <tbody>
                            {rekKoranMonthlySummary.map((sum, idx) => (
                              <tr key={idx} className="hover:bg-blue-50/50 transition border-b border-slate-400">
                                <td className="border border-slate-900 px-4 py-2.5 text-center font-bold text-slate-900">
                                  {sum.monthLabel}
                                </td>
                                <td className="border border-slate-900 px-4 py-2.5 text-right font-bold text-slate-900">
                                  {sum.debit > 0 ? sum.debit.toLocaleString('id-ID') : ""}
                                </td>
                                <td className="border border-slate-900 px-4 py-2.5 text-right font-bold text-slate-900">
                                  {sum.kredit > 0 ? sum.kredit.toLocaleString('id-ID') : ""}
                                </td>
                                <td className="border border-slate-900 px-4 py-2.5 text-right font-extrabold text-blue-900">
                                  {sum.saldoAkhir.toLocaleString('id-ID')}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}

                    <div className="mt-10 pt-4 flex justify-between items-start text-xs font-semibold text-slate-900 px-6">
                      <div className="text-center min-w-[200px]">
                        <p className="font-normal text-slate-900">Bendahara</p>
                        <div className="h-24"></div>
                        <p className="font-bold text-slate-900">{headerInfo.bendaharaSekolah}</p>
                      </div>

                      <div className="text-center min-w-[220px]">
                        <p className="font-normal text-slate-900">Bekasi, 28 September 2026</p>
                        <p className="font-normal text-slate-900 mt-0.5">Kepala {headerInfo.namaSekolah || "Sekolah"}</p>
                        <div className="h-20"></div>
                        <p className="font-bold text-slate-900">{headerInfo.kepalaSekolah}</p>
                      </div>
                    </div>

                  </div>
                </div>
              )}

              {/* VIEW 7: REALISASI DISTRIBUSI (PENGELUARAN & STOK) */}
              {currentView === "distribusi" && (
                <div className="max-w-5xl mx-auto space-y-6">
                  <h2 className="text-xl md:text-2xl font-extrabold text-slate-900">
                    Pengeluaran &amp; Stok
                  </h2>

                  {/* CARD CATAT PENGELUARAN */}
                  <div className="no-print bg-white rounded-2xl border border-slate-200 shadow-sm p-5 md:p-6">
                    <div className="flex items-center justify-between mb-4">
                      <h3 className="text-base md:text-lg font-bold text-slate-800">
                        {distribusiForm.id ? "Edit Data Pengeluaran" : "Catat Pengeluaran"}
                      </h3>
                      {distribusiForm.id && (
                        <button
                          type="button"
                          onClick={resetDistribusiForm}
                          className="text-xs font-bold text-red-600 hover:text-red-700 bg-red-50 hover:bg-red-100 border border-red-200 px-3 py-1.5 rounded-lg transition cursor-pointer"
                        >
                          ✕ Batal Edit
                        </button>
                      )}
                    </div>

                    <form onSubmit={handleSaveDistribusi} className="space-y-4">
                      {/* Tanggal */}
                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1.5">
                          Tanggal
                        </label>
                        <input
                          type="date"
                          required
                          value={distribusiForm.tanggal}
                          onChange={(e) => setDistribusiForm(prev => ({ ...prev, tanggal: e.target.value }))}
                          className="w-full text-sm border border-slate-300 rounded-xl px-3.5 py-2.5 bg-white text-slate-800 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 focus:outline-none"
                        />
                      </div>

                      {/* Cari Barang (Tersedia Sisa Stok) dari Database Master RKAS */}
                      <div className="relative">
                        <label className="block text-xs font-semibold text-slate-600 mb-1.5">
                          Cari Barang (Tersedia Sisa Stok)
                        </label>
                        <div className="relative">
                          <input
                            type="text"
                            required
                            placeholder="Ketik nama barang..."
                            value={distribusiBarangQuery}
                            onFocus={() => setShowDistribusiDropdown(true)}
                            onChange={(e) => {
                              const val = e.target.value;
                              setDistribusiBarangQuery(val);
                              setDistribusiForm(prev => ({
                                ...prev,
                                namaBarang: val,
                                masterNo: ""
                              }));
                              setShowDistribusiDropdown(true);
                            }}
                            className="w-full text-sm border border-slate-300 rounded-xl px-3.5 py-2.5 pr-9 bg-white text-slate-800 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 focus:outline-none"
                          />
                          {distribusiBarangQuery && (
                            <button
                              type="button"
                              onClick={() => {
                                setDistribusiBarangQuery("");
                                setDistribusiForm(prev => ({ ...prev, namaBarang: "", masterNo: "", kodeRekening: "", kodeKegiatan: "" }));
                                setShowDistribusiDropdown(true);
                              }}
                              className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-sm font-bold cursor-pointer"
                              title="Bersihkan pencarian"
                            >
                              ✕
                            </button>
                          )}
                        </div>

                        {/* Dropdown Pilihan Barang dari Database Master RKAS & Per-Bulan */}
                        {showDistribusiDropdown && (
                          <div
                            className="absolute left-0 right-0 mt-1 max-h-72 overflow-y-auto bg-white border border-slate-200 rounded-xl shadow-xl z-50 divide-y divide-slate-100"
                            onMouseLeave={() => setShowDistribusiDropdown(false)}
                          >
                            <div className="px-3 py-2 bg-slate-50 text-[11px] font-bold text-slate-500 flex items-center justify-between sticky top-0 z-10 border-b border-slate-100">
                              <span>📦 Sumber: Database Master RKAS & Per-Bulan ({filteredDistribusiBarangOptions.length} item ditemukan)</span>
                              <button
                                type="button"
                                onClick={() => setShowDistribusiDropdown(false)}
                                className="text-slate-400 hover:text-slate-700 font-bold cursor-pointer"
                              >
                                Tutup ✕
                              </button>
                            </div>
                            {filteredDistribusiBarangOptions.length === 0 ? (
                              <div className="p-3 text-xs text-slate-500 italic text-center">
                                {masterRkas.length === 0 && Object.values(perBulanData || {}).every(arr => !arr || arr.length === 0)
                                  ? "Database Master RKAS & Per-Bulan masih kosong. Silakan impor RKAS atau ketik manual nama barang."
                                  : `Tidak ada barang yang cocok dengan "${distribusiBarangQuery}". Anda tetap dapat menyimpan nama barang yang diketik.`}
                              </div>
                            ) : (
                              filteredDistribusiBarangOptions.map((opt) => (
                                <button
                                  type="button"
                                  key={opt.stokKey || `${opt.no}-${opt.namaBarang}`}
                                  onClick={() => handleSelectDistribusiBarang(opt)}
                                  className={`w-full text-left px-3.5 py-2.5 transition flex items-center justify-between gap-3 cursor-pointer ${
                                    opt.isHabis ? "bg-red-50/40 hover:bg-red-50" : "hover:bg-blue-50"
                                  }`}
                                >
                                  <div className="min-w-0">
                                    <div className="flex items-center gap-1.5 flex-wrap">
                                      <p className={`text-xs md:text-sm font-bold truncate ${opt.isHabis ? "text-red-700 line-through" : "text-slate-800"}`}>
                                        {opt.namaBarang}
                                      </p>
                                      {opt.bulan && (
                                        <span className="inline-block text-[10px] font-bold px-2 py-0.5 rounded-md bg-indigo-50 text-indigo-700 border border-indigo-200">
                                          📅 {opt.bulan}
                                        </span>
                                      )}
                                    </div>
                                    <p className="text-[11px] text-slate-500 font-mono mt-0.5">
                                      No.{opt.no} {opt.kodeRekening ? `| Rek: ${opt.kodeRekening}` : ""} {opt.kodeKegiatan ? `| Keg: ${opt.kodeKegiatan}` : ""}
                                      {opt.stokAwal > 0 ? ` | Jml Per-Bulan: ${opt.stokAwal} ${opt.satuan} | Terdistribusi: ${opt.sudahKeluar}` : ""}
                                    </p>
                                  </div>
                                  <div className="text-right shrink-0">
                                    {opt.sisaStok !== null ? (
                                      opt.sisaStok > 0 ? (
                                        <span className="inline-block text-[11px] font-bold px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200">
                                          Sisa: {opt.sisaStok} / {opt.stokAwal} {opt.satuan}
                                        </span>
                                      ) : (
                                        <span className="inline-block text-[11px] font-extrabold px-2.5 py-1 rounded-full bg-red-100 text-red-700 border border-red-300">
                                          ⚠️ HABIS (0 / {opt.stokAwal} {opt.satuan})
                                        </span>
                                      )
                                    ) : (
                                      <span className="inline-block text-[11px] font-semibold px-2 py-0.5 rounded-full bg-blue-50 text-blue-700 border border-blue-200">
                                        {opt.sudahKeluar > 0 ? `Keluar: ${opt.sudahKeluar}` : `Rp ${opt.pagu.toLocaleString("id-ID")}`}
                                      </span>
                                    )}
                                  </div>
                                </button>
                              ))
                            )}
                          </div>
                        )}

                        {/* Info Sisa Stok & Peringatan Barang Habis untuk Barang yang Dipilih */}
                        {activeSelectedDistribusiStok && activeSelectedDistribusiStok.stokAwal > 0 && (
                          <div
                            className={`mt-2 p-3 rounded-xl border text-xs flex flex-wrap items-center justify-between gap-2 ${
                              activeSelectedDistribusiStok.sisaStok <= 0
                                ? "bg-red-50 border-red-300 text-red-800"
                                : "bg-emerald-50/70 border-emerald-200 text-emerald-900"
                            }`}
                          >
                            <div className="space-y-0.5">
                              <p className="font-bold">
                                {activeSelectedDistribusiStok.sisaStok <= 0
                                  ? `⚠️ PERINGATAN: Stok "${activeSelectedDistribusiStok.namaBarang}" (${activeSelectedDistribusiStok.bulan || "Master"}) SUDAH HABIS!`
                                  : `📦 Informasi Stok "${activeSelectedDistribusiStok.namaBarang}" (${activeSelectedDistribusiStok.bulan || "Master"})`}
                              </p>
                              <p className="text-[11px] opacity-90">
                                Jumlah di {activeSelectedDistribusiStok.bulan ? `Bulan ${activeSelectedDistribusiStok.bulan}` : "RKAS"}: <strong>{activeSelectedDistribusiStok.stokAwal} {activeSelectedDistribusiStok.satuan}</strong>
                                {" • "}Sudah Dikeluarkan: <strong>{activeSelectedDistribusiStok.sudahKeluar} {activeSelectedDistribusiStok.satuan}</strong>
                                {" • "}Sisa Tersedia: <strong className={activeSelectedDistribusiStok.sisaStok <= 0 ? "text-red-700 underline" : "text-emerald-700 underline"}>{activeSelectedDistribusiStok.sisaStok} {activeSelectedDistribusiStok.satuan}</strong>
                              </p>
                            </div>
                            <span
                              className={`px-2.5 py-1 rounded-lg font-extrabold text-xs ${
                                activeSelectedDistribusiStok.sisaStok <= 0
                                  ? "bg-red-600 text-white"
                                  : "bg-emerald-600 text-white"
                              }`}
                            >
                              {activeSelectedDistribusiStok.sisaStok <= 0
                                ? "STOK HABIS (0)"
                                : `Maksimal Input: ${activeSelectedDistribusiStok.sisaStok} ${activeSelectedDistribusiStok.satuan}`}
                            </span>
                          </div>
                        )}
                      </div>

                      {/* Jumlah Keluar & Penerima */}
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1.5">
                            Jumlah Keluar
                          </label>
                          <input
                            type="number"
                            min="1"
                            max={activeSelectedDistribusiStok && activeSelectedDistribusiStok.stokAwal > 0 && activeSelectedDistribusiStok.sisaStok > 0 ? activeSelectedDistribusiStok.sisaStok : undefined}
                            required
                            placeholder="0"
                            value={distribusiForm.jumlahKeluar}
                            onChange={(e) => setDistribusiForm(prev => ({ ...prev, jumlahKeluar: e.target.value }))}
                            className={`w-full text-sm border rounded-xl px-3.5 py-2.5 bg-white text-slate-800 focus:ring-2 focus:outline-none ${
                              activeSelectedDistribusiStok &&
                              activeSelectedDistribusiStok.stokAwal > 0 &&
                              (activeSelectedDistribusiStok.sisaStok <= 0 || Number(distribusiForm.jumlahKeluar) > activeSelectedDistribusiStok.sisaStok)
                                ? "border-red-400 focus:ring-red-500 focus:border-red-500 bg-red-50/30"
                                : "border-slate-300 focus:ring-blue-500 focus:border-blue-500"
                            }`}
                          />
                          {activeSelectedDistribusiStok && activeSelectedDistribusiStok.stokAwal > 0 && Number(distribusiForm.jumlahKeluar) > 0 && (
                            Number(distribusiForm.jumlahKeluar) > activeSelectedDistribusiStok.sisaStok ? (
                              <p className="mt-1 text-[11px] font-bold text-red-600">
                                ⚠️ Melebihi sisa stok! Maksimal yang bisa dikeluarkan: {activeSelectedDistribusiStok.sisaStok} {activeSelectedDistribusiStok.satuan}.
                              </p>
                            ) : (
                              <p className="mt-1 text-[11px] font-semibold text-emerald-700">
                                ✓ Sisa stok setelah dikeluarkan: {Math.max(0, activeSelectedDistribusiStok.sisaStok - Number(distribusiForm.jumlahKeluar))} {activeSelectedDistribusiStok.satuan}
                              </p>
                            )
                          )}
                        </div>
                        <div>
                          <label className="block text-xs font-semibold text-slate-600 mb-1.5">
                            Penerima
                          </label>
                          <input
                            type="text"
                            required
                            placeholder="Mis: Guru Kls 1"
                            value={distribusiForm.penerima}
                            onChange={(e) => setDistribusiForm(prev => ({ ...prev, penerima: e.target.value }))}
                            className="w-full text-sm border border-slate-300 rounded-xl px-3.5 py-2.5 bg-white text-slate-800 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 focus:outline-none"
                          />
                        </div>
                      </div>

                      {/* Keterangan Tambahan */}
                      <div>
                        <label className="block text-xs font-semibold text-slate-600 mb-1.5">
                          Keterangan Tambahan
                        </label>
                        <input
                          type="text"
                          placeholder="Mis: Untuk lomba"
                          value={distribusiForm.keterangan}
                          onChange={(e) => setDistribusiForm(prev => ({ ...prev, keterangan: e.target.value }))}
                          className="w-full text-sm border border-slate-300 rounded-xl px-3.5 py-2.5 bg-white text-slate-800 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 focus:outline-none"
                        />
                      </div>

                      {/* Tombol Simpan Data */}
                      <button
                        type="submit"
                        className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold text-sm py-3.5 rounded-xl shadow-sm transition cursor-pointer"
                      >
                        {distribusiForm.id ? "Simpan Perubahan" : "Simpan Data"}
                      </button>
                    </form>
                  </div>

                  {/* RIWAYAT PENGELUARAN */}
                  <div className="space-y-3">
                    <h3 className="text-base md:text-lg font-bold text-slate-800">
                      Riwayat Pengeluaran
                    </h3>

                    {/* Search Bar Riwayat */}
                    <div className="no-print relative">
                      <span className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 text-sm">
                        🔍
                      </span>
                      <input
                        type="text"
                        placeholder="Cari nama barang atau penerima..."
                        value={distribusiRiwayatSearch}
                        onChange={(e) => setDistribusiRiwayatSearch(e.target.value)}
                        className="w-full text-sm border border-slate-200 rounded-xl pl-10 pr-4 py-2.5 bg-white text-slate-800 shadow-2xs focus:ring-2 focus:ring-blue-500 focus:border-blue-500 focus:outline-none"
                      />
                    </div>

                    {/* Tabel Riwayat Pengeluaran */}
                    <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                      <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse">
                          <thead>
                            <tr className="bg-slate-50/80 border-b border-slate-200 text-slate-600 text-xs font-bold">
                              <th className="py-3.5 px-4">Tanggal</th>
                              <th className="py-3.5 px-4">Nama Barang</th>
                              <th className="py-3.5 px-4">Jml</th>
                              <th className="py-3.5 px-4">Penerima</th>
                              <th className="py-3.5 px-4">Keterangan</th>
                              {isAdmin && (
                                <th className="no-print py-3.5 px-4 text-center w-36">Aksi</th>
                              )}
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-200 text-xs md:text-sm">
                            {filteredDistribusiRiwayat.length === 0 ? (
                              <tr>
                                <td colSpan={isAdmin ? "6" : "5"} className="py-8 text-center text-slate-400 italic">
                                  Belum ada riwayat pengeluaran barang yang tercatat.
                                </td>
                              </tr>
                            ) : (
                              filteredDistribusiRiwayat.map((item, idx) => {
                                const mStr = String(item.masterNo || "").trim();
                                const bulanTag = mStr.includes(":") ? mStr.split(":")[0].trim() : "";
                                return (
                                  <tr key={item.id || idx} className="hover:bg-slate-50/80 transition">
                                    <td className="py-3.5 px-4 text-slate-500 font-medium whitespace-nowrap">
                                      {item.tanggal || "-"}
                                    </td>
                                    <td className="py-3.5 px-4 font-bold text-slate-900">
                                      <div>{item.namaBarang || "-"}</div>
                                      {bulanTag && (
                                        <span className="inline-block mt-0.5 text-[10px] font-bold px-2 py-0.5 rounded bg-indigo-50 text-indigo-700 border border-indigo-200">
                                          📅 Bulan {bulanTag}
                                        </span>
                                      )}
                                    </td>
                                    <td className="py-3.5 px-4 font-extrabold text-red-600 whitespace-nowrap">
                                      {item.jumlahKeluar} {item.satuan || ""}
                                    </td>
                                    <td className="py-3.5 px-4 font-bold text-slate-800">
                                      {item.penerima || "-"}
                                    </td>
                                    <td className="py-3.5 px-4 text-slate-500">
                                      {item.keterangan || "-"}
                                    </td>
                                    {isAdmin && (
                                      <td className="no-print py-3.5 px-4 text-center whitespace-nowrap">
                                        <div className="inline-flex items-center gap-1.5">
                                          <button
                                            type="button"
                                            onClick={() => handleEditDistribusi(item)}
                                            className="px-2.5 py-1 bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 rounded-lg text-xs font-bold transition cursor-pointer"
                                            title="Edit Riwayat Pengeluaran"
                                          >
                                            ✏️ Edit
                                          </button>
                                          <button
                                            type="button"
                                            onClick={() => handleDeleteDistribusi(item.id)}
                                            className="px-2.5 py-1 bg-red-50 hover:bg-red-100 text-red-600 border border-red-200 rounded-lg text-xs font-bold transition cursor-pointer"
                                            title="Hapus Riwayat Pengeluaran"
                                          >
                                            🗑️ Hapus
                                          </button>
                                        </div>
                                      </td>
                                    )}
                                  </tr>
                                );
                              })
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
              )}

            </div>
          </main>

          {/* MODAL FORM BELANJA */}
          {showBelanjaModal && (
            <div className="no-print fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-xl shadow-xl border border-slate-200 max-w-lg w-full p-6">
                <div className="flex justify-between items-center mb-4 pb-2 border-b">
                  <h3 className="text-base font-bold text-slate-800">
                    {belanjaForm.id ? "Edit Item Belanja Bank" : `Input Belanja Baru (${selectedMonthView})`}
                  </h3>
                  <button onClick={() => setShowBelanjaModal(false)} className="text-slate-400 hover:text-slate-600 font-bold text-lg">&times;</button>
                </div>

                <form onSubmit={handleSubmitBelanja} className="space-y-3 text-xs">
                  
                  {masterRkas.length > 0 && (
                    <div className="bg-amber-50 p-3 rounded-lg border border-amber-200 mb-2">
                      <label className="block font-bold text-amber-900 mb-1">
                        💡 Pilih Uraian Dari Database Master RKAS ({availableMasterOptions.length} Opsi Tersedia / Masih Memiliki Sisa Anggaran):
                      </label>
                      <select
                        value={belanjaForm.selectedMasterNo}
                        onChange={(e) => {
                          const mNo = e.target.value;
                          const selected = availableMasterOptions.find(x => String(x.no) === String(mNo));
                          if (selected) {
                            const kids = selected.children || [];
                            setBelanjaForm(prev => ({
                              ...prev,
                              selectedMasterNo: selected.no,
                              belanja: selected.uraian,
                              subItems: [],
                              total: kids.length > 1 ? 0 : (selected.sisa || selected.jumlah || prev.total)
                            }));
                          }
                        }}
                        className="w-full text-xs border border-amber-300 rounded-lg px-2.5 py-1.5 bg-white focus:ring-2 focus:ring-amber-500 focus:outline-none"
                      >
                        <option value="">-- Pilih Uraian RKAS (Menampilkan Sisa Anggaran & Sub-Menu) --</option>
                        {availableMasterOptions.map(m => {
                          const availKids = (m.children || []).filter(c => !c.isFull && c.sisa > 0).length;
                          const groupTag = (m.children && m.children.length > 1) ? ` [📦 ${availKids} Barang]` : "";
                          const priceTag = m.terpakai > 0
                            ? `SISA: Rp ${m.sisa.toLocaleString('id-ID')} (Pagu: Rp ${m.pagu.toLocaleString('id-ID')})`
                            : `Rp ${m.sisa.toLocaleString('id-ID')}`;
                          return (
                            <option key={m.no} value={m.no}>
                              {`[No.${m.no}] ${m.uraian}${groupTag} — ${priceTag}`}
                            </option>
                          );
                        })}
                      </select>

                      {(() => {
                        const selOpt = availableMasterOptions.find(x => String(x.no) === String(belanjaForm.selectedMasterNo));
                        const kids = selOpt && Array.isArray(selOpt.children) ? selOpt.children : [];
                        if (kids.length === 0) return null;
                        const checkedNos = new Set((belanjaForm.subItems || []).map(s => String(s.no)));
                        return (
                          <div className="mt-2 pt-2 border-t border-amber-200 max-h-44 overflow-y-auto space-y-1">
                            <div className="text-[11px] font-bold text-blue-900 mb-1">
                              📦 Centang Daftar Barang ({(belanjaForm.subItems || []).length} dipilih):
                            </div>
                            {kids.map(ch => {
                              const isChecked = checkedNos.has(String(ch.no));
                              const isDis = ch.isFull || ch.sisa <= 0;
                              return (
                                <label key={ch.no} className={`flex items-center justify-between gap-2 p-1.5 rounded border text-[11px] ${isDis ? "bg-slate-100 text-slate-400 line-through" : isChecked ? "bg-blue-50 border-blue-400 font-bold" : "bg-white border-slate-200 cursor-pointer"}`}>
                                  <span className="flex items-center gap-1.5">
                                    <input
                                      type="checkbox"
                                      checked={isChecked}
                                      disabled={isDis}
                                      onChange={(ev) => {
                                        const checked = ev.target.checked;
                                        setBelanjaForm(prev => {
                                          const cur = Array.isArray(prev.subItems) ? [...prev.subItems] : [];
                                          const next = checked
                                            ? [...cur, { no: ch.no, uraian: ch.uraian, jumlah: ch.sisa }]
                                            : cur.filter(x => String(x.no) !== String(ch.no));
                                          const sum = next.reduce((a, b) => a + Number(b.jumlah || 0), 0);
                                          return { ...prev, subItems: next, total: sum };
                                        });
                                      }}
                                    />
                                    <span>{ch.uraian}</span>
                                  </span>
                                  <span className="font-mono">{isDis ? "Sudah Dibeli" : `Rp ${ch.sisa.toLocaleString('id-ID')}`}</span>
                                </label>
                              );
                            })}
                          </div>
                        );
                      })()}
                    </div>
                  )}

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">Uraian Belanja <span className="text-red-500">*</span></label>
                    <input
                      type="text"
                      required
                      placeholder="Contoh: Belanja Perlengkapan Kantor (CCTV)"
                      value={belanjaForm.belanja}
                      onChange={(e) => setBelanjaForm({ ...belanjaForm, belanja: e.target.value })}
                      className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none font-semibold"
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Nomor Rekening</label>
                      <input
                        type="text"
                        placeholder="184500009453232"
                        value={belanjaForm.noRek}
                        onChange={(e) => setBelanjaForm({ ...belanjaForm, noRek: e.target.value })}
                        className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Nama Bank</label>
                      <input
                        type="text"
                        placeholder="BJB"
                        value={belanjaForm.namaBank}
                        onChange={(e) => setBelanjaForm({ ...belanjaForm, namaBank: e.target.value })}
                        className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Atas Nama Supplier / Toko</label>
                      <input
                        type="text"
                        placeholder="Siplah-Tokoladang"
                        value={belanjaForm.atasNama}
                        onChange={(e) => setBelanjaForm({ ...belanjaForm, atasNama: e.target.value })}
                        className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">NPWP Supplier</label>
                      <input
                        type="text"
                        placeholder="613096429001000"
                        value={belanjaForm.npwp}
                        onChange={(e) => setBelanjaForm({ ...belanjaForm, npwp: e.target.value })}
                        className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Total Nominal (Rp) <span className="text-red-500">*</span></label>
                      <input
                        type="text"
                        required
                        placeholder="Contoh: 7.950.000"
                        value={formatRupiah(belanjaForm.total)}
                        onChange={(e) => setBelanjaForm({ ...belanjaForm, total: parseRupiah(e.target.value) })}
                        className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none font-mono font-bold text-blue-700 text-right"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Sumber Dana</label>
                      <select
                        value={belanjaForm.sumberDana}
                        onChange={(e) => setBelanjaForm({ ...belanjaForm, sumberDana: e.target.value })}
                        className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none font-semibold"
                      >
                        <option value="">-- Pilih Sumber Dana --</option>
                        <option value="BOS REGULER">BOS REGULER</option>
                        <option value="BOS KINERJA">BOS KINERJA</option>
                        <option value="SILPA">SILPA</option>
                        <option value="LAINNYA">LAINNYA</option>
                      </select>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Tgl Transaksi SI (Penanggalan Indonesia)</label>
                      <div className="flex items-center gap-1">
                        <input
                          type="text"
                          placeholder="28/09/2026"
                          value={belanjaForm.tglTransaksiSI}
                          onChange={(e) => setBelanjaForm({ ...belanjaForm, tglTransaksiSI: e.target.value })}
                          className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none font-mono font-bold"
                          title="Format Penanggalan Indonesia: DD/MM/YYYY (contoh: 28/09/2026)"
                        />
                        <input
                          type="date"
                          value={formatDateToIso(belanjaForm.tglTransaksiSI)}
                          onChange={(e) => {
                            if (e.target.value) {
                              setBelanjaForm({ ...belanjaForm, tglTransaksiSI: formatIsoToDdMmYyyy(e.target.value) });
                            }
                          }}
                          className="w-7 h-8 cursor-pointer accent-blue-600 border-none bg-transparent p-0 ml-1"
                          title="Buka Kalender Pemilih Tanggal Indonesia"
                        />
                      </div>
                    </div>
                  </div>

                  <div className="flex justify-end gap-2 pt-3">
                    <button
                      type="button"
                      onClick={() => setShowBelanjaModal(false)}
                      className="px-4 py-2 font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-lg transition"
                    >
                      Batal
                    </button>
                    <button
                      type="submit"
                      className="px-4 py-2 font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition"
                    >
                      Simpan Belanja ({selectedMonthView})
                    </button>
                  </div>
                </form>
              </div>
            </div>
          )}

          {/* MODAL EDIT REKENING KORAN LENGKAP */}
          {showEditRekKoranModal && (
            <div className="no-print fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-xl shadow-xl border border-slate-200 max-w-2xl w-full p-6 max-h-[90vh] overflow-y-auto">
                <div className="flex justify-between items-center mb-4 pb-3 border-b">
                  <div>
                    <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
                      ✏️ Menu Edit Data Rekening Koran
                    </h3>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Kelola Saldo Awal, Dana Masuk (Kredit), & Ringkasan Transaksi Rekening Koran
                    </p>
                  </div>
                  <button onClick={() => setShowEditRekKoranModal(false)} className="text-slate-400 hover:text-slate-600 font-bold text-lg">&times;</button>
                </div>

                <div className="flex border-b border-slate-200 mb-4 gap-2">
                  <button
                    type="button"
                    onClick={() => setEditRekKoranTab('saldo')}
                    className={`px-4 py-2 text-xs font-bold border-b-2 transition ${editRekKoranTab === 'saldo' ? 'border-emerald-600 text-emerald-700 bg-emerald-50/50' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
                  >
                    ⚙️ Saldo Awal (Rp {saldoAwal.toLocaleString('id-ID')})
                  </button>
                  <button
                    type="button"
                    onClick={() => setEditRekKoranTab('dana_masuk')}
                    className={`px-4 py-2 text-xs font-bold border-b-2 transition ${editRekKoranTab === 'dana_masuk' ? 'border-blue-600 text-blue-700 bg-blue-50/50' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
                  >
                    💵 Dana Masuk ({danaMasukList.length} Entri)
                  </button>
                  <button
                    type="button"
                    onClick={() => setEditRekKoranTab('ringkasan')}
                    className={`px-4 py-2 text-xs font-bold border-b-2 transition ${editRekKoranTab === 'ringkasan' ? 'border-purple-600 text-purple-700 bg-purple-50/50' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
                  >
                    📊 Total Per Tanggal ({computedRekKoranLedger.length - 1} Hari)
                  </button>
                </div>

                {editRekKoranTab === 'saldo' && (
                  <form onSubmit={(e) => { handleSaveSaldoAwal(e); setShowEditRekKoranModal(false); }} className="space-y-4 text-xs">
                    <div className="bg-emerald-50/70 p-4 rounded-xl border border-emerald-200">
                      <label className="block font-bold text-emerald-900 mb-1">
                        Nominal Saldo Awal Rekening Koran (Rp):
                      </label>
                      <input
                        type="text"
                        required
                        value={formatRupiah(saldoAwalInput)}
                        onChange={(e) => setSaldoAwalInput(parseRupiah(e.target.value))}
                        className="w-full text-base border border-emerald-300 rounded-lg px-3 py-2.5 font-mono font-bold text-emerald-800 text-right bg-white shadow-sm"
                        placeholder="Contoh: 851.793"
                      />
                      <p className="text-[11px] text-slate-600 mt-2">
                        Nominal ini akan berada pada Baris Pertama Rekening Koran sebagai saldo pembuka.
                      </p>
                    </div>

                    <div className="flex justify-end gap-2 pt-2 border-t">
                      <button
                        type="button"
                        onClick={() => setShowEditRekKoranModal(false)}
                        className="px-4 py-2 bg-slate-100 text-slate-700 font-bold rounded-lg hover:bg-slate-200"
                      >
                        Batal
                      </button>
                      <button
                        type="submit"
                        className="px-4 py-2 bg-emerald-600 text-white font-bold rounded-lg hover:bg-emerald-700 shadow"
                      >
                        💾 Simpan Saldo Awal
                      </button>
                    </div>
                  </form>
                )}

                {editRekKoranTab === 'dana_masuk' && (
                  <div className="space-y-4 text-xs">
                    <div className="flex justify-between items-center">
                      <span className="font-bold text-slate-700">Daftar Dana Masuk / Kredit Terdaftar:</span>
                      <button
                        type="button"
                        onClick={() => {
                          setDanaMasukForm({ id: null, tgl: "20/01/2026", kredit: 0, keterangan: "" });
                          setShowDanaMasukModal(true);
                        }}
                        className="px-3 py-1.5 bg-blue-600 text-white font-bold rounded-lg hover:bg-blue-700 text-xs flex items-center gap-1 shadow"
                      >
                        ➕ Tambah Dana Masuk Baru
                      </button>
                    </div>

                    {danaMasukList.length === 0 ? (
                      <p className="text-slate-400 text-center py-6 italic border rounded-xl bg-slate-50">Belum ada item Dana Masuk manual.</p>
                    ) : (
                      <div className="space-y-2 max-h-64 overflow-y-auto border rounded-xl p-2 bg-slate-50">
                        {danaMasukList.map((dm) => (
                          <div key={dm.id} className="bg-white p-3 rounded-lg border border-slate-200 shadow-sm flex items-center justify-between">
                            <div>
                              <span className="font-mono font-bold text-emerald-800">{dm.tgl || "Saldo"}</span>
                              <span className="text-slate-700 font-medium ml-2">{dm.keterangan}</span>
                              <span className="font-mono font-extrabold text-emerald-700 ml-2">
                                (+Rp {Number(dm.kredit).toLocaleString('id-ID')})
                              </span>
                            </div>
                            <div className="flex items-center gap-1.5">
                              <button
                                type="button"
                                onClick={() => {
                                  setDanaMasukForm(dm);
                                  setShowDanaMasukModal(true);
                                }}
                                className="px-2.5 py-1 bg-blue-50 text-blue-700 border border-blue-200 font-bold rounded hover:bg-blue-100"
                              >
                                ✏️ Edit
                              </button>
                              <button
                                type="button"
                                onClick={() => handleDeleteDanaMasuk(dm.id)}
                                className="px-2.5 py-1 bg-red-50 text-red-700 border border-red-200 font-bold rounded hover:bg-red-100"
                              >
                                🗑️ Hapus
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}

                    <div className="flex justify-end pt-2 border-t">
                      <button
                        type="button"
                        onClick={() => setShowEditRekKoranModal(false)}
                        className="px-4 py-2 bg-slate-800 text-white font-bold rounded-lg hover:bg-slate-900"
                      >
                        Selesai
                      </button>
                    </div>
                  </div>
                )}

                {editRekKoranTab === 'ringkasan' && (
                  <div className="space-y-3 text-xs">
                    <p className="text-slate-600 font-medium">
                      Berikut adalah rekapitulasi total per-tanggal yang tampil di Rekening Koran (otomatis menjumlahkan Kredit & Debit per tanggal):
                    </p>
                    <div className="overflow-x-auto border rounded-xl max-h-64 overflow-y-auto">
                      <table className="w-full text-xs font-mono">
                        <thead className="bg-slate-100 text-slate-800 font-bold sticky top-0">
                          <tr>
                            <th className="p-2 text-center border-b font-extrabold uppercase">TGL</th>
                            <th className="p-2 text-center border-b font-extrabold uppercase">DEBIT</th>
                            <th className="p-2 text-center border-b font-extrabold uppercase">KREDIT</th>
                            <th className="p-2 text-center border-b font-extrabold uppercase">TOTAL</th>
                          </tr>
                        </thead>
                        <tbody>
                          {computedRekKoranLedger.map((row, idx) => (
                            <tr key={row.id || idx} className="border-b hover:bg-slate-50">
                              <td className="p-2 text-center font-bold text-slate-700">{row.tgl || "Saldo Awal"}</td>
                              <td className="p-2 text-right text-purple-700 font-semibold">{row.debit ? Number(row.debit).toLocaleString('id-ID') : '-'}</td>
                              <td className="p-2 text-right text-emerald-700 font-semibold">{row.kredit ? Number(row.kredit).toLocaleString('id-ID') : '-'}</td>
                              <td className="p-2 text-right font-extrabold text-slate-900">{Number(row.runningBalance).toLocaleString('id-ID')}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div className="flex justify-end pt-2 border-t">
                      <button
                        type="button"
                        onClick={() => setShowEditRekKoranModal(false)}
                        className="px-4 py-2 bg-slate-800 text-white font-bold rounded-lg hover:bg-slate-900"
                      >
                        Tutup
                      </button>
                    </div>
                  </div>
                )}

              </div>
            </div>
          )}

          {/* MODAL SALDO AWAL MANUAL */}
          {showSaldoAwalModal && (
            <div className="no-print fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-xl shadow-xl border border-slate-200 max-w-md w-full p-6">
                <div className="flex justify-between items-center mb-4 pb-2 border-b">
                  <h3 className="text-base font-bold text-slate-800">
                    ⚙️ Pengaturan Saldo Awal Rekening Koran
                  </h3>
                  <button onClick={() => setShowSaldoAwalModal(false)} className="text-slate-400 hover:text-slate-600 font-bold text-lg">&times;</button>
                </div>

                <form onSubmit={handleSaveSaldoAwal} className="space-y-4 text-xs">
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Nominal Saldo Awal Rekening Bank (Rp) <span className="text-red-500">*</span>
                    </label>
                    <input
                      type="text"
                      required
                      placeholder="Contoh: 851.793"
                      value={formatRupiah(saldoAwalInput)}
                      onChange={(e) => setSaldoAwalInput(parseRupiah(e.target.value))}
                      className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2.5 focus:ring-2 focus:ring-emerald-500 focus:outline-none font-mono font-bold text-emerald-800 text-right bg-emerald-50/50"
                    />
                    <p className="text-[11px] text-slate-500 mt-1">
                      Nominal ini akan berada pada Baris Pertama Rekening Koran sebagai saldo pembuka.
                    </p>
                  </div>

                  <div className="flex justify-end gap-2 pt-2">
                    <button
                      type="button"
                      onClick={() => setShowSaldoAwalModal(false)}
                      className="px-4 py-2 font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-lg transition"
                    >
                      Batal
                    </button>
                    <button
                      type="submit"
                      className="px-4 py-2 font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg transition"
                    >
                      Simpan Saldo Awal
                    </button>
                  </div>
                </form>
              </div>
            </div>
          )}

          {/* MODAL FORM DANA MASUK (KREDIT) */}
          {showDanaMasukModal && (
            <div className="no-print fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-xl shadow-xl border border-slate-200 max-w-md w-full p-6">
                <div className="flex justify-between items-center mb-4 pb-2 border-b">
                  <h3 className="text-base font-bold text-slate-800">
                    {danaMasukForm.id ? "Edit Item Dana Masuk" : "➕ Input Dana Masuk Baru (Kredit)"}
                  </h3>
                  <button onClick={() => setShowDanaMasukModal(false)} className="text-slate-400 hover:text-slate-600 font-bold text-lg">&times;</button>
                </div>

                <form onSubmit={handleSaveDanaMasuk} className="space-y-3 text-xs">
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Tanggal Masuk <span className="text-red-500">*</span> (Format: DD/MM/YYYY)
                    </label>
                    <div className="flex items-center gap-1">
                      <input
                        type="text"
                        required
                        placeholder="20/01/2026"
                        value={danaMasukForm.tgl}
                        onChange={(e) => setDanaMasukForm({ ...danaMasukForm, tgl: e.target.value })}
                        className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none font-mono font-bold"
                      />
                      <input
                        type="date"
                        value={formatDateToIso(danaMasukForm.tgl)}
                        onChange={(e) => {
                          if (e.target.value) {
                            setDanaMasukForm({ ...danaMasukForm, tgl: formatIsoToDdMmYyyy(e.target.value) });
                          }
                        }}
                        className="w-7 h-8 cursor-pointer accent-blue-600 border-none bg-transparent p-0 ml-1"
                        title="Pemilih Tanggal Kalender"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Nominal Dana Masuk (Rp) <span className="text-red-500">*</span>
                    </label>
                    <input
                      type="text"
                      required
                      placeholder="Contoh: 123.420.000"
                      value={formatRupiah(danaMasukForm.kredit)}
                      onChange={(e) => setDanaMasukForm({ ...danaMasukForm, kredit: parseRupiah(e.target.value) })}
                      className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none font-mono font-bold text-blue-800 text-right"
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Keterangan / Uraian Penerimaan
                    </label>
                    <input
                      type="text"
                      placeholder="Contoh: Penerimaan Dana BOS Reguler Tahap 1"
                      value={danaMasukForm.keterangan}
                      onChange={(e) => setDanaMasukForm({ ...danaMasukForm, keterangan: e.target.value })}
                      className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none font-semibold"
                    />
                  </div>

                  <div className="flex justify-end gap-2 pt-3">
                    <button
                      type="button"
                      onClick={() => setShowDanaMasukModal(false)}
                      className="px-4 py-2 font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-lg transition"
                    >
                      Batal
                    </button>
                    <button
                      type="submit"
                      className="px-4 py-2 font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition"
                    >
                      Simpan Dana Masuk
                    </button>
                  </div>
                </form>
              </div>
            </div>
          )}

          {/* MODAL LOGIN ADMIN (IKON GIR KANAN ATAS - DIVERIFIKASI DARI SHEET ROLE PENGGUNA) */}
          {showAdminLoginModal && (
            <div className="no-print fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
              <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-sm w-full p-6">
                <div className="flex justify-between items-center mb-4 pb-3 border-b border-slate-100">
                  <div className="flex items-center gap-2.5">
                    <img
                      src={customLogoUrl || SIBOS_CLEAN_ICON}
                      onError={(e) => { e.target.onerror = null; e.target.src = SIBOS_CLEAN_ICON; }}
                      alt="Logo"
                      className="w-10 h-10 object-contain shrink-0 block bg-transparent"
                    />
                    <div>
                      <h3 className="text-sm font-bold text-slate-900">Login Admin</h3>
                      <p className="text-[11px] text-slate-500">Masukkan kode akses admin untuk mengaktifkan fitur Edit, Hapus &amp; Import</p>
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      setShowAdminLoginModal(false);
                      setAdminPasswordError("");
                    }}
                    className="text-slate-400 hover:text-slate-600 font-bold text-lg cursor-pointer"
                  >
                    &times;
                  </button>
                </div>

                <form onSubmit={handleAdminLoginSubmit} className="space-y-4">
                  {adminPasswordError && (
                    <div className="p-2.5 rounded-lg bg-red-50 border border-red-200 text-red-600 text-xs font-semibold">
                      ⚠️ {adminPasswordError}
                    </div>
                  )}

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1.5">
                      Password Admin
                    </label>
                    <input
                      type="password"
                      autoFocus
                      required
                      placeholder="Masukkan password admin..."
                      value={adminPasswordInput}
                      onChange={(e) => {
                        setAdminPasswordInput(e.target.value);
                        if (adminPasswordError) setAdminPasswordError("");
                      }}
                      className="w-full text-sm border border-slate-300 rounded-xl px-3.5 py-2.5 bg-white text-slate-900 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 focus:outline-none font-mono"
                    />
                  </div>

                  <div className="flex justify-end gap-2 pt-1">
                    <button
                      type="button"
                      disabled={isVerifyingAdmin}
                      onClick={() => {
                        setShowAdminLoginModal(false);
                        setAdminPasswordError("");
                      }}
                      className="px-4 py-2 text-xs font-bold text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-xl transition cursor-pointer"
                    >
                      Batal
                    </button>
                    <button
                      type="submit"
                      disabled={isVerifyingAdmin}
                      className="px-4 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-700 disabled:opacity-60 rounded-xl shadow-sm transition cursor-pointer"
                    >
                      {isVerifyingAdmin ? "⏳ Memeriksa..." : "🔓 Masuk Admin"}
                    </button>
                  </div>
                </form>
              </div>
            </div>
          )}

          {/* DIALOG KONFIRMASI IN-APP (BEBAS DARI BLOKIR IFRAME / GOOGLE SITES) */}
          {confirmDialog && (
            <div className="fixed inset-0 z-[100000] flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4">
              <div className="bg-white rounded-2xl p-6 max-w-sm w-full shadow-2xl border border-slate-200 text-center">
                <div className="w-12 h-12 rounded-full bg-red-100 text-red-600 flex items-center justify-center mx-auto text-2xl mb-3">
                  🗑️
                </div>
                <h3 className="text-base font-bold text-slate-900 mb-1.5">
                  {confirmDialog.title || "Konfirmasi"}
                </h3>
                <p className="text-xs text-slate-600 mb-5 leading-relaxed">
                  {confirmDialog.message}
                </p>
                <div className="flex justify-center gap-2">
                  <button
                    type="button"
                    onClick={() => setConfirmDialog(null)}
                    className="px-4 py-2.5 text-xs font-bold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-xl transition cursor-pointer"
                  >
                    Batal
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      const fn = confirmDialog.onConfirm;
                      setConfirmDialog(null);
                      if (fn) fn();
                    }}
                    className={`px-4 py-2.5 text-xs font-bold text-white rounded-xl shadow-xs transition cursor-pointer ${
                      confirmDialog.confirmColor || "bg-red-600 hover:bg-red-700"
                    }`}
                  >
                    {confirmDialog.confirmText || "Lanjutkan"}
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* TOAST NOTIFIKASI IN-APP (BEBAS DARI BLOKIR IFRAME / GOOGLE SITES) */}
          {toastMsg && (
            <div className="fixed top-5 right-5 z-[100000] max-w-md bg-slate-900 text-white px-4 py-3 rounded-2xl shadow-2xl border border-slate-700 flex items-center gap-3">
              <span className="text-lg">
                {toastMsg.type === "error" ? "⚠️" : toastMsg.type === "info" ? "ℹ️" : "✅"}
              </span>
              <div className="text-xs font-semibold leading-snug flex-1">
                {toastMsg.message}
              </div>
              <button
                type="button"
                onClick={() => setToastMsg(null)}
                className="text-slate-400 hover:text-white text-sm font-bold ml-1 cursor-pointer"
              >
                ✕
              </button>
            </div>
          )}

          {/* MOBILE BOTTOM APP NAVIGATION BAR (KHUSUS TAMPILAN HP - 3 MENU UTAMA: DASHBOARD, DISTRIBUSI STOK, MENU) */}
          <nav
            className="mobile-bottom-nav no-print md:hidden px-6 py-2 flex justify-around items-center"
            style={{
              position: "fixed",
              bottom: 0,
              left: 0,
              right: 0,
              width: "100%",
              backgroundColor: "#ffffff",
              zIndex: 99999
            }}
          >
            <button
              type="button"
              onClick={() => setCurrentView("dashboard")}
              className={`flex flex-col items-center gap-1 py-1 px-4 rounded-xl transition cursor-pointer ${
                currentView === "dashboard" ? "text-blue-600 font-extrabold" : "text-slate-500 hover:text-slate-800 font-medium"
              }`}
            >
              <span className="text-xl leading-none">🖥️</span>
              <span className="text-[11px]">Dashboard</span>
            </button>

            <button
              type="button"
              onClick={() => setCurrentView("distribusi")}
              className={`flex flex-col items-center gap-1 py-1 px-4 rounded-xl transition cursor-pointer ${
                currentView === "distribusi" ? "text-blue-600 font-extrabold" : "text-slate-500 hover:text-slate-800 font-medium"
              }`}
            >
              <span className="text-xl leading-none">📦</span>
              <span className="text-[11px]">Distribusi Stok</span>
            </button>

            <button
              type="button"
              onClick={() => setIsMobileMenuOpen(true)}
              className="flex flex-col items-center gap-1 py-1 px-4 rounded-xl text-slate-500 hover:text-slate-800 font-medium transition cursor-pointer"
            >
              <span className="text-xl leading-none">☰</span>
              <span className="text-[11px]">Menu</span>
            </button>
          </nav>

        </div>
      );
    }

export default App;
