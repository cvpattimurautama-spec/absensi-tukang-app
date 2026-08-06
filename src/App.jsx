import React, { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { db } from './firebaseClient.js';
import {
  collection, doc, getDocs, setDoc, updateDoc, deleteDoc, addDoc, query,
} from 'firebase/firestore';
import bcrypt from 'bcryptjs';
import {
  LogIn, Loader2, AlertTriangle, Users, Plus, Trash2, Wallet,
  Camera, MapPin, ChevronDown, ChevronUp, Check, X, CalendarDays, Edit3, Printer, Receipt, Copy, Save, Banknote, ShoppingCart, Warehouse, FileBarChart, LayoutDashboard, Bluetooth, Video, RefreshCw
} from 'lucide-react';

const THEME = {
  charcoal: '#0D1930',
  paper: '#FAF8F1',
  amber: '#C9A227',
  amberSoft: '#E9D48A',
  concrete: '#F1ECDD',
  ink: '#14213D',
  inkSoft: '#5B6478',
  line: '#DDD5BE',
  rust: '#B23A2E',
  green: '#2F7A52',
};

const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700;9..144,900&family=Work+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');`;

const DAYS = ['SN', 'SL', 'RB', 'KM', 'JM', 'SB', 'MG'];
const DAY_LABELS = { SN: 'Senin', SL: 'Selasa', RB: 'Rabu', KM: 'Kamis', JM: 'Jumat', SB: 'Sabtu', MG: 'Minggu' };

const KATEGORI_LIST = ['material', 'peralatan', 'logistik', 'suku_cadang', 'operasional', 'non_kategori'];
const KATEGORI_LABELS = {
  material: 'Material', peralatan: 'Peralatan', logistik: 'Logistik',
  suku_cadang: 'Suku Cadang', operasional: 'Operasional', non_kategori: 'Non Kategori',
};

const PROJECT_STATUS_LABELS = { berjalan: 'Berjalan', selesai: 'Selesai', ditunda: 'Ditunda' };
const PROJECT_STATUS_COLORS = { berjalan: '#C9A227', selesai: '#2F7A52', ditunda: '#B23A2E' };

const GUDANG_UTAMA_ID = 'utama';

const COMPANY_LOGO_DATA_URI = 'data:image/jpeg;base64,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';


let CURRENT_COMPANY = { name: '', address: '', phone: '', tagline: '' };

const DEFAULT_COMPANY_SERVICES = [
  'Jasa Konstruksi',
  'Jasa Transportasi',
  'Jasa Cor Beton',
  'Jasa Penyewaan Alat Konstruksi / Alat Berat',
  'Jasa Pengelasan, CNC Cutting, CNC Router',
  'Industri Kreatif',
  'Perdagangan Besar dan Eceran',
].join('\n');

const DEFAULT_COMPANY_MATERIALS_NOTE = 'Melayani pembelian Pasir, Batu Split semua ukuran, Batu Gunung, Krokos, dan material lainnya.';


function companyHeaderText() {
  if (!CURRENT_COMPANY.name) return '';
  const lines = [CURRENT_COMPANY.name];
  if (CURRENT_COMPANY.tagline) lines.push(CURRENT_COMPANY.tagline);
  if (CURRENT_COMPANY.address) lines.push(CURRENT_COMPANY.address);
  if (CURRENT_COMPANY.phone) lines.push(`Telp: ${CURRENT_COMPANY.phone}`);
  return lines.join('\n') + '\n================================\n';
}

function companyHeaderHtml() {
  return `<div style="text-align:center;border-bottom:2px solid #000;padding-bottom:8px;margin-bottom:10px;">
    <img src="${CURRENT_COMPANY.logoDataUri || COMPANY_LOGO_DATA_URI}" alt="Logo" style="width:56px;height:56px;object-fit:cover;border-radius:6px;margin:0 auto 4px;display:block;" />
    ${CURRENT_COMPANY.name ? `<div style="font-weight:700;font-size:1.15em;">${CURRENT_COMPANY.name}</div>` : ''}
    ${CURRENT_COMPANY.tagline ? `<div style="font-size:0.85em;">${CURRENT_COMPANY.tagline}</div>` : ''}
    ${CURRENT_COMPANY.address ? `<div style="font-size:0.8em;">${CURRENT_COMPANY.address}</div>` : ''}
    ${CURRENT_COMPANY.phone ? `<div style="font-size:0.8em;">Telp: ${CURRENT_COMPANY.phone}</div>` : ''}
  </div>`;
}
function getGudangList(projects) {
  return [
    { id: GUDANG_UTAMA_ID, name: 'Gudang Utama' },
    ...projects.map((p) => ({ id: `proj:${p.id}`, name: `Gudang ${p.name}`, projectId: p.id })),
  ];
}

function uid() {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

function formatRupiah(n) {
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(n || 0);
}

function terbilang(n) {
  const satuan = ['', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan', 'Sembilan', 'Sepuluh',
    'Sebelas', 'Dua Belas', 'Tiga Belas', 'Empat Belas', 'Lima Belas', 'Enam Belas', 'Tujuh Belas', 'Delapan Belas', 'Sembilan Belas'];
  function convert(num) {
    num = Math.floor(num);
    if (num < 20) return satuan[num];
    if (num < 100) return `${satuan[Math.floor(num / 10)]} Puluh${num % 10 ? ' ' + satuan[num % 10] : ''}`;
    if (num < 200) return `Seratus${num % 100 ? ' ' + convert(num % 100) : ''}`;
    if (num < 1000) return `${satuan[Math.floor(num / 100)]} Ratus${num % 100 ? ' ' + convert(num % 100) : ''}`;
    if (num < 2000) return `Seribu${num % 1000 ? ' ' + convert(num % 1000) : ''}`;
    if (num < 1000000) return `${convert(Math.floor(num / 1000))} Ribu${num % 1000 ? ' ' + convert(num % 1000) : ''}`;
    if (num < 1000000000) return `${convert(Math.floor(num / 1000000))} Juta${num % 1000000 ? ' ' + convert(num % 1000000) : ''}`;
    if (num < 1000000000000) return `${convert(Math.floor(num / 1000000000))} Miliar${num % 1000000000 ? ' ' + convert(num % 1000000000) : ''}`;
    return `${convert(Math.floor(num / 1000000000000))} Triliun${num % 1000000000000 ? ' ' + convert(num % 1000000000000) : ''}`;
  }
  const val = Math.round(Number(n) || 0);
  if (val === 0) return 'Nol Rupiah';
  return `${convert(val)} Rupiah`.replace(/\s+/g, ' ').trim();
}

function resizeImage(file, maxWidth = 360, quality = 0.55) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Gagal membaca file'));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error('Gagal memuat gambar'));
      img.onload = () => {
        const scale = Math.min(1, maxWidth / img.width);
        const canvas = document.createElement('canvas');
        canvas.width = img.width * scale;
        canvas.height = img.height * scale;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL('image/jpeg', quality));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

const DAY_NAMES_ID = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
const MONTH_NAMES_ID = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

function getWeekDates(week) {
  let start = week?.startDate ? new Date(week.startDate + 'T00:00:00') : null;
  let end = week?.endDate ? new Date(week.endDate + 'T00:00:00') : null;
  if (!start || isNaN(start.getTime())) start = new Date();
  if (!end || isNaN(end.getTime()) || end < start) end = new Date(start.getTime() + 6 * 86400000);
  const maxDays = 62; // pengaman agar tidak render ribuan baris kalau tanggal salah input
  const dates = [];
  const cur = new Date(start);
  let count = 0;
  while (cur <= end && count < maxDays) {
    const key = cur.toISOString().slice(0, 10);
    const dayName = DAY_NAMES_ID[cur.getDay()];
    dates.push({
      key,
      dayName,
      label: `${dayName}, ${cur.getDate()} ${MONTH_NAMES_ID[cur.getMonth()]} ${cur.getFullYear()}`,
      shortLabel: `${dayName.slice(0, 3)} ${cur.getDate()}/${cur.getMonth() + 1}`,
    });
    cur.setDate(cur.getDate() + 1);
    count++;
  }
  return dates;
}

function dayStatus(rec) {
  if (!rec) return 'off';
  if (rec.status) return rec.status;
  return rec.hadir ? 'full' : 'off';
}

function purchaseItems(p) {
  if (Array.isArray(p.items) && p.items.length > 0) return p.items;
  if (p.materialId) {
    return [{ id: p.id, materialId: p.materialId, materialName: p.materialName, unit: p.unit, kategori: p.kategori, qty: p.qty, price: p.price }];
  }
  return [];
}
function purchaseTotal(p) {
  return purchaseItems(p).reduce((s, it) => s + (Number(it.qty) || 0) * (Number(it.price) || 0), 0);
}

function calcWorkerWeekTotals(worker, weekRecord, dateKeys) {
  const keys = (dateKeys && dateKeys.length ? dateKeys : DAYS.map((d) => d));
  let totalHariHadir = 0;
  let hariTambahan = 0;
  let totalJamLembur = 0;
  keys.forEach((d) => {
    const rec = (weekRecord && weekRecord[d]) || {};
    const status = dayStatus(rec);
    if (status === 'full') totalHariHadir += 1;
    else if (status === 'half') totalHariHadir += 0.5;
    const jam = Number(rec.lembur) || 0;
    if (jam >= 4) {
      hariTambahan += 1;
    } else if (jam > 0) {
      totalJamLembur += jam;
    }
  });
  const totalHariBayar = totalHariHadir + hariTambahan;
  const upahHarian = Number(worker.upahHarian) || 0;
  const upahLembur = Number(worker.upahLembur) || 0;
  const totalHarianRp = totalHariBayar * upahHarian;
  const totalLemburRp = totalJamLembur * upahLembur;
  return {
    totalHariHadir,
    hariTambahan,
    totalHariBayar,
    totalJamLembur,
    totalHarianRp,
    totalLemburRp,
    totalUpah: totalHarianRp + totalLemburRp,
  };
}

function GlobalStyle() {
  return (
    <style>{`
      ${FONT_IMPORT}
      .att-display { font-family: 'Archivo Black', sans-serif; }
      .att-body { font-family: 'Barlow', sans-serif; }
      .att-mono { font-family: 'IBM Plex Mono', monospace; }
      .att-scroll::-webkit-scrollbar { width: 6px; height: 6px; }
      .att-scroll::-webkit-scrollbar-thumb { background: ${THEME.line}; border-radius: 4px; }
      @keyframes punchIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
      .att-punch-anim { animation: punchIn 0.25s ease-out; }
    `}</style>
  );
}

function escapeHtml(str) {
  return String(str ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

// Membuka halaman cetak di tab baru (di luar sandbox artifact) supaya dialog print browser
// benar-benar terpanggil dan bisa memilih printer apa saja, termasuk printer thermal.
function printToRawBT(bodyText) {
  const fullText = companyHeaderText() + bodyText;
  try {
    const b64 = btoa(unescape(encodeURIComponent(fullText)));
    const url = `rawbt:base64,${b64}`;
    const a = document.createElement('a');
    a.href = url;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } catch (err) {
    console.error('Gagal membuka RawBT:', err);
    alert('Gagal membuka aplikasi RawBT. Pastikan aplikasi RawBT sudah terpasang di HP (Android) dan browser diizinkan membuka link "rawbt:".');
  }
}

function printKasbonInvoice(title, worker, amount, date, note, jenis) {
  const lines = [
    title.toUpperCase(),
    '================================',
    `Nama: ${worker?.name || '-'}`,
    `Tanggal: ${date}`,
    `Jenis: ${jenis === 'bayar' ? 'Pembayaran Kasbon' : 'Pemberian Kasbon'}`,
    `Jumlah: ${formatRupiah(amount)}`,
    note ? `Catatan: ${note}` : '',
    '================================',
    '',
    'Tanda tangan,',
    '',
    '',
    '______________',
  ].filter((l) => l !== undefined).join('\n');
  printToRawBT(lines);
}

function openPrintDocument(title, bodyHtml, thermal) {
  const pageCss = thermal
    ? `@page { size: 58mm auto; margin: 2mm; } body{width:54mm;margin:0 auto;font-family:'Courier New',monospace;font-size:10px;color:#000;}`
    : `@page { size: landscape; margin: 10mm; } body{font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#000;}`;
  const htmlWithAutoPrint = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title>
<style>
  ${pageCss}
  table{width:100%;border-collapse:collapse;margin-top:6px;}
  th,td{border:${thermal ? '0' : '1px solid #999'};padding:${thermal ? '1px 0' : '3px 5px'};text-align:left;font-size:${thermal ? '9px' : '11px'};}
  th{background:${thermal ? 'transparent' : '#eee'};border-bottom:1px solid #000;}
  .right{text-align:right;}
  .center{text-align:center;}
  .bold{font-weight:700;}
  .dash{border-top:1px dashed #000;margin:4px 0;}
  h2{font-size:${thermal ? '12px' : '16px'};margin:0 0 4px;}
  p{margin:1px 0;font-size:${thermal ? '9px' : '11px'};}
</style>
</head><body>
${companyHeaderHtml()}
${bodyHtml}
<script>window.onload = function () { setTimeout(function () { window.print(); }, 250); };</script>
</body></html>`;

  // Coba buka tab cetak baru dulu (pengalaman terbaik: langsung muncul dialog print browser).
  try {
    const win = window.open('', '_blank');
    if (win) {
      win.document.write(htmlWithAutoPrint);
      win.document.close();
      return true;
    }
  } catch (err) {
    console.error('Gagal membuka tab cetak:', err);
  }

  // Fallback: browser memblokir popup -> unduh sebagai file HTML yang bisa dibuka manual.
  let ok = false;
  try {
    const htmlWithManualButton = htmlWithAutoPrint.replace(
      '<body>',
      '<body><button onclick="window.print()" style="display:block;width:100%;padding:12px;margin-bottom:10px;font-size:14px;font-weight:700;background:#000;color:#fff;border:none;border-radius:4px;">Cetak Sekarang</button>'
    );
    const blob = new Blob([htmlWithManualButton], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = title.replace(/\s+/g, '_') + (thermal ? '_struk' : '') + '.html';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    ok = true;
  } catch (err) {
    console.error('Gagal membuat file cetak:', err);
    alert('Gagal membuat file cetak. Coba lagi, atau screenshot data ini sebagai alternatif.');
  }
  return ok;
}

function buildPekerjaHtml(workers, thermal) {
  if (thermal) {
    const rows = workers.map((w) => `
      <p class="bold">${escapeHtml(w.name)}</p>
      <p>${escapeHtml(w.position || '-')}</p>
      <p>Harian: ${formatRupiah(w.upahHarian)}</p>
      <p>Lembur/jam: ${formatRupiah(w.upahLembur)}</p>
      <div class="dash"></div>`).join('');
    return `<h2>DATA PEKERJA</h2><div class="dash"></div>${rows}`;
  }
  const rows = workers.map((w, i) => `
    <tr><td>${i + 1}</td><td>${escapeHtml(w.name)}</td><td>${escapeHtml(w.address)}</td><td>${escapeHtml(w.position)}</td>
    <td>${escapeHtml(w.phone)}</td><td>${escapeHtml(w.ktp)}</td>
    <td>${formatRupiah(w.upahHarian)}</td><td>${formatRupiah(w.upahLembur)}</td></tr>`).join('');
  return `<h2>DATA PEKERJA</h2>
    <table><thead><tr><th>No</th><th>Nama</th><th>Alamat</th><th>Jabatan</th><th>No. HP</th><th>No. KTP</th>
    <th>Upah Harian (Rp)</th><th>Upah Lembur/Jam (Rp)</th></tr></thead><tbody>${rows}</tbody></table>`;
}

function idCardHtml(w) {
  const logo = CURRENT_COMPANY.logoDataUri || COMPANY_LOGO_DATA_URI;
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>ID Card - ${escapeHtml(w.name)}</title>
<style>
  @page { size: 340px 540px; margin: 0; }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: 'Arial', sans-serif; }
  .card { width: 340px; height: 540px; background: #FAF8F1; position: relative; overflow: hidden; border: 1px solid #DDD5BE; }
  .head { background: #0D1930; color: #F2ECD9; padding: 18px 16px 14px; text-align: center; border-bottom: 4px solid #C9A227; }
  .head img { width: 42px; height: 42px; object-fit: cover; border-radius: 8px; margin: 0 auto 6px; display: block; border: 1px solid #C9A227; }
  .head .cname { font-weight: 800; font-size: 15px; letter-spacing: 0.02em; }
  .head .ctag { font-size: 8px; color: #C9A227; margin-top: 3px; letter-spacing: 0.06em; text-transform: uppercase; }
  .photo-wrap { text-align: center; margin-top: 22px; }
  .photo { width: 140px; height: 140px; border-radius: 10px; object-fit: cover; border: 3px solid #C9A227; margin: 0 auto; display: block; background: #EFE9D8; }
  .name { text-align: center; font-weight: 800; font-size: 17px; color: #0D1930; margin-top: 14px; padding: 0 12px; }
  .pos { text-align: center; font-size: 11px; color: #C9A227; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 2px; }
  .divider { width: 60px; height: 2px; background: #C9A227; margin: 12px auto; }
  .info { padding: 0 24px; font-size: 10.5px; color: #14213D; }
  .info .row { display: flex; justify-content: space-between; padding: 5px 0; border-bottom: 1px solid #DDD5BE; }
  .info .lbl { color: #5B6478; }
  .info .val { font-weight: 700; }
  .foot { position: absolute; bottom: 0; left: 0; right: 0; background: #0D1930; color: #C9A227; font-size: 8px; text-align: center; padding: 8px; letter-spacing: 0.03em; }
  .noprint { text-align: center; padding: 10px; }
</style>
</head><body>
  <div class="card">
    <div class="head">
      <img src="${logo}" alt="Logo" />
      <div class="cname">${escapeHtml(CURRENT_COMPANY.name || 'ABSENSI TUKANG')}</div>
      ${CURRENT_COMPANY.tagline ? `<div class="ctag">${escapeHtml(CURRENT_COMPANY.tagline)}</div>` : ''}
    </div>
    <div class="photo-wrap">
      ${w.photoDataUri ? `<img class="photo" src="${w.photoDataUri}" alt="${escapeHtml(w.name)}" />` : `<div class="photo"></div>`}
    </div>
    <div class="name">${escapeHtml(w.name)}</div>
    <div class="pos">${escapeHtml(w.position || '-')}</div>
    <div class="divider"></div>
    <div class="info">
      ${w.joinDate ? `<div class="row"><span class="lbl">Bergabung</span><span class="val">${escapeHtml(w.joinDate)}</span></div>` : ''}
      ${w.phone ? `<div class="row"><span class="lbl">No. HP</span><span class="val">${escapeHtml(w.phone)}</span></div>` : ''}
      ${w.ktp ? `<div class="row"><span class="lbl">No. KTP</span><span class="val">${escapeHtml(w.ktp)}</span></div>` : ''}
    </div>
    <div class="foot">${escapeHtml(CURRENT_COMPANY.website || CURRENT_COMPANY.phone || '')}</div>
  </div>
  <script>window.onload = function () { setTimeout(function () { window.print(); }, 250); };</script>
</body></html>`;
}

function openIdCardPrint(worker) {
  try {
    const win = window.open('', '_blank');
    if (win) {
      win.document.write(idCardHtml(worker));
      win.document.close();
      return;
    }
  } catch (err) {
    console.error('Gagal membuka ID card:', err);
  }
  alert('Gagal membuka jendela cetak. Pastikan pop-up tidak diblokir browser.');
}

function statusChar(rec) {
  const s = dayStatus(rec);
  if (s === 'full') return 'H' + (rec.lembur ? '+' + rec.lembur + 'j' : '');
  if (s === 'half') return '½' + (rec.lembur ? '+' + rec.lembur + 'j' : '');
  return '-';
}

function buildAbsensiHtml(activeWeek, project, workers, grandTotal, dates, thermal) {
  const dateKeys = dates.map((d) => d.key);
  if (thermal) {
    const rows = workers.map((w) => {
      const rec = activeWeek.records[w.id] || {};
      const totals = calcWorkerWeekTotals(w, rec, dateKeys);
      const days = dates.map((d) => `${d.shortLabel}:${statusChar(rec[d.key])}`).join(' ');
      return `
        <p class="bold">${escapeHtml(w.name)}</p>
        <p>${days}</p>
        <p>Hari: ${totals.totalHariBayar} &nbsp; Lembur: ${totals.totalJamLembur}j</p>
        <p>Upah: ${formatRupiah(totals.totalUpah)}</p>
        <div class="dash"></div>`;
    }).join('');
    return `<h2>ABSENSI UPAH</h2>
      <p>${escapeHtml(project?.name || '')}</p>
      <p>${escapeHtml(activeWeek.weekLabel)}</p>
      <p>${escapeHtml(activeWeek.startDate)} s/d ${escapeHtml(activeWeek.endDate)}</p>
      <div class="dash"></div>
      ${rows}
      <p class="bold">TOTAL: ${formatRupiah(grandTotal)}</p>`;
  }
  const rows = workers.map((w, i) => {
    const rec = activeWeek.records[w.id] || {};
    const totals = calcWorkerWeekTotals(w, rec, dateKeys);
    const dayCells = dates.map((d) => `<td>${statusChar(rec[d.key])}</td>`).join('');
    return `<tr><td>${i + 1}</td><td>${escapeHtml(w.name)}</td><td>${escapeHtml(w.position)}</td>${dayCells}
      <td>${totals.totalHariBayar}</td><td>${totals.totalJamLembur}</td>
      <td>${formatRupiah(totals.totalHarianRp)}</td><td>${formatRupiah(totals.totalLemburRp)}</td>
      <td>${formatRupiah(totals.totalUpah)}</td></tr>`;
  }).join('');
  const dayHeaders = dates.map((d) => `<th>${escapeHtml(d.shortLabel)}</th>`).join('');
  return `<h2>ABSENSI UPAH HARIAN DAN LEMBURAN TUKANG</h2>
    <p>PROYEK: ${escapeHtml(project?.name || '')}</p>
    <p>Alamat: ${escapeHtml(project?.address || '')}</p>
    <p>${escapeHtml(activeWeek.weekLabel)} &nbsp; Tanggal: ${escapeHtml(activeWeek.startDate)} s/d ${escapeHtml(activeWeek.endDate)}</p>
    <table><thead><tr><th>No</th><th>Nama</th><th>Posisi</th>${dayHeaders}
      <th>Total Hari</th><th>Total Jam Lembur</th><th>Total Harian (Rp)</th><th>Total Lembur (Rp)</th><th>Total Upah (Rp)</th></tr></thead>
      <tbody>${rows}</tbody>
      <tfoot><tr><td colspan="${dates.length + 8}" class="right bold">JUMLAH TOTAL UPAH: ${formatRupiah(grandTotal)}</td></tr></tfoot></table>
    <p style="margin-top:24px;display:flex;justify-content:space-between;">
      <span>Dibuat oleh: ______________</span><span>Disetujui oleh: ______________</span><span>Diterima oleh: ______________</span>
    </p>`;
}

function buildRekapHtml(rows, grand, thermal) {
  if (thermal) {
    const body = rows.map(({ worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa }) => `
      <p class="bold">${escapeHtml(worker.name)}</p>
      <p>${totalHari} hari, ${totalJamLembur} jam lembur</p>
      <p>Total: ${formatRupiah(totalUpah)}</p>
      <p>Diterima: ${formatRupiah(diterima)}</p>
      <p>Kasbon: ${formatRupiah(totalKasbon)}</p>
      <p>Sisa: ${formatRupiah(sisa)}</p>
      <div class="dash"></div>`).join('');
    return `<h2>REKAP UPAH</h2><div class="dash"></div>${body}
      <p class="bold">TOTAL SISA: ${formatRupiah(grand.sisa)}</p>`;
  }
  const body = rows.map(({ worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa }, i) => `
    <tr><td>${i + 1}</td><td>${escapeHtml(worker.name)}</td><td>${escapeHtml(worker.position)}</td>
    <td>${totalHari}</td><td>${totalJamLembur}</td>
    <td>${formatRupiah(totalUpah)}</td><td>${formatRupiah(diterima)}</td><td>${formatRupiah(totalKasbon)}</td><td>${formatRupiah(sisa)}</td><td></td></tr>`).join('');
  return `<h2>REKAP UPAH TUKANG - SELURUH PERIODE</h2>
    <table><thead><tr><th>No</th><th>Nama</th><th>Posisi</th><th>Hari</th><th>Jam Lembur</th><th>Jumlah Total Upah (Rp)</th>
    <th>Upah Diterima (Rp)</th><th>Kasbon (Rp)</th><th>Sisa Upah (Rp)</th><th>Paraf Diterima</th></tr></thead>
    <tbody>${body}</tbody>
    <tfoot><tr class="bold"><td colspan="5">TOTAL</td><td>${formatRupiah(grand.totalUpah)}</td>
    <td>${formatRupiah(grand.diterima)}</td><td>${formatRupiah(grand.totalKasbon)}</td><td>${formatRupiah(grand.sisa)}</td><td></td></tr></tfoot></table>`;
}

function buildKasbonHtml(sorted, workerName, total, thermal) {
  if (thermal) {
    const body = sorted.map((k) => `
      <p class="bold">${escapeHtml(workerName(k.workerId))} [${k.jenis === 'bayar' ? 'BAYAR' : 'KASBON'}]</p>
      <p>${escapeHtml(k.date)} - ${formatRupiah(k.amount)}</p>
      ${k.note ? `<p>${escapeHtml(k.note)}</p>` : ''}
      <div class="dash"></div>`).join('');
    return `<h2>KASBON PEKERJA</h2><div class="dash"></div>${body}<p class="bold">TOTAL DIBERIKAN: ${formatRupiah(total)}</p>`;
  }
  const body = sorted.map((k, i) => `
    <tr><td>${i + 1}</td><td>${escapeHtml(k.date)}</td><td>${escapeHtml(workerName(k.workerId))}</td><td>${k.jenis === 'bayar' ? 'Bayar' : 'Kasbon'}</td>
    <td>${formatRupiah(k.amount)}</td><td>${escapeHtml(k.note || '')}</td></tr>`).join('');
  return `<h2>KASBON PEKERJA</h2>
    <table><thead><tr><th>No</th><th>Tanggal</th><th>Nama</th><th>Jenis</th><th>Jumlah (Rp)</th><th>Catatan</th></tr></thead>
    <tbody>${body}</tbody>
    <tfoot><tr class="bold"><td colspan="4">TOTAL DIBERIKAN</td><td>${formatRupiah(total)}</td><td></td></tr></tfoot></table>`;
}

function buildBelanjaHtml(sorted, total, thermal) {
  if (thermal) {
    const body = sorted.map((p) => {
      const items = purchaseItems(p);
      const itemLines = items.map((it) => `<p>${escapeHtml(it.materialName)}: ${it.qty} ${escapeHtml(it.unit || '')} x ${formatRupiah(it.price)} = ${formatRupiah(it.qty * it.price)}</p>`).join('');
      return `
      <p class="bold">${escapeHtml(p.date)}${p.noNota ? ' - Nota ' + escapeHtml(p.noNota) : ''}</p>
      ${p.supplierName ? `<p>${escapeHtml(p.supplierName)}</p>` : ''}
      ${itemLines}
      <p class="bold">= ${formatRupiah(purchaseTotal(p))}</p>
      <div class="dash"></div>`;
    }).join('');
    return `<h2>BELANJA HARIAN</h2><div class="dash"></div>${body}<p class="bold">TOTAL: ${formatRupiah(total)}</p>`;
  }
  const body = sorted.map((p, i) => {
    const items = purchaseItems(p);
    return items.map((it, j) => `
    <tr><td>${j === 0 ? i + 1 : ''}</td><td>${j === 0 ? escapeHtml(p.date) : ''}</td><td>${j === 0 ? escapeHtml(p.noNota || '') : ''}</td>
    <td>${j === 0 ? escapeHtml(p.supplierName || '') : ''}</td><td>${escapeHtml(it.materialName)}</td>
    <td>${it.qty} ${escapeHtml(it.unit || '')}</td><td>${formatRupiah(it.price)}</td><td>${formatRupiah(it.qty * it.price)}</td></tr>`).join('');
  }).join('');
  return `<h2>BELANJA HARIAN MATERIAL</h2>
    <table><thead><tr><th>No</th><th>Tanggal</th><th>No. Nota</th><th>Suplier</th><th>Material</th><th>Qty</th><th>Harga</th><th>Subtotal</th></tr></thead>
    <tbody>${body}</tbody>
    <tfoot><tr class="bold"><td colspan="7">TOTAL</td><td>${formatRupiah(total)}</td></tr></tfoot></table>`;
}

/* ---- Versi teks polos (untuk disalin ke clipboard, fallback saat unduh/print diblokir) ---- */

function buildPekerjaText(workers) {
  const lines = ['DATA PEKERJA', '================================'];
  workers.forEach((w, i) => {
    lines.push(`${i + 1}. ${w.name}`);
    if (w.position) lines.push(`   ${w.position}`);
    if (w.address) lines.push(`   ${w.address}`);
    if (w.phone) lines.push(`   HP: ${w.phone}`);
    lines.push(`   Harian: ${formatRupiah(w.upahHarian)}`);
    lines.push(`   Lembur/jam: ${formatRupiah(w.upahLembur)}`);
    lines.push('--------------------------------');
  });
  lines.push(`Total: ${workers.length} pekerja`);
  return lines.join('\n');
}

function buildAbsensiText(activeWeek, project, workers, grandTotal, dates) {
  const dateKeys = dates.map((d) => d.key);
  const lines = [
    'ABSENSI UPAH HARIAN DAN LEMBURAN TUKANG',
    `PROYEK: ${project?.name || ''}`,
    project?.address ? `Alamat: ${project.address}` : '',
    `${activeWeek.weekLabel}  (${activeWeek.startDate} s/d ${activeWeek.endDate})`,
    '================================',
  ].filter(Boolean);
  workers.forEach((w, i) => {
    const rec = activeWeek.records[w.id] || {};
    const totals = calcWorkerWeekTotals(w, rec, dateKeys);
    const days = dates.map((d) => `${d.shortLabel}:${statusChar(rec[d.key])}`).join(' ');
    lines.push(`${i + 1}. ${w.name} (${w.position || '-'})`);
    lines.push(`   ${days}`);
    lines.push(`   Hari: ${totals.totalHariBayar}  Lembur: ${totals.totalJamLembur}j`);
    lines.push(`   Upah: ${formatRupiah(totals.totalUpah)}`);
    lines.push('--------------------------------');
  });
  lines.push(`JUMLAH TOTAL UPAH: ${formatRupiah(grandTotal)}`);
  return lines.join('\n');
}

function buildRekapText(rows, grand) {
  const lines = ['REKAP UPAH TUKANG - SELURUH PERIODE', '================================'];
  rows.forEach(({ worker, totalUpah, diterima, sisa }, i) => {
    lines.push(`${i + 1}. ${worker.name}`);
    lines.push(`   Total Upah : ${formatRupiah(totalUpah)}`);
    lines.push(`   Diterima   : ${formatRupiah(diterima)}`);
    lines.push(`   Sisa       : ${formatRupiah(sisa)}`);
    lines.push('--------------------------------');
  });
  lines.push(`TOTAL Upah    : ${formatRupiah(grand.totalUpah)}`);
  lines.push(`TOTAL Diterima: ${formatRupiah(grand.diterima)}`);
  lines.push(`TOTAL Sisa    : ${formatRupiah(grand.sisa)}`);
  return lines.join('\n');
}

function buildSlipGajiText(row, filterLabel) {
  const { worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown, evidenceFilled, evidenceExpected } = row;
  const lines = [
    'SLIP GAJI', '================================',
    `Nama       : ${worker.name}`,
    `Jabatan    : ${worker.position || '-'}`,
    `Periode    : ${filterLabel}`,
    `Tanggal    : ${new Date().toISOString().slice(0, 10)}`,
    '================================',
    'RINCIAN KERJA',
  ];
  weekBreakdown.forEach((wk) => {
    lines.push(`${wk.weekLabel} (${wk.projectName})`);
    lines.push(`  ${wk.startDate} s/d ${wk.endDate}`);
    lines.push(`  ${wk.totalHariBayar} hari, ${wk.totalJamLembur} jam lembur = ${formatRupiah(wk.totalUpah)}`);
  });
  lines.push('--------------------------------');
  lines.push(`Total Hari Kerja   : ${totalHari}`);
  lines.push(`Total Jam Lembur   : ${totalJamLembur}`);
  if (evidenceExpected > 0) lines.push(`Bukti Foto Kehadiran: ${evidenceFilled}/${evidenceExpected} periode terisi`);
  lines.push(`TOTAL UPAH         : ${formatRupiah(totalUpah)}`);
  lines.push('');
  lines.push('POTONGAN & PEMBAYARAN');
  workerKasbon.forEach((k) => lines.push(`  Kasbon ${k.date}: ${formatRupiah(k.amount)}${k.note ? ' (' + k.note + ')' : ''}`));
  workerKasbonPayments.forEach((p) => lines.push(`  Bayar kasbon ${p.date}: -${formatRupiah(p.amount)}${p.note ? ' (' + p.note + ')' : ''}`));
  lines.push(`  Sisa Kasbon Belum Lunas: ${formatRupiah(totalKasbon)}`);
  lines.push('');
  workerPayments.forEach((p) => lines.push(`  Upah dibayar ${p.date}: ${formatRupiah(p.amount)}`));
  lines.push(`  Total Diterima: ${formatRupiah(diterima)}`);
  lines.push('================================');
  lines.push(`SISA UPAH BELUM DIBAYAR: ${formatRupiah(sisa)}`);
  lines.push('================================');
  lines.push('', 'Diterima oleh,', '', '', `(${worker.name})`);
  return lines.join('\n');
}

function buildSlipGajiHtml(row, filterLabel, thermal) {
  const { worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown, evidenceFilled, evidenceExpected } = row;
  const workRows = weekBreakdown.map((wk) => `<tr><td>${escapeHtml(wk.weekLabel)}</td><td>${escapeHtml(wk.projectName)}</td><td>${escapeHtml(wk.startDate)} s/d ${escapeHtml(wk.endDate)}</td><td>${wk.totalHariBayar}</td><td>${wk.totalJamLembur}</td><td>${formatRupiah(wk.totalUpah)}</td></tr>`).join('');
  return `<h2>SLIP GAJI</h2>
    <p><b>${escapeHtml(worker.name)}</b> &nbsp; ${escapeHtml(worker.position || '-')}</p>
    <p>Periode: ${escapeHtml(filterLabel)} &nbsp; Tanggal cetak: ${new Date().toISOString().slice(0, 10)}</p>
    <table><thead><tr><th>Periode</th><th>Proyek</th><th>Tanggal</th><th>Hari</th><th>Jam Lembur</th><th>Upah</th></tr></thead>
    <tbody>${workRows}</tbody>
    <tfoot><tr class="bold"><td colspan="3">TOTAL</td><td>${totalHari}</td><td>${totalJamLembur}</td><td>${formatRupiah(totalUpah)}</td></tr></tfoot></table>
    ${evidenceExpected > 0 ? `<p style="margin-top:6px;">Bukti Foto Kehadiran: <b>${evidenceFilled}/${evidenceExpected}</b> periode terisi</p>` : ''}
    <p style="margin-top:10px;"><b>Kasbon &amp; Pembayaran</b></p>
    <table><tbody>
      ${workerKasbon.map((k) => `<tr><td>Kasbon ${escapeHtml(k.date)}</td><td>${formatRupiah(k.amount)}</td></tr>`).join('')}
      ${workerKasbonPayments.map((p) => `<tr><td>Bayar kasbon ${escapeHtml(p.date)}</td><td>-${formatRupiah(p.amount)}</td></tr>`).join('')}
      <tr class="bold"><td>Sisa Kasbon Belum Lunas</td><td>${formatRupiah(totalKasbon)}</td></tr>
      ${workerPayments.map((p) => `<tr><td>Upah dibayar ${escapeHtml(p.date)}</td><td>${formatRupiah(p.amount)}</td></tr>`).join('')}
      <tr class="bold"><td>Total Diterima</td><td>${formatRupiah(diterima)}</td></tr>
    </tbody></table>
    <p class="bold" style="margin-top:10px;font-size:${thermal ? '11px' : '14px'};">SISA UPAH BELUM DIBAYAR: ${formatRupiah(sisa)}</p>
    <p style="margin-top:24px;">Diterima oleh,</p>
    <p style="margin-top:36px;">(${escapeHtml(worker.name)})</p>`;
}

function TextPreviewModal({ title, text, onClose }) {
  const fullText = companyHeaderText() + text;
  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(33,29,26,0.7)' }}>
      <div className="w-full max-w-sm rounded-xl p-4 att-body att-punch-anim" style={{ background: THEME.paper }}>
        <div className="flex items-center justify-between mb-2">
          <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>{title}</h3>
          <button type="button" onClick={onClose}><X size={18} color={THEME.inkSoft} /></button>
        </div>
        <p className="att-mono text-[10px] mb-2" style={{ color: THEME.inkSoft }}>
          Tekan &amp; tahan teks di bawah ini, lalu pilih "Pilih Semua" / "Select All" &rarr; "Salin" / "Copy". Setelah itu tempel ke aplikasi catatan atau printer thermal kamu.
        </p>
        <textarea
          readOnly
          value={fullText}
          onFocus={(e) => e.target.select()}
          rows={12}
          className="w-full p-2 rounded att-mono text-xs outline-none"
          style={{ background: THEME.concrete, color: THEME.ink, border: `1px solid ${THEME.line}` }}
        />
      </div>
    </div>
  );
}

function ManageUsersModal({ users, currentUsername, onAdd, onSetPassword, onDelete, onClose }) {
  const [newUsername, setNewUsername] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const [pwFor, setPwFor] = useState(null);
  const [pwValue, setPwValue] = useState('');

  const handleAdd = async () => {
    setError('');
    if (!newUsername.trim() || newPassword.length < 4) {
      setError('Username wajib diisi, password minimal 4 karakter.');
      return;
    }
    setBusy(true);
    const err = await onAdd(newUsername.trim(), newPassword);
    setBusy(false);
    if (err) { setError(err); return; }
    setNewUsername('');
    setNewPassword('');
  };

  const submitPassword = async (userId) => {
    if (pwValue.length < 4) return;
    await onSetPassword(userId, pwValue);
    setPwFor(null);
    setPwValue('');
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(33,29,26,0.7)' }}>
      <div className="w-full max-w-sm rounded-xl p-4 att-body att-punch-anim max-h-[85vh] overflow-y-auto" style={{ background: THEME.paper }}>
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>Kelola User</h3>
          <button type="button" onClick={onClose}><X size={18} color={THEME.inkSoft} /></button>
        </div>
        <p className="att-mono text-[10px] mb-3" style={{ color: THEME.inkSoft }}>
          Semua user punya akses admin penuh yang sama. Cocok untuk beberapa admin/staff kantor yang saling percaya.
        </p>

        <div className="space-y-2 mb-4">
          {users.map((u) => (
            <div key={u.id} className="p-2.5 rounded" style={{ background: THEME.concrete }}>
              <div className="flex items-center justify-between">
                <span className="att-body text-sm font-semibold" style={{ color: THEME.ink }}>
                  {u.username}{u.username === currentUsername ? ' (kamu)' : ''}
                </span>
                <div className="flex gap-1">
                  <button type="button" onClick={() => { setPwFor(pwFor === u.id ? null : u.id); setPwValue(''); }}
                    className="att-mono text-[10px] px-2 py-1 rounded" style={{ border: `1px solid ${THEME.line}`, color: THEME.inkSoft }}>
                    Ganti Password
                  </button>
                  {users.length > 1 && (
                    <button type="button" onClick={() => onDelete(u.id)} className="p-1.5 rounded" style={{ background: THEME.paper }}>
                      <Trash2 size={13} color={THEME.rust} />
                    </button>
                  )}
                </div>
              </div>
              {pwFor === u.id && (
                <div className="flex gap-1.5 mt-2">
                  <input type="password" value={pwValue} onChange={(e) => setPwValue(e.target.value)} placeholder="Password baru (min. 4 karakter)"
                    className="flex-1 px-2 py-1.5 rounded att-mono text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }} />
                  <button type="button" onClick={() => submitPassword(u.id)} className="px-2.5 rounded" style={{ background: THEME.amber }}>
                    <Check size={13} color={THEME.charcoal} />
                  </button>
                </div>
              )}
            </div>
          ))}
        </div>

        <p className="att-body font-semibold text-sm mb-2" style={{ color: THEME.ink }}>Tambah User Baru</p>
        <div className="space-y-2">
          <input value={newUsername} onChange={(e) => setNewUsername(e.target.value)} placeholder="Username"
            className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input type="password" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} placeholder="Password (min. 4 karakter)"
            className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>
        {error && <p className="text-sm mt-2" style={{ color: THEME.rust }}>{error}</p>}
        <button type="button" onClick={handleAdd} disabled={busy}
          className="w-full mt-3 flex items-center justify-center gap-1.5 py-2 rounded att-body font-semibold text-sm disabled:opacity-60"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          {busy ? <Loader2 size={15} className="animate-spin" /> : <Plus size={15} />} Tambah User
        </button>
      </div>
    </div>
  );
}

function CompanyProfileModal({ company, onSave, onClose }) {
  const [name, setName] = useState(company.name || '');
  const [tagline, setTagline] = useState(company.tagline || '');
  const [address, setAddress] = useState(company.address || '');
  const [phone, setPhone] = useState(company.phone || '');
  const [website, setWebsite] = useState(company.website || '');
  const [email, setEmail] = useState(company.email || '');
  const [services, setServices] = useState(company.services || DEFAULT_COMPANY_SERVICES);
  const [materialsNote, setMaterialsNote] = useState(company.materialsNote || DEFAULT_COMPANY_MATERIALS_NOTE);
  const [logoDataUri, setLogoDataUri] = useState(company.logoDataUri || '');
  const [logoBusy, setLogoBusy] = useState(false);

  const handleLogoFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLogoBusy(true);
    try {
      const dataUri = await resizeImage(file, 200, 0.7); // logo: kecil saja, cukup untuk kop dokumen
      setLogoDataUri(dataUri);
    } catch {
      alert('Gagal memuat logo. Coba foto/gambar lain.');
    }
    setLogoBusy(false);
    e.target.value = '';
  };

  const handleSave = async () => {
    await onSave({
      name: name.trim(), tagline: tagline.trim(), address: address.trim(), phone: phone.trim(),
      website: website.trim(), email: email.trim(), services: services.trim(), materialsNote: materialsNote.trim(),
      logoDataUri,
    });
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(33,29,26,0.7)' }}>
      <div className="w-full max-w-sm rounded-xl p-4 att-body att-punch-anim max-h-[90vh] overflow-y-auto" style={{ background: THEME.paper }}>
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>Profil Perusahaan</h3>
          <button type="button" onClick={onClose}><X size={18} color={THEME.inkSoft} /></button>
        </div>
        <p className="att-mono text-[10px] mb-3" style={{ color: THEME.inkSoft }}>
          Muncul di judul aplikasi dan sebagai kop di setiap dokumen cetak/struk/PDF/ID card.
        </p>

        <div className="flex items-center gap-3 mb-3">
          <div className="w-16 h-16 rounded-lg overflow-hidden flex items-center justify-center shrink-0" style={{ background: THEME.concrete, border: `1px solid ${THEME.line}` }}>
            {logoDataUri
              ? <img src={logoDataUri} alt="Logo" className="w-full h-full object-cover" />
              : <img src={COMPANY_LOGO_DATA_URI} alt="Logo bawaan" className="w-full h-full object-cover opacity-40" />}
          </div>
          <div className="flex-1">
            <label className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded att-mono text-xs font-semibold cursor-pointer"
              style={{ border: `1px solid ${THEME.amber}`, color: THEME.charcoal, background: THEME.amber }}>
              {logoBusy ? 'Memuat...' : 'Ganti Logo'}
              <input type="file" accept="image/*" className="hidden" onChange={handleLogoFile} disabled={logoBusy} />
            </label>
            {logoDataUri && (
              <button type="button" onClick={() => setLogoDataUri('')} className="block mt-1 att-mono text-[10px]" style={{ color: THEME.rust }}>Hapus logo, pakai bawaan</button>
            )}
          </div>
        </div>

        <div className="space-y-2">
          <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Nama perusahaan"
            className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={tagline} onChange={(e) => setTagline(e.target.value)} placeholder="Tagline / jenis usaha (opsional)"
            className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={address} onChange={(e) => setAddress(e.target.value)} placeholder="Alamat"
            className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="No. Telepon"
            className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={website} onChange={(e) => setWebsite(e.target.value)} placeholder="Website (contoh: pattimurautama.com)"
            className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={email} onChange={(e) => setEmail(e.target.value)} placeholder="Email perusahaan"
            className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <div>
            <p className="att-mono text-[10px] mb-1" style={{ color: THEME.inkSoft }}>Layanan (1 baris = 1 layanan)</p>
            <textarea value={services} onChange={(e) => setServices(e.target.value)} rows={7}
              className="w-full px-3 py-2 rounded att-body text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          </div>
          <div>
            <p className="att-mono text-[10px] mb-1" style={{ color: THEME.inkSoft }}>Catatan tambahan (opsional)</p>
            <textarea value={materialsNote} onChange={(e) => setMaterialsNote(e.target.value)} rows={2}
              className="w-full px-3 py-2 rounded att-body text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          </div>
        </div>
        <button type="button" onClick={handleSave}
          className="w-full mt-3 flex items-center justify-center gap-1.5 py-2 rounded att-body font-semibold text-sm"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Check size={15} /> Simpan Profil
        </button>
      </div>
    </div>
  );
}

function BackupModal({ data, onImport, onClose }) {
  const [mode, setMode] = useState('export'); // 'export' | 'import'
  const [importText, setImportText] = useState('');
  const [importError, setImportError] = useState('');
  const exportText = JSON.stringify(data, null, 0);

  const handleImport = () => {
    setImportError('');
    try {
      const parsed = JSON.parse(importText);
      if (!parsed || typeof parsed !== 'object') throw new Error('Format tidak valid');
      onImport(parsed);
      onClose();
    } catch (err) {
      setImportError('Teks yang ditempel bukan cadangan yang valid. Pastikan disalin utuh dari tombol Backup sebelumnya.');
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(33,29,26,0.7)' }}>
      <div className="w-full max-w-sm rounded-xl p-4 att-body att-punch-anim" style={{ background: THEME.paper }}>
        <div className="flex items-center justify-between mb-2">
          <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>Backup Data</h3>
          <button type="button" onClick={onClose}><X size={18} color={THEME.inkSoft} /></button>
        </div>

        <div className="flex gap-2 mb-3">
          <button type="button" onClick={() => setMode('export')}
            className="flex-1 py-1.5 rounded att-mono text-xs font-semibold"
            style={{ background: mode === 'export' ? THEME.amber : THEME.concrete, color: THEME.charcoal }}>
            Ekspor
          </button>
          <button type="button" onClick={() => setMode('import')}
            className="flex-1 py-1.5 rounded att-mono text-xs font-semibold"
            style={{ background: mode === 'import' ? THEME.amber : THEME.concrete, color: THEME.charcoal }}>
            Impor
          </button>
        </div>

        {mode === 'export' ? (
          <>
            <p className="att-mono text-[10px] mb-2" style={{ color: THEME.inkSoft }}>
              Penyimpanan otomatis di preview ini sedang tidak bisa diandalkan. Tekan &amp; tahan teks di bawah &rarr; "Pilih Semua" &rarr; "Salin", lalu tempel ke Notes/WhatsApp sebagai cadangan sebelum menutup halaman.
            </p>
            <textarea
              readOnly
              value={exportText}
              onFocus={(e) => e.target.select()}
              rows={10}
              className="w-full p-2 rounded att-mono text-xs outline-none"
              style={{ background: THEME.concrete, color: THEME.ink, border: `1px solid ${THEME.line}` }}
            />
          </>
        ) : (
          <>
            <p className="att-mono text-[10px] mb-2" style={{ color: THEME.inkSoft }}>
              Tempel di sini teks cadangan yang sebelumnya kamu salin, lalu tekan "Pulihkan Data". Ini akan menimpa data yang sedang tampil saat ini.
            </p>
            <textarea
              value={importText}
              onChange={(e) => setImportText(e.target.value)}
              rows={10}
              placeholder="Tempel teks cadangan di sini..."
              className="w-full p-2 rounded att-mono text-xs outline-none"
              style={{ background: THEME.concrete, color: THEME.ink, border: `1px solid ${THEME.line}` }}
            />
            {importError && <p className="text-xs mt-2" style={{ color: THEME.rust }}>{importError}</p>}
            <button type="button" onClick={handleImport}
              className="w-full mt-3 py-2 rounded att-body font-semibold text-sm"
              style={{ background: THEME.amber, color: THEME.charcoal }}>
              Pulihkan Data
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function PrintMenu({ onCopy }) {
  return (
    <button
      type="button"
      onClick={onCopy}
      className="flex items-center gap-1.5 px-3 py-1.5 rounded att-mono text-xs font-semibold"
      style={{ border: `1px solid ${THEME.amber}`, color: THEME.charcoal, background: THEME.amber }}
    >
      <Copy size={13} /> Lihat &amp; Salin Teks
    </button>
  );
}

function TopBar({ onLogout, onBackup, onProfile, onManageUsers, company, username }) {
  return (
    <div className="flex items-center justify-between px-5 py-4 sticky top-0 z-10" style={{ background: THEME.charcoal }}>
      <div className="flex items-center gap-2 min-w-0">
        <img src={company?.logoDataUri || COMPANY_LOGO_DATA_URI} alt="Logo" className="w-9 h-9 rounded object-cover shrink-0" style={{ border: `1px solid ${THEME.line}` }} />
        <div className="min-w-0">
          <p className="att-display text-sm leading-none truncate" style={{ color: THEME.paper }}>{company?.name || 'ABSENSI TUKANG'}</p>
          <p className="att-mono text-[10px] mt-0.5" style={{ color: '#9C948A' }}>{username ? `Login: ${username}` : 'ADMIN'}</p>
        </div>
      </div>
      <div className="flex items-center gap-2 shrink-0 flex-wrap justify-end">
        <button onClick={onManageUsers} className="att-mono text-xs px-3 py-1.5 rounded" style={{ color: THEME.paper, border: `1px solid ${THEME.line}` }}>
          User
        </button>
        <button onClick={onProfile} className="att-mono text-xs px-3 py-1.5 rounded" style={{ color: THEME.paper, border: `1px solid ${THEME.line}` }}>
          Profil
        </button>
        <button onClick={onBackup} className="att-mono text-xs px-3 py-1.5 rounded flex items-center gap-1.5" style={{ color: THEME.charcoal, background: THEME.amber }}>
          <Save size={13} /> Backup
        </button>
        <button onClick={onLogout} className="att-mono text-xs px-3 py-1.5 rounded" style={{ color: THEME.paper, border: `1px solid ${THEME.line}` }}>
          Keluar
        </button>
      </div>
    </div>
  );
}

function LoginScreen({ onLogin, error, busy }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  return (
    <div className="min-h-screen flex items-center justify-center p-6 att-body" style={{ background: THEME.charcoal }}>
      <GlobalStyle />
      <div className="w-full max-w-xs">
        <div className="text-center mb-8">
          <img src={COMPANY_LOGO_DATA_URI} alt="Logo" className="w-16 h-16 rounded-lg mx-auto mb-3 object-cover" style={{ border: `1px solid ${THEME.line}` }} />
          <h1 className="att-display text-xl" style={{ color: THEME.paper }}>ABSENSI TUKANG</h1>
          <p className="att-mono text-xs mt-1" style={{ color: '#9C948A' }}>Login Admin</p>
        </div>
        <div className="space-y-3">
          <div>
            <label className="att-mono text-xs block mb-1" style={{ color: '#9C948A' }}>USERNAME</label>
            <input
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Enter') onLogin(username.trim(), password); }}
              className="w-full px-3 py-2 rounded outline-none att-body"
              style={{ background: THEME.paper, color: THEME.ink }}
              placeholder="admin"
              autoFocus
            />
          </div>
          <div>
            <label className="att-mono text-xs block mb-1" style={{ color: '#9C948A' }}>PASSWORD</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Enter') onLogin(username.trim(), password); }}
              className="w-full px-3 py-2 rounded outline-none att-body"
              style={{ background: THEME.paper, color: THEME.ink }}
              placeholder="••••••"
            />
          </div>
          {error && (
            <div className="flex items-center gap-2 text-sm px-3 py-2 rounded" style={{ background: '#3A241D', color: '#E8A08F' }}>
              <AlertTriangle size={16} />
              {error}
            </div>
          )}
          <button
            type="button"
            onClick={() => onLogin(username.trim(), password)}
            disabled={busy}
            className="w-full py-2.5 rounded att-body font-semibold flex items-center justify-center gap-2 disabled:opacity-60"
            style={{ background: THEME.amber, color: THEME.charcoal }}
          >
            {busy ? <Loader2 size={18} className="animate-spin" /> : <LogIn size={18} />}
            Masuk
          </button>
          <p className="att-mono text-[10px] text-center mt-2" style={{ color: '#78706660' }}>demo: admin / admin123</p>
        </div>
      </div>
    </div>
  );
}

function TabBar({ tab, setTab }) {
  const tabs = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'pekerja', label: 'Data Pekerja', icon: Users },
    { id: 'absensi', label: 'Absensi Mingguan', icon: CalendarDays },
    { id: 'kasbon', label: 'Kasbon', icon: Banknote },
    { id: 'belanja', label: 'Belanja', icon: ShoppingCart },
    { id: 'gudang', label: 'Gudang', icon: Warehouse },
    { id: 'rekap', label: 'Rekap Upah', icon: Wallet },
    { id: 'proyek', label: 'Rekap Proyek', icon: FileBarChart },
    { id: 'progres', label: 'Progres Proyek', icon: Camera },
    { id: 'penawaran', label: 'Penawaran', icon: Receipt },
    { id: 'klien', label: 'Data Klien', icon: Users },
    { id: 'ahsp', label: 'AHSP Master', icon: FileBarChart },
    { id: 'klien', label: 'Data Klien', icon: Users },
  ];
  return (
    <div className="flex gap-1 px-3 pt-3 pb-2 att-scroll overflow-x-auto" style={{ background: THEME.charcoal }}>
      {tabs.map((t) => {
        const Icon = t.icon;
        const active = tab === t.id;
        return (
          <button
            key={t.id}
            onClick={() => setTab(t.id)}
            className="flex items-center gap-1.5 px-3 py-2 rounded-t-lg att-body text-xs font-semibold whitespace-nowrap"
            style={{
              background: active ? THEME.concrete : 'transparent',
              color: active ? THEME.ink : '#9C948A',
            }}
          >
            <Icon size={14} /> {t.label}
          </button>
        );
      })}
    </div>
  );
}

/* ---------------- Data Pekerja ---------------- */

function emptyForm() {
  return { name: '', address: '', position: '', phone: '', ktp: '', upahHarian: '', upahLembur: '', joinDate: '', photoDataUri: '' };
}

function DataPekerjaTab({ workers, onAdd, onUpdate, onDelete }) {
  const [form, setForm] = useState(emptyForm());
  const [editingId, setEditingId] = useState(null);
  const [formError, setFormError] = useState('');
  const [saving, setSaving] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [showTextPreview, setShowTextPreview] = useState(false);
  const [photoBusy, setPhotoBusy] = useState(false);
  const [idCardFor, setIdCardFor] = useState(null); // worker sedang buka menu cetak ID card

  const handlePhotoFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setPhotoBusy(true);
    try {
      const dataUri = await resizeImage(file, 320, 0.6); // foto pekerja: cukup untuk ID card & daftar
      setForm((f) => ({ ...f, photoDataUri: dataUri }));
    } catch {
      alert('Gagal memuat foto. Coba foto lain.');
    }
    setPhotoBusy(false);
    e.target.value = '';
  };

  const startEdit = (w) => {
    setEditingId(w.id);
    setForm({
      name: w.name || '', address: w.address || '', position: w.position || '',
      phone: w.phone || '', ktp: w.ktp || '',
      upahHarian: String(w.upahHarian ?? ''), upahLembur: String(w.upahLembur ?? ''),
      joinDate: w.joinDate || '', photoDataUri: w.photoDataUri || '',
    });
  };

  const cancelEdit = () => {
    setEditingId(null);
    setForm(emptyForm());
    setFormError('');
  };

  const handleSave = async () => {
    setFormError('');
    if (!form.name || form.upahHarian === '' || form.upahLembur === '') {
      setFormError('Nama, upah harian, dan upah lembur/jam wajib diisi.');
      return;
    }
    setSaving(true);
    const payload = {
      name: form.name.trim(),
      address: form.address.trim(),
      position: form.position.trim(),
      phone: form.phone.trim(),
      ktp: form.ktp.trim(),
      upahHarian: Number(form.upahHarian) || 0,
      upahLembur: Number(form.upahLembur) || 0,
      joinDate: form.joinDate,
      photoDataUri: form.photoDataUri,
    };
    if (editingId) {
      await onUpdate(editingId, payload);
    } else {
      await onAdd({ id: uid(), ...payload });
    }
    setSaving(false);
    cancelEdit();
  };

  return (
    <div className="p-4 pb-24">
      <div className="flex items-center justify-between mb-3">
        <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{workers.length} PEKERJA TERDAFTAR</p>
        <PrintMenu
          onPrint={(thermal) => openPrintDocument('Data Pekerja', buildPekerjaHtml(workers, thermal), thermal)}
          onCopy={() => setShowTextPreview(true)}
        />
      </div>
      {showTextPreview && (
        <TextPreviewModal title="Data Pekerja" text={buildPekerjaText(workers)} onClose={() => setShowTextPreview(false)} />
      )}
      <div className="p-4 rounded-lg mb-5" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
        <p className="att-body font-semibold text-sm mb-3" style={{ color: THEME.ink }}>
          {editingId ? 'Edit Pekerja' : 'Tambah Pekerja Baru'}
        </p>
        <div className="flex items-center gap-3 mb-3">
          <div className="w-16 h-16 rounded-lg overflow-hidden flex items-center justify-center shrink-0" style={{ background: THEME.concrete, border: `1px solid ${THEME.line}` }}>
            {form.photoDataUri
              ? <img src={form.photoDataUri} alt="Foto pekerja" className="w-full h-full object-cover" />
              : <Users size={22} color={THEME.inkSoft} />}
          </div>
          <div className="flex-1">
            <label className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded att-mono text-xs font-semibold cursor-pointer"
              style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.concrete }}>
              {photoBusy ? 'Memuat...' : 'Pilih Foto'}
              <input type="file" accept="image/*" className="hidden" onChange={handlePhotoFile} disabled={photoBusy} />
            </label>
            {form.photoDataUri && (
              <button type="button" onClick={() => setForm((f) => ({ ...f, photoDataUri: '' }))} className="block mt-1 att-mono text-[10px]" style={{ color: THEME.rust }}>Hapus foto</button>
            )}
          </div>
        </div>
        <div className="grid sm:grid-cols-2 gap-3">
          <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })}
            placeholder="Nama lengkap" className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={form.position} onChange={(e) => setForm({ ...form, position: e.target.value })}
            placeholder="Jabatan (mis. Tukang, Kepala Tukang)" className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })}
            placeholder="Alamat" className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })}
            placeholder="No. HP" className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={form.ktp} onChange={(e) => setForm({ ...form, ktp: e.target.value })}
            placeholder="No. KTP" className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <div>
            <p className="att-mono text-[10px] mb-1" style={{ color: THEME.inkSoft }}>Tanggal Bergabung</p>
            <input type="date" value={form.joinDate} onChange={(e) => setForm({ ...form, joinDate: e.target.value })}
              className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          </div>
          <input value={form.upahHarian} onChange={(e) => setForm({ ...form, upahHarian: e.target.value.replace(/[^0-9]/g, '') })}
            placeholder="Upah harian (Rp)" inputMode="numeric" className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={form.upahLembur} onChange={(e) => setForm({ ...form, upahLembur: e.target.value.replace(/[^0-9]/g, '') })}
            placeholder="Upah lembur/jam (Rp)" inputMode="numeric" className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>
        {formError && <p className="text-sm mt-2" style={{ color: THEME.rust }}>{formError}</p>}
        <div className="flex gap-2 mt-3">
          <button type="button" onClick={handleSave} disabled={saving}
            className="flex-1 flex items-center justify-center gap-1.5 py-2 rounded att-body font-semibold text-sm disabled:opacity-60"
            style={{ background: THEME.amber, color: THEME.charcoal }}>
            {saving ? <Loader2 size={15} className="animate-spin" /> : (editingId ? <Check size={15} /> : <Plus size={15} />)}
            {editingId ? 'Simpan Perubahan' : 'Tambah Pekerja'}
          </button>
          {editingId && (
            <button type="button" onClick={cancelEdit} className="px-4 py-2 rounded att-body text-sm" style={{ border: `1px solid ${THEME.line}`, color: THEME.inkSoft }}>
              Batal
            </button>
          )}
        </div>
      </div>

      <div className="space-y-2">
        {workers.map((w) => (
          <div key={w.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div className="flex items-start justify-between gap-2">
              <div className="flex items-start gap-2.5 min-w-0">
                <div className="w-11 h-11 rounded-lg overflow-hidden shrink-0 flex items-center justify-center" style={{ background: THEME.concrete, border: `1px solid ${THEME.line}` }}>
                  {w.photoDataUri
                    ? <img src={w.photoDataUri} alt={w.name} className="w-full h-full object-cover" />
                    : <Users size={18} color={THEME.inkSoft} />}
                </div>
                <div className="min-w-0">
                  <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{w.name}</p>
                  <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{w.position || '-'}</p>
                  <p className="att-mono text-xs mt-1" style={{ color: THEME.inkSoft }}>
                    Harian {formatRupiah(w.upahHarian)} &middot; Lembur {formatRupiah(w.upahLembur)}/jam
                  </p>
                  {w.joinDate && (
                    <p className="att-mono text-[11px] mt-0.5" style={{ color: THEME.inkSoft }}>Bergabung: {w.joinDate}</p>
                  )}
                </div>
              </div>
              <div className="flex gap-1 shrink-0">
                <button type="button" onClick={() => startEdit(w)} className="p-2 rounded" style={{ background: THEME.concrete }}>
                  <Edit3 size={14} color={THEME.inkSoft} />
                </button>
                {confirmDelete === w.id ? (
                  <button type="button" onClick={() => { onDelete(w.id); setConfirmDelete(null); }} className="p-2 rounded" style={{ background: THEME.rust }}>
                    <Check size={14} color={THEME.paper} />
                  </button>
                ) : (
                  <button type="button" onClick={() => setConfirmDelete(w.id)} className="p-2 rounded" style={{ background: THEME.concrete }}>
                    <Trash2 size={14} color={THEME.rust} />
                  </button>
                )}
              </div>
            </div>
            <div className="flex items-center gap-1.5 mt-2">
              <button type="button" onClick={() => setIdCardFor(idCardFor === w.id ? null : w.id)}
                className="flex items-center gap-1.5 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold"
                style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.concrete }}>
                <Receipt size={12} /> ID Card
              </button>
              {idCardFor === w.id && (
                <>
                  <button type="button" onClick={() => exportIdCardPDF(w)}
                    className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.paper }}>
                    <Save size={12} /> PDF
                  </button>
                  <button type="button" onClick={() => openIdCardPrint(w)}
                    className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.paper }}>
                    <Printer size={12} /> Cetak
                  </button>
                </>
              )}
            </div>
          </div>
        ))}
        {workers.length === 0 && (
          <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada pekerja. Tambahkan di atas.</p>
        )}
      </div>
    </div>
  );
}

/* ---------------- Evidence Modal (GPS + Foto) ---------------- */

const EVIDENCE_PERIODS = [
  { key: 'pagi', label: 'Pagi' },
  { key: 'siang', label: 'Siang' },
  { key: 'sore', label: 'Sore' },
  { key: 'lembur', label: 'Lembur' },
];

function nowTimeLabel() {
  const d = new Date();
  return `${d.toLocaleDateString('id-ID')} ${d.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })}`;
}

function EvidenceModal({ onSave, onCancel, existing }) {
  const [geoState, setGeoState] = useState('idle');
  const [coords, setCoords] = useState(existing ? { lat: existing.lat, lng: existing.lng } : null);
  const [periods, setPeriods] = useState(existing?.periods || {});
  const [activePeriod, setActivePeriod] = useState('pagi');
  const [saving, setSaving] = useState(false);
  const [busyPeriod, setBusyPeriod] = useState(null);

  const captureLocation = () => {
    setGeoState('loading');
    if (!navigator.geolocation) {
      setGeoState('error');
      return;
    }
    navigator.geolocation.getCurrentPosition(
      (pos) => {
        setCoords({ lat: pos.coords.latitude, lng: pos.coords.longitude });
        setGeoState('ok');
      },
      () => setGeoState('error'),
      { enableHighAccuracy: true, timeout: 8000 }
    );
  };

  const onFile = async (e, periodKey) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setBusyPeriod(periodKey);
    try {
      const dataUrl = await resizeImage(file);
      setPeriods((prev) => ({ ...prev, [periodKey]: { photo: dataUrl, time: nowTimeLabel() } }));
    } catch {
      alert('Gagal memuat foto. Coba lagi.');
    }
    setBusyPeriod(null);
    e.target.value = '';
  };

  const removePeriodPhoto = (periodKey) => {
    setPeriods((prev) => {
      const next = { ...prev };
      delete next[periodKey];
      return next;
    });
  };

  const handleSave = async () => {
    setSaving(true);
    await onSave({ lat: coords?.lat ?? null, lng: coords?.lng ?? null, periods });
    setSaving(false);
  };

  const filledCount = EVIDENCE_PERIODS.filter((p) => periods[p.key]?.photo).length;

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(33,29,26,0.7)' }}>
      <div className="w-full max-w-sm rounded-xl p-5 att-punch-anim att-body max-h-[92vh] overflow-y-auto" style={{ background: THEME.paper }}>
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>Bukti Kehadiran ({filledCount}/4)</h3>
          <button type="button" onClick={onCancel}><X size={18} color={THEME.inkSoft} /></button>
        </div>

        {coords ? (
          <div className="flex items-center gap-1.5 mb-3">
            <a href={`https://www.google.com/maps?q=${coords.lat},${coords.lng}`} target="_blank" rel="noopener noreferrer"
              className="flex-1 flex items-center justify-center gap-2 py-2.5 rounded text-sm font-semibold"
              style={{ background: THEME.concrete, color: THEME.ink }}>
              <MapPin size={16} /> {coords.lat.toFixed(4)}, {coords.lng.toFixed(4)} — Buka Maps
            </a>
            <button type="button" onClick={captureLocation} className="px-3 py-2.5 rounded shrink-0" style={{ background: THEME.concrete }}>
              <RefreshCw size={16} color={THEME.inkSoft} />
            </button>
          </div>
        ) : (
          <button type="button" onClick={captureLocation}
            className="w-full flex items-center justify-center gap-2 py-2.5 rounded mb-3 text-sm font-semibold"
            style={{ background: THEME.concrete, color: THEME.ink }}>
            <MapPin size={16} />
            {geoState === 'loading' ? 'Mengambil lokasi...' : 'Ambil Lokasi GPS (sekali untuk hari ini)'}
          </button>
        )}
        {geoState === 'error' && <p className="text-xs mb-3" style={{ color: THEME.rust }}>Gagal mengambil lokasi. Coba lagi atau lewati.</p>}

        <div className="grid grid-cols-4 gap-1.5 mb-3">
          {EVIDENCE_PERIODS.map((p) => (
            <button key={p.key} type="button" onClick={() => setActivePeriod(p.key)}
              className="py-2 rounded att-mono text-[11px] font-semibold relative"
              style={{
                background: activePeriod === p.key ? THEME.amber : THEME.concrete,
                color: activePeriod === p.key ? THEME.charcoal : THEME.inkSoft,
              }}>
              {p.label}
              {periods[p.key]?.photo && (
                <span className="absolute -top-1 -right-1 w-3 h-3 rounded-full" style={{ background: THEME.green }} />
              )}
            </button>
          ))}
        </div>

        {EVIDENCE_PERIODS.filter((p) => p.key === activePeriod).map((p) => (
          <div key={p.key}>
            {periods[p.key]?.photo ? (
              <div className="mb-3">
                <img src={periods[p.key].photo} alt={p.label} className="w-full rounded-lg mb-1.5" style={{ maxHeight: 180, objectFit: 'cover' }} />
                <div className="flex items-center justify-between">
                  <p className="att-mono text-[10px]" style={{ color: THEME.inkSoft }}>Diambil: {periods[p.key].time}</p>
                  <button type="button" onClick={() => removePeriodPhoto(p.key)} className="att-mono text-[10px]" style={{ color: THEME.rust }}>Hapus</button>
                </div>
              </div>
            ) : (
              <div className="flex gap-2 mb-3">
                <label className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded text-sm font-semibold cursor-pointer"
                  style={{ background: THEME.concrete, color: THEME.ink }}>
                  <Camera size={16} />
                  {busyPeriod === p.key ? '...' : 'Kamera'}
                  <input type="file" accept="image/*" capture="environment" className="hidden" onChange={(e) => onFile(e, p.key)} disabled={busyPeriod === p.key} />
                </label>
                <label className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded text-sm font-semibold cursor-pointer"
                  style={{ background: THEME.concrete, color: THEME.ink }}>
                  <Save size={16} />
                  {busyPeriod === p.key ? '...' : 'File'}
                  <input type="file" accept="image/*" className="hidden" onChange={(e) => onFile(e, p.key)} disabled={busyPeriod === p.key} />
                </label>
              </div>
            )}
          </div>
        ))}

        <div className="flex gap-2">
          <button type="button" onClick={handleSave} disabled={saving}
            className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded font-semibold text-sm disabled:opacity-60"
            style={{ background: THEME.amber, color: THEME.charcoal }}>
            {saving ? <Loader2 size={15} className="animate-spin" /> : <Check size={15} />}
            Simpan Bukti
          </button>
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded text-sm" style={{ border: `1px solid ${THEME.line}`, color: THEME.inkSoft }}>
            Tutup
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------------- Absensi Mingguan ---------------- */

function WorkerWeekCard({ worker, weekRecord, dates, onSetStatus, onSetLembur, onOpenEvidence, evidenceMap }) {
  const [expanded, setExpanded] = useState(true);
  const [page, setPage] = useState(0);
  const totals = calcWorkerWeekTotals(worker, weekRecord, dates.map((d) => d.key));
  const STATUS_CYCLE = ['off', 'full', 'half'];
  const STATUS_LABEL = { off: 'Tidak Hadir', full: 'Hadir Penuh', half: '½ Hari' };
  const STATUS_COLOR = { off: THEME.paper, full: THEME.amber, half: THEME.amberSoft };

  const perPage = 7;
  const totalPages = Math.max(1, Math.ceil(dates.length / perPage));
  const pageDates = dates.slice(page * perPage, page * perPage + perPage);

  return (
    <div className="rounded-lg mb-3" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
      <button type="button" onClick={() => setExpanded(!expanded)} className="w-full flex items-center justify-between p-3">
        <div className="text-left">
          <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{worker.name}</p>
          <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>
            {totals.totalHariBayar} hari &middot; {totals.totalJamLembur} jam lembur &middot; {formatRupiah(totals.totalUpah)}
          </p>
        </div>
        {expanded ? <ChevronUp size={16} color={THEME.inkSoft} /> : <ChevronDown size={16} color={THEME.inkSoft} />}
      </button>

      {expanded && (
        <div className="px-3 pb-3">
          {totalPages > 1 && (
            <div className="flex items-center justify-between mb-2">
              <button type="button" onClick={() => setPage(Math.max(0, page - 1))} disabled={page === 0}
                className="att-mono text-[10px] px-2 py-1 rounded disabled:opacity-30" style={{ border: `1px solid ${THEME.line}`, color: THEME.inkSoft }}>
                &larr; Sebelumnya
              </button>
              <span className="att-mono text-[10px]" style={{ color: THEME.inkSoft }}>Hari {page * perPage + 1}-{Math.min(dates.length, page * perPage + perPage)} dari {dates.length}</span>
              <button type="button" onClick={() => setPage(Math.min(totalPages - 1, page + 1))} disabled={page >= totalPages - 1}
                className="att-mono text-[10px] px-2 py-1 rounded disabled:opacity-30" style={{ border: `1px solid ${THEME.line}`, color: THEME.inkSoft }}>
                Selanjutnya &rarr;
              </button>
            </div>
          )}
          <div className="grid grid-cols-2 gap-2 mb-2">
            {pageDates.map((d) => {
              const rec = (weekRecord && weekRecord[d.key]) || {};
              const status = dayStatus(rec);
              const hasEvidence = !!(evidenceMap && evidenceMap[d.key] && Object.values(evidenceMap[d.key].periods || {}).some((p) => p?.photo));
              return (
                <div key={d.key} className="p-2 rounded" style={{ background: THEME.concrete }}>
                  <div className="flex items-center justify-between mb-1">
                    <span className="att-mono text-[9px] leading-tight" style={{ color: THEME.inkSoft }}>{d.shortLabel}</span>
                    <button
                      type="button"
                      onClick={() => onOpenEvidence(d.key)}
                      disabled={status === 'off'}
                      className="p-1 rounded disabled:opacity-30"
                      style={{ background: hasEvidence ? THEME.green : 'transparent' }}
                    >
                      <Camera size={12} color={hasEvidence ? THEME.paper : THEME.inkSoft} />
                    </button>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      const next = STATUS_CYCLE[(STATUS_CYCLE.indexOf(status) + 1) % STATUS_CYCLE.length];
                      onSetStatus(d.key, next);
                    }}
                    className="w-full py-1.5 rounded att-body text-xs font-semibold mb-1"
                    style={{ background: STATUS_COLOR[status], color: status === 'off' ? THEME.inkSoft : THEME.charcoal, border: `1px solid ${THEME.line}` }}
                  >
                    {STATUS_LABEL[status]}
                  </button>
                  <select
                    value={rec.lembur || 0}
                    disabled={status === 'off'}
                    onChange={(e) => onSetLembur(d.key, Number(e.target.value))}
                    className="w-full py-1 rounded att-mono text-xs outline-none disabled:opacity-40"
                    style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }}
                  >
                    {[...Array(11)].map((_, i) => (
                      <option key={i} value={i}>{i} jam lembur</option>
                    ))}
                  </select>
                </div>
              );
            })}
          </div>
          <div className="grid grid-cols-3 gap-2 att-mono text-[10px]" style={{ color: THEME.inkSoft }}>
            <div className="p-2 rounded text-center" style={{ background: THEME.concrete }}>
              <p style={{ color: THEME.ink }} className="text-sm font-semibold">{totals.totalHariBayar}</p>
              Total Hari
            </div>
            <div className="p-2 rounded text-center" style={{ background: THEME.concrete }}>
              <p style={{ color: THEME.ink }} className="text-sm font-semibold">{totals.totalJamLembur}</p>
              Jam Lembur
            </div>
            <div className="p-2 rounded text-center" style={{ background: THEME.concrete }}>
              <p style={{ color: THEME.ink }} className="text-sm font-semibold">{formatRupiah(totals.totalUpah)}</p>
              Total Upah
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function WeekForm({ projects, workers, onSubmit, onCancel, initial, onAddProject }) {
  const [projectId, setProjectId] = useState(initial?.projectId || '');
  const [newProjectName, setNewProjectName] = useState('');
  const [newProjectAddress, setNewProjectAddress] = useState('');
  const [addingProject, setAddingProject] = useState(projects.length === 0);
  const [weekLabel, setWeekLabel] = useState(initial?.weekLabel || '');
  const [startDate, setStartDate] = useState(initial?.startDate || '');
  const [endDate, setEndDate] = useState(initial?.endDate || '');
  const [workerIds, setWorkerIds] = useState(initial?.workerIds && initial.workerIds.length > 0 ? initial.workerIds : workers.map((w) => w.id));
  const [error, setError] = useState('');

  const toggleWorker = (id) => {
    setWorkerIds((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  };
  const allSelected = workers.length > 0 && workerIds.length === workers.length;

  const handleSubmit = async () => {
    setError('');
    let finalProjectId = projectId;
    if (addingProject || !projectId) {
      if (!newProjectName.trim()) {
        setError('Nama proyek wajib diisi.');
        return;
      }
      const newProject = { id: uid(), name: newProjectName.trim(), address: newProjectAddress.trim() };
      await onAddProject(newProject);
      finalProjectId = newProject.id;
    }
    if (!weekLabel) {
      setError('Label minggu wajib diisi.');
      return;
    }
    if (workerIds.length === 0) {
      setError('Pilih minimal 1 pekerja untuk minggu ini.');
      return;
    }
    onSubmit({ projectId: finalProjectId, weekLabel: weekLabel.trim(), startDate, endDate, workerIds });
  };

  return (
    <div className="p-4 rounded-lg mb-4" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
      <p className="att-body font-semibold text-sm mb-3" style={{ color: THEME.ink }}>{initial ? 'Edit Minggu Absensi' : 'Buat Minggu Absensi Baru'}</p>

      {!addingProject ? (
        <div className="flex gap-2 mb-3">
          <select value={projectId} onChange={(e) => setProjectId(e.target.value)}
            className="flex-1 px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }}>
            <option value="">Pilih proyek...</option>
            {projects.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
          <button type="button" onClick={() => setAddingProject(true)} className="px-3 py-2 rounded att-mono text-xs" style={{ background: THEME.concrete, color: THEME.ink }}>
            + Baru
          </button>
        </div>
      ) : (
        <div className="grid sm:grid-cols-2 gap-3 mb-3">
          <input value={newProjectName} onChange={(e) => setNewProjectName(e.target.value)} placeholder="Nama proyek baru"
            className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={newProjectAddress} onChange={(e) => setNewProjectAddress(e.target.value)} placeholder="Alamat proyek"
            className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
          {projects.length > 0 && (
            <button type="button" onClick={() => setAddingProject(false)} className="att-mono text-xs text-left sm:col-span-2" style={{ color: THEME.inkSoft }}>
              &larr; pilih proyek yang sudah ada
            </button>
          )}
        </div>
      )}

      <div className="grid sm:grid-cols-2 gap-3">
        <input value={weekLabel} onChange={(e) => setWeekLabel(e.target.value)} placeholder="mis. Minggu ke-1"
          className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        <div className="flex gap-2">
          <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)}
            className="flex-1 px-2 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)}
            className="flex-1 px-2 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>
      </div>

      <div className="mt-3">
        <div className="flex items-center justify-between mb-1.5">
          <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>PEKERJA IKUT MINGGU INI</p>
          <button type="button" onClick={() => setWorkerIds(allSelected ? [] : workers.map((w) => w.id))} className="att-mono text-xs" style={{ color: THEME.amber }}>
            {allSelected ? 'Kosongkan' : 'Pilih Semua'}
          </button>
        </div>
        <div className="flex flex-wrap gap-1.5">
          {workers.map((w) => {
            const checked = workerIds.includes(w.id);
            return (
              <button
                key={w.id}
                type="button"
                onClick={() => toggleWorker(w.id)}
                className="px-2.5 py-1 rounded-full att-mono text-xs"
                style={{ background: checked ? THEME.amber : THEME.concrete, color: checked ? THEME.charcoal : THEME.inkSoft }}
              >
                {checked ? '✓ ' : ''}{w.name}
              </button>
            );
          })}
          {workers.length === 0 && <p className="att-body text-xs" style={{ color: THEME.inkSoft }}>Belum ada pekerja terdaftar.</p>}
        </div>
      </div>

      {error && <p className="text-sm mt-2" style={{ color: THEME.rust }}>{error}</p>}
      <div className="flex gap-2 mt-3">
        <button type="button" onClick={handleSubmit}
          className="flex-1 flex items-center justify-center gap-1.5 py-2 rounded att-body font-semibold text-sm"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Plus size={15} /> {initial ? 'Simpan Perubahan' : 'Buat Minggu'}
        </button>
        {onCancel && (
          <button type="button" onClick={onCancel} className="px-4 py-2 rounded att-body text-sm" style={{ border: `1px solid ${THEME.line}`, color: THEME.inkSoft }}>
            Batal
          </button>
        )}
      </div>
    </div>
  );
}

function AbsensiMingguanTab({ workers, weeks, projects, onCreateWeek, onUpdateWeek, onAddProject, onUpdateWeekRecord, evidence, onSaveEvidence }) {
  const [activeWeekId, setActiveWeekId] = useState(weeks[0]?.id || null);
  const [showNewWeek, setShowNewWeek] = useState(weeks.length === 0);
  const [editingWeek, setEditingWeek] = useState(false);
  const [evidenceTarget, setEvidenceTarget] = useState(null); // { workerId, day }
  const [showTextPreview, setShowTextPreview] = useState(false);

  useEffect(() => {
    if (!activeWeekId && weeks.length > 0) setActiveWeekId(weeks[0].id);
  }, [weeks, activeWeekId]);

  const activeWeek = weeks.find((w) => w.id === activeWeekId) || null;
  const activeProject = activeWeek ? projects.find((p) => p.id === activeWeek.projectId) : null;
  const weekWorkers = activeWeek
    ? (activeWeek.workerIds && activeWeek.workerIds.length > 0
        ? workers.filter((w) => activeWeek.workerIds.includes(w.id))
        : workers)
    : [];
  const weekDates = activeWeek ? getWeekDates(activeWeek) : [];

  const handleCreate = (data) => {
    const week = { id: uid(), ...data, records: {} };
    onCreateWeek(week);
    setActiveWeekId(week.id);
    setShowNewWeek(false);
  };

  const handleEditSave = (data) => {
    onUpdateWeek(activeWeek.id, data);
    setEditingWeek(false);
  };

  const grandTotal = activeWeek
    ? weekWorkers.reduce((sum, w) => sum + calcWorkerWeekTotals(w, activeWeek.records[w.id], weekDates.map((d) => d.key)).totalUpah, 0)
    : 0;

  const projectLabel = (w) => projects.find((p) => p.id === w.projectId)?.name || '(proyek dihapus)';

  return (
    <div className="p-4 pb-24">
      {weeks.length > 0 && (
        <div className="flex items-center gap-2 mb-2">
          <select
            value={activeWeekId || ''}
            onChange={(e) => { setActiveWeekId(e.target.value); setShowNewWeek(false); setEditingWeek(false); }}
            className="flex-1 px-3 py-2 rounded att-body text-sm outline-none"
            style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }}
          >
            {weeks.map((w) => (
              <option key={w.id} value={w.id}>{w.weekLabel} &middot; {projectLabel(w)}</option>
            ))}
          </select>
          <button type="button" onClick={() => { setShowNewWeek(!showNewWeek); setEditingWeek(false); }}
            className="p-2.5 rounded" style={{ background: THEME.amber }}>
            <Plus size={16} color={THEME.charcoal} />
          </button>
          {activeWeek && (
            <button type="button" onClick={() => { setEditingWeek(!editingWeek); setShowNewWeek(false); }}
              className="p-2.5 rounded" style={{ background: THEME.concrete }}>
              <Edit3 size={16} color={THEME.ink} />
            </button>
          )}
        </div>
      )}

      {activeWeek && !editingWeek && !showNewWeek && (
        <div className="mb-4">
          <PrintMenu
            onPrint={(thermal) => openPrintDocument('Absensi Mingguan', buildAbsensiHtml(activeWeek, activeProject, weekWorkers, grandTotal, weekDates, thermal), thermal)}
            onCopy={() => setShowTextPreview(true)}
          />
        </div>
      )}

      {activeWeek && showTextPreview && (
        <TextPreviewModal
          title="Absensi Mingguan"
          text={buildAbsensiText(activeWeek, activeProject, weekWorkers, grandTotal, weekDates)}
          onClose={() => setShowTextPreview(false)}
        />
      )}

      {showNewWeek && (
        <WeekForm projects={projects} workers={workers} onAddProject={onAddProject} onSubmit={handleCreate} onCancel={() => setShowNewWeek(false)} />
      )}

      {editingWeek && activeWeek && (
        <WeekForm
          projects={projects}
          workers={workers}
          onAddProject={onAddProject}
          initial={activeWeek}
          onSubmit={handleEditSave}
          onCancel={() => setEditingWeek(false)}
        />
      )}

      {activeWeek && !editingWeek && !showNewWeek && (
        <>
          <div className="mb-3 p-3 rounded-lg att-mono text-xs" style={{ background: THEME.concrete, color: THEME.inkSoft }}>
            <p style={{ color: THEME.ink }} className="att-body font-semibold text-sm mb-0.5">{activeProject?.name || '(proyek dihapus)'}</p>
            {activeProject?.address && <p>{activeProject.address}</p>}
            {(activeWeek.startDate || activeWeek.endDate) && <p>{activeWeek.startDate} s/d {activeWeek.endDate} &middot; {weekDates.length} hari</p>}
            <p className="mt-1">{weekWorkers.length} dari {workers.length} pekerja ikut minggu ini</p>
          </div>

          {weekWorkers.length > 0 && weekDates.length > 0 && (
            <div className="mb-3">
              <p className="att-mono text-[10px] mb-1" style={{ color: THEME.inkSoft }}>STATUS ISI ABSEN PER HARI</p>
              <div className="flex gap-1 overflow-x-auto att-scroll pb-1">
                {weekDates.map((d) => {
                  const hadirCount = weekWorkers.filter((w) => {
                    const st = dayStatus((activeWeek.records[w.id] || {})[d.key]);
                    return st === 'full' || st === 'half';
                  }).length;
                  const kosong = hadirCount === 0;
                  return (
                    <div key={d.key} title={kosong ? 'Belum ada kehadiran tercatat' : `${hadirCount} pekerja hadir`}
                      className="shrink-0 px-2 py-1.5 rounded text-center"
                      style={{ background: kosong ? '#3A241D' : THEME.paper, border: `1px solid ${kosong ? THEME.rust : THEME.line}`, minWidth: 46 }}>
                      <p className="att-mono text-[9px]" style={{ color: kosong ? '#E8A08F' : THEME.inkSoft }}>{d.dayName.slice(0, 3)}</p>
                      <p className="att-mono text-[9px]" style={{ color: kosong ? '#E8A08F' : THEME.inkSoft }}>{d.shortLabel.split(' ')[1]}</p>
                      <p className="att-mono text-[10px] font-bold mt-0.5" style={{ color: kosong ? THEME.rust : THEME.green }}>{kosong ? '!' : hadirCount}</p>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {weekWorkers.length === 0 && (
            <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>
              Belum ada pekerja dipilih untuk minggu ini. Tap tombol edit di atas untuk memilih pekerja.
            </p>
          )}

          {weekWorkers.map((w) => (
            <WorkerWeekCard
              key={w.id}
              worker={w}
              weekRecord={activeWeek.records[w.id]}
              dates={weekDates}
              evidenceMap={evidence[`${activeWeek.id}:${w.id}`]}
              onSetStatus={(day, status) => onUpdateWeekRecord(activeWeek.id, w.id, day, { status })}
              onSetLembur={(day, lembur) => onUpdateWeekRecord(activeWeek.id, w.id, day, { lembur })}
              onOpenEvidence={(day) => setEvidenceTarget({ weekId: activeWeek.id, workerId: w.id, day })}
            />
          ))}

          {weekWorkers.length > 0 && (
            <div className="mt-2 p-3 rounded-lg flex items-center justify-between" style={{ background: THEME.charcoal }}>
              <span className="att-body font-semibold text-sm" style={{ color: THEME.paper }}>JUMLAH TOTAL UPAH</span>
              <span className="att-mono text-sm font-bold" style={{ color: THEME.amber }}>{formatRupiah(grandTotal)}</span>
            </div>
          )}
        </>
      )}

      {!activeWeek && weeks.length > 0 && !showNewWeek && (
        <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Pilih minggu di atas.</p>
      )}

      {evidenceTarget && (
        <EvidenceModal
          existing={evidence[`${evidenceTarget.weekId}:${evidenceTarget.workerId}`]?.[evidenceTarget.day]}
          onCancel={() => setEvidenceTarget(null)}
          onSave={async (data) => {
            await onSaveEvidence(evidenceTarget.weekId, evidenceTarget.workerId, evidenceTarget.day, data);
            setEvidenceTarget(null);
          }}
        />
      )}
    </div>
  );
}

/* ---------------- Rekap Upah ---------------- */

function buildUpahRows(workers, weeks, projects, payments, kasbon, kasbonPayments, evidence, filter) {
  let filteredWeeks = weeks;
  if (filter && filter !== 'all') {
    if (filter.startsWith('week:')) {
      const weekId = filter.slice(5);
      filteredWeeks = weeks.filter((w) => w.id === weekId);
    } else if (filter.startsWith('project:')) {
      const projectId = filter.slice(8);
      filteredWeeks = weeks.filter((w) => w.projectId === projectId);
    }
  }
  return workers.map((w) => {
    let evidenceFilled = 0;
    let evidenceExpected = 0;
    const weekBreakdown = filteredWeeks
      .map((week) => {
        const dateKeys = getWeekDates(week).map((d) => d.key);
        const totals = calcWorkerWeekTotals(w, week.records[w.id], dateKeys);
        const project = projects.find((p) => p.id === week.projectId);
        const dayMap = (evidence && evidence[`${week.id}:${w.id}`]) || {};
        dateKeys.forEach((day) => {
          const dayRec = (week.records[w.id] || {})[day];
          if (!dayRec || dayStatus(dayRec) === 'off') return;
          evidenceExpected += 4;
          const periods = dayMap[day]?.periods || {};
          evidenceFilled += EVIDENCE_PERIODS.filter((p) => periods[p.key]?.photo).length;
        });
        return { weekId: week.id, weekLabel: week.weekLabel, projectName: project?.name || '-', startDate: week.startDate, endDate: week.endDate, ...totals };
      })
      .filter((x) => x.totalHariBayar > 0 || x.totalJamLembur > 0);
    const totalUpah = weekBreakdown.reduce((sum, x) => sum + x.totalUpah, 0);
    const totalHari = weekBreakdown.reduce((sum, x) => sum + x.totalHariBayar, 0);
    const totalJamLembur = weekBreakdown.reduce((sum, x) => sum + x.totalJamLembur, 0);
    const workerPayments = payments.filter((p) => p.workerId === w.id);
    const workerKasbon = kasbon.filter((k) => k.workerId === w.id);
    const workerKasbonPayments = kasbonPayments.filter((p) => p.workerId === w.id);
    const diterima = workerPayments.reduce((s, p) => s + (Number(p.amount) || 0), 0);
    const totalKasbonDiberi = workerKasbon.reduce((s, k) => s + (Number(k.amount) || 0), 0);
    const totalKasbonDibayar = workerKasbonPayments.reduce((s, p) => s + (Number(p.amount) || 0), 0);
    const totalKasbon = totalKasbonDiberi - totalKasbonDibayar; // sisa kasbon yang belum dilunasi
    const sisa = totalUpah - diterima - totalKasbon;
    return { worker: w, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown, evidenceFilled, evidenceExpected };
  });
}

function RekapUpahTab({ workers, weeks, projects, payments, kasbon, kasbonPayments, evidence, onAddPayment, onDeletePayment }) {
  const [payFormFor, setPayFormFor] = useState(null);
  const [payAmount, setPayAmount] = useState('');
  const [expanded, setExpanded] = useState(null);
  const [showTextPreview, setShowTextPreview] = useState(false);
  const [filter, setFilter] = useState('all'); // 'all' | 'project:<id>' | 'week:<id>'
  const [slipFor, setSlipFor] = useState(null); // worker.id sedang buka slip gaji
  const [showSlipText, setShowSlipText] = useState(null); // worker.id sedang tampil teks slip gaji

  const rows = buildUpahRows(workers, weeks, projects, payments, kasbon, kasbonPayments, evidence, filter);
  const grand = rows.reduce((acc, r) => ({
    totalUpah: acc.totalUpah + r.totalUpah,
    diterima: acc.diterima + r.diterima,
    totalKasbon: acc.totalKasbon + r.totalKasbon,
    sisa: acc.sisa + r.sisa,
  }), { totalUpah: 0, diterima: 0, totalKasbon: 0, sisa: 0 });

  const submitPayment = async (workerId) => {
    const amount = Number(payAmount.replace(/[^0-9]/g, '')) || 0;
    if (amount <= 0) return;
    await onAddPayment({ id: uid(), workerId, amount, date: new Date().toISOString().slice(0, 10) });
    setPayFormFor(null);
    setPayAmount('');
  };

  const filterLabel = () => {
    if (filter === 'all') return 'REKAP UPAH SELURUH PERIODE';
    if (filter.startsWith('project:')) {
      const p = projects.find((pr) => pr.id === filter.slice(8));
      return `REKAP UPAH PROYEK: ${p?.name || '-'}`;
    }
    const week = weeks.find((w) => w.id === filter.slice(5));
    return `REKAP UPAH: ${week?.weekLabel || '-'}`;
  };

  return (
    <div className="p-4 pb-24">
      <select value={filter} onChange={(e) => setFilter(e.target.value)}
        className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-3" style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }}>
        <option value="all">Rekap Seluruh Periode</option>
        {projects.map((p) => {
          const projectWeeks = weeks.filter((w) => w.projectId === p.id);
          if (projectWeeks.length === 0) return null;
          return (
            <optgroup key={p.id} label={p.name}>
              <option value={`project:${p.id}`}>Semua periode - {p.name}</option>
              {projectWeeks.map((w) => (
                <option key={w.id} value={`week:${w.id}`}>{w.weekLabel} ({w.startDate} s/d {w.endDate})</option>
              ))}
            </optgroup>
          );
        })}
      </select>

      <div className="flex items-center justify-between mb-3">
        <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>
          {filterLabel()}
        </p>
        {rows.length > 0 && (
          <div className="flex items-center gap-1.5">
            <button type="button" onClick={() => exportRekapUpahPDF(rows, grand, filterLabel())}
              className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.paper }}>
              <Save size={12} /> PDF
            </button>
            <PrintMenu
              onPrint={(thermal) => openPrintDocument('Rekap Upah', buildRekapHtml(rows, grand, thermal), thermal)}
              onCopy={() => setShowTextPreview(true)}
              onRawBT={() => printToRawBT(buildRekapText(rows, grand))}
            />
          </div>
        )}
      </div>
      {rows.length > 0 && showTextPreview && (
        <TextPreviewModal title="Rekap Upah" text={buildRekapText(rows, grand)} onClose={() => setShowTextPreview(false)} />
      )}
      {weeks.length === 0 && (
        <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>
          Belum ada data absensi mingguan untuk direkap.
        </p>
      )}
      <div className="space-y-2">
        {rows.map(({ worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown, evidenceFilled, evidenceExpected }) => (
          <div key={worker.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div className="flex items-center justify-between mb-1">
              <div>
                <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{worker.name}</p>
                <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{worker.position || '-'}</p>
              </div>
              <p className="att-mono text-sm font-bold" style={{ color: THEME.ink }}>{formatRupiah(totalUpah)}</p>
            </div>
            <p className="att-mono text-[11px] mb-2" style={{ color: THEME.inkSoft }}>
              {totalHari} hari kerja &middot; {totalJamLembur} jam lembur
              {evidenceExpected > 0 && (
                <> &middot; <Camera size={10} className="inline -mt-0.5" /> {evidenceFilled}/{evidenceExpected} bukti foto</>
              )}
            </p>
            <div className="grid grid-cols-3 gap-2 att-mono text-xs">
              <div className="p-2 rounded" style={{ background: THEME.concrete }}>
                <p style={{ color: THEME.inkSoft }}>Diterima</p>
                <p className="mt-0.5 font-semibold" style={{ color: THEME.ink }}>{formatRupiah(diterima)}</p>
              </div>
              <div className="p-2 rounded" style={{ background: THEME.concrete }}>
                <p style={{ color: THEME.inkSoft }}>Sisa Kasbon</p>
                <p className="mt-0.5 font-semibold" style={{ color: THEME.rust }}>{formatRupiah(totalKasbon)}</p>
              </div>
              <div className="p-2 rounded" style={{ background: THEME.concrete }}>
                <p style={{ color: THEME.inkSoft }}>Sisa Upah</p>
                <p className="mt-0.5 font-semibold" style={{ color: sisa > 0 ? THEME.green : THEME.rust }}>{formatRupiah(sisa)}</p>
              </div>
            </div>

            <button type="button" onClick={() => setSlipFor(slipFor === worker.id ? null : worker.id)}
              className="w-full mt-2 flex items-center justify-center gap-1.5 py-1.5 rounded att-mono text-xs font-semibold"
              style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.concrete }}>
              <Receipt size={13} /> Slip Gaji
            </button>
            {slipFor === worker.id && (
              <div className="mt-2 flex items-center gap-1.5">
                <button type="button" onClick={() => exportSlipGajiPDF(
                  { worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown, evidenceFilled, evidenceExpected },
                  filterLabel())}
                  className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.paper }}>
                  <Save size={12} /> PDF
                </button>
                <PrintMenu
                  onPrint={(thermal) => openPrintDocument(`Slip Gaji - ${worker.name}`, buildSlipGajiHtml(
                    { worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown, evidenceFilled, evidenceExpected },
                    filterLabel(), thermal), thermal)}
                  onCopy={() => setShowSlipText(worker.id)}
                  onRawBT={() => printToRawBT(buildSlipGajiText(
                    { worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown, evidenceFilled, evidenceExpected },
                    filterLabel()))}
                />
              </div>
            )}
            {showSlipText === worker.id && (
              <TextPreviewModal
                title={`Slip Gaji - ${worker.name}`}
                text={buildSlipGajiText(
                  { worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown, evidenceFilled, evidenceExpected },
                  filterLabel())}
                onClose={() => setShowSlipText(null)}
              />
            )}

            <div className="flex items-center gap-2 mt-2">
              {payFormFor === worker.id ? (
                <>
                  <input
                    autoFocus
                    value={payAmount}
                    onChange={(e) => setPayAmount(e.target.value.replace(/[^0-9]/g, ''))}
                    placeholder="Jumlah dibayar (Rp)"
                    inputMode="numeric"
                    className="flex-1 px-2 py-1.5 rounded att-mono text-xs outline-none"
                    style={{ background: THEME.concrete, color: THEME.ink, border: `1px solid ${THEME.line}` }}
                  />
                  <button type="button" onClick={() => submitPayment(worker.id)} className="px-2.5 py-1.5 rounded" style={{ background: THEME.amber }}>
                    <Check size={13} color={THEME.charcoal} />
                  </button>
                  <button type="button" onClick={() => { setPayFormFor(null); setPayAmount(''); }} className="px-2.5 py-1.5 rounded" style={{ border: `1px solid ${THEME.line}` }}>
                    <X size={13} color={THEME.inkSoft} />
                  </button>
                </>
              ) : (
                <>
                  <button type="button" onClick={() => { setPayFormFor(worker.id); setPayAmount(''); }}
                    className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-xs font-semibold"
                    style={{ background: THEME.amber, color: THEME.charcoal }}>
                    <Plus size={12} /> Tambah Pembayaran
                  </button>
                  {(workerPayments.length > 0 || workerKasbon.length > 0 || workerKasbonPayments.length > 0 || weekBreakdown.length > 0) && (
                    <button type="button" onClick={() => setExpanded(expanded === worker.id ? null : worker.id)}
                      className="att-mono text-xs" style={{ color: THEME.inkSoft }}>
                      {expanded === worker.id ? 'Sembunyikan riwayat' : 'Lihat riwayat'}
                    </button>
                  )}
                </>
              )}
            </div>

            {expanded === worker.id && (
              <div className="mt-2 space-y-1 att-mono text-[11px]" style={{ color: THEME.inkSoft }}>
                {weekBreakdown.map((wk) => (
                  <div key={wk.weekId} className="p-1.5 rounded" style={{ background: THEME.concrete }}>
                    <div className="flex justify-between items-center">
                      <span style={{ color: THEME.ink }} className="font-semibold">{wk.weekLabel} &middot; {wk.projectName}</span>
                      <span>{formatRupiah(wk.totalUpah)}</span>
                    </div>
                    <div>{wk.startDate} s/d {wk.endDate}</div>
                    <div>{wk.totalHariBayar} hari ({formatRupiah(wk.totalHarianRp)}) + {wk.totalJamLembur} jam lembur ({formatRupiah(wk.totalLemburRp)})</div>
                  </div>
                ))}
                {workerPayments.map((p) => (
                  <div key={p.id} className="flex justify-between items-center p-1.5 rounded" style={{ background: THEME.concrete }}>
                    <span>Dibayar {p.date}</span>
                    <span className="flex items-center gap-2">
                      {formatRupiah(p.amount)}
                      <button type="button" onClick={() => onDeletePayment(p.id)}><Trash2 size={11} color={THEME.rust} /></button>
                    </span>
                  </div>
                ))}
                {workerKasbon.map((k) => (
                  <div key={k.id} className="flex justify-between items-center p-1.5 rounded" style={{ background: THEME.concrete }}>
                    <span>Kasbon {k.date}{k.note ? ` - ${k.note}` : ''}</span>
                    <span>{formatRupiah(k.amount)}</span>
                  </div>
                ))}
                {workerKasbonPayments.map((p) => (
                  <div key={p.id} className="flex justify-between items-center p-1.5 rounded" style={{ background: THEME.concrete }}>
                    <span>Bayar kasbon {p.date}{p.note ? ` - ${p.note}` : ''}</span>
                    <span style={{ color: THEME.green }}>{formatRupiah(p.amount)}</span>
                  </div>
                ))}
                {sisa > 0 && (
                  <div className="p-1.5 rounded" style={{ background: '#3A241D', color: '#E8A08F' }}>
                    Upah belum lunas {formatRupiah(sisa)}. Tercatat otomatis sebagai utang ke {worker.name} di tab Rekap Proyek.
                  </div>
                )}
              </div>
            )}
          </div>
        ))}
      </div>

      {rows.length > 0 && (
        <div className="mt-3 p-3 rounded-lg" style={{ background: THEME.charcoal }}>
          <div className="flex justify-between att-mono text-xs mb-1" style={{ color: '#9C948A' }}>
            <span>Total Upah</span><span>{formatRupiah(grand.totalUpah)}</span>
          </div>
          <div className="flex justify-between att-mono text-xs mb-1" style={{ color: '#9C948A' }}>
            <span>Total Diterima</span><span>{formatRupiah(grand.diterima)}</span>
          </div>
          <div className="flex justify-between att-mono text-xs mb-1" style={{ color: '#9C948A' }}>
            <span>Total Kasbon</span><span>{formatRupiah(grand.totalKasbon)}</span>
          </div>
          <div className="flex justify-between att-body text-sm font-semibold" style={{ color: THEME.amber }}>
            <span>Total Sisa</span><span>{formatRupiah(grand.sisa)}</span>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------- Kasbon Pekerja ---------------- */

function KasbonTab({ workers, kasbon, kasbonPayments, onAdd, onDelete, onAddKasbonPayment, onDeleteKasbonPayment }) {
  const [workerId, setWorkerId] = useState('');
  const [amount, setAmount] = useState('');
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [note, setNote] = useState('');
  const [error, setError] = useState('');
  const [showTextPreview, setShowTextPreview] = useState(false);
  const [payFormFor, setPayFormFor] = useState(null);
  const [payAmount, setPayAmount] = useState('');
  const [payDate, setPayDate] = useState(new Date().toISOString().slice(0, 10));
  const [payNote, setPayNote] = useState('');

  const workerName = (id) => workers.find((w) => w.id === id)?.name || '(pekerja dihapus)';

  const handleAdd = async () => {
    setError('');
    if (!workerId || !amount) {
      setError('Pilih pekerja dan isi jumlah kasbon.');
      return;
    }
    await onAdd({ id: uid(), workerId, amount: Number(amount) || 0, date, note: note.trim() });
    setAmount('');
    setNote('');
  };

  const submitPayment = async (wId) => {
    const amt = Number(payAmount.replace(/[^0-9]/g, '')) || 0;
    if (amt <= 0) return;
    await onAddKasbonPayment({ id: uid(), workerId: wId, amount: amt, date: payDate, note: payNote.trim() });
    setPayFormFor(null);
    setPayAmount('');
    setPayNote('');
  };

  const ringkasan = workers.map((w) => {
    const totalKasbon = kasbon.filter((k) => k.workerId === w.id).reduce((s, k) => s + (Number(k.amount) || 0), 0);
    const totalDibayar = kasbonPayments.filter((p) => p.workerId === w.id).reduce((s, p) => s + (Number(p.amount) || 0), 0);
    return { worker: w, totalKasbon, totalDibayar, sisa: totalKasbon - totalDibayar };
  }).filter((x) => x.totalKasbon > 0);

  const sorted = [
    ...kasbon.map((k) => ({ ...k, jenis: 'beri' })),
    ...kasbonPayments.map((p) => ({ ...p, jenis: 'bayar' })),
  ].sort((a, b) => (a.date < b.date ? 1 : -1));
  const total = kasbon.reduce((s, k) => s + (Number(k.amount) || 0), 0);
  const text = ['KASBON PEKERJA', '================================',
    ...sorted.map((k) => `${k.date}  ${workerName(k.workerId)}  [${k.jenis === 'bayar' ? 'BAYAR' : 'KASBON'}]\n  ${formatRupiah(k.amount)}${k.note ? '  (' + k.note + ')' : ''}`),
    '================================', `TOTAL KASBON DIBERIKAN: ${formatRupiah(total)}`].join('\n');

  return (
    <div className="p-4 pb-24">
      {ringkasan.length > 0 && (
        <div className="mb-5">
          <p className="att-mono text-xs mb-2" style={{ color: THEME.inkSoft }}>SISA KASBON PER PEKERJA</p>
          <div className="space-y-2">
            {ringkasan.map(({ worker, totalKasbon, totalDibayar, sisa }) => (
              <div key={worker.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
                <div className="flex items-center justify-between mb-1">
                  <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{worker.name}</p>
                  <p className="att-mono text-sm font-bold" style={{ color: sisa > 0 ? THEME.rust : THEME.green }}>{formatRupiah(sisa)}</p>
                </div>
                <p className="att-mono text-[11px] mb-2" style={{ color: THEME.inkSoft }}>
                  Total kasbon {formatRupiah(totalKasbon)} &middot; Sudah dibayar {formatRupiah(totalDibayar)}
                </p>
                {sisa > 0 && (
                  payFormFor === worker.id ? (
                    <div className="space-y-1.5">
                      <div className="flex items-center gap-2">
                        <input
                          autoFocus
                          value={payAmount}
                          onChange={(e) => setPayAmount(e.target.value.replace(/[^0-9]/g, ''))}
                          placeholder="Jumlah bayar (Rp)"
                          inputMode="numeric"
                          className="flex-1 px-2 py-1.5 rounded att-mono text-xs outline-none"
                          style={{ background: THEME.concrete, color: THEME.ink, border: `1px solid ${THEME.line}` }}
                        />
                        <input
                          type="date"
                          value={payDate}
                          onChange={(e) => setPayDate(e.target.value)}
                          className="px-2 py-1.5 rounded att-mono text-xs outline-none"
                          style={{ background: THEME.concrete, color: THEME.ink, border: `1px solid ${THEME.line}` }}
                        />
                      </div>
                      <input
                        value={payNote}
                        onChange={(e) => setPayNote(e.target.value)}
                        placeholder="Catatan (opsional)"
                        className="w-full px-2 py-1.5 rounded att-body text-xs outline-none"
                        style={{ background: THEME.concrete, color: THEME.ink, border: `1px solid ${THEME.line}` }}
                      />
                      <div className="flex items-center gap-2">
                        <button type="button" onClick={() => submitPayment(worker.id)} className="flex-1 py-1.5 rounded att-mono text-xs font-semibold" style={{ background: THEME.amber, color: THEME.charcoal }}>
                          Simpan
                        </button>
                        <button type="button" onClick={() => printKasbonInvoice('Bukti Bayar Kasbon', worker, Number(payAmount.replace(/[^0-9]/g, '')) || 0, payDate, payNote, 'bayar')}
                          className="px-2.5 py-1.5 rounded" style={{ border: `1px solid ${THEME.line}` }}>
                          <Receipt size={13} color={THEME.inkSoft} />
                        </button>
                        <button type="button" onClick={() => { setPayFormFor(null); setPayAmount(''); setPayNote(''); }} className="px-2.5 py-1.5 rounded" style={{ border: `1px solid ${THEME.line}` }}>
                          <X size={13} color={THEME.inkSoft} />
                        </button>
                      </div>
                    </div>
                  ) : (
                    <button type="button" onClick={() => { setPayFormFor(worker.id); setPayAmount(''); setPayDate(new Date().toISOString().slice(0, 10)); }}
                      className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-xs font-semibold"
                      style={{ background: THEME.amber, color: THEME.charcoal }}>
                      <Plus size={12} /> Bayar Kasbon
                    </button>
                  )
                )}
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="p-4 rounded-lg mb-5" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
        <p className="att-body font-semibold text-sm mb-3" style={{ color: THEME.ink }}>Catat Kasbon</p>
        <div className="grid sm:grid-cols-2 gap-3">
          <select value={workerId} onChange={(e) => setWorkerId(e.target.value)}
            className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }}>
            <option value="">Pilih pekerja...</option>
            {workers.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
          </select>
          <input value={amount} onChange={(e) => setAmount(e.target.value.replace(/[^0-9]/g, ''))}
            placeholder="Jumlah kasbon (Rp)" inputMode="numeric" className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input type="date" value={date} onChange={(e) => setDate(e.target.value)}
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={note} onChange={(e) => setNote(e.target.value)}
            placeholder="Catatan (opsional)" className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>
        {error && <p className="text-sm mt-2" style={{ color: THEME.rust }}>{error}</p>}
        <div className="flex gap-2 mt-3">
          <button type="button" onClick={handleAdd}
            className="flex-1 flex items-center justify-center gap-1.5 py-2 rounded att-body font-semibold text-sm"
            style={{ background: THEME.amber, color: THEME.charcoal }}>
            <Plus size={15} /> Tambah Kasbon
          </button>
          <button type="button"
            onClick={() => printKasbonInvoice('Bukti Pemberian Kasbon', workers.find((w) => w.id === workerId), Number(amount) || 0, date, note, 'beri')}
            disabled={!workerId || !amount}
            className="px-3 py-2 rounded disabled:opacity-40" style={{ border: `1px solid ${THEME.line}` }} title="Cetak invoice ke RawBT">
            <Bluetooth size={16} color={THEME.inkSoft} />
          </button>
        </div>
      </div>

      <div className="flex items-center justify-between mb-2">
        <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{kasbon.length} CATATAN KASBON</p>
        {kasbon.length > 0 && (
          <PrintMenu
            onPrint={(thermal) => openPrintDocument('Kasbon Pekerja', buildKasbonHtml(sorted, workerName, total, thermal), thermal)}
            onCopy={() => setShowTextPreview(true)}
          />
        )}
      </div>
      {showTextPreview && <TextPreviewModal title="Kasbon Pekerja" text={text} onClose={() => setShowTextPreview(false)} />}

      <div className="space-y-2">
        {sorted.map((k) => (
          <div key={`${k.jenis}-${k.id}`} className="p-3 rounded-lg flex items-center justify-between" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div>
              <div className="flex items-center gap-1.5">
                <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{workerName(k.workerId)}</p>
                <span className="att-mono text-[9px] px-1.5 py-0.5 rounded-full font-semibold"
                  style={{ background: k.jenis === 'bayar' ? THEME.green : THEME.rust, color: THEME.paper }}>
                  {k.jenis === 'bayar' ? 'BAYAR' : 'KASBON'}
                </span>
              </div>
              <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{k.date}{k.note ? ` - ${k.note}` : ''}</p>
            </div>
            <div className="flex items-center gap-2">
              <p className="att-mono text-sm font-bold" style={{ color: k.jenis === 'bayar' ? THEME.green : THEME.rust }}>{formatRupiah(k.amount)}</p>
              <button type="button" onClick={() => (k.jenis === 'bayar' ? onDeleteKasbonPayment(k.id) : onDelete(k.id))}><Trash2 size={14} color={THEME.rust} /></button>
            </div>
          </div>
        ))}
        {sorted.length === 0 && (
          <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada catatan kasbon.</p>
        )}
      </div>

      {kasbon.length > 0 && (
        <div className="mt-3 p-3 rounded-lg flex items-center justify-between" style={{ background: THEME.charcoal }}>
          <span className="att-body font-semibold text-sm" style={{ color: THEME.paper }}>TOTAL KASBON</span>
          <span className="att-mono text-sm font-bold" style={{ color: THEME.amber }}>{formatRupiah(total)}</span>
        </div>
      )}
    </div>
  );
}

/* ---------------- Data Material & Belanja Harian ---------------- */

function MaterialForm({ onAdd, onUpdate, editing, onCancelEdit }) {
  const [name, setName] = useState(editing?.name || '');
  const [merk, setMerk] = useState(editing?.merk || '');
  const [unit, setUnit] = useState(editing?.unit || '');
  const [price, setPrice] = useState(editing ? String(editing.price) : '');
  const [kategori, setKategori] = useState(editing?.kategori || 'material');
  const [stokMinimum, setStokMinimum] = useState(editing ? String(editing.stokMinimum ?? '') : '');
  const [error, setError] = useState('');

  useEffect(() => {
    setName(editing?.name || '');
    setMerk(editing?.merk || '');
    setUnit(editing?.unit || '');
    setPrice(editing ? String(editing.price) : '');
    setKategori(editing?.kategori || 'material');
    setStokMinimum(editing ? String(editing.stokMinimum ?? '') : '');
  }, [editing]);

  const submit = async () => {
    setError('');
    if (!name || price === '') {
      setError('Nama material dan harga wajib diisi.');
      return;
    }
    const payload = {
      name: name.trim(), merk: merk.trim(), unit: unit.trim(), price: Number(price) || 0, kategori,
      stokMinimum: stokMinimum === '' ? 0 : Number(stokMinimum) || 0,
    };
    if (editing) {
      await onUpdate(editing.id, payload);
    } else {
      await onAdd({ id: uid(), ...payload });
      setName(''); setMerk(''); setUnit(''); setPrice(''); setKategori('material'); setStokMinimum('');
    }
  };

  return (
    <div className="p-4 rounded-lg mb-4" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
      <p className="att-body font-semibold text-sm mb-3" style={{ color: THEME.ink }}>{editing ? 'Edit Material' : 'Tambah Material'}</p>
      <div className="grid sm:grid-cols-3 gap-3">
        <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Nama material"
          className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
        <input value={merk} onChange={(e) => setMerk(e.target.value)} placeholder="Merk"
          className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        <select value={kategori} onChange={(e) => setKategori(e.target.value)}
          className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }}>
          {KATEGORI_LIST.map((k) => <option key={k} value={k}>{KATEGORI_LABELS[k]}</option>)}
        </select>
        <input value={unit} onChange={(e) => setUnit(e.target.value)} placeholder="Satuan (sak, kg, btg)"
          className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        <input value={price} onChange={(e) => setPrice(e.target.value.replace(/[^0-9]/g, ''))} placeholder="Harga (Rp)" inputMode="numeric"
          className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        <input value={stokMinimum} onChange={(e) => setStokMinimum(e.target.value.replace(/[^0-9]/g, ''))} placeholder="Stok minimum (opsional)" inputMode="numeric"
          className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
      </div>
      {error && <p className="text-sm mt-2" style={{ color: THEME.rust }}>{error}</p>}
      <div className="flex gap-2 mt-3">
        <button type="button" onClick={submit}
          className="flex-1 flex items-center justify-center gap-1.5 py-2 rounded att-body font-semibold text-sm"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          {editing ? <Check size={15} /> : <Plus size={15} />} {editing ? 'Simpan Perubahan' : 'Tambah Material'}
        </button>
        {editing && (
          <button type="button" onClick={onCancelEdit} className="px-4 py-2 rounded att-body text-sm" style={{ border: `1px solid ${THEME.line}`, color: THEME.inkSoft }}>
            Batal
          </button>
        )}
      </div>
    </div>
  );
}

function SupplierForm({ onAdd, onUpdate, editing, onCancelEdit }) {
  const [name, setName] = useState(editing?.name || '');
  const [address, setAddress] = useState(editing?.address || '');
  const [phone, setPhone] = useState(editing?.phone || '');
  const [error, setError] = useState('');

  useEffect(() => {
    setName(editing?.name || '');
    setAddress(editing?.address || '');
    setPhone(editing?.phone || '');
  }, [editing]);

  const submit = async () => {
    setError('');
    if (!name) {
      setError('Nama toko/suplier wajib diisi.');
      return;
    }
    const payload = { name: name.trim(), address: address.trim(), phone: phone.trim() };
    if (editing) {
      await onUpdate(editing.id, payload);
    } else {
      await onAdd({ id: uid(), ...payload });
      setName(''); setAddress(''); setPhone('');
    }
  };

  return (
    <div className="p-4 rounded-lg mb-4" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
      <p className="att-body font-semibold text-sm mb-3" style={{ color: THEME.ink }}>{editing ? 'Edit Toko/Suplier' : 'Tambah Toko/Suplier'}</p>
      <div className="grid sm:grid-cols-2 gap-3">
        <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Nama toko/distributor"
          className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
        <input value={address} onChange={(e) => setAddress(e.target.value)} placeholder="Alamat"
          className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
        <input value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="No. Telepon"
          className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
      </div>
      {error && <p className="text-sm mt-2" style={{ color: THEME.rust }}>{error}</p>}
      <div className="flex gap-2 mt-3">
        <button type="button" onClick={submit}
          className="flex-1 flex items-center justify-center gap-1.5 py-2 rounded att-body font-semibold text-sm"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          {editing ? <Check size={15} /> : <Plus size={15} />} {editing ? 'Simpan Perubahan' : 'Tambah Toko'}
        </button>
        {editing && (
          <button type="button" onClick={onCancelEdit} className="px-4 py-2 rounded att-body text-sm" style={{ border: `1px solid ${THEME.line}`, color: THEME.inkSoft }}>
            Batal
          </button>
        )}
      </div>
    </div>
  );
}

const STATUS_BAYAR_LABEL = { cash: 'Cash', tempo: 'Tempo', utang: 'Utang' };
const STATUS_BAYAR_COLOR = { cash: THEME.green, tempo: THEME.amber, utang: THEME.rust };

function PeralatanForm({ onAdd, onUpdate, editing, onCancelEdit }) {
  const [nama, setNama] = useState(editing?.nama || '');
  const [merk, setMerk] = useState(editing?.merk || '');
  const [kapasitas, setKapasitas] = useState(editing?.kapasitas || '');
  const [unit, setUnit] = useState(editing?.unit || 'unit');
  const [jumlah, setJumlah] = useState(editing ? String(editing.jumlah) : '1');
  const [waktuPembelian, setWaktuPembelian] = useState(editing?.waktuPembelian || new Date().toISOString().slice(0, 10));
  const [error, setError] = useState('');

  useEffect(() => {
    setNama(editing?.nama || '');
    setMerk(editing?.merk || '');
    setKapasitas(editing?.kapasitas || '');
    setUnit(editing?.unit || 'unit');
    setJumlah(editing ? String(editing.jumlah) : '1');
    setWaktuPembelian(editing?.waktuPembelian || new Date().toISOString().slice(0, 10));
  }, [editing]);

  const submit = async () => {
    setError('');
    if (!nama) {
      setError('Nama peralatan wajib diisi.');
      return;
    }
    const payload = { nama: nama.trim(), merk: merk.trim(), kapasitas: kapasitas.trim(), unit: unit.trim() || 'unit', jumlah: Number(jumlah) || 0, waktuPembelian };
    if (editing) {
      await onUpdate(editing.id, payload);
    } else {
      await onAdd({ id: uid(), ...payload });
      setNama(''); setMerk(''); setKapasitas(''); setUnit('unit'); setJumlah('1'); setWaktuPembelian(new Date().toISOString().slice(0, 10));
    }
  };

  return (
    <div className="p-4 rounded-lg mb-4" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
      <p className="att-body font-semibold text-sm mb-3" style={{ color: THEME.ink }}>{editing ? 'Edit Peralatan' : 'Tambah Peralatan Manual'}</p>
      <div className="grid sm:grid-cols-2 gap-3">
        <input value={nama} onChange={(e) => setNama(e.target.value)} placeholder="Nama peralatan"
          className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
        <input value={merk} onChange={(e) => setMerk(e.target.value)} placeholder="Merk"
          className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        <input value={kapasitas} onChange={(e) => setKapasitas(e.target.value)} placeholder="Kapasitas (mis. 5000 watt, 3 ton)"
          className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        <input value={jumlah} onChange={(e) => setJumlah(e.target.value.replace(/[^0-9]/g, ''))} placeholder="Jumlah" inputMode="numeric"
          className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        <input value={unit} onChange={(e) => setUnit(e.target.value)} placeholder="Satuan (unit, buah)"
          className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        <input type="date" value={waktuPembelian} onChange={(e) => setWaktuPembelian(e.target.value)}
          className="px-3 py-2 rounded att-mono text-xs outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
      </div>
      {error && <p className="text-sm mt-2" style={{ color: THEME.rust }}>{error}</p>}
      <div className="flex gap-2 mt-3">
        <button type="button" onClick={submit}
          className="flex-1 flex items-center justify-center gap-1.5 py-2 rounded att-body font-semibold text-sm"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          {editing ? <Check size={15} /> : <Plus size={15} />} {editing ? 'Simpan Perubahan' : 'Tambah Peralatan'}
        </button>
        {editing && (
          <button type="button" onClick={onCancelEdit} className="px-4 py-2 rounded att-body text-sm" style={{ border: `1px solid ${THEME.line}`, color: THEME.inkSoft }}>
            Batal
          </button>
        )}
      </div>
    </div>
  );
}

function BelanjaTab({ materials, onAddMaterial, onUpdateMaterial, onDeleteMaterial, suppliers, onAddSupplier, onUpdateSupplier, onDeleteSupplier, projects, purchases, onAddPurchase, onDeletePurchase }) {
  const [showMaterials, setShowMaterials] = useState(false);
  const [editingMaterial, setEditingMaterial] = useState(null);
  const [showSuppliers, setShowSuppliers] = useState(false);
  const [editingSupplier, setEditingSupplier] = useState(null);
  const [materialId, setMaterialId] = useState('');
  const [itemQty, setItemQty] = useState('1');
  const [itemPriceOverride, setItemPriceOverride] = useState('');
  const [cart, setCart] = useState([]);
  const [supplierId, setSupplierId] = useState('');
  const [gudangId, setGudangId] = useState(GUDANG_UTAMA_ID);
  const [noNota, setNoNota] = useState('');
  const [statusBayar, setStatusBayar] = useState('cash');
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [note, setNote] = useState('');
  const [notaPhoto, setNotaPhoto] = useState(null); // { photo, time }
  const [photoBusy, setPhotoBusy] = useState(false);
  const [error, setError] = useState('');
  const [showTextPreview, setShowTextPreview] = useState(false);

  const gudangList = getGudangList(projects);
  const selectedMaterial = materials.find((m) => m.id === materialId);
  const selectedSupplier = suppliers.find((s) => s.id === supplierId);
  const cartTotal = cart.reduce((s, it) => s + it.qty * it.price, 0);

  const addToCart = () => {
    setError('');
    if (!materialId || !itemQty) {
      setError('Pilih material dan isi jumlah dulu.');
      return;
    }
    const price = itemPriceOverride !== '' ? Number(itemPriceOverride) : (selectedMaterial?.price || 0);
    setCart([...cart, {
      id: uid(), materialId, materialName: selectedMaterial?.name || '-', unit: selectedMaterial?.unit || '',
      kategori: selectedMaterial?.kategori || 'non_kategori', qty: Number(itemQty) || 0, price,
    }]);
    setMaterialId('');
    setItemQty('1');
    setItemPriceOverride('');
  };

  const removeFromCart = (id) => setCart(cart.filter((it) => it.id !== id));

  const handleNotaPhoto = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setPhotoBusy(true);
    try {
      const dataUri = await resizeImage(file);
      setNotaPhoto({ photo: dataUri, time: nowTimeLabel() });
    } catch {
      alert('Gagal memuat foto. Coba lagi.');
    }
    setPhotoBusy(false);
    e.target.value = '';
  };

  const handleSaveNota = async () => {
    setError('');
    if (cart.length === 0) {
      setError('Tambahkan minimal 1 item belanja ke nota ini dulu.');
      return;
    }
    await onAddPurchase({
      id: uid(), date, supplierId, supplierName: selectedSupplier?.name || '', noNota: noNota.trim(), statusBayar,
      gudangId, note: note.trim(), items: cart, photo: notaPhoto,
    });
    setCart([]);
    setNoNota('');
    setNote('');
    setNotaPhoto(null);
  };

  const sorted = [...purchases].sort((a, b) => (a.date < b.date ? 1 : -1));
  const total = purchases.reduce((s, p) => s + purchaseTotal(p), 0);
  const text = ['BELANJA HARIAN MATERIAL', '================================',
    ...sorted.map((p) => {
      const items = purchaseItems(p);
      const itemLines = items.map((it) => `  - ${it.materialName}: ${it.qty} ${it.unit || ''} x ${formatRupiah(it.price)} = ${formatRupiah(it.qty * it.price)}`).join('\n');
      return `${p.date}${p.noNota ? '  Nota ' + p.noNota : ''}${p.supplierName ? '  (' + p.supplierName + ')' : ''}\n${itemLines}\n  Total nota: ${formatRupiah(purchaseTotal(p))}${p.note ? '  - ' + p.note : ''}`;
    }),
    '================================', `TOTAL BELANJA: ${formatRupiah(total)}`].join('\n');

  return (
    <div className="p-4 pb-24">
      <div className="grid grid-cols-2 gap-1.5 mb-3">
        <button type="button" onClick={() => setShowMaterials(!showMaterials)}
          className="flex items-center justify-center gap-1 px-2 py-2 rounded-lg att-body text-xs font-semibold"
          style={{ background: THEME.charcoal, color: THEME.paper }}>
          Material ({materials.length}) {showMaterials ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
        </button>
        <button type="button" onClick={() => setShowSuppliers(!showSuppliers)}
          className="flex items-center justify-center gap-1 px-2 py-2 rounded-lg att-body text-xs font-semibold"
          style={{ background: THEME.charcoal, color: THEME.paper }}>
          Suplier ({suppliers.length}) {showSuppliers ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
        </button>
      </div>

      {showSuppliers && (
        <>
          <SupplierForm
            onAdd={onAddSupplier}
            onUpdate={async (id, patch) => { await onUpdateSupplier(id, patch); setEditingSupplier(null); }}
            editing={editingSupplier}
            onCancelEdit={() => setEditingSupplier(null)}
          />
          <div className="space-y-2 mb-5">
            {suppliers.map((s) => (
              <div key={s.id} className="p-3 rounded-lg flex items-center justify-between" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
                <div>
                  <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{s.name}</p>
                  <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{s.address}{s.phone ? ` · ${s.phone}` : ''}</p>
                </div>
                <div className="flex gap-1">
                  <button type="button" onClick={() => setEditingSupplier(s)} className="p-2 rounded" style={{ background: THEME.concrete }}>
                    <Edit3 size={14} color={THEME.inkSoft} />
                  </button>
                  <button type="button" onClick={() => onDeleteSupplier(s.id)} className="p-2 rounded" style={{ background: THEME.concrete }}>
                    <Trash2 size={14} color={THEME.rust} />
                  </button>
                </div>
              </div>
            ))}
            {suppliers.length === 0 && (
              <p className="att-body text-sm text-center py-4" style={{ color: THEME.inkSoft }}>Belum ada toko/suplier. Tambahkan di atas.</p>
            )}
          </div>
        </>
      )}

      {showMaterials && (
        <>
          <MaterialForm
            onAdd={onAddMaterial}
            onUpdate={async (id, patch) => { await onUpdateMaterial(id, patch); setEditingMaterial(null); }}
            editing={editingMaterial}
            onCancelEdit={() => setEditingMaterial(null)}
          />
          <div className="space-y-2 mb-5">
            {materials.map((m) => (
              <div key={m.id} className="p-3 rounded-lg flex items-center justify-between" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
                <div>
                  <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{m.name}</p>
                  <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{formatRupiah(m.price)} / {m.unit || 'unit'} &middot; {KATEGORI_LABELS[m.kategori] || KATEGORI_LABELS.non_kategori}</p>
                </div>
                <div className="flex gap-1">
                  <button type="button" onClick={() => setEditingMaterial(m)} className="p-2 rounded" style={{ background: THEME.concrete }}>
                    <Edit3 size={14} color={THEME.inkSoft} />
                  </button>
                  <button type="button" onClick={() => onDeleteMaterial(m.id)} className="p-2 rounded" style={{ background: THEME.concrete }}>
                    <Trash2 size={14} color={THEME.rust} />
                  </button>
                </div>
              </div>
            ))}
            {materials.length === 0 && (
              <p className="att-body text-sm text-center py-4" style={{ color: THEME.inkSoft }}>Belum ada material. Tambahkan di atas.</p>
            )}
          </div>
        </>
      )}

      <div className="p-4 rounded-lg mb-5" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
        <p className="att-body font-semibold text-sm mb-3" style={{ color: THEME.ink }}>Nota Belanja Baru</p>

        <div className="grid sm:grid-cols-2 gap-3 mb-3">
          <select value={supplierId} onChange={(e) => setSupplierId(e.target.value)}
            className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }}>
            <option value="">Beli dari toko/suplier mana? (opsional)</option>
            {suppliers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
          <select value={gudangId} onChange={(e) => setGudangId(e.target.value)}
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }}>
            {gudangList.map((g) => <option key={g.id} value={g.id}>Masuk ke: {g.name}</option>)}
          </select>
          <input type="date" value={date} onChange={(e) => setDate(e.target.value)}
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={noNota} onChange={(e) => setNoNota(e.target.value)} placeholder="No. Nota (opsional)"
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <select value={statusBayar} onChange={(e) => setStatusBayar(e.target.value)}
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }}>
            <option value="cash">Status Bayar: Cash</option>
            <option value="tempo">Status Bayar: Tempo</option>
            <option value="utang">Status Bayar: Utang</option>
          </select>
          <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="Catatan nota (opsional)"
            className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>

        <div className="mb-3">
          {notaPhoto ? (
            <div className="flex items-center gap-2">
              <img src={notaPhoto.photo} alt="Foto nota" className="w-14 h-14 rounded object-cover" style={{ border: `1px solid ${THEME.line}` }} />
              <div className="flex-1">
                <p className="att-mono text-[10px]" style={{ color: THEME.inkSoft }}>Diambil: {notaPhoto.time}</p>
                <button type="button" onClick={() => setNotaPhoto(null)} className="att-mono text-[10px]" style={{ color: THEME.rust }}>Hapus foto</button>
              </div>
            </div>
          ) : (
            <div className="flex gap-2">
              <label className="flex-1 flex items-center justify-center gap-1.5 py-2 rounded att-mono text-xs font-semibold cursor-pointer"
                style={{ background: THEME.concrete, color: THEME.ink }}>
                <Camera size={14} /> {photoBusy ? '...' : 'Foto Kamera'}
                <input type="file" accept="image/*" capture="environment" className="hidden" onChange={handleNotaPhoto} disabled={photoBusy} />
              </label>
              <label className="flex-1 flex items-center justify-center gap-1.5 py-2 rounded att-mono text-xs font-semibold cursor-pointer"
                style={{ background: THEME.concrete, color: THEME.ink }}>
                <Save size={14} /> {photoBusy ? '...' : 'Foto File'}
                <input type="file" accept="image/*" className="hidden" onChange={handleNotaPhoto} disabled={photoBusy} />
              </label>
            </div>
          )}
        </div>

        <div className="p-3 rounded-lg mb-3" style={{ background: THEME.concrete }}>
          <p className="att-mono text-[10px] mb-2" style={{ color: THEME.inkSoft }}>TAMBAH ITEM KE NOTA INI</p>
          <div className="grid sm:grid-cols-2 gap-2">
            <select value={materialId} onChange={(e) => { setMaterialId(e.target.value); setItemPriceOverride(''); }}
              className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.paper, color: THEME.ink }}>
              <option value="">Pilih material...</option>
              {materials.map((m) => <option key={m.id} value={m.id}>{m.name}{m.merk ? ` - ${m.merk}` : ''} ({formatRupiah(m.price)}/{m.unit || 'unit'})</option>)}
            </select>
            <input value={itemQty} onChange={(e) => setItemQty(e.target.value.replace(/[^0-9.]/g, ''))} placeholder="Jumlah/Qty" inputMode="decimal"
              className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.paper, color: THEME.ink }} />
            <input value={itemPriceOverride} onChange={(e) => setItemPriceOverride(e.target.value.replace(/[^0-9]/g, ''))}
              placeholder={`Harga (default ${formatRupiah(selectedMaterial?.price || 0)})`} inputMode="numeric"
              className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.paper, color: THEME.ink }} />
          </div>
          <button type="button" onClick={addToCart}
            className="w-full mt-2 flex items-center justify-center gap-1.5 py-1.5 rounded att-mono text-xs font-semibold"
            style={{ background: THEME.paper, border: `1px solid ${THEME.line}`, color: THEME.ink }}>
            <Plus size={13} /> Tambah ke Nota
          </button>
        </div>

        {cart.length > 0 && (
          <div className="space-y-1.5 mb-3">
            {cart.map((it) => (
              <div key={it.id} className="flex items-center justify-between p-2 rounded" style={{ background: THEME.concrete }}>
                <div>
                  <p className="att-body text-sm" style={{ color: THEME.ink }}>{it.materialName}</p>
                  <p className="att-mono text-[10px]" style={{ color: THEME.inkSoft }}>{it.qty} {it.unit} &times; {formatRupiah(it.price)} = {formatRupiah(it.qty * it.price)}</p>
                </div>
                <button type="button" onClick={() => removeFromCart(it.id)}><Trash2 size={13} color={THEME.rust} /></button>
              </div>
            ))}
            <div className="flex justify-between att-mono text-xs font-semibold px-1" style={{ color: THEME.ink }}>
              <span>Total Nota</span><span>{formatRupiah(cartTotal)}</span>
            </div>
          </div>
        )}

        {error && <p className="text-sm mt-2" style={{ color: THEME.rust }}>{error}</p>}
        <button type="button" onClick={handleSaveNota}
          className="w-full mt-1 flex items-center justify-center gap-1.5 py-2 rounded att-body font-semibold text-sm"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Check size={15} /> Simpan Nota Belanja ({cart.length} item)
        </button>
        <p className="att-mono text-[10px] mt-2" style={{ color: THEME.inkSoft }}>Semua item otomatis masuk ke stok gudang tujuan. Untuk pemakaian, catat di tab Gudang.</p>
      </div>

      <div className="flex items-center justify-between mb-2">
        <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{purchases.length} CATATAN BELANJA</p>
        {purchases.length > 0 && (
          <PrintMenu
            onPrint={(thermal) => openPrintDocument('Belanja Harian', buildBelanjaHtml(sorted, total, thermal), thermal)}
            onCopy={() => setShowTextPreview(true)}
          />
        )}
      </div>
      {showTextPreview && <TextPreviewModal title="Belanja Harian" text={text} onClose={() => setShowTextPreview(false)} />}

      <div className="space-y-2">
        {sorted.map((p) => {
          const items = purchaseItems(p);
          return (
            <div key={p.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
              <div className="flex items-start justify-between mb-1.5">
                <div>
                  <div className="flex items-center gap-1.5 flex-wrap">
                    <p className="att-mono text-xs font-semibold" style={{ color: THEME.ink }}>{p.date}</p>
                    {p.noNota && <span className="att-mono text-[10px]" style={{ color: THEME.inkSoft }}>Nota: {p.noNota}</span>}
                    {p.statusBayar && (
                      <span className="att-mono text-[9px] px-1.5 py-0.5 rounded-full font-semibold" style={{ background: STATUS_BAYAR_COLOR[p.statusBayar], color: THEME.charcoal }}>
                        {STATUS_BAYAR_LABEL[p.statusBayar]}
                      </span>
                    )}
                  </div>
                  {p.supplierName && <p className="att-body text-xs mt-0.5" style={{ color: THEME.inkSoft }}>{p.supplierName}</p>}
                  {p.gudangId && (
                    <p className="att-mono text-[10px]" style={{ color: THEME.inkSoft }}>&rarr; {gudangList.find((g) => g.id === p.gudangId)?.name || 'Gudang Utama'}</p>
                  )}
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <p className="att-mono text-sm font-bold" style={{ color: THEME.ink }}>{formatRupiah(purchaseTotal(p))}</p>
                  <button type="button" onClick={() => onDeletePurchase(p.id)}><Trash2 size={14} color={THEME.rust} /></button>
                </div>
              </div>
              <div className="space-y-0.5">
                {items.map((it, idx) => (
                  <p key={it.id || idx} className="att-mono text-[11px]" style={{ color: THEME.inkSoft }}>
                    &bull; {it.materialName}: {it.qty} {it.unit} &times; {formatRupiah(it.price)} = {formatRupiah(it.qty * it.price)}
                  </p>
                ))}
              </div>
              {p.note && <p className="att-mono text-[10px] mt-1" style={{ color: THEME.inkSoft }}>{p.note}</p>}
            </div>
          );
        })}
        {purchases.length === 0 && (
          <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada catatan belanja.</p>
        )}
      </div>

      {purchases.length > 0 && (
        <div className="mt-3 p-3 rounded-lg flex items-center justify-between" style={{ background: THEME.charcoal }}>
          <span className="att-body font-semibold text-sm" style={{ color: THEME.paper }}>TOTAL BELANJA</span>
          <span className="att-mono text-sm font-bold" style={{ color: THEME.amber }}>{formatRupiah(total)}</span>
        </div>
      )}
    </div>
  );
}

/* ---------------- Gudang: Stok & Pemakaian ---------------- */

function GudangTab({ materials, projects, purchases, usage, onAddUsage, onDeleteUsage, peralatan, onAddPeralatan, onUpdatePeralatan, onDeletePeralatan, peralatanUsage, onAddPeralatanUsage, onDeletePeralatanUsage, workers }) {
  const gudangList = getGudangList(projects);
  const [gudangId, setGudangId] = useState(GUDANG_UTAMA_ID);
  const [materialId, setMaterialId] = useState('');
  const [qty, setQty] = useState('');
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [note, setNote] = useState('');
  const [error, setError] = useState('');
  const [showTextPreview, setShowTextPreview] = useState(false);
  const [showPeralatan, setShowPeralatan] = useState(false);
  const [editingPeralatan, setEditingPeralatan] = useState(null);
  const [peralatanId, setPeralatanId] = useState('');
  const [pemakaiId, setPemakaiId] = useState('');
  const [peralatanDate, setPeralatanDate] = useState(new Date().toISOString().slice(0, 10));
  const [peralatanNote, setPeralatanNote] = useState('');
  const [peralatanError, setPeralatanError] = useState('');

  const stokPerMaterial = materials.map((m) => {
    const totalMasuk = purchases
      .filter((p) => p.gudangId === gudangId)
      .reduce((s, p) => s + purchaseItems(p).filter((it) => it.materialId === m.id).reduce((s2, it) => s2 + (Number(it.qty) || 0), 0), 0);
    const totalKeluar = usage.filter((u) => u.materialId === m.id && u.gudangId === gudangId)
      .reduce((s, u) => s + (Number(u.qty) || 0), 0);
    return { material: m, totalMasuk, totalKeluar, sisaStok: totalMasuk - totalKeluar };
  }).filter((x) => x.totalMasuk > 0 || x.totalKeluar > 0);

  const selectedMaterial = materials.find((m) => m.id === materialId);
  const stokTersedia = stokPerMaterial.find((x) => x.material.id === materialId)?.sisaStok ?? 0;

  const handleAddUsage = async () => {
    setError('');
    if (!materialId || !qty) {
      setError('Pilih material dan isi jumlah pemakaian.');
      return;
    }
    const qtyNum = Number(qty) || 0;
    await onAddUsage({
      id: uid(), date, gudangId, materialId, materialName: selectedMaterial?.name || '-', unit: selectedMaterial?.unit || '',
      qty: qtyNum, note: note.trim(),
    });
    setQty('');
    setNote('');
  };

  const gudangUsage = [...usage].filter((u) => u.gudangId === gudangId).sort((a, b) => (a.date < b.date ? 1 : -1));
  const text = [`STOK GUDANG - ${gudangList.find((g) => g.id === gudangId)?.name || ''}`, '================================',
    ...stokPerMaterial.map((x) => `${x.material.name}: masuk ${x.totalMasuk}, dipakai ${x.totalKeluar}, sisa ${x.sisaStok} ${x.material.unit || ''}`),
  ].join('\n');

  const peralatanGudang = peralatan.filter((x) => !x.gudangId || x.gudangId === gudangId);
  const selectedPeralatan = peralatan.find((x) => x.id === peralatanId);

  const handleAddPeralatanUsageSubmit = async () => {
    setPeralatanError('');
    if (!peralatanId || !pemakaiId) {
      setPeralatanError('Pilih peralatan dan pekerja yang menggunakan.');
      return;
    }
    const worker = workers.find((w) => w.id === pemakaiId);
    await onAddPeralatanUsage({
      id: uid(), date: peralatanDate, gudangId, peralatanId, peralatanNama: selectedPeralatan?.nama || '-',
      workerId: pemakaiId, workerName: worker?.name || '-', note: peralatanNote.trim(),
    });
    setPeralatanId('');
    setPemakaiId('');
    setPeralatanNote('');
  };

  const peralatanUsageGudang = [...peralatanUsage].filter((u) => u.gudangId === gudangId).sort((a, b) => (a.date < b.date ? 1 : -1));

  return (
    <div className="p-4 pb-24">
      <select value={gudangId} onChange={(e) => { setGudangId(e.target.value); setMaterialId(''); }}
        className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-3" style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }}>
        {gudangList.map((g) => <option key={g.id} value={g.id}>{g.name}</option>)}
      </select>

      <div className="flex items-center justify-between mb-2">
        <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>STOK SAAT INI</p>
        {stokPerMaterial.length > 0 && (
          <PrintMenu
            onPrint={(thermal) => openPrintDocument('Stok Gudang', `<pre style="white-space:pre-wrap;font-family:inherit;">${escapeHtml(text)}</pre>`, thermal)}
            onCopy={() => setShowTextPreview(true)}
          />
        )}
      </div>
      {showTextPreview && <TextPreviewModal title="Stok Gudang" text={text} onClose={() => setShowTextPreview(false)} />}

      <div className="space-y-2 mb-5">
        {stokPerMaterial.map(({ material, totalMasuk, totalKeluar, sisaStok }) => (
          <div key={material.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <p className="att-body font-semibold text-sm mb-1" style={{ color: THEME.ink }}>
              {material.name}{material.merk ? ` (${material.merk})` : ''}
              <span className="att-mono text-[10px] ml-1.5" style={{ color: THEME.inkSoft }}>{KATEGORI_LABELS[material.kategori] || KATEGORI_LABELS.non_kategori}</span>
            </p>
            <div className="grid grid-cols-3 gap-2 att-mono text-[11px]" style={{ color: THEME.inkSoft }}>
              <div className="p-1.5 rounded text-center" style={{ background: THEME.concrete }}>
                <p style={{ color: THEME.ink }} className="font-semibold">{totalMasuk} {material.unit}</p>
                Masuk
              </div>
              <div className="p-1.5 rounded text-center" style={{ background: THEME.concrete }}>
                <p style={{ color: THEME.ink }} className="font-semibold">{totalKeluar} {material.unit}</p>
                Dipakai
              </div>
              <div className="p-1.5 rounded text-center" style={{ background: THEME.concrete }}>
                <p style={{ color: sisaStok > 0 ? THEME.green : THEME.rust }} className="font-semibold">{sisaStok} {material.unit}</p>
                Sisa Stok
              </div>
            </div>
          </div>
        ))}
        {stokPerMaterial.length === 0 && (
          <p className="att-body text-sm text-center py-4" style={{ color: THEME.inkSoft }}>Belum ada belanja masuk ke gudang ini.</p>
        )}
      </div>

      <div className="p-4 rounded-lg mb-5" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
        <p className="att-body font-semibold text-sm mb-3" style={{ color: THEME.ink }}>Catat Pemakaian</p>
        <div className="grid sm:grid-cols-2 gap-3">
          <select value={materialId} onChange={(e) => setMaterialId(e.target.value)}
            className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }}>
            <option value="">Pilih material...</option>
            {stokPerMaterial.map(({ material, sisaStok }) => (
              <option key={material.id} value={material.id}>{material.name} (sisa: {sisaStok} {material.unit})</option>
            ))}
          </select>
          <input type="date" value={date} onChange={(e) => setDate(e.target.value)}
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={qty} onChange={(e) => setQty(e.target.value.replace(/[^0-9.]/g, ''))} placeholder="Jumlah dipakai/Qty" inputMode="decimal"
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="Untuk keperluan apa (opsional)"
            className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>
        {materialId && Number(qty) > stokTersedia && (
          <p className="text-xs mt-2" style={{ color: THEME.rust }}>Jumlah melebihi sisa stok ({stokTersedia} {selectedMaterial?.unit}). Tetap bisa disimpan, tapi periksa kembali catatan belanja.</p>
        )}
        {error && <p className="text-sm mt-2" style={{ color: THEME.rust }}>{error}</p>}
        <button type="button" onClick={handleAddUsage}
          className="w-full mt-3 flex items-center justify-center gap-1.5 py-2 rounded att-body font-semibold text-sm"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Plus size={15} /> Catat Pemakaian
        </button>
      </div>

      <p className="att-mono text-xs mb-2" style={{ color: THEME.inkSoft }}>{gudangUsage.length} RIWAYAT PEMAKAIAN</p>
      <div className="space-y-2">
        {gudangUsage.map((u) => (
          <div key={u.id} className="p-3 rounded-lg flex items-center justify-between" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div>
              <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{u.materialName}</p>
              <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{u.date} &middot; {u.qty} {u.unit}{u.note ? ` - ${u.note}` : ''}</p>
            </div>
            <button type="button" onClick={() => onDeleteUsage(u.id)}><Trash2 size={14} color={THEME.rust} /></button>
          </div>
        ))}
        {gudangUsage.length === 0 && (
          <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada catatan pemakaian di gudang ini.</p>
        )}
      </div>

      <button type="button" onClick={() => setShowPeralatan(!showPeralatan)}
        className="w-full flex items-center justify-between px-3 py-2 rounded-lg mt-6 mb-3 att-body text-sm font-semibold"
        style={{ background: THEME.charcoal, color: THEME.paper }}>
        Data Peralatan ({peralatan.length})
        {showPeralatan ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
      </button>

      {showPeralatan && (
        <>
          <PeralatanForm
            onAdd={onAddPeralatan}
            onUpdate={async (id, patch) => { await onUpdatePeralatan(id, patch); setEditingPeralatan(null); }}
            editing={editingPeralatan}
            onCancelEdit={() => setEditingPeralatan(null)}
          />
          <div className="space-y-2 mb-5">
            {peralatan.map((x) => (
              <div key={x.id} className="p-3 rounded-lg flex items-center justify-between" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
                <div>
                  <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{x.nama}{x.merk ? ` (${x.merk})` : ''}</p>
                  <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>
                    {x.kapasitas ? `${x.kapasitas} · ` : ''}{x.jumlah} {x.unit} &middot; dibeli {x.waktuPembelian}
                    {x.sourceNoNota ? ` · Nota ${x.sourceNoNota}` : ''}
                  </p>
                </div>
                <div className="flex gap-1 shrink-0">
                  <button type="button" onClick={() => setEditingPeralatan(x)} className="p-2 rounded" style={{ background: THEME.concrete }}>
                    <Edit3 size={14} color={THEME.inkSoft} />
                  </button>
                  <button type="button" onClick={() => onDeletePeralatan(x.id)} className="p-2 rounded" style={{ background: THEME.concrete }}>
                    <Trash2 size={14} color={THEME.rust} />
                  </button>
                </div>
              </div>
            ))}
            {peralatan.length === 0 && (
              <p className="att-body text-sm text-center py-4" style={{ color: THEME.inkSoft }}>Belum ada peralatan. Tambahkan manual di atas, atau otomatis terisi saat belanja kategori "Peralatan" di tab Belanja.</p>
            )}
          </div>
        </>
      )}

      <div className="p-4 rounded-lg mb-5" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
        <p className="att-body font-semibold text-sm mb-3" style={{ color: THEME.ink }}>Catat Pemakaian Peralatan</p>
        <div className="grid sm:grid-cols-2 gap-3">
          <select value={peralatanId} onChange={(e) => setPeralatanId(e.target.value)}
            className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }}>
            <option value="">Pilih peralatan...</option>
            {peralatanGudang.map((x) => <option key={x.id} value={x.id}>{x.nama}{x.merk ? ` - ${x.merk}` : ''}</option>)}
          </select>
          <select value={pemakaiId} onChange={(e) => setPemakaiId(e.target.value)}
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }}>
            <option value="">Pekerja pemakai...</option>
            {workers.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
          </select>
          <input type="date" value={peralatanDate} onChange={(e) => setPeralatanDate(e.target.value)}
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={peralatanNote} onChange={(e) => setPeralatanNote(e.target.value)} placeholder="Catatan (opsional)"
            className="px-3 py-2 rounded att-body text-sm outline-none sm:col-span-2" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>
        {peralatanError && <p className="text-sm mt-2" style={{ color: THEME.rust }}>{peralatanError}</p>}
        <button type="button" onClick={handleAddPeralatanUsageSubmit}
          className="w-full mt-3 flex items-center justify-center gap-1.5 py-2 rounded att-body font-semibold text-sm"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Plus size={15} /> Catat Pemakaian Peralatan
        </button>
      </div>

      <p className="att-mono text-xs mb-2" style={{ color: THEME.inkSoft }}>{peralatanUsageGudang.length} RIWAYAT PEMAKAIAN PERALATAN</p>
      <div className="space-y-2">
        {peralatanUsageGudang.map((u) => (
          <div key={u.id} className="p-3 rounded-lg flex items-center justify-between" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div>
              <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{u.peralatanNama}</p>
              <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{u.date} &middot; dipakai {u.workerName}{u.note ? ` - ${u.note}` : ''}</p>
            </div>
            <button type="button" onClick={() => onDeletePeralatanUsage(u.id)}><Trash2 size={14} color={THEME.rust} /></button>
          </div>
        ))}
        {peralatanUsageGudang.length === 0 && (
          <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada catatan pemakaian peralatan di gudang ini.</p>
        )}
      </div>
    </div>
  );
}

/* ---------------- Rekap Proyek: Hutang/Cash & Pemakaian Material ---------------- */

function buildRekapProyekTSV(perGudang, utangPekerja, usageByMaterial) {
  const lines = [];
  lines.push('=== HUTANG & CASH PER PROYEK ===');
  lines.push(['Gudang', 'Cash (Rp)', 'Belum Lunas (Rp)'].join('\t'));
  perGudang.forEach((x) => {
    lines.push([x.gudang.name, x.totalCash, x.totalBelumLunas].join('\t'));
    x.supplierList.forEach((s) => lines.push([`  -> ${s.name}`, '', s.sisa].join('\t')));
  });
  lines.push('');
  lines.push('=== RIWAYAT NOTA ===');
  lines.push(['Gudang', 'Tanggal', 'No Nota', 'Suplier', 'Status', 'Kategori', 'Material', 'Qty', 'Satuan', 'Harga', 'Subtotal'].join('\t'));
  perGudang.forEach((x) => {
    x.notaList.forEach((p) => {
      purchaseItems(p).forEach((it) => {
        lines.push([
          x.gudang.name, p.date, p.noNota || '', p.supplierName || '', STATUS_BAYAR_LABEL[p.statusBayar] || '',
          KATEGORI_LABELS[it.kategori] || KATEGORI_LABELS.non_kategori, it.materialName, it.qty, it.unit || '',
          it.price, (Number(it.qty) || 0) * (Number(it.price) || 0),
        ].join('\t'));
      });
    });
  });
  lines.push('');
  lines.push('=== PEMAKAIAN MATERIAL/PERALATAN ===');
  lines.push(['Material', 'Kategori', 'Jumlah', 'Satuan'].join('\t'));
  usageByMaterial.forEach((x) => lines.push([x.material.name, KATEGORI_LABELS[x.material.kategori] || KATEGORI_LABELS.non_kategori, x.totalQty, x.material.unit || ''].join('\t')));
  lines.push('');
  lines.push('=== UTANG UPAH KE PEKERJA ===');
  lines.push(['Nama', 'Total Upah', 'Dibayar', 'Sisa (Utang)'].join('\t'));
  utangPekerja.forEach((r) => lines.push([r.worker.name, r.totalUpah, r.diterima, r.sisa].join('\t')));
  return lines.join('\n');
}

function exportRekapProyekExcel(perGudang, utangPekerja, usageByMaterial) {
  try {
    const wb = XLSX.utils.book_new();

    const hutangRows = [];
    perGudang.forEach((x) => {
      hutangRows.push({ Gudang: x.gudang.name, Cash: x.totalCash, 'Belum Lunas': x.totalBelumLunas });
      x.supplierList.forEach((s) => {
        hutangRows.push({ Gudang: `   -> ${s.name}`, Cash: '', 'Belum Lunas': s.sisa });
      });
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(hutangRows), 'Hutang-Cash');

    const notaRows = [];
    perGudang.forEach((x) => {
      x.notaList.forEach((p) => {
        purchaseItems(p).forEach((it) => {
          notaRows.push({
            Gudang: x.gudang.name, Tanggal: p.date, 'No Nota': p.noNota || '', Suplier: p.supplierName || '',
            Status: STATUS_BAYAR_LABEL[p.statusBayar] || '', Kategori: KATEGORI_LABELS[it.kategori] || KATEGORI_LABELS.non_kategori,
            Material: it.materialName, Qty: it.qty, Satuan: it.unit, Harga: it.price, Subtotal: (Number(it.qty) || 0) * (Number(it.price) || 0),
          });
        });
      });
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(notaRows), 'Riwayat Nota');

    const usageRows = usageByMaterial.map((x) => ({ Material: x.material.name, Kategori: KATEGORI_LABELS[x.material.kategori] || KATEGORI_LABELS.non_kategori, Jumlah: x.totalQty, Satuan: x.material.unit }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(usageRows), 'Pemakaian Material');

    const pekerjaRows = utangPekerja.map((r) => ({ Nama: r.worker.name, 'Total Upah': r.totalUpah, Dibayar: r.diterima, 'Sisa (Utang)': r.sisa }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pekerjaRows), 'Utang Pekerja');

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'rekap-proyek.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  } catch (err) {
    console.error('Gagal ekspor Excel:', err);
    alert('Gagal membuat file Excel. Coba lagi, atau gunakan tombol Cetak (bisa disimpan sebagai PDF lewat dialog cetak).');
  }
}

function exportRekapProyekPDF(perGudang, utangPekerja, usageByMaterial) {
  try {
    const doc = new window.jspdf.jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    let y = pdfAddHeader(doc, 'Rekap Proyek', 'Hutang, Cash, Material & Utang Pekerja');

    const section = (title, head, rows) => {
      if (y > 740) { doc.addPage(); y = 40; }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...PDF_NAVY);
      doc.text(title, 40, y); y += 8;
      doc.autoTable({
        startY: y, head: [head], body: rows, margin: { left: 40, right: 40 },
        styles: { fontSize: 8.5, textColor: PDF_INK },
        headStyles: { fillColor: PDF_NAVY, textColor: [242, 236, 217] },
        alternateRowStyles: { fillColor: [250, 248, 241] },
      });
      y = doc.lastAutoTable.finalY + 22;
    };

    const hutangRows = [];
    perGudang.forEach((x) => {
      hutangRows.push([x.gudang.name, formatRupiah(x.totalCash), formatRupiah(x.totalBelumLunas)]);
      x.supplierList.forEach((s) => hutangRows.push([`   → ${s.name}`, '', formatRupiah(s.sisa)]));
    });
    if (hutangRows.length) section('Hutang & Cash per Gudang', ['Gudang', 'Cash', 'Belum Lunas'], hutangRows);

    const usageRows = usageByMaterial.map((x) => [x.material.name, KATEGORI_LABELS[x.material.kategori] || KATEGORI_LABELS.non_kategori, `${x.totalQty} ${x.material.unit}`]);
    if (usageRows.length) section('Pemakaian Material', ['Material', 'Kategori', 'Jumlah'], usageRows);

    const pekerjaRows = utangPekerja.map((r) => [r.worker.name, formatRupiah(r.totalUpah), formatRupiah(r.diterima), formatRupiah(r.sisa)]);
    if (pekerjaRows.length) section('Utang Pekerja', ['Nama', 'Total Upah', 'Dibayar', 'Sisa (Utang)'], pekerjaRows);

    pdfFinishAllPages(doc);
    doc.save('rekap-proyek.pdf');
  } catch (err) {
    console.error('Gagal ekspor PDF:', err);
    alert('Gagal membuat file PDF. Coba lagi, atau gunakan tombol Cetak (bisa disimpan sebagai PDF lewat dialog cetak).');
  }
}

/* ---- Helper bersama: header elegan navy-emas + watermark, dipakai semua dokumen PDF ---- */
const PDF_NAVY = [13, 25, 48];
const PDF_GOLD = [201, 162, 39];
const PDF_INK = [40, 40, 40];
const PDF_SLATE = [110, 116, 130];

function pdfAddWatermark(doc) {
  const w = doc.internal.pageSize.getWidth();
  const h = doc.internal.pageSize.getHeight();
  doc.saveGraphicsState();
  try { doc.setGState(new doc.GState({ opacity: 0.07 })); } catch { /* fallback tanpa transparansi kalau tidak didukung */ }
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(46);
  doc.setTextColor(...PDF_GOLD);
  const label = (CURRENT_COMPANY?.name || 'PATTIMURA UTAMA').toUpperCase();
  doc.text(label, w / 2, h / 2, { angle: 33, align: 'center' });
  doc.restoreGraphicsState();
}

function pdfAddHeader(doc, title, subtitle) {
  const w = doc.internal.pageSize.getWidth();
  doc.setFillColor(...PDF_NAVY);
  doc.rect(0, 0, w, 72, 'F');
  doc.setDrawColor(...PDF_GOLD);
  doc.setLineWidth(2);
  doc.line(0, 72, w, 72);

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(15);
  doc.setTextColor(242, 236, 217);
  doc.text(CURRENT_COMPANY?.name || 'ABSENSI TUKANG', 40, 30);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8.5);
  doc.setTextColor(...PDF_GOLD);
  const infoLine = [CURRENT_COMPANY?.address, CURRENT_COMPANY?.phone ? `Telp: ${CURRENT_COMPANY.phone}` : ''].filter(Boolean).join('  |  ');
  if (infoLine) doc.text(infoLine, 40, 43);

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(11);
  doc.setTextColor(...PDF_GOLD);
  doc.text(title.toUpperCase(), w - 40, 30, { align: 'right' });
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8);
  doc.setTextColor(200, 200, 200);
  if (subtitle) doc.text(subtitle, w - 40, 43, { align: 'right' });
  doc.text(`Dicetak: ${new Date().toLocaleDateString('id-ID')}`, w - 40, 55, { align: 'right' });

  return 96; // posisi y untuk mulai konten
}

function pdfFinishAllPages(doc) {
  const pages = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pages; i += 1) {
    doc.setPage(i);
    pdfAddWatermark(doc);
    const w = doc.internal.pageSize.getWidth();
    const h = doc.internal.pageSize.getHeight();
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.setTextColor(...PDF_SLATE);
    doc.text(`Halaman ${i} / ${pages}`, w - 40, h - 20, { align: 'right' });
  }
}

function exportIdCardPDF(worker) {
  try {
    const doc = new window.jspdf.jsPDF({ orientation: 'portrait', unit: 'pt', format: [230, 360] });
    const w = 230, h = 360;

    doc.setFillColor(...PDF_NAVY);
    doc.rect(0, 0, w, h, 'F');
    doc.setFillColor(250, 248, 241);
    doc.roundedRect(6, 92, w - 12, h - 98, 6, 6, 'F');

    // Header: logo + nama perusahaan
    try {
      const logo = CURRENT_COMPANY.logoDataUri || COMPANY_LOGO_DATA_URI;
      doc.addImage(logo, 'JPEG', w / 2 - 18, 14, 36, 36);
    } catch { /* lewati kalau format logo tidak didukung addImage */ }
    doc.setFont('helvetica', 'bold'); doc.setFontSize(10.5); doc.setTextColor(242, 236, 217);
    doc.text(CURRENT_COMPANY?.name || 'ABSENSI TUKANG', w / 2, 62, { align: 'center', maxWidth: w - 24 });
    if (CURRENT_COMPANY?.tagline) {
      doc.setFont('helvetica', 'normal'); doc.setFontSize(6); doc.setTextColor(...PDF_GOLD);
      doc.text(CURRENT_COMPANY.tagline.toUpperCase(), w / 2, 72, { align: 'center', maxWidth: w - 24 });
    }
    doc.setDrawColor(...PDF_GOLD); doc.setLineWidth(2.5);
    doc.line(0, 88, w, 88);

    // Foto pekerja
    const photoSize = 84;
    const photoX = w / 2 - photoSize / 2;
    const photoY = 106;
    doc.setDrawColor(...PDF_GOLD); doc.setLineWidth(2);
    if (worker.photoDataUri) {
      try { doc.addImage(worker.photoDataUri, 'JPEG', photoX, photoY, photoSize, photoSize); } catch { /* skip jika gagal */ }
    } else {
      doc.setFillColor(239, 233, 216); doc.rect(photoX, photoY, photoSize, photoSize, 'F');
    }
    doc.rect(photoX, photoY, photoSize, photoSize);

    doc.setFont('helvetica', 'bold'); doc.setFontSize(12); doc.setTextColor(...PDF_NAVY);
    doc.text(worker.name, w / 2, photoY + photoSize + 20, { align: 'center', maxWidth: w - 24 });
    doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(...PDF_GOLD);
    doc.text((worker.position || '-').toUpperCase(), w / 2, photoY + photoSize + 32, { align: 'center' });

    let iy = photoY + photoSize + 46;
    doc.setDrawColor(...PDF_GOLD); doc.setLineWidth(1.5);
    doc.line(w / 2 - 28, iy, w / 2 + 28, iy);
    iy += 14;

    const infoRows = [];
    if (worker.joinDate) infoRows.push(['Bergabung', worker.joinDate]);
    if (worker.phone) infoRows.push(['No. HP', worker.phone]);
    if (worker.ktp) infoRows.push(['No. KTP', worker.ktp]);
    doc.setFontSize(7.5);
    infoRows.forEach(([lbl, val]) => {
      doc.setFont('helvetica', 'normal'); doc.setTextColor(...PDF_SLATE);
      doc.text(lbl, 18, iy);
      doc.setFont('helvetica', 'bold'); doc.setTextColor(...PDF_INK);
      doc.text(val, w - 18, iy, { align: 'right' });
      iy += 13;
    });

    doc.setFillColor(...PDF_NAVY);
    doc.rect(0, h - 22, w, 22, 'F');
    doc.setFont('helvetica', 'normal'); doc.setFontSize(6.5); doc.setTextColor(...PDF_GOLD);
    doc.text(CURRENT_COMPANY?.website || CURRENT_COMPANY?.phone || '', w / 2, h - 9, { align: 'center' });

    doc.save(`id-card-${worker.name.replace(/\s+/g, '-').toLowerCase()}.pdf`);
  } catch (err) {
    console.error('Gagal ekspor ID Card PDF:', err);
    alert('Gagal membuat file PDF. Coba lagi, atau gunakan tombol Cetak.');
  }
}

function exportRekapUpahPDF(rows, grand, filterLabel) {
  try {
    const doc = new window.jspdf.jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    let y = pdfAddHeader(doc, 'Rekap Upah Tukang', filterLabel);

    const body = rows.map((r) => [
      r.worker.name, r.worker.position || '-', `${r.totalHari} hr / ${r.totalJamLembur} j`,
      formatRupiah(r.totalUpah), formatRupiah(r.diterima), formatRupiah(r.totalKasbon), formatRupiah(r.sisa),
    ]);
    doc.autoTable({
      startY: y, margin: { left: 40, right: 40 },
      head: [['Nama', 'Jabatan', 'Hari/Lembur', 'Total Upah', 'Diterima', 'Sisa Kasbon', 'Sisa Upah']],
      body,
      foot: [['', '', 'TOTAL', formatRupiah(grand.totalUpah), formatRupiah(grand.diterima), formatRupiah(grand.totalKasbon), formatRupiah(grand.sisa)]],
      styles: { fontSize: 8, textColor: PDF_INK },
      headStyles: { fillColor: PDF_NAVY, textColor: [242, 236, 217] },
      footStyles: { fillColor: [241, 236, 221], textColor: PDF_NAVY, fontStyle: 'bold' },
      alternateRowStyles: { fillColor: [250, 248, 241] },
    });

    pdfFinishAllPages(doc);
    doc.save('rekap-upah.pdf');
  } catch (err) {
    console.error('Gagal ekspor PDF Rekap Upah:', err);
    alert('Gagal membuat file PDF. Coba lagi, atau gunakan tombol Cetak.');
  }
}

function exportSlipGajiPDF(row, filterLabel) {
  try {
    const { worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown, evidenceFilled, evidenceExpected } = row;
    const doc = new window.jspdf.jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    let y = pdfAddHeader(doc, 'Slip Gaji', filterLabel);

    doc.setFont('helvetica', 'bold'); doc.setFontSize(12); doc.setTextColor(...PDF_NAVY);
    doc.text(worker.name, 40, y);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...PDF_SLATE);
    doc.text(worker.position || '-', 40, y + 14);
    y += 34;

    // Ringkasan
    const summary = [
      ['Hari Kerja', `${totalHari} hari`], ['Jam Lembur', `${totalJamLembur} jam`],
      ['Total Upah', formatRupiah(totalUpah)], ['Sudah Diterima', formatRupiah(diterima)],
      ['Sisa Kasbon', formatRupiah(totalKasbon)], ['Sisa Upah', formatRupiah(sisa)],
    ];
    if (evidenceExpected > 0) summary.push(['Bukti Foto Kehadiran', `${evidenceFilled}/${evidenceExpected} periode`]);
    doc.autoTable({
      startY: y, margin: { left: 40, right: 40 }, theme: 'plain',
      body: summary.map(([a, b]) => [a, b]),
      styles: { fontSize: 9.5, textColor: PDF_INK, cellPadding: { top: 4, bottom: 4, left: 0, right: 0 } },
      columnStyles: { 0: { textColor: PDF_SLATE, cellWidth: 150 }, 1: { fontStyle: 'bold', textColor: PDF_NAVY } },
    });
    y = doc.lastAutoTable.finalY + 20;

    const section = (title, head, rows2) => {
      if (!rows2.length) return;
      if (y > 700) { doc.addPage(); y = 40; }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10.5); doc.setTextColor(...PDF_NAVY);
      doc.text(title, 40, y); y += 6;
      doc.autoTable({
        startY: y, margin: { left: 40, right: 40 }, head: [head], body: rows2,
        styles: { fontSize: 8, textColor: PDF_INK },
        headStyles: { fillColor: [241, 236, 221], textColor: PDF_NAVY },
        alternateRowStyles: { fillColor: [250, 248, 241] },
      });
      y = doc.lastAutoTable.finalY + 20;
    };

    section('Rincian per Periode', ['Periode', 'Proyek', 'Hari', 'Lembur (j)', 'Upah'],
      weekBreakdown.map((wk) => [wk.weekLabel, wk.projectName, String(wk.totalHariBayar), String(wk.totalJamLembur), formatRupiah(wk.totalUpah)]));

    section('Riwayat Pembayaran Upah', ['Tanggal', 'Jumlah'],
      workerPayments.map((p) => [p.date, formatRupiah(p.amount)]));

    section('Riwayat Kasbon Diberikan', ['Tanggal', 'Jumlah', 'Catatan'],
      workerKasbon.map((k) => [k.date, formatRupiah(k.amount), k.note || '']));

    section('Riwayat Pelunasan Kasbon', ['Tanggal', 'Jumlah'],
      workerKasbonPayments.map((p) => [p.date, formatRupiah(p.amount)]));

    pdfFinishAllPages(doc);
    doc.save(`slip-gaji-${worker.name.replace(/\s+/g, '-').toLowerCase()}.pdf`);
  } catch (err) {
    console.error('Gagal ekspor PDF Slip Gaji:', err);
    alert('Gagal membuat file PDF. Coba lagi, atau gunakan tombol Cetak.');
  }
}

/* ---------------- AHSP (Analisa Harga Satuan Pekerjaan) ---------------- */

function ahspBaseHarga(komponen) {
  return (komponen || []).reduce((s, k) => s + (Number(k.koefisien) || 0) * (Number(k.hargaSatuan) || 0), 0);
}
function ahspFinalHarga(komponen, overheadProfitPercent) {
  const base = ahspBaseHarga(komponen);
  return base * (1 + (Number(overheadProfitPercent) || 0) / 100);
}
function emptyAhspKomponen() {
  return { id: uid(), jenis: 'bahan', nama: '', koefisien: '', hargaSatuan: '' };
}
const AHSP_JENIS_LABEL = { bahan: 'Bahan', upah: 'Upah', alat: 'Alat' };
const AHSP_SUMBER_LIST = ['AHSP Cipta Karya', 'AHSP Bina Marga', 'AHSP SDA', 'AHSP SNI', 'Custom'];
const UPAH_KATEGORI = [
  { key: 'pekerja', label: 'Pekerja' },
  { key: 'tukang', label: 'Tukang' },
  { key: 'tukangBatu', label: 'Tukang Batu' },
  { key: 'tukangKayu', label: 'Tukang Kayu' },
  { key: 'tukangBesi', label: 'Tukang Besi' },
  { key: 'tukangCat', label: 'Tukang Cat' },
  { key: 'mandor', label: 'Mandor' },
  { key: 'kepalaTukang', label: 'Kepala Tukang' },
  { key: 'penjagaMalam', label: 'Penjaga Malam' },
];

const HSPK_PALEMBANG_2026 = [
  { no: 1, nama: 'Bata Merah', satuan: 'Buah', harga: 1100 },
  { no: 2, nama: 'Batu Bata / Merah', satuan: 'Buah', harga: 1500 },
  { no: 3, nama: 'Sok Pipa Pvc 1 Inci', satuan: 'Buah', harga: 5400 },
  { no: 4, nama: 'Karung', satuan: 'Lembar', harga: 5800 },
  { no: 5, nama: 'Seal Tape', satuan: 'Buah', harga: 7700 },
  { no: 6, nama: 'Sdl Pipa Pvc 1 Inci', satuan: 'Buah', harga: 10900 },
  { no: 7, nama: 'Kuas', satuan: 'Buah', harga: 12300 },
  { no: 8, nama: 'Ring Nepple 1/2 Inci', satuan: 'Buah', harga: 16100 },
  { no: 9, nama: 'Paku Beton 0,5-3 Cm', satuan: 'Kotak', harga: 22500 },
  { no: 10, nama: 'Amplas Ukuran 9X11 Merk 3M', satuan: 'Lembar', harga: 12800 },
  { no: 11, nama: 'Ember Bangunan', satuan: 'Buah', harga: 16500 },
  { no: 12, nama: 'Baja Tulangan', satuan: 'Kg', harga: 30600 },
  { no: 13, nama: 'Lem Pipa 60 Gr', satuan: 'Botol', harga: 21200 },
  { no: 14, nama: 'Kawat Beton', satuan: 'Kg', harga: 30500 },
  { no: 15, nama: 'Paku Beton 4-7Cm', satuan: 'Kotak', harga: 30300 },
  { no: 16, nama: 'Centong Bangunan', satuan: 'Buah', harga: 30000 },
  { no: 17, nama: 'Stop Kran 1/2', satuan: 'Buah', harga: 29000 },
  { no: 18, nama: 'Paku', satuan: 'Kg', harga: 27000 },
  { no: 19, nama: 'Baja', satuan: 'Kg', harga: 30000 },
  { no: 20, nama: 'Besi Cor', satuan: 'Kg', harga: 33000 },
  { no: 21, nama: 'Thiner', satuan: 'Kg', harga: 33200 },
  { no: 22, nama: 'Thiner', satuan: 'Botol', harga: 29000 },
  { no: 23, nama: 'Thiner', satuan: 'Kaleng', harga: 37000 },
  { no: 24, nama: 'Stop Kran 3/4', satuan: 'Buah', harga: 31000 },
  { no: 25, nama: 'Pipa Pvc Aw 1/2 Per 4M', satuan: 'Batang', harga: 39000 },
  { no: 26, nama: 'Stop Kran 1"', satuan: 'Buah', harga: 38400 },
  { no: 27, nama: 'Meteran', satuan: 'Unit', harga: 40400 },
  { no: 28, nama: 'Baja Tulangan', satuan: 'Kg', harga: 21700 },
  { no: 29, nama: 'Ampelas', satuan: 'Lembar', harga: 19000 },
  { no: 30, nama: 'Besi Beton Polos 6 Mm -Hjj, Hanil,Ks (Sni)', satuan: 'Batang', harga: 49500 },
  { no: 31, nama: 'Kawat No.12', satuan: 'Kg', harga: 40000 },
  { no: 32, nama: 'Material Aspal-Aspal', satuan: 'Liter', harga: 34000 },
  { no: 33, nama: 'Kawat Las Listrik', satuan: 'Kg', harga: 84000 },
  { no: 34, nama: 'Paku Beton 10-12,5 Cm', satuan: 'Kotak', harga: 41000 },
  { no: 35, nama: 'Paku Beton', satuan: 'Kotak', harga: 45000 },
  { no: 36, nama: 'Paku Biasa', satuan: 'Kg', harga: 42300 },
  { no: 37, nama: 'Benang Bangunan', satuan: 'Gulung', harga: 35000 },
  { no: 38, nama: 'Meni,Plamir,Cat Dasar/0,9 Liter', satuan: 'Kaleng', harga: 67000 },
  { no: 39, nama: 'Pipa Pvc Aw 3/4 Per 4M', satuan: 'Batang', harga: 55000 },
  { no: 40, nama: 'Meni Besi', satuan: 'Kg', harga: 71200 },
  { no: 41, nama: 'Besi Hollow', satuan: 'Batang', harga: 43000 },
  { no: 42, nama: 'Lem Pipa 400 Gr', satuan: 'Kaleng', harga: 70000 },
  { no: 43, nama: 'Dempul', satuan: 'Kg', harga: 78000 },
  { no: 44, nama: 'Pipa Pvc Aw 1 Per 4M', satuan: 'Batang', harga: 78000 },
  { no: 45, nama: 'Centong/Roskam Bangunan', satuan: 'Buah', harga: 66000 },
  { no: 46, nama: 'Besi Beton Polos 8 Mm -Hjj, Hanil,Ks (Sni)', satuan: 'Batang', harga: 71000 },
  { no: 47, nama: 'Fiber Pagar', satuan: 'Roll', harga: 62000 },
  { no: 48, nama: 'Pipa Pvc 1 Inci', satuan: 'Buah', harga: 70000 },
  { no: 49, nama: 'Cat Batu Candi / Liter', satuan: 'Kaleng', harga: 130000 },
  { no: 50, nama: 'Cat Tembok', satuan: 'Kg', harga: 91000 },
  { no: 51, nama: 'Material Bahan Lainnya-Semen', satuan: 'Zak', harga: 97000 },
  { no: 52, nama: 'Besi Hollow', satuan: 'Batang', harga: 109000 },
  { no: 53, nama: 'Wd 40', satuan: 'Buah', harga: 107000 },
  { no: 54, nama: 'Pas Bata & Plester', satuan: 'Zak', harga: 104000 },
  { no: 55, nama: 'Pengisi Nat Keramik', satuan: 'Kg', harga: 74000 },
  { no: 56, nama: 'Material Bahan Lainnya-Cat Besi', satuan: 'Kg', harga: 111000 },
  { no: 57, nama: 'Thinner', satuan: 'Kaleng', harga: 99000 },
  { no: 58, nama: 'Besi Siku', satuan: 'Batang', harga: 113000 },
  { no: 59, nama: 'Cat Tembok Exterior / Cat Tembok Interior', satuan: 'Kg', harga: 69000 },
  { no: 60, nama: 'Sarana Dan Prasarana Air Mancur-Pipa Pvc Aw 1 1/4', satuan: 'Batang', harga: 103000 },
  { no: 61, nama: 'Sarana Dan Prasarana Air Mancur-Pipa Pvc Aw 1 1/2', satuan: 'Batang', harga: 115000 },
  { no: 62, nama: 'Cat Kusen & Besi', satuan: 'Kg', harga: 102000 },
  { no: 63, nama: 'Sarana Dan Prasarana Air Mancur-Pipa Pvc Aw 2', satuan: 'Batang', harga: 169000 },
  { no: 64, nama: 'Cat Anti Karat Untuk Seng /Meta Galvanis (Water - Based)1 Liter', satuan: 'Kaleng', harga: 165000 },
  { no: 65, nama: 'Paving Blok', satuan: 'M2', harga: 191000 },
  { no: 66, nama: 'Lem Pipa', satuan: 'Buah', harga: 134000 },
  { no: 67, nama: 'Material Agregat/Batu Pecah-Agregat Base Kelas A', satuan: 'M3', harga: 370000 },
  { no: 68, nama: 'Pipa Pvc Tee Elbow', satuan: 'Buah', harga: 172000 },
  { no: 69, nama: 'Meteran Kecil', satuan: 'Buah', harga: 117000 },
  { no: 70, nama: 'Pasir Urug', satuan: 'M3', harga: 243000 },
  { no: 71, nama: 'Kayu Bakar', satuan: 'Mobil', harga: 275000 },
  { no: 72, nama: 'Besi Hollow', satuan: 'Batang', harga: 224000 },
  { no: 73, nama: 'Pipa Pvc Aw 2,5', satuan: 'Batang', harga: 203000 },
  { no: 74, nama: 'Kawat Las', satuan: 'Pak', harga: 230000 },
  { no: 75, nama: 'Gembok Sherlock 40 Mm', satuan: 'Buah', harga: 281000 },
  { no: 76, nama: 'Portland Semen', satuan: 'Kg', harga: 119000 },
  { no: 77, nama: 'Selang Spiral 3 Inci', satuan: 'Buah', harga: 183000 },
  { no: 78, nama: 'Slang Bangunan', satuan: 'Buah', harga: 218000 },
  { no: 79, nama: 'Besi Beton Polos12 Mm -Hjj, Hanil,Ks (Sni)', satuan: 'Batang', harga: 183000 },
  { no: 80, nama: 'Besi Cnp', satuan: 'Batang', harga: 480000 },
  { no: 81, nama: 'Kakak Tua (Catut)', satuan: 'Buah', harga: 108000 },
  { no: 82, nama: 'Pasir Halus (Untuk Hrs)', satuan: 'M3', harga: 284000 },
  { no: 83, nama: 'Waterpass', satuan: 'Buah', harga: 282000 },
  { no: 84, nama: 'Besi Hollow', satuan: 'Batang', harga: 252000 },
  { no: 85, nama: 'Besi Cnp', satuan: 'Batang', harga: 367500 },
  { no: 86, nama: 'Besi Unp', satuan: 'Batang', harga: 338000 },
  { no: 87, nama: 'Pelat Besi/Baja 3 Inc', satuan: 'Meter', harga: 288000 },
  { no: 88, nama: 'Besi Siku', satuan: 'Batang', harga: 280000 },
  { no: 89, nama: 'Kunci Pintu', satuan: 'Set', harga: 474000 },
  { no: 90, nama: 'Pipa Pvc Aw 3', satuan: 'Batang', harga: 281400 },
  { no: 91, nama: 'Pasir', satuan: 'Pick Up', harga: 222000 },
  { no: 92, nama: 'Besi Unp', satuan: 'Batang', harga: 310500 },
  { no: 93, nama: 'Acian Plester Dan Beton', satuan: 'Kg', harga: 285500 },
  { no: 94, nama: 'Besi Cnp', satuan: 'Batang', harga: 456000 },
  { no: 95, nama: 'Besi Hollow', satuan: 'Batang', harga: 315000 },
  { no: 96, nama: 'Batu Kali/Belah', satuan: 'M3', harga: 495000 },
  { no: 97, nama: 'Tangkat Lampu Rambu Kerja', satuan: 'Buah', harga: 266000 },
  { no: 98, nama: 'Meteran Besar', satuan: 'Buah', harga: 272000 },
  { no: 99, nama: 'Cat Tembok Exterior/Interior', satuan: 'Galon', harga: 314000 },
  { no: 100, nama: 'Besi Cnp', satuan: 'Batang', harga: 436000 },
  { no: 101, nama: 'Acian Profil', satuan: 'Zak', harga: 260000 },
  { no: 102, nama: 'Besi Cnp', satuan: 'Batang', harga: 467000 },
  { no: 103, nama: 'Buis Beton', satuan: 'Buah', harga: 429000 },
  { no: 104, nama: 'Pasir Beton', satuan: 'M3', harga: 434000 },
  { no: 105, nama: 'Sarana Dan Prasarana Air Mancur-Pipa Pvc Aw 4', satuan: 'Batang', harga: 451000 },
  { no: 106, nama: 'Material Agregat/Batu Pecah-Koral/Kerikil/Agregat Beton', satuan: 'M3', harga: 518000 },
  { no: 107, nama: 'Besi Siku', satuan: 'Batang', harga: 470000 },
  { no: 108, nama: 'Pasir Pasang', satuan: 'M3', harga: 430000 },
  { no: 109, nama: 'Besi Beton Polos', satuan: 'Batang', harga: 492000 },
  { no: 110, nama: 'Material Bahan Lainnya-Pasir', satuan: 'M3', harga: 400000 },
  { no: 111, nama: 'Besi Unp', satuan: 'Batang', harga: 494000 },
  { no: 112, nama: 'Pipa Pvc Elbow', satuan: 'Buah', harga: 286000 },
  { no: 113, nama: 'Atap Bitumen Transparant', satuan: 'Pcs', harga: 562000 },
  { no: 114, nama: 'Besi Siku', satuan: 'Batang', harga: 633000 },
  { no: 115, nama: 'Pasir Silika', satuan: 'Zak', harga: 412000 },
  { no: 116, nama: 'Pasir Karbon Aktif', satuan: 'Zak', harga: 506000 },
  { no: 117, nama: 'Pasir Zeolite', satuan: 'Zak', harga: 362000 },
  { no: 118, nama: 'Perbaikan Permukaan Beton', satuan: 'Zak', harga: 420000 },
  { no: 119, nama: 'Besi Siku', satuan: 'Batang', harga: 661000 },
  { no: 120, nama: 'Besi Unp', satuan: 'Batang', harga: 757000 },
  { no: 121, nama: 'Pipa Pvc', satuan: 'Buah', harga: 685000 },
  { no: 122, nama: 'Mesin Pompa Air', satuan: 'Unit', harga: 825000 },
  { no: 123, nama: 'Box Culvert', satuan: 'Buah', harga: 995000 },
  { no: 124, nama: 'Besi Unp', satuan: 'Batang', harga: 1000000 },
  { no: 125, nama: 'Kerikil/Koral/Agregat Beton', satuan: 'M3', harga: 685000 },
  { no: 126, nama: 'Besi Plat', satuan: 'Lembar', harga: 818000 },
  { no: 127, nama: 'Besi Plat', satuan: 'Keping', harga: 858000 },
  { no: 128, nama: 'Batu Split', satuan: 'M2', harga: 692000 },
  { no: 129, nama: 'Batu / Batu Kali/ Batu Belah', satuan: 'M3', harga: 845000 },
  { no: 130, nama: 'Agregat Base Kelas A/B/C', satuan: 'M3', harga: 750000 },
  { no: 131, nama: 'Keramik', satuan: 'M2', harga: 695000 },
  { no: 132, nama: 'Troli / Gerobak Sorong', satuan: 'Buah', harga: 978000 },
  { no: 133, nama: 'Baja Pelat', satuan: 'M2', harga: 1005000 },
  { no: 134, nama: 'Material Beton Ready Mix-Ready Mix', satuan: 'M3', harga: 1209000 },
  { no: 135, nama: 'Besi Unp', satuan: 'Batang', harga: 1357000 },
  { no: 136, nama: 'Besi Plat', satuan: 'Lembar', harga: 1261000 },
  { no: 137, nama: 'Pipa Pvc Untuk Air Limbah', satuan: 'M3', harga: 1342000 },
  { no: 138, nama: 'Allumunium Pelat', satuan: 'Meter', harga: 1306000 },
  { no: 139, nama: 'Selang Air Spiral 8 Inch', satuan: 'Roll', harga: 1874000 },
  { no: 140, nama: 'Pipa Baja Lainnya', satuan: 'Buah', harga: 1587000 },
  { no: 141, nama: 'Ready Mix Cor Beton', satuan: 'M3', harga: 1976000 },
  { no: 142, nama: 'Besi Plat', satuan: 'Buah', harga: 1713000 },
  { no: 143, nama: 'Besi Plat', satuan: 'Keping', harga: 1717000 },
  { no: 144, nama: 'Besi Unp', satuan: 'Batang', harga: 2023000 },
  { no: 145, nama: 'Besi Plat', satuan: 'Lembar', harga: 2448000 },
  { no: 146, nama: 'Sarana Dan Prasarana Air Mancur-Pipa Pvc Aw 10', satuan: 'Batang', harga: 3042000 },
  { no: 147, nama: 'Cat Tembok', satuan: 'Liter', harga: 2152000 },
  { no: 148, nama: 'Besi Siku', satuan: 'Batang', harga: 2739000 },
  { no: 149, nama: 'Pipa Lainnya', satuan: 'Buah', harga: 5283000 },
  { no: 150, nama: 'Aspal Drum', satuan: 'Drum', harga: 2887000 },
  { no: 151, nama: 'Besi Wf', satuan: 'Batang', harga: 3334000 },
  { no: 152, nama: 'Kayu Racuk', satuan: 'M3', harga: 2066000 },
  { no: 153, nama: 'Besi Pipa', satuan: 'Meter', harga: 3230000 },
  { no: 154, nama: 'Besi Plat', satuan: 'Buah', harga: 5649000 },
  { no: 155, nama: 'Besi Plat', satuan: 'Buah', harga: 10513000 },
  { no: 156, nama: 'Besi UNP', satuan: 'Batang', harga: 638000 },
  { no: 157, nama: 'Besi UNP', satuan: 'Btg (6m)', harga: 920000 },
  { no: 158, nama: 'Besi UNP', satuan: 'Batang', harga: 1646000 },
  { no: 159, nama: 'Besi UNP', satuan: 'Batang', harga: 2070000 },
];

function AhspMasterForm({ onSave, onCancel, initial, overheadProfit, upahRates }) {
  const [kode, setKode] = useState(initial?.kode || '');
  const [uraian, setUraian] = useState(initial?.uraian || '');
  const [satuan, setSatuan] = useState(initial?.satuan || '');
  const [sumber, setSumber] = useState(initial?.sumber || AHSP_SUMBER_LIST[4]);
  const [komponen, setKomponen] = useState(initial?.komponen?.length ? initial.komponen : [emptyAhspKomponen()]);

  const updateK = (id, patch) => setKomponen(komponen.map((k) => (k.id === id ? { ...k, ...patch } : k)));
  const addK = () => setKomponen([...komponen, emptyAhspKomponen()]);
  const removeK = (id) => setKomponen(komponen.filter((k) => k.id !== id));
  const pickUpah = (id, kategoriKey) => {
    const kategori = UPAH_KATEGORI.find((u) => u.key === kategoriKey);
    if (!kategori) return;
    updateK(id, { nama: kategori.label, hargaSatuan: String(upahRates?.[kategoriKey] || 0) });
  };
  const pickBahan = (id, no) => {
    const item = HSPK_PALEMBANG_2026.find((h) => String(h.no) === no);
    if (!item) return;
    updateK(id, { nama: item.nama, hargaSatuan: String(item.harga) });
  };

  const base = ahspBaseHarga(komponen);
  const final = ahspFinalHarga(komponen, overheadProfit);

  const handleSave = async () => {
    if (!kode.trim() || !uraian.trim() || !satuan.trim()) {
      alert('Isi kode AHSP, uraian pekerjaan, dan satuan dulu.');
      return;
    }
    await onSave({
      id: initial?.id || uid(), kode: kode.trim(), uraian: uraian.trim(), satuan: satuan.trim(), sumber,
      komponen: komponen.filter((k) => k.nama.trim()).map((k) => ({ ...k, koefisien: Number(k.koefisien) || 0, hargaSatuan: Number(k.hargaSatuan) || 0 })),
    });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(13,25,48,0.7)' }}>
      <div className="w-full max-w-md rounded-xl p-4 att-body max-h-[92vh] overflow-y-auto" style={{ background: THEME.paper }}>
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>{initial ? 'Edit' : 'Tambah'} Item AHSP</h3>
          <button type="button" onClick={onCancel}><X size={18} color={THEME.inkSoft} /></button>
        </div>
        <div className="grid grid-cols-2 gap-2 mb-2">
          <input value={kode} onChange={(e) => setKode(e.target.value)} placeholder="Kode AHSP (mis. A.4.1.1.1)"
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={satuan} onChange={(e) => setSatuan(e.target.value)} placeholder="Satuan (m3, m2, dll)"
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>
        <select value={sumber} onChange={(e) => setSumber(e.target.value)}
          className="w-full px-3 py-2 rounded att-mono text-xs outline-none mb-2" style={{ background: THEME.concrete, color: THEME.ink }}>
          {AHSP_SUMBER_LIST.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <input value={uraian} onChange={(e) => setUraian(e.target.value)} placeholder="Uraian pekerjaan (mis. 1 m3 Beton Mutu K-225)"
          className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-3" style={{ background: THEME.concrete, color: THEME.ink }} />

        <p className="att-mono text-[10px] mb-2" style={{ color: THEME.inkSoft }}>STRUKTUR HARGA (BAHAN / UPAH / ALAT)</p>
        <div className="space-y-2 mb-2">
          {komponen.map((k) => (
            <div key={k.id} className="p-2.5 rounded-lg" style={{ background: THEME.concrete }}>
              <div className="flex gap-1.5 mb-1.5">
                <select value={k.jenis} onChange={(e) => updateK(k.id, { jenis: e.target.value })}
                  className="px-2 py-1.5 rounded att-mono text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink }}>
                  {Object.entries(AHSP_JENIS_LABEL).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                </select>
                <input value={k.nama} onChange={(e) => updateK(k.id, { nama: e.target.value })} placeholder="Nama item"
                  className="flex-1 px-2 py-1.5 rounded att-body text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink }} />
                <button type="button" onClick={() => removeK(k.id)} className="p-1.5 rounded shrink-0" style={{ background: THEME.paper }}>
                  <Trash2 size={13} color={THEME.rust} />
                </button>
              </div>
              {k.jenis === 'upah' && (
                <select value="" onChange={(e) => pickUpah(k.id, e.target.value)}
                  className="w-full px-2 py-1.5 rounded att-mono text-[10.5px] outline-none mb-1.5" style={{ background: THEME.paper, color: THEME.inkSoft }}>
                  <option value="">— Isi cepat dari Data Upah —</option>
                  {UPAH_KATEGORI.map((u) => <option key={u.key} value={u.key}>{u.label} ({formatRupiah(upahRates?.[u.key] || 0)})</option>)}
                </select>
              )}
              {k.jenis === 'bahan' && (
                <select value="" onChange={(e) => pickBahan(k.id, e.target.value)}
                  className="w-full px-2 py-1.5 rounded att-mono text-[10.5px] outline-none mb-1.5" style={{ background: THEME.paper, color: THEME.inkSoft }}>
                  <option value="">— Isi cepat dari HSPK Palembang 2026 —</option>
                  {HSPK_PALEMBANG_2026.map((h) => <option key={h.no} value={h.no}>{h.nama} — {formatRupiah(h.harga)}/{h.satuan}</option>)}
                </select>
              )}
              <div className="grid grid-cols-2 gap-1.5">
                <input value={k.koefisien} onChange={(e) => updateK(k.id, { koefisien: e.target.value.replace(/[^0-9.]/g, '') })} placeholder="Koefisien" inputMode="decimal"
                  className="px-2 py-1.5 rounded att-mono text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink }} />
                <input value={k.hargaSatuan} onChange={(e) => updateK(k.id, { hargaSatuan: e.target.value.replace(/[^0-9]/g, '') })} placeholder="Harga satuan" inputMode="numeric"
                  className="px-2 py-1.5 rounded att-mono text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink }} />
              </div>
              {k.koefisien && k.hargaSatuan && (
                <p className="att-mono text-[10px] mt-1" style={{ color: THEME.amber }}>= {formatRupiah((Number(k.koefisien) || 0) * (Number(k.hargaSatuan) || 0))}</p>
              )}
            </div>
          ))}
        </div>
        <button type="button" onClick={addK}
          className="w-full flex items-center justify-center gap-1.5 py-2 rounded att-mono text-xs font-semibold mb-3"
          style={{ border: `1px dashed ${THEME.line}`, color: THEME.inkSoft }}>
          <Plus size={13} /> Tambah Komponen
        </button>

        <div className="p-2.5 rounded-lg mb-3" style={{ background: THEME.charcoal }}>
          <div className="flex items-center justify-between">
            <span className="att-mono text-[11px]" style={{ color: THEME.paper }}>Harga Dasar (tanpa profit)</span>
            <span className="att-mono text-xs" style={{ color: THEME.paper }}>{formatRupiah(base)}</span>
          </div>
          <div className="flex items-center justify-between mt-1">
            <span className="att-mono text-[11px]" style={{ color: THEME.amber }}>Harga Satuan Final (+{overheadProfit}%)</span>
            <span className="att-body font-bold text-sm" style={{ color: THEME.amber }}>{formatRupiah(final)}</span>
          </div>
        </div>

        <button type="button" onClick={handleSave}
          className="w-full flex items-center justify-center gap-1.5 py-2.5 rounded font-semibold text-sm"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Check size={15} /> Simpan Item AHSP
        </button>
      </div>
    </div>
  );
}

function AhspMasterTab({ ahspList, onAdd, onUpdate, onDelete, company, onSaveCompany }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [search, setSearch] = useState('');
  const [overheadInput, setOverheadInput] = useState(String(company.ahspOverheadProfit ?? 15));
  const overheadProfit = Number(company.ahspOverheadProfit ?? 15);
  const [upahInput, setUpahInput] = useState(() => {
    const rates = company.upahPekerja || {};
    const init = {};
    UPAH_KATEGORI.forEach((u) => { init[u.key] = String(rates[u.key] || ''); });
    return init;
  });

  const handleSaveOverhead = async () => {
    await onSaveCompany({ ...company, ahspOverheadProfit: Number(overheadInput) || 0 });
  };

  const handleSaveUpah = async () => {
    const rates = {};
    UPAH_KATEGORI.forEach((u) => { rates[u.key] = Number(upahInput[u.key]) || 0; });
    await onSaveCompany({ ...company, upahPekerja: rates });
  };

  const handleSave = async (data) => {
    if (editing) await onUpdate(editing.id, data);
    else await onAdd(data);
    setShowForm(false);
    setEditing(null);
  };

  const filtered = ahspList.filter((a) =>
    !search.trim() || a.uraian.toLowerCase().includes(search.toLowerCase()) || a.kode.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div className="p-4 pb-24">
      <p className="att-mono text-xs mb-2" style={{ color: THEME.inkSoft }}>SETTING AHSP</p>
      <div className="p-3 rounded-lg mb-3 flex items-center gap-2" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
        <div className="flex-1">
          <p className="att-body text-sm font-semibold" style={{ color: THEME.ink }}>Profit & Overhead</p>
          <p className="att-mono text-[10px]" style={{ color: THEME.inkSoft }}>Ditambahkan ke semua harga dasar AHSP saat dipakai di Penawaran</p>
        </div>
        <input value={overheadInput} onChange={(e) => setOverheadInput(e.target.value.replace(/[^0-9.]/g, ''))}
          className="w-16 px-2 py-2 rounded att-mono text-sm text-center outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        <span className="att-mono text-sm" style={{ color: THEME.ink }}>%</span>
        <button type="button" onClick={handleSaveOverhead} className="p-2 rounded" style={{ background: THEME.amber }}>
          <Check size={15} color={THEME.charcoal} />
        </button>
      </div>

      <div className="p-3 rounded-lg mb-5" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
        <p className="att-body text-sm font-semibold mb-0.5" style={{ color: THEME.ink }}>Data Upah Pekerja</p>
        <p className="att-mono text-[10px] mb-2.5" style={{ color: THEME.inkSoft }}>Diisi manual — dipakai untuk isi cepat komponen "Upah" di AHSP</p>
        <div className="grid grid-cols-2 gap-2 mb-2">
          {UPAH_KATEGORI.map((u) => (
            <div key={u.key} className="flex items-center gap-1.5">
              <span className="att-mono text-[9.5px] flex-1" style={{ color: THEME.inkSoft }}>{u.label}</span>
              <input value={upahInput[u.key]} onChange={(e) => setUpahInput({ ...upahInput, [u.key]: e.target.value.replace(/[^0-9]/g, '') })}
                placeholder="Rp" inputMode="numeric"
                className="w-20 px-2 py-1.5 rounded att-mono text-[10.5px] outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
            </div>
          ))}
        </div>
        <button type="button" onClick={handleSaveUpah}
          className="w-full flex items-center justify-center gap-1.5 py-2 rounded att-mono text-xs font-semibold" style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Check size={13} /> Simpan Data Upah
        </button>
      </div>

      <div className="flex items-center justify-between mb-2">
        <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>MASTER AHSP ({ahspList.length})</p>
        <button type="button" onClick={() => { setEditing(null); setShowForm(true); }}
          className="flex items-center gap-1.5 px-3 py-1.5 rounded att-mono text-xs font-semibold" style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Plus size={13} /> Tambah Item
        </button>
      </div>
      <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Cari kode atau uraian pekerjaan..."
        className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-3" style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }} />

      <div className="space-y-2">
        {filtered.map((a) => (
          <div key={a.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div className="flex items-center justify-between">
              <p className="att-mono text-[10px]" style={{ color: THEME.amber }}>{a.kode}</p>
              {a.sumber && <span className="att-mono text-[9px] px-1.5 py-0.5 rounded" style={{ background: THEME.concrete, color: THEME.inkSoft }}>{a.sumber}</span>}
            </div>
            <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{a.uraian}</p>
            <p className="att-mono text-xs mt-1" style={{ color: THEME.inkSoft }}>
              Satuan: {a.satuan} &middot; Harga final: <b style={{ color: THEME.ink }}>{formatRupiah(ahspFinalHarga(a.komponen, overheadProfit))}</b>
            </p>
            <div className="flex items-center gap-1.5 mt-2">
              <button type="button" onClick={() => { setEditing(a); setShowForm(true); }}
                className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.concrete }}>
                <Edit3 size={12} /> Edit
              </button>
              {confirmDelete === a.id ? (
                <button type="button" onClick={() => { onDelete(a.id); setConfirmDelete(null); }} className="p-1.5 rounded ml-auto" style={{ background: THEME.rust }}>
                  <Check size={13} color={THEME.paper} />
                </button>
              ) : (
                <button type="button" onClick={() => setConfirmDelete(a.id)} className="p-1.5 rounded ml-auto" style={{ background: THEME.concrete }}>
                  <Trash2 size={13} color={THEME.rust} />
                </button>
              )}
            </div>
          </div>
        ))}
        {filtered.length === 0 && (
          <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada item AHSP. Tap "Tambah Item" untuk mulai.</p>
        )}
      </div>

      {showForm && (
        <AhspMasterForm initial={editing} overheadProfit={overheadProfit} upahRates={company.upahPekerja || {}} onSave={handleSave} onCancel={() => { setShowForm(false); setEditing(null); }} />
      )}
    </div>
  );
}

function AhspPickerModal({ ahspList, overheadProfit, onPick, onCancel }) {
  const [search, setSearch] = useState('');
  const filtered = ahspList.filter((a) =>
    !search.trim() || a.uraian.toLowerCase().includes(search.toLowerCase()) || a.kode.toLowerCase().includes(search.toLowerCase())
  );
  return (
    <div className="fixed inset-0 z-[60] flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(13,25,48,0.75)' }}>
      <div className="w-full max-w-md rounded-xl p-4 att-body max-h-[85vh] overflow-y-auto" style={{ background: THEME.paper }}>
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>Pilih dari Master AHSP</h3>
          <button type="button" onClick={onCancel}><X size={18} color={THEME.inkSoft} /></button>
        </div>
        <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Cari kode atau uraian pekerjaan..."
          className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-3" style={{ background: THEME.concrete, color: THEME.ink }} />
        <div className="space-y-1.5">
          {filtered.map((a) => (
            <button key={a.id} type="button" onClick={() => onPick(a)}
              className="w-full text-left p-2.5 rounded-lg" style={{ background: THEME.concrete }}>
              <p className="att-mono text-[10px]" style={{ color: THEME.amber }}>{a.kode}</p>
              <p className="att-body text-sm font-semibold" style={{ color: THEME.ink }}>{a.uraian}</p>
              <p className="att-mono text-[10.5px] mt-0.5" style={{ color: THEME.inkSoft }}>
                {a.satuan} &middot; {formatRupiah(ahspFinalHarga(a.komponen, overheadProfit))}
              </p>
            </button>
          ))}
          {filtered.length === 0 && (
            <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>
              Tidak ada item AHSP yang cocok. Tambahkan dulu lewat tab "AHSP Master".
            </p>
          )}
        </div>
      </div>
    </div>
  );
}

function penawaranTotal(items) {
  return (items || []).reduce((s, it) => s + (Number(it.volume) || 0) * (Number(it.hargaSatuan) || 0), 0);
}

function penawaranBreakdown(p) {
  const subtotal = penawaranTotal(p.items);
  const ppn = p.includePPN ? subtotal * 0.11 : 0;
  return { subtotal, ppn, grandTotal: subtotal + ppn };
}

function penawaranHtml(p) {
  const items = p.items || [];
  const rows = items.map((it, i) => `
    <tr>
      <td style="text-align:center;">${i + 1}</td>
      <td>${escapeHtml(it.uraian)}</td>
      <td style="text-align:center;">${escapeHtml(it.kodeAhsp || '-')}</td>
      <td style="text-align:center;">${it.volume}</td>
      <td style="text-align:center;">${escapeHtml(it.satuan)}</td>
      <td style="text-align:right;">${formatRupiah(it.hargaSatuan)}</td>
      <td style="text-align:right;">${formatRupiah((Number(it.volume) || 0) * (Number(it.hargaSatuan) || 0))}</td>
    </tr>`).join('');
  const { subtotal, ppn, grandTotal } = penawaranBreakdown(p);
  return `
    <p><b>No</b>: ${escapeHtml(p.nomorSurat || '-')}<br/>
    <b>Tanggal</b>: ${escapeHtml(p.tanggal || '-')}<br/>
    <b>Kepada Yth.</b>: ${escapeHtml(p.namaKlien || '-')}<br/>
    <b>Perihal</b>: Penawaran Harga — ${escapeHtml(p.perihal || '-')}</p>
    <p>${escapeHtml(p.pembuka || 'Dengan hormat, bersama ini kami sampaikan penawaran harga untuk pekerjaan tersebut di atas sebagai berikut:')}</p>
    <table style="width:100%;border-collapse:collapse;font-size:0.85em;" border="1" cellpadding="5">
      <thead><tr style="background:#eee;"><th>No</th><th>Uraian Pekerjaan</th><th>Kode AHSP</th><th>Vol.</th><th>Satuan</th><th>Harga Satuan</th><th>Jumlah</th></tr></thead>
      <tbody>${rows}</tbody>
      <tfoot>
        <tr><td colspan="6" style="text-align:right;">Subtotal</td><td style="text-align:right;">${formatRupiah(subtotal)}</td></tr>
        ${p.includePPN ? `<tr><td colspan="6" style="text-align:right;">PPN 11%</td><td style="text-align:right;">${formatRupiah(ppn)}</td></tr>` : ''}
        <tr style="font-weight:bold;"><td colspan="6" style="text-align:right;">TOTAL</td><td style="text-align:right;">${formatRupiah(grandTotal)}</td></tr>
      </tfoot>
    </table>
    <p style="font-style:italic;">Terbilang: ${terbilang(grandTotal)}</p>
    ${p.syarat ? `<p><b>Syarat & Ketentuan:</b><br/>${escapeHtml(p.syarat).replace(/\n/g, '<br/>')}</p>` : ''}
    <p style="margin-top:40px;">Hormat kami,</p>
    <p style="margin-top:60px;">${escapeHtml(CURRENT_COMPANY.name || '-')}</p>
  `;
}

function exportPenawaranPDF(p) {
  try {
    const doc = new window.jspdf.jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    let y = pdfAddHeader(doc, 'Surat Penawaran', p.nomorSurat);

    doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(...PDF_INK);
    doc.text(`Tanggal: ${p.tanggal || '-'}`, 40, y);
    doc.text(`Kepada Yth.: ${p.namaKlien || '-'}`, 40, y + 14);
    doc.text(`Perihal: Penawaran Harga — ${p.perihal || '-'}`, 40, y + 28);
    y += 46;

    doc.setFontSize(9);
    const openText = p.pembuka || 'Dengan hormat, bersama ini kami sampaikan penawaran harga untuk pekerjaan tersebut di atas sebagai berikut:';
    const openLines = doc.splitTextToSize(openText, 515);
    doc.text(openLines, 40, y);
    y += openLines.length * 12 + 10;

    const items = p.items || [];
    const body = items.map((it, i) => [
      String(i + 1), it.uraian, it.kodeAhsp || '-', String(it.volume), it.satuan, formatRupiah(it.hargaSatuan),
      formatRupiah((Number(it.volume) || 0) * (Number(it.hargaSatuan) || 0)),
    ]);
    const { subtotal, ppn, grandTotal } = penawaranBreakdown(p);
    const footRows = [['', '', '', '', '', 'Subtotal', formatRupiah(subtotal)]];
    if (p.includePPN) footRows.push(['', '', '', '', '', 'PPN 11%', formatRupiah(ppn)]);
    footRows.push(['', '', '', '', '', 'TOTAL', formatRupiah(grandTotal)]);
    doc.autoTable({
      startY: y, margin: { left: 40, right: 40 },
      head: [['No', 'Uraian Pekerjaan', 'Kode AHSP', 'Vol.', 'Satuan', 'Harga Satuan', 'Jumlah']],
      body,
      foot: footRows,
      styles: { fontSize: 8, textColor: PDF_INK },
      headStyles: { fillColor: PDF_NAVY, textColor: [242, 236, 217] },
      footStyles: { fillColor: [241, 236, 221], textColor: PDF_NAVY, fontStyle: 'bold' },
      columnStyles: { 1: { cellWidth: 140 } },
      alternateRowStyles: { fillColor: [250, 248, 241] },
    });
    y = doc.lastAutoTable.finalY + 16;

    doc.setFont('helvetica', 'italic'); doc.setFontSize(9); doc.setTextColor(...PDF_SLATE);
    const terbilangLines = doc.splitTextToSize(`Terbilang: ${terbilang(grandTotal)}`, 515);
    doc.text(terbilangLines, 40, y);
    y += terbilangLines.length * 12 + 12;

    if (p.syarat) {
      if (y > 680) { doc.addPage(); y = 40; }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5); doc.setTextColor(...PDF_NAVY);
      doc.text('Syarat & Ketentuan:', 40, y);
      y += 14;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(...PDF_INK);
      const syaratLines = doc.splitTextToSize(p.syarat, 515);
      doc.text(syaratLines, 40, y);
      y += syaratLines.length * 11 + 20;
    }

    if (y > 680) { doc.addPage(); y = 40; }
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(...PDF_INK);
    doc.text('Hormat kami,', 40, y);
    doc.setFont('helvetica', 'bold');
    doc.text(CURRENT_COMPANY.name || '-', 40, y + 56);

    pdfFinishAllPages(doc);
    doc.save(`penawaran-${(p.nomorSurat || 'draft').replace(/[\/\s]+/g, '-').toLowerCase()}.pdf`);
  } catch (err) {
    console.error('Gagal ekspor PDF Penawaran:', err);
    alert('Gagal membuat file PDF. Coba lagi.');
  }
}

function ahspItemsUsedIn(p, ahspList) {
  const kodeSet = new Set((p.items || []).map((it) => it.kodeAhsp).filter(Boolean));
  return ahspList.filter((a) => kodeSet.has(a.kode));
}

function exportAhspAnalysisPDF(p, ahspList, overheadProfit) {
  try {
    const usedAhsp = ahspItemsUsedIn(p, ahspList);
    if (usedAhsp.length === 0) {
      alert('Tidak ada item di penawaran ini yang memakai kode AHSP dari Master. Tidak ada yang bisa dicetak.');
      return;
    }
    const doc = new window.jspdf.jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    let y = pdfAddHeader(doc, 'Lampiran Analisa Harga Satuan Pekerjaan', p.nomorSurat || p.perihal);

    usedAhsp.forEach((a, idx) => {
      if (idx > 0 && y > 620) { doc.addPage(); y = 40; }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(...PDF_NAVY);
      doc.text(`${a.kode} — ${a.uraian}`, 40, y);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...PDF_SLATE);
      doc.text(`Satuan: ${a.satuan}  |  Sumber: ${a.sumber || '-'}`, 40, y + 12);
      y += 24;

      const body = (a.komponen || []).map((k) => [
        AHSP_JENIS_LABEL[k.jenis] || k.jenis, k.nama, String(k.koefisien), formatRupiah(k.hargaSatuan),
        formatRupiah((Number(k.koefisien) || 0) * (Number(k.hargaSatuan) || 0)),
      ]);
      const base = ahspBaseHarga(a.komponen);
      const final = ahspFinalHarga(a.komponen, overheadProfit);
      doc.autoTable({
        startY: y, margin: { left: 40, right: 40 },
        head: [['Jenis', 'Uraian', 'Koefisien', 'Harga Satuan', 'Jumlah']],
        body,
        foot: [
          ['', '', '', 'Harga Dasar', formatRupiah(base)],
          ['', '', '', `Overhead & Profit (${overheadProfit}%)`, formatRupiah(final - base)],
          ['', '', '', 'HARGA SATUAN AKHIR', formatRupiah(final)],
        ],
        styles: { fontSize: 8, textColor: PDF_INK },
        headStyles: { fillColor: PDF_NAVY, textColor: [242, 236, 217] },
        footStyles: { fillColor: [241, 236, 221], textColor: PDF_NAVY, fontStyle: 'bold', fontSize: 7.5 },
        alternateRowStyles: { fillColor: [250, 248, 241] },
      });
      y = doc.lastAutoTable.finalY + 26;
    });

    pdfFinishAllPages(doc);
    doc.save(`analisa-ahsp-${(p.nomorSurat || 'draft').replace(/[\/\s]+/g, '-').toLowerCase()}.pdf`);
  } catch (err) {
    console.error('Gagal ekspor PDF Analisa AHSP:', err);
    alert('Gagal membuat file PDF. Coba lagi.');
  }
}

function ahspAnalysisHtml(p, ahspList, overheadProfit) {
  const usedAhsp = ahspItemsUsedIn(p, ahspList);
  const sections = usedAhsp.map((a) => {
    const rows = (a.komponen || []).map((k) => `
      <tr>
        <td>${escapeHtml(AHSP_JENIS_LABEL[k.jenis] || k.jenis)}</td>
        <td>${escapeHtml(k.nama)}</td>
        <td style="text-align:center;">${k.koefisien}</td>
        <td style="text-align:right;">${formatRupiah(k.hargaSatuan)}</td>
        <td style="text-align:right;">${formatRupiah((Number(k.koefisien) || 0) * (Number(k.hargaSatuan) || 0))}</td>
      </tr>`).join('');
    const base = ahspBaseHarga(a.komponen);
    const final = ahspFinalHarga(a.komponen, overheadProfit);
    return `
      <p style="margin-bottom:2px;"><b>${escapeHtml(a.kode)} — ${escapeHtml(a.uraian)}</b></p>
      <p style="font-size:0.8em;margin-top:0;">Satuan: ${escapeHtml(a.satuan)} | Sumber: ${escapeHtml(a.sumber || '-')}</p>
      <table style="width:100%;border-collapse:collapse;font-size:0.82em;margin-bottom:18px;" border="1" cellpadding="5">
        <thead><tr style="background:#eee;"><th>Jenis</th><th>Uraian</th><th>Koefisien</th><th>Harga Satuan</th><th>Jumlah</th></tr></thead>
        <tbody>${rows}</tbody>
        <tfoot>
          <tr><td colspan="4" style="text-align:right;">Harga Dasar</td><td style="text-align:right;">${formatRupiah(base)}</td></tr>
          <tr><td colspan="4" style="text-align:right;">Overhead & Profit (${overheadProfit}%)</td><td style="text-align:right;">${formatRupiah(final - base)}</td></tr>
          <tr style="font-weight:bold;"><td colspan="4" style="text-align:right;">HARGA SATUAN AKHIR</td><td style="text-align:right;">${formatRupiah(final)}</td></tr>
        </tfoot>
      </table>`;
  }).join('');
  return `<p>Lampiran Analisa Harga Satuan Pekerjaan untuk: <b>${escapeHtml(p.perihal || '-')}</b></p>${sections || '<p>Tidak ada item yang memakai kode AHSP dari Master.</p>'}`;
}

function emptyPenawaranItem() {
  return { id: uid(), uraian: '', kodeAhsp: '', volume: '', satuan: '', hargaSatuan: '' };
}

function PenawaranForm({ onSave, onCancel, initial, ahspList, overheadProfit, projects }) {
  const [nomorSurat, setNomorSurat] = useState(initial?.nomorSurat || '');
  const [tanggal, setTanggal] = useState(initial?.tanggal || new Date().toISOString().slice(0, 10));
  const [namaKlien, setNamaKlien] = useState(initial?.namaKlien || '');
  const [projectId, setProjectId] = useState(initial?.projectId || '');
  const [perihal, setPerihal] = useState(initial?.perihal || '');
  const [pembuka, setPembuka] = useState(initial?.pembuka || '');
  const [items, setItems] = useState(initial?.items?.length ? initial.items : [emptyPenawaranItem()]);
  const [syarat, setSyarat] = useState(initial?.syarat || 'Harga sudah termasuk material dan upah kerja.\nPembayaran: 50% DP, 50% saat serah terima.\nPenawaran berlaku 14 hari sejak tanggal surat.');
  const [includePPN, setIncludePPN] = useState(initial?.includePPN || false);
  const [pickerFor, setPickerFor] = useState(null); // id item yang lagi pilih AHSP

  const updateItem = (id, patch) => setItems(items.map((it) => (it.id === id ? { ...it, ...patch } : it)));
  const addItem = () => setItems([...items, emptyPenawaranItem()]);
  const removeItem = (id) => setItems(items.filter((it) => it.id !== id));
  const { subtotal, ppn, grandTotal } = penawaranBreakdown({ items, includePPN });

  const handlePickAhsp = (ahspItem) => {
    updateItem(pickerFor, {
      uraian: ahspItem.uraian, kodeAhsp: ahspItem.kode, satuan: ahspItem.satuan,
      hargaSatuan: Math.round(ahspFinalHarga(ahspItem.komponen, overheadProfit)),
    });
    setPickerFor(null);
  };

  const handleSave = async () => {
    if (!namaKlien.trim() || items.every((it) => !it.uraian.trim())) {
      alert('Isi minimal nama klien dan 1 uraian pekerjaan.');
      return;
    }
    await onSave({
      id: initial?.id || uid(), nomorSurat: nomorSurat.trim(), tanggal, namaKlien: namaKlien.trim(), projectId,
      perihal: perihal.trim(), pembuka: pembuka.trim(), syarat: syarat.trim(), includePPN,
      items: items.filter((it) => it.uraian.trim()).map((it) => ({ ...it, volume: Number(it.volume) || 0, hargaSatuan: Number(it.hargaSatuan) || 0 })),
    });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(13,25,48,0.7)' }}>
      <div className="w-full max-w-md rounded-xl p-4 att-body max-h-[92vh] overflow-y-auto" style={{ background: THEME.paper }}>
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>{initial ? 'Edit' : 'Buat'} Surat Penawaran</h3>
          <button type="button" onClick={onCancel}><X size={18} color={THEME.inkSoft} /></button>
        </div>
        <div className="grid grid-cols-2 gap-2 mb-2">
          <input value={nomorSurat} onChange={(e) => setNomorSurat(e.target.value)} placeholder="No. Surat (opsional)"
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input type="date" value={tanggal} onChange={(e) => setTanggal(e.target.value)}
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>
        <input value={namaKlien} onChange={(e) => setNamaKlien(e.target.value)} placeholder="Kepada Yth. (nama klien/perusahaan)"
          className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-2" style={{ background: THEME.concrete, color: THEME.ink }} />
        <select value={projectId} onChange={(e) => setProjectId(e.target.value)}
          className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-2" style={{ background: THEME.concrete, color: THEME.ink }}>
          <option value="">— Hubungkan ke Proyek (opsional, untuk Progres Proyek) —</option>
          {projects.map((pr) => <option key={pr.id} value={pr.id}>{pr.name}</option>)}
        </select>
        <input value={perihal} onChange={(e) => setPerihal(e.target.value)} placeholder="Perihal (nama pekerjaan/proyek)"
          className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-2" style={{ background: THEME.concrete, color: THEME.ink }} />
        <textarea value={pembuka} onChange={(e) => setPembuka(e.target.value)} rows={2} placeholder="Kalimat pembuka (opsional, ada default kalau kosong)"
          className="w-full px-3 py-2 rounded att-body text-xs outline-none mb-3" style={{ background: THEME.concrete, color: THEME.ink }} />

        <p className="att-mono text-[10px] mb-2" style={{ color: THEME.inkSoft }}>DAFTAR PEKERJAAN</p>
        <div className="space-y-2 mb-2">
          {items.map((it) => (
            <div key={it.id} className="p-2.5 rounded-lg" style={{ background: THEME.concrete }}>
              <div className="flex gap-1.5 mb-1.5">
                <input value={it.uraian} onChange={(e) => updateItem(it.id, { uraian: e.target.value })} placeholder="Uraian pekerjaan"
                  className="flex-1 px-2 py-1.5 rounded att-body text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink }} />
                <button type="button" onClick={() => setPickerFor(it.id)} title="Pilih dari AHSP Master" className="p-1.5 rounded shrink-0" style={{ background: THEME.amberSoft }}>
                  <Receipt size={13} color={THEME.charcoal} />
                </button>
                <button type="button" onClick={() => removeItem(it.id)} className="p-1.5 rounded shrink-0" style={{ background: THEME.paper }}>
                  <Trash2 size={13} color={THEME.rust} />
                </button>
              </div>
              {it.kodeAhsp && (
                <p className="att-mono text-[10px] mb-1.5" style={{ color: THEME.amber }}>Kode AHSP: {it.kodeAhsp}</p>
              )}
              <DimensiCalculator onApply={(hasil) => updateItem(it.id, { volume: String(hasil) })} />
              <div className="grid grid-cols-3 gap-1.5">
                <input value={it.volume} onChange={(e) => updateItem(it.id, { volume: e.target.value.replace(/[^0-9.]/g, '') })} placeholder="Vol." inputMode="decimal"
                  className="px-2 py-1.5 rounded att-mono text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink }} />
                <input value={it.satuan} onChange={(e) => updateItem(it.id, { satuan: e.target.value })} placeholder="Satuan"
                  className="px-2 py-1.5 rounded att-mono text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink }} />
                <input value={it.hargaSatuan} onChange={(e) => updateItem(it.id, { hargaSatuan: e.target.value.replace(/[^0-9]/g, ''), kodeAhsp: '' })} placeholder="Harga satuan" inputMode="numeric"
                  className="px-2 py-1.5 rounded att-mono text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink }} />
              </div>
              {it.volume && it.hargaSatuan && (
                <div className="flex items-center justify-between mt-1">
                  <p className="att-mono text-[10px]" style={{ color: THEME.amber }}>= {formatRupiah((Number(it.volume) || 0) * (Number(it.hargaSatuan) || 0))}</p>
                  <p className="att-mono text-[10px]" style={{ color: THEME.inkSoft }}>
                    Bobot: {subtotal > 0 ? (((Number(it.volume) || 0) * (Number(it.hargaSatuan) || 0) / subtotal) * 100).toFixed(1) : '0.0'}%
                  </p>
                </div>
              )}
            </div>
          ))}
        </div>
        <button type="button" onClick={addItem}
          className="w-full flex items-center justify-center gap-1.5 py-2 rounded att-mono text-xs font-semibold mb-3"
          style={{ border: `1px dashed ${THEME.line}`, color: THEME.inkSoft }}>
          <Plus size={13} /> Tambah Baris Pekerjaan
        </button>

        <label className="flex items-center gap-2 mb-2 px-1">
          <input type="checkbox" checked={includePPN} onChange={(e) => setIncludePPN(e.target.checked)} className="w-4 h-4" />
          <span className="att-body text-sm" style={{ color: THEME.ink }}>Tambahkan PPN 11%</span>
        </label>

        <div className="p-2.5 rounded-lg mb-3" style={{ background: THEME.charcoal }}>
          <div className="flex items-center justify-between">
            <span className="att-mono text-[11px]" style={{ color: THEME.paper }}>Subtotal</span>
            <span className="att-mono text-xs" style={{ color: THEME.paper }}>{formatRupiah(subtotal)}</span>
          </div>
          {includePPN && (
            <div className="flex items-center justify-between mt-1">
              <span className="att-mono text-[11px]" style={{ color: THEME.paper }}>PPN 11%</span>
              <span className="att-mono text-xs" style={{ color: THEME.paper }}>{formatRupiah(ppn)}</span>
            </div>
          )}
          <div className="flex items-center justify-between mt-1.5 pt-1.5" style={{ borderTop: `1px solid rgba(233,222,190,0.2)` }}>
            <span className="att-mono text-xs" style={{ color: THEME.amber }}>TOTAL</span>
            <span className="att-body font-bold text-sm" style={{ color: THEME.amber }}>{formatRupiah(grandTotal)}</span>
          </div>
        </div>

        <p className="att-mono text-[10px] mb-1.5" style={{ color: THEME.inkSoft }}>SYARAT & KETENTUAN</p>
        <textarea value={syarat} onChange={(e) => setSyarat(e.target.value)} rows={3}
          className="w-full px-3 py-2 rounded att-body text-xs outline-none mb-3" style={{ background: THEME.concrete, color: THEME.ink }} />

        <button type="button" onClick={handleSave}
          className="w-full flex items-center justify-center gap-1.5 py-2.5 rounded font-semibold text-sm"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Check size={15} /> Simpan Penawaran
        </button>
      </div>

      {pickerFor && (
        <AhspPickerModal ahspList={ahspList} overheadProfit={overheadProfit} onPick={handlePickAhsp} onCancel={() => setPickerFor(null)} />
      )}
    </div>
  );
}

const CLIENT_JENIS_LABEL = { perorangan: 'Perorangan', perusahaan: 'Perusahaan', instansi: 'Instansi Pemerintah' };
const CLIENT_STATUS_LABEL = { prospek: 'Prospek', aktif: 'Aktif', selesai: 'Selesai' };
const CLIENT_STATUS_COLORS = { prospek: '#C9A227', aktif: '#2F7A52', selesai: '#5B6478' };
const CLIENT_SUMBER_LIST = ['Website', 'Referensi', 'Datang Langsung', 'Media Sosial', 'Lainnya'];

function emptyClientForm() {
  return {
    nama: '', jenis: 'perusahaan', pic: '', jabatanPic: '', hp: '', email: '', alamat: '', npwp: '',
    status: 'prospek', sumber: CLIENT_SUMBER_LIST[0], projectId: '', catatan: '',
  };
}

function exportClientPDF(c, projects) {
  try {
    const doc = new window.jspdf.jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    let y = pdfAddHeader(doc, 'Profil Klien', c.nama);

    const project = projects.find((p) => p.id === c.projectId);
    const rows = [
      ['Nama', c.nama || '-'], ['Jenis', CLIENT_JENIS_LABEL[c.jenis] || '-'], ['Status', CLIENT_STATUS_LABEL[c.status] || '-'],
      ['PIC', c.pic || '-'], ['Jabatan PIC', c.jabatanPic || '-'], ['No. HP', c.hp || '-'], ['Email', c.email || '-'],
      ['Alamat', c.alamat || '-'], ['NPWP', c.npwp || '-'], ['Sumber', c.sumber || '-'], ['Proyek Terkait', project?.name || '-'],
    ];
    doc.autoTable({
      startY: y, margin: { left: 40, right: 40 }, theme: 'plain',
      body: rows, styles: { fontSize: 9.5, textColor: PDF_INK, cellPadding: { top: 5, bottom: 5, left: 0, right: 0 } },
      columnStyles: { 0: { textColor: PDF_SLATE, cellWidth: 110 }, 1: { fontStyle: 'bold', textColor: PDF_NAVY } },
    });
    y = doc.lastAutoTable.finalY + 16;

    if (c.catatan) {
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5); doc.setTextColor(...PDF_NAVY);
      doc.text('Catatan:', 40, y);
      y += 14;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(...PDF_INK);
      const catatanLines = doc.splitTextToSize(c.catatan, 515);
      doc.text(catatanLines, 40, y);
    }

    pdfFinishAllPages(doc);
    doc.save(`klien-${(c.nama || 'profil').replace(/\s+/g, '-').toLowerCase()}.pdf`);
  } catch (err) {
    console.error('Gagal ekspor PDF Klien:', err);
    alert('Gagal membuat file PDF. Coba lagi.');
  }
}

function ClientForm({ onSave, onCancel, initial, projects }) {
  const [form, setForm] = useState(initial || emptyClientForm());

  const handleSave = async () => {
    if (!form.nama.trim()) {
      alert('Isi nama klien/perusahaan dulu.');
      return;
    }
    await onSave({ ...form, id: initial?.id || uid(), nama: form.nama.trim() });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(13,25,48,0.7)' }}>
      <div className="w-full max-w-md rounded-xl p-4 att-body max-h-[92vh] overflow-y-auto" style={{ background: THEME.paper }}>
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>{initial ? 'Edit' : 'Tambah'} Klien</h3>
          <button type="button" onClick={onCancel}><X size={18} color={THEME.inkSoft} /></button>
        </div>
        <input value={form.nama} onChange={(e) => setForm({ ...form, nama: e.target.value })} placeholder="Nama klien / perusahaan"
          className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-2" style={{ background: THEME.concrete, color: THEME.ink }} />
        <div className="grid grid-cols-2 gap-2 mb-2">
          <select value={form.jenis} onChange={(e) => setForm({ ...form, jenis: e.target.value })}
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }}>
            {Object.entries(CLIENT_JENIS_LABEL).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select>
          <select value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })}
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }}>
            {Object.entries(CLIENT_STATUS_LABEL).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select>
        </div>
        <div className="grid grid-cols-2 gap-2 mb-2">
          <input value={form.pic} onChange={(e) => setForm({ ...form, pic: e.target.value })} placeholder="Nama PIC"
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={form.jabatanPic} onChange={(e) => setForm({ ...form, jabatanPic: e.target.value })} placeholder="Jabatan PIC"
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>
        <div className="grid grid-cols-2 gap-2 mb-2">
          <input value={form.hp} onChange={(e) => setForm({ ...form, hp: e.target.value })} placeholder="No. HP"
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} placeholder="Email"
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>
        <input value={form.alamat} onChange={(e) => setForm({ ...form, alamat: e.target.value })} placeholder="Alamat"
          className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-2" style={{ background: THEME.concrete, color: THEME.ink }} />
        <div className="grid grid-cols-2 gap-2 mb-2">
          <input value={form.npwp} onChange={(e) => setForm({ ...form, npwp: e.target.value })} placeholder="NPWP (opsional)"
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <select value={form.sumber} onChange={(e) => setForm({ ...form, sumber: e.target.value })}
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }}>
            {CLIENT_SUMBER_LIST.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
        <select value={form.projectId} onChange={(e) => setForm({ ...form, projectId: e.target.value })}
          className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-2" style={{ background: THEME.concrete, color: THEME.ink }}>
          <option value="">— Hubungkan ke Proyek (opsional) —</option>
          {projects.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
        <textarea value={form.catatan} onChange={(e) => setForm({ ...form, catatan: e.target.value })} rows={3} placeholder="Catatan / riwayat komunikasi (opsional)"
          className="w-full px-3 py-2 rounded att-body text-xs outline-none mb-3" style={{ background: THEME.concrete, color: THEME.ink }} />

        <button type="button" onClick={handleSave}
          className="w-full flex items-center justify-center gap-1.5 py-2.5 rounded font-semibold text-sm"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Check size={15} /> Simpan Klien
        </button>
      </div>
    </div>
  );
}

function ClientTab({ clients, onAdd, onUpdate, onDelete, projects }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [search, setSearch] = useState('');
  const [filterStatus, setFilterStatus] = useState('all');

  const filtered = clients.filter((c) => {
    const matchSearch = !search.trim() || c.nama.toLowerCase().includes(search.toLowerCase()) || (c.pic || '').toLowerCase().includes(search.toLowerCase());
    const matchStatus = filterStatus === 'all' || c.status === filterStatus;
    return matchSearch && matchStatus;
  });

  const handleSave = async (data) => {
    if (editing) await onUpdate(editing.id, data);
    else await onAdd(data);
    setShowForm(false);
    setEditing(null);
  };

  return (
    <div className="p-4 pb-24">
      <div className="flex items-center justify-between mb-3">
        <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>DATA KLIEN ({clients.length})</p>
        <button type="button" onClick={() => { setEditing(null); setShowForm(true); }}
          className="flex items-center gap-1.5 px-3 py-1.5 rounded att-mono text-xs font-semibold" style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Plus size={13} /> Tambah Klien
        </button>
      </div>

      <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Cari nama klien atau PIC..."
        className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-2" style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }} />
      <div className="flex gap-1.5 mb-3 overflow-x-auto">
        {['all', 'prospek', 'aktif', 'selesai'].map((s) => (
          <button key={s} type="button" onClick={() => setFilterStatus(s)}
            className="px-3 py-1.5 rounded att-mono text-[11px] font-semibold shrink-0"
            style={{ background: filterStatus === s ? THEME.amber : THEME.concrete, color: filterStatus === s ? THEME.charcoal : THEME.inkSoft }}>
            {s === 'all' ? 'Semua' : CLIENT_STATUS_LABEL[s]}
          </button>
        ))}
      </div>

      <div className="space-y-2">
        {filtered.map((c) => (
          <div key={c.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div className="flex items-start justify-between mb-1">
              <div className="min-w-0">
                <p className="att-body font-semibold text-sm truncate" style={{ color: THEME.ink }}>{c.nama}</p>
                <p className="att-mono text-[11px]" style={{ color: THEME.inkSoft }}>{CLIENT_JENIS_LABEL[c.jenis]} &middot; {c.pic || '-'}</p>
              </div>
              <span className="att-mono text-[9px] px-1.5 py-0.5 rounded-full font-semibold shrink-0" style={{ background: CLIENT_STATUS_COLORS[c.status], color: '#fff' }}>
                {CLIENT_STATUS_LABEL[c.status]}
              </span>
            </div>
            {(c.hp || c.email) && (
              <p className="att-mono text-[11px] mb-2" style={{ color: THEME.inkSoft }}>{[c.hp, c.email].filter(Boolean).join(' · ')}</p>
            )}
            <div className="flex flex-wrap items-center gap-1.5">
              <button type="button" onClick={() => { setEditing(c); setShowForm(true); }}
                className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.concrete }}>
                <Edit3 size={12} /> Edit
              </button>
              <button type="button" onClick={() => exportClientPDF(c, projects)}
                className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.concrete }}>
                <Save size={12} /> PDF
              </button>
              {confirmDelete === c.id ? (
                <button type="button" onClick={() => { onDelete(c.id); setConfirmDelete(null); }} className="p-1.5 rounded ml-auto" style={{ background: THEME.rust }}>
                  <Check size={13} color={THEME.paper} />
                </button>
              ) : (
                <button type="button" onClick={() => setConfirmDelete(c.id)} className="p-1.5 rounded ml-auto" style={{ background: THEME.concrete }}>
                  <Trash2 size={13} color={THEME.rust} />
                </button>
              )}
            </div>
          </div>
        ))}
        {filtered.length === 0 && (
          <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada klien. Tap "Tambah Klien" untuk mulai.</p>
        )}
      </div>

      {showForm && (
        <ClientForm initial={editing} projects={projects} onSave={handleSave} onCancel={() => { setShowForm(false); setEditing(null); }} />
      )}
    </div>
  );
}

function PenawaranTab({ penawaranList, onAdd, onUpdate, onDelete, ahspList, overheadProfit, projects }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);

  const sorted = [...penawaranList].sort((a, b) => (a.tanggal < b.tanggal ? 1 : -1));

  const handleSave = async (data) => {
    if (editing) await onUpdate(editing.id, data);
    else await onAdd(data);
    setShowForm(false);
    setEditing(null);
  };

  return (
    <div className="p-4 pb-24">
      <div className="flex items-center justify-between mb-3">
        <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>SISTEM PENAWARAN / RAB ({sorted.length})</p>
        <button type="button" onClick={() => { setEditing(null); setShowForm(true); }}
          className="flex items-center gap-1.5 px-3 py-1.5 rounded att-mono text-xs font-semibold" style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Plus size={13} /> Buat Baru
        </button>
      </div>

      <div className="space-y-2">
        {sorted.map((p) => (
          <div key={p.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div className="flex items-start justify-between mb-1">
              <div className="min-w-0">
                <p className="att-body font-semibold text-sm truncate" style={{ color: THEME.ink }}>{p.perihal || '(tanpa perihal)'}</p>
                <p className="att-mono text-[11px]" style={{ color: THEME.inkSoft }}>Kepada: {p.namaKlien} &middot; {p.tanggal}</p>
              </div>
            </div>
            <p className="att-body font-bold text-sm mb-2" style={{ color: THEME.amber }}>{formatRupiah(penawaranBreakdown(p).grandTotal)}{p.includePPN ? ' (+PPN)' : ''}</p>
            <div className="flex flex-wrap items-center gap-1.5">
              <button type="button" onClick={() => { setEditing(p); setShowForm(true); }}
                className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.concrete }}>
                <Edit3 size={12} /> Edit
              </button>
              <button type="button" onClick={() => exportPenawaranPDF(p)}
                className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.concrete }}>
                <Save size={12} /> PDF
              </button>
              <button type="button" onClick={() => openPrintDocument(`Penawaran - ${p.perihal}`, penawaranHtml(p), false)}
                className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.concrete }}>
                <Printer size={12} /> Cetak
              </button>
              <button type="button" onClick={() => exportAhspAnalysisPDF(p, ahspList, overheadProfit)}
                className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.amber}`, color: THEME.charcoal, background: THEME.amberSoft }}>
                <FileBarChart size={12} /> PDF Analisa AHSP
              </button>
              <button type="button" onClick={() => openPrintDocument(`Analisa AHSP - ${p.perihal}`, ahspAnalysisHtml(p, ahspList, overheadProfit), false)}
                className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.amber}`, color: THEME.charcoal, background: THEME.amberSoft }}>
                <Printer size={12} /> Cetak Analisa AHSP
              </button>
              {confirmDelete === p.id ? (
                <button type="button" onClick={() => { onDelete(p.id); setConfirmDelete(null); }} className="p-1.5 rounded ml-auto" style={{ background: THEME.rust }}>
                  <Check size={13} color={THEME.paper} />
                </button>
              ) : (
                <button type="button" onClick={() => setConfirmDelete(p.id)} className="p-1.5 rounded ml-auto" style={{ background: THEME.concrete }}>
                  <Trash2 size={13} color={THEME.rust} />
                </button>
              )}
            </div>
          </div>
        ))}
        {sorted.length === 0 && (
          <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada penawaran. Tap "Buat Baru" untuk mulai.</p>
        )}
      </div>

      {showForm && (
        <PenawaranForm initial={editing} ahspList={ahspList} overheadProfit={overheadProfit} projects={projects} onSave={handleSave} onCancel={() => { setShowForm(false); setEditing(null); }} />
      )}
    </div>
  );
}

function filterEntriesByPeriode(entries, periode) {
  const now = new Date();
  const todayStr = now.toISOString().slice(0, 10);
  if (periode === 'harian') {
    return entries.filter((e) => e.date === todayStr);
  }
  if (periode === 'mingguan') {
    const weekAgo = new Date(now); weekAgo.setDate(weekAgo.getDate() - 7);
    const weekAgoStr = weekAgo.toISOString().slice(0, 10);
    return entries.filter((e) => e.date >= weekAgoStr && e.date <= todayStr);
  }
  // bulanan
  const monthPrefix = todayStr.slice(0, 7); // YYYY-MM
  return entries.filter((e) => e.date.slice(0, 7) === monthPrefix);
}

function gambarRekapBobot(doc, yStart, rabItems, allEntries, rabTotalNilai, projectId, sampaiTanggal) {
  let bobotTotal = 0;
  const rows = rabItems.map((item) => {
    const matching = allEntries.filter((e) => e.projectId === projectId && e.uraian === item.uraian && e.date <= sampaiTanggal);
    const latest = [...matching].sort((a, b) => (a.date < b.date ? 1 : -1))[0];
    const volSelesai = latest ? Number(latest.volumeSelesai) || 0 : 0;
    const volTotal = Number(item.volume) || 0;
    const nilaiItem = volTotal * (Number(item.hargaSatuan) || 0);
    const bobot = rabTotalNilai > 0 ? (nilaiItem / rabTotalNilai) * 100 : 0;
    const rasio = volTotal > 0 ? Math.min(1, volSelesai / volTotal) : 0;
    const bobotTercapai = bobot * rasio;
    bobotTotal += bobotTercapai;
    const nilaiTercapai = nilaiItem * rasio;
    const status = volTotal > 0 && volSelesai >= volTotal ? 'SELESAI' : (volSelesai > 0 ? 'ON PROSES' : 'OFF');
    return [item.uraian, status, `${volSelesai}/${volTotal} ${item.satuan}`, `${bobot.toFixed(2)}%`, `${bobotTercapai.toFixed(2)}% (${formatRupiah(nilaiTercapai)})`];
  });

  doc.setFont('helvetica', 'bold'); doc.setFontSize(10.5); doc.setTextColor(...PDF_NAVY);
  doc.text(`Rekap Bobot Pekerjaan (per tanggal ${sampaiTanggal})`, 40, yStart);
  let y = yStart + 8;
  doc.autoTable({
    startY: y, margin: { left: 40, right: 40 },
    head: [['Uraian Pekerjaan', 'Status', 'Volume', 'Bobot Rencana', 'Bobot Tercapai (+Nilai)']],
    body: rows,
    foot: [['', '', '', 'TOTAL BOBOT TERCAPAI', `${bobotTotal.toFixed(2)}% (${formatRupiah(rabTotalNilai * bobotTotal / 100)})`]],
    styles: { fontSize: 6.5, textColor: PDF_INK },
    headStyles: { fillColor: PDF_NAVY, textColor: [242, 236, 217] },
    footStyles: { fillColor: [241, 236, 221], textColor: PDF_NAVY, fontStyle: 'bold' },
    alternateRowStyles: { fillColor: [250, 248, 241] },
  });
  return { y: doc.lastAutoTable.finalY + 20, bobotTotal };
}

function bobotProyekPadaTanggal(rabItems, allEntries, rabTotalNilai, projectId, tanggal) {
  // Bobot capaian proyek keseluruhan (kumulatif) per tanggal tertentu (inklusif)
  let total = 0;
  rabItems.forEach((item) => {
    const matching = allEntries.filter((e) => e.projectId === projectId && e.uraian === item.uraian && e.date <= tanggal);
    const latest = [...matching].sort((a, b) => (a.date < b.date ? 1 : -1))[0];
    const volSelesai = latest ? Number(latest.volumeSelesai) || 0 : 0;
    const volTotal = Number(item.volume) || 0;
    if (volTotal <= 0 || rabTotalNilai <= 0) return;
    const nilaiItem = volTotal * (Number(item.hargaSatuan) || 0);
    const bobotItem = (nilaiItem / rabTotalNilai) * 100;
    const rasio = Math.min(1, volSelesai / volTotal);
    total += bobotItem * rasio;
  });
  return total;
}

function addDaysStr(dateStr, n) {
  const d = new Date(dateStr);
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
}

function buildPeriodeBuckets(allEntries, projectId, hariPerBucket) {
  const projEntries = allEntries.filter((e) => e.projectId === projectId);
  if (projEntries.length === 0) return [];
  const startStr = [...projEntries.map((e) => e.date)].sort()[0];
  const todayStr = new Date().toISOString().slice(0, 10);
  const totalDays = Math.floor((new Date(todayStr) - new Date(startStr)) / 86400000) + 1;
  const totalBuckets = Math.max(1, Math.ceil(totalDays / hariPerBucket));
  const buckets = [];
  for (let b = 1; b <= totalBuckets; b += 1) {
    const bStart = addDaysStr(startStr, (b - 1) * hariPerBucket);
    let bEnd = addDaysStr(startStr, b * hariPerBucket - 1);
    if (bEnd > todayStr) bEnd = todayStr;
    buckets.push({ index: b, startStr: bStart, endStr: bEnd });
  }
  return buckets;
}

function hitungRekapBobot(rabItems, allEntries, rabTotalNilai, projectId, sampaiTanggal) {
  let bobotTotal = 0;
  let nilaiTotal = 0;
  const rows = rabItems.map((item) => {
    const matching = allEntries.filter((e) => e.projectId === projectId && e.uraian === item.uraian && e.date <= sampaiTanggal);
    const latest = [...matching].sort((a, b) => (a.date < b.date ? 1 : -1))[0];
    const volSelesai = latest ? Number(latest.volumeSelesai) || 0 : 0;
    const volTotal = Number(item.volume) || 0;
    const nilaiItem = volTotal * (Number(item.hargaSatuan) || 0);
    const bobot = rabTotalNilai > 0 ? (nilaiItem / rabTotalNilai) * 100 : 0;
    const rasio = volTotal > 0 ? Math.min(1, volSelesai / volTotal) : 0;
    const bobotTercapai = bobot * rasio;
    const nilaiTercapai = nilaiItem * rasio;
    bobotTotal += bobotTercapai;
    nilaiTotal += nilaiTercapai;
    const status = volTotal > 0 && volSelesai >= volTotal ? 'SELESAI' : (volSelesai > 0 ? 'ON PROSES' : 'OFF');
    return { uraian: item.uraian, status, volSelesai, volTotal, satuan: item.satuan, bobot, bobotTercapai, nilaiTercapai };
  });
  return { rows, bobotTotal, nilaiTotal };
}

function bucketLabelDanEntri(project, allEntries, rabItems, rabTotalNilai, periode) {
  const projectId = project?.id;
  if (periode === 'harian') {
    const todayStr = new Date().toISOString().slice(0, 10);
    const entri = allEntries.filter((e) => e.projectId === projectId && e.date === todayStr);
    return [{ label: `Tanggal ${todayStr}`, endStr: todayStr, entri }];
  }
  const hariPerBucket = periode === 'mingguan' ? 7 : 28;
  const buckets = buildPeriodeBuckets(allEntries, projectId, hariPerBucket);
  const judul = periode === 'mingguan' ? 'Minggu ke-' : 'Bulan ke-';
  return buckets.map((b) => ({
    label: `${judul}${b.index} (${b.startStr} s/d ${b.endStr})`,
    endStr: b.endStr,
    entri: allEntries.filter((e) => e.projectId === projectId && e.date >= b.startStr && e.date <= b.endStr).sort((a, c) => (a.date < c.date ? -1 : 1)),
  }));
}

function exportLaporanPeriodeExcel(project, allEntries, rabItems, rabTotalNilai, periode) {
  try {
    const wb = XLSX.utils.book_new();
    const bucketList = bucketLabelDanEntri(project, allEntries, rabItems, rabTotalNilai, periode);

    bucketList.forEach((bucket, idx) => {
      const rekap = hitungRekapBobot(rabItems, allEntries, rabTotalNilai, project?.id, bucket.endStr);
      const rekapRows = rekap.rows.map((r) => ({
        'Uraian Pekerjaan': r.uraian, Status: r.status, Volume: `${r.volSelesai}/${r.volTotal} ${r.satuan}`,
        'Bobot Rencana (%)': Number(r.bobot.toFixed(2)), 'Bobot Tercapai (%)': Number(r.bobotTercapai.toFixed(2)), 'Nilai Tercapai (Rp)': Math.round(r.nilaiTercapai),
      }));
      rekapRows.push({ 'Uraian Pekerjaan': 'TOTAL', Status: '', Volume: '', 'Bobot Rencana (%)': '', 'Bobot Tercapai (%)': Number(rekap.bobotTotal.toFixed(2)), 'Nilai Tercapai (Rp)': Math.round(rekap.nilaiTotal) });

      const aktivitasRows = bucket.entri.map((e) => {
        const bh = bobotHarianEntry(e, allEntries, rabItems, rabTotalNilai);
        const rabItem = rabItems.find((it) => it.uraian === e.uraian);
        const hargaSatuan = Number(rabItem?.hargaSatuan) || 0;
        return {
          Tanggal: e.date, Uraian: e.uraian || '-', 'Volume Kumulatif': e.satuan ? `${e.volumeSelesai}/${e.volumeTotal} ${e.satuan}` : '-',
          'Volume Hari Itu': bh ? bh.deltaVolume : '', 'Nilai Hari Itu (Rp)': bh ? Math.round(bh.deltaVolume * hargaSatuan) : '',
          'Bobot Hari Itu thd Proyek (%)': bh ? Number(bh.bobotHarian.toFixed(2)) : '',
          'Kumulatif thd Pekerjaan (%)': bh ? Number(bh.persenPekerjaanIni.toFixed(1)) : '',
          Cuaca: e.cuaca ? CUACA_LIST[e.cuaca]?.label : '', Catatan: e.note || '',
        };
      });

      const sheetName = (bucket.label.split(' (')[0]).slice(0, 28) || `Bagian ${idx + 1}`;
      const sheet = XLSX.utils.json_to_sheet(rekapRows);
      XLSX.utils.sheet_add_json(sheet, aktivitasRows, { origin: `A${rekapRows.length + 3}`, skipHeader: false });
      XLSX.utils.book_append_sheet(wb, sheet, sheetName);
    });

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${periode}-${(project?.name || 'proyek').replace(/\s+/g, '-').toLowerCase()}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  } catch (err) {
    console.error('Gagal ekspor Excel laporan:', err);
    alert('Gagal membuat file Excel. Coba lagi.');
  }
}

function laporanPeriodeHtml(project, allEntries, rabItems, rabTotalNilai, periode) {
  const bucketList = bucketLabelDanEntri(project, allEntries, rabItems, rabTotalNilai, periode);
  const judulLaporan = { harian: 'Laporan Harian', mingguan: 'Laporan Mingguan', bulanan: 'Laporan Bulanan' }[periode];

  const sections = bucketList.map((bucket) => {
    const rekap = hitungRekapBobot(rabItems, allEntries, rabTotalNilai, project?.id, bucket.endStr);
    const rekapRows = rekap.rows.map((r) => `
      <tr>
        <td>${escapeHtml(r.uraian)}</td><td>${r.status}</td><td>${r.volSelesai}/${r.volTotal} ${escapeHtml(r.satuan)}</td>
        <td style="text-align:right;">${r.bobot.toFixed(2)}%</td><td style="text-align:right;">${r.bobotTercapai.toFixed(2)}% (${formatRupiah(r.nilaiTercapai)})</td>
      </tr>`).join('');
    const aktivitasRows = bucket.entri.map((e) => {
      const bh = bobotHarianEntry(e, allEntries, rabItems, rabTotalNilai);
      return `<tr>
        <td>${e.date}</td><td>${escapeHtml(e.uraian || '-')}</td>
        <td>${e.satuan ? `${e.volumeSelesai}/${e.volumeTotal} ${escapeHtml(e.satuan)}` : '-'}</td>
        <td style="text-align:right;">${bh ? `+${bh.deltaVolume} ${escapeHtml(e.satuan)} (${formatRupiah(bh.deltaVolume * (Number(rabItems.find((it) => it.uraian === e.uraian)?.hargaSatuan) || 0))})` : '-'}</td>
        <td style="text-align:right;">${bh ? `+${bh.bobotHarian.toFixed(2)}%` : '-'}</td>
      </tr>`;
    }).join('');
    return `
      <h3 style="margin-bottom:4px;">${escapeHtml(bucket.label)}</h3>
      <p style="font-size:0.85em;font-weight:bold;margin:8px 0 2px;">Rekap Bobot Pekerjaan (per tanggal ${bucket.endStr})</p>
      <table style="width:100%;border-collapse:collapse;font-size:0.8em;" border="1" cellpadding="4">
        <thead><tr style="background:#eee;"><th>Uraian</th><th>Status</th><th>Volume</th><th>Bobot Rencana</th><th>Bobot Tercapai</th></tr></thead>
        <tbody>${rekapRows}</tbody>
        <tfoot><tr style="font-weight:bold;"><td colspan="4" style="text-align:right;">TOTAL BOBOT TERCAPAI</td><td style="text-align:right;">${rekap.bobotTotal.toFixed(2)}% (${formatRupiah(rekap.nilaiTotal)})</td></tr></tfoot>
      </table>
      <p style="font-size:0.85em;font-weight:bold;margin:12px 0 2px;">Aktivitas (${bucket.entri.length} entri)</p>
      ${bucket.entri.length === 0 ? '<p style="font-size:0.8em;font-style:italic;">Tidak ada aktivitas pada periode ini.</p>' : `
      <table style="width:100%;border-collapse:collapse;font-size:0.78em;margin-bottom:18px;" border="1" cellpadding="4">
        <thead><tr style="background:#eee;"><th>Tanggal</th><th>Uraian</th><th>Volume</th><th>Vol.+Nilai Hari Itu</th><th>Bobot thd Proyek</th></tr></thead>
        <tbody>${aktivitasRows}</tbody>
      </table>`}
    `;
  }).join('<hr style="margin:18px 0;"/>');

  return `<p><b>${escapeHtml(judulLaporan)}</b> — ${escapeHtml(project?.name || '-')}</p>${sections || '<p>Belum ada aktivitas progres tercatat.</p>'}`;
}

function exportLaporanPeriodePDF(project, allEntries, rabItems, rabTotalNilai, periode) {
  try {
    const doc = new window.jspdf.jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    const projectId = project?.id;
    let y;

    if (periode === 'harian') {
      const todayStr = new Date().toISOString().slice(0, 10);
      const todayEntries = allEntries.filter((e) => e.projectId === projectId && e.date === todayStr);
      y = pdfAddHeader(doc, 'Laporan Harian', project?.name || '-');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10.5); doc.setTextColor(...PDF_NAVY);
      doc.text(`Tanggal: ${todayStr}`, 40, y);
      y += 18;

      const rekap = gambarRekapBobot(doc, y, rabItems, allEntries, rabTotalNilai, projectId, todayStr);
      y = rekap.y;

      if (y > 640) { doc.addPage(); y = 40; }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10.5); doc.setTextColor(...PDF_NAVY);
      doc.text('Aktivitas Hari Ini', 40, y);
      y += 8;
      if (todayEntries.length === 0) {
        doc.setFont('helvetica', 'italic'); doc.setFontSize(9); doc.setTextColor(...PDF_SLATE);
        doc.text('Tidak ada aktivitas/foto progres tercatat hari ini.', 40, y + 14);
      } else {
        const rows = todayEntries.map((e) => {
          const bh = bobotHarianEntry(e, allEntries, rabItems, rabTotalNilai);
          const rabItem = rabItems.find((it) => it.uraian === e.uraian);
          const hargaSatuan = Number(rabItem?.hargaSatuan) || 0;
          const nilaiHariIni = bh ? bh.deltaVolume * hargaSatuan : 0;
          const nilaiKumulatifItem = bh ? (Number(e.volumeSelesai) || 0) * hargaSatuan : 0;
          return [
            e.date, e.uraian || '-', e.satuan ? `${e.volumeSelesai}/${e.volumeTotal} ${e.satuan}` : '-',
            bh ? `+${bh.deltaVolume} ${e.satuan} (${formatRupiah(nilaiHariIni)})` : '-',
            bh ? `${(bh.bobotHarian / (bh.bobotItem || 1) * 100).toFixed(1)}%` : '-',
            bh ? `+${bh.bobotHarian.toFixed(2)}% (${formatRupiah(rabTotalNilai * bh.bobotHarian / 100)})` : '-',
            bh ? `${bh.persenPekerjaanIni.toFixed(1)}%${bh.selesai ? ' ✓' : ''} (${formatRupiah(nilaiKumulatifItem)})` : '-',
            e.cuaca ? CUACA_LIST[e.cuaca]?.label : '-',
          ];
        });
        doc.autoTable({
          startY: y + 6, margin: { left: 40, right: 40 },
          head: [['Tgl', 'Uraian', 'Volume', 'Vol. + Nilai Hari Ini', 'Bobot thd Pekerjaan', 'Bobot + Nilai thd Proyek', 'Kumulatif thd Pekerjaan + Nilai', 'Cuaca']],
          body: rows,
          styles: { fontSize: 6, textColor: PDF_INK },
          headStyles: { fillColor: PDF_NAVY, textColor: [242, 236, 217] },
          alternateRowStyles: { fillColor: [250, 248, 241] },
        });
      }
      pdfFinishAllPages(doc);
      doc.save(`harian-${(project?.name || 'proyek').replace(/\s+/g, '-').toLowerCase()}.pdf`);
      return;
    }

    // mingguan (7 hari/bucket) atau bulanan (28 hari = 4 minggu/bucket)
    const hariPerBucket = periode === 'mingguan' ? 7 : 28;
    const buckets = buildPeriodeBuckets(allEntries, projectId, hariPerBucket);
    const judulBucket = periode === 'mingguan' ? 'Minggu ke-' : 'Bulan ke-';
    y = pdfAddHeader(doc, periode === 'mingguan' ? 'Laporan Mingguan' : 'Laporan Bulanan', project?.name || '-');

    if (buckets.length === 0) {
      doc.setFont('helvetica', 'italic'); doc.setFontSize(9); doc.setTextColor(...PDF_SLATE);
      doc.text('Belum ada aktivitas progres tercatat untuk proyek ini.', 40, y);
    }

    buckets.forEach((bucket) => {
      if (y > 640) { doc.addPage(); y = 40; }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...PDF_NAVY);
      doc.text(`${judulBucket}${bucket.index} (${bucket.startStr} s/d ${bucket.endStr})`, 40, y);
      y += 16;

      const entriBucket = allEntries
        .filter((e) => e.projectId === projectId && e.date >= bucket.startStr && e.date <= bucket.endStr)
        .sort((a, b) => (a.date < b.date ? -1 : 1));

      if (entriBucket.length === 0) {
        doc.setFont('helvetica', 'italic'); doc.setFontSize(8.5); doc.setTextColor(...PDF_SLATE);
        doc.text('Tidak ada aktivitas pada periode ini.', 40, y);
        y += 16;
      } else {
        const rows = entriBucket.map((e) => {
          const bh = bobotHarianEntry(e, allEntries, rabItems, rabTotalNilai);
          const rabItem = rabItems.find((it) => it.uraian === e.uraian);
          const hargaSatuan = Number(rabItem?.hargaSatuan) || 0;
          const nilaiHariItu = bh ? bh.deltaVolume * hargaSatuan : 0;
          const nilaiKumulatifItem = bh ? (Number(e.volumeSelesai) || 0) * hargaSatuan : 0;
          return [
            e.date, e.uraian || '-', e.satuan ? `${e.volumeSelesai}/${e.volumeTotal} ${e.satuan}` : '-',
            bh ? `+${bh.deltaVolume} ${e.satuan} (${formatRupiah(nilaiHariItu)})` : '-',
            bh ? `+${bh.bobotHarian.toFixed(2)}% (${formatRupiah(rabTotalNilai * bh.bobotHarian / 100)})` : '-',
            bh ? `${bh.persenPekerjaanIni.toFixed(1)}%${bh.selesai ? ' ✓' : ''} (${formatRupiah(nilaiKumulatifItem)})` : '-',
          ];
        });
        doc.autoTable({
          startY: y, margin: { left: 40, right: 40 },
          head: [['Tanggal', 'Uraian', 'Volume', 'Vol. + Nilai Hari Itu', 'Bobot + Nilai thd Proyek', 'Kumulatif thd Pekerjaan + Nilai']],
          body: rows,
          styles: { fontSize: 6, textColor: PDF_INK },
          headStyles: { fillColor: PDF_NAVY, textColor: [242, 236, 217] },
          alternateRowStyles: { fillColor: [250, 248, 241] },
        });
        y = doc.lastAutoTable.finalY + 10;
      }

      // Rekap Bobot Pekerjaan kumulatif s.d. akhir bucket ini (bucket 1 + 2 + ... + bucket ini)
      if (y > 600) { doc.addPage(); y = 40; }
      const rekap = gambarRekapBobot(doc, y, rabItems, allEntries, rabTotalNilai, projectId, bucket.endStr);
      y = rekap.y + 14;
    });

    pdfFinishAllPages(doc);
    doc.save(`${periode}-${(project?.name || 'proyek').replace(/\s+/g, '-').toLowerCase()}.pdf`);
  } catch (err) {
    console.error('Gagal ekspor laporan periode:', err);
    alert('Gagal membuat file PDF. Coba lagi.');
  }
}

function exportProgresFotoPDF(project, entries) {
  try {
    const doc = new window.jspdf.jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    let y = pdfAddHeader(doc, 'Progres Proyek', project?.name || '-');

    entries.forEach((entry) => {
      if (y > 600) { doc.addPage(); y = 40; }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(...PDF_NAVY);
      doc.text(entry.date, 40, y);
      if (entry.uraian) doc.text(entry.uraian, 215, y, { maxWidth: 320 });
      y += 14;
      if (entry.satuan && (entry.volumeTotal > 0 || entry.volumeSelesai > 0)) {
        doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(...PDF_GOLD);
        doc.text(`Volume: ${entry.volumeSelesai}/${entry.volumeTotal} ${entry.satuan}  |  Sisa: ${entry.volumeSisa} ${entry.satuan}`, 215, y, { maxWidth: 320 });
        y += 14;
      }
      if (entry.photo) {
        try {
          doc.addImage(entry.photo, 'JPEG', 40, y - 14, 160, 120);
        } catch { /* lewati kalau format foto tidak didukung addImage */ }
      }
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...PDF_INK);
      doc.text(entry.note || '(tanpa catatan)', 215, y + 6, { maxWidth: 320 });
      y += 130;
    });

    pdfFinishAllPages(doc);
    doc.save(`progres-${(project?.name || 'proyek').replace(/\s+/g, '-').toLowerCase()}.pdf`);
  } catch (err) {
    console.error('Gagal ekspor PDF Progres Proyek:', err);
    alert('Gagal membuat file PDF. Coba lagi.');
  }
}

/* ---------------- Data Klien (CRM ringan) ---------------- */
const KLIEN_JENIS = { perorangan: 'Perorangan', perusahaan: 'Perusahaan/Instansi' };
const KLIEN_STATUS = {
  prospek: 'Prospek', follow_up: 'Follow-up', deal: 'Deal', tidak_jadi: 'Tidak Jadi',
};
const KLIEN_STATUS_COLORS = { prospek: '#5B6478', follow_up: '#C9A227', deal: '#2F7A52', tidak_jadi: '#B23A2E' };
const KLIEN_SUMBER = ['Website', 'Rekomendasi', 'Medsos', 'Datang Langsung', 'Lainnya'];

function emptyKlienForm() {
  return {
    nama: '', jenis: 'perusahaan', namaPIC: '', jabatanPIC: '', telepon: '', email: '',
    alamat: '', sumber: KLIEN_SUMBER[0], status: 'prospek', proyekTerkait: '', catatan: '',
  };
}

function KlienForm({ onSave, onCancel, initial }) {
  const [form, setForm] = useState(initial ? { ...emptyKlienForm(), ...initial } : emptyKlienForm());
  const set = (patch) => setForm((f) => ({ ...f, ...patch }));

  const handleSave = async () => {
    if (!form.nama.trim()) {
      alert('Isi nama klien/perusahaan dulu.');
      return;
    }
    await onSave({ id: initial?.id || uid(), ...form, nama: form.nama.trim() });
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(13,25,48,0.7)' }}>
      <div className="w-full max-w-md rounded-xl p-4 att-body max-h-[92vh] overflow-y-auto" style={{ background: THEME.paper }}>
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>{initial ? 'Edit' : 'Tambah'} Data Klien</h3>
          <button type="button" onClick={onCancel}><X size={18} color={THEME.inkSoft} /></button>
        </div>

        <input value={form.nama} onChange={(e) => set({ nama: e.target.value })} placeholder="Nama Klien / Perusahaan"
          className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-2" style={{ background: THEME.concrete, color: THEME.ink }} />

        <div className="grid grid-cols-2 gap-2 mb-2">
          <select value={form.jenis} onChange={(e) => set({ jenis: e.target.value })}
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }}>
            {Object.entries(KLIEN_JENIS).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select>
          <select value={form.status} onChange={(e) => set({ status: e.target.value })}
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }}>
            {Object.entries(KLIEN_STATUS).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select>
        </div>

        <p className="att-mono text-[10px] mb-1.5" style={{ color: THEME.inkSoft }}>NARAHUBUNG (PIC)</p>
        <div className="grid grid-cols-2 gap-2 mb-2">
          <input value={form.namaPIC} onChange={(e) => set({ namaPIC: e.target.value })} placeholder="Nama PIC"
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={form.jabatanPIC} onChange={(e) => set({ jabatanPIC: e.target.value })} placeholder="Jabatan PIC"
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>
        <div className="grid grid-cols-2 gap-2 mb-2">
          <input value={form.telepon} onChange={(e) => set({ telepon: e.target.value })} placeholder="No. HP/Telepon"
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={form.email} onChange={(e) => set({ email: e.target.value })} placeholder="Email"
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>
        <input value={form.alamat} onChange={(e) => set({ alamat: e.target.value })} placeholder="Alamat"
          className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-2" style={{ background: THEME.concrete, color: THEME.ink }} />

        <div className="grid grid-cols-2 gap-2 mb-2">
          <select value={form.sumber} onChange={(e) => set({ sumber: e.target.value })}
            className="px-3 py-2 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }}>
            {KLIEN_SUMBER.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <input value={form.proyekTerkait} onChange={(e) => set({ proyekTerkait: e.target.value })} placeholder="Proyek terkait (opsional)"
            className="px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
        </div>

        <p className="att-mono text-[10px] mb-1.5" style={{ color: THEME.inkSoft }}>CATATAN / RIWAYAT KOMUNIKASI</p>
        <textarea value={form.catatan} onChange={(e) => set({ catatan: e.target.value })} rows={4}
          placeholder="Mis. 12/07 - Telepon awal, tertarik renovasi kantor. 20/07 - Kirim penawaran, tunggu konfirmasi."
          className="w-full px-3 py-2 rounded att-body text-xs outline-none mb-3" style={{ background: THEME.concrete, color: THEME.ink }} />

        <button type="button" onClick={handleSave}
          className="w-full flex items-center justify-center gap-1.5 py-2.5 rounded font-semibold text-sm"
          style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Check size={15} /> Simpan Data Klien
        </button>
      </div>
    </div>
  );
}

function DataKlienTab({ klienList, onAdd, onUpdate, onDelete }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [search, setSearch] = useState('');
  const [filterStatus, setFilterStatus] = useState('all');

  const handleSave = async (data) => {
    if (editing) await onUpdate(editing.id, data);
    else await onAdd(data);
    setShowForm(false);
    setEditing(null);
  };

  const filtered = klienList
    .filter((k) => filterStatus === 'all' || k.status === filterStatus)
    .filter((k) => !search.trim() || k.nama.toLowerCase().includes(search.toLowerCase()) || (k.namaPIC || '').toLowerCase().includes(search.toLowerCase()))
    .sort((a, b) => (a.nama || '').localeCompare(b.nama || ''));

  return (
    <div className="p-4 pb-24">
      <div className="flex items-center justify-between mb-3">
        <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>DATA KLIEN ({klienList.length})</p>
        <button type="button" onClick={() => { setEditing(null); setShowForm(true); }}
          className="flex items-center gap-1.5 px-3 py-1.5 rounded att-mono text-xs font-semibold" style={{ background: THEME.amber, color: THEME.charcoal }}>
          <Plus size={13} /> Tambah Klien
        </button>
      </div>

      <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Cari nama klien atau PIC..."
        className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-2" style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }} />

      <div className="flex gap-1.5 mb-3 overflow-x-auto att-scroll">
        {['all', ...Object.keys(KLIEN_STATUS)].map((s) => (
          <button key={s} type="button" onClick={() => setFilterStatus(s)}
            className="px-3 py-1.5 rounded-full att-mono text-[10.5px] font-semibold shrink-0"
            style={{ background: filterStatus === s ? THEME.charcoal : THEME.concrete, color: filterStatus === s ? THEME.paper : THEME.inkSoft }}>
            {s === 'all' ? 'Semua' : KLIEN_STATUS[s]}
          </button>
        ))}
      </div>

      <div className="space-y-2">
        {filtered.map((k) => (
          <div key={k.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div className="flex items-start justify-between mb-1">
              <div className="min-w-0">
                <p className="att-body font-semibold text-sm truncate" style={{ color: THEME.ink }}>{k.nama}</p>
                <p className="att-mono text-[11px]" style={{ color: THEME.inkSoft }}>{KLIEN_JENIS[k.jenis]}{k.namaPIC ? ` · PIC: ${k.namaPIC}` : ''}</p>
              </div>
              <span className="att-mono text-[9px] px-1.5 py-0.5 rounded-full font-semibold shrink-0" style={{ background: KLIEN_STATUS_COLORS[k.status], color: '#fff' }}>
                {KLIEN_STATUS[k.status]}
              </span>
            </div>
            {(k.telepon || k.email) && (
              <p className="att-mono text-[11px] mt-1" style={{ color: THEME.inkSoft }}>{[k.telepon, k.email].filter(Boolean).join(' · ')}</p>
            )}
            {k.proyekTerkait && (
              <p className="att-mono text-[11px] mt-0.5" style={{ color: THEME.amber }}>Proyek: {k.proyekTerkait}</p>
            )}
            <div className="flex items-center gap-1.5 mt-2">
              <button type="button" onClick={() => { setEditing(k); setShowForm(true); }}
                className="flex items-center gap-1 px-2.5 py-1.5 rounded att-mono text-[11px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.concrete }}>
                <Edit3 size={12} /> Edit
              </button>
              {confirmDelete === k.id ? (
                <button type="button" onClick={() => { onDelete(k.id); setConfirmDelete(null); }} className="p-1.5 rounded ml-auto" style={{ background: THEME.rust }}>
                  <Check size={13} color={THEME.paper} />
                </button>
              ) : (
                <button type="button" onClick={() => setConfirmDelete(k.id)} className="p-1.5 rounded ml-auto" style={{ background: THEME.concrete }}>
                  <Trash2 size={13} color={THEME.rust} />
                </button>
              )}
            </div>
          </div>
        ))}
        {filtered.length === 0 && (
          <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada data klien. Tap "Tambah Klien" untuk mulai.</p>
        )}
      </div>

      {showForm && (
        <KlienForm initial={editing} onSave={handleSave} onCancel={() => { setShowForm(false); setEditing(null); }} />
      )}
    </div>
  );
}

/* ---------------- Kalkulator Dimensi (RAB & Progres) ---------------- */
const DIMENSI_FORMULA = {
  manual: { label: 'Manual (langsung isi angka)', fields: [] },
  kotak: { label: 'Volume Kotak/Balok — P×L×T×N', fields: ['p', 'l', 't', 'n'] },
  luas: { label: 'Luas — P×L×N', fields: ['p', 'l', 'n'] },
  trapesium: { label: 'Trapesium (mis. Pondasi Batu Kali) — ((La+Lb)/2×T)×P×N', fields: ['p', 'la', 'lb', 't', 'n'] },
  batang: { label: 'Berat Batang (mis. Kolom IWF/Besi) — P×N×Berat/m', fields: ['p', 'n', 'beratPerM'] },
};
const DIMENSI_FIELD_LABEL = {
  p: 'Panjang (P)', l: 'Lebar (L)', t: 'Tinggi (T)', n: 'Jumlah (N)',
  la: 'Lebar Atas (La)', lb: 'Lebar Bawah (Lb)', beratPerM: 'Berat per Meter (kg/m)',
};
function hitungDimensi(formulaType, d) {
  const p = Number(d.p) || 0, l = Number(d.l) || 0, t = Number(d.t) || 0, n = Number(d.n) || 1;
  const la = Number(d.la) || 0, lb = Number(d.lb) || 0, beratPerM = Number(d.beratPerM) || 0;
  switch (formulaType) {
    case 'kotak': return p * l * t * n;
    case 'luas': return p * l * n;
    case 'trapesium': return ((la + lb) / 2) * t * p * n;
    case 'batang': return p * n * beratPerM;
    default: return null;
  }
}

function DimensiCalculator({ onApply }) {
  const [open, setOpen] = useState(false);
  const [formulaType, setFormulaType] = useState('manual');
  const [d, setD] = useState({});
  const hasil = hitungDimensi(formulaType, d);

  if (!open) {
    return (
      <button type="button" onClick={() => setOpen(true)}
        className="att-mono text-[10px] font-semibold mb-2" style={{ color: THEME.amber }}>
        📐 Pakai Kalkulator Dimensi
      </button>
    );
  }
  return (
    <div className="p-2.5 rounded-lg mb-2" style={{ background: THEME.concrete }}>
      <div className="flex items-center justify-between mb-1.5">
        <p className="att-mono text-[10px]" style={{ color: THEME.inkSoft }}>KALKULATOR DIMENSI</p>
        <button type="button" onClick={() => setOpen(false)}><X size={13} color={THEME.inkSoft} /></button>
      </div>
      <select value={formulaType} onChange={(e) => { setFormulaType(e.target.value); setD({}); }}
        className="w-full px-2 py-1.5 rounded att-mono text-[10.5px] outline-none mb-1.5" style={{ background: THEME.paper, color: THEME.ink }}>
        {Object.entries(DIMENSI_FORMULA).map(([v, f]) => <option key={v} value={v}>{f.label}</option>)}
      </select>
      {formulaType !== 'manual' && (
        <>
          <div className="grid grid-cols-2 gap-1.5 mb-1.5">
            {DIMENSI_FORMULA[formulaType].fields.map((f) => (
              <input key={f} value={d[f] || ''} onChange={(e) => setD({ ...d, [f]: e.target.value.replace(/[^0-9.]/g, '') })}
                placeholder={DIMENSI_FIELD_LABEL[f]} inputMode="decimal"
                className="px-2 py-1.5 rounded att-mono text-[10.5px] outline-none" style={{ background: THEME.paper, color: THEME.ink }} />
            ))}
          </div>
          <p className="att-mono text-[10.5px] mb-1.5" style={{ color: THEME.amber }}>Hasil: {hasil !== null ? hasil.toLocaleString('id-ID', { maximumFractionDigits: 3 }) : '-'}</p>
          <button type="button" onClick={() => { if (hasil !== null) { onApply(hasil); setOpen(false); } }}
            className="w-full py-1.5 rounded att-mono text-[10.5px] font-semibold" style={{ background: THEME.amber, color: THEME.charcoal }}>
            Terapkan ke Volume
          </button>
        </>
      )}
    </div>
  );
}

/* ---------------- Cuaca ---------------- */
const CUACA_LIST = {
  cerah: { label: 'Cerah', color: '#2F7A52', kerja: 'auto_on' },
  mendung: { label: 'Mendung', color: '#C9A227', kerja: 'pilihan' },
  gerimis: { label: 'Hujan Gerimis', color: '#E091A8', kerja: 'pilihan' },
  deras: { label: 'Hujan Deras', color: '#8A1F1F', kerja: 'auto_stop' },
};

function rabItemStatus(rabItem, allEntries, projectId) {
  const matching = allEntries.filter((e) => e.projectId === projectId && e.uraian === rabItem.uraian);
  const latest = [...matching].sort((a, b) => (a.date < b.date ? 1 : -1))[0];
  const volumeSelesai = latest ? Number(latest.volumeSelesai) || 0 : 0;
  const volumeTotal = Number(rabItem.volume) || 0;
  let status = 'off';
  if (volumeTotal > 0 && volumeSelesai >= volumeTotal) status = 'sukses';
  else if (volumeSelesai > 0) status = 'proses';
  return { volumeSelesai, volumeTotal, status, latestDate: latest?.date };
}
const RAB_STATUS_LABEL = { off: 'OFF', proses: 'ON PROSES', sukses: 'SELESAI' };
const RAB_STATUS_COLOR = { off: '#B23A2E', proses: '#C9A227', sukses: '#2F7A52' };

function bobotHarianEntry(entry, allEntries, rabItems, rabTotalNilai) {
  if (!entry.uraian || !entry.satuan) return null;
  const rabItem = rabItems.find((it) => it.uraian === entry.uraian);
  if (!rabItem) return null;
  const volumeTotalRab = Number(rabItem.volume) || 0;
  if (volumeTotalRab <= 0 || rabTotalNilai <= 0) return null;

  const nilaiItem = volumeTotalRab * (Number(rabItem.hargaSatuan) || 0);
  const bobotItem = (nilaiItem / rabTotalNilai) * 100;

  // Entri lain untuk pekerjaan yang sama, di proyek yang sama, sebelum tanggal entri ini
  const riwayatSama = allEntries
    .filter((e) => e.projectId === entry.projectId && e.uraian === entry.uraian && e.date <= entry.date && e.id !== entry.id)
    .sort((a, b) => (a.date < b.date ? 1 : -1));
  const sebelumnya = riwayatSama.find((e) => e.date < entry.date);
  const volSebelum = sebelumnya ? Number(sebelumnya.volumeSelesai) || 0 : 0;
  const volSekarang = Number(entry.volumeSelesai) || 0;
  const deltaVolume = Math.max(0, volSekarang - volSebelum);
  const bobotHarian = bobotItem * (deltaVolume / volumeTotalRab);
  const persenPekerjaanIni = Math.min(100, (volSekarang / volumeTotalRab) * 100);
  const selesai = volSekarang >= volumeTotalRab;
  return { bobotItem, deltaVolume, bobotHarian, persenPekerjaanIni, selesai };
}

function ProgresFotoTab({ projects, entries, onAdd, onDelete, penawaranList }) {
  const [projectId, setProjectId] = useState(projects[0]?.id || '');
  const [activeItem, setActiveItem] = useState(null); // rabItem yang lagi diisi progresnya (atau { free: true } utk catatan bebas)
  const [uraian, setUraian] = useState('');
  const [volumeTotal, setVolumeTotal] = useState('');
  const [volumeSelesai, setVolumeSelesai] = useState('');
  const [satuan, setSatuan] = useState('');
  const [note, setNote] = useState('');
  const [cuaca, setCuaca] = useState('cerah');
  const [statusKerja, setStatusKerja] = useState('lanjut'); // 'lanjut' | 'stop', dipakai saat cuaca mendung/gerimis
  const [photoDataUri, setPhotoDataUri] = useState('');
  const [photoBusy, setPhotoBusy] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(null);

  const rabItems = (penawaranList || [])
    .filter((p) => p.projectId === projectId)
    .flatMap((p) => p.items || []);
  const rabTotalNilai = rabItems.reduce((s, it) => s + (Number(it.volume) || 0) * (Number(it.hargaSatuan) || 0), 0);

  const openForRab = (item, currentSelesai) => {
    setActiveItem(item);
    setUraian(item.uraian);
    setVolumeTotal(String(item.volume || ''));
    setSatuan(item.satuan || '');
    setVolumeSelesai(currentSelesai > 0 ? String(currentSelesai) : '');
    setNote(''); setCuaca('cerah'); setStatusKerja('lanjut');
    setPhotoDataUri('');
  };
  const openFree = () => {
    setActiveItem({ free: true });
    setUraian(''); setVolumeTotal(''); setSatuan(''); setVolumeSelesai(''); setNote('');
    setCuaca('cerah'); setStatusKerja('lanjut'); setPhotoDataUri('');
  };
  const closeForm = () => setActiveItem(null);

  const handlePhotoFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setPhotoBusy(true);
    try {
      const dataUri = await resizeImage(file);
      setPhotoDataUri(dataUri);
    } catch {
      alert('Gagal memuat foto. Coba lagi.');
    }
    setPhotoBusy(false);
    e.target.value = '';
  };

  const handleSave = async () => {
    if (!projectId || !photoDataUri) {
      alert('Ambil/pilih foto dulu.');
      return;
    }
    if (CUACA_LIST[cuaca].kerja === 'auto_stop') {
      alert('Cuaca Hujan Deras — pekerjaan otomatis dihentikan (Stop). Progres hari ini tidak bisa disimpan sebagai lanjutan kerja.');
      return;
    }
    const today = new Date().toISOString().slice(0, 10);
    const vTotal = Number(volumeTotal) || 0;
    const vSelesai = Number(volumeSelesai) || 0;
    const efektifStatusKerja = CUACA_LIST[cuaca].kerja === 'auto_on' ? 'lanjut' : statusKerja;
    await onAdd({
      id: uid(), projectId, date: today, note: note.trim(), photo: photoDataUri,
      uraian: uraian.trim(), satuan: satuan.trim(), cuaca, statusKerja: efektifStatusKerja,
      volumeTotal: vTotal, volumeSelesai: vSelesai, volumeSisa: Math.max(0, vTotal - vSelesai),
    });
    closeForm();
  };

  const projectEntries = entries
    .filter((e) => e.projectId === projectId)
    .sort((a, b) => (a.date < b.date ? 1 : -1));
  const activeProject = projects.find((p) => p.id === projectId);

  return (
    <div className="p-4 pb-24">
      <p className="att-mono text-xs mb-3" style={{ color: THEME.inkSoft }}>PROGRES FOTO PROYEK</p>

      <select value={projectId} onChange={(e) => { setProjectId(e.target.value); setActiveItem(null); }}
        className="w-full px-3 py-2.5 rounded att-body text-sm outline-none mb-3" style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }}>
        <option value="">— Pilih Proyek —</option>
        {projects.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
      </select>

      {projects.length === 0 && (
        <p className="att-body text-sm text-center py-4" style={{ color: THEME.inkSoft }}>Belum ada proyek. Buat lewat tab Absensi Mingguan.</p>
      )}

      {projectId && (
        <>
          <div className="flex items-center justify-between mb-2">
            <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>PEKERJAAN DARI RAB ({rabItems.length})</p>
            <button type="button" onClick={openFree}
              className="att-mono text-[10.5px] font-semibold" style={{ color: THEME.amber }}>+ Catatan Bebas</button>
          </div>
          <div className="space-y-2 mb-5">
            {rabItems.map((item, idx) => {
              const st = rabItemStatus(item, entries, projectId);
              const nilaiItem = (Number(item.volume) || 0) * (Number(item.hargaSatuan) || 0);
              const bobot = rabTotalNilai > 0 ? (nilaiItem / rabTotalNilai) * 100 : 0;
              const rasioSelesai = st.volumeTotal > 0 ? Math.min(1, st.volumeSelesai / st.volumeTotal) : 0;
              const bobotTercapai = bobot * rasioSelesai;
              return (
                <button key={item.id || idx} type="button" onClick={() => openForRab(item, st.volumeSelesai)}
                  className="w-full text-left p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
                  <div className="flex items-start justify-between gap-2">
                    <p className="att-body font-semibold text-sm flex-1" style={{ color: THEME.ink }}>{item.uraian}</p>
                    <span className="att-mono text-[9px] px-2 py-0.5 rounded-full font-bold shrink-0" style={{ background: RAB_STATUS_COLOR[st.status], color: '#fff' }}>
                      {RAB_STATUS_LABEL[st.status]}
                    </span>
                  </div>
                  <p className="att-mono text-[10.5px] mt-1" style={{ color: THEME.inkSoft }}>
                    {st.volumeSelesai}/{st.volumeTotal} {item.satuan} {st.latestDate ? `· terakhir ${st.latestDate}` : '· belum ada foto'}
                  </p>
                  <p className="att-mono text-[10px] mt-1" style={{ color: THEME.amber }}>
                    Bobot pekerjaan: {bobot.toFixed(1)}% dari proyek &middot; Tercapai: {bobotTercapai.toFixed(1)}%
                  </p>
                </button>
              );
            })}
            {rabItems.length === 0 && (
              <p className="att-body text-sm text-center py-4" style={{ color: THEME.inkSoft }}>
                Belum ada RAB yang terhubung ke proyek ini. Buat/hubungkan dulu lewat tab "Penawaran", atau tap "+ Catatan Bebas" di atas.
              </p>
            )}
          </div>
        </>
      )}

      {activeItem && (
        <div className="fixed inset-0 z-[60] flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(13,25,48,0.75)' }}>
          <div className="w-full max-w-md rounded-xl p-4 att-body max-h-[92vh] overflow-y-auto" style={{ background: THEME.paper }}>
            <div className="flex items-center justify-between mb-3">
              <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>Tambah Foto Progres Hari Ini</h3>
              <button type="button" onClick={closeForm}><X size={18} color={THEME.inkSoft} /></button>
            </div>

            {photoDataUri ? (
              <div className="mb-2">
                <img src={photoDataUri} alt="Progres" className="w-full rounded-lg mb-1.5" style={{ maxHeight: 180, objectFit: 'cover' }} />
                <button type="button" onClick={() => setPhotoDataUri('')} className="att-mono text-[10px]" style={{ color: THEME.rust }}>Hapus foto, ambil ulang</button>
              </div>
            ) : (
              <div className="flex gap-2 mb-3">
                <label className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded text-sm font-semibold cursor-pointer"
                  style={{ background: THEME.concrete, color: THEME.ink }}>
                  <Camera size={16} /> {photoBusy ? '...' : 'Kamera'}
                  <input type="file" accept="image/*" capture="environment" className="hidden" onChange={handlePhotoFile} disabled={photoBusy} />
                </label>
                <label className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded text-sm font-semibold cursor-pointer"
                  style={{ background: THEME.concrete, color: THEME.ink }}>
                  <Save size={16} /> {photoBusy ? '...' : 'File'}
                  <input type="file" accept="image/*" className="hidden" onChange={handlePhotoFile} disabled={photoBusy} />
                </label>
              </div>
            )}

            {activeItem.free ? (
              <input value={uraian} onChange={(e) => setUraian(e.target.value)} placeholder="Uraian pekerjaan (mis. Pengecoran Kolom Lantai 1)"
                className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-2" style={{ background: THEME.concrete, color: THEME.ink }} />
            ) : (
              <div className="p-2.5 rounded-lg mb-2" style={{ background: THEME.concrete }}>
                <p className="att-mono text-[9.5px]" style={{ color: THEME.inkSoft }}>UraIAN PEKERJAAN (dari RAB, terkunci)</p>
                <p className="att-body text-sm font-semibold" style={{ color: THEME.ink }}>{uraian}</p>
              </div>
            )}

            <DimensiCalculator onApply={(hasil) => setVolumeSelesai(String(hasil))} />
            <div className="grid grid-cols-3 gap-2 mb-2">
              <input value={volumeTotal} onChange={(e) => activeItem.free && setVolumeTotal(e.target.value.replace(/[^0-9.]/g, ''))} placeholder="Volume total"
                readOnly={!activeItem.free} inputMode="decimal"
                className="px-2 py-2 rounded att-body text-sm outline-none" style={{ background: activeItem.free ? THEME.concrete : THEME.line, color: THEME.ink }} />
              <input value={volumeSelesai} onChange={(e) => setVolumeSelesai(e.target.value.replace(/[^0-9.]/g, ''))} placeholder="Vol. selesai" inputMode="decimal"
                className="px-2 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
              <input value={satuan} onChange={(e) => activeItem.free && setSatuan(e.target.value)} placeholder="Satuan" readOnly={!activeItem.free}
                className="px-2 py-2 rounded att-body text-sm outline-none" style={{ background: activeItem.free ? THEME.concrete : THEME.line, color: THEME.ink }} />
            </div>
            {volumeTotal && volumeSelesai && (
              <p className="att-mono text-[10.5px] mb-2" style={{ color: Number(volumeSelesai) >= Number(volumeTotal) ? THEME.green : THEME.inkSoft }}>
                {Number(volumeSelesai) >= Number(volumeTotal) ? '✓ Volume tercapai — pekerjaan SELESAI' : `Sisa volume: ${Math.max(0, (Number(volumeTotal) || 0) - (Number(volumeSelesai) || 0))} ${satuan || ''}`}
              </p>
            )}

            <p className="att-mono text-[10px] mb-1.5" style={{ color: THEME.inkSoft }}>CUACA HARI INI</p>
            <div className="grid grid-cols-4 gap-1.5 mb-2">
              {Object.entries(CUACA_LIST).map(([key, c]) => (
                <button key={key} type="button" onClick={() => { setCuaca(key); if (c.kerja !== 'pilihan') setStatusKerja('lanjut'); }}
                  className="py-2 rounded att-mono text-[9px] font-bold text-center"
                  style={{ background: cuaca === key ? c.color : THEME.concrete, color: cuaca === key ? '#fff' : THEME.inkSoft }}>
                  {c.label}
                </button>
              ))}
            </div>
            {CUACA_LIST[cuaca].kerja === 'auto_on' && (
              <p className="att-mono text-[10px] mb-2" style={{ color: THEME.green }}>● Status: ON — kerja berjalan normal</p>
            )}
            {CUACA_LIST[cuaca].kerja === 'pilihan' && (
              <div className="flex gap-1.5 mb-2">
                <button type="button" onClick={() => setStatusKerja('lanjut')}
                  className="flex-1 py-2 rounded att-mono text-[10.5px] font-semibold" style={{ background: statusKerja === 'lanjut' ? THEME.green : THEME.concrete, color: statusKerja === 'lanjut' ? '#fff' : THEME.inkSoft }}>
                  Lanjut Kerja
                </button>
                <button type="button" onClick={() => setStatusKerja('stop')}
                  className="flex-1 py-2 rounded att-mono text-[10.5px] font-semibold" style={{ background: statusKerja === 'stop' ? THEME.rust : THEME.concrete, color: statusKerja === 'stop' ? '#fff' : THEME.inkSoft }}>
                  Stop Kerja
                </button>
              </div>
            )}
            {CUACA_LIST[cuaca].kerja === 'auto_stop' && (
              <p className="att-mono text-[10px] mb-2" style={{ color: THEME.rust }}>● Status: STOP OTOMATIS — hujan deras, pekerjaan dihentikan</p>
            )}

            <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="Catatan tambahan (opsional)"
              className="w-full px-3 py-2 rounded att-body text-sm outline-none mb-3" style={{ background: THEME.concrete, color: THEME.ink }} />
            <button type="button" onClick={handleSave} disabled={CUACA_LIST[cuaca].kerja === 'auto_stop'}
              className="w-full flex items-center justify-center gap-1.5 py-2.5 rounded font-semibold text-sm disabled:opacity-40"
              style={{ background: THEME.amber, color: THEME.charcoal }}>
              <Check size={15} /> Simpan Progres Hari Ini
            </button>
          </div>
        </div>
      )}

      {projectId && projectEntries.length > 0 && (
        <>
          <p className="att-mono text-xs mb-2" style={{ color: THEME.inkSoft }}>LAPORAN</p>
          <div className="space-y-1.5 mb-4">
            {[
              { key: 'harian', label: 'Harian' },
              { key: 'mingguan', label: 'Mingguan' },
              { key: 'bulanan', label: 'Bulanan' },
            ].map((p) => (
              <div key={p.key} className="flex items-center gap-1.5">
                <span className="att-body text-xs font-semibold w-16 shrink-0" style={{ color: THEME.ink }}>{p.label}</span>
                <button type="button" onClick={() => exportLaporanPeriodePDF(activeProject, projectEntries, rabItems, rabTotalNilai, p.key)}
                  className="flex-1 flex items-center justify-center gap-1 py-1.5 rounded att-mono text-[10px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.paper }}>
                  <Save size={11} /> PDF
                </button>
                <button type="button" onClick={() => exportLaporanPeriodeExcel(activeProject, projectEntries, rabItems, rabTotalNilai, p.key)}
                  className="flex-1 flex items-center justify-center gap-1 py-1.5 rounded att-mono text-[10px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.paper }}>
                  <Save size={11} /> Excel
                </button>
                <button type="button" onClick={() => openPrintDocument(`Laporan ${p.label} - ${activeProject?.name || ''}`, laporanPeriodeHtml(activeProject, projectEntries, rabItems, rabTotalNilai, p.key), false)}
                  className="flex-1 flex items-center justify-center gap-1 py-1.5 rounded att-mono text-[10px] font-semibold" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.paper }}>
                  <Printer size={11} /> Cetak
                </button>
              </div>
            ))}
          </div>
          <p className="att-mono text-xs mb-2" style={{ color: THEME.inkSoft }}>RIWAYAT ({projectEntries.length})</p>
          <div className="space-y-2">
            {projectEntries.map((entry) => (
              <div key={entry.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
                <div className="flex gap-2.5">
                  {entry.photo && (
                    <img src={entry.photo} alt={entry.date} className="w-20 h-20 rounded-lg object-cover shrink-0" style={{ border: `1px solid ${THEME.line}` }} />
                  )}
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-1.5">
                      <p className="att-mono text-xs font-semibold" style={{ color: THEME.ink }}>{entry.date}</p>
                      {entry.cuaca && CUACA_LIST[entry.cuaca] && (
                        <span className="att-mono text-[8px] px-1.5 py-0.5 rounded font-bold" style={{ background: CUACA_LIST[entry.cuaca].color, color: '#fff' }}>
                          {CUACA_LIST[entry.cuaca].label}{entry.statusKerja === 'stop' ? ' · STOP' : ''}
                        </span>
                      )}
                    </div>
                    {entry.uraian && <p className="att-body text-sm font-semibold mt-0.5" style={{ color: THEME.ink }}>{entry.uraian}</p>}
                    {entry.satuan && (entry.volumeTotal > 0 || entry.volumeSelesai > 0) && (
                      <p className="att-mono text-[10.5px] mt-0.5" style={{ color: entry.volumeSelesai >= entry.volumeTotal && entry.volumeTotal > 0 ? THEME.green : THEME.amber }}>
                        {entry.volumeSelesai}/{entry.volumeTotal} {entry.satuan} &middot; sisa {entry.volumeSisa} {entry.satuan}
                      </p>
                    )}
                    {(() => {
                      const bh = bobotHarianEntry(entry, entries, rabItems, rabTotalNilai);
                      if (!bh) return null;
                      return (
                        <div className="mt-1 space-y-1">
                          <p className="att-mono text-[10px] inline-block mr-1" style={{ color: THEME.charcoal, background: THEME.amberSoft, padding: '1px 6px', borderRadius: 4 }}>
                            +{bh.deltaVolume} {entry.satuan} hari ini &middot; bobot proyek: +{bh.bobotHarian.toFixed(2)}%
                          </p>
                          <p className="att-mono text-[10px]" style={{ color: bh.selesai ? THEME.green : THEME.inkSoft }}>
                            Bobot terhadap pekerjaan ini sendiri: {bh.persenPekerjaanIni.toFixed(1)}%
                            {bh.selesai && <span style={{ color: THEME.green, fontWeight: 700 }}> &middot; ✓ SELESAI</span>}
                          </p>
                        </div>
                      );
                    })()}
                    <p className="att-body text-sm mt-0.5" style={{ color: THEME.inkSoft }}>{entry.note || '(tanpa catatan)'}</p>
                  </div>
                  {confirmDelete === entry.id ? (
                    <button type="button" onClick={() => { onDelete(entry.id); setConfirmDelete(null); }} className="p-1.5 rounded shrink-0 h-fit" style={{ background: THEME.rust }}>
                      <Check size={13} color={THEME.paper} />
                    </button>
                  ) : (
                    <button type="button" onClick={() => setConfirmDelete(entry.id)} className="p-1.5 rounded shrink-0 h-fit" style={{ background: THEME.concrete }}>
                      <Trash2 size={13} color={THEME.rust} />
                    </button>
                  )}
                </div>
              </div>
            ))}
          </div>
        </>
      )}
      {projectId && projectEntries.length === 0 && (
        <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada progres foto untuk proyek ini.</p>
      )}
    </div>
  );
}

function RekapProyekTab({ projects, purchases, usage, materials, utangPayments, onAddUtangPayment, workers, weeks, payments, kasbon, kasbonPayments, onUpdateProject, peralatanUsage }) {
  const [showTextPreview, setShowTextPreview] = useState(false);
  const [showTablePreview, setShowTablePreview] = useState(false);
  const [showKelolaProyek, setShowKelolaProyek] = useState(false);
  const [openNota, setOpenNota] = useState(null); // gudangId sedang buka riwayat nota
  const [openStok, setOpenStok] = useState(null); // gudangId sedang buka riwayat stok
  const [openPeralatanUsage, setOpenPeralatanUsage] = useState(null); // gudangId sedang buka pemakaian peralatan
  const [payFor, setPayFor] = useState(null); // { gudangId, supplierKey, supplierName }
  const [payAmount, setPayAmount] = useState('');
  const [payDate, setPayDate] = useState(new Date().toISOString().slice(0, 10));
  const [payMethod, setPayMethod] = useState('cash');
  const [payWho, setPayWho] = useState(''); // nama kasir (cash) atau nama bank (transfer)
  const [payNote, setPayNote] = useState('');
  const [openHistory, setOpenHistory] = useState(null); // `${gudangId}:${supplierKey}`
  const gudangList = getGudangList(projects);

  const perGudang = gudangList.map((g) => {
    const gPurchases = purchases.filter((p) => p.gudangId === g.id);
    const totalBelanja = gPurchases.reduce((s, p) => s + purchaseTotal(p), 0);
    const totalCash = gPurchases.filter((p) => p.statusBayar === 'cash').reduce((s, p) => s + purchaseTotal(p), 0);
    const totalUtangKotor = gPurchases.filter((p) => p.statusBayar !== 'cash').reduce((s, p) => s + purchaseTotal(p), 0);
    const perSupplier = {};
    gPurchases.filter((p) => p.statusBayar !== 'cash').forEach((p) => {
      const key = p.supplierId || p.supplierName || 'lainnya';
      const name = p.supplierName || '(tanpa suplier)';
      if (!perSupplier[key]) perSupplier[key] = { key, name, amount: 0 };
      perSupplier[key].amount += purchaseTotal(p);
    });
    Object.values(perSupplier).forEach((s) => {
      const dibayar = utangPayments
        .filter((up) => up.gudangId === g.id && (up.supplierId || up.supplierName || 'lainnya') === s.key)
        .reduce((sum, up) => sum + (Number(up.amount) || 0), 0);
      s.dibayar = dibayar;
      s.sisa = s.amount - dibayar;
    });
    const totalDibayarUtang = Object.values(perSupplier).reduce((s, x) => s + x.dibayar, 0);
    const totalBelumLunas = totalUtangKotor - totalDibayarUtang;
    return {
      gudang: g, totalBelanja, totalCash, totalUtangKotor, totalBelumLunas,
      supplierList: Object.values(perSupplier).filter((s) => s.sisa > 0 || s.amount > 0),
      notaList: [...gPurchases].sort((a, b) => (a.date < b.date ? 1 : -1)),
      count: gPurchases.length,
    };
  }).filter((x) => x.count > 0);

  const usageByMaterial = materials.map((m) => {
    const mUsage = usage.filter((u) => u.materialId === m.id);
    const totalQty = mUsage.reduce((s, u) => s + (Number(u.qty) || 0), 0);
    return { material: m, totalQty, count: mUsage.length };
  }).filter((x) => x.totalQty > 0).sort((a, b) => b.totalQty - a.totalQty);

  const grandUtang = perGudang.reduce((s, x) => s + x.totalBelumLunas, 0);
  const utangPekerja = buildUpahRows(workers, weeks, projects, payments, kasbon, kasbonPayments, null, 'all')
    .filter((r) => r.sisa > 0)
    .sort((a, b) => b.sisa - a.sisa);
  const grandUtangPekerja = utangPekerja.reduce((s, r) => s + r.sisa, 0);
  const grandCash = perGudang.reduce((s, x) => s + x.totalCash, 0);

  const stokGudang = (gudangId) => materials.map((m) => {
    const totalMasuk = purchases.filter((p) => p.gudangId === gudangId)
      .reduce((s, p) => s + purchaseItems(p).filter((it) => it.materialId === m.id).reduce((s2, it) => s2 + (Number(it.qty) || 0), 0), 0);
    const totalKeluar = usage.filter((u) => u.materialId === m.id && u.gudangId === gudangId).reduce((s, u) => s + (Number(u.qty) || 0), 0);
    return { material: m, totalMasuk, totalKeluar, sisaStok: totalMasuk - totalKeluar };
  }).filter((x) => x.totalMasuk > 0 || x.totalKeluar > 0);

  const submitUtangPayment = async () => {
    const amt = Number(payAmount.replace(/[^0-9]/g, '')) || 0;
    if (amt <= 0 || !payFor) return;
    await onAddUtangPayment({
      id: uid(), date: payDate, gudangId: payFor.gudangId,
      supplierId: payFor.supplierKey, supplierName: payFor.supplierName,
      amount: amt, method: payMethod, kasirName: payMethod === 'cash' ? payWho.trim() : '',
      bankName: payMethod === 'transfer' ? payWho.trim() : '', note: payNote.trim(),
    });
    setPayFor(null); setPayAmount(''); setPayWho(''); setPayNote(''); setPayMethod('cash'); setPayDate(new Date().toISOString().slice(0, 10));
  };

  const text = [
    'REKAP HUTANG/CASH PER PROYEK', '================================',
    ...perGudang.map((x) => [
      x.gudang.name,
      `  Cash: ${formatRupiah(x.totalCash)}`,
      `  Belum lunas: ${formatRupiah(x.totalBelumLunas)}`,
      ...x.supplierList.map((s) => `    - ${s.name}: sisa ${formatRupiah(s.sisa)} (total ${formatRupiah(s.amount)}, dibayar ${formatRupiah(s.dibayar)})`),
      '  Riwayat nota:',
      ...x.notaList.map((p) => `    ${p.date} Nota ${p.noNota || '-'} (${p.supplierName || 'tanpa suplier'}) - ${formatRupiah(purchaseTotal(p))}`),
    ].join('\n')),
    '================================',
    `TOTAL CASH: ${formatRupiah(grandCash)}`, `TOTAL BELUM LUNAS: ${formatRupiah(grandUtang)}`,
    '', 'UTANG UPAH KE PEKERJA', '================================',
    ...utangPekerja.map((r) => `${r.worker.name}: ${formatRupiah(r.sisa)}`),
    `TOTAL UTANG UPAH: ${formatRupiah(grandUtangPekerja)}`,
    '', 'REKAP PEMAKAIAN MATERIAL/PERALATAN', '================================',
    ...usageByMaterial.map((x) => `${x.material.name}: ${x.totalQty} ${x.material.unit || ''}`),
  ].join('\n');

  return (
    <div className="p-4 pb-24">
      <button type="button" onClick={() => setShowKelolaProyek(!showKelolaProyek)}
        className="w-full flex items-center justify-between px-3 py-2 rounded-lg mb-3 att-body text-sm font-semibold"
        style={{ background: THEME.charcoal, color: THEME.paper }}>
        Kelola Status Proyek ({projects.length})
        {showKelolaProyek ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
      </button>

      {showKelolaProyek && (
        <div className="space-y-2 mb-5">
          {projects.map((p) => {
            const status = p.status || 'berjalan';
            const overdue = p.targetDate && status !== 'selesai' && new Date(p.targetDate) < new Date(new Date().toDateString());
            const daysLeft = p.targetDate ? Math.ceil((new Date(p.targetDate) - new Date(new Date().toDateString())) / 86400000) : null;
            return (
              <div key={p.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${overdue ? THEME.rust : THEME.line}` }}>
                <div className="flex items-center justify-between mb-2">
                  <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{p.name}</p>
                  <span className="att-mono text-[9px] px-1.5 py-0.5 rounded-full font-semibold" style={{ background: PROJECT_STATUS_COLORS[status], color: '#fff' }}>
                    {PROJECT_STATUS_LABELS[status]}
                  </span>
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <select value={status} onChange={(e) => onUpdateProject(p.id, { status: e.target.value })}
                    className="px-2 py-1.5 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }}>
                    {Object.keys(PROJECT_STATUS_LABELS).map((s) => <option key={s} value={s}>{PROJECT_STATUS_LABELS[s]}</option>)}
                  </select>
                  <input type="date" value={p.targetDate || ''} onChange={(e) => onUpdateProject(p.id, { targetDate: e.target.value })}
                    className="px-2 py-1.5 rounded att-mono text-xs outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
                </div>
                {p.targetDate && status !== 'selesai' && (
                  <p className="att-mono text-[10px] mt-1.5" style={{ color: overdue ? THEME.rust : THEME.inkSoft }}>
                    {overdue ? `Terlambat ${Math.abs(daysLeft)} hari dari target selesai` : `${daysLeft} hari lagi menuju target selesai`}
                  </p>
                )}
                <div className="mt-2 space-y-1.5">
                  <div className="flex items-center gap-1.5">
                    <input type="text" defaultValue={p.cctv1 || ''} placeholder="Kamera 1 — mis. EZVIZ, Kamera Gerbang Depan"
                      onBlur={(e) => { if (e.target.value !== (p.cctv1 || '')) onUpdateProject(p.id, { cctv1: e.target.value.trim() }); }}
                      className="flex-1 px-2 py-1.5 rounded att-mono text-[11px] outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
                    <Video size={13} color={THEME.inkSoft} className="shrink-0" />
                  </div>
                  <div className="flex items-center gap-1.5">
                    <input type="text" defaultValue={p.cctv2 || ''} placeholder="Kamera 2 — mis. EZVIZ, Kamera Gudang Material"
                      onBlur={(e) => { if (e.target.value !== (p.cctv2 || '')) onUpdateProject(p.id, { cctv2: e.target.value.trim() }); }}
                      className="flex-1 px-2 py-1.5 rounded att-mono text-[11px] outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
                    <Video size={13} color={THEME.inkSoft} className="shrink-0" />
                  </div>
                  {(p.cctv1 || p.cctv2) && (
                    <p className="att-mono text-[9.5px]" style={{ color: THEME.inkSoft }}>
                      Buka lewat app kamera merek terkait di HP — kolom ini catatan pengingat, bukan link langsung.
                    </p>
                  )}
                </div>
              </div>
            );
          })}
          {projects.length === 0 && (
            <p className="att-body text-sm text-center py-4" style={{ color: THEME.inkSoft }}>Belum ada proyek. Buat lewat tab Absensi Mingguan.</p>
          )}
        </div>
      )}

      <div className="flex items-center justify-between mb-2 gap-2">
        <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>HUTANG &amp; CASH PER PROYEK</p>
        {(perGudang.length > 0 || usageByMaterial.length > 0) && (
          <div className="flex gap-1.5">
            <button type="button" onClick={() => setShowTablePreview(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded att-mono text-xs font-semibold" style={{ border: `1px solid ${THEME.amber}`, color: THEME.charcoal, background: THEME.amber }}>
              <FileBarChart size={13} /> Salin Tabel
            </button>
            <button type="button" onClick={() => exportRekapProyekExcel(perGudang, utangPekerja, usageByMaterial)}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded att-mono text-xs" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.paper }}>
              <Save size={13} /> Unduh .xlsx
            </button>
            <button type="button" onClick={() => exportRekapProyekPDF(perGudang, utangPekerja, usageByMaterial)}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded att-mono text-xs" style={{ border: `1px solid ${THEME.line}`, color: THEME.ink, background: THEME.paper }}>
              <Save size={13} /> Unduh .pdf
            </button>
          </div>
        )}
      </div>
      {(perGudang.length > 0 || usageByMaterial.length > 0) && (
        <div className="mb-3">
          <PrintMenu
            onPrint={(thermal) => openPrintDocument('Rekap Proyek', `<pre style="white-space:pre-wrap;font-family:inherit;">${escapeHtml(text)}</pre>`, thermal)}
            onCopy={() => setShowTextPreview(true)}
          />
          <p className="att-mono text-[10px] mt-1" style={{ color: THEME.inkSoft }}>
            "Salin Tabel" paling disarankan — hasilnya bisa langsung ditempel ke Excel/Google Sheets dan otomatis terbagi per kolom.
          </p>
        </div>
      )}
      {showTextPreview && <TextPreviewModal title="Rekap Proyek" text={text} onClose={() => setShowTextPreview(false)} />}
      {showTablePreview && (
        <TextPreviewModal
          title="Tabel Rekap Proyek (untuk Excel/Sheets)"
          text={buildRekapProyekTSV(perGudang, utangPekerja, usageByMaterial)}
          onClose={() => setShowTablePreview(false)}
        />
      )}

      <div className="space-y-2 mb-6">
        {perGudang.map(({ gudang, totalCash, totalBelumLunas, supplierList, notaList }) => (
          <div key={gudang.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <p className="att-body font-semibold text-sm mb-2" style={{ color: THEME.ink }}>{gudang.name}</p>
            <div className="grid grid-cols-2 gap-2 att-mono text-xs mb-2">
              <div className="p-2 rounded text-center" style={{ background: THEME.concrete }}>
                <p style={{ color: THEME.green }} className="font-semibold">{formatRupiah(totalCash)}</p>
                Cash
              </div>
              <div className="p-2 rounded text-center" style={{ background: THEME.concrete }}>
                <p style={{ color: THEME.rust }} className="font-semibold">{formatRupiah(totalBelumLunas)}</p>
                Belum Lunas
              </div>
            </div>

            {supplierList.length > 0 && (
              <div className="space-y-1.5 mb-2">
                {supplierList.map((s) => {
                  const histKey = `${gudang.id}:${s.key}`;
                  const history = utangPayments
                    .filter((up) => up.gudangId === gudang.id && (up.supplierId || up.supplierName || 'lainnya') === s.key)
                    .sort((a, b) => (a.date < b.date ? 1 : -1));
                  return (
                    <div key={s.key} className="p-2 rounded" style={{ background: THEME.concrete }}>
                      <div className="flex justify-between att-mono text-[11px] mb-1">
                        <span style={{ color: THEME.inkSoft }}>{s.name}</span>
                        <span style={{ color: s.sisa > 0 ? THEME.rust : THEME.green }} className="font-semibold">
                          {s.sisa > 0 ? `Sisa ${formatRupiah(s.sisa)}` : 'Lunas'}
                        </span>
                      </div>
                      {s.sisa > 0 && (
                        payFor && payFor.gudangId === gudang.id && payFor.supplierKey === s.key ? (
                          <div className="space-y-1.5">
                            <div className="flex gap-1.5">
                              <button type="button" onClick={() => setPayMethod('cash')} className="flex-1 py-1 rounded att-mono text-[10px] font-semibold"
                                style={{ background: payMethod === 'cash' ? THEME.amber : THEME.paper, color: payMethod === 'cash' ? THEME.charcoal : THEME.inkSoft }}>Cash</button>
                              <button type="button" onClick={() => setPayMethod('transfer')} className="flex-1 py-1 rounded att-mono text-[10px] font-semibold"
                                style={{ background: payMethod === 'transfer' ? THEME.amber : THEME.paper, color: payMethod === 'transfer' ? THEME.charcoal : THEME.inkSoft }}>Transfer Bank</button>
                            </div>
                            <div className="flex gap-1.5">
                              <input value={payAmount} onChange={(e) => setPayAmount(e.target.value.replace(/[^0-9]/g, ''))} placeholder="Jumlah bayar (Rp)" inputMode="numeric"
                                className="flex-1 px-2 py-1.5 rounded att-mono text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }} />
                              <input type="date" value={payDate} onChange={(e) => setPayDate(e.target.value)}
                                className="px-2 py-1.5 rounded att-mono text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }} />
                            </div>
                            <input value={payWho} onChange={(e) => setPayWho(e.target.value)} placeholder={payMethod === 'cash' ? 'Nama kasir (opsional)' : 'Nama bank (opsional)'}
                              className="w-full px-2 py-1.5 rounded att-body text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }} />
                            <input value={payNote} onChange={(e) => setPayNote(e.target.value)} placeholder="Catatan (opsional)"
                              className="w-full px-2 py-1.5 rounded att-body text-xs outline-none" style={{ background: THEME.paper, color: THEME.ink, border: `1px solid ${THEME.line}` }} />
                            <div className="flex gap-1.5">
                              <button type="button" onClick={submitUtangPayment} className="flex-1 py-1.5 rounded att-mono text-xs font-semibold" style={{ background: THEME.amber, color: THEME.charcoal }}>
                                Simpan
                              </button>
                              <button type="button" onClick={() => { setPayFor(null); setPayAmount(''); setPayWho(''); setPayNote(''); }} className="px-3 py-1.5 rounded att-mono text-xs" style={{ border: `1px solid ${THEME.line}`, color: THEME.inkSoft }}>
                                Batal
                              </button>
                            </div>
                          </div>
                        ) : (
                          <button type="button" onClick={() => { setPayFor({ gudangId: gudang.id, supplierKey: s.key, supplierName: s.name }); setPayAmount(String(s.sisa)); setPayDate(new Date().toISOString().slice(0, 10)); }}
                            className="w-full py-1 rounded att-mono text-[10px] font-semibold" style={{ background: THEME.amber, color: THEME.charcoal }}>
                            Bayar Utang
                          </button>
                        )
                      )}
                      {history.length > 0 && (
                        <button type="button" onClick={() => setOpenHistory(openHistory === histKey ? null : histKey)}
                          className="att-mono text-[10px] mt-1" style={{ color: THEME.inkSoft }}>
                          {openHistory === histKey ? 'Sembunyikan riwayat bayar' : `Riwayat bayar (${history.length})`}
                        </button>
                      )}
                      {openHistory === histKey && (
                        <div className="mt-1 space-y-1">
                          {history.map((h) => (
                            <div key={h.id} className="att-mono text-[10px] p-1.5 rounded" style={{ background: THEME.paper, color: THEME.inkSoft }}>
                              <div className="flex justify-between">
                                <span>{h.date} &middot; {h.method === 'cash' ? 'Cash' : 'Transfer'}{h.kasirName ? ` (${h.kasirName})` : ''}{h.bankName ? ` (${h.bankName})` : ''}</span>
                                <span style={{ color: THEME.green }} className="font-semibold">{formatRupiah(h.amount)}</span>
                              </div>
                              {h.note && <div>{h.note}</div>}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}

            <div className="flex gap-3 flex-wrap">
              <button type="button" onClick={() => setOpenNota(openNota === gudang.id ? null : gudang.id)} className="att-mono text-[11px]" style={{ color: THEME.inkSoft }}>
                {openNota === gudang.id ? 'Sembunyikan riwayat nota' : `Lihat riwayat nota (${notaList.length})`}
              </button>
              <button type="button" onClick={() => setOpenStok(openStok === gudang.id ? null : gudang.id)} className="att-mono text-[11px]" style={{ color: THEME.inkSoft }}>
                {openStok === gudang.id ? 'Sembunyikan stok' : 'Lihat stok gudang'}
              </button>
              <button type="button" onClick={() => setOpenPeralatanUsage(openPeralatanUsage === gudang.id ? null : gudang.id)} className="att-mono text-[11px]" style={{ color: THEME.inkSoft }}>
                {openPeralatanUsage === gudang.id ? 'Sembunyikan pemakaian peralatan' : `Lihat pemakaian peralatan (${peralatanUsage.filter((u) => u.gudangId === gudang.id).length})`}
              </button>
            </div>

            {openNota === gudang.id && (
              <div className="mt-2 space-y-1.5">
                {notaList.map((p) => (
                  <div key={p.id} className="p-2 rounded att-mono text-[11px]" style={{ background: THEME.concrete, color: THEME.inkSoft }}>
                    <div className="flex justify-between">
                      <span style={{ color: THEME.ink }} className="font-semibold">{p.date} &middot; {p.noNota || 'tanpa no. nota'}</span>
                      <span>{formatRupiah(purchaseTotal(p))}</span>
                    </div>
                    <div>{p.supplierName || '(tanpa suplier)'} &middot; {STATUS_BAYAR_LABEL[p.statusBayar] || '-'}</div>
                    {purchaseItems(p).map((it, idx) => (
                      <div key={it.id || idx}>&bull; [{KATEGORI_LABELS[it.kategori] || KATEGORI_LABELS.non_kategori}] {it.materialName}: {it.qty} {it.unit} &times; {formatRupiah(it.price)}</div>
                    ))}
                    {p.photo?.photo && (
                      <div className="flex items-center gap-1.5 mt-1.5">
                        <img src={p.photo.photo} alt="Foto nota" className="w-10 h-10 rounded object-cover" style={{ border: `1px solid ${THEME.line}` }} />
                        <span className="flex items-center gap-1"><Camera size={10} /> {p.photo.time}</span>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}

            {openStok === gudang.id && (
              <div className="mt-2 space-y-1.5">
                {stokGudang(gudang.id).map(({ material, totalMasuk, totalKeluar, sisaStok }) => (
                  <div key={material.id} className="p-2 rounded att-mono text-[11px] flex justify-between" style={{ background: THEME.concrete, color: THEME.inkSoft }}>
                    <span style={{ color: THEME.ink }}>{material.name}</span>
                    <span>Masuk {totalMasuk} &middot; Pakai {totalKeluar} &middot; <span style={{ color: sisaStok > 0 ? THEME.green : THEME.rust, fontWeight: 600 }}>Sisa {sisaStok}</span></span>
                  </div>
                ))}
                {stokGudang(gudang.id).length === 0 && <p className="att-body text-xs" style={{ color: THEME.inkSoft }}>Belum ada stok tercatat.</p>}
              </div>
            )}

            {openPeralatanUsage === gudang.id && (
              <div className="mt-2 space-y-1.5">
                {peralatanUsage.filter((u) => u.gudangId === gudang.id).sort((a, b) => (a.date < b.date ? 1 : -1)).map((u) => (
                  <div key={u.id} className="p-2 rounded att-mono text-[11px] flex justify-between" style={{ background: THEME.concrete, color: THEME.inkSoft }}>
                    <span style={{ color: THEME.ink }}>{u.date} &middot; {u.peralatanNama}</span>
                    <span>{u.workerName}{u.note ? ` - ${u.note}` : ''}</span>
                  </div>
                ))}
                {peralatanUsage.filter((u) => u.gudangId === gudang.id).length === 0 && <p className="att-body text-xs" style={{ color: THEME.inkSoft }}>Belum ada pemakaian peralatan tercatat.</p>}
              </div>
            )}
          </div>
        ))}
        {perGudang.length === 0 && (
          <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada data belanja untuk direkap.</p>
        )}
      </div>

      {perGudang.length > 0 && (
        <div className="mb-6 p-3 rounded-lg" style={{ background: THEME.charcoal }}>
          <div className="flex justify-between att-mono text-xs mb-1" style={{ color: '#9C948A' }}>
            <span>Total Cash</span><span>{formatRupiah(grandCash)}</span>
          </div>
          <div className="flex justify-between att-body text-sm font-semibold" style={{ color: THEME.amber }}>
            <span>Total Belum Lunas</span><span>{formatRupiah(grandUtang)}</span>
          </div>
        </div>
      )}

      <p className="att-mono text-xs mb-2" style={{ color: THEME.inkSoft }}>UTANG UPAH KE PEKERJA</p>
      <div className="space-y-2 mb-6">
        {utangPekerja.map((r) => (
          <div key={r.worker.id} className="p-3 rounded-lg flex items-center justify-between" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div>
              <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{r.worker.name}</p>
              <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>Total upah {formatRupiah(r.totalUpah)} &middot; Dibayar {formatRupiah(r.diterima)}</p>
            </div>
            <p className="att-mono text-sm font-bold" style={{ color: THEME.rust }}>{formatRupiah(r.sisa)}</p>
          </div>
        ))}
        {utangPekerja.length === 0 && (
          <p className="att-body text-sm text-center py-4" style={{ color: THEME.inkSoft }}>Tidak ada upah yang belum lunas.</p>
        )}
        {utangPekerja.length > 0 && (
          <div className="p-3 rounded-lg flex items-center justify-between" style={{ background: THEME.charcoal }}>
            <span className="att-body font-semibold text-sm" style={{ color: THEME.paper }}>TOTAL UTANG UPAH</span>
            <span className="att-mono text-sm font-bold" style={{ color: THEME.amber }}>{formatRupiah(grandUtangPekerja)}</span>
          </div>
        )}
        <p className="att-mono text-[10px]" style={{ color: THEME.inkSoft }}>Untuk membayar, buka tab Rekap Upah &rarr; Tambah Pembayaran pada pekerja terkait.</p>
      </div>

      <p className="att-mono text-xs mb-2" style={{ color: THEME.inkSoft }}>REKAP PEMAKAIAN MATERIAL/PERALATAN</p>
      <div className="space-y-2">
        {usageByMaterial.map(({ material, totalQty }) => (
          <div key={material.id} className="p-3 rounded-lg flex items-center justify-between" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div>
              <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{material.name}</p>
              <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{KATEGORI_LABELS[material.kategori] || KATEGORI_LABELS.non_kategori}</p>
            </div>
            <p className="att-mono text-sm font-bold" style={{ color: THEME.ink }}>{totalQty} {material.unit}</p>
          </div>
        ))}
        {usageByMaterial.length === 0 && (
          <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada catatan pemakaian.</p>
        )}
      </div>
    </div>
  );
}

/* ---------------- Dashboard Ringkasan ---------------- */

function DashboardTab({ workers, projects, weeks, payments, kasbon, kasbonPayments, materials, purchases, usage, utangPayments, setTab }) {
  const now = new Date();
  const curMonth = now.getMonth();
  const curYear = now.getFullYear();

  const upahBulanIni = weeks
    .filter((w) => {
      if (!w.startDate) return false;
      const d = new Date(w.startDate);
      return d.getMonth() === curMonth && d.getFullYear() === curYear;
    })
    .reduce((sum, w) => {
      const weekWorkers = w.workerIds && w.workerIds.length > 0 ? workers.filter((wk) => w.workerIds.includes(wk.id)) : workers;
      const dk = getWeekDates(w).map((d) => d.key);
      return sum + weekWorkers.reduce((s, wk) => s + calcWorkerWeekTotals(wk, w.records[wk.id], dk).totalUpah, 0);
    }, 0);

  const totalKasbonBelumLunas = workers.reduce((sum, w) => {
    const diberi = kasbon.filter((k) => k.workerId === w.id).reduce((s, k) => s + (Number(k.amount) || 0), 0);
    const dibayar = kasbonPayments.filter((p) => p.workerId === w.id).reduce((s, p) => s + (Number(p.amount) || 0), 0);
    return sum + Math.max(0, diberi - dibayar);
  }, 0);

  const gudangList = getGudangList(projects);
  const totalUtangKotor = purchases.filter((p) => p.statusBayar && p.statusBayar !== 'cash').reduce((s, p) => s + purchaseTotal(p), 0);
  const totalUtangDibayar = utangPayments.reduce((s, up) => s + (Number(up.amount) || 0), 0);
  const totalUtangBelumLunas = Math.max(0, totalUtangKotor - totalUtangDibayar);

  const stokMenipis = materials
    .filter((m) => Number(m.stokMinimum) > 0)
    .map((m) => {
      const totalMasuk = purchases.reduce((s, p) => s + purchaseItems(p).filter((it) => it.materialId === m.id).reduce((s2, it) => s2 + (Number(it.qty) || 0), 0), 0);
      const totalKeluar = usage.filter((u) => u.materialId === m.id).reduce((s, u) => s + (Number(u.qty) || 0), 0);
      const sisa = totalMasuk - totalKeluar;
      return { material: m, sisa };
    })
    .filter((x) => x.sisa <= Number(x.material.stokMinimum));

  const upahBelumLunas = buildUpahRows(workers, weeks, projects, payments, kasbon, kasbonPayments, null, 'all').filter((r) => r.sisa > 0);

  const today = new Date(new Date().toDateString());
  const proyekTerlambat = projects.filter((p) => p.targetDate && (p.status || 'berjalan') !== 'selesai' && new Date(p.targetDate) < today);
  const proyekBerjalan = projects.filter((p) => (p.status || 'berjalan') === 'berjalan').length;

  const prioritas = [
    ...stokMenipis.map((x) => ({ type: 'stok', key: `stok-${x.material.id}`, text: `Stok ${x.material.name} tersisa ${x.sisa} ${x.material.unit || ''}`, tab: 'gudang' })),
    ...upahBelumLunas.map((r) => ({ type: 'upah', key: `upah-${r.worker.id}`, text: `Upah ${r.worker.name} belum lunas ${formatRupiah(r.sisa)}`, tab: 'rekap' })),
    ...(totalUtangBelumLunas > 0 ? [{ type: 'utang', key: 'utang-total', text: `Ada utang ke suplier belum lunas ${formatRupiah(totalUtangBelumLunas)}`, tab: 'proyek' }] : []),
    ...proyekTerlambat.map((p) => ({ type: 'proyek', key: `proyek-${p.id}`, text: `Proyek ${p.name} sudah lewat target selesai`, tab: 'proyek' })),
  ];

  const cards = [
    { label: 'Pekerja Aktif', value: workers.length, tab: 'pekerja' },
    { label: 'Proyek Berjalan', value: proyekBerjalan, tab: 'absensi' },
    { label: 'Upah Bulan Ini', value: formatRupiah(upahBulanIni), tab: 'rekap' },
    { label: 'Kasbon Belum Lunas', value: formatRupiah(totalKasbonBelumLunas), tab: 'kasbon' },
    { label: 'Utang ke Suplier', value: formatRupiah(totalUtangBelumLunas), tab: 'proyek' },
  ];

  const proyekDenganCctv = projects.filter((p) => p.cctv1 || p.cctv2);

  return (
    <div className="p-4 pb-24">
      <p className="att-mono text-xs mb-2" style={{ color: THEME.inkSoft }}>
        RINGKASAN &middot; {now.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}
      </p>
      {proyekDenganCctv.length > 0 && (
        <div className="mb-5">
          <p className="att-mono text-xs mb-2" style={{ color: THEME.inkSoft }}>KAMERA CCTV PROYEK</p>
          <div className="space-y-1.5">
            {proyekDenganCctv.map((p) => (
              <div key={p.id} className="p-2.5 rounded-lg" style={{ background: THEME.charcoal }}>
                <p className="att-body text-sm font-semibold mb-1" style={{ color: THEME.paper }}>{p.name}</p>
                {p.cctv1 && (
                  <p className="att-mono text-[11px] flex items-center gap-1.5" style={{ color: THEME.amberSoft }}>
                    <Video size={11} /> {p.cctv1}
                  </p>
                )}
                {p.cctv2 && (
                  <p className="att-mono text-[11px] flex items-center gap-1.5 mt-0.5" style={{ color: THEME.amberSoft }}>
                    <Video size={11} /> {p.cctv2}
                  </p>
                )}
              </div>
            ))}
          </div>
        </div>
      )}
      <div className="grid grid-cols-2 gap-2 mb-5">
        {cards.map((c) => (
          <button key={c.label} type="button" onClick={() => setTab(c.tab)}
            className="p-3 rounded-lg text-left" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <p className="att-mono text-[10px]" style={{ color: THEME.inkSoft }}>{c.label}</p>
            <p className="att-body font-bold text-base mt-0.5" style={{ color: THEME.ink }}>{c.value}</p>
          </button>
        ))}
      </div>

      <p className="att-mono text-xs mb-2" style={{ color: THEME.inkSoft }}>PRIORITAS HARI INI ({prioritas.length})</p>
      <div className="space-y-2 mb-5">
        {prioritas.map((p) => (
          <button key={p.key} type="button" onClick={() => setTab(p.tab)}
            className="w-full p-3 rounded-lg flex items-center gap-2 text-left" style={{ background: '#3A241D', border: `1px solid ${THEME.rust}` }}>
            <AlertTriangle size={16} color={THEME.rust} className="shrink-0" />
            <span className="att-body text-sm" style={{ color: '#E8A08F' }}>{p.text}</span>
          </button>
        ))}
        {prioritas.length === 0 && (
          <p className="att-body text-sm text-center py-4" style={{ color: THEME.inkSoft }}>Semua aman, tidak ada yang perlu ditindaklanjuti hari ini. 👍</p>
        )}
      </div>
    </div>
  );
}

/* ---------------- App Root ---------------- */

function stripUndefined(obj) {
  const out = {};
  Object.entries(obj).forEach(([k, v]) => {
    if (v !== undefined) out[k] = v;
  });
  return out;
}

async function fetchAll(collectionName) {
  const snap = await getDocs(collection(db, collectionName));
  return snap.docs.map((d) => ({ id: d.id, ...d.data() }));
}

export default function App() {
  const [loading, setLoading] = useState(false);
  const [persistenceOk, setPersistenceOk] = useState(true);
  const [currentUser, setCurrentUser] = useState(null);
  const [loginError, setLoginError] = useState('');
  const [loginBusy, setLoginBusy] = useState(false);

  const [workers, setWorkers] = useState([]);
  const [weeks, setWeeks] = useState([]);
  const [projects, setProjects] = useState([]);
  const [payments, setPayments] = useState([]);
  const [evidence, setEvidence] = useState({});
  const [kasbon, setKasbon] = useState([]);
  const [kasbonPayments, setKasbonPayments] = useState([]);
  const [materials, setMaterials] = useState([]);
  const [suppliers, setSuppliers] = useState([]);
  const [purchases, setPurchases] = useState([]);
  const [usage, setUsage] = useState([]);
  const [utangPayments, setUtangPayments] = useState([]);
  const [peralatan, setPeralatan] = useState([]);
  const [peralatanUsage, setPeralatanUsage] = useState([]);
  const [progressPhotos, setProgressPhotos] = useState([]);
  const [penawaranList, setPenawaranList] = useState([]);
  const [ahspList, setAhspList] = useState([]);
  const [klienList, setKlienList] = useState([]);
  const [clients, setClients] = useState([]);
  const [tab, setTab] = useState('dashboard');
  const [showBackup, setShowBackup] = useState(false);
  const [showProfile, setShowProfile] = useState(false);
  const [showManageUsers, setShowManageUsers] = useState(false);
  const [appUsers, setAppUsers] = useState([]);
  const [company, setCompany] = useState({ name: '', tagline: '', address: '', phone: '' });

  useEffect(() => {
    CURRENT_COMPANY = company;
  }, [company]);

  // Pulihkan sesi login dari localStorage (bertahan lintas refresh/tab)
  useEffect(() => {
    try {
      const saved = localStorage.getItem('tukang_session');
      if (saved) setCurrentUser(JSON.parse(saved));
    } catch {
      // abaikan sesi tersimpan yang rusak
    }
  }, []);

  // Bootstrap: buat user admin default kalau koleksi users masih kosong.
  // Ini WAJIB jalan sebelum login (bukan sesudahnya), karena login pertama
  // butuh setidaknya satu user sudah ada di Firestore.
  useEffect(() => {
    (async () => {
      try {
        const rows = await fetchAll('users');
        if (rows.length === 0) {
          const hash = await bcrypt.hash('admin123', 10);
          await setDoc(doc(db, 'users', uid()), { username: 'admin', passwordHash: hash });
        }
      } catch (err) {
        console.error('Gagal bootstrap user admin default:', err.message);
      }
    })();
  }, []);

  const loadAppUsers = useCallback(async () => {
    try {
      const rows = await fetchAll('users');
      // Buat admin default kalau belum ada user sama sekali (bootstrap sekali di awal)
      if (rows.length === 0) {
        const hash = await bcrypt.hash('admin123', 10);
        await setDoc(doc(db, 'users', uid()), { username: 'admin', passwordHash: hash });
        const rows2 = await fetchAll('users');
        setAppUsers(rows2.map((r) => ({ id: r.id, username: r.username, created_at: r.created_at })));
        return;
      }
      setAppUsers(rows.map((r) => ({ id: r.id, username: r.username, created_at: r.created_at })));
    } catch (err) {
      console.error('Gagal memuat daftar user:', err.message);
    }
  }, []);

  const loadAll = useCallback(async () => {
    setLoading(true);
    let ok = true;
    try {
      const [w, wk, pr, pay, ev, ka, kp, mat, sup, pu, us, up, pe, peu, pgf, pnw, ahsp, klien, compSnap] = await Promise.all([
        fetchAll('workers'), fetchAll('weeks'), fetchAll('projects'), fetchAll('payments'),
        fetchAll('evidence'), fetchAll('kasbon'), fetchAll('kasbonPayments'), fetchAll('materials'),
        fetchAll('suppliers'), fetchAll('purchases'), fetchAll('usage'), fetchAll('utangPayments'),
        fetchAll('peralatan'), fetchAll('peralatanUsage'), fetchAll('progressPhotos'), fetchAll('penawaran'), fetchAll('ahspMaster'), fetchAll('klien'),
        getDocs(collection(db, 'companyProfile')),
      ]);
      setWorkers(w);
      setWeeks(wk.map((x) => ({ ...x, workerIds: x.workerIds || [], records: x.records || {} })));
      setProjects(pr);
      setPayments(pay);
      const evMap = {};
      ev.forEach((row) => { evMap[row.id] = row.data || {}; });
      setEvidence(evMap);
      setKasbon(ka);
      setKasbonPayments(kp);
      setMaterials(mat);
      setSuppliers(sup);
      setPurchases(pu.map((x) => ({ ...x, items: x.items || [] })));
      setUsage(us);
      setUtangPayments(up);
      setPeralatan(pe);
      setPeralatanUsage(peu);
      setProgressPhotos(pgf);
      setPenawaranList(pnw);
      setAhspList(ahsp);
      setKlienList(klien);
      const compDoc = compSnap.docs.find((d) => d.id === 'main');
      const compRow = compDoc ? compDoc.data() : { name: '', tagline: '', address: '', phone: '' };
      setCompany(compRow);
      CURRENT_COMPANY = compRow;
      await loadAppUsers();
    } catch (err) {
      console.error('Gagal memuat data dari server:', err.message);
      ok = false;
    }
    setPersistenceOk(ok);
    setLoading(false);
  }, [loadAppUsers]);

  useEffect(() => {
    if (currentUser) loadAll();
  }, [currentUser, loadAll]);

  // ---- Helper CRUD generik: update state lokal langsung + kirim ke Firestore ----
  function makeCrud(collectionName, setState) {
    return {
      add: async (item) => {
        setState((prev) => [item, ...prev]);
        try {
          await setDoc(doc(db, collectionName, item.id), stripUndefined(item));
        } catch (err) {
          console.error(`Gagal menambah data (${collectionName}):`, err.message);
          setPersistenceOk(false);
        }
      },
      update: async (id, patch) => {
        setState((prev) => prev.map((x) => (x.id === id ? { ...x, ...patch } : x)));
        try {
          await updateDoc(doc(db, collectionName, id), stripUndefined(patch));
        } catch (err) {
          console.error(`Gagal mengubah data (${collectionName}):`, err.message);
          setPersistenceOk(false);
        }
      },
      remove: async (id) => {
        setState((prev) => prev.filter((x) => x.id !== id));
        try {
          await deleteDoc(doc(db, collectionName, id));
        } catch (err) {
          console.error(`Gagal menghapus data (${collectionName}):`, err.message);
          setPersistenceOk(false);
        }
      },
    };
  }

  const workerCrud = makeCrud('workers', setWorkers);
  const projectCrud = makeCrud('projects', setProjects);
  const weekCrud = makeCrud('weeks', setWeeks);
  const paymentCrud = makeCrud('payments', setPayments);
  const kasbonCrud = makeCrud('kasbon', setKasbon);
  const kasbonPaymentCrud = makeCrud('kasbonPayments', setKasbonPayments);
  const materialCrud = makeCrud('materials', setMaterials);
  const supplierCrud = makeCrud('suppliers', setSuppliers);
  const purchaseCrud = makeCrud('purchases', setPurchases);
  const usageCrud = makeCrud('usage', setUsage);
  const utangPaymentCrud = makeCrud('utangPayments', setUtangPayments);
  const peralatanCrud = makeCrud('peralatan', setPeralatan);
  const peralatanUsageCrud = makeCrud('peralatanUsage', setPeralatanUsage);
  const progressPhotoCrud = makeCrud('progressPhotos', setProgressPhotos);
  const penawaranCrud = makeCrud('penawaran', setPenawaranList);
  const ahspCrud = makeCrud('ahspMaster', setAhspList);
  const klienCrud = makeCrud('klien', setKlienList);

  const handleLogin = async (username, password) => {
    setLoginError('');
    setLoginBusy(true);
    try {
      const rows = await fetchAll('users');
      const match = rows.find((r) => r.username === username);
      if (match && await bcrypt.compare(password, match.passwordHash)) {
        const user = { id: match.id, username: match.username };
        setCurrentUser(user);
        localStorage.setItem('tukang_session', JSON.stringify(user));
      } else {
        setLoginError('Username atau password salah.');
      }
    } catch (err) {
      console.error(err);
      setLoginError('Gagal terhubung ke server. Periksa koneksi internet atau pengaturan Firebase.');
    }
    setLoginBusy(false);
  };

  const handleLogout = () => {
    setCurrentUser(null);
    localStorage.removeItem('tukang_session');
  };

  const handleAddAppUser = async (username, password) => {
    try {
      const rows = await fetchAll('users');
      if (rows.some((r) => r.username === username)) return 'Username sudah dipakai.';
      const hash = await bcrypt.hash(password, 10);
      await setDoc(doc(db, 'users', uid()), { username, passwordHash: hash });
      await loadAppUsers();
      return null;
    } catch (err) {
      return 'Gagal menambah user: ' + err.message;
    }
  };
  const handleSetAppUserPassword = async (userId, newPassword) => {
    try {
      const hash = await bcrypt.hash(newPassword, 10);
      await updateDoc(doc(db, 'users', userId), { passwordHash: hash });
    } catch (err) {
      console.error('Gagal ganti password:', err.message);
    }
  };
  const handleDeleteAppUser = async (userId) => {
    try {
      await deleteDoc(doc(db, 'users', userId));
      await loadAppUsers();
    } catch (err) {
      console.error('Gagal hapus user:', err.message);
    }
  };

  const handleAddWorker = workerCrud.add;
  const handleUpdateWorker = workerCrud.update;
  const handleDeleteWorker = workerCrud.remove;

  const handleAddProject = projectCrud.add;
  const handleUpdateProject = projectCrud.update;

  const handleCreateWeek = weekCrud.add;
  const handleUpdateWeek = weekCrud.update;
  const handleUpdateWeekRecord = async (weekId, workerId, day, patch) => {
    const week = weeks.find((w) => w.id === weekId);
    if (!week) return;
    const workerRec = { ...(week.records[workerId] || {}) };
    const dayRec = { hadir: false, lembur: 0, ...(workerRec[day] || {}), ...patch };
    const nextRecords = { ...week.records, [workerId]: { ...workerRec, [day]: dayRec } };
    await weekCrud.update(weekId, { records: nextRecords });
  };

  const handleAddPayment = paymentCrud.add;
  const handleDeletePayment = paymentCrud.remove;

  const handleSaveEvidence = async (weekId, workerId, day, data) => {
    const mapKey = `${weekId}:${workerId}`;
    const nextDayMap = { ...(evidence[mapKey] || {}), [day]: data };
    setEvidence((prev) => ({ ...prev, [mapKey]: nextDayMap }));
    try {
      await setDoc(doc(db, 'evidence', mapKey), { weekId, workerId, data: nextDayMap });
    } catch (err) {
      console.error('Gagal menyimpan bukti kehadiran:', err.message);
      setPersistenceOk(false);
    }
  };

  const handleAddKasbon = kasbonCrud.add;
  const handleDeleteKasbon = kasbonCrud.remove;
  const handleAddKasbonPayment = kasbonPaymentCrud.add;
  const handleDeleteKasbonPayment = kasbonPaymentCrud.remove;

  const handleAddMaterial = materialCrud.add;
  const handleUpdateMaterial = materialCrud.update;
  const handleDeleteMaterial = materialCrud.remove;

  const handleAddSupplier = supplierCrud.add;
  const handleUpdateSupplier = supplierCrud.update;
  const handleDeleteSupplier = supplierCrud.remove;

  const handleAddPurchase = async (p) => {
    await purchaseCrud.add(p);
    const peralatanItems = purchaseItems(p).filter((it) => it.kategori === 'peralatan');
    if (peralatanItems.length > 0) {
      const findMaterial = (it) => materials.find((mm) => mm.id === it.materialId);
      for (const it of peralatanItems) {
        const newEntry = {
          id: uid(),
          nama: it.materialName,
          merk: findMaterial(it)?.merk || '',
          kapasitas: '',
          unit: it.unit || 'unit',
          jumlah: Number(it.qty) || 0,
          waktuPembelian: p.date,
          sourcePurchaseId: p.id,
          sourceNoNota: p.noNota || '',
          gudangId: p.gudangId || '',
        };
        await peralatanCrud.add(newEntry);
      }
    }
  };
  const handleDeletePurchase = purchaseCrud.remove;

  const handleAddPeralatan = peralatanCrud.add;
  const handleUpdatePeralatan = peralatanCrud.update;
  const handleDeletePeralatan = peralatanCrud.remove;

  const handleAddPeralatanUsage = peralatanUsageCrud.add;
  const handleDeletePeralatanUsage = peralatanUsageCrud.remove;
  const handleAddProgressPhoto = progressPhotoCrud.add;
  const handleDeleteProgressPhoto = progressPhotoCrud.remove;
  const handleAddPenawaran = penawaranCrud.add;
  const handleUpdatePenawaran = penawaranCrud.update;
  const handleDeletePenawaran = penawaranCrud.remove;
  const handleAddAhsp = ahspCrud.add;
  const handleUpdateAhsp = ahspCrud.update;
  const handleDeleteAhsp = ahspCrud.remove;
  const handleAddKlien = klienCrud.add;
  const handleUpdateKlien = klienCrud.update;
  const handleDeleteKlien = klienCrud.remove;

  const handleAddUsage = usageCrud.add;
  const handleDeleteUsage = usageCrud.remove;

  const handleAddUtangPayment = utangPaymentCrud.add;

  const saveCompany = async (next) => {
    setCompany(next);
    CURRENT_COMPANY = next;
    try {
      await setDoc(doc(db, 'companyProfile', 'main'), stripUndefined(next));
    } catch (err) {
      console.error('Gagal menyimpan profil perusahaan:', err.message);
      setPersistenceOk(false);
    }
  };

  // Backup/restore teks tetap tersedia sebagai cadangan darurat (mis. mau pindah server),
  // meski di versi web ini data utamanya sudah otomatis tersimpan permanen di Firestore.
  const bulkUpsert = async (collectionName, rows) => {
    if (!Array.isArray(rows) || rows.length === 0) return;
    try {
      for (const r of rows) {
        await setDoc(doc(db, collectionName, r.id), stripUndefined(r));
      }
    } catch (err) {
      console.error(`Gagal impor ${collectionName}:`, err.message);
    }
  };
  const handleImportBackup = async (parsed) => {
    if (Array.isArray(parsed.workers)) await bulkUpsert('workers', parsed.workers);
    if (Array.isArray(parsed.projects)) await bulkUpsert('projects', parsed.projects);
    if (Array.isArray(parsed.weeks)) await bulkUpsert('weeks', parsed.weeks);
    if (Array.isArray(parsed.payments)) await bulkUpsert('payments', parsed.payments);
    if (parsed.evidence && typeof parsed.evidence === 'object') {
      const rows = Object.entries(parsed.evidence).map(([id, data]) => {
        const [weekId, workerId] = id.split(':');
        return { id, weekId, workerId, data };
      });
      await bulkUpsert('evidence', rows);
    }
    if (Array.isArray(parsed.kasbon)) await bulkUpsert('kasbon', parsed.kasbon);
    if (Array.isArray(parsed.kasbonPayments)) await bulkUpsert('kasbonPayments', parsed.kasbonPayments);
    if (Array.isArray(parsed.materials)) await bulkUpsert('materials', parsed.materials);
    if (Array.isArray(parsed.suppliers)) await bulkUpsert('suppliers', parsed.suppliers);
    if (Array.isArray(parsed.purchases)) await bulkUpsert('purchases', parsed.purchases);
    if (Array.isArray(parsed.usage)) await bulkUpsert('usage', parsed.usage);
    if (Array.isArray(parsed.utangPayments)) await bulkUpsert('utangPayments', parsed.utangPayments);
    if (Array.isArray(parsed.peralatan)) await bulkUpsert('peralatan', parsed.peralatan);
    if (Array.isArray(parsed.peralatanUsage)) await bulkUpsert('peralatanUsage', parsed.peralatanUsage);
    if (parsed.company && typeof parsed.company === 'object') await saveCompany(parsed.company);
    await loadAll();
  };

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: THEME.charcoal }}>
        <GlobalStyle />
        <Loader2 size={28} className="animate-spin" color={THEME.amber} />
      </div>
    );
  }

  const persistenceBanner = !persistenceOk && (
    <div className="flex items-center gap-2 px-4 py-2 att-mono text-xs flex-wrap" style={{ background: '#3A241D', color: '#E8A08F' }}>
      <AlertTriangle size={14} className="shrink-0" />
      <span className="flex-1 min-w-[200px]">
        Gagal terhubung ke database Firestore. Periksa koneksi internet dan pengaturan VITE_FIREBASE_URL/VITE_FIREBASE_ANON_KEY di file .env.
      </span>
      <button type="button" onClick={loadAll} className="px-2 py-1 rounded font-semibold shrink-0" style={{ background: '#E8A08F', color: THEME.charcoal }}>
        Coba Lagi
      </button>
    </div>
  );

  if (!currentUser) {
    return (
      <div>
        {persistenceBanner}
        <LoginScreen onLogin={handleLogin} error={loginError} busy={loginBusy} />
      </div>
    );
  }

  return (
    <div className="min-h-screen att-body" style={{ background: THEME.concrete }}>
      <GlobalStyle />
      <div className="screen-only">
        {persistenceBanner}
        <TopBar
          onLogout={handleLogout}
          onBackup={() => setShowBackup(true)}
          onProfile={() => setShowProfile(true)}
          onManageUsers={() => setShowManageUsers(true)}
          company={company}
          username={currentUser?.username}
        />
        <TabBar tab={tab} setTab={setTab} />
      </div>

      {showProfile && (
        <CompanyProfileModal company={company} onSave={saveCompany} onClose={() => setShowProfile(false)} />
      )}

      {showManageUsers && (
        <ManageUsersModal
          users={appUsers}
          currentUsername={currentUser?.username}
          onAdd={handleAddAppUser}
          onSetPassword={handleSetAppUserPassword}
          onDelete={handleDeleteAppUser}
          onClose={() => setShowManageUsers(false)}
        />
      )}

      {showBackup && (
        <BackupModal
          data={{ workers, weeks, projects, payments, evidence, kasbon, kasbonPayments, materials, suppliers, purchases, usage, utangPayments, company, peralatan, peralatanUsage }}
          onImport={handleImportBackup}
          onClose={() => setShowBackup(false)}
        />
      )}

      {tab === 'dashboard' && (
        <DashboardTab
          workers={workers}
          projects={projects}
          weeks={weeks}
          payments={payments}
          kasbon={kasbon}
          kasbonPayments={kasbonPayments}
          materials={materials}
          purchases={purchases}
          usage={usage}
          utangPayments={utangPayments}
          setTab={setTab}
        />
      )}
      {tab === 'pekerja' && (
        <DataPekerjaTab workers={workers} onAdd={handleAddWorker} onUpdate={handleUpdateWorker} onDelete={handleDeleteWorker} />
      )}
      {tab === 'absensi' && (
        <AbsensiMingguanTab
          workers={workers}
          weeks={weeks}
          projects={projects}
          onCreateWeek={handleCreateWeek}
          onUpdateWeek={handleUpdateWeek}
          onAddProject={handleAddProject}
          onUpdateWeekRecord={handleUpdateWeekRecord}
          evidence={evidence}
          onSaveEvidence={handleSaveEvidence}
        />
      )}
      {tab === 'kasbon' && (
        <KasbonTab
          workers={workers}
          kasbon={kasbon}
          kasbonPayments={kasbonPayments}
          onAdd={handleAddKasbon}
          onDelete={handleDeleteKasbon}
          onAddKasbonPayment={handleAddKasbonPayment}
          onDeleteKasbonPayment={handleDeleteKasbonPayment}
        />
      )}
      {tab === 'belanja' && (
        <BelanjaTab
          materials={materials}
          onAddMaterial={handleAddMaterial}
          onUpdateMaterial={handleUpdateMaterial}
          onDeleteMaterial={handleDeleteMaterial}
          suppliers={suppliers}
          onAddSupplier={handleAddSupplier}
          onUpdateSupplier={handleUpdateSupplier}
          onDeleteSupplier={handleDeleteSupplier}
          projects={projects}
          purchases={purchases}
          onAddPurchase={handleAddPurchase}
          onDeletePurchase={handleDeletePurchase}
        />
      )}
      {tab === 'gudang' && (
        <GudangTab
          materials={materials}
          projects={projects}
          purchases={purchases}
          usage={usage}
          onAddUsage={handleAddUsage}
          onDeleteUsage={handleDeleteUsage}
          peralatan={peralatan}
          onAddPeralatan={handleAddPeralatan}
          onUpdatePeralatan={handleUpdatePeralatan}
          onDeletePeralatan={handleDeletePeralatan}
          peralatanUsage={peralatanUsage}
          onAddPeralatanUsage={handleAddPeralatanUsage}
          onDeletePeralatanUsage={handleDeletePeralatanUsage}
          workers={workers}
        />
      )}
      {tab === 'rekap' && (
        <RekapUpahTab
          workers={workers}
          weeks={weeks}
          projects={projects}
          payments={payments}
          kasbon={kasbon}
          kasbonPayments={kasbonPayments}
          evidence={evidence}
          onAddPayment={handleAddPayment}
          onDeletePayment={handleDeletePayment}
        />
      )}
      {tab === 'proyek' && (
        <RekapProyekTab
          projects={projects}
          purchases={purchases}
          usage={usage}
          materials={materials}
          utangPayments={utangPayments}
          onAddUtangPayment={handleAddUtangPayment}
          workers={workers}
          weeks={weeks}
          payments={payments}
          kasbon={kasbon}
          kasbonPayments={kasbonPayments}
          onUpdateProject={handleUpdateProject}
          peralatanUsage={peralatanUsage}
        />
      )}
      {tab === 'progres' && (
        <ProgresFotoTab
          projects={projects}
          entries={progressPhotos}
          onAdd={handleAddProgressPhoto}
          onDelete={handleDeleteProgressPhoto}
          penawaranList={penawaranList}
        />
      )}
      {tab === 'penawaran' && (
        <PenawaranTab
          penawaranList={penawaranList}
          onAdd={handleAddPenawaran}
          onUpdate={handleUpdatePenawaran}
          onDelete={handleDeletePenawaran}
          ahspList={ahspList}
          overheadProfit={Number(company.ahspOverheadProfit ?? 15)}
          projects={projects}
        />
      )}
      {tab === 'ahsp' && (
        <AhspMasterTab
          ahspList={ahspList}
          onAdd={handleAddAhsp}
          onUpdate={handleUpdateAhsp}
          onDelete={handleDeleteAhsp}
          company={company}
          onSaveCompany={saveCompany}
        />
      )}
      {tab === 'klien' && (
        <DataKlienTab
          klienList={klienList}
          onAdd={handleAddKlien}
          onUpdate={handleUpdateKlien}
          onDelete={handleDeleteKlien}
        />
      )}
    </div>
  );
}
