import React, { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { db } from './firebaseClient.js';
import {
  collection, doc, getDocs, setDoc, updateDoc, deleteDoc, addDoc, query,
} from 'firebase/firestore';
import bcrypt from 'bcryptjs';
import {
  LogIn, Loader2, AlertTriangle, Users, Plus, Trash2, Wallet,
  Camera, MapPin, ChevronDown, ChevronUp, Check, X, CalendarDays, Edit3, Printer, Receipt, Copy, Save, Banknote, ShoppingCart, Warehouse, FileBarChart, LayoutDashboard, Bluetooth
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
    <img src="${COMPANY_LOGO_DATA_URI}" alt="Logo" style="width:56px;height:56px;object-fit:cover;border-radius:6px;margin:0 auto 4px;display:block;" />
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

function resizeImage(file, maxWidth = 480, quality = 0.7) {
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
  const { worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown } = row;
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
  const { worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown } = row;
  const workRows = weekBreakdown.map((wk) => `<tr><td>${escapeHtml(wk.weekLabel)}</td><td>${escapeHtml(wk.projectName)}</td><td>${escapeHtml(wk.startDate)} s/d ${escapeHtml(wk.endDate)}</td><td>${wk.totalHariBayar}</td><td>${wk.totalJamLembur}</td><td>${formatRupiah(wk.totalUpah)}</td></tr>`).join('');
  return `<h2>SLIP GAJI</h2>
    <p><b>${escapeHtml(worker.name)}</b> &nbsp; ${escapeHtml(worker.position || '-')}</p>
    <p>Periode: ${escapeHtml(filterLabel)} &nbsp; Tanggal cetak: ${new Date().toISOString().slice(0, 10)}</p>
    <table><thead><tr><th>Periode</th><th>Proyek</th><th>Tanggal</th><th>Hari</th><th>Jam Lembur</th><th>Upah</th></tr></thead>
    <tbody>${workRows}</tbody>
    <tfoot><tr class="bold"><td colspan="3">TOTAL</td><td>${totalHari}</td><td>${totalJamLembur}</td><td>${formatRupiah(totalUpah)}</td></tr></tfoot></table>
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

  const handleSave = async () => {
    await onSave({ name: name.trim(), tagline: tagline.trim(), address: address.trim(), phone: phone.trim() });
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(33,29,26,0.7)' }}>
      <div className="w-full max-w-sm rounded-xl p-4 att-body att-punch-anim" style={{ background: THEME.paper }}>
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>Profil Perusahaan</h3>
          <button type="button" onClick={onClose}><X size={18} color={THEME.inkSoft} /></button>
        </div>
        <p className="att-mono text-[10px] mb-3" style={{ color: THEME.inkSoft }}>
          Muncul di judul aplikasi dan sebagai kop di setiap dokumen cetak/struk/salin teks.
        </p>
        <div className="space-y-2">
          <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Nama perusahaan"
            className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={tagline} onChange={(e) => setTagline(e.target.value)} placeholder="Tagline / jenis usaha (opsional)"
            className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={address} onChange={(e) => setAddress(e.target.value)} placeholder="Alamat"
            className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
          <input value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="No. Telepon"
            className="w-full px-3 py-2 rounded att-body text-sm outline-none" style={{ background: THEME.concrete, color: THEME.ink }} />
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
        <img src={COMPANY_LOGO_DATA_URI} alt="Logo" className="w-9 h-9 rounded object-cover shrink-0" style={{ border: `1px solid ${THEME.line}` }} />
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
  return { name: '', address: '', position: '', phone: '', ktp: '', upahHarian: '', upahLembur: '' };
}

function DataPekerjaTab({ workers, onAdd, onUpdate, onDelete }) {
  const [form, setForm] = useState(emptyForm());
  const [editingId, setEditingId] = useState(null);
  const [formError, setFormError] = useState('');
  const [saving, setSaving] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [showTextPreview, setShowTextPreview] = useState(false);

  const startEdit = (w) => {
    setEditingId(w.id);
    setForm({
      name: w.name || '', address: w.address || '', position: w.position || '',
      phone: w.phone || '', ktp: w.ktp || '',
      upahHarian: String(w.upahHarian ?? ''), upahLembur: String(w.upahLembur ?? ''),
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
          <div key={w.id} className="p-3 rounded-lg flex items-start justify-between gap-2" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div className="min-w-0">
              <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{w.name}</p>
              <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{w.position || '-'}</p>
              <p className="att-mono text-xs mt-1" style={{ color: THEME.inkSoft }}>
                Harian {formatRupiah(w.upahHarian)} &middot; Lembur {formatRupiah(w.upahLembur)}/jam
              </p>
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
        ))}
        {workers.length === 0 && (
          <p className="att-body text-sm text-center py-6" style={{ color: THEME.inkSoft }}>Belum ada pekerja. Tambahkan di atas.</p>
        )}
      </div>
    </div>
  );
}

/* ---------------- Evidence Modal (GPS + Foto) ---------------- */

function EvidenceModal({ onSave, onCancel, existing }) {
  const [geoState, setGeoState] = useState('idle');
  const [coords, setCoords] = useState(existing ? { lat: existing.lat, lng: existing.lng } : null);
  const [photo, setPhoto] = useState(existing ? existing.photo : null);
  const [saving, setSaving] = useState(false);

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

  const onFile = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      const dataUrl = await resizeImage(file);
      setPhoto(dataUrl);
    } catch {
      setPhoto(null);
    }
  };

  const handleSave = async () => {
    setSaving(true);
    await onSave({ lat: coords?.lat ?? null, lng: coords?.lng ?? null, photo });
    setSaving(false);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-4" style={{ background: 'rgba(33,29,26,0.7)' }}>
      <div className="w-full max-w-sm rounded-xl p-5 att-punch-anim att-body" style={{ background: THEME.paper }}>
        <div className="flex items-center justify-between mb-4">
          <h3 className="font-bold text-sm" style={{ color: THEME.ink }}>Bukti Kehadiran</h3>
          <button type="button" onClick={onCancel}><X size={18} color={THEME.inkSoft} /></button>
        </div>

        <button type="button" onClick={captureLocation}
          className="w-full flex items-center justify-center gap-2 py-2.5 rounded mb-2 text-sm font-semibold"
          style={{ background: THEME.concrete, color: THEME.ink }}>
          <MapPin size={16} />
          {geoState === 'loading' ? 'Mengambil lokasi...' : coords ? `Lokasi: ${coords.lat.toFixed(4)}, ${coords.lng.toFixed(4)}` : 'Ambil Lokasi GPS'}
        </button>
        {geoState === 'error' && <p className="text-xs mb-2" style={{ color: THEME.rust }}>Gagal mengambil lokasi. Coba lagi atau lewati.</p>}

        <label className="w-full flex items-center justify-center gap-2 py-2.5 rounded mb-3 text-sm font-semibold cursor-pointer"
          style={{ background: THEME.concrete, color: THEME.ink }}>
          <Camera size={16} />
          {photo ? 'Ganti Foto' : 'Ambil Foto'}
          <input type="file" accept="image/*" capture="environment" className="hidden" onChange={onFile} />
        </label>
        {photo && <img src={photo} alt="bukti" className="w-full rounded-lg mb-3" style={{ maxHeight: 180, objectFit: 'cover' }} />}

        <div className="flex gap-2">
          <button type="button" onClick={handleSave} disabled={saving}
            className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded font-semibold text-sm disabled:opacity-60"
            style={{ background: THEME.amber, color: THEME.charcoal }}>
            {saving ? <Loader2 size={15} className="animate-spin" /> : <Check size={15} />}
            Simpan Bukti
          </button>
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded text-sm" style={{ border: `1px solid ${THEME.line}`, color: THEME.inkSoft }}>
            Lewati
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
              const hasEvidence = !!(evidenceMap && evidenceMap[d.key]);
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

function buildUpahRows(workers, weeks, projects, payments, kasbon, kasbonPayments, filter) {
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
    const weekBreakdown = filteredWeeks
      .map((week) => {
        const totals = calcWorkerWeekTotals(w, week.records[w.id], getWeekDates(week).map((d) => d.key));
        const project = projects.find((p) => p.id === week.projectId);
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
    return { worker: w, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown };
  });
}

function RekapUpahTab({ workers, weeks, projects, payments, kasbon, kasbonPayments, onAddPayment, onDeletePayment }) {
  const [payFormFor, setPayFormFor] = useState(null);
  const [payAmount, setPayAmount] = useState('');
  const [expanded, setExpanded] = useState(null);
  const [showTextPreview, setShowTextPreview] = useState(false);
  const [filter, setFilter] = useState('all'); // 'all' | 'project:<id>' | 'week:<id>'
  const [slipFor, setSlipFor] = useState(null); // worker.id sedang buka slip gaji
  const [showSlipText, setShowSlipText] = useState(null); // worker.id sedang tampil teks slip gaji

  const rows = buildUpahRows(workers, weeks, projects, payments, kasbon, kasbonPayments, filter);
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
          <PrintMenu
            onPrint={(thermal) => openPrintDocument('Rekap Upah', buildRekapHtml(rows, grand, thermal), thermal)}
            onCopy={() => setShowTextPreview(true)}
            onRawBT={() => printToRawBT(buildRekapText(rows, grand))}
          />
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
        {rows.map(({ worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown }) => (
          <div key={worker.id} className="p-3 rounded-lg" style={{ background: THEME.paper, border: `1px solid ${THEME.line}` }}>
            <div className="flex items-center justify-between mb-1">
              <div>
                <p className="att-body font-semibold text-sm" style={{ color: THEME.ink }}>{worker.name}</p>
                <p className="att-mono text-xs" style={{ color: THEME.inkSoft }}>{worker.position || '-'}</p>
              </div>
              <p className="att-mono text-sm font-bold" style={{ color: THEME.ink }}>{formatRupiah(totalUpah)}</p>
            </div>
            <p className="att-mono text-[11px] mb-2" style={{ color: THEME.inkSoft }}>{totalHari} hari kerja &middot; {totalJamLembur} jam lembur</p>
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
              <div className="mt-2">
                <PrintMenu
                  onPrint={(thermal) => openPrintDocument(`Slip Gaji - ${worker.name}`, buildSlipGajiHtml(
                    { worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown },
                    filterLabel(), thermal), thermal)}
                  onCopy={() => setShowSlipText(worker.id)}
                  onRawBT={() => printToRawBT(buildSlipGajiText(
                    { worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown },
                    filterLabel()))}
                />
              </div>
            )}
            {showSlipText === worker.id && (
              <TextPreviewModal
                title={`Slip Gaji - ${worker.name}`}
                text={buildSlipGajiText(
                  { worker, totalUpah, totalHari, totalJamLembur, diterima, totalKasbon, sisa, workerPayments, workerKasbon, workerKasbonPayments, weekBreakdown },
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

  const handleSaveNota = async () => {
    setError('');
    if (cart.length === 0) {
      setError('Tambahkan minimal 1 item belanja ke nota ini dulu.');
      return;
    }
    await onAddPurchase({
      id: uid(), date, supplierId, supplierName: selectedSupplier?.name || '', noNota: noNota.trim(), statusBayar,
      gudangId, note: note.trim(), items: cart,
    });
    setCart([]);
    setNoNota('');
    setNote('');
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
    const gold = [201, 162, 39];
    const navy = [13, 25, 48];
    let y = 40;

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.setTextColor(...navy);
    doc.text(CURRENT_COMPANY?.name || 'Rekap Proyek', 40, y);
    y += 18;
    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(90, 90, 90);
    doc.text(`Rekap Hutang, Cash, Material & Utang Pekerja  |  Dicetak: ${new Date().toLocaleDateString('id-ID')}`, 40, y);
    y += 10;
    doc.setDrawColor(...gold);
    doc.setLineWidth(1.2);
    doc.line(40, y, 555, y);
    y += 18;

    const section = (title, head, rows) => {
      if (y > 740) { doc.addPage(); y = 40; }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...navy);
      doc.text(title, 40, y); y += 8;
      doc.autoTable({
        startY: y, head: [head], body: rows, margin: { left: 40, right: 40 },
        styles: { fontSize: 8.5, textColor: [30, 30, 30] },
        headStyles: { fillColor: navy, textColor: [242, 236, 217] },
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

    doc.save('rekap-proyek.pdf');
  } catch (err) {
    console.error('Gagal ekspor PDF:', err);
    alert('Gagal membuat file PDF. Coba lagi, atau gunakan tombol Cetak (bisa disimpan sebagai PDF lewat dialog cetak).');
  }
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
  const utangPekerja = buildUpahRows(workers, weeks, projects, payments, kasbon, kasbonPayments, 'all')
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

  const upahBelumLunas = buildUpahRows(workers, weeks, projects, payments, kasbon, kasbonPayments, 'all').filter((r) => r.sisa > 0);

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

  return (
    <div className="p-4 pb-24">
      <p className="att-mono text-xs mb-2" style={{ color: THEME.inkSoft }}>
        RINGKASAN &middot; {now.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}
      </p>
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
      const [w, wk, pr, pay, ev, ka, kp, mat, sup, pu, us, up, pe, peu, compSnap] = await Promise.all([
        fetchAll('workers'), fetchAll('weeks'), fetchAll('projects'), fetchAll('payments'),
        fetchAll('evidence'), fetchAll('kasbon'), fetchAll('kasbonPayments'), fetchAll('materials'),
        fetchAll('suppliers'), fetchAll('purchases'), fetchAll('usage'), fetchAll('utangPayments'),
        fetchAll('peralatan'), fetchAll('peralatanUsage'),
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
        Gagal terhubung ke database Supabase. Periksa koneksi internet dan pengaturan VITE_SUPABASE_URL/VITE_SUPABASE_ANON_KEY di file .env.
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
    </div>
  );
}
