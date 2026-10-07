// ============================================================
// Auction Result List Checker
// ============================================================

let allRows = [];       // Data mentah dari Excel
let checkedData = [];   // Data setelah dicek
let currentFilter = 'all';
let currentSearch = '';

// ---------- Helper: Normalisasi string ----------
function norm(v) {
  return String(v ?? '').trim().toUpperCase();
}

// ---------- Helper: Deteksi nama PT ----------
function isPTName(name) {
  if (!name) return false;
  const n = norm(name);
  // Diawali "PT" atau "PT." atau "PT " atau "PT."
  return /^(PT|P\.T)\.?\s/.test(n) || n === 'PT' || n.startsWith('PT.');
}

// ============ PENGECEKAN 1: Tahun unit dari 10 digit nomor rangka ============
// Standar VIN Honda/Yamaha dll: posisi ke-10 = kode tahun
// Kode tahun (standar internasional):
// 2010=A, 2011=B, 2012=C, 2013=D, 2014=E, 2015=F, 2016=G, 2017=H,
// 2018=J, 2019=K, 2020=L, 2021=M, 2022=N, 2023=P, 2024=R, 2025=S, 2026=T, 2027=V...
const YEAR_CODE = {
  'A': 2010, 'B': 2011, 'C': 2012, 'D': 2013, 'E': 2014,
  'F': 2015, 'G': 2016, 'H': 2017, 'J': 2018, 'K': 2019,
  'L': 2020, 'M': 2021, 'N': 2022, 'P': 2023, 'R': 2024,
  'S': 2025, 'T': 2026, 'V': 2027, 'W': 2028, 'X': 2029,
  'Y': 2030
};

function checkYearFromRangka(noRangka, tahunInput) {
  const rangka = norm(noRangka).replace(/[^A-Z0-9]/g, '');
  if (rangka.length < 10) {
    return { ok: false, reason: 'Nomor rangka kurang dari 10 digit', yearFromRangka: null };
  }
  const code = rangka.charAt(9); // index ke-9 = digit ke-10
  const yearFromRangka = YEAR_CODE[code] || null;

  if (!yearFromRangka) {
    return { ok: false, reason: `Kode tahun "${code}" tidak dikenal`, yearFromRangka: null };
  }

  const tahun = parseInt(String(tahunInput).replace(/[^0-9]/g, ''), 10);
  if (isNaN(tahun)) {
    return { ok: false, reason: 'Tahun unit tidak valid', yearFromRangka };
  }

  const ok = tahun === yearFromRangka;
  return {
    ok,
    reason: ok
      ? `Tahun sesuai (${tahun})`
      : `Tahun unit ${tahun} ≠ tahun rangka ${yearFromRangka} (kode "${code}")`,
    yearFromRangka
  };
}

// ============ PENGECEKAN 2: Tipe unit dari CC vs Merk/Type ============
// Ambil angka CC dari merk/type lalu bandingkan dengan kolom CC
function extractCCFromText(text) {
  if (!text) return null;
  const t = norm(text);
  // Cari pola angka 2-4 digit yang diikuti "CC" atau berdiri sendiri (mis. "125", "160")
  const matchCC = t.match(/(\d{2,4})\s*CC/);
  if (matchCC) return parseInt(matchCC[1], 10);

  // Ambil semua angka 2-4 digit, ambil yang paling masuk akal (100-2500)
  const nums = [...t.matchAll(/\b(\d{2,4})\b/g)]
    .map(m => parseInt(m[1], 10))
    .filter(n => n >= 100 && n <= 2500);
  if (nums.length) return nums[nums.length - 1]; // biasanya di akhir
  return null;
}

function checkTypeFromCC(merkType, ccInput) {
  const cc = parseFloat(String(ccInput).replace(/[^0-9.]/g, ''));
  if (isNaN(cc)) {
    return { ok: false, reason: 'CC tidak valid' };
  }
  const ccFromType = extractCCFromText(merkType);
  if (ccFromType === null) {
    return { ok: false, reason: `Tidak bisa ekstrak CC dari "${merkType}"` };
  }

  // Toleransi ±2 CC (karena sering ada pembulatan misal 149.16 vs 150)
  const diff = Math.abs(cc - ccFromType);
  const ok = diff <= 2;

  return {
    ok,
    reason: ok
      ? `CC sesuai (${cc} ≈ ${ccFromType})`
      : `CC tidak sesuai: kolom CC=${cc}, dari tipe=${ccFromType}`,
    ccFromType
  };
}

// ============ PENGECEKAN 3: Nama STNK vs BPKB ============
function checkNameSTNKvsBPKB(namaBPKB, namaSTNK) {
  const a = norm(namaBPKB);
  const b = norm(namaSTNK);

  if (!a && !b) return { ok: true, reason: 'Keduanya kosong' };
  if (!a) return { ok: false, reason: 'Nama BPKB kosong' };
  if (!b) return { ok: false, reason: 'Nama STNK kosong' };

  // Bersihkan gelar/tanda baca untuk perbandingan
  const clean = s => s.replace(/[^A-Z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  const ca = clean(a);
  const cb = clean(b);

  const ok = ca === cb;
  return {
    ok,
    reason: ok
      ? 'Nama STNK & BPKB sama'
      : `Nama STNK "${namaSTNK}" ≠ BPKB "${namaBPKB}"`
  };
}

// ============ PENGECEKAN 4: Note 1 wajib "AN PT" jika nama diawali PT ============
function checkNote1ANPT(namaBPKB, namaSTNK, note1) {
  const isPT = isPTName(namaBPKB) || isPTName(namaSTNK);
  const n1 = norm(note1);

  if (!isPT) {
    return { ok: true, reason: 'Nama bukan PT, tidak perlu AN PT', applicable: false };
  }

  // Cari pola "AN PT", "A.N PT", "A.N. PT", "AN.PT", "A N PT", "A.N PT" dsb.
  const hasANPT = /A\.?\s*N\.?\s*P\.?\s*T/i.test(note1 || '');

  return {
    ok: hasANPT,
    reason: hasANPT
      ? 'Note 1 sudah tertulis AN PT'
      : 'Nama diawali PT tapi Note 1 tidak ada "AN PT"',
    applicable: true
  };
}

// ---------- Baca file Excel ----------
function readExcel(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });

      let allData = [];
      wb.SheetNames.forEach(sheetName => {
        const sheet = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        allData = allData.concat(json.map(r => ({ sheet: sheetName, row: r })));
      });

      processExcel(allData);
    } catch (err) {
      showFileInfo('❌ Gagal membaca file: ' + err.message, true);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ---------- Proses data Excel ----------
function processExcel(rawData) {
  // Cari baris header (yang mengandung "AUCTION NAME" & "LOT NO")
  let headerIdx = -1;
  let headerRow = null;

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i].row.map(c => norm(c));
    if (row.includes('AUCTION NAME') && row.includes('LOT NO')) {
      headerIdx = i;
      headerRow = rawData[i].row;
      break;
    }
  }

  if (headerIdx === -1) {
    showFileInfo('❌ Header tidak ditemukan. Pastikan file berisi kolom "AUCTION NAME" dan "LOT NO".', true);
    return;
  }

  // Mapping nama kolom -> index
  const col = {};
  headerRow.forEach((h, i) => {
    const key = norm(h);
    if (key) col[key] = i;
  });

  // Cari index kolom yang dibutuhkan (dengan toleransi variasi nama)
  const findCol = (...names) => {
    for (const n of names) {
      const key = norm(n);
      if (col[key] !== undefined) return col[key];
    }
    // cari partial
    for (const key of Object.keys(col)) {
      for (const n of names) {
        if (key.includes(norm(n))) return col[key];
      }
    }
    return -1;
  };

  const idx = {
    auction: findCol('AUCTION NAME'),
    lotNo: findCol('LOT NO'),
    merk: findCol('MERK'),
    model: findCol('MODEL'),
    noPol: findCol('NO POL'),
    merkType: findCol('MERK/TYPE'),
    tahun: findCol('TAHUN'),
    cc: findCol('CC'),
    noRangka: findCol('NO.RANGKA', 'NO RANGKA'),
    namaBPKB: findCol('Nama BPKB'),
    namaSTNK: findCol('NAMA STNK'),
    note1: findCol('NOTE 1')
  };

  // Ambil baris data (setelah header)
  allRows = [];
  for (let i = headerIdx + 1; i < rawData.length; i++) {
    const row = rawData[i].row;
    const lotNo = row[idx.lotNo];
    // skip baris kosong / bukan data (harus ada Lot No atau No Pol)
    if (!lotNo && !row[idx.noPol]) continue;
    if (norm(lotNo) === 'LOT NO') continue;

    allRows.push({
      sheet: rawData[i].sheet,
      auction: row[idx.auction] || '',
      lotNo: row[idx.lotNo] || '',
      merk: row[idx.merk] || '',
      model: row[idx.model] || '',
      noPol: row[idx.noPol] || '',
      merkType: row[idx.merkType] || '',
      tahun: row[idx.tahun] || '',
      cc: row[idx.cc] || '',
      noRangka: row[idx.noRangka] || '',
      namaBPKB: row[idx.namaBPKB] || '',
      namaSTNK: row[idx.namaSTNK] || '',
      note1: row[idx.note1] || ''
    });
  }

  // Lakukan pengecekan
  checkedData = allRows.map((d, i) => {
    const yearCheck = checkYearFromRangka(d.noRangka, d.tahun);
    const typeCheck = checkTypeFromCC(d.merkType, d.cc);
    const nameCheck = checkNameSTNKvsBPKB(d.namaBPKB, d.namaSTNK);
    const noteCheck = checkNote1ANPT(d.namaBPKB, d.namaSTNK, d.note1);

    const hasError = !yearCheck.ok || !typeCheck.ok || !nameCheck.ok || !noteCheck.ok;

    return {
      ...d,
      no: i + 1,
      yearCheck,
      typeCheck,
      nameCheck,
      noteCheck,
      hasError
    };
  });

  // Tampilkan
  showFileInfo(`✅ Berhasil memuat ${checkedData.length} unit dari ${new Set(checkedData.map(d => d.sheet)).size} sheet.`);
  renderSummary();
  renderTable();
  document.getElementById('summarySection').classList.remove('hidden');
  document.getElementById('tableSection').classList.remove('hidden');
}

// ---------- Render summary ----------
function renderSummary() {
  const total = checkedData.length;
  const totalYear = checkedData.filter(d => !d.yearCheck.ok).length;
  const totalType = checkedData.filter(d => !d.typeCheck.ok).length;
  const totalName = checkedData.filter(d => !d.nameCheck.ok).length;
  const totalNote = checkedData.filter(d => !d.noteCheck.ok).length;

  document.getElementById('totalUnits').textContent = total;
  document.getElementById('totalYear').textContent = totalYear;
  document.getElementById('totalType').textContent = totalType;
  document.getElementById('totalName').textContent = totalName;
  document.getElementById('totalNote').textContent = totalNote;
}

// ---------- Render tabel ----------
function renderTable() {
  const tbody = document.querySelector('#resultTable tbody');
  const filter = currentFilter;
  const search = currentSearch.toLowerCase();

  let data = checkedData.filter(d => {
    // filter
    if (filter === 'year' && d.yearCheck.ok) return false;
    if (filter === 'type' && d.typeCheck.ok) return false;
    if (filter === 'name' && d.nameCheck.ok) return false;
    if (filter === 'note' && d.noteCheck.ok) return false;
    if (filter === 'any' && !d.hasError) return false;

    // search
    if (search) {
      const hay = [
        d.lotNo, d.noPol, d.merkType, d.merk, d.model,
        d.tahun, d.cc, d.noRangka, d.namaBPKB, d.namaSTNK, d.note1, d.auction
      ].join(' ').toLowerCase();
      if (!hay.includes(search)) return false;
    }
    return true;
  });

  if (!data.length) {
    tbody.innerHTML = `<tr><td colspan="12" style="text-align:center; padding:2rem; color:#6b7280;">
      Tidak ada data yang cocok dengan filter.
    </td></tr>`;
    return;
  }

  tbody.innerHTML = data.map(d => {
    const badges = [];
    if (d.yearCheck.ok) badges.push('<span class="badge badge-ok">TAHUN ✓</span>');
    else badges.push('<span class="badge badge-err">TAHUN ✗</span>');

    if (d.typeCheck.ok) badges.push('<span class="badge badge-ok">TIPE ✓</span>');
    else badges.push('<span class="badge badge-err">TIPE ✗</span>');

    if (d.nameCheck.ok) badges.push('<span class="badge badge-ok">NAMA ✓</span>');
    else badges.push('<span class="badge badge-err">NAMA ✗</span>');

    if (d.noteCheck.applicable) {
      if (d.noteCheck.ok) badges.push('<span class="badge badge-ok">AN PT ✓</span>');
      else badges.push('<span class="badge badge-err">AN PT ✗</span>');
    } else {
      badges.push('<span class="badge badge-warn">AN PT -</span>');
    }

    const detailList = [
      d.yearCheck,
      d.typeCheck,
      d.nameCheck,
      d.noteCheck
    ].map(c => `<li class="${c.ok ? 'ok' : ''}">${c.reason}</li>`).join('');

    return `
      <tr class="${d.hasError ? 'row-error' : ''}">
        <td>${d.no}</td>
        <td><strong>${d.lotNo}</strong></td>
        <td>${d.noPol}</td>
        <td>${d.merkType || (d.merk + ' ' + d.model)}</td>
        <td>${d.tahun}</td>
        <td>${d.cc}</td>
        <td style="font-family:monospace; font-size:0.78rem;">${d.noRangka}</td>
        <td>${d.namaBPKB}</td>
        <td>${d.namaSTNK}</td>
        <td style="max-width:180px;">${d.note1}</td>
        <td>${badges.join(' ')}</td>
        <td><ul class="detail-list">${detailList}</ul></td>
      </tr>
    `;
  }).join('');
}

// ---------- UI Handlers ----------
function showFileInfo(msg, isError = false) {
  const el = document.getElementById('fileInfo');
  el.textContent = msg;
  el.classList.remove('hidden');
  el.classList.toggle('error', isError);
}

// Drop zone
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');

dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});

fileInput.addEventListener('change', (e) => {
  if (e.target.files.length) handleFile(e.target.files[0]);
});

function handleFile(file) {
  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    showFileInfo('❌ File harus berformat .xlsx atau .xls', true);
    return;
  }
  showFileInfo(`⏳ Membaca file: ${file.name} (${(file.size/1024).toFixed(1)} KB)...`);
  readExcel(file);
}

// Filter & search
document.getElementById('filterSelect').addEventListener('change', (e) => {
  currentFilter = e.target.value;
  renderTable();
});

let searchTimer;
document.getElementById('searchInput').addEventListener('input', (e) => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => {
    currentSearch = e.target.value;
    renderTable();
  }, 250);
});

// Export hasil
document.getElementById('exportBtn').addEventListener('click', () => {
  if (!checkedData.length) return alert('Belum ada data untuk di-export.');

  const rows = checkedData.map(d => ({
    'No': d.no,
    'Auction': d.auction,
    'Lot No': d.lotNo,
    'No Pol': d.noPol,
    'Merk/Type': d.merkType,
    'Tahun': d.tahun,
    'CC': d.cc,
    'No Rangka': d.noRangka,
    'Nama BPKB': d.namaBPKB,
    'Nama STNK': d.namaSTNK,
    'Note 1': d.note1,
    'Cek Tahun': d.yearCheck.ok ? 'OK' : 'TIDAK SESUAI',
    'Ket Tahun': d.yearCheck.reason,
    'Cek Tipe/CC': d.typeCheck.ok ? 'OK' : 'TIDAK SESUAI',
    'Ket Tipe/CC': d.typeCheck.reason,
    'Cek Nama STNK vs BPKB': d.nameCheck.ok ? 'OK' : 'TIDAK SAMA',
    'Ket Nama': d.nameCheck.reason,
    'Cek AN PT': d.noteCheck.applicable ? (d.noteCheck.ok ? 'OK' : 'MISSING') : 'N/A',
    'Ket AN PT': d.noteCheck.reason
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Hasil Cek');
  XLSX.writeFile(wb, `Hasil_Cek_Auction_${new Date().toISOString().slice(0,10)}.xlsx`);
});