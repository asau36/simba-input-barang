/**
 * Merapikan sheet Master_Item:
 *  1. Mengelompokkan baris berdasarkan Kategori (lalu urut Kode Barang).
 *  2. Memberi warna kategori (kolom C:E) dan garis pemisah antar kategori.
 *  3. Memberi fill Stok Akhir (kolom L): merah = habis, oranye = menipis, hijau = aman.
 *
 * Cara pakai: buka spreadsheet > Extensions > Apps Script > tempel file ini,
 * lalu jalankan rapikanMasterItem() sekali (izinkan akses bila diminta).
 * Jalankan lagi setiap ada barang baru supaya tetap terkelompok.
 *
 * Blok "Ringkasan" (kolom Q:V, termasuk checkbox Update Aplikasi) TIDAK ikut
 * dipindah; hanya kolom A:P dan W:X yang diurutkan.
 */
var MI_SHEET = 'Master_Item';
var MI_KOLOM_KODE = 3;      // C
var MI_KOLOM_KATEGORI = 5;  // E
var MI_KOLOM_STOK_AKHIR = 12; // L
var MI_URUTAN_KATEGORI = ['GWP', 'GWP & Tools', 'POSM', 'Product Sampel', 'Visual', 'ATK', 'Elektronik', 'Aset', 'Seragam'];
var MI_WARNA_KATEGORI = {
  'GWP':            { terang: '#E3F2FD', tegas: '#90CAF9' },
  'GWP & Tools':    { terang: '#E0F2F1', tegas: '#80CBC4' },
  'POSM':           { terang: '#FFF8E1', tegas: '#FFE082' },
  'Product Sampel': { terang: '#F3E5F5', tegas: '#CE93D8' },
  'Visual':         { terang: '#FBE9E7', tegas: '#FFAB91' },
  'ATK':            { terang: '#F1F8E9', tegas: '#C5E1A5' },
  'Elektronik':     { terang: '#E8EAF6', tegas: '#9FA8DA' },
  'Aset':           { terang: '#EFEBE9', tegas: '#BCAAA4' },
  'Seragam':        { terang: '#FCE4EC', tegas: '#F48FB1' }
};

function rapikanMasterItem() {
  var sh = SpreadsheetApp.getActive().getSheetByName(MI_SHEET);
  if (!sh) throw new Error('Sheet ' + MI_SHEET + ' tidak ditemukan');

  var akhir = miBarisTerakhir_(sh);
  if (akhir < 2) return;
  var n = akhir - 1;

  miUrutkan_(sh, n);
  miGarisKategori_(sh, akhir);
  miAturanWarna_(sh);
  SpreadsheetApp.flush();
}

// Urutkan via R1C1 agar rumus relatif (mis. =I2+J2-K2) tetap benar di baris barunya.
function miUrutkan_(sh, n) {
  var kiri = sh.getRange(2, 1, n, 16);   // A:P
  var kanan = sh.getRange(2, 23, n, 2);  // W:X
  var fKiri = kiri.getFormulasR1C1();
  var fKanan = kanan.getFormulasR1C1();
  var nilaiKiri = kiri.getValues();

  var idx = [];
  for (var i = 0; i < n; i++) idx.push(i);
  idx.sort(function (a, b) {
    var ka = String(nilaiKiri[a][MI_KOLOM_KATEGORI - 1]).trim();
    var kb = String(nilaiKiri[b][MI_KOLOM_KATEGORI - 1]).trim();
    var ra = miUrutKategori_(ka), rb = miUrutKategori_(kb);
    if (ra !== rb) return ra - rb;
    if (ra === 999 && ka !== kb) return ka < kb ? -1 : 1;
    var ca = String(nilaiKiri[a][MI_KOLOM_KODE - 1]);
    var cb = String(nilaiKiri[b][MI_KOLOM_KODE - 1]);
    return ca.localeCompare(cb, undefined, { numeric: true }) || a - b;
  });

  kiri.setFormulasR1C1(idx.map(function (i) { return fKiri[i]; }));
  kanan.setFormulasR1C1(idx.map(function (i) { return fKanan[i]; }));
}

function miUrutKategori_(k) {
  var p = MI_URUTAN_KATEGORI.indexOf(k);
  return p < 0 ? 999 : p;
}

function miBarisTerakhir_(sh) {
  var kode = sh.getRange(1, MI_KOLOM_KODE, sh.getMaxRows(), 1).getValues();
  for (var i = kode.length - 1; i >= 1; i--) {
    if (String(kode[i][0]).trim() !== '') return i + 1;
  }
  return 1;
}

function miGarisKategori_(sh, akhir) {
  var area = sh.getRange(2, 1, akhir - 1, 16);
  area.setBorder(false, null, false, null, null, null); // reset garis atas/bawah lama
  var kat = sh.getRange(2, MI_KOLOM_KATEGORI, akhir - 1, 1).getValues();
  for (var i = 0; i < kat.length; i++) {
    if (i === 0 || kat[i][0] !== kat[i - 1][0]) {
      sh.getRange(i + 2, 1, 1, 16).setBorder(true, null, null, null, null, null,
        '#37474F', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }
}

function miAturanWarna_(sh) {
  var batas = sh.getMaxRows();
  var lama = sh.getConditionalFormatRules();
  var tetap = lama.filter(function (r) {
    var rg = r.getRanges().map(function (x) { return x.getA1Notation(); }).join(',');
    return !/^(L2:L|C2:E|E2:E)/.test(rg) || /^C1:C/.test(rg);
  });

  var rL = 'L2:L' + batas;
  var rCE = 'C2:E' + batas;
  var rE = 'E2:E' + batas;
  var baru = [];

  // Stok Akhir — urutan = prioritas
  baru.push(miAturan_(sh, rL, '=AND($C2<>"",$L2<=0)', '#F4C7C3', '#B71C1C'));
  baru.push(miAturan_(sh, rL, '=AND($C2<>"",$L2<=$H2)', '#FFE0B2', '#E65100'));
  baru.push(miAturan_(sh, rL, '=AND($C2<>"",$L2>$H2)', '#C8E6C9', '#1B5E20'));

  // Kategori: kolom E lebih tegas, C:D lebih terang
  Object.keys(MI_WARNA_KATEGORI).forEach(function (k) {
    var w = MI_WARNA_KATEGORI[k];
    var f = '=$E2="' + k + '"';
    baru.push(miAturan_(sh, rE, f, w.tegas, null));
    baru.push(miAturan_(sh, rCE, f, w.terang, null));
  });

  sh.setConditionalFormatRules(tetap.concat(baru));
  sh.getRange(rL).setFontWeight('bold').setHorizontalAlignment('center');
}

function miAturan_(sh, a1, rumus, latar, teks) {
  var b = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(rumus)
    .setBackground(latar)
    .setRanges([sh.getRange(a1)]);
  if (teks) b.setFontColor(teks);
  return b.build();
}
