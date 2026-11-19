/* Almeyra's Shop - perhitungan rekap + pembuat laporan Word (.docx)
   Murni fungsi: tidak menyentuh DOM, dapat diuji di Node. */
(function (root) {
  'use strict';

  var BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  var BLN = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  var HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  var SHOP = "Almeyra's Shop";

  /* ---------- tanggal (semua dalam format ISO YYYY-MM-DD, hitung di UTC) ---------- */
  function pad(n) { return n < 10 ? '0' + n : '' + n; }
  function parseISO(s) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || '');
    if (!m) return null;
    var y = +m[1], mo = +m[2], d = +m[3];
    var dt = new Date(Date.UTC(y, mo - 1, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return null;
    return dt;
  }
  function toISO(dt) { return dt.getUTCFullYear() + '-' + pad(dt.getUTCMonth() + 1) + '-' + pad(dt.getUTCDate()); }
  function addDays(iso, n) { var dt = parseISO(iso); dt.setUTCDate(dt.getUTCDate() + n); return toISO(dt); }
  function weekStart(iso) { return addDays(iso, -((parseISO(iso).getUTCDay() + 6) % 7)); } // Senin
  function daysBetween(a, b) { return Math.round((parseISO(b) - parseISO(a)) / 86400000); }
  function fmtDate(iso) { var d = parseISO(iso); return d.getUTCDate() + ' ' + BULAN[d.getUTCMonth()] + ' ' + d.getUTCFullYear(); }
  function fmtShort(iso) { var d = parseISO(iso); return d.getUTCDate() + ' ' + BLN[d.getUTCMonth()] + ' ' + d.getUTCFullYear(); }
  function fmtDay(iso) { var d = parseISO(iso); return HARI[d.getUTCDay()] + ', ' + fmtShort(iso); }
  function monthLabel(ym) { var m = /^(\d{4})-(\d{2})$/.exec(ym); return BULAN[+m[2] - 1] + ' ' + m[1]; }

  /* ---------- angka ---------- */
  function num(n) { return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
  function rp(n) { return (n < 0 ? '-' : '') + 'Rp ' + num(Math.abs(n)); }

  /* ---------- periode laporan ---------- */
  // mode: 'mingguan' (input = tanggal mana pun dalam minggu), 'bulanan' (input = 'YYYY-MM'),
  //       'rentang' (input = {from, to})
  function makePeriod(mode, input) {
    var start, end, label, title;
    if (mode === 'mingguan') {
      if (!parseISO(input)) return null;
      start = weekStart(input); end = addDays(start, 6);
      title = 'Laporan Penjualan Mingguan';
      label = fmtDate(start) + ' – ' + fmtDate(end) + ' (Senin–Minggu)';
    } else if (mode === 'bulanan') {
      var m = /^(\d{4})-(\d{2})$/.exec(input || '');
      if (!m || +m[2] < 1 || +m[2] > 12) return null;
      var last = new Date(Date.UTC(+m[1], +m[2], 0)).getUTCDate();
      start = m[1] + '-' + m[2] + '-01'; end = m[1] + '-' + m[2] + '-' + pad(last);
      title = 'Laporan Penjualan Bulanan';
      label = monthLabel(input) + ' (' + fmtDate(start) + ' – ' + fmtDate(end) + ')';
    } else if (mode === 'rentang') {
      if (!input || !parseISO(input.from) || !parseISO(input.to) || input.from > input.to) return null;
      start = input.from; end = input.to;
      title = 'Laporan Penjualan Periode';
      label = fmtDate(start) + ' – ' + fmtDate(end);
    } else return null;
    return { mode: mode, start: start, end: end, label: label, title: title };
  }

  /* ---------- agregasi ---------- */
  function norm(s) { return String(s).trim().replace(/\s+/g, ' ').toLowerCase(); }
  function lineTotal(t) { return t.jumlah * t.harga; }
  function cmpTx(a, b) {
    if (a.tgl !== b.tgl) return a.tgl < b.tgl ? -1 : 1;
    return (a.createdAt || 0) - (b.createdAt || 0);
  }
  function byTotalDesc(a, b) {
    if (b.total !== a.total) return b.total - a.total;
    return a.nama < b.nama ? -1 : a.nama > b.nama ? 1 : 0;
  }
  function group(list, keyFn, nameFn) {
    var map = {}, order = [];
    list.forEach(function (t) {
      var k = keyFn(t);
      if (!map[k]) { map[k] = { nama: nameFn(t), unit: 0, trx: 0, total: 0 }; order.push(k); }
      map[k].unit += t.jumlah; map[k].trx += 1; map[k].total += lineTotal(t);
    });
    return order.map(function (k) { return map[k]; });
  }

  function aggregate(txs, start, end) {
    var list = txs.filter(function (t) { return t.tgl >= start && t.tgl <= end; }).sort(cmpTx);
    var sum = { trx: list.length, unit: 0, total: 0 };
    list.forEach(function (t) { sum.unit += t.jumlah; sum.total += lineTotal(t); });
    var byItem = group(list, function (t) { return norm(t.barang); }, function (t) { return t.barang.trim().replace(/\s+/g, ' '); }).sort(byTotalDesc);
    var byBuyer = group(list, function (t) { return norm(t.pembeli); }, function (t) { return t.pembeli.trim().replace(/\s+/g, ' '); }).sort(byTotalDesc);
    var byDay = group(list, function (t) { return t.tgl; }, function (t) { return t.tgl; })
      .sort(function (a, b) { return a.nama < b.nama ? -1 : 1; });
    var byWeek = group(list, function (t) { return weekStart(t.tgl); }, function (t) { return weekStart(t.tgl); })
      .sort(function (a, b) { return a.nama < b.nama ? -1 : 1; })
      .map(function (w) {
        var ws = w.nama, we = addDays(ws, 6);
        var a = ws < start ? start : ws, b = we > end ? end : we;
        w.label = a === b ? fmtShort(a) : fmtShort(a) + ' – ' + fmtShort(b);
        return w;
      });
    // Tampilan rekap: seluruh minggu dalam periode (termasuk yang bernilai 0);
    // seluruh tanggal bila periode <= 7 hari, selain itu hanya hari yang ada penjualan.
    var wkMap = {}; byWeek.forEach(function (w) { wkMap[w.nama] = w; });
    var weekRows = [];
    for (var ws = weekStart(start); ws <= end; ws = addDays(ws, 7)) {
      var we = addDays(ws, 6), a = ws < start ? start : ws, bb = we > end ? end : we;
      var lbl = a === bb ? fmtShort(a) : fmtShort(a) + ' – ' + fmtShort(bb);
      var w0 = wkMap[ws];
      weekRows.push({ label: lbl, trx: w0 ? w0.trx : 0, unit: w0 ? w0.unit : 0, total: w0 ? w0.total : 0 });
    }
    var dayFull = daysBetween(start, end) <= 6; // periode <= 7 hari: semua tanggal tampil
    var dayRows;
    if (dayFull) {
      var dMap = {}; byDay.forEach(function (d) { dMap[d.nama] = d; });
      dayRows = [];
      for (var dd = start; dd <= end; dd = addDays(dd, 1)) {
        var d0 = dMap[dd];
        dayRows.push({ nama: dd, trx: d0 ? d0.trx : 0, unit: d0 ? d0.unit : 0, total: d0 ? d0.total : 0 });
      }
    } else dayRows = byDay;
    var topUnit = byItem.slice().sort(function (a, b) {
      if (b.unit !== a.unit) return b.unit - a.unit;
      return a.nama < b.nama ? -1 : 1;
    })[0] || null;
    var topDay = byDay.slice().sort(function (a, b) {
      if (b.total !== a.total) return b.total - a.total;
      return a.nama < b.nama ? -1 : 1;
    })[0] || null;
    return {
      list: list, sum: sum, byItem: byItem, byBuyer: byBuyer, byDay: byDay, byWeek: byWeek,
      weekRows: weekRows, dayRows: dayRows, dayFull: dayFull,
      avgTrx: list.length ? Math.round(sum.total / list.length) : 0,
      topValue: byItem[0] || null, topUnit: topUnit, topDay: topDay
    };
  }

  /* ---------- WordprocessingML ---------- */
  var C_ROSE = 'B5456B', C_HEAD_TXT = 'FFFFFF', C_LINE = 'D9C7C0', C_SOFT = 'FBE9EF', C_TOTAL = 'F6E3D6', C_MUTED = '6F5F63';
  var PAGE_W = 11906, PAGE_H = 16838, MARGIN = 1134, CONTENT_W = PAGE_W - 2 * MARGIN; // 9638

  function esc(s) {
    return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function run(text, o) {
    o = o || {};
    var p = '';
    if (o.b) p += '<w:b/>';
    if (o.i) p += '<w:i/>';
    if (o.color) p += '<w:color w:val="' + o.color + '"/>';
    if (o.sz) p += '<w:sz w:val="' + o.sz + '"/><w:szCs w:val="' + o.sz + '"/>';
    return '<w:r>' + (p ? '<w:rPr>' + p + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
  }
  function para(runs, o) {
    o = o || {};
    var pp = '';
    if (o.style) pp += '<w:pStyle w:val="' + o.style + '"/>';
    if (o.keepNext) pp += '<w:keepNext/>';
    if (o.before != null || o.after != null) {
      pp += '<w:spacing' + (o.before != null ? ' w:before="' + o.before + '"' : '') + (o.after != null ? ' w:after="' + o.after + '"' : '') + '/>';
    }
    if (o.jc) pp += '<w:jc w:val="' + o.jc + '"/>';
    return '<w:p>' + (pp ? '<w:pPr>' + pp + '</w:pPr>' : '') + (runs || '') + '</w:p>';
  }
  function heading(text) { return para(run(text), { style: 'Heading1' }); }

  function cell(c, width, o) {
    // c: string atau {t, span, b}; o: {head, total, jc, label}
    var obj = typeof c === 'object' && c !== null ? c : { t: c };
    var w = width;
    var tcPr = '<w:tcW w:w="' + w + '" w:type="dxa"/>';
    if (obj.span && obj.span > 1) tcPr += '<w:gridSpan w:val="' + obj.span + '"/>';
    var fill = o.head ? C_ROSE : o.total ? C_TOTAL : o.label ? C_SOFT : null;
    if (fill) tcPr += '<w:shd w:val="clear" w:color="auto" w:fill="' + fill + '"/>';
    tcPr += '<w:vAlign w:val="center"/>';
    var bold = o.head || o.total || o.label || obj.b;
    var r = run(obj.t == null ? '' : obj.t, { b: bold, color: o.head ? C_HEAD_TXT : null });
    return '<w:tc><w:tcPr>' + tcPr + '</w:tcPr>' + para(r, { before: 0, after: 0, jc: o.jc }) + '</w:tc>';
  }
  // table(widths, head, rows, aligns, opts{totalRow:[...], labelCol:true})
  function table(widths, head, rows, aligns, opts) {
    opts = opts || {};
    var total = widths.reduce(function (a, b) { return a + b; }, 0);
    var x = '<w:tbl><w:tblPr><w:tblW w:w="' + total + '" w:type="dxa"/><w:tblBorders>';
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].forEach(function (s) {
      x += '<w:' + s + ' w:val="single" w:sz="4" w:space="0" w:color="' + C_LINE + '"/>';
    });
    x += '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="50" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>';
    widths.forEach(function (w) { x += '<w:gridCol w:w="' + w + '"/>'; });
    x += '</w:tblGrid>';
    function row(cells, ro) {
      var out = '<w:tr><w:trPr><w:cantSplit/>' + (ro.head ? '<w:tblHeader/>' : '') + '</w:trPr>';
      var col = 0;
      cells.forEach(function (c) {
        var span = typeof c === 'object' && c !== null && c.span ? c.span : 1;
        var w = 0;
        for (var i = 0; i < span; i++) w += widths[col + i];
        var jc = ro.head ? (aligns && aligns[col] === 'right' ? 'right' : (aligns && aligns[col] === 'center' ? 'center' : 'left')) : (aligns && aligns[col]) || 'left';
        if (ro.labelCol && col === 0) jc = 'left';
        out += cell(c, w, { head: ro.head, total: ro.total, label: ro.labelCol && col === 0, jc: jc });
        col += span;
      });
      return out + '</w:tr>';
    }
    if (head) x += row(head, { head: true });
    rows.forEach(function (r) { x += row(r, { labelCol: opts.labelCol }); });
    if (opts.totalRow) x += row(opts.totalRow, { total: true });
    return x + '</w:tbl>';
  }

  function buildBody(txs, period, todayIso) {
    var A = aggregate(txs, period.start, period.end);
    var b = '';
    b += para(run(SHOP.toUpperCase(), { b: true, color: C_ROSE, sz: 22 }), { before: 0, after: 40 });
    b += para(run(period.title), { style: 'Title' });
    b += para(run(period.label, { color: C_MUTED, sz: 22 }), { before: 0, after: 40 });
    b += para(run('Dibuat pada ' + fmtDate(todayIso), { color: C_MUTED, sz: 18 }), { before: 0, after: 200 });

    // Ringkasan
    b += heading('Ringkasan');
    var rows = [
      ['Jumlah transaksi', num(A.sum.trx)],
      ['Total unit terjual', num(A.sum.unit)],
      ['Total penjualan', rp(A.sum.total)],
      ['Rata-rata nilai per transaksi', rp(A.avgTrx)],
      ['Jumlah pembeli', num(A.byBuyer.length)],
      ['Jumlah jenis barang terjual', num(A.byItem.length)]
    ];
    if (A.topValue) rows.push(['Barang dengan nilai penjualan tertinggi', A.topValue.nama + ' (' + rp(A.topValue.total) + ')']);
    if (A.topUnit) rows.push(['Barang terjual terbanyak (unit)', A.topUnit.nama + ' (' + num(A.topUnit.unit) + ' unit)']);
    if (A.topDay) rows.push(['Hari dengan penjualan tertinggi', fmtDay(A.topDay.nama) + ' (' + rp(A.topDay.total) + ')']);
    b += table([3900, CONTENT_W - 3900], null, rows, ['left', 'left'], { labelCol: true });

    // Rekap per minggu (hanya jika periode lebih dari 7 hari)
    if (daysBetween(period.start, period.end) > 6) {
      b += heading('Rekap per Minggu');
      b += table([3638, 1400, 1700, 2900], ['Minggu (Senin–Minggu)', 'Transaksi', 'Unit', 'Total Penjualan'],
        A.weekRows.map(function (w) { return [w.label, num(w.trx), num(w.unit), rp(w.total)]; }),
        ['left', 'right', 'right', 'right'],
        { totalRow: ['Total', num(A.sum.trx), num(A.sum.unit), rp(A.sum.total)] });
    }

    // Rekap per hari
    b += heading(A.dayFull ? 'Rekap per Hari' : 'Rekap per Hari (hanya hari yang ada penjualan)');
    b += table([3638, 1400, 1700, 2900], ['Tanggal', 'Transaksi', 'Unit', 'Total Penjualan'],
      A.dayRows.map(function (d) { return [fmtDay(d.nama), num(d.trx), num(d.unit), rp(d.total)]; }),
      ['left', 'right', 'right', 'right'],
      { totalRow: ['Total', num(A.sum.trx), num(A.sum.unit), rp(A.sum.total)] });

    // Rekap per barang
    b += heading('Rekap per Jenis Barang');
    b += table([3038, 1100, 1000, 1900, 2600], ['Jenis Barang', 'Transaksi', 'Unit', 'Harga Rata-rata/Unit', 'Total Penjualan'],
      A.byItem.map(function (i) { return [i.nama, num(i.trx), num(i.unit), rp(i.total / i.unit), rp(i.total)]; }),
      ['left', 'right', 'right', 'right', 'right'],
      { totalRow: ['Total', num(A.sum.trx), num(A.sum.unit), '', rp(A.sum.total)] });

    // Rekap per pembeli
    b += heading('Rekap per Pembeli');
    b += table([3638, 1400, 1700, 2900], ['Nama Pembeli', 'Transaksi', 'Unit', 'Total Pembelian'],
      A.byBuyer.map(function (i) { return [i.nama, num(i.trx), num(i.unit), rp(i.total)]; }),
      ['left', 'right', 'right', 'right'],
      { totalRow: ['Total', num(A.sum.trx), num(A.sum.unit), rp(A.sum.total)] });

    // Rincian
    b += heading('Rincian Transaksi');
    b += table([500, 1400, 1900, 2038, 700, 1500, 1600], ['No', 'Tanggal', 'Pembeli', 'Jenis Barang', 'Jml', 'Harga/Unit', 'Total'],
      A.list.map(function (t, i) { return [String(i + 1), fmtShort(t.tgl), t.pembeli, t.barang, num(t.jumlah), rp(t.harga), rp(lineTotal(t))]; }),
      ['center', 'left', 'left', 'left', 'right', 'right', 'right'],
      { totalRow: [{ t: 'Total', span: 4 }, num(A.sum.unit), '', rp(A.sum.total)] });

    b += para(run('Dokumen ini dibuat otomatis oleh aplikasi ' + SHOP + '. Total = jumlah × harga per satuan; semua nilai dalam Rupiah.', { i: true, color: C_MUTED, sz: 16 }), { before: 240, after: 0 });
    return b;
  }

  var NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  var XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  function stylesXml() {
    return XML + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="id-ID" w:eastAsia="id-ID" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="0" w:after="60"/></w:pPr><w:rPr><w:b/><w:color w:val="2D2326"/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="320" w:after="100"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="' + C_ROSE + '"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>' +
      '</w:styles>';
  }
  function footerXml() {
    function fld(instr) {
      return '<w:r><w:rPr><w:color w:val="' + C_MUTED + '"/><w:sz w:val="16"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:rPr><w:color w:val="' + C_MUTED + '"/><w:sz w:val="16"/></w:rPr><w:instrText xml:space="preserve"> ' + instr + ' </w:instrText></w:r>' +
        '<w:r><w:rPr><w:color w:val="' + C_MUTED + '"/><w:sz w:val="16"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:rPr><w:color w:val="' + C_MUTED + '"/><w:sz w:val="16"/></w:rPr><w:t>1</w:t></w:r>' +
        '<w:r><w:rPr><w:color w:val="' + C_MUTED + '"/><w:sz w:val="16"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>';
    }
    var p = para(run(SHOP + ' — Halaman ', { color: C_MUTED, sz: 16 }) + fld('PAGE') + run(' dari ', { color: C_MUTED, sz: 16 }) + fld('NUMPAGES'), { jc: 'center', before: 0, after: 0 });
    return XML + '<w:ftr ' + NS + '>' + p + '</w:ftr>';
  }

  // hasil: { 'path/dalam/zip': 'isi xml' }
  function buildDocxFiles(txs, period, todayIso) {
    var doc = XML + '<w:document ' + NS + '><w:body>' + buildBody(txs, period, todayIso) +
      '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/><w:pgSz w:w="' + PAGE_W + '" w:h="' + PAGE_H + '"/>' +
      '<w:pgMar w:top="' + MARGIN + '" w:right="' + MARGIN + '" w:bottom="' + MARGIN + '" w:left="' + MARGIN + '" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>';
    var files = {};
    files['[Content_Types].xml'] = XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>';
    files['_rels/.rels'] = XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>';
    files['word/_rels/document.xml.rels'] = XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>';
    files['word/document.xml'] = doc;
    files['word/styles.xml'] = stylesXml();
    files['word/footer1.xml'] = footerXml();
    files['docProps/core.xml'] = XML + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + esc(period.title + ' - ' + SHOP) + '</dc:title><dc:creator>' + esc(SHOP) + '</dc:creator>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + todayIso + 'T00:00:00Z</dcterms:created></cp:coreProperties>';
    return files;
  }

  // JSZipCtor: konstruktor JSZip; type: 'blob' (browser) atau 'nodebuffer' (Node)
  function makeDocx(JSZipCtor, txs, period, todayIso, type) {
    var files = buildDocxFiles(txs, period, todayIso);
    var zip = new JSZipCtor();
    // [Content_Types].xml harus entri pertama
    zip.file('[Content_Types].xml', files['[Content_Types].xml']);
    Object.keys(files).forEach(function (k) { if (k !== '[Content_Types].xml') zip.file(k, files[k]); });
    return zip.generateAsync({
      type: type || 'blob', compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
  }

  function fileName(period) {
    var kind = period.mode === 'mingguan' ? 'Mingguan' : period.mode === 'bulanan' ? 'Bulanan' : 'Periode';
    return 'Laporan-Penjualan-' + kind + '_' + period.start + '_sd_' + period.end + '.docx';
  }

  var API = {
    pad: pad, parseISO: parseISO, toISO: toISO, addDays: addDays, weekStart: weekStart, daysBetween: daysBetween,
    fmtDate: fmtDate, fmtShort: fmtShort, fmtDay: fmtDay, monthLabel: monthLabel,
    num: num, rp: rp, lineTotal: lineTotal, makePeriod: makePeriod, aggregate: aggregate,
    buildDocxFiles: buildDocxFiles, makeDocx: makeDocx, fileName: fileName
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  else root.AlmeyraReport = API;
})(typeof window !== 'undefined' ? window : this);
