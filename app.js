/* Almeyra's Shop - aplikasi pencatatan penjualan souvenir (mandiri, berjalan offline) */
(function () {
  'use strict';
  var R = window.AlmeyraReport;
  var $ = function (id) { return document.getElementById(id); };
  var DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  var LS_TX = 'almeyra-shop-tx-v1';
  var LS_BK = 'almeyra-shop-lastbackup';
  var BK_REMIND_DAYS = 30, BK_REMIND_MIN_TX = 10;

  /* ---------- util DOM ---------- */
  function h(tag, props) {
    var e = document.createElement(tag);
    if (props) for (var k in props) {
      var v = props[k];
      if (v == null || v === false) continue;
      if (k === 'class') e.className = v;
      else if (k === 'text') e.textContent = v;
      else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v === true ? '' : v);
    }
    function add(c) {
      if (c == null || c === false) return;
      if (Array.isArray(c)) { c.forEach(add); return; }
      e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    }
    for (var i = 2; i < arguments.length; i++) add(arguments[i]);
    return e;
  }
  function clear(el) { while (el.firstChild) el.removeChild(el.firstChild); }
  function pad(n) { return n < 10 ? '0' + n : '' + n; }
  function todayISO() { var d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function clean(s) { return String(s).trim().replace(/\s+/g, ' '); }
  function digits(s) { var d = String(s).replace(/\D/g, ''); return d ? parseInt(d, 10) : NaN; }

  var toastTimer = null;
  function toast(msg, kind) {
    var t = $('toast'); t.textContent = msg; t.className = kind === 'err' ? 'err' : ''; t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.hidden = true; }, kind === 'err' ? 7000 : 3500);
  }

  /* ---------- state ---------- */
  var today = todayISO();
  var S = {
    txs: [], tab: 'catat', editId: null, confirmId: null, busy: false, blocked: false,
    repMode: 'mingguan', repWeek: today, repMonth: today.slice(0, 7), repFrom: today.slice(0, 8) + '01', repTo: today,
    hMonth: today.slice(0, 7), hAll: false, hQuery: '', canShare: false
  };

  function validTx(t) {
    return !!t && typeof t.tgl === 'string' && !!R.parseISO(t.tgl) && typeof t.pembeli === 'string' && typeof t.barang === 'string' &&
      t.pembeli.trim() !== '' && t.barang.trim() !== '' &&
      Number.isInteger(t.jumlah) && t.jumlah >= 1 && t.jumlah <= 999999 &&
      Number.isInteger(t.harga) && t.harga >= 1 && t.harga <= 999999999;
  }
  function find(id) { for (var i = 0; i < S.txs.length; i++) if (S.txs[i].id === id) return S.txs[i]; return null; }
  function sortDesc(list) {
    return list.slice().sort(function (a, b) {
      if (a.tgl !== b.tgl) return a.tgl < b.tgl ? 1 : -1;
      return (b.createdAt || 0) - (a.createdAt || 0);
    });
  }
  function newId() { return 't' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }

  /* ---------- penyimpanan di perangkat (localStorage) ---------- */
  function loadTx() {
    try {
      var raw = localStorage.getItem(LS_TX), a = raw ? JSON.parse(raw) : [];
      return Array.isArray(a) ? a.filter(validTx) : [];
    } catch (e) { return []; }
  }
  // commit: tulis dulu, baru ubah memori; bila gagal, data lama tetap utuh.
  function commit(next) {
    try { localStorage.setItem(LS_TX, JSON.stringify(next)); }
    catch (e) { return Promise.reject({ code: 'storage' }); }
    S.txs = next; renderAll();
    return Promise.resolve();
  }
  var store = {
    add: function (tx) { return commit(S.txs.concat([Object.assign({ id: newId() }, tx)])); },
    update: function (id, tx) { return commit(S.txs.map(function (t) { return t.id === id ? Object.assign({ id: id }, tx) : t; })); },
    remove: function (id) { return commit(S.txs.filter(function (t) { return t.id !== id; })); }
  };
  function errMsg() { return 'Data tidak dapat disimpan. Ruang penyimpanan penuh atau penyimpanan browser diblokir (mode penyamaran/privat).'; }

  /* ---------- banner ---------- */
  function lastBackup() { try { return parseInt(localStorage.getItem(LS_BK), 10) || 0; } catch (e) { return 0; } }
  function updateBanner() {
    var b = $('banner'), text = '', kind = '';
    if (S.blocked) { text = 'Penyimpanan browser diblokir sehingga data tidak dapat disimpan. Buka aplikasi di jendela biasa (bukan mode privat).'; kind = 'err'; }
    else if (S.txs.length >= BK_REMIND_MIN_TX && Date.now() - lastBackup() > BK_REMIND_DAYS * 86400000) {
      text = 'Sudah ada ' + R.num(S.txs.length) + ' transaksi dan belum ada cadangan dalam ' + BK_REMIND_DAYS + ' hari terakhir. Buka menu Riwayat, lalu pilih "Unduh cadangan".';
    }
    if (!text) { b.hidden = true; return; }
    b.textContent = text; b.className = 'banner' + (kind ? ' ' + kind : ''); b.hidden = false;
  }

  /* ---------- form ---------- */
  function formatDigits(input, maxDigits) {
    var pos = input.selectionStart == null ? input.value.length : input.selectionStart;
    var before = input.value.slice(0, pos).replace(/\D/g, '').length;
    var d = input.value.replace(/\D/g, '').replace(/^0+(?=\d)/, '').slice(0, maxDigits);
    var f = d ? R.num(parseInt(d, 10)) : '';
    input.value = f;
    var cnt = 0, i = 0;
    for (; i < f.length && cnt < before; i++) if (/\d/.test(f.charAt(i))) cnt++;
    try { input.setSelectionRange(i, i); } catch (e) { /* abaikan */ }
  }
  function readForm() {
    return { tgl: $('fTgl').value, pembeli: clean($('fPembeli').value), barang: clean($('fBarang').value), jumlah: digits($('fJumlah').value), harga: digits($('fHarga').value) };
  }
  function validate(f) {
    var e = {};
    if (!R.parseISO(f.tgl) || f.tgl < '2000-01-01' || f.tgl > '2100-12-31') e.Tgl = 'Isi tanggal penjualan yang valid.';
    if (!f.pembeli) e.Pembeli = 'Nama pembeli wajib diisi.';
    if (!f.barang) e.Barang = 'Jenis barang wajib diisi.';
    if (!(f.jumlah >= 1)) e.Jumlah = 'Jumlah barang minimal 1.';
    else if (f.jumlah > 999999) e.Jumlah = 'Jumlah barang maksimal 999.999.';
    if (!(f.harga >= 1)) e.Harga = 'Harga per satuan harus lebih dari 0.';
    else if (f.harga > 999999999) e.Harga = 'Harga per satuan maksimal Rp 999.999.999.';
    return e;
  }
  function showErrors(errs) {
    ['Tgl', 'Pembeli', 'Barang', 'Jumlah', 'Harga'].forEach(function (k) {
      $('e' + k).textContent = errs[k] || '';
      var inp = $('f' + k);
      if (errs[k]) inp.setAttribute('aria-invalid', 'true'); else inp.removeAttribute('aria-invalid');
    });
  }
  function updateTotal() {
    var j = digits($('fJumlah').value), p = digits($('fHarga').value);
    $('fTotal').textContent = R.rp(j >= 1 && p >= 1 ? j * p : 0);
  }
  function updateFutureHint() {
    var v = $('fTgl').value;
    $('hTgl').hidden = !(R.parseISO(v) && v > todayISO());
  }
  function renderSaveButton() {
    var btn = $('btnSave');
    btn.disabled = S.busy;
    btn.textContent = S.busy ? 'Menyimpan…' : (S.editId ? 'Simpan perubahan' : 'Simpan transaksi');
  }
  function enterEdit(id) {
    var t = find(id); if (!t) return;
    S.editId = id; S.confirmId = null;
    $('fTgl').value = t.tgl; $('fPembeli').value = t.pembeli; $('fBarang').value = t.barang;
    $('fJumlah').value = R.num(t.jumlah); $('fHarga').value = R.num(t.harga);
    showErrors({}); updateTotal(); updateFutureHint();
    $('formTitle').textContent = 'Ubah Transaksi'; $('btnCancel').hidden = false;
    setTab('catat'); renderSaveButton(); renderDay();
    $('form').scrollIntoView({ behavior: 'smooth', block: 'start' });
    $('fPembeli').focus();
  }
  function exitEdit() {
    S.editId = null;
    $('formTitle').textContent = 'Catat Penjualan'; $('btnCancel').hidden = true;
    ['fPembeli', 'fBarang', 'fJumlah', 'fHarga'].forEach(function (i) { $(i).value = ''; });
    showErrors({}); updateTotal(); renderSaveButton();
  }
  function onSave(ev) {
    ev.preventDefault();
    if (S.busy) return;
    var f = readForm(), errs = validate(f);
    showErrors(errs);
    var keys = ['Tgl', 'Pembeli', 'Barang', 'Jumlah', 'Harga'];
    for (var i = 0; i < keys.length; i++) if (errs[keys[i]]) { $('f' + keys[i]).focus(); return; }
    S.busy = true; renderSaveButton();
    var editing = S.editId, p;
    if (editing) {
      var old = find(editing);
      p = store.update(editing, Object.assign({}, f, { createdAt: (old && old.createdAt) || Date.now() }));
    } else {
      p = store.add(Object.assign({}, f, { createdAt: Date.now() }));
    }
    p.then(function () {
      toast((editing ? 'Perubahan disimpan: ' : 'Transaksi disimpan: ') + R.num(f.jumlah) + ' × ' + f.barang + ' = ' + R.rp(f.jumlah * f.harga));
      exitEdit();
      $('fPembeli').focus();
    }).catch(function (e) {
      toast(errMsg(e), 'err');
    }).then(function () { S.busy = false; renderSaveButton(); });
  }

  /* ---------- tampilan: daftar transaksi ---------- */
  function txList(list, showDate) {
    var wrap = h('div', { class: 'txlist' });
    list.forEach(function (t) {
      var actions;
      if (S.confirmId === t.id) {
        actions = h('div', { class: 'confirm' }, 'Hapus transaksi ini?',
          h('button', { type: 'button', class: 'btn small danger', onclick: function () { doDelete(t.id); } }, 'Ya, hapus'),
          h('button', { type: 'button', class: 'btn small secondary', onclick: function () { S.confirmId = null; renderAll(); } }, 'Batal'));
      } else {
        actions = h('div', { class: 'row-actions' },
          h('button', { type: 'button', class: 'btn small secondary', onclick: function () { enterEdit(t.id); }, 'aria-label': 'Ubah transaksi ' + t.barang + ' untuk ' + t.pembeli }, 'Ubah'),
          h('button', { type: 'button', class: 'btn small secondary', onclick: function () { S.confirmId = t.id; renderAll(); }, 'aria-label': 'Hapus transaksi ' + t.barang + ' untuk ' + t.pembeli }, 'Hapus'));
      }
      wrap.appendChild(h('div', { class: 'tx' + (S.editId === t.id ? ' editing' : '') },
        h('div', { class: 'tx-main' }, h('div', { class: 'tx-title' }, t.barang), h('div', { class: 'tx-total' }, R.rp(R.lineTotal(t)))),
        h('div', { class: 'tx-sub' }, (showDate ? R.fmtShort(t.tgl) + ' · ' : '') + t.pembeli + ' · ' + R.num(t.jumlah) + ' × ' + R.rp(t.harga)),
        actions));
    });
    return wrap;
  }
  function doDelete(id) {
    store.remove(id).then(function () {
      S.confirmId = null;
      if (S.editId === id) exitEdit();
      toast('Transaksi dihapus.');
      renderAll();
    }).catch(function (e) { toast(errMsg(e), 'err'); });
  }

  function renderDatalists() {
    function fill(id, getter) {
      var seen = {}, vals = [];
      sortDesc(S.txs).forEach(function (t) {
        var v = getter(t), k = v.trim().toLowerCase();
        if (!seen[k]) { seen[k] = 1; vals.push(v); }
      });
      var dl = $(id); clear(dl);
      vals.slice(0, 200).forEach(function (v) { dl.appendChild(h('option', { value: v })); });
    }
    fill('dlBuyers', function (t) { return t.pembeli; });
    fill('dlItems', function (t) { return t.barang; });
  }

  function renderDay() {
    var date = $('fTgl').value;
    var box = $('dayList'); clear(box);
    if (!R.parseISO(date)) { $('dayTitle').textContent = 'Transaksi pada tanggal ini'; $('daySummary').textContent = ''; return; }
    $('dayTitle').textContent = 'Transaksi tanggal ' + R.fmtDate(date);
    var list = sortDesc(S.txs.filter(function (t) { return t.tgl === date; }));
    var A = R.aggregate(S.txs, date, date);
    if (!list.length) { $('daySummary').textContent = ''; box.appendChild(h('div', { class: 'empty' }, 'Belum ada transaksi pada tanggal ini.')); return; }
    clear($('daySummary'));
    $('daySummary').appendChild(h('b', null, R.num(A.sum.trx) + ' transaksi'));
    $('daySummary').appendChild(document.createTextNode(' · ' + R.num(A.sum.unit) + ' unit · total '));
    $('daySummary').appendChild(h('b', null, R.rp(A.sum.total)));
    box.appendChild(txList(list, false));
  }

  function renderHistory() {
    var box = $('hList'); clear(box);
    $('hMonth').disabled = S.hAll;
    var q = S.hQuery.trim().toLowerCase();
    var list = sortDesc(S.txs.filter(function (t) {
      if (!S.hAll && t.tgl.slice(0, 7) !== S.hMonth) return false;
      if (q && t.pembeli.toLowerCase().indexOf(q) < 0 && t.barang.toLowerCase().indexOf(q) < 0) return false;
      return true;
    }));
    var unit = 0, total = 0;
    list.forEach(function (t) { unit += t.jumlah; total += R.lineTotal(t); });
    var CAP = 300, shown = list.length > CAP ? list.slice(0, CAP) : list;
    var sum = $('hSummary'); clear(sum);
    if (!list.length) {
      box.appendChild(h('div', { class: 'empty' }, S.txs.length ? 'Tidak ada transaksi yang cocok dengan filter.' : 'Belum ada transaksi. Mulai dari menu Catat.'));
    } else {
      sum.appendChild(h('b', null, R.num(list.length) + ' transaksi'));
      sum.appendChild(document.createTextNode(' · ' + R.num(unit) + ' unit · total '));
      sum.appendChild(h('b', null, R.rp(total)));
      if (list.length > CAP) {
        box.appendChild(h('div', { class: 'hint mb8' }, 'Menampilkan ' + CAP + ' transaksi terbaru dari ' + R.num(list.length) + '. Persempit filter untuk melihat sisanya. Ringkasan di atas mencakup seluruh ' + R.num(list.length) + ' transaksi.'));
      }
      box.appendChild(txList(shown, true));
    }
    renderBackupStatus();
  }
  function renderBackupStatus() {
    var t = lastBackup();
    var d = t ? new Date(t) : null;
    $('bkStatus').textContent = d
      ? 'Cadangan terakhir: ' + d.getDate() + ' ' + ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'][d.getMonth()] + ' ' + d.getFullYear() + ' pukul ' + pad(d.getHours()) + '.' + pad(d.getMinutes())
      : 'Belum pernah dicadangkan.';
  }

  /* ---------- berkas: unduh / bagikan ---------- */
  function saveBlob(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = name; a.rel = 'noopener'; a.style.display = 'none';
    document.body.appendChild(a); a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 4000);
  }

  /* ---------- cadangan & pemulihan ---------- */
  function onBackup() {
    var data = { app: 'almeyras-shop', versi: 1, dibuat: new Date().toISOString(), jumlah: S.txs.length, transaksi: S.txs };
    var blob = new Blob([JSON.stringify(data, null, 1)], { type: 'application/json' });
    var name = 'Cadangan-Almeyras-Shop_' + todayISO() + '.json';
    saveBlob(blob, name);
    try { localStorage.setItem(LS_BK, String(Date.now())); } catch (e) { /* abaikan */ }
    renderBackupStatus(); updateBanner();
    toast('Cadangan diunduh: ' + name + ' (' + R.num(S.txs.length) + ' transaksi). Simpan file ini di tempat aman.');
  }
  function onRestoreFile(file) {
    if (!file) return;
    var fr = new FileReader();
    fr.onerror = function () { toast('File tidak dapat dibaca.', 'err'); };
    fr.onload = function () {
      var obj;
      try { obj = JSON.parse(String(fr.result)); } catch (e) { obj = null; }
      if (!obj || obj.app !== 'almeyras-shop' || !Array.isArray(obj.transaksi)) {
        toast('File ini bukan cadangan Almeyra\'s Shop.', 'err'); return;
      }
      var have = {}; S.txs.forEach(function (t) { have[t.id] = 1; });
      var add = [], dupe = 0, bad = 0;
      obj.transaksi.forEach(function (t) {
        if (!t || typeof t !== 'object') { bad++; return; }
        var c = {
          id: typeof t.id === 'string' && t.id ? t.id : newId(), tgl: t.tgl,
          pembeli: typeof t.pembeli === 'string' ? clean(t.pembeli).slice(0, 80) : '',
          barang: typeof t.barang === 'string' ? clean(t.barang).slice(0, 80) : '',
          jumlah: t.jumlah, harga: t.harga, createdAt: Number.isFinite(t.createdAt) ? t.createdAt : Date.now()
        };
        if (!validTx(c)) { bad++; return; }
        if (have[c.id]) { dupe++; return; }
        have[c.id] = 1; add.push(c);
      });
      if (!add.length) {
        toast('Tidak ada transaksi baru. ' + dupe + ' sudah ada' + (bad ? ', ' + bad + ' tidak valid' : '') + '.', dupe ? undefined : 'err'); return;
      }
      commit(S.txs.concat(add)).then(function () {
        toast('Dipulihkan: ' + R.num(add.length) + ' transaksi ditambahkan' + (dupe ? ', ' + dupe + ' sudah ada' : '') + (bad ? ', ' + bad + ' dilewati (tidak valid)' : '') + '.');
      }).catch(function (e) { toast(errMsg(e), 'err'); });
    };
    fr.readAsText(file);
  }

  /* ---------- laporan ---------- */
  function currentPeriod() {
    if (S.repMode === 'mingguan') return R.makePeriod('mingguan', S.repWeek);
    if (S.repMode === 'bulanan') return R.makePeriod('bulanan', S.repMonth);
    return R.makePeriod('rentang', { from: S.repFrom, to: S.repTo });
  }
  function stat(k, v, main) { return h('div', { class: 'stat' + (main ? ' main' : '') }, h('div', { class: 'k' }, k), h('div', { class: 'v' }, v)); }

  function renderReport() {
    var pv = $('preview'); clear(pv);
    ['mingguan', 'bulanan', 'rentang'].forEach(function (m) {
      $('m-' + m).setAttribute('aria-pressed', String(S.repMode === m));
      $('c-' + m).hidden = S.repMode !== m;
    });
    var p = currentPeriod();
    var btn = $('btnDownload'), sh = $('btnShare');
    sh.hidden = !S.canShare;
    $('rErr').textContent = '';
    function off() { btn.disabled = true; sh.disabled = true; }
    if (!p) {
      if (S.repMode === 'rentang' && R.parseISO(S.repFrom) && R.parseISO(S.repTo) && S.repFrom > S.repTo) $('rErr').textContent = 'Tanggal "Dari" tidak boleh setelah tanggal "Sampai".';
      else $('rErr').textContent = 'Lengkapi pilihan periode dengan benar.';
      off(); return;
    }
    var A = R.aggregate(S.txs, p.start, p.end);
    pv.appendChild(h('div', { class: 'periodlabel' }, p.label));
    if (!A.sum.trx) {
      pv.appendChild(h('div', { class: 'empty' }, 'Belum ada transaksi pada periode ini.'));
      off(); return;
    }
    pv.appendChild(h('div', { class: 'stats' },
      stat('Total penjualan', R.rp(A.sum.total), true), stat('Transaksi', R.num(A.sum.trx)),
      stat('Unit terjual', R.num(A.sum.unit)), stat('Rata-rata/transaksi', R.rp(A.avgTrx))));
    pv.appendChild(h('h3', null, 'Rekap per jenis barang'));
    pv.appendChild(h('div', { class: 'tablewrap' }, h('table', null,
      h('thead', null, h('tr', null, h('th', null, 'Jenis barang'), h('th', { class: 'r' }, 'Transaksi'), h('th', { class: 'r' }, 'Unit'), h('th', { class: 'r' }, 'Total'))),
      h('tbody', null, A.byItem.map(function (i) {
        return h('tr', null, h('td', null, i.nama), h('td', { class: 'r' }, R.num(i.trx)), h('td', { class: 'r' }, R.num(i.unit)), h('td', { class: 'r' }, R.rp(i.total)));
      })),
      h('tfoot', null, h('tr', null, h('td', null, 'Total'), h('td', { class: 'r' }, R.num(A.sum.trx)), h('td', { class: 'r' }, R.num(A.sum.unit)), h('td', { class: 'r' }, R.rp(A.sum.total)))))));
    pv.appendChild(h('p', { class: 'muted note13' },
      'Dokumen Word memuat ringkasan, rekap per hari' + (R.daysBetween(p.start, p.end) > 6 ? ' dan per minggu' : '') + ', rekap per jenis barang, rekap per pembeli, serta rincian seluruh transaksi.'));
    btn.disabled = false; sh.disabled = false;
  }
  // mode: 'unduh' | 'bagikan'
  function makeReport(mode) {
    var p = currentPeriod(); if (!p) return;
    var A = R.aggregate(S.txs, p.start, p.end); if (!A.sum.trx) return;
    if (typeof JSZip === 'undefined') { toast('Pembuat dokumen belum termuat. Muat ulang halaman.', 'err'); return; }
    var btn = mode === 'unduh' ? $('btnDownload') : $('btnShare');
    var label = btn.textContent; btn.disabled = true; btn.textContent = 'Menyiapkan…';
    var name = R.fileName(p);
    R.makeDocx(JSZip, S.txs, p, todayISO(), 'blob').then(function (blob) {
      if (mode === 'unduh') {
        saveBlob(blob, name);
        toast('Laporan diunduh: ' + name + ' (cek folder Download atau notifikasi unduhan).');
        return;
      }
      var file = new File([blob], name, { type: DOCX_MIME });
      return navigator.share({ files: [file], title: p.title + ' - Almeyra\'s Shop', text: p.title + ' ' + p.label });
    }).catch(function (e) {
      if (e && e.name === 'AbortError') toast('Dibatalkan.');
      else toast('Laporan gagal dibuat atau dibagikan. Coba lagi.', 'err');
    }).then(function () { btn.textContent = label; renderReport(); });
  }

  /* ---------- tab & render ---------- */
  function setTab(t) {
    S.tab = t;
    ['catat', 'riwayat', 'laporan'].forEach(function (k) {
      $('t-' + k).setAttribute('aria-selected', String(k === t));
      $('p-' + k).hidden = k !== t;
    });
    renderAll();
  }
  function renderAll() {
    renderSaveButton(); renderDatalists(); updateBanner();
    if (S.tab === 'catat') renderDay();
    if (S.tab === 'riwayat') renderHistory();
    if (S.tab === 'laporan') renderReport();
  }

  /* ---------- event ---------- */
  function bind() {
    $('fTgl').value = today; $('hMonth').value = S.hMonth; $('rWeek').value = S.repWeek; $('rMonth').value = S.repMonth;
    $('rFrom').value = S.repFrom; $('rTo').value = S.repTo;
    ['catat', 'riwayat', 'laporan'].forEach(function (k) { $('t-' + k).addEventListener('click', function () { setTab(k); }); });
    $('form').addEventListener('submit', onSave);
    $('btnCancel').addEventListener('click', function () { exitEdit(); renderDay(); });
    $('fTgl').addEventListener('input', function () { updateFutureHint(); $('eTgl').textContent = ''; $('fTgl').removeAttribute('aria-invalid'); renderDay(); });
    $('fJumlah').addEventListener('input', function () { formatDigits($('fJumlah'), 6); updateTotal(); });
    $('fHarga').addEventListener('input', function () { formatDigits($('fHarga'), 9); updateTotal(); });
    ['fPembeli', 'fBarang', 'fJumlah', 'fHarga'].forEach(function (id) {
      $(id).addEventListener('input', function () { $('e' + id.slice(1)).textContent = ''; $(id).removeAttribute('aria-invalid'); });
    });

    $('hMonth').addEventListener('input', function () { if (/^\d{4}-\d{2}$/.test($('hMonth').value)) { S.hMonth = $('hMonth').value; S.confirmId = null; renderHistory(); } });
    $('hAll').addEventListener('change', function () { S.hAll = $('hAll').checked; S.confirmId = null; renderHistory(); });
    $('hQuery').addEventListener('input', function () { S.hQuery = $('hQuery').value; S.confirmId = null; renderHistory(); });
    $('btnBackup').addEventListener('click', onBackup);
    $('btnRestore').addEventListener('click', function () { $('fRestore').value = ''; $('fRestore').click(); });
    $('fRestore').addEventListener('change', function () { onRestoreFile($('fRestore').files[0]); });

    ['mingguan', 'bulanan', 'rentang'].forEach(function (m) { $('m-' + m).addEventListener('click', function () { S.repMode = m; renderReport(); }); });
    $('rWeek').addEventListener('input', function () { if (R.parseISO($('rWeek').value)) { S.repWeek = $('rWeek').value; renderReport(); } });
    $('wPrev').addEventListener('click', function () { var b = R.parseISO(S.repWeek) ? S.repWeek : today; S.repWeek = R.addDays(b, -7); $('rWeek').value = S.repWeek; renderReport(); });
    $('wNext').addEventListener('click', function () { var b = R.parseISO(S.repWeek) ? S.repWeek : today; S.repWeek = R.addDays(b, 7); $('rWeek').value = S.repWeek; renderReport(); });
    $('rMonth').addEventListener('input', function () { if (/^\d{4}-\d{2}$/.test($('rMonth').value)) { S.repMonth = $('rMonth').value; renderReport(); } });
    $('rFrom').addEventListener('input', function () { S.repFrom = $('rFrom').value; renderReport(); });
    $('rTo').addEventListener('input', function () { S.repTo = $('rTo').value; renderReport(); });
    $('btnDownload').addEventListener('click', function () { makeReport('unduh'); });
    $('btnShare').addEventListener('click', function () { makeReport('bagikan'); });
  }

  /* ---------- pasang sebagai aplikasi + offline ---------- */
  var installEvt = null;
  function setupPwa() {
    window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); installEvt = e; $('btnInstall').hidden = false; });
    window.addEventListener('appinstalled', function () { installEvt = null; $('btnInstall').hidden = true; toast('Aplikasi terpasang di layar utama.'); });
    $('btnInstall').addEventListener('click', function () {
      if (!installEvt) return;
      installEvt.prompt();
      installEvt.userChoice.then(function () { installEvt = null; $('btnInstall').hidden = true; });
    });
    if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
      window.addEventListener('load', function () { navigator.serviceWorker.register('sw.js').catch(function () { /* tanpa mode offline */ }); });
    }
    try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch (e) { /* abaikan */ }
  }

  /* ---------- mulai ---------- */
  function init() {
    try { localStorage.setItem('almeyra-probe', '1'); localStorage.removeItem('almeyra-probe'); } catch (e) { S.blocked = true; }
    S.txs = loadTx();
    try {
      var probe = new File([''], 'a.docx', { type: DOCX_MIME });
      S.canShare = !!(navigator.share && navigator.canShare && navigator.canShare({ files: [probe] }));
    } catch (e) { S.canShare = false; }
    bind(); updateTotal(); updateFutureHint(); setupPwa(); renderAll();
  }
  init();
})();
