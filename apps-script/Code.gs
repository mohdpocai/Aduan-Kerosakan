/**
 * Aduan Kerosakan LP2 - simpan aduan ke Google Sheet dan gambar ke Google Drive.
 * Tampal kod ini dalam Apps Script, kemudian Deploy sebagai Web app.
 */
const SHEET_ID  = '1_ZgsW-EDqbtIH2MK4MzrCu_GjVR8QjVQLAOxk35h4ZQ'; // Google Sheet "Aduan Kerosakan LP2"
const FOLDER_ID = '1FntAmCKj7ZXeID7lCEnxrWQNKSNRRzNT';            // Folder "Aduan Kerosakan LP2 - Gambar"

const HEADERS = ['Masa Hantar','ID Aduan','No. Rumah','Tarikh','Nama Pelapor','No. Telefon',
  'Keutamaan','Bil. Kerosakan','Kerosakan','Catatan','Gambar','Status','Tindakan / Catatan Penyelia','Status Kerosakan'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'status') return out(kemaskiniStatus(d));
    const ss = SpreadsheetApp.openById(SHEET_ID);
    if (ss.getSpreadsheetTimeZone() !== 'Asia/Kuala_Lumpur') ss.setSpreadsheetTimeZone('Asia/Kuala_Lumpur');
    const sh = ss.getSheets()[0];
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS);

    // Elak rekod berganda jika telefon hantar dua kali
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, 2, last - 1, 1).getValues().flat();
      if (ids.indexOf(d.id) !== -1) return out({ ok: true, id: d.id, dup: true });
    }

    // Simpan gambar dalam subfolder ikut bulan, cth "2026-11"
    const month = (d.tarikh || '').slice(0, 7) || Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyy-MM');
    const root = DriveApp.getFolderById(FOLDER_ID);
    const it = root.getFoldersByName(month);
    const folder = it.hasNext() ? it.next() : root.createFolder(month);

    const urls = (d.photos || []).map((b64, i) => {
      const name = d.id + '_' + (d.rumah || 'rumah').replace(/[^\w-]+/g, '') + '_' + (i + 1) + '.jpg';
      const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name);
      return folder.createFile(blob).getUrl();
    });

    const items = (d.items || []).map((x, i) => (i + 1) + '. ' + x.name + (x.note ? ' - ' + x.note : '')).join('\n');

    sh.appendRow([new Date(), d.id, d.rumah, d.tarikh, d.nama, "'" + (d.tel || ''), d.pri,
      (d.items || []).length, items, d.catatan || '', '', 'Baru', '', '']);
    const row = sh.getLastRow();

    // Pautan gambar boleh ditekan: "Gambar 1, Gambar 2, ..."
    if (urls.length) {
      const labels = urls.map((_, i) => 'Gambar ' + (i + 1));
      const text = labels.join(', ');
      const rt = SpreadsheetApp.newRichTextValue().setText(text);
      let pos = 0;
      labels.forEach((l, i) => { rt.setLinkUrl(pos, pos + l.length, urls[i]); pos += l.length + 2; });
      sh.getRange(row, 11).setRichTextValue(rt.build());
    }
    sh.getRange(row, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange(row, 9).setWrap(true);
    if (d.pri === 'KECEMASAN') sh.getRange(row, 7).setBackground('#f8d3bd').setFontWeight('bold');

    return out({ ok: true, id: d.id, row: row });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Buka link Web app dalam browser untuk semak ia berfungsi.
// Tambah ?pdf=LP2 pada hujung link untuk buat Laporan PDF bulan ini.
// Contoh bulan lain: ?pdf=LP2&bulan=2026-10
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.pdf === PDF_KEY) return buatPDF(p.bulan);
  if (p.data !== undefined) return p.data === DATA_PIN ? out(senaraiAduan()) : out({ ok: false, error: 'PIN salah' });
  return out({ ok: true, msg: 'Aduan Kerosakan LP2 sedia menerima aduan' });
}

/* ---------- DATA UNTUK HALAMAN RINGKASAN (laporan.html) ---------- */
const DATA_PIN = '708602';

// Tukar status setiap kerosakan dari halaman Ringkasan.
// Lajur N "Status Kerosakan": satu baris setiap item, cth "1. Tingkap: Selesai - tukar kaca"
const ST_OK = ['Baru', 'Dalam Tindakan', 'Selesai'];

function bacaStatusItem(text, count) {
  const out = [];
  String(text || '').split('\n').forEach(line => {
    const m = line.match(/^(\d+)\.\s*.*?:\s*(Baru|Dalam Tindakan|Selesai)(?:\s+-\s+([\s\S]*))?$/);
    if (m) out[+m[1] - 1] = { status: m[2], tindakan: (m[3] || '').trim() };
  });
  const res = [];
  for (let i = 0; i < count; i++) res.push(out[i] || { status: 'Baru', tindakan: '' });
  return res;
}

function statusKeseluruhan(st) {
  if (!st.length) return 'Baru';
  if (st.every(x => x.status === 'Selesai')) return 'Selesai';
  if (st.some(x => x.status !== 'Baru')) return 'Dalam Tindakan';
  return 'Baru';
}

function kemaskiniStatus(d) {
  if (d.pin !== DATA_PIN) return { ok: false, error: 'PIN salah' };
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const n = sh.getLastRow() - 1;
  if (n < 1) return { ok: false, error: 'Tiada aduan' };
  const ids = sh.getRange(2, 2, n, 1).getValues().flat().map(String);
  const i = ids.indexOf(String(d.id));
  if (i === -1) return { ok: false, error: 'Aduan tidak dijumpai' };
  const row = i + 2;
  if (sh.getRange(1, 14).getValue() === '') sh.getRange(1, 14).setValue('Status Kerosakan').setFontWeight('bold');

  const names = String(sh.getRange(row, 9).getValue() || '').split('\n').filter(String)
    .map(t => t.replace(/^\d+\.\s*/, '').split(' - ')[0].trim());
  const items = (d.items || []).slice(0, names.length).map(x => ({
    status: ST_OK.indexOf(x.status) > -1 ? x.status : 'Baru',
    tindakan: String(x.tindakan || '').replace(/\s*\n\s*/g, ' ').slice(0, 300)
  }));
  while (items.length < names.length) items.push({ status: 'Baru', tindakan: '' });

  const text = names.map((nm, k) => (k + 1) + '. ' + nm + ': ' + items[k].status + (items[k].tindakan ? ' - ' + items[k].tindakan : '')).join('\n');
  const status = statusKeseluruhan(items);
  sh.getRange(row, 14).setValue(text).setWrap(true);
  sh.getRange(row, 12).setValue(status)
    .setBackground(status === 'Selesai' ? '#e2f1e6' : (status === 'Dalam Tindakan' ? '#f8f0d6' : null));
  if (d.tindakan !== undefined) sh.getRange(row, 13).setValue(String(d.tindakan || '').slice(0, 1000));
  return { ok: true, id: d.id, status: status, items: items };
}

function senaraiAduan() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const n = sh.getLastRow() - 1;
  if (n < 1) return { ok: true, aduan: [] };
  const rows = sh.getRange(2, 1, n, 14).getValues();
  const rich = sh.getRange(2, 11, n, 1).getRichTextValues();
  const aduan = rows.map((r, i) => {
    const id = String(r[1] || '');
    const m = id.match(/^AK-(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})/);
    const photos = [];
    (rich[i][0] ? rich[i][0].getRuns() : []).forEach(run => {
      const u = run.getLinkUrl(); const f = u && u.match(/\/d\/([\w-]+)/); if (f) photos.push(f[1]);
    });
    return {
      id: id, row: i + 2,
      masa: m ? m[1] + '-' + m[2] + '-' + m[3] + 'T' + m[4] + ':' + m[5] : '',
      rumah: String(r[2] || ''), nama: String(r[4] || ''), tel: String(r[5] || ''), pri: String(r[6] || 'BIASA'),
      items: (() => {
        const lines = String(r[8] || '').split('\n').filter(String);
        const st = bacaStatusItem(r[13], lines.length);
        return lines.map((t, k) => {
          const s = t.replace(/^\d+\.\s*/, ''), c = s.indexOf(' - ');
          const it = c > -1 ? { name: s.slice(0, c), note: s.slice(c + 3) } : { name: s, note: '' };
          it.status = st[k].status; it.tindakan = st[k].tindakan; return it;
        });
      })(),
      catatan: String(r[9] || ''), status: String(r[11] || 'Baru'), tindakan: String(r[12] || ''), photos: photos
    };
  }).filter(a => a.id);
  return { ok: true, aduan: aduan };
}

/* ---------- LAPORAN PDF ---------- */
const PDF_KEY = 'LP2';
const TZ = 'Asia/Kuala_Lumpur';

function buatPDF(bulan) {
  try {
    const now = new Date();
    bulan = /^\d{4}-\d{2}$/.test(bulan || '') ? bulan : Utilities.formatDate(now, TZ, 'yyyy-MM');
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    const n = sh.getLastRow() - 1;
    const rows = n > 0 ? sh.getRange(2, 1, n, 14).getValues() : [];
    const rich = n > 0 ? sh.getRange(2, 11, n, 1).getRichTextValues() : [];

    const list = [];
    rows.forEach((r, i) => {
      const id = String(r[1] || '');
      const m = id.match(/^AK-(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})/);
      const ym = m ? m[1] + '-' + m[2] : String(r[3]).slice(0, 7);
      if (ym !== bulan) return;
      const links = [];
      (rich[i][0] ? rich[i][0].getRuns() : []).forEach(run => { const u = run.getLinkUrl(); if (u) links.push(u); });
      list.push({
        id: id, masa: m ? m[3] + '/' + m[2] + '/' + m[1] + ', ' + m[4] + ':' + m[5] : '',
        rumah: r[2], nama: r[4] || '(tidak dinyatakan)', tel: r[5], pri: r[6],
        items: String(r[8] || '').split('\n').filter(String), catatan: r[9],
        status: r[11] || 'Baru', tindakan: r[12], links: links,
        itemSt: bacaStatusItem(r[13], String(r[8] || '').split('\n').filter(String).length)
      });
    });

    const html = laporanHTML(list, bulan, now);
    const name = 'Laporan_Aduan_Kerosakan_LP2_' + bulan + '.pdf';
    const pdf = Utilities.newBlob(html, 'text/html', 'laporan.html').getAs('application/pdf').setName(name);

    const root = DriveApp.getFolderById(FOLDER_ID);
    const it = root.getFoldersByName('Laporan PDF');
    const folder = it.hasNext() ? it.next() : root.createFolder('Laporan PDF');
    const old = folder.getFilesByName(name);
    while (old.hasNext()) old.next().setTrashed(true);   // ganti versi lama bulan yang sama
    const f = folder.createFile(pdf);

    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;padding:24px;max-width:480px">' +
      '<h2 style="color:#2f5d3a">Laporan PDF siap ✓</h2>' +
      '<p>' + list.length + ' aduan untuk bulan ' + bulan + '.</p>' +
      '<p><a style="display:inline-block;background:#2f5d3a;color:#fff;padding:12px 18px;border-radius:8px;text-decoration:none;font-weight:bold" target="_blank" href="' + f.getUrl() + '">Buka PDF</a></p>' +
      '<p style="color:#666;font-size:13px">Disimpan dalam Google Drive: Aduan Kerosakan LP2 - Gambar / Laporan PDF</p></div>'
    ).setTitle('Laporan PDF');
  } catch (err) {
    return HtmlService.createHtmlOutput('<p style="font-family:Arial;padding:24px">Gagal buat PDF: ' + esc(String(err)) + '</p>');
  }
}

function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

function laporanHTML(list, bulan, now) {
  const BM = ['Januari','Februari','Mac','April','Mei','Jun','Julai','Ogos','September','Oktober','November','Disember'];
  const label = BM[+bulan.slice(5) - 1] + ' ' + bulan.slice(0, 4);
  const G = '#2f5d3a', L = '#d9dfd8';
  const td = 'border:1px solid ' + L + ';padding:5px 7px;vertical-align:top;font-size:9.5pt;';
  const th = td + 'background:#e8efe9;font-size:8.5pt;text-align:left;';

  // Ringkasan
  let totalItems = 0, kecemasan = 0, selesai = 0, itemSiap = 0;
  const jenis = {};
  list.forEach(a => {
    totalItems += a.items.length;
    if (a.pri === 'KECEMASAN') kecemasan++;
    if (a.status === 'Selesai') selesai++;
    itemSiap += a.itemSt.filter(x => x.status === 'Selesai').length;
    a.items.forEach(t => {
      const k = t.replace(/^\d+\.\s*/, '').split(' - ')[0].trim();
      (jenis[k] = jenis[k] || []).push(a.rumah);
    });
  });
  const tile = (num, lbl) => '<td style="border:1px solid ' + L + ';padding:8px 10px;width:25%"><div style="font-size:18pt;font-weight:bold">' + num + '</div><div style="font-size:8pt;color:#5f6b63">' + lbl + '</div></td>';

  let h = '<html><head><meta charset="utf-8"></head><body style="font-family:Arial,sans-serif;color:#1d2420;font-size:10pt">';
  h += '<div style="background:' + G + ';color:#fff;padding:12px 16px"><div style="font-size:8pt;letter-spacing:2px">LABOUR QUARTERS · LEEPANG 2 (LP2)</div>' +
       '<div style="font-size:18pt;font-weight:bold;margin-top:3px">Laporan Aduan Kerosakan Rumah</div><div style="font-size:9.5pt">Bulan ' + label + '</div></div>';
  h += '<p style="font-size:9pt;color:#4a564e">Tarikh laporan: <b>' + Utilities.formatDate(now, TZ, 'dd/MM/yyyy HH:mm') + '</b> &nbsp;·&nbsp; Sumber: <b>Borang Aduan QR</b></p>';
  h += '<h3 style="color:' + G + ';font-size:11pt;margin:14px 0 6px">RINGKASAN</h3><table style="border-collapse:collapse;width:100%"><tr>' +
       tile(list.length, 'ADUAN') + tile(totalItems, 'ITEM KEROSAKAN') + tile(itemSiap + ' / ' + totalItems, 'ITEM SELESAI') + tile(selesai + ' / ' + list.length, 'ADUAN SELESAI') + '</tr></table>';

  if (!list.length) return h + '<p>Tiada aduan untuk bulan ini.</p></body></html>';

  h += '<h3 style="color:' + G + ';font-size:11pt;margin:14px 0 6px">KEROSAKAN MENGIKUT JENIS</h3><table style="border-collapse:collapse;width:100%">' +
       '<tr><th style="' + th + '">Jenis</th><th style="' + th + 'width:50px">Bil.</th><th style="' + th + '">Rumah</th></tr>';
  Object.keys(jenis).sort((a, b) => jenis[b].length - jenis[a].length).forEach(k => {
    h += '<tr><td style="' + td + '">' + esc(k) + '</td><td style="' + td + 'text-align:center">' + jenis[k].length + '</td><td style="' + td + '">' + esc([...new Set(jenis[k])].join(', ')) + '</td></tr>';
  });
  h += '</table>';

  h += '<h3 style="color:' + G + ';font-size:11pt;margin:16px 0 6px">BUTIRAN ADUAN</h3>';
  list.forEach(a => {
    const pc = a.status === 'Selesai' ? '#2f5d3a' : (a.pri === 'KECEMASAN' ? '#c2561a' : '#a34712');
    h += '<div style="border:1px solid #cfd8d0;margin:10px 0;page-break-inside:avoid">';
    h += '<table style="width:100%;border-collapse:collapse;background:#f3f5f1"><tr><td style="padding:7px 10px;font-size:12pt;font-weight:bold">' + esc(a.rumah) + '</td>' +
         '<td style="padding:7px 10px;text-align:right;font-size:8.5pt;font-weight:bold;color:' + pc + '">' + esc(String(a.status).toUpperCase()) + ' · ' + esc(a.pri) + '</td></tr></table>';
    h += '<table style="width:100%;font-size:9pt;margin:6px 4px"><tr><td style="color:#5f6b63;width:70px">ID</td><td>' + esc(a.id) + '</td><td style="color:#5f6b63;width:60px">Masa</td><td>' + esc(a.masa) + '</td></tr>' +
         '<tr><td style="color:#5f6b63">Pelapor</td><td>' + esc(a.nama) + '</td><td style="color:#5f6b63">Telefon</td><td>' + esc(a.tel) + '</td></tr></table>';
    h += '<table style="border-collapse:collapse;width:96%;margin:4px 2%"><tr><th style="' + th + 'width:24px">#</th><th style="' + th + 'width:140px">Item</th><th style="' + th + '">Keterangan pelapor</th><th style="' + th + 'width:95px">Status</th><th style="' + th + '">Tindakan</th></tr>';
    a.items.forEach((t, i) => {
      const s = t.replace(/^\d+\.\s*/, ''), cut = s.indexOf(' - ');
      const st = a.itemSt[i] || { status: 'Baru', tindakan: '' };
      const sc = st.status === 'Selesai' ? 'color:#2f7a45;background:#e2f1e6;' : (st.status === 'Dalam Tindakan' ? 'color:#8a6d0f;background:#f8f0d6;' : 'color:#a34712;background:#fbeee5;');
      h += '<tr><td style="' + td + 'text-align:center">' + (i + 1) + '</td><td style="' + td + '">' + esc(cut > -1 ? s.slice(0, cut) : s) + '</td><td style="' + td + '">' + esc(cut > -1 ? s.slice(cut + 3) : '–') + '</td>' +
           '<td style="' + td + sc + 'font-weight:bold;font-size:8.5pt">' + esc(st.status) + '</td><td style="' + td + '">' + esc(st.tindakan || '') + '</td></tr>';
    });
    h += '</table>';
    if (a.catatan) h += '<p style="margin:6px 10px;font-size:9pt"><b>Catatan:</b> ' + esc(a.catatan) + '</p>';
    if (a.tindakan) h += '<p style="margin:6px 10px;font-size:9pt"><b>Tindakan penyelia:</b> ' + esc(a.tindakan) + '</p>';
    if (a.links.length) {
      h += '<table style="margin:6px 8px 10px"><tr>';
      a.links.forEach((u, i) => {
        const m = u.match(/\/d\/([\w-]+)/);
        let img = '';
        try { if (m) img = '<img src="data:image/jpeg;base64,' + Utilities.base64Encode(DriveApp.getFileById(m[1]).getBlob().getBytes()) + '" style="width:150px;height:112px;object-fit:cover">'; } catch (e) {}
        h += '<td style="padding:3px;vertical-align:top">' + (img || '[Gambar ' + (i + 1) + ']') + '</td>';
        if (i % 4 === 3) h += '</tr><tr>';
      });
      h += '</tr></table>';
    } else {
      h += '<p style="margin:6px 10px 10px;font-size:9pt;color:#5f6b63">Gambar: tiada dilampirkan</p>';
    }
    h += '</div>';
  });
  h += '<table style="width:100%;margin-top:36px;font-size:9pt"><tr>' +
       ['Disediakan oleh', 'Disemak oleh', 'Diluluskan oleh'].map(t => '<td style="width:33%;padding-right:24px"><div style="border-top:1px solid #1d2420;padding-top:5px">' + t + '</div></td>').join('') + '</tr></table>';
  return h + '</body></html>';
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Jalankan sekali dari editor untuk beri kebenaran (Run > setup)
function setup() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  sh.setName('Aduan');
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#2f5d3a').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  const status = SpreadsheetApp.newDataValidation().requireValueInList(['Baru', 'Dalam Tindakan', 'Selesai'], true).build();
  sh.getRange('L2:L').setDataValidation(status);
  sh.setColumnWidth(9, 320);
  sh.setColumnWidth(13, 260);
  DriveApp.getFolderById(FOLDER_ID).getName();
}
