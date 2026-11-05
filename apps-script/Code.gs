/**
 * Aduan Kerosakan LP2 - simpan aduan ke Google Sheet dan gambar ke Google Drive.
 * Tampal kod ini dalam Apps Script, kemudian Deploy sebagai Web app.
 */
const SHEET_ID  = '1_ZgsW-EDqbtIH2MK4MzrCu_GjVR8QjVQLAOxk35h4ZQ'; // Google Sheet "Aduan Kerosakan LP2"
const FOLDER_ID = '1FntAmCKj7ZXeID7lCEnxrWQNKSNRRzNT';            // Folder "Aduan Kerosakan LP2 - Gambar"

const HEADERS = ['Masa Hantar','ID Aduan','No. Rumah','Tarikh','Nama Pelapor','No. Telefon',
  'Keutamaan','Bil. Kerosakan','Kerosakan','Catatan','Gambar','Status','Tindakan / Catatan Penyelia'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
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
      (d.items || []).length, items, d.catatan || '', '', 'Baru', '']);
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

// Buka link Web app dalam browser untuk semak ia berfungsi
function doGet() {
  return out({ ok: true, msg: 'Aduan Kerosakan LP2 sedia menerima aduan' });
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
