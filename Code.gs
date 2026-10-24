// ===== غيّر الرمزين =====
const STUDENT_CODE = 'READ2026';
const ADMIN_PIN = 'CHANGE-ME-9999';

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('subs');
  if (!s) { s = ss.insertSheet('subs'); s.appendRow(['id','student','cls','level','status','date','books']); }
  return s;
}
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const safe_ = v => { v = String(v == null ? '' : v).slice(0, 5000); return /^[=+\-@]/.test(v) ? "'" + v : v; };
function rows_(s) {
  const v = s.getDataRange().getValues(); v.shift();
  return v.filter(r => r[0] !== '').map(r => {
    let books = []; try { books = JSON.parse(r[6]); } catch (e) {}
    return { id: Number(r[0]), student: String(r[1]), cls: String(r[2]), level: Number(r[3]), status: String(r[4]), date: String(r[5]), books: books };
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const r = JSON.parse(e.postData.contents), s = sheet_();
    if (r.action === 'submit') {
      if (r.code !== STUDENT_CODE) return out_({ error: 'bad_code' });
      const p = r.p || {}, lv = Number(p.level);
      if (!p.student || !p.cls || !(lv >= 1 && lv <= 5) || !Array.isArray(p.books)) return out_({ error: 'bad_data' });
      if (rows_(s).some(x => x.student === String(p.student) && x.cls === String(p.cls) && x.level === lv && x.status !== 'rejected')) return out_({ error: 'exists' });
      const books = p.books.slice(0, 5).map(b => { const o = {}; ['title','author','pub','year','ed','pages','sum'].forEach(k => o[k] = String((b && b[k]) || '').slice(0, 5000)); return o; });
      s.appendRow([Number(p.id) || Date.now(), safe_(p.student), safe_(p.cls), lv, 'pending', safe_(p.date), JSON.stringify(books)]);
      return out_({ ok: 1 });
    }
    if (r.action === 'mine') {
      if (r.code !== STUDENT_CODE) return out_({ error: 'bad_code' });
      return out_({ rows: rows_(s).filter(x => x.student === String(r.student) && x.cls === String(r.cls)) });
    }
    if (r.pin !== ADMIN_PIN) return out_({ error: 'bad_pin' });
    if (r.action === 'list') return out_({ rows: rows_(s) });
    if (r.action === 'set') {
      if (['pending','approved','rejected'].indexOf(r.status) < 0) return out_({ error: 'bad_status' });
      const v = s.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) if (Number(v[i][0]) === Number(r.id)) { s.getRange(i + 1, 5).setValue(r.status); return out_({ ok: 1 }); }
      return out_({ error: 'not_found' });
    }
    return out_({ error: 'bad_action' });
  } finally { lock.releaseLock(); }
}
