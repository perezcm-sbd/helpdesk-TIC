// Helpdesk TIC · Flor de Maig — backend (Google Apps Script)
// Google Sheet buit > Extensions > Apps Script > enganxa-ho.
// Propietats de l'script (Configuració del projecte):
//   CLIENT_ID      = el Client ID OAuth de Google (el mateix que a index.html)
//   TIC_EMAILS     = correus de l'equip TIC separats per comes
//   ANTHROPIC_API_KEY = (opcional) clau de l'API de Claude; si hi és, s'activa el xat d'ajuda
// Desplega com a aplicació web: executa com a tu, accés "Qualsevol" (el codi verifica el compte @flordemaig.cat).
// Si vens de la versió anterior, esborra la pestanya "Tickets" del Sheet.
const DOMAIN = 'flordemaig.cat';
const COLS = ['id','data','professor','email','assumpte','espai','categoria','prioritat','descripcio','estat','assignat','fils','actualitzat'];
const P = PropertiesService.getScriptProperties();
const SYS = "Ets l'assistent d'ajuda TIC del centre d'educació especial CEE Flor de Maig. Respon en català, de manera breu, clara i amable, amb passos senzills per resoldre problemes informàtics habituals (ordinador, projector, Wi-Fi, impressora, comptes). No demanis ni acceptis contrasenyes. Si no es pot resoldre amb pocs passos, o hi ha risc, recomana crear un ticket amb el botó 'Nou ticket'.";

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Tickets');
  if (!sh) { sh = ss.insertSheet('Tickets'); sh.appendRow(COLS); sh.setFrozenRows(1); }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
let WHY = '';
function auth_(token) {
  if (!token) { WHY = 'sense token'; return null; }
  const cache = CacheService.getScriptCache();
  const key = 't' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token)).slice(0, 40);
  let u = cache.get(key);
  if (u) return JSON.parse(u);
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) { WHY = 'tokeninfo ' + r.getResponseCode(); return null; }
  const j = JSON.parse(r.getContentText());
  const email = String(j.email || '').toLowerCase();
  if (j.aud !== P.getProperty('CLIENT_ID')) { WHY = 'aud diferent: ' + j.aud; return null; }
  if (j.email_verified !== 'true') { WHY = 'correu no verificat'; return null; }
  if (j.hd !== DOMAIN || !email.endsWith('@' + DOMAIN)) { WHY = 'domini: ' + j.hd; return null; }
  const tics = (P.getProperty('TIC_EMAILS') || '').toLowerCase().split(',').map(s => s.trim());
  const user = { email, name: j.name || email, tic: tics.includes(email) };
  cache.put(key, JSON.stringify(user), 300);
  return user;
}
function doGet() { return out_({ ok: true, msg: 'Helpdesk TIC actiu' }); }
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const b = JSON.parse(e.postData.contents), me = auth_(b.token);
    if (!me) return out_({ ok: false, error: 'auth' });
    const sh = sheet_(), now = new Date().toISOString();
    if (b.action === 'list') {
      const rows = sh.getDataRange().getValues().slice(1).map(r => {
        const t = Object.fromEntries(COLS.map((c, i) => [c, r[i]]));
        try { t.fils = JSON.parse(t.fils || '[]'); } catch (x) { t.fils = []; }
        return t;
      }).filter(t => me.tic || t.email === me.email).reverse();
      return out_({ ok: true, me, tickets: rows, chat: !!P.getProperty('ANTHROPIC_API_KEY') });
    }
    if (b.action === 'create') {
      const id = Number(P.getProperty('N') || 0) + 1; P.setProperty('N', String(id));
      sh.appendRow([id, now, me.name, me.email, b.assumpte, b.espai, b.categoria, b.prioritat, b.descripcio, 'Obert', '', '[]', now]);
      return out_({ ok: true, id });
    }
    if (b.action === 'chat') return out_(chat_(b.messages));
    const data = sh.getDataRange().getValues(), row = data.findIndex((r, i) => i > 0 && r[0] == b.id) + 1;
    if (row < 2) return out_({ ok: false, error: 'Ticket no trobat' });
    if (!me.tic && data[row - 1][COLS.indexOf('email')] !== me.email) return out_({ ok: false, error: 'Sense permís' });
    if (b.action === 'comment') {
      const c = COLS.indexOf('fils') + 1, fils = JSON.parse(data[row - 1][c - 1] || '[]');
      fils.push({ autor: me.name, tic: me.tic, text: String(b.text).slice(0, 2000), data: now });
      sh.getRange(row, c).setValue(JSON.stringify(fils));
      sh.getRange(row, COLS.length).setValue(now);
      return out_({ ok: true });
    }
    if (b.action === 'update') {
      if (!me.tic) return out_({ ok: false, error: 'Sense permís' });
      ['estat', 'prioritat', 'assignat'].forEach(k => { if (b[k] !== undefined) sh.getRange(row, COLS.indexOf(k) + 1).setValue(b[k]); });
      sh.getRange(row, COLS.length).setValue(now);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'Acció desconeguda' });
  } finally { lock.releaseLock(); }
}
function chat_(msgs) {
  const key = P.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, error: 'Xat no activat' };
  const m = (msgs || []).slice(-10).map(x => ({ role: x.role === 'assistant' ? 'assistant' : 'user', content: String(x.content).slice(0, 1500) }));
  const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: 'claude-haiku-5-5', max_tokens: 500, system: SYS, messages: m })
  });
  if (r.getResponseCode() !== 200) return { ok: false, error: 'El xat no està disponible ara mateix.' };
  return { ok: true, text: JSON.parse(r.getContentText()).content[0].text };
}
