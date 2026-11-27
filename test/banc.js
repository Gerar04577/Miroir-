// Banc d'essai du moteur Miroir : faux OneDrive + faux Google Drive, vrai engine.js
'use strict';
const assert = require('assert');
const path = require('path');
const crypto = require('crypto');
const sha1 = (t) => crypto.createHash('sha1').update(Buffer.from(t, 'latin1')).digest('hex');
const L1 = (t) => new Blob([Uint8Array.from(t, (c) => c.charCodeAt(0) & 255)]);
// contenu réaliste selon le type de fichier (sinon le contrôle d'intégrité le refuserait, à raison)
function wrap(name, c) {
  if (c == null) return c;
  const ext = (name.split('.').pop() || '').toLowerCase();
  const pad = (n) => (c.length < n ? ' '.repeat(n - c.length) : '');   // une vraie photo n'est jamais minuscule
  if (ext === 'jpg' || ext === 'jpeg') return '\xFF\xD8\xFF\xE0' + c + pad(21000) + '\xFF\xD9';
  if (ext === 'pdf') return '%PDF-1.4\n' + c + pad(1200) + '\n%%EOF\n';
  if (ext === 'docx') return 'PK\x03\x04' + c + pad(2200) + 'PK\x05\x06' + '\0'.repeat(18);
  if (ext === 'mp4' || ext === 'mov' || ext === 'heic') return '\0\0\0\x18ftypisom' + c + pad(21000);
  return c;
}
const { createEngine, PauseError } = require(path.join(__dirname, '..', 'engine.js'));

let checks = 0;
function ok(cond, msg) { assert.ok(cond, msg); checks++; console.log('  ✔ ' + msg); }

/* ---------------- Faux OneDrive ---------------- */
class FakeOneDrive {
  constructor() { this.items = new Map(); this.n = 0; this.driveId = 'DRV1'; this.failList = new Set(); this.calls = 0; this.add('ROOT', null, 'root', true); }
  add(id, parent, name, folder, content, raw) {
    const it = { id, parent, name, folder, content: content == null ? null : (raw ? content : wrap(name, content)), ctag: 'c' + (++this.n), mod: new Date().toISOString() };
    this.items.set(id, it); return it;
  }
  setContent(id, c, raw) { const it = this.items.get(id); it.content = raw ? c : wrap(it.name, c); it.ctag = 'c' + (++this.n); it.mod = new Date().toISOString(); }
  rename(id, name) { this.items.get(id).name = name; }
  move(id, parent) { this.items.get(id).parent = parent; }
  remove(id) { for (const k of this.kids(id)) this.remove(k.id); this.items.delete(id); }
  kids(id) { return [...this.items.values()].filter((i) => i.parent === id).sort((a, b) => a.id.localeCompare(b.id)); }
  pathOf(id, rootId) { const parts = []; let c = this.items.get(id); while (c && c.id !== rootId) { parts.unshift(c.name); c = this.items.get(c.parent); } return parts.join('/'); }
  tree(rootId) {
    const out = {}; const st = [rootId];
    while (st.length) { const c = st.pop(); for (const k of this.kids(c)) { if (k.pkg) continue; out[this.pathOf(k.id, rootId)] = k.folder ? '<dir>' : k.content; if (k.folder) st.push(k.id); } }
    return out;
  }
}

/* ---------------- Faux Google Drive ---------------- */
class FakeDrive {
  constructor() { this.files = new Map(); this.n = 0; this.token = 'G1'; this.sessions = new Map(); this.uploadsCount = 0; this.files.set('root', { id: 'root', name: 'My Drive', parents: [], folder: true, trashed: false }); }
  create(meta, content) {
    const id = 'g' + (++this.n);
    this.files.set(id, { id, name: meta.name, parents: meta.parents || ['root'], folder: meta.mimeType === 'application/vnd.google-apps.folder', trashed: false, content: content === undefined ? null : content, appProperties: meta.appProperties });
    return id;
  }
  tree(rootId) {
    const out = {}; const st = [rootId];
    const pathOf = (f) => { const parts = []; let c = f; while (c && c.id !== rootId) { parts.unshift(c.name); c = this.files.get(c.parents[0]); } return parts.join('/'); };
    while (st.length) {
      const c = st.pop();
      for (const f of this.files.values()) if (!f.trashed && f.parents.includes(c)) { const p = pathOf(f); if (out[p] !== undefined) out[p + ' (DOUBLON)'] = 1; out[p] = f.folder ? '<dir>' : f.content; if (f.folder) st.push(f.id); }
    }
    return out;
  }
}

function resp(status, body, headers = {}, bytes = false) {
  const h = new Map(Object.entries(headers).map(([k, v]) => [k.toLowerCase(), v]));
  const text = body === undefined || body === null ? '' : (typeof body === 'string' ? body : JSON.stringify(body));
  return {
    status, ok: status >= 200 && status < 300,
    headers: { get: (k) => h.get(k.toLowerCase()) || null },
    text: async () => bytes ? Buffer.from(text, 'latin1').toString('utf8') : text,
    json: async () => JSON.parse(text),
    blob: async () => bytes ? L1(text) : new Blob([text])
  };
}
async function bodyText(b) { if (b == null) return ''; if (typeof b === 'string') return b; return Buffer.from(await b.arrayBuffer()).toString('latin1'); }

function makeWorld() {
  const od = new FakeOneDrive();
  const gd = new FakeDrive();
  const world = { od, gd, pageSize: 3, gExpireAfterCalls: Infinity, gCalls: 0, crashOnIndexSave: false, indexSaves: 0, crashAfter: 0 };

  async function fetch(url, opt = {}) {
    const u = new URL(url);
    const method = (opt.method || 'GET').toUpperCase();
    const auth = (opt.headers && opt.headers.Authorization) || '';
    /* --- OneDrive --- */
    if (u.host === 'graph.microsoft.com') {
      od.calls++;
      assert.strictEqual(method, 'GET', 'jamais d\'écriture vers OneDrive');
      let m = /\/drives\/([^/]+)\/items\/([^/]+)\/children$/.exec(decodeURIComponent(u.pathname));
      if (m) {
        const id = m[2];
        if (od.failList.has(id)) return resp(403, 'interdit');
        const all = od.kids(id);
        const skip = Number(u.searchParams.get('skip') || 0);
        const page = all.slice(skip, skip + world.pageSize).map((k) => {
          const base = { id: k.id, name: k.name, lastModifiedDateTime: k.mod };
          if (k.pkg) return Object.assign(base, { package: { type: 'oneNote' } });
          if (k.folder) return Object.assign(base, { folder: { childCount: 0 }, cTag: 'f' + Math.random() });
          return Object.assign(base, { file: {}, size: k.content.length, cTag: k.ctag, eTag: 'e' + k.ctag });
        });
        const body = { value: page };
        if (skip + world.pageSize < all.length) body['@odata.nextLink'] = `https://graph.microsoft.com/v1.0/drives/${m[1]}/items/${id}/children?skip=${skip + world.pageSize}`;
        return resp(200, body);
      }
      m = /\/drives\/([^/]+)\/items\/([^/]+)\/content$/.exec(decodeURIComponent(u.pathname));
      if (m) {   // comportement réel : /content renvoie le fichier (la redirection est suivie par fetch)
        if (world.contentBroken) throw new TypeError('Load failed');
        const it = od.items.get(m[2]); if (!it) return resp(404, 'nope');
        return resp(200, it.content, {}, true);
      }
      m = /\/drives\/([^/]+)\/items\/([^/]+)$/.exec(decodeURIComponent(u.pathname));
      if (m) {   // comportement réel : un $select fait disparaître le lien de téléchargement
        const it = od.items.get(m[2]); if (!it) return resp(404, 'nope');
        const body = { id: it.id, name: it.name };
        if (!u.searchParams.has('$select')) body['@microsoft.graph.downloadUrl'] = 'https://dl.files.1drv.com/' + it.id + '?v=' + it.ctag;
        return resp(200, body);
      }
      return resp(400, 'graph route inconnue ' + u.pathname);
    }
    if (u.host === 'dl.files.1drv.com') {
      assert.ok(!auth, 'téléchargement sans en-tête Authorization');
      const it = od.items.get(u.pathname.slice(1));
      return resp(200, it.content, {}, true);
    }
    /* --- Google --- */
    if (u.host === 'www.googleapis.com') {
      world.gCalls++;
      if (auth !== 'Bearer ' + gd.token || world.gCalls > world.gExpireAfterCalls) return resp(401, 'expired');
      const p = u.pathname;
      if (p === '/drive/v3/files' && method === 'GET') {
        const qs = u.searchParams.get('q');
        const nm = /name='((?:[^'\\]|\\.)*)'/.exec(qs);
        const name = nm ? nm[1].replace(/\\'/g, "'").replace(/\\\\/g, '\\') : null;
        const parent = /'([^']+)' in parents/.exec(qs)[1];
        const folderOnly = /mimeType=/.test(qs);
        const all = [...gd.files.values()].filter((f) => !f.trashed && (name === null || f.name === name) && f.parents.includes(parent) && (!folderOnly || f.folder));
        const ps = 4, start = Number(u.searchParams.get('pageToken') || 0);   // petite pagination pour tester nextPageToken
        const files = all.slice(start, start + ps).map((f) => ({ id: f.id, name: f.name, mimeType: f.folder ? 'application/vnd.google-apps.folder' : 'application/octet-stream', size: f.folder ? undefined : String((f.content || '').length), sha1Checksum: f.folder ? undefined : sha1(f.content || ''), appProperties: f.appProperties }));
        const body = { files };
        if (start + ps < all.length) body.nextPageToken = String(start + ps);
        return resp(200, body);
      }
      if (p === '/drive/v3/files' && method === 'POST') {
        const meta = JSON.parse(opt.body);
        return resp(200, { id: gd.create(meta) });
      }
      let m = /^\/drive\/v3\/files\/([^/]+)$/.exec(p);
      if (m) {
        const f = gd.files.get(decodeURIComponent(m[1]));
        if (!f) return resp(404, 'not found');
        if (method === 'GET') {
          if (u.searchParams.get('alt') === 'media') return resp(200, f.content, {}, true);
          return resp(200, { id: f.id, trashed: f.trashed, parents: f.parents, headRevisionId: 'rev-' + f.id + '-' + (f.rev || 0) });
        }
        if (method === 'PATCH') {
          const body = opt.body ? JSON.parse(opt.body) : {};
          const add = u.searchParams.get('addParents'), rem = u.searchParams.get('removeParents');
          if (rem) f.parents = f.parents.filter((x) => !rem.split(',').includes(x));
          if (add) f.parents.push(add);
          if (body.name) f.name = body.name;
          if (body.trashed) f.trashed = true;
          return resp(200, { id: f.id });
        }
      }
      m = /^\/drive\/v3\/files\/([^/]+)\/revisions\/([^/]+)$/.exec(p);
      if (m && method === 'PATCH') {
        const f = gd.files.get(decodeURIComponent(m[1])); if (!f) return resp(404, 'not found');
        const body = JSON.parse(opt.body);
        (f.kept = f.kept || []).push({ rev: decodeURIComponent(m[2]), content: f.content, keepForever: body.keepForever });
        return resp(200, { id: m[2], keepForever: body.keepForever });
      }
      if (p === '/upload/drive/v3/files' && method === 'POST') {
        const type = u.searchParams.get('uploadType');
        if (type === 'multipart') {
          const raw = await bodyText(opt.body);
          const boundary = /boundary=(.+)$/.exec(opt.headers['Content-Type'])[1];
          const parts = raw.split('--' + boundary);
          const meta = JSON.parse(Buffer.from(parts[1].split('\r\n\r\n')[1].trim(), 'latin1').toString('utf8'));
          const content = parts[2].split('\r\n\r\n').slice(1).join('\r\n\r\n').replace(/\r\n$/, '');
          if (meta.name.startsWith('_index-') && world.crashOnIndexSave && ++world.indexSaves > world.crashAfter) throw new Error('CRASH simulé');
          if (!meta.name.startsWith('_index-')) gd.uploadsCount++;
          return resp(200, { id: gd.create(meta, content) });
        }
        if (type === 'resumable') {
          const meta = JSON.parse(opt.body);
          const sid = 'S' + (++gd.n);
          gd.sessions.set(sid, { meta, total: Number(opt.headers['X-Upload-Content-Length']), parts: [], existing: null });
          return resp(200, '', { Location: 'https://www.googleapis.com/upload/session/' + sid });
        }
      }
      m = /^\/upload\/drive\/v3\/files\/([^/]+)$/.exec(p);
      if (m && method === 'PATCH') {
        const f = gd.files.get(decodeURIComponent(m[1]));
        if (!f) return resp(404, 'not found');
        const type = u.searchParams.get('uploadType');
        if (type === 'media') {
          if (f.name.startsWith('_index-') && world.crashOnIndexSave && ++world.indexSaves > world.crashAfter) throw new Error('CRASH simulé');
          f.content = await bodyText(opt.body); f.rev = (f.rev || 0) + 1; if (!f.name.startsWith('_index-')) gd.uploadsCount++; return resp(200, { id: f.id });
        }
        if (type === 'resumable') {
          const sid = 'S' + (++gd.n);
          gd.sessions.set(sid, { meta: {}, total: Number(opt.headers['X-Upload-Content-Length']), parts: [], existing: f.id });
          return resp(200, '', { Location: 'https://www.googleapis.com/upload/session/' + sid });
        }
      }
      return resp(400, 'google route inconnue ' + method + ' ' + p);
    }
    throw new Error('hôte inconnu ' + u.host);
  }

  async function xhrPut(url, headers, blob) {
    if (world.failNextChunk && blob) { world.failNextChunk = false; throw new Error('réseau coupé'); }
    const sid = url.split('/').pop();
    const s = gd.sessions.get(sid);
    if (!s) return { status: 404 };
    const cr = headers['Content-Range'];
    if (/\*\//.test(cr)) { const got = s.parts.join('').length; return { status: 308, range: got ? 'bytes=0-' + (got - 1) : null }; }
    const [, a, b, tot] = /bytes (\d+)-(\d+)\/(\d+)/.exec(cr).map(Number);
    assert.strictEqual(a, s.parts.join('').length, 'morceau envoyé au bon décalage');
    const txt = Buffer.from(await blob.arrayBuffer()).toString('latin1');
    assert.strictEqual(txt.length, b - a + 1, 'taille du morceau');
    if (b + 1 < tot) assert.strictEqual((b - a + 1) % (256 * 1024), 0, 'morceau multiple de 256 Kio');
    s.parts.push(txt);
    if (b + 1 < tot) return { status: 308, range: 'bytes=0-' + b };
    const content = s.parts.join('');
    gd.uploadsCount++;
    if (s.existing) { const f = gd.files.get(s.existing); f.content = content; f.rev = (f.rev || 0) + 1; return { status: 200, json: { id: s.existing } }; }
    return { status: 200, json: { id: gd.create(s.meta, content) } };
  }

  const journalStore = new Map();
  const journal = {
    load: (id) => JSON.parse(journalStore.get(id) || '[]'),
    add: (id, e) => { const a = JSON.parse(journalStore.get(id) || '[]'); a.push(e); journalStore.set(id, JSON.stringify(a)); },
    clear: (id) => journalStore.delete(id)
  };
  world.journalStore = journalStore;
  world.engine = (extra = {}) => createEngine(Object.assign({
    fetch, xhrPut,
    msToken: async () => 'M1',
    gToken: async () => { if (world.gCalls >= world.gExpireAfterCalls) throw new PauseError('google'); return gd.token; },
    gTokenQuiet: () => gd.token,
    journal, sleep: async () => {}, log: () => {}
  }, extra));
  return world;
}

function sameTree(world, src, gRoot, label) {
  const a = world.od.tree(src.id), b = world.gd.tree(gRoot);
  const ka = Object.keys(a).sort(), kb = Object.keys(b).sort();
  if (JSON.stringify(ka) !== JSON.stringify(kb)) {
    console.log('DIFF', ka.length, kb.length, ka.filter(k=>!kb.includes(k)), kb.filter(k=>!ka.includes(k)));
  }
  ok(JSON.stringify(ka) === JSON.stringify(kb), label + ' — mêmes chemins (' + ka.length + ')');
  ok(ka.every((k) => a[k] === b[k]), label + ' — mêmes contenus');
}
function mirrorRoot(world) {
  const top = [...world.gd.files.values()].find((f) => f.name === 'Miroir OneDrive' && !f.trashed);
  return [...world.gd.files.values()].find((f) => f.parents.includes(top.id) && f.folder && f.name === 'Immobilier' && !f.trashed).id;
}
const allSusp = (r) => (r.suspects || []).concat(r.sizeSuspects || []);
const big = (ch, mb) => ch.repeat(Math.floor(mb * 1024 * 1024));

function buildTree(od) {
  od.add('IMMO', 'ROOT', 'Immobilier', true);
  od.add('NIMY', 'IMMO', 'Nimy', true);
  od.add('BICHE', 'IMMO', 'Biche', true);
  od.add('S1', 'NIMY', 'Studio 1', true);
  od.add('S2', 'NIMY', 'Studio 2', true);
  for (let i = 1; i <= 7; i++) od.add('P' + i, 'S1', `photo ${i}.jpg`, false, 'photo-' + i);
  od.add('BAIL', 'S1', "Bail d'Aurélie.pdf", false, 'bail v1');
  od.add('GROS', 'S2', 'video EDLE.mp4', false, big('x', 20.3));
  od.add('B1', 'BICHE', 'STUDIO 1', true);
  od.add('EDL', 'B1', 'EDLE.docx', false, 'edl');
  od.add('VIDE', 'BICHE', 'Dossier vide', true);
  const nb = od.add('NOTE', 'IMMO', 'Bloc-notes', false, 'x'); nb.pkg = true;
  od.add('AUTRE', 'ROOT', 'Hors périmètre.txt', false, 'non');
}

(async () => {
  const src = { driveId: 'DRV1', id: 'IMMO', name: 'Immobilier' };

  console.log('\n1) Première copie complète');
  let W = makeWorld(); buildTree(W.od);
  let r = await W.engine().run(src, {});
  ok(!r.paused && r.failed.length === 0, 'aucune erreur, pas de pause');
  ok(r.skipped.includes('Bloc-notes'), 'bloc-notes OneNote ignoré et signalé');
  sameTree(W, src, mirrorRoot(W), 'copie initiale');
  ok(!Object.keys(W.gd.tree(mirrorRoot(W))).some((k) => k.includes('Hors')), 'rien hors du dossier choisi');
  ok(W.journalStore.size === 0, 'journal vidé après enregistrement de l\'index');

  console.log('\n2) Deuxième passage sans changement');
  const before = W.gd.uploadsCount;
  r = await W.engine().run(src, {});
  const c = r.plan.counts;
  ok(c.nouveauxFichiers === 0 && c.modifies === 0 && c.deplaces === 0 && c.supprimes === 0 && c.nouveauxDossiers === 0, 'plan vide');
  ok(W.gd.uploadsCount === before, 'aucun fichier renvoyé (seul l\'index est réécrit)');

  console.log('\n3) Modifs, renommages, déplacements, ajouts, suppressions');
  W.od.setContent('BAIL', 'bail v2 avenant');
  W.od.rename('P1', 'photo 1 renommée.jpg');
  W.od.move('P2', 'B1');
  W.od.rename('S2', 'Studio 2 (Appartement)');
  W.od.move('S2', 'BICHE');
  W.od.add('NEW', 'S1', 'nouveau.pdf', false, 'neuf');
  W.od.add('NEWD', 'IMMO', 'Havré', true); W.od.add('NEWF', 'NEWD', 'bail havré.pdf', false, 'h');
  W.od.remove('P3');
  W.od.remove('VIDE');
  W.od.setContent('GROS', big('y', 6));
  r = await W.engine().run(src, {});
  ok(r.failed.length === 0, 'aucune erreur');
  ok(r.plan.counts.supprimes === 2 && r.plan.topDels.length === 2, '2 suppressions détectées');
  sameTree(W, src, mirrorRoot(W), 'après changements');
  const bin = [...W.gd.files.values()].find((f) => f.name === '_Supprimés');
  const binTree = W.gd.tree(bin.id);
  ok(Object.keys(binTree).some((k) => k.endsWith('/photo 3.jpg')) && Object.keys(binTree).some((k) => k.endsWith('/Dossier vide')), 'fichier et dossier supprimés rangés dans _Supprimés/<date>');

  console.log('\n4) Lecture OneDrive en échec → aucune suppression');
  W.od.remove('P4');
  W.od.failList.add('B1');
  r = await W.engine().run(src, {});
  ok(r.walkErrors.length === 1, 'erreur de lecture signalée');
  ok(r.plan.dels.length === 0, 'aucune suppression planifiée');
  ok(Object.keys(W.gd.tree(mirrorRoot(W))).some((k) => k.endsWith('photo 4.jpg')), 'photo 4 toujours présente côté Google');
  ok(Object.keys(W.gd.tree(mirrorRoot(W))).some((k) => k.endsWith('EDLE.docx')), 'contenu du dossier illisible NON supprimé');
  W.od.failList.clear();
  r = await W.engine().run(src, {});
  sameTree(W, src, mirrorRoot(W), 'lecture rétablie');

  console.log('\n5) Jeton Google expiré en cours de copie, puis reprise');
  W = makeWorld(); buildTree(W.od);
  for (let i = 0; i < 40; i++) W.od.add('X' + i, 'S1', `scan ${i}.pdf`, false, 'scan' + i);
  W.gExpireAfterCalls = 30;
  r = await W.engine().run(src, {});
  ok(r.paused === 'google', 'copie mise en pause proprement (jeton Google)');
  W.gd.token = 'G2'; W.gCalls = 0; W.gExpireAfterCalls = Infinity;
  r = await W.engine().run(src, {});
  ok(!r.paused && r.failed.length === 0, 'reprise sans erreur');
  sameTree(W, src, mirrorRoot(W), 'après reprise');

  console.log('\n6) Plantage au moment d\'enregistrer l\'index → journal local, pas de doublon');
  W = makeWorld(); buildTree(W.od);
  W.crashOnIndexSave = true; W.crashAfter = 0;          // plantage dès le tout premier enregistrement
  await W.engine().run(src, {}).catch(() => {});
  W.crashOnIndexSave = false;
  await W.engine().run(src, {});
  const roots = [...W.gd.files.values()].filter((f) => f.name === 'Immobilier' && f.folder);
  ok(roots.length === 1, 'un seul dossier miroir créé malgré le plantage initial');
  sameTree(W, src, mirrorRoot(W), 'après plantage initial');
  W = makeWorld(); buildTree(W.od);
  W.crashOnIndexSave = true; W.crashAfter = 2;          // plantage à l'enregistrement final (1 = création, 2 = prise de main)
  r = await W.engine().run(src, {}).catch((e) => ({ crashed: e }));
  ok(W.journalStore.size === 1, 'journal local conservé après plantage');
  W.crashOnIndexSave = false;
  const up = W.gd.uploadsCount;
  r = await W.engine().run(src, {});
  ok(r.plan.counts.nouveauxFichiers === 0, 'rien à renvoyer après reprise du journal');
  ok(W.gd.uploadsCount === up, 'aucun doublon envoyé');
  sameTree(W, src, mirrorRoot(W), 'après plantage');

  console.log('\n7) Mode « corbeille Google »');
  W.od.remove('P5');
  r = await W.engine().run(src, { delMode: 'trash' });
  const f5 = [...W.gd.files.values()].find((f) => f.name === 'photo 5.jpg');
  ok(f5.trashed, 'photo 5 mise à la corbeille Google');
  sameTree(W, src, mirrorRoot(W), 'mode corbeille');

  console.log('\n8) Fichiers / dossier effacés à la main dans Google → réparation');
  const gBail = [...W.gd.files.values()].find((f) => f.name === "Bail d'Aurélie.pdf");
  W.gd.files.delete(gBail.id);
  W.od.setContent('BAIL', 'bail v3');
  r = await W.engine().run(src, {});
  sameTree(W, src, mirrorRoot(W), 'fichier recréé');

  console.log('\n9) Analyse seule : rien n\'est écrit');
  W.od.add('DRY', 'S1', 'test.pdf', false, 't');
  const nFiles = W.gd.files.size;
  r = await W.engine().run(src, { dryRun: true });
  ok(r.plan.counts.nouveauxFichiers === 1, 'analyse : 1 nouveau fichier annoncé');
  ok(W.gd.files.size === nFiles, 'analyse : aucun fichier créé dans Google');

  console.log('\n10) Confirmation refusée → rien n\'est fait');
  r = await W.engine().run(src, { confirm: async () => false });
  ok(r.paused === 'annulé' && W.gd.files.size === nFiles, 'annulation respectée');

  console.log('\n11) Arrêt demandé par l\'utilisateur, puis reprise');
  for (let i = 0; i < 30; i++) W.od.add('Y' + i, 'B1', `y${i}.jpg`, false, 'y' + i);
  const e11 = W.engine({ progress: (p) => { if (p.phase === 'transfer' && p.done === 5) e11.stop(); } });
  r = await e11.run(src, {});
  ok(r.paused === 'stop', 'arrêt pris en compte');
  r = await W.engine().run(src, {});
  sameTree(W, src, mirrorRoot(W), 'après arrêt et reprise');

  console.log('\n12) Coupure réseau pendant l\'envoi d\'un gros fichier + même moteur réutilisé');
  W.od.setContent('GROS', big('z', 17));
  W.failNextChunk = true;
  const eng = W.engine();
  r = await eng.run(src, {});
  ok(r.failed.length === 0, 'morceau renvoyé après coupure, sans erreur');
  sameTree(W, src, mirrorRoot(W), 'gros fichier après coupure');
  W.od.add('Z1', 'S1', 'z1.pdf', false, 'z1');
  r = await eng.run(src, {});
  sameTree(W, src, mirrorRoot(W), 'second passage avec le même moteur');

  console.log('\n13) Copie Google modifiée à la main → remise exacte en miroir');
  const G = W.gd; const byName = (n) => [...G.files.values()].find((f) => f.name === n && !f.trashed);
  G.files.delete(byName('photo 6.jpg').id);                                  // fichier effacé dans Google
  byName('EDLE.docx').content = 'abîmé';                                      // contenu altéré dans Google
  byName('photo 7.jpg').name = 'photo 7 renommée à la main.jpg';             // renommé dans Google
  const gS1 = byName('Studio 1').id, gB1 = byName('STUDIO 1').id;
  const nf = byName('test.pdf'); nf.parents = [gB1];                        // déplacé dans Google
  G.create({ name: 'intrus.txt', parents: [gS1] }, 'pas dans OneDrive');      // fichier en trop
  const intrusDir = G.create({ name: 'Dossier intrus', parents: [mirrorRoot(W)], mimeType: 'application/vnd.google-apps.folder' });
  G.create({ name: 'dedans.txt', parents: [intrusDir] }, 'x');                 // dossier en trop avec contenu
  const odBefore = JSON.stringify(W.od.tree('IMMO'));
  r = await W.engine().run(src, { delMode: 'trash' });
  ok(r.failed.length === 0, 'aucune erreur');
  ok(r.plan.counts.enTropDansGoogle === 3, 'éléments en trop dans Google détectés (3)');
  ok(r.plan.verif.manquants === 1 && r.plan.verif.contenuDifferent === 1 && r.plan.verif.deplacesGoogle === 2, 'manquant, contenu altéré et 2 déplacements/renommages détectés');
  sameTree(W, src, mirrorRoot(W), 'après réparation de la copie Google');
  ok(JSON.stringify(W.od.tree('IMMO')) === odBefore, 'OneDrive strictement inchangé');
  r = await W.engine().run(src, {});
  ok(r.plan.counts.reparesGoogle === 0 && r.plan.counts.enTropDansGoogle === 0 && r.plan.fileOps.length === 0, 'passage suivant : plus rien à corriger');

  console.log('\n14) Adresse /content inutilisable → secours par le lien de téléchargement');
  W.contentBroken = true;
  W.od.setContent('BAIL', 'bail v4');
  W.od.add('SEC', 'S1', 'secours.pdf', false, 'sec');
  r = await W.engine().run(src, {});
  ok(r.failed.length === 0, 'secours utilisé sans erreur');
  sameTree(W, src, mirrorRoot(W), 'après téléchargement de secours');
  W.contentBroken = false;

  console.log('\n15) Copie interrompue sur l\'iPhone, reprise sur l\'iPad');
  W = makeWorld(); buildTree(W.od);
  let clock = 1e12;
  const iphone = { load: (id) => JSON.parse(W.journalStore.get(id) || '[]'), add: (id, e) => { const a = JSON.parse(W.journalStore.get(id) || '[]'); a.push(e); W.journalStore.set(id, JSON.stringify(a)); }, clear: (id) => W.journalStore.delete(id) };
  const ipadStore = new Map();
  const ipad = { load: (id) => JSON.parse(ipadStore.get(id) || '[]'), add: (id, e) => { const a = JSON.parse(ipadStore.get(id) || '[]'); a.push(e); ipadStore.set(id, JSON.stringify(a)); }, clear: (id) => ipadStore.delete(id) };
  W.crashOnIndexSave = true; W.crashAfter = 2;                 // l'iPhone copie tout, puis s'arrête avant d'enregistrer l'index
  await W.engine({ journal: iphone, device: { id: 'IPH', label: 'iPhone' }, now: () => clock }).run(src, {}).catch(() => {});
  W.crashOnIndexSave = false;
  ok(W.journalStore.size === 1, 'les opérations de l\'iPhone sont restées dans sa mémoire');
  clock += 2 * 60000;
  r = await W.engine({ journal: ipad, device: { id: 'IPD', label: 'iPad' }, now: () => clock }).run(src, {});
  ok(r.paused === 'autre-appareil' && r.lock.label === 'iPhone', 'iPad refuse de démarrer : l\'iPhone a copié il y a 2 min');
  clock += 11 * 60000;
  const upB = W.gd.uploadsCount;
  r = await W.engine({ journal: ipad, device: { id: 'IPD', label: 'iPad' }, now: () => clock }).run(src, {});
  ok(!r.paused && r.failed.length === 0, 'iPad démarre une fois le délai passé');
  ok(r.plan.counts.enTropDansGoogle === 0, 'aucun fichier de l\'iPhone pris pour un intrus');
  ok(r.plan.counts.reprisDansIndex > 0 && r.plan.counts.nouveauxFichiers === 0 && r.plan.counts.nouveauxDossiers === 0, 'fichiers et dossiers repris dans l\'index (' + r.plan.counts.reprisDansIndex + '), rien à recopier');
  ok(W.gd.uploadsCount === upB, 'aucun fichier renvoyé');
  ok([...W.gd.files.values()].every((f) => !f.trashed), 'rien mis à la corbeille');
  sameTree(W, src, mirrorRoot(W), 'après reprise sur l\'iPad');

  console.log('\n16) Fichiers copiés par la v1.4 (sans version notée) → contenu renvoyé une fois, sans corbeille');
  const ids = new Map([...W.gd.files.values()].filter((f) => !f.folder && f.appProperties).map((f) => [f.id, f.name]));
  for (const f of W.gd.files.values()) if (f.appProperties) delete f.appProperties.t;
  // l'index perd ces fichiers (comme une copie interrompue)
  const st = W.engine({ journal: ipad, device: { id: 'IPD', label: 'iPad' }, now: () => clock });
  const idxFile = [...W.gd.files.values()].find((f) => f.name.startsWith('_index-'));
  const idx = JSON.parse(Buffer.from(idxFile.content, 'latin1').toString('utf8'));
  for (const [k, v] of Object.entries(idx.items)) if (!v.f) delete idx.items[k];
  idxFile.content = Buffer.from(JSON.stringify(idx), 'utf8').toString('latin1');
  r = await st.run(src, {});
  ok(r.plan.counts.enTropDansGoogle === 0 && r.plan.counts.nouveauxFichiers === 0, 'repris, pas d\'intrus, pas de nouveau fichier');
  ok(r.plan.counts.modifies === ids.size, 'contenu renvoyé une fois par sécurité (' + r.plan.counts.modifies + ')');
  ok([...ids.keys()].every((id) => W.gd.files.has(id) && !W.gd.files.get(id).trashed), 'mêmes fichiers Google conservés, rien à la corbeille');
  sameTree(W, src, mirrorRoot(W), 'après reprise des fichiers v1.4');
  r = await st.run(src, {});
  ok(r.plan.fileOps.length === 0 && r.plan.counts.enTropDansGoogle === 0, 'passage suivant : plus rien à faire');

  console.log('\n17) Fichier corrompu (photo de 135 octets) : JAMAIS copié, signalé');
  W = makeWorld(); buildTree(W.od);
  const bad135 = 'x'.repeat(135);
  W.od.add('CORR', 'S1', 'IMG_3376.jpeg', false, bad135, true);
  r = await W.engine().run(src, {});
  const gHas = (name) => [...W.gd.files.values()].some((f) => f.name === name && !f.trashed);
  ok(!gHas('IMG_3376.jpeg'), 'la photo corrompue n\'a PAS été copiée dans Google');
  ok(allSusp(r).length === 1 && /taille impossible|n'est pas/.test(allSusp(r)[0].reason) && r.plan.counts.corrompus === 1 && !r.plan.fileOps.some((o) => o.it.id === 'CORR'), 'signalée, comptée « corrompue » et non « nouvelle » : ' + allSusp(r)[0].reason);
  r = await W.engine().run(src, {});
  ok(!gHas('IMG_3376.jpeg') && allSusp(r).length === 1, 'au passage suivant : toujours pas copiée, toujours signalée');

  console.log('\n18) Variantes de corruption : vide, coupée, remplie de zéros, remplacée par du texte');
  const mk = (id, name, raw) => W.od.add(id, 'B1', name, false, raw, true);
  const good = W.od.items.get('P6').content;
  mk('C1', 'vide.jpg', '');
  mk('C2', 'coupee.jpg', good.slice(0, 15000) + ' '.repeat(10000));          // assez lourde mais sans fin d'image
  mk('C3', 'zeros.jpg', '\0'.repeat(30000));
  mk('C4', 'texte.jpg', '<html>Erreur 503 : service indisponible</html>' + ' '.repeat(30000));
  mk('C5', 'coupe.pdf', '%PDF-1.4\n' + ' '.repeat(5000));
  r = await W.engine().run(src, {});
  for (const n of ['vide.jpg', 'coupee.jpg', 'zeros.jpg', 'texte.jpg', 'coupe.pdf']) ok(!gHas(n), n + ' non copié — ' + (allSusp(r).find((x) => x.path.endsWith(n)) || {}).reason);
  ok(/texte/.test(r.suspects.find((x) => x.path.endsWith('texte.jpg')).preview), 'le contenu réel est décrit : ' + r.suspects.find((x) => x.path.endsWith('texte.jpg')).preview);

  console.log('\n19) Une bonne photo déjà copiée se corrompt dans OneDrive : la copie Google reste intacte');
  const before7 = [...W.gd.files.values()].find((f) => f.name === 'photo 7.jpg').content;
  W.od.setContent('P7', 'y'.repeat(135), true);
  r = await W.engine().run(src, {});
  const g7 = [...W.gd.files.values()].find((f) => f.name === 'photo 7.jpg');
  ok(g7.content === before7, 'copie Google de photo 7 inchangée (bonne version conservée)');
  ok(allSusp(r).some((x) => x.path.endsWith('photo 7.jpg') && x.inGoogle), 'signalée comme corrompue dans OneDrive, copie Google protégée');

  console.log('\n20) Vraie modification : l\'ancienne version Google est conservée pour toujours');
  const gBail2 = [...W.gd.files.values()].find((f) => f.name === "Bail d'Aurélie.pdf");
  const oldBail = gBail2.content;
  W.od.setContent('BAIL', 'bail v9 signé');
  r = await W.engine().run(src, {});
  ok(gBail2.kept && gBail2.kept.some((k) => k.keepForever && k.content === oldBail), 'ancienne version marquée « à conserver pour toujours »');
  ok(gBail2.content === W.od.items.get('BAIL').content, 'nouvelle version en place');

  console.log('\n21) Rapport d\'intégrité (lecture seule) + retrait de Google sur demande');
  // fichier corrompu copié autrefois (avant la v2.0) : on le simule en le corrompant dans OneDrive sans changer sa version
  W.od.items.get('P5').content = 'z'.repeat(135);
  const nGoogle = W.gd.files.size, odCalls = W.od.calls;
  const e21 = W.engine();
  const rep1 = await e21.integrity(src, {});
  ok(W.gd.files.size === nGoogle, 'contrôle rapide : rien écrit dans Google');
  const names = rep1.suspects.map((x) => x.path.split('/').pop()).sort();
  ok(['IMG_3376.jpeg', 'photo 5.jpg', 'photo 7.jpg', 'vide.jpg'].every((n) => names.includes(n)), 'contrôle rapide : trouve les fichiers trop petits ou vides (' + names.join(', ') + ')');
  ok(rep1.suspects.find((x) => x.path.endsWith('photo 5.jpg')).inGoogle === true, 'photo 5 signalée comme présente dans Google');
  const rep2 = await e21.integrity(src, { deep: true });
  const names2 = rep2.suspects.map((x) => x.path.split('/').pop()).sort();
  ok(['coupee.jpg', 'zeros.jpg', 'texte.jpg', 'coupe.pdf'].every((n) => names2.includes(n)), 'contrôle complet : trouve aussi les fichiers coupés ou remplacés (' + names2.length + ' au total)');
  ok(!names2.includes("Bail d'Aurélie.pdf") && !names2.includes('photo 1 renommée.jpg') && !names2.includes('photo 6.jpg'), 'aucun faux positif sur les bons fichiers');
  const gP5 = [...W.gd.files.values()].find((f) => f.name === 'photo 5.jpg' && !f.trashed);
  const rr = await e21.removeFromGoogle(src, [rep1.suspects.find((x) => x.path.endsWith('photo 5.jpg')).id]);
  ok(rr.removed === 1 && rr.failed.length === 0 && gP5.trashed, 'retrait sur demande : photo 5 mise à la corbeille Google');
  r = await W.engine().run(src, {});
  ok(!gHas('photo 5.jpg') && allSusp(r).some((x) => x.path.endsWith('photo 5.jpg')), 'ensuite : photo 5 ni recopiée ni oubliée — toujours signalée');

  console.log('\n22) Retrait en nombre (220 copies corrompues comme chez vous)');
  W = makeWorld(); buildTree(W.od);
  for (let i = 0; i < 220; i++) W.od.add('Q' + i, 'B1', 'IMG_' + (3300 + i) + '.jpeg', false, 'ok' + i);
  await W.engine().run(src, {});                                    // copiées alors qu'elles étaient bonnes…
  for (let i = 0; i < 220; i++) W.od.items.get('Q' + i).content = 'bplist00' + 'x'.repeat(127);   // …puis corrompues (même version)
  const e22 = W.engine();
  const rep22 = await e22.integrity(src, {});
  ok(rep22.suspects.length === 220 && rep22.suspects.every((x) => x.inGoogle), '220 copies corrompues repérées dans Google');
  let lastProg = null;
  const e22b = W.engine({ progress: (pp) => { if (pp.phase === 'remove') lastProg = pp; } });
  const rr22 = await e22b.removeFromGoogle(src, rep22.suspects.map((x) => x.id));
  ok(rr22.removed === 220 && rr22.failed.length === 0, '220 retirées, aucun échec');
  ok(lastProg && lastProg.done === 220 && lastProg.total === 220, 'avancement affiché jusqu\'à 220 / 220');
  r = await W.engine().run(src, {});
  ok(r.plan.counts.corrompus === 220 && r.plan.counts.nouveauxFichiers === 0 && r.plan.fileOps.length === 0, 'passage suivant : 220 « corrompus, non copiés », 0 nouveau, rien téléchargé');
  ok(r.plan.counts.enTropDansGoogle === 0, 'rien de « en trop »');

  console.log('\n23) Check files (Julien) : fichiers envoyés ces dernières 24 heures');
  W = makeWorld(); buildTree(W.od);
  const old = new Date(Date.now() - 3 * 86400000).toISOString();
  for (const it of W.od.items.values()) it.mod = old;                         // tout ce qui existe date de 3 jours
  W.od.add('J1', 'S2', 'IMG_9001.jpeg', false, 'photo envoyée ce matin');      // envoyée aujourd'hui, intacte
  W.od.add('J2', 'S2', 'IMG_9002.jpeg', false, 'bplist00' + 'x'.repeat(127), true);   // envoyée aujourd'hui, corrompue (135 octets)
  W.od.add('J3', 'S2', 'IMG_9003.jpeg', false, '\xFF\xD8\xFF\xE0' + ' '.repeat(40000), true);  // coupée (pas de fin)
  const nG = W.gd.files.size;
  const cf = await W.engine().integrity(src, { sinceMs: Date.now() - 86400000 });
  ok(cf.recent.length === 3, 'seuls les 3 fichiers des dernières 24 h sont contrôlés (' + cf.recent.length + ' sur ' + cf.total + ')');
  ok(cf.recent.filter((x) => x.ok).length === 1 && cf.recent.find((x) => x.ok).path.endsWith('IMG_9001.jpeg'), 'IMG_9001 : ✅ intacte');
  ok(cf.suspects.length === 2 && cf.recent[0].ok === false, 'IMG_9002 et IMG_9003 : ❌ corrompues, listées en premier');
  ok(W.gd.files.size === nG, 'rien n\'est écrit dans Google');

  console.log('\n24) Vérification d\'une visite EDL (empreintes SHA-256 notées par EDL avant l\'envoi)');
  W = makeWorld(); buildTree(W.od);
  const sha256 = (t) => crypto.createHash('sha256').update(Buffer.from(t, 'latin1')).digest('hex');
  const ph = (id, nom, content) => { W.od.add(id, 'S2', nom, false, content); return { photo_id: id, nom_fichier: nom, onedrive_item_id: id, statut_transfert: 'confirme',
    empreinte_sha256: sha256(W.od.items.get(id).content), taille_octets: W.od.items.get(id).content.length }; };
  const p1 = ph('E1', 'EDLE_CH1-A_001.jpg', 'mur A'), p2 = ph('E2', 'EDLE_CH1-B_002.jpg', 'mur B'), p3 = ph('E3', 'EDLE_CH1-C_003.jpg', 'mur C');
  W.od.items.get('E2').content = W.od.items.get('E2').content.slice(0, 15000);      // arrivée coupée dans OneDrive
  W.od.remove('E3');                                                                 // disparue de OneDrive
  const p4 = { photo_id: 'E4', nom_fichier: 'EDLE_SDB-A_001.jpg', statut_transfert: 'en_attente', empreinte_sha256: 'x' };
  const p5 = { photo_id: 'E5', nom_fichier: 'EDLE_SDB-B_002.jpg', statut_transfert: 'echec', empreinte_sha256: 'y' };
  W.od.add('PV', 'S2', 'PV-EDLE.pdf', false, 'proces-verbal signe');
  const visite = { visit_id: 'v123', bien: { unite_source: 'Studio 2' }, photos: [p1, p2, p3, p4, p5], preuve: { pv_onedrive_item_id: 'PV' } };
  W.od.add('VJ', 'S2', 'visite_v123.json', false, Buffer.from(JSON.stringify(visite), 'utf8').toString('latin1'), true);
  const nG24 = W.gd.files.size;
  const rv = await W.engine().verifyEdlVisit({ visitId: 'v123', driveId: 'DRV1', folderId: 'S2' });
  const st24 = (n) => (rv.items.find((x) => x.nom === n) || {}).state;
  ok(st24('EDLE_CH1-A_001.jpg') === 'ok', 'photo intacte : ✅ identique au bit près');
  ok(st24('EDLE_CH1-B_002.jpg') === 'bad' && /DIFFÉRENTE/.test(rv.items.find((x) => x.nom === 'EDLE_CH1-B_002.jpg').reason), 'photo coupée : ❌ ' + rv.items.find((x) => x.nom === 'EDLE_CH1-B_002.jpg').reason);
  ok(st24('EDLE_CH1-C_003.jpg') === 'bad' && /absente/.test(rv.items.find((x) => x.nom === 'EDLE_CH1-C_003.jpg').reason), 'photo disparue : ❌ absente de OneDrive');
  ok(st24('EDLE_SDB-A_001.jpg') === 'attente', 'photo pas encore envoyée : ⏳');
  ok(st24('EDLE_SDB-B_002.jpg') === 'bad', 'photo refusée à l\'envoi : ❌');
  ok(st24('Procès-verbal (PDF)') === 'ok', 'procès-verbal : ✅ PDF complet');
  ok(rv.ok === 2 && rv.bad === 3 && rv.pending === 1 && rv.items[0].state === 'bad', 'bilan 2 ✅ / 3 ❌ / 1 ⏳, les ❌ en premier');
  ok(W.gd.files.size === nG24, 'rien écrit dans Google');
  let err24 = null;
  try { await W.engine().verifyEdlVisit({ visitId: 'inconnue', driveId: 'DRV1', folderId: 'S2' }); } catch (e) { err24 = e.message; }
  ok(/pas encore dans OneDrive/.test(err24 || ''), 'visite introuvable : message clair');

  console.log(`\n✅ ${checks} contrôles réussis`);
})().catch((e) => { console.error('\n❌ ÉCHEC :', e); process.exit(1); });
