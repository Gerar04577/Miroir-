// Banc d'essai du moteur Miroir : faux OneDrive + faux Google Drive, vrai engine.js
'use strict';
const assert = require('assert');
const path = require('path');
const crypto = require('crypto');
const sha1 = (t) => crypto.createHash('sha1').update(Buffer.from(t, 'utf8')).digest('hex');
const { createEngine, PauseError } = require(path.join(__dirname, '..', 'engine.js'));

let checks = 0;
function ok(cond, msg) { assert.ok(cond, msg); checks++; console.log('  ✔ ' + msg); }

/* ---------------- Faux OneDrive ---------------- */
class FakeOneDrive {
  constructor() { this.items = new Map(); this.n = 0; this.driveId = 'DRV1'; this.failList = new Set(); this.calls = 0; this.add('ROOT', null, 'root', true); }
  add(id, parent, name, folder, content) {
    const it = { id, parent, name, folder, content: content || null, ctag: 'c' + (++this.n) };
    this.items.set(id, it); return it;
  }
  setContent(id, c) { const it = this.items.get(id); it.content = c; it.ctag = 'c' + (++this.n); }
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

function resp(status, body, headers = {}) {
  const h = new Map(Object.entries(headers).map(([k, v]) => [k.toLowerCase(), v]));
  const text = body === undefined || body === null ? '' : (typeof body === 'string' ? body : JSON.stringify(body));
  return {
    status, ok: status >= 200 && status < 300,
    headers: { get: (k) => h.get(k.toLowerCase()) || null },
    text: async () => text, json: async () => JSON.parse(text), blob: async () => new Blob([text])
  };
}
async function bodyText(b) { if (b == null) return ''; if (typeof b === 'string') return b; return await b.text(); }

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
          const base = { id: k.id, name: k.name };
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
        return resp(200, it.content);
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
      return resp(200, it.content);
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
        const files = all.slice(start, start + ps).map((f) => ({ id: f.id, name: f.name, mimeType: f.folder ? 'application/vnd.google-apps.folder' : 'application/octet-stream', size: f.folder ? undefined : String((f.content || '').length), sha1Checksum: f.folder ? undefined : sha1(f.content || '') }));
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
          if (u.searchParams.get('alt') === 'media') return resp(200, f.content);
          return resp(200, { id: f.id, trashed: f.trashed, parents: f.parents });
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
      if (p === '/upload/drive/v3/files' && method === 'POST') {
        const type = u.searchParams.get('uploadType');
        if (type === 'multipart') {
          const raw = await bodyText(opt.body);
          const boundary = /boundary=(.+)$/.exec(opt.headers['Content-Type'])[1];
          const parts = raw.split('--' + boundary);
          const meta = JSON.parse(parts[1].split('\r\n\r\n')[1].trim());
          const content = parts[2].split('\r\n\r\n').slice(1).join('\r\n\r\n').replace(/\r\n$/, '');
          if (meta.name.startsWith('_index-') && world.crashOnIndexSave && ++world.indexSaves > world.crashAfter) throw new Error('CRASH simulé');
          gd.uploadsCount++;
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
          f.content = await bodyText(opt.body); gd.uploadsCount++; return resp(200, { id: f.id });
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
    const txt = await blob.text();
    assert.strictEqual(txt.length, b - a + 1, 'taille du morceau');
    if (b + 1 < tot) assert.strictEqual((b - a + 1) % (256 * 1024), 0, 'morceau multiple de 256 Kio');
    s.parts.push(txt);
    if (b + 1 < tot) return { status: 308, range: 'bytes=0-' + b };
    const content = s.parts.join('');
    gd.uploadsCount++;
    if (s.existing) { gd.files.get(s.existing).content = content; return { status: 200, json: { id: s.existing } }; }
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
  ok(W.gd.uploadsCount === before + 1, 'aucun fichier renvoyé (seul l\'index est réécrit)');

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
  W.crashOnIndexSave = true; W.crashAfter = 1;          // plantage à l'enregistrement final
  r = await W.engine().run(src, {}).catch((e) => ({ crashed: e }));
  ok(W.journalStore.size === 1, 'journal local conservé après plantage');
  W.crashOnIndexSave = false;
  const up = W.gd.uploadsCount;
  r = await W.engine().run(src, {});
  ok(r.plan.counts.nouveauxFichiers === 0, 'rien à renvoyer après reprise du journal');
  ok(W.gd.uploadsCount === up + 2, 'aucun doublon envoyé (seulement 2 écritures de l\'index)');
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

  console.log(`\n✅ ${checks} contrôles réussis`);
})().catch((e) => { console.error('\n❌ ÉCHEC :', e); process.exit(1); });
