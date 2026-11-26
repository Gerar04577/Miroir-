/* Miroir OneDrive -> Google Drive — moteur de synchronisation
 * Fichier : engine.js — VERSION 1.3
 * Sens unique : OneDrive (LECTURE SEULE, jamais aucune écriture ni suppression) -> Google Drive.
 * À chaque passage, le miroir Google est relu et remis exactement à l'image de OneDrive.
 * Aucune dépendance externe. Testable hors navigateur (fetch / xhrPut injectés).
 */
(function (root) {
  'use strict';

  const ENGINE_VERSION = '1.3';
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const GD = 'https://www.googleapis.com/drive/v3';
  const GU = 'https://www.googleapis.com/upload/drive/v3';
  const FOLDER_MIME = 'application/vnd.google-apps.folder';
  const TOP_NAME = 'Miroir OneDrive';
  const BIN_NAME = '_Supprimés';
  const SMALL = 5 * 1024 * 1024;          // <= 5 Mo : envoi en une requête
  const CHUNK = 8 * 1024 * 1024;          // morceaux de 8 Mo (multiple de 256 Kio exigé par Google)
  const LIST_PARALLEL = 6;
  const TRANSFER_PARALLEL = 3;
  const SAVE_EVERY_OPS = 25;
  const SAVE_EVERY_MS = 60000;

  class PauseError extends Error {
    constructor(reason) { super('pause:' + reason); this.pause = reason; }
  }
  class HttpError extends Error {
    constructor(service, status, body) {
      super(service + ' ' + status + (body ? ' — ' + String(body).slice(0, 300) : ''));
      this.status = status; this.service = service;
    }
  }

  const q = (s) => String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");

  function createEngine(env) {
    const fetchFn = env.fetch;
    const xhrPut = env.xhrPut;
    const msToken = env.msToken;           // async (force?) => jeton Microsoft valide
    const gToken = env.gToken;             // async (margeSecondes) => jeton Google ou PauseError('google')
    const gTokenQuiet = env.gTokenQuiet || (() => null);
    const log = env.log || (() => {});
    const progress = env.progress || (() => {});
    const journal = env.journal;           // {load(id), add(id, entry), clear(id)}
    const sleep = env.sleep || ((ms) => new Promise((r) => setTimeout(r, ms)));
    const now = env.now || (() => Date.now());
    let stopRequested = false;

    async function safeText(r) { try { return await r.text(); } catch (e) { return ''; } }

    /* ---------- Microsoft Graph (lecture seule) ---------- */
    async function graph(url) {
      for (let attempt = 0; ; attempt++) {
        const tok = await msToken(attempt === 1);
        let r;
        try {
          r = await fetchFn(url.startsWith('http') ? url : GRAPH + url, { headers: { Authorization: 'Bearer ' + tok } });
        } catch (e) {
          if (attempt < 4) { await sleep(1000 * 2 ** attempt); continue; }
          throw e;
        }
        if ((r.status === 429 || r.status >= 500) && attempt < 5) {
          const ra = Number(r.headers.get('Retry-After')) || 2 ** attempt;
          await sleep(ra * 1000); continue;
        }
        if (r.status === 401 && attempt === 0) continue; // rafraîchit le jeton puis réessaie
        if (!r.ok) throw new HttpError('Microsoft', r.status, await safeText(r));
        return r.json();
      }
    }

    /* ---------- Google Drive ---------- */
    async function drive(url, opt = {}, want = 'json') {
      for (let attempt = 0; ; attempt++) {
        const tok = await gToken(120);
        const headers = Object.assign({}, opt.headers || {}, { Authorization: 'Bearer ' + tok });
        let r;
        try {
          r = await fetchFn(url, Object.assign({}, opt, { headers }));
        } catch (e) {
          if (attempt < 4) { await sleep(1000 * 2 ** attempt); continue; }
          throw e;
        }
        if (r.status === 401) throw new PauseError('google');
        if ((r.status === 429 || r.status >= 500) && attempt < 6) { await sleep(1000 * 2 ** attempt); continue; }
        if (r.status === 403 && attempt < 6) {
          const t = await safeText(r);
          if (/rateLimitExceeded|userRateLimitExceeded/.test(t)) { await sleep(1000 * 2 ** attempt); continue; }
          throw new HttpError('Google', 403, t);
        }
        if (!r.ok) throw new HttpError('Google', r.status, await safeText(r));
        if (want === 'response') return r;
        if (want === 'none') return null;
        const txt = await r.text();
        return txt ? JSON.parse(txt) : {};
      }
    }

    async function gFindChild(parentId, name, folderOnly) {
      let query = `name='${q(name)}' and '${q(parentId)}' in parents and trashed=false`;
      if (folderOnly) query += ` and mimeType='${FOLDER_MIME}'`;
      const j = await drive(GD + '/files?pageSize=10&fields=files(id,name)&q=' + encodeURIComponent(query));
      return (j.files && j.files[0]) ? j.files[0].id : null;
    }
    async function gMkdir(name, parentId) {
      const j = await drive(GD + '/files?fields=id', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, mimeType: FOLDER_MIME, parents: [parentId] })
      });
      return j.id;
    }
    async function gEnsureFolder(name, parentId) {
      return (await gFindChild(parentId, name, true)) || (await gMkdir(name, parentId));
    }
    async function gExists(id) {
      try {
        const j = await drive(GD + '/files/' + encodeURIComponent(id) + '?fields=id,trashed');
        return !j.trashed;
      } catch (e) { if (e.status === 404) return false; throw e; }
    }
    async function gMove(id, newParent, newName) {
      const cur = await drive(GD + '/files/' + encodeURIComponent(id) + '?fields=parents');
      const remove = (cur.parents || []).filter((p) => p !== newParent).join(',');
      let url = GD + '/files/' + encodeURIComponent(id) + '?fields=id&addParents=' + encodeURIComponent(newParent);
      if (remove) url += '&removeParents=' + encodeURIComponent(remove);
      const body = newName !== undefined ? { name: newName } : {};
      await drive(url, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    }

    /* Envoi d'un contenu : création (meta.parents) ou remplacement (existingId) */
    async function gUpload(blob, meta, existingId) {
      const mime = blob.type || 'application/octet-stream';
      if (blob.size <= SMALL) {
        if (existingId) {
          const j = await drive(GU + '/files/' + encodeURIComponent(existingId) + '?uploadType=media&fields=id', {
            method: 'PATCH', headers: { 'Content-Type': mime }, body: blob
          });
          return j.id;
        }
        const boundary = 'miroir' + Math.random().toString(36).slice(2);
        const body = new Blob([
          `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n` +
          `--${boundary}\r\nContent-Type: ${mime}\r\n\r\n`, blob, `\r\n--${boundary}--`
        ]);
        const j = await drive(GU + '/files?uploadType=multipart&fields=id', {
          method: 'POST', headers: { 'Content-Type': 'multipart/related; boundary=' + boundary }, body
        });
        return j.id;
      }
      // Gros fichier : envoi reprenable par morceaux
      const initUrl = existingId
        ? GU + '/files/' + encodeURIComponent(existingId) + '?uploadType=resumable&fields=id'
        : GU + '/files?uploadType=resumable&fields=id';
      const r = await drive(initUrl, {
        method: existingId ? 'PATCH' : 'POST',
        headers: { 'Content-Type': 'application/json; charset=UTF-8', 'X-Upload-Content-Type': mime, 'X-Upload-Content-Length': String(blob.size) },
        body: JSON.stringify(existingId ? {} : meta)
      }, 'response');
      const session = r.headers.get('Location');
      if (!session) throw new Error('Google n\'a pas renvoyé d\'adresse d\'envoi (Location)');
      let start = 0, fails = 0;
      while (start < blob.size) {
        if (stopRequested) throw new PauseError('stop');
        const end = Math.min(start + CHUNK, blob.size);
        const headers = { 'Content-Range': `bytes ${start}-${end - 1}/${blob.size}` };
        const tok = gTokenQuiet();
        if (tok) headers.Authorization = 'Bearer ' + tok;
        let res;
        try { res = await xhrPut(session, headers, blob.slice(start, end)); }
        catch (e) { res = { status: 0 }; }
        if (res.status === 200 || res.status === 201) return (res.json && res.json.id) || existingId;
        if (res.status === 308) {
          const range = res.range;                // "bytes=0-12345" si lisible
          const m = range && /-(\d+)$/.exec(range);
          start = m ? Number(m[1]) + 1 : end;
          fails = 0; progress({ chunk: true, bytes: end - start }); continue;
        }
        if (res.status === 404 || res.status === 410) throw new Error('Session d\'envoi Google expirée');
        if (++fails > 4) throw new HttpError('Google', res.status, 'échec d\'envoi d\'un morceau');
        await sleep(1000 * 2 ** fails);
        // Demande à Google où il en est
        try {
          const st = await xhrPut(session, { 'Content-Range': `bytes */${blob.size}` }, null);
          if (st.status === 200 || st.status === 201) return (st.json && st.json.id) || existingId;
          if (st.status === 308) { const m = st.range && /-(\d+)$/.exec(st.range); start = m ? Number(m[1]) + 1 : 0; }
        } catch (e) { /* on réessaie le même morceau */ }
      }
      throw new Error('Envoi incomplet');
    }

    function sha1FromOd(fileFacet) {
      const h = fileFacet && fileFacet.hashes && fileFacet.hashes.sha1Hash;
      return h ? String(h).toLowerCase() : '';
    }
    async function sha1Of(blob) {
      try {
        if (blob.size > 500 * 1024 * 1024 || !(globalThis.crypto && crypto.subtle)) return '';
        const d = await crypto.subtle.digest('SHA-1', await blob.arrayBuffer());
        return Array.from(new Uint8Array(d)).map((b) => b.toString(16).padStart(2, '0')).join('');
      } catch (e) { return ''; }
    }

    /* ---------- Lecture complète du miroir Google (vérification) ---------- */
    async function listGoogle(rootG) {
      const nodes = new Map(), errors = [];
      const queue = [rootG];
      let idx = 0, running = 0, pauseErr = null;
      async function listOne(fid) {
        let token = null;
        do {
          const url = GD + '/files?pageSize=1000&fields=nextPageToken,files(id,name,mimeType,size,sha1Checksum)&q=' +
            encodeURIComponent(`'${q(fid)}' in parents and trashed=false`) + (token ? '&pageToken=' + encodeURIComponent(token) : '');
          const j = await drive(url);
          for (const f of j.files || []) {
            const folder = f.mimeType === FOLDER_MIME;
            nodes.set(f.id, { id: f.id, name: f.name, folder, size: f.size != null ? Number(f.size) : null, sha1: String(f.sha1Checksum || '').toLowerCase(), parent: fid });
            if (folder) queue.push(f.id);
          }
          token = j.nextPageToken || null;
        } while (token);
      }
      await new Promise((resolve) => {
        const pump = () => {
          while (running < LIST_PARALLEL && idx < queue.length && !pauseErr && !stopRequested) {
            const fid = queue[idx++]; running++;
            listOne(fid)
              .catch((e) => { if (e instanceof PauseError) pauseErr = e; else errors.push(String(e.message || e)); })
              .finally(() => { running--; progress({ phase: 'verify', folders: idx, items: nodes.size }); pump(); });
          }
          if (running === 0 && (idx >= queue.length || pauseErr || stopRequested)) resolve();
        };
        pump();
      });
      if (pauseErr) throw pauseErr;
      return { nodes, errors, complete: errors.length === 0 && !stopRequested };
    }

    /* Confronte l'index à ce qui existe réellement dans Google.
     * Corrige l'index en mémoire pour que le plan remette tout d'aplomb,
     * et renvoie les éléments présents dans Google mais absents de OneDrive (en trop). */
    function verify(index, w, gt, source) {
      const g2od = new Map();
      for (const [od, ex] of Object.entries(index.items)) g2od.set(ex.g, od);
      const walkById = new Map(w.items.map((i) => [i.id, i]));
      const fix = { manquants: 0, deplacesGoogle: 0, contenuDifferent: 0 };
      for (const [od, ex] of Object.entries(index.items)) {
        const node = gt.nodes.get(ex.g);
        if (!node || node.folder !== !!ex.f) { delete index.items[od]; fix.manquants++; continue; }
        const parentOd = node.parent === index.rootG ? source.id : g2od.get(node.parent);
        if (parentOd === undefined) { delete index.items[od]; fix.manquants++; continue; }
        if (ex.n !== node.name || ex.p !== parentOd) { ex.n = node.name; ex.p = parentOd; fix.deplacesGoogle++; }
        if (!ex.f) {
          const it = walkById.get(od);
          let diff = false;
          if (node.sha1 && ex.h) diff = node.sha1 !== ex.h;                  // empreinte de ce qui a été envoyé
          else if (node.sha1 && it && it.sha1) diff = node.sha1 !== it.sha1;  // empreinte fournie par OneDrive
          else if (node.size != null && ex.s != null) diff = node.size !== ex.s;
          if (diff) { ex.t = '__verif__'; fix.contenuDifferent++; }
        }
      }
      const mapped = new Set(Object.values(index.items).map((e) => e.g));
      const extra = [...gt.nodes.values()].filter((n) => !mapped.has(n.id));
      const extraSet = new Set(extra.map((n) => n.id));
      const topExtras = extra.filter((n) => !extraSet.has(n.parent)).map((n) => n.id);
      return { fix, extras: topExtras, extrasCount: extra.length };
    }

    /* ---------- Lecture complète du dossier OneDrive ---------- */
    async function walk(source) {
      const items = [], errors = [], skipped = [];
      const queue = [{ driveId: source.driveId, id: source.id, odId: source.id, path: '' }];
      let idx = 0, running = 0;
      async function listFolder(f) {
        let url = `/drives/${encodeURIComponent(f.driveId)}/items/${encodeURIComponent(f.id)}/children?$top=200&$select=id,name,size,file,folder,cTag,eTag,remoteItem,package`;
        const local = [];
        while (url) {
          const j = await graph(url);
          for (const it of j.value || []) {
            const ri = it.remoteItem || null;
            const isFolder = !!(it.folder || (ri && ri.folder));
            const isFile = !!(it.file || (ri && ri.file));
            const rel = f.path ? f.path + '/' + it.name : it.name;
            if (!isFolder && !isFile) { skipped.push(rel); continue; } // ex. bloc-notes OneNote
            const rec = {
              id: it.id, name: it.name, parent: f.odId, folder: isFolder,
              size: isFolder ? 0 : ((ri && ri.size) || it.size || 0),
              tag: isFolder ? '' : ((ri && (ri.cTag || ri.eTag)) || it.cTag || it.eTag || ''),
              path: rel,
              sha1: isFolder ? '' : sha1FromOd((ri && ri.file) || it.file),
              driveId: (ri && ri.parentReference && ri.parentReference.driveId) || f.driveId,
              realId: (ri && ri.id) || it.id
            };
            local.push(rec);
          }
          url = j['@odata.nextLink'] || null;
        }
        // tout le dossier est lu : on publie d'un bloc
        for (const rec of local) {
          items.push(rec);
          if (rec.folder) queue.push({ driveId: rec.driveId, id: rec.realId, odId: rec.id, path: rec.path });
        }
      }
      await new Promise((resolve) => {
        const pump = () => {
          while (running < LIST_PARALLEL && idx < queue.length && !stopRequested) {
            const f = queue[idx++]; running++;
            listFolder(f)
              .catch((e) => { errors.push({ path: f.path || '(racine)', err: String((e && e.message) || e) }); })
              .finally(() => { running--; progress({ phase: 'walk', folders: idx, items: items.length }); pump(); });
          }
          if (running === 0 && (idx >= queue.length || stopRequested)) resolve();
        };
        pump();
      });
      // Ordre garanti : un dossier apparaît toujours avant son contenu
      return { items, errors, skipped, complete: errors.length === 0 && !stopRequested };
    }

    /* ---------- Index (correspondance OneDrive -> Google), stocké dans Google ---------- */
    async function loadIndex(source) {
      const top = await gEnsureFolder(TOP_NAME, 'root');
      const fname = '_index-' + source.id + '.json';
      const fileG = await gFindChild(top, fname, false);
      let index = null;
      if (fileG) {
        const r = await drive(GD + '/files/' + encodeURIComponent(fileG) + '?alt=media', {}, 'response');
        try { index = JSON.parse(await r.text()); } catch (e) { throw new Error('Index Google illisible : ' + e.message); }
        index.fileG = fileG;
      }
      if (!index) index = { v: 1, sourceId: source.id, rootG: null, binG: null, fileG: null, items: {} };
      index.top = top; index.fname = fname;
      // Le dossier miroir existe-t-il toujours ?
      if (index.rootG && !(await gExists(index.rootG))) {
        log('⚠️ Le dossier miroir Google a disparu : tout sera recopié.');
        index.rootG = null; index.items = {};
      }
      // Rejoue le journal local (opérations faites mais pas encore enregistrées dans l'index)
      const pending = journal.load(source.id) || [];
      for (const e of pending) {
        if (e.root) { if (!index.rootG) index.rootG = e.root; continue; }
        if (e.del) delete index.items[e.id]; else index.items[e.id] = e.rec;
      }
      if (pending.length) log('Reprise : ' + pending.length + ' opération(s) récupérée(s) du journal local.');
      if (!index.rootG) {
        index.rootG = await gMkdir(source.name || 'OneDrive', top);
        journal.add(source.id, { root: index.rootG });
      }
      // L'index existe dans Google dès le départ : un dossier miroir n'est jamais créé deux fois
      if (!index.fileG || pending.length) await saveIndex(index, source);
      return index;
    }

    let saveChain = Promise.resolve();
    function saveIndex(index, source) {
      saveChain = saveChain.catch(() => {}).then(async () => {
        const data = JSON.stringify({ v: 1, sourceId: index.sourceId, sourceName: source.name, rootG: index.rootG, binG: index.binG, savedAt: new Date(now()).toISOString(), items: index.items });
        const blob = new Blob([data], { type: 'application/json' });
        if (index.fileG) {
          await drive(GU + '/files/' + encodeURIComponent(index.fileG) + '?uploadType=media&fields=id', {
            method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: blob
          });
        } else {
          index.fileG = await gUpload(blob, { name: index.fname, parents: [index.top], mimeType: 'application/json' });
        }
        journal.clear(source.id);
      });
      return saveChain;
    }

    /* ---------- Plan : ce qu'il faut faire ---------- */
    function plan(walkRes, index) {
      const seen = new Set();
      const folderOps = [], fileOps = [];
      let bytes = 0;
      for (const it of walkRes.items) {
        seen.add(it.id);
        const ex = index.items[it.id];
        const moved = ex && (ex.n !== it.name || ex.p !== it.parent);
        if (it.folder) {
          if (!ex) folderOps.push({ k: 'mkdir', it });
          else if (moved) folderOps.push({ k: 'move', it, ex });
        } else if (!ex) { fileOps.push({ k: 'new', it }); bytes += it.size; }
        else if (ex.t !== it.tag || ex.s !== it.size) { fileOps.push({ k: 'update', it, ex, moved }); bytes += it.size; }
        else if (moved) fileOps.push({ k: 'move', it, ex });
      }
      const dels = [];
      if (walkRes.complete) for (const id of Object.keys(index.items)) if (!seen.has(id)) dels.push(id);
      const delSet = new Set(dels);
      const topDels = dels.filter((id) => !delSet.has(index.items[id].p));
      return {
        folderOps, fileOps, dels, topDels, bytes,
        counts: {
          nouveauxDossiers: folderOps.filter((o) => o.k === 'mkdir').length,
          nouveauxFichiers: fileOps.filter((o) => o.k === 'new').length,
          modifies: fileOps.filter((o) => o.k === 'update').length,
          deplaces: folderOps.filter((o) => o.k === 'move').length + fileOps.filter((o) => o.k === 'move' || (o.k === 'update' && o.moved)).length,
          supprimes: dels.length,
          dejaAJour: walkRes.items.filter((i) => !i.folder).length - fileOps.filter((o) => o.k === 'new' || o.k === 'update').length
        }
      };
    }

    function parentG(index, source, odParent) {
      if (odParent === source.id) return index.rootG;
      const p = index.items[odParent];
      return p ? p.g : null;
    }

    /* Téléchargement d'un fichier OneDrive.
     * 1) Méthode principale : GET …/items/{id}/content avec le jeton — même méthode que
     *    Gestion Loyers et EDL, éprouvée sur ce OneDrive.
     * 2) Secours : demander l'élément SANS $select (un $select fait disparaître le lien
     *    @microsoft.graph.downloadUrl), puis télécharger ce lien pré-authentifié. */
    async function download(it) {
      const base = `/drives/${encodeURIComponent(it.driveId)}/items/${encodeURIComponent(it.realId)}`;
      let lastErr = null;
      for (let attempt = 0; attempt < 4; attempt++) {
        if (attempt) await sleep(1500 * 2 ** (attempt - 1));
        // 1) /content
        try {
          const tok = await msToken(attempt > 0);
          const r = await fetchFn(GRAPH + base + '/content', { headers: { Authorization: 'Bearer ' + tok } });
          if (r.ok) return await r.blob();
          if (r.status === 404) throw new HttpError('Microsoft (téléchargement)', 404, 'fichier introuvable');
          lastErr = new HttpError('Microsoft (téléchargement /content)', r.status, await safeText(r));
        } catch (e) {
          if (e instanceof PauseError || e.status === 404) throw e;
          lastErr = e;
        }
        // 2) lien pré-authentifié, réponse complète
        try {
          const meta = await graph(base);
          const url = meta['@microsoft.graph.downloadUrl'] || meta['@content.downloadUrl'];
          if (!url) throw new Error('Microsoft n\'a pas fourni de lien de téléchargement');
          const r = await fetchFn(url);
          if (r.ok) return await r.blob();
          lastErr = new HttpError('Microsoft (lien de téléchargement)', r.status, '');
        } catch (e) {
          if (e instanceof PauseError) throw e;
          lastErr = e;
        }
      }
      throw new Error('Téléchargement impossible : ' + ((lastErr && lastErr.message) || lastErr));
    }

    /* ---------- Exécution ---------- */
    async function run(source, opts = {}) {
      stopRequested = false;
      const t0 = now();
      const res = { paused: null, done: 0, failed: [], plan: null, skipped: [], walkErrors: [] };
      progress({ phase: 'start' });
      log('Lecture de OneDrive…');
      const w = await walk(source);
      res.skipped = w.skipped; res.walkErrors = w.errors;
      if (stopRequested) { res.paused = 'stop'; return res; }
      w.errors.forEach((e) => log('❌ Lecture impossible : ' + e.path + ' — ' + e.err));
      log(`OneDrive : ${w.items.filter((i) => i.folder).length} dossiers, ${w.items.filter((i) => !i.folder).length} fichiers.`);
      let index;
      try { index = await loadIndex(source); }
      catch (e) { if (e instanceof PauseError) { res.paused = e.pause; return res; } throw e; }
      let ver = { fix: { manquants: 0, deplacesGoogle: 0, contenuDifferent: 0 }, extras: [], extrasCount: 0 };
      log('Vérification de la copie Google…');
      try {
        const gt = await listGoogle(index.rootG);
        if (gt.complete) ver = verify(index, w, gt, source);
        else { log('⚠️ Lecture du miroir Google incomplète : vérification sautée cette fois-ci.'); gt.errors.forEach((e) => log('❌ Google : ' + e)); }
        res.verifyComplete = gt.complete;
      } catch (e) { if (e instanceof PauseError) { res.paused = e.pause; return res; } throw e; }
      if (stopRequested) { res.paused = 'stop'; return res; }
      const p = plan(w, index);
      p.verif = ver.fix;
      p.extras = (w.complete && res.verifyComplete) ? ver.extras : [];
      p.counts.enTropDansGoogle = p.extras.length ? ver.extrasCount : 0;
      p.counts.reparesGoogle = ver.fix.manquants + ver.fix.deplacesGoogle + ver.fix.contenuDifferent;
      res.plan = p;
      if (!w.complete) log('⚠️ Lecture incomplète : AUCUNE suppression ne sera reportée cette fois-ci.');
      if (opts.dryRun) return res;
      if (opts.confirm && !(await opts.confirm(p))) { res.paused = 'annulé'; return res; }

      let opsSinceSave = 0, lastSave = now(), pause = null;
      const total = p.folderOps.length + p.fileOps.length + p.topDels.length + p.extras.length;
      let bytesDone = 0;
      const current = new Set();
      const report = () => progress({ phase: 'transfer', done: res.done, total, failed: res.failed.length, bytesDone, bytesTotal: p.bytes, current: [...current] });
      const record = async (id, rec) => {
        if (rec) index.items[id] = rec; else delete index.items[id];
        journal.add(source.id, rec ? { id, rec } : { id, del: true });
        res.done++; opsSinceSave++;
        if (opsSinceSave >= SAVE_EVERY_OPS || now() - lastSave > SAVE_EVERY_MS) {
          opsSinceSave = 0; lastSave = now();
          await saveIndex(index, source);
        }
        report();
      };
      const fail = (op, e) => {
        res.failed.push({ path: op.it ? op.it.path : op.id, err: String((e && e.message) || e) });
        log('❌ ' + (op.it ? op.it.path : op.id) + ' — ' + ((e && e.message) || e));
      };

      // 1) Dossiers, un par un, dans l'ordre (le parent existe toujours avant l'enfant)
      for (const op of p.folderOps) {
        if (pause || stopRequested) break;
        const it = op.it;
        try {
          const pg = parentG(index, source, it.parent);
          if (!pg) throw new Error('dossier parent introuvable dans le miroir');
          if (op.k === 'mkdir') {
            const g = await gMkdir(it.name, pg);
            await record(it.id, { g, p: it.parent, n: it.name, f: 1 });
          } else {
            try {
              await gMove(op.ex.g, pg, it.name);
              await record(it.id, Object.assign({}, op.ex, { p: it.parent, n: it.name }));
            } catch (e) {
              if (e.status !== 404) throw e;
              // dossier effacé à la main côté Google : on le recrée et on oublie son contenu (recopié au prochain passage)
              const g = await gMkdir(it.name, pg);
              forgetDescendants(index, it.id);
              await record(it.id, { g, p: it.parent, n: it.name, f: 1 });
              log('⚠️ Dossier recréé dans Google : ' + it.path + ' — relancez une copie pour compléter son contenu.');
            }
          }
        } catch (e) {
          if (e instanceof PauseError) { pause = e.pause; break; }
          fail(op, e);
        }
      }

      // 2) Fichiers, plusieurs à la fois
      const queue = p.fileOps.slice();
      async function worker() {
        while (queue.length && !pause && !stopRequested) {
          const op = queue.shift();
          const it = op.it;
          current.add(it.path); report();
          try {
            const pg = parentG(index, source, it.parent);
            if (!pg) throw new Error('dossier parent absent du miroir (erreur plus haut ?)');
            if (op.k === 'move') {
              try {
                await gMove(op.ex.g, pg, it.name);
                await record(it.id, Object.assign({}, op.ex, { p: it.parent, n: it.name }));
              } catch (e) {
                if (e.status !== 404) throw e;
                const blob = await download(it);
                const h = await sha1Of(blob);
                const g = await gUpload(blob, { name: it.name, parents: [pg], appProperties: { od: it.id } });
                await record(it.id, { g, p: it.parent, n: it.name, f: 0, t: it.tag, s: it.size, h });
              }
            } else {
              await gToken(it.size > SMALL ? 900 : 300); // pas de démarrage d'un envoi avec un jeton presque expiré
              const blob = await download(it);
              const h = await sha1Of(blob);
              let g;
              if (op.k === 'update') {
                try {
                  if (op.moved) await gMove(op.ex.g, pg, it.name);
                  g = await gUpload(blob, null, op.ex.g);
                } catch (e) {
                  if (e.status !== 404) throw e;
                  g = await gUpload(blob, { name: it.name, parents: [pg], appProperties: { od: it.id } });
                }
              } else {
                g = await gUpload(blob, { name: it.name, parents: [pg], appProperties: { od: it.id } });
              }
              bytesDone += it.size;
              await record(it.id, { g, p: it.parent, n: it.name, f: 0, t: it.tag, s: it.size, h });
            }
          } catch (e) {
            if (e instanceof PauseError) { pause = e.pause; queue.unshift(op); }
            else fail(op, e);
          } finally { current.delete(it.path); report(); }
        }
      }
      if (!pause && !stopRequested) await Promise.all(Array.from({ length: TRANSFER_PARALLEL }, worker));

      // 3) Suppressions (seulement si tout le reste est passé et la lecture était complète)
      if (!pause && !stopRequested && w.complete && (p.topDels.length || p.extras.length)) {
        try {
          let binDated = null;
          const targets = p.topDels.map((id) => ({ id, g: index.items[id].g })).concat(p.extras.map((g) => ({ id: null, g })));
          for (const tgt of targets) {
            const id = tgt.id;
            const ex = { g: tgt.g };
            try {
              if (opts.delMode === 'trash') {
                await drive(GD + '/files/' + encodeURIComponent(ex.g) + '?fields=id', {
                  method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ trashed: true })
                });
              } else {
                if (!binDated) {
                  if (!index.binG || !(await gExists(index.binG))) index.binG = await gEnsureFolder(BIN_NAME, index.top);
                  const d = new Date(now());
                  const pad = (n) => String(n).padStart(2, '0');
                  binDated = await gMkdir(`${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}h${pad(d.getMinutes())}`, index.binG);
                }
                await gMove(ex.g, binDated);
              }
            } catch (e) { if (e.status !== 404) throw e; }
            // le dossier et tout ce qu'il contient sortent de l'index
            if (id) {
              const gone = [id].concat(descendants(index, id));
              for (const gid of gone) { delete index.items[gid]; journal.add(source.id, { id: gid, del: true }); }
            }
            res.done++; report();
          }
        } catch (e) {
          if (e instanceof PauseError) pause = e.pause; else { log('❌ Suppressions : ' + e.message); res.failed.push({ path: '(suppressions)', err: e.message }); }
        }
      }

      if (stopRequested && !pause) pause = 'stop';
      try { await saveIndex(index, source); }
      catch (e) { if (e instanceof PauseError) pause = pause || e.pause; else log('❌ Enregistrement de l\'index : ' + e.message + ' (le journal local le garde)'); }
      res.paused = pause;
      res.seconds = Math.round((now() - t0) / 1000);
      return res;
    }

    function descendants(index, id) {
      const kids = {};
      for (const [k, v] of Object.entries(index.items)) (kids[v.p] = kids[v.p] || []).push(k);
      const out = [], st = [id];
      while (st.length) { const c = st.pop(); for (const k of kids[c] || []) { out.push(k); st.push(k); } }
      return out;
    }
    function forgetDescendants(index, id) {
      for (const k of descendants(index, id)) { delete index.items[k]; journal.add(index.sourceId, { id: k, del: true }); }
    }

    return {
      run, walk, plan, loadIndex,
      stop() { stopRequested = true; },
      _graph: graph
    };
  }

  const api = { createEngine, PauseError, HttpError, ENGINE_VERSION, TOP_NAME, BIN_NAME };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.MiroirEngine = api;
})(typeof self !== 'undefined' ? self : globalThis);
