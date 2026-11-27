/* Miroir OneDrive -> Google Drive — moteur de synchronisation
 * Fichier : engine.js — VERSION 3.0
 * Sens unique : OneDrive (LECTURE SEULE, jamais aucune écriture ni suppression) -> Google Drive.
 * À chaque passage, le miroir Google est relu et remis exactement à l'image de OneDrive.
 * Aucune dépendance externe. Testable hors navigateur (fetch / xhrPut injectés).
 */
(function (root) {
  'use strict';

  const ENGINE_VERSION = '3.0';
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
  const SAVE_EVERY_OPS = 300;   // l'index d'« Immobilier » pèse ~3 Mo : le réécrire tous les 25 fichiers enverrait ~2 Go inutiles
  const SAVE_EVERY_MS = 180000;
  const LOCK_MS = 10 * 60000;                             // un autre appareil actif il y a moins de 10 min = refus

  /* ---------- Contrôle d'intégrité des fichiers ---------- */
  const EXT_KIND = { jpg: 'jpeg', jpeg: 'jpeg', png: 'png', gif: 'gif', heic: 'heic', heif: 'heic', mov: 'video',
    mp4: 'video', m4v: 'video', '3gp': 'video', pdf: 'pdf', docx: 'zip', xlsx: 'zip', pptx: 'zip', zip: 'zip', doc: 'ole', xls: 'ole' };
  // en dessous de ces tailles, un fichier de ce type est anormalement petit (une photo d'iPhone pèse 1 à 5 Mo)
  const MIN_SIZE = { jpeg: 20480, png: 2048, gif: 512, heic: 20480, video: 10240, pdf: 1024, zip: 2048, ole: 2048 };
  const KIND_LABEL = { jpeg: 'une photo JPEG', png: 'une image PNG', gif: 'une image GIF', heic: 'une photo HEIC', video: 'une vidéo', pdf: 'un PDF', zip: 'un document Word/Excel', ole: 'un ancien document Word/Excel' };
  const kindOf = (name) => EXT_KIND[(String(name).split('.').pop() || '').toLowerCase()] || null;
  const bytesOf = (str) => Array.from(str, (c) => c.charCodeAt(0));
  function findIn(arr, pat) {
    outer: for (let i = 0; i <= arr.length - pat.length; i++) {
      for (let j = 0; j < pat.length; j++) if (arr[i + j] !== pat[j]) continue outer;
      return i;
    }
    return -1;
  }
  const startsWith = (arr, pat, off = 0) => pat.every((b, i) => arr[off + i] === b);
  function preview(arr) {
    const n = Math.min(arr.length, 120);
    let printable = 0;
    for (let i = 0; i < n; i++) { const b = arr[i]; if (b === 9 || b === 10 || b === 13 || (b >= 32 && b < 127) || b >= 160) printable++; }
    if (n && printable / n > 0.9) {
      let t = ''; try { t = new TextDecoder('utf-8', { fatal: false }).decode(arr.subarray(0, n)); } catch (e) { t = ''; }
      return 'contient du texte : « ' + t.replace(/\s+/g, ' ').trim().slice(0, 80) + ' »';
    }
    return 'commence par ' + Array.from(arr.subarray(0, Math.min(8, arr.length)), (b) => b.toString(16).padStart(2, '0')).join(' ');
  }
  /* Vérifie la structure d'un fichier. Renvoie { ok, level: 'abîmé' | 'à vérifier' | null, reason, preview } */
  async function checkBlob(name, blob) {
    const kind = kindOf(name);
    if (blob.size === 0) return { ok: false, level: 'abîmé', reason: 'fichier vide (0 octet)' };
    const head = new Uint8Array(await blob.slice(0, Math.min(blob.size, 1024)).arrayBuffer());
    const tailN = Math.min(blob.size, (kind === 'zip' || kind === 'jpeg') ? 65536 : 8192);
    const tail = new Uint8Array(await blob.slice(blob.size - tailN).arrayBuffer());
    const zero = (a) => a.length > 0 && a.every((b) => b === 0);
    if (zero(head.subarray(0, 64)) && zero(tail.subarray(Math.max(0, tail.length - 64)))) {
      return { ok: false, level: 'abîmé', reason: 'rempli de zéros', preview: preview(head) };
    }
    if (!kind) return { ok: true };
    const bad = (reason) => ({ ok: false, level: 'abîmé', reason, preview: preview(head) });
    const label = KIND_LABEL[kind];
    switch (kind) {
      case 'jpeg':
        if (!startsWith(head, [0xFF, 0xD8, 0xFF])) return bad('ce n\'est pas ' + label);
        if (findIn(tail, [0xFF, 0xD9]) < 0) return bad('photo coupée (la fin de l\'image manque)');
        break;
      case 'png':
        if (!startsWith(head, [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])) return bad('ce n\'est pas ' + label);
        if (findIn(tail, bytesOf('IEND')) < 0) return bad('image coupée (la fin de l\'image manque)');
        break;
      case 'gif':
        if (!startsWith(head, bytesOf('GIF8'))) return bad('ce n\'est pas ' + label);
        break;
      case 'heic':
        if (!startsWith(head, bytesOf('ftyp'), 4)) return bad('ce n\'est pas ' + label);
        break;
      case 'video':   // MP4 / MOV : le premier bloc peut être ftyp, mais aussi wide, mdat, moov, free… sur d'anciens fichiers QuickTime
        if (!['ftyp', 'wide', 'mdat', 'moov', 'free', 'skip', 'pnot'].some((a) => startsWith(head, bytesOf(a), 4))) return bad('ce n\'est pas ' + label);
        break;
      case 'pdf':
        if (findIn(head, bytesOf('%PDF')) < 0) return bad('ce n\'est pas un PDF');
        if (findIn(tail, bytesOf('%%EOF')) < 0) return bad('PDF coupé (la fin du fichier manque)');
        break;
      case 'zip':
        if (!startsWith(head, [0x50, 0x4B, 0x03, 0x04])) return bad('ce n\'est pas ' + label);
        if (findIn(tail, [0x50, 0x4B, 0x05, 0x06]) < 0) return bad('document coupé (la fin du fichier manque)');
        break;
      case 'ole':
        if (!startsWith(head, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1])) return bad('ce n\'est pas ' + label);
        break;
    }
    if (blob.size < MIN_SIZE[kind]) return { ok: false, level: 'abîmé', reason: 'taille impossible pour ' + label + ' (' + blob.size + ' octets)', preview: preview(head) };
    return { ok: true };
  }
  /* Repérage sans téléchargement, d'après la taille seule */
  function suspectBySize(name, size) {
    if (size === 0) return 'fichier vide (0 octet)';
    const kind = kindOf(name);
    if (kind && size < MIN_SIZE[kind]) return 'taille impossible pour ' + KIND_LABEL[kind] + ' (' + size + ' octets)';
    return null;
  }

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
    const myDevId = (env.device && env.device.id) || 'inconnu';

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
    async function gMkdir(name, parentId, odId) {
      const meta = { name, mimeType: FOLDER_MIME, parents: [parentId] };
      if (odId) meta.appProperties = { od: odId };   // permet de reconnaître ce dossier comme le sien
      const j = await drive(GD + '/files?fields=id', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(meta)
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

    /* Avant de remplacer un contenu : l'ancienne version Google est marquée « à conserver pour toujours » */
    async function keepCurrentVersion(fileId) {
      try {
        const f = await drive(GD + '/files/' + encodeURIComponent(fileId) + '?fields=headRevisionId');
        if (!f.headRevisionId) return false;
        await drive(GD + '/files/' + encodeURIComponent(fileId) + '/revisions/' + encodeURIComponent(f.headRevisionId) + '?fields=id,keepForever', {
          method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ keepForever: true })
        });
        return true;
      } catch (e) {
        if (e instanceof PauseError || e.status === 404) throw e;
        log('⚠️ Ancienne version non marquée « à conserver » (' + e.message + ')');
        return false;
      }
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
          const url = GD + '/files?pageSize=1000&fields=nextPageToken,files(id,name,mimeType,size,sha1Checksum,appProperties)&q=' +
            encodeURIComponent(`'${q(fid)}' in parents and trashed=false`) + (token ? '&pageToken=' + encodeURIComponent(token) : '');
          const j = await drive(url);
          for (const f of j.files || []) {
            const folder = f.mimeType === FOLDER_MIME;
            nodes.set(f.id, { id: f.id, name: f.name, folder, size: f.size != null ? Number(f.size) : null, sha1: String(f.sha1Checksum || '').toLowerCase(), parent: fid, ap: f.appProperties || {} });
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
      // Reprise des éléments que Miroir a lui-même copiés mais qui manquent à l'index
      // (ex. copie interrompue sur un autre appareil) : reconnus par l'identifiant OneDrive inscrit dessus.
      fix.adoptes = 0;
      const mappedNow = new Set(Object.values(index.items).map((e) => e.g));
      for (const node of gt.nodes.values()) {              // ordre de lecture : un dossier avant son contenu
        if (mappedNow.has(node.id)) continue;
        const od = node.ap && node.ap.od;
        const it = od && walkById.get(od);
        if (!it || index.items[od] || it.folder !== node.folder) continue;
        const parentOd = node.parent === index.rootG ? source.id : g2od.get(node.parent);
        if (parentOd === undefined) continue;
        if (node.folder) {
          index.items[od] = { g: node.id, p: parentOd, n: node.name, f: 1 };
        } else {
          // version de l'original connue et identique → rien à renvoyer ; sinon le contenu sera renvoyé une fois
          const sameVersion = node.ap.t && node.ap.t === String(it.tag).slice(0, 100) && node.size === it.size;
          index.items[od] = { g: node.id, p: parentOd, n: node.name, f: 0, t: sameVersion ? it.tag : '__repris__', s: node.size, h: node.sha1 };
        }
        g2od.set(node.id, od); mappedNow.add(node.id); fix.adoptes++;
      }
      const mapped = mappedNow;
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
        let url = `/drives/${encodeURIComponent(f.driveId)}/items/${encodeURIComponent(f.id)}/children?$top=200&$select=id,name,size,file,folder,cTag,eTag,remoteItem,package,lastModifiedDateTime`;
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
              mod: it.lastModifiedDateTime || (ri && ri.lastModifiedDateTime) || '',
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
        if (index.lock && index.lock.dev === myDevId) index.lock.at = index.lock.done ? 0 : now();   // ne touche jamais au verrou d'un autre appareil
        const data = JSON.stringify({ v: 1, sourceId: index.sourceId, sourceName: source.name, rootG: index.rootG, binG: index.binG, lock: index.lock || null, savedAt: new Date(now()).toISOString(), items: index.items });
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
    function plan(walkRes, index, corrupt) {
      corrupt = corrupt || new Set();
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
        } else if (corrupt.has(it.id)) {
          if (ex && moved) fileOps.push({ k: 'move', it, ex });   // corrompu : jamais copié ; seul un déplacement est suivi
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
      const res = { paused: null, done: 0, failed: [], plan: null, skipped: [], walkErrors: [], suspects: [], protectedCount: 0 };
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
      // Garde-fou : un autre appareil a-t-il copié ce dossier il y a moins de 10 minutes ?
      const dev = env.device || { id: 'inconnu', label: 'cet appareil' };
      if (index.lock && index.lock.dev !== dev.id && now() - index.lock.at < LOCK_MS) {
        res.paused = 'autre-appareil';
        res.lock = { label: index.lock.label, minutes: Math.round((now() - index.lock.at) / 60000) };
        log('⛔ ' + (index.lock.label || 'Un autre appareil') + ' a copié ce dossier il y a ' + res.lock.minutes + ' min : arrêt par sécurité.');
        return res;
      }
      let ver = { fix: { manquants: 0, deplacesGoogle: 0, contenuDifferent: 0 }, extras: [], extrasCount: 0 };
      log('Vérification de la copie Google…');
      try {
        const gt = await listGoogle(index.rootG);
        if (gt.complete) ver = verify(index, w, gt, source);
        else { log('⚠️ Lecture du miroir Google incomplète : vérification sautée cette fois-ci.'); gt.errors.forEach((e) => log('❌ Google : ' + e)); }
        res.verifyComplete = gt.complete;
      } catch (e) { if (e instanceof PauseError) { res.paused = e.pause; return res; } throw e; }
      if (stopRequested) { res.paused = 'stop'; return res; }
      // fichiers corrompus repérés à la taille (vide, taille impossible) : jamais copiés, jamais téléchargés
      const corruptFiles = w.items.filter((i) => !i.folder && suspectBySize(i.name, i.size));
      const p = plan(w, index, new Set(corruptFiles.map((i) => i.id)));
      res.sizeSuspects = corruptFiles.map((i) => ({ id: i.id, path: i.path, size: i.size, mod: i.mod, reason: suspectBySize(i.name, i.size), inGoogle: !!index.items[i.id] }));
      p.counts.corrompus = corruptFiles.length;
      p.verif = ver.fix;
      p.extras = (w.complete && res.verifyComplete) ? ver.extras : [];
      p.counts.enTropDansGoogle = p.extras.length ? ver.extrasCount : 0;
      p.counts.reparesGoogle = ver.fix.manquants + ver.fix.deplacesGoogle + ver.fix.contenuDifferent;
      p.counts.reprisDansIndex = ver.fix.adoptes || 0;
      res.plan = p;
      if (!w.complete) log('⚠️ Lecture incomplète : AUCUNE suppression ne sera reportée cette fois-ci.');
      if (opts.dryRun) return res;
      if (opts.confirm && !(await opts.confirm(p))) { res.paused = 'annulé'; return res; }
      index.lock = { dev: dev.id, label: dev.label, at: now() };
      await saveIndex(index, source);                      // prend la main et enregistre les éléments repris

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
            const g = await gMkdir(it.name, pg, it.id);
            await record(it.id, { g, p: it.parent, n: it.name, f: 1 });
          } else {
            try {
              await gMove(op.ex.g, pg, it.name);
              await record(it.id, Object.assign({}, op.ex, { p: it.parent, n: it.name }));
            } catch (e) {
              if (e.status !== 404) throw e;
              // dossier effacé à la main côté Google : on le recrée et on oublie son contenu (recopié au prochain passage)
              const g = await gMkdir(it.name, pg, it.id);
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
                const chk = await checkBlob(it.name, blob);
                if (!chk.ok) {
                  res.suspects.push({ id: it.id, path: it.path, size: it.size, mod: it.mod, reason: chk.reason, preview: chk.preview || '', inGoogle: false });
                  log('❌ CORROMPU, non copié : ' + it.path + ' — ' + chk.reason);
                  delete index.items[it.id]; journal.add(source.id, { id: it.id, del: true });
                  continue;
                }
                const h = await sha1Of(blob);
                const g = await gUpload(blob, { name: it.name, parents: [pg], appProperties: { od: it.id, t: String(it.tag).slice(0, 100) } });
                await record(it.id, { g, p: it.parent, n: it.name, f: 0, t: it.tag, s: it.size, h });
              }
            } else {
              await gToken(it.size > SMALL ? 900 : 300); // pas de démarrage d'un envoi avec un jeton presque expiré
              const blob = await download(it);
              const chk = await checkBlob(it.name, blob);
              if (!chk.ok) {
                // RÈGLE : un fichier corrompu n'est JAMAIS copié. Il est signalé ; une copie Google existante reste intacte.
                res.suspects.push({ id: it.id, path: it.path, size: it.size, mod: it.mod, reason: chk.reason, preview: chk.preview || '', inGoogle: !!op.ex });
                log('❌ CORROMPU, non copié : ' + it.path + ' — ' + chk.reason);
                if (op.k === 'update' && op.moved) {
                  await gMove(op.ex.g, pg, it.name);
                  await record(it.id, Object.assign({}, op.ex, { p: it.parent, n: it.name }));
                }
                if (op.ex) { res.protectedCount++; log('🛡️ La copie Google existante est conservée telle quelle : ' + it.path); }
                continue;
              }
              const h = await sha1Of(blob);
              let g;
              if (op.k === 'update') {
                try {
                  if (op.moved) await gMove(op.ex.g, pg, it.name);
                  if (!String(op.ex.t).startsWith('__')) await keepCurrentVersion(op.ex.g);   // vraie modification dans OneDrive
                  g = await gUpload(blob, null, op.ex.g);
                } catch (e) {
                  if (e.status !== 404) throw e;
                  g = await gUpload(blob, { name: it.name, parents: [pg], appProperties: { od: it.id, t: String(it.tag).slice(0, 100) } });
                }
              } else {
                g = await gUpload(blob, { name: it.name, parents: [pg], appProperties: { od: it.id, t: String(it.tag).slice(0, 100) } });
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
      if (index.lock) index.lock.done = true;               // la main est rendue
      try { await saveIndex(index, source); }
      catch (e) { if (e instanceof PauseError) pause = pause || e.pause; else log('❌ Enregistrement de l\'index : ' + e.message + ' (le journal local le garde)'); }
      res.paused = pause;
      res.seconds = Math.round((now() - t0) / 1000);
      return res;
    }

    /* Lecture de l'index sans rien créer dans Google */
    async function peekIndex(source) {
      const top = await gFindChild('root', TOP_NAME, true);
      if (!top) return null;
      const fileG = await gFindChild(top, '_index-' + source.id + '.json', false);
      if (!fileG) return null;
      const r = await drive(GD + '/files/' + encodeURIComponent(fileG) + '?alt=media', {}, 'response');
      const idx = JSON.parse(await r.text());
      for (const e of journal.load(source.id) || []) { if (e.root) continue; if (e.del) delete idx.items[e.id]; else idx.items[e.id] = e.rec; }
      return idx;
    }

    /* Retire de Google des fichiers corrompus copiés autrefois — UNIQUEMENT sur demande explicite (corbeille Google) */
    async function removeFromGoogle(source, odIds) {
      progress({ phase: 'remove', done: 0, total: odIds.length });
      const index = await loadIndex(source);
      const todo = odIds.filter((id) => index.items[id] && !index.items[id].f);
      let removed = 0, done = 0, i = 0, pauseErr = null;
      const failed = [];
      async function worker() {
        while (i < todo.length && !pauseErr) {
          const id = todo[i++];
          const ex = index.items[id];
          try {
            try {
              await drive(GD + '/files/' + encodeURIComponent(ex.g) + '?fields=id', {
                method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ trashed: true })
              });
            } catch (e) { if (e.status !== 404) throw e; }
            delete index.items[id]; journal.add(source.id, { id, del: true });
            removed++;
          } catch (e) {
            if (e instanceof PauseError) { pauseErr = e; break; }
            failed.push({ id, err: e.message }); log('❌ Retrait impossible (' + id + ') : ' + e.message);
          }
          done++; progress({ phase: 'remove', done, total: todo.length });
        }
      }
      await Promise.all(Array.from({ length: 4 }, worker));
      await saveIndex(index, source);
      log('🗑️ ' + removed + ' copie(s) corrompue(s) retirée(s) de Google (corbeille).');
      if (pauseErr) throw pauseErr;
      return { removed, failed };
    }

    /* ---------- Rapport d'intégrité de OneDrive (lecture seule) ---------- */
    async function integrity(source, opts = {}) {
      stopRequested = false;
      const t0 = now();
      log(opts.deep ? 'Contrôle complet de OneDrive…' : 'Contrôle rapide de OneDrive…');
      const w = await walk(source);
      const files = w.items.filter((i) => !i.folder);
      const out = { suspects: [], checked: 0, total: files.length, walkErrors: w.errors, complete: w.complete, deep: !!opts.deep };
      // Check files : tous les fichiers envoyés/modifiés depuis opts.sinceMs, lus en entier ;
      // rapide : seulement les fichiers anormalement petits ; complet : tous les fichiers d'un type contrôlable
      const recentMode = typeof opts.sinceMs === 'number';
      if (recentMode) opts.deep = true;
      const toCheck = files.filter((f) => recentMode ? (Date.parse(f.mod) >= opts.sinceMs)
        : opts.deep ? (kindOf(f.name) || f.size === 0) : suspectBySize(f.name, f.size));
      out.recent = recentMode ? [] : null;
      let i = 0, done = 0, bytes = 0;
      const totalBytes = toCheck.reduce((a, f) => a + f.size, 0);
      async function worker() {
        while (i < toCheck.length && !stopRequested) {
          const f = toCheck[i++];
          let r;
          try {
            const blob = f.size === 0 ? new Blob([]) : await download(f);
            r = await checkBlob(f.name, blob);
            if (!opts.deep && r.ok) { const why = suspectBySize(f.name, f.size); r = { ok: false, level: 'abîmé', reason: why }; }
          } catch (e) {
            if (e instanceof PauseError) { stopRequested = true; break; }
            r = { ok: false, level: 'illisible', reason: 'lecture impossible : ' + e.message };
          }
          if (!r.ok) out.suspects.push({ id: f.id, path: f.path, size: f.size, mod: f.mod, level: r.level, reason: r.reason, preview: r.preview || '' });
          if (out.recent) out.recent.push({ id: f.id, path: f.path, size: f.size, mod: f.mod, ok: !!r.ok, reason: r.ok ? '' : r.reason });
          done++; bytes += f.size; out.checked = done;
          progress({ phase: 'integrity', done, total: toCheck.length, bytesDone: bytes, bytesTotal: totalBytes, files: files.length });
        }
      }
      await Promise.all(Array.from({ length: TRANSFER_PARALLEL }, worker));
      out.stopped = stopRequested;
      out.suspects.sort((a, b) => a.path.localeCompare(b.path, 'fr'));
      if (out.recent) out.recent.sort((a, b) => (a.ok - b.ok) || a.path.localeCompare(b.path, 'fr'));
      try {
        const idx = await peekIndex(source);
        for (const s2 of out.suspects) s2.inGoogle = !!(idx && idx.items[s2.id]);
        out.googleKnown = !!idx;
      } catch (e) { if (e instanceof PauseError) out.googlePaused = true; else log('⚠️ Copie Google non consultée : ' + e.message); }
      out.seconds = Math.round((now() - t0) / 1000);
      log(`Contrôle terminé : ${out.suspects.length} fichier(s) suspect(s) sur ${files.length}.`);
      return out;
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
      run, walk, plan, loadIndex, integrity, removeFromGoogle,
      stop() { stopRequested = true; },
      _graph: graph
    };
  }

  const api = { createEngine, PauseError, HttpError, ENGINE_VERSION, TOP_NAME, BIN_NAME, checkBlob, suspectBySize };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.MiroirEngine = api;
})(typeof self !== 'undefined' ? self : globalThis);
