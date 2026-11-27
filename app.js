/* Miroir OneDrive -> Google Drive — interface et connexions
 * Fichier : app.js — VERSION 2.2
 */
(function () {
  'use strict';
  const APP_VERSION = '2.2';
  const DEFAULT_MS_CLIENT = ''; // inscription Entra DÉDIÉE « Miroir », avec la seule permission Files.Read : Microsoft refuse alors toute écriture dans OneDrive
  const MS_AUTH = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
  const MS_SCOPE = 'Files.Read offline_access';
  const G_SCOPE = 'https://www.googleapis.com/auth/drive'; // accès complet : nécessaire pour VOIR les fichiers ajoutés à la main dans la copie Google (l'app n'agit que dans « Miroir OneDrive »)
  const REDIRECT = location.origin + location.pathname.replace(/index\.html$/, '');
  const LS = {
    cfg: 'miroir.cfg', ms: 'miroir.ms', g: 'miroir.g', pkce: 'miroir.pkce', state: 'miroir.state',
    resume: 'miroir.resume', lastRenew: 'miroir.lastRenew', last: 'miroir.last'
  };
  const $ = (id) => document.getElementById(id);
  const get = (k, d) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } };
  const set = (k, v) => localStorage.setItem(k, JSON.stringify(v));
  const del = (k) => localStorage.removeItem(k);

  let cfg = Object.assign({ msClient: DEFAULT_MS_CLIENT, gClient: '', source: null, delMode: 'trash' }, get(LS.cfg, {}));
  const saveCfg = () => set(LS.cfg, cfg);
  // Identité de cet appareil (garde-fou : deux appareils ne copient jamais en même temps)
  const ua = navigator.userAgent || '';
  const DEVICE = {
    id: get('miroir.dev', null) || (() => { const v = 'D' + Math.random().toString(36).slice(2, 10); set('miroir.dev', v); return v; })(),
    label: /iPhone/.test(ua) ? 'iPhone' : (/iPad/.test(ua) || (/Macintosh/.test(ua) && navigator.maxTouchPoints > 1)) ? 'iPad' : 'Ordinateur'
  };
  let engine = null, running = false, wakeLock = null;

  /* ---------------- Journal ---------------- */
  function log(msg) {
    const box = $('log');
    const line = document.createElement('div');
    const t = new Date();
    line.textContent = t.toLocaleTimeString('fr-BE') + '  ' + msg;
    box.prepend(line);
    while (box.children.length > 300) box.lastChild.remove();
  }

  /* ---------------- Outils ---------------- */
  const b64url = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const rnd = (n) => b64url(crypto.getRandomValues(new Uint8Array(n)));
  const fmtBytes = (b) => b >= 1e12 ? (b / 1e12).toFixed(2) + ' To' : b >= 1e9 ? (b / 1e9).toFixed(2) + ' Go' : b >= 1e6 ? (b / 1e6).toFixed(1) + ' Mo' : (b >= 1e3 ? Math.round(b / 1e3) + ' Ko' : b + ' octets');

  /* ---------------- Microsoft (code + PKCE, sans librairie) ---------------- */
  async function msLogin() {
    if (!cfg.msClient) { alert('Indiquez d\'abord l\'identifiant d\'application Microsoft dans les Réglages.'); return; }
    const verifier = rnd(48);
    const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
    const state = 'ms-' + rnd(12);
    set(LS.pkce, verifier); set(LS.state, state);
    const p = new URLSearchParams({ client_id: cfg.msClient, response_type: 'code', redirect_uri: REDIRECT, scope: MS_SCOPE, code_challenge: challenge, code_challenge_method: 'S256', state, prompt: 'select_account' });
    location.href = MS_AUTH + '/authorize?' + p;
  }
  async function msTokenRequest(params) {
    const r = await fetch(MS_AUTH + '/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams(Object.assign({ client_id: cfg.msClient, scope: MS_SCOPE }, params)) });
    const j = await r.json();
    if (!r.ok) throw new Error(j.error_description || j.error || ('Microsoft ' + r.status));
    const old = get(LS.ms, {});
    const t = { access: j.access_token, refresh: j.refresh_token || old.refresh, exp: Date.now() + j.expires_in * 1000 };
    set(LS.ms, t);
    return t;
  }
  let msRefreshing = null; // une seule demande de rafraîchissement à la fois
  async function msToken(force) {
    const t = get(LS.ms, null);
    if (!t || !t.refresh) throw new Error('Microsoft non connecté');
    if (!force && t.exp - 120000 > Date.now()) return t.access;
    if (!msRefreshing) msRefreshing = msTokenRequest({ grant_type: 'refresh_token', refresh_token: t.refresh }).finally(() => { msRefreshing = null; });
    try { return (await msRefreshing).access; }
    catch (e) { throw new Error('Connexion Microsoft expirée — reconnectez Microsoft (' + e.message + ')'); }
  }

  /* ---------------- Google (jeton par redirection, sans librairie) ---------------- */
  function gLogin(silent) {
    if (!cfg.gClient) { alert('Indiquez d\'abord l\'identifiant client Google dans les Réglages.'); return; }
    const state = 'g-' + rnd(12);
    set(LS.state, state);
    const p = new URLSearchParams({ client_id: cfg.gClient, redirect_uri: REDIRECT, response_type: 'token', scope: G_SCOPE, include_granted_scopes: 'true', state });
    if (silent) p.set('prompt', 'none');
    const g = get(LS.g, null);
    if (g && g.email) p.set('login_hint', g.email);
    location.href = 'https://accounts.google.com/o/oauth2/v2/auth?' + p;
  }
  async function gToken(margin) {
    const g = get(LS.g, null);
    if (g && g.exp - (margin || 120) * 1000 > Date.now()) return g.access;
    throw new MiroirEngine.PauseError('google');
  }
  function gTokenQuiet() {
    const g = get(LS.g, null);
    return g && g.exp - 30000 > Date.now() ? g.access : null;
  }

  /* ---------------- Retour des pages de connexion ---------------- */
  async function handleRedirect() {
    const qs = new URLSearchParams(location.search);
    const hs = new URLSearchParams(location.hash.slice(1));
    const expected = get(LS.state, null);
    if (qs.get('state') && qs.get('state').startsWith('ms-')) {
      history.replaceState(null, '', REDIRECT);
      if (qs.get('state') !== expected) { log('❌ Retour Microsoft refusé (état inattendu)'); return; }
      if (qs.get('error')) { log('❌ Microsoft : ' + (qs.get('error_description') || qs.get('error'))); return; }
      try {
        await msTokenRequest({ grant_type: 'authorization_code', code: qs.get('code'), redirect_uri: REDIRECT, code_verifier: get(LS.pkce, '') });
        log('✅ Microsoft connecté');
      } catch (e) { log('❌ Microsoft : ' + e.message); }
      del(LS.pkce); del(LS.state);
    } else if (hs.get('state') && hs.get('state').startsWith('g-')) {
      history.replaceState(null, '', REDIRECT);
      if (hs.get('state') !== expected) { log('❌ Retour Google refusé (état inattendu)'); return; }
      del(LS.state);
      if (hs.get('error')) {
        log('Google : reconnexion manuelle nécessaire (' + hs.get('error') + ')');
        $('gNeed').hidden = false;
        return;
      }
      if (!String(hs.get('scope') || G_SCOPE).includes('auth/drive')) { log('❌ Google : l\'accès à Drive n\'a pas été accordé (cochez la case Drive).'); return; }
      const old = get(LS.g, {});
      set(LS.g, { access: hs.get('access_token'), exp: Date.now() + Number(hs.get('expires_in') || 3600) * 1000, email: old.email });
      log('✅ Google connecté');
      try {
        const r = await fetch('https://www.googleapis.com/drive/v3/about?fields=user(emailAddress),storageQuota', { headers: { Authorization: 'Bearer ' + hs.get('access_token') } });
        if (r.ok) {
          const j = await r.json();
          const g = get(LS.g, {}); g.email = j.user && j.user.emailAddress; g.quota = j.storageQuota; set(LS.g, g);
        }
      } catch (e) { /* facultatif */ }
    }
  }

  /* ---------------- Choix du dossier OneDrive (par identifiant) ---------------- */
  let pickStack = [];
  async function graphGet(url) {
    const r = await fetch(url.startsWith('http') ? url : 'https://graph.microsoft.com/v1.0' + url, { headers: { Authorization: 'Bearer ' + await msToken() } });
    if (!r.ok) throw new Error('Microsoft ' + r.status);
    return r.json();
  }
  async function openPicker() {
    try { await msToken(); } catch (e) { alert(e.message); return; }
    $('picker').hidden = false;
    pickStack = [{ name: 'Mon OneDrive', driveId: null, id: 'root' }];
    renderPicker();
  }
  async function renderPicker() {
    const cur = pickStack[pickStack.length - 1];
    $('pickPath').textContent = pickStack.map((f) => f.name).join(' › ');
    $('pickBack').disabled = pickStack.length < 2;
    $('pickChoose').disabled = cur.id === 'root';
    const list = $('pickList');
    list.innerHTML = '<div class="muted">Chargement…</div>';
    try {
      let url = cur.id === 'root'
        ? '/me/drive/root/children?$top=200&$select=id,name,folder,remoteItem,parentReference'
        : `/drives/${encodeURIComponent(cur.driveId)}/items/${encodeURIComponent(cur.id)}/children?$top=200&$select=id,name,folder,remoteItem,parentReference`;
      const folders = [];
      while (url) {
        const j = await graphGet(url);
        for (const it of j.value) {
          const ri = it.remoteItem;
          if (it.folder || (ri && ri.folder)) folders.push({
            name: it.name,
            driveId: ri ? ri.parentReference.driveId : it.parentReference.driveId,
            id: ri ? ri.id : it.id
          });
        }
        url = j['@odata.nextLink'];
      }
      if (cur.id === 'root' && folders[0]) cur.driveId = folders[0].driveId;
      folders.sort((a, b) => a.name.localeCompare(b.name, 'fr'));
      list.innerHTML = '';
      if (!folders.length) list.innerHTML = '<div class="muted">Aucun sous-dossier.</div>';
      for (const f of folders) {
        const b = document.createElement('button');
        b.className = 'row'; b.textContent = '📁  ' + f.name;
        b.onclick = () => { pickStack.push(f); renderPicker(); };
        list.appendChild(b);
      }
    } catch (e) { list.innerHTML = '<div class="err">' + e.message + '</div>'; }
  }
  function choosePicked() {
    const cur = pickStack[pickStack.length - 1];
    if (cfg.source && cfg.source.id !== cur.id && !confirm('Changer de dossier source ? Un nouveau miroir sera créé dans Google (l\'ancien reste en place).')) return;
    cfg.source = { driveId: cur.driveId, id: cur.id, name: cur.name, path: pickStack.slice(1).map((f) => f.name).join(' › ') };
    saveCfg(); $('picker').hidden = true; render();
  }

  /* ---------------- Envoi par morceaux (XHR : lit le statut 308 de Google) ---------------- */
  function xhrPut(url, headers, blob) {
    return new Promise((resolve, reject) => {
      const x = new XMLHttpRequest();
      x.open('PUT', url);
      for (const k of Object.keys(headers)) x.setRequestHeader(k, headers[k]);
      x.onload = () => {
        let json = null; try { json = JSON.parse(x.responseText); } catch (e) { /* vide */ }
        resolve({ status: x.status, json, range: x.getResponseHeader('Range') });
      };
      x.onerror = () => reject(new Error('réseau'));
      x.ontimeout = () => reject(new Error('délai dépassé'));
      x.timeout = 300000;
      x.send(blob);
    });
  }

  const journal = {
    load: (id) => get('miroir.journal.' + id, []),
    add: (id, e) => { const a = get('miroir.journal.' + id, []); a.push(e); set('miroir.journal.' + id, a); },
    clear: (id) => del('miroir.journal.' + id)
  };

  /* ---------------- Écran allumé pendant la copie ---------------- */
  async function keepAwake(on) {
    try {
      if (on && 'wakeLock' in navigator) wakeLock = await navigator.wakeLock.request('screen');
      if (!on && wakeLock) { await wakeLock.release(); wakeLock = null; }
    } catch (e) { /* non disponible : l'utilisateur garde l'écran allumé */ }
  }
  document.addEventListener('visibilitychange', () => { if (running && document.visibilityState === 'visible') keepAwake(true); });

  /* ---------------- Lancement ---------------- */
  let intMode = false, intT0 = 0;
  const elapsed = () => { const s2 = Math.round((Date.now() - intT0) / 1000); return Math.floor(s2 / 60) + ' min ' + String(s2 % 60).padStart(2, '0') + ' s'; };
  function progress(p) {
    if (intMode && p.phase === 'walk') {
      $('iText').textContent = `Lecture de OneDrive : ${p.folders} dossiers lus, ${p.items} éléments trouvés… (${elapsed()})`;
      $('iBar').style.width = '5%';
      return;
    }
    if (p.phase === 'walk') {
      $('pText').textContent = `Lecture de OneDrive : ${p.folders} dossiers lus, ${p.items} éléments trouvés…`;
      $('pBar').style.width = '5%';
    } else if (p.phase === 'remove') {
      $('iProgress').hidden = false;
      $('iBar').style.width = (p.total ? Math.max(3, Math.round(100 * p.done / p.total)) : 3) + '%';
      $('iText').textContent = p.done ? `Retrait de Google : ${p.done} / ${p.total}… (${elapsed()})` : 'Lecture de la liste Google… (' + elapsed() + ')';
    } else if (p.phase === 'integrity') {
      const pct = p.total ? Math.round(100 * p.done / p.total) : 100;
      $('iBar').style.width = Math.max(3, pct) + '%';
      $('iText').textContent = `Contrôle : ${p.done} / ${p.total} fichiers lus (${fmtBytes(p.bytesDone)} / ${fmtBytes(p.bytesTotal)}) sur ${p.files} fichiers au total… (${elapsed()})`;
    } else if (p.phase === 'verify') {
      $('pText').textContent = `Vérification de la copie Google : ${p.folders} dossiers relus, ${p.items} éléments…`;
      $('pBar').style.width = '8%';
    } else if (p.phase === 'transfer') {
      const pct = p.total ? Math.round(100 * p.done / p.total) : 100;
      $('pBar').style.width = Math.max(5, pct) + '%';
      $('pText').textContent = `${p.done} / ${p.total} opérations — ${fmtBytes(p.bytesDone)} / ${fmtBytes(p.bytesTotal)}` + (p.failed ? ` — ${p.failed} erreur(s)` : '');
      $('pCur').textContent = p.current.length ? '⇢ ' + p.current.join('\n⇢ ') : '';
    }
  }
  function planText(p) {
    const c = p.counts;
    return `Nouveaux dossiers : ${c.nouveauxDossiers}\nNouveaux fichiers : ${c.nouveauxFichiers}\nFichiers modifiés : ${c.modifies}\nRenommés / déplacés : ${c.deplaces}\nDisparus de OneDrive (à retirer de la copie) : ${c.supprimes}` +
      `\nDéjà à jour : ${c.dejaAJour}` +
      (c.corrompus ? `\n❌ Corrompus dans OneDrive, ne seront PAS copiés : ${c.corrompus}` : '') +
      `\nCopie Google à corriger (manquant, altéré, déplacé à la main) : ${c.reparesGoogle || 0}` +
      `\nEn trop dans Google (absent de OneDrive) : ${c.enTropDansGoogle || 0}` +
      ((c.supprimes || c.enTropDansGoogle) ? (cfg.delMode === 'trash' ? ' → corbeille Google' : ' → _Supprimés') : '') +
      (c.reprisDansIndex ? `\nDéjà copiés, repris dans la liste : ${c.reprisDansIndex}` : '') +
      `\nVolume à copier : ${fmtBytes(p.bytes)}\n\nOneDrive n'est jamais modifié.`;
  }
  function askConfirm(p) {
    return new Promise((resolve) => {
      $('cText').textContent = planText(p);
      const warn = (p.counts.supprimes + (p.counts.enTropDansGoogle || 0)) > 20 ? `⚠️ ${p.counts.supprimes + (p.counts.enTropDansGoogle || 0)} éléments vont être retirés de la copie Google. Vérifiez avant de continuer.` : '';
      $('cWarn').textContent = warn; $('cWarn').hidden = !warn;
      $('confirmBox').hidden = false;
      $('cGo').onclick = () => { $('confirmBox').hidden = true; resolve(true); };
      $('cNo').onclick = () => { $('confirmBox').hidden = true; resolve(false); };
    });
  }

  async function start(dryRun, auto) {
    if (running) return;
    if (!cfg.source) { alert('Choisissez d\'abord le dossier OneDrive.'); return; }
    try { await msToken(); } catch (e) { alert(e.message); return; }
    if (!gTokenQuiet()) { set(LS.resume, dryRun ? 'dry' : 'run'); gLogin(false); return; }
    running = true; render();
    $('progress').hidden = false; $('pBar').style.width = '2%'; $('pCur').textContent = '';
    if (!dryRun) await keepAwake(true);
    engine = MiroirEngine.createEngine({ fetch: (u, o) => fetch(u, o), xhrPut, msToken, gToken, gTokenQuiet, log, progress, journal, device: DEVICE });
    log(dryRun ? '— Analyse —' : (auto ? '— Reprise automatique —' : '— Copie —'));
    let res;
    try {
      res = await engine.run(cfg.source, {
        dryRun, delMode: cfg.delMode,
        confirm: (p) => (auto && (p.counts.supprimes + (p.counts.enTropDansGoogle || 0)) <= 20) ? Promise.resolve(true) : askConfirm(p)
      });
    } catch (e) {
      log('❌ ' + e.message); res = { error: e.message };
    }
    running = false; await keepAwake(false);
    finish(res, dryRun);
  }

  function finish(res, dryRun) {
    $('pCur').textContent = '';
    if (res.error) { $('pText').textContent = '❌ ' + res.error; render(); return; }
    if (dryRun && res.plan) {
      $('pBar').style.width = '100%';
      $('pText').textContent = 'Analyse terminée — rien n\'a été écrit.';
      $('summary').hidden = false; $('summary').textContent = planText(res.plan);
      if (res.walkErrors.length) $('summary').textContent += `\n\n⚠️ ${res.walkErrors.length} dossier(s) illisible(s) : les suppressions ne seraient pas reportées.`;
      if (res.sizeSuspects && res.sizeSuspects.length) {
        $('summary').textContent += `\n\n❌ ${res.sizeSuspects.length} fichier(s) déjà repérés comme corrompus (taille impossible) : voir « Intégrité ».`;
        $('iProgress').hidden = false; $('iText').textContent = 'Repérés pendant l\'analyse :'; $('iBar').style.width = '100%';
        showSuspects(res.sizeSuspects, 'Fichiers corrompus dans OneDrive (jamais copiés vers Google) :');
      }
      render(); return;
    }
    if (res.paused === 'google') {
      log('Jeton Google arrivé à expiration : renouvellement puis reprise automatique…');
      const last = Number(localStorage.getItem(LS.lastRenew) || 0);
      set(LS.resume, 'run');
      if (Date.now() - last > 120000) { localStorage.setItem(LS.lastRenew, String(Date.now())); gLogin(true); }
      else { $('gNeed').hidden = false; $('pText').textContent = 'En pause : reconnectez Google pour reprendre.'; }
      render(); return;
    }
    if (res.paused === 'autre-appareil') {
      $('pText').textContent = `⛔ ${res.lock.label || 'Un autre appareil'} a copié ce dossier il y a ${res.lock.minutes} min. Arrêtez-le, ou attendez 10 minutes après sa dernière activité, puis relancez.`;
      $('pBar').style.width = '0'; render(); return;
    }
    const sec = res.seconds || 0;
    const msg = res.paused === 'stop' ? '⏸ Arrêté. Relancez pour continuer là où la copie s\'est arrêtée.'
      : res.paused === 'annulé' ? 'Annulé, rien n\'a été fait.'
      : (res.failed.length ? `⚠️ Terminé avec ${res.failed.length} erreur(s) (voir le journal). Relancez pour réessayer.` : '✅ Miroir à jour.');
    $('pText').textContent = msg + (sec ? ` (${Math.floor(sec / 60)} min ${sec % 60} s)` : '');
    if (!res.paused) $('pBar').style.width = '100%';
    if (res.paused === 'annulé') $('pBar').style.width = '0';
    if (res.skipped && res.skipped.length) log(`ℹ️ ${res.skipped.length} élément(s) non copiables ignorés (ex. blocs-notes OneNote) : ` + res.skipped.slice(0, 5).join(', '));
    if (res.plan && !res.paused) set(LS.last, { at: new Date().toISOString(), failed: res.failed.length, counts: res.plan.counts });
    const corrupt = (res.suspects || []).concat((res.sizeSuspects || []).filter((x) => !(res.suspects || []).some((y) => y.id === x.id)));
    if (corrupt.length) {
      $('pText').textContent = $('pText').textContent.replace('✅ Miroir à jour.', '⚠️ Copie terminée, sauf les fichiers corrompus.') + ` — ❌ ${corrupt.length} fichier(s) corrompu(s) NON copié(s) : voir « Intégrité ».`;
      $('iProgress').hidden = false; $('iText').textContent = 'Repérés pendant la copie :'; $('iBar').style.width = '100%';
      showSuspects(corrupt.sort((a, b) => a.path.localeCompare(b.path, 'fr')), 'Fichiers corrompus dans OneDrive (jamais copiés vers Google) :');
    }
    render();
  }

  /* ---------------- Intégrité ---------------- */
  let lastSuspects = [];
  const fmtDate = (d) => d ? new Date(d).toLocaleDateString('fr-BE') : '';
  function showSuspects(list, title) {
    const key = 'miroir.suspects.' + (cfg.source ? cfg.source.id : '');
    const known = new Set(get(key, []));
    lastSuspects = list;
    const box = $('iList');
    box.innerHTML = '';
    $('iSummary').textContent = title;
    for (const x of list) {
      const d = document.createElement('div');
      d.className = 'susp';
      d.textContent = (known.has(x.id) ? '' : '🆕 ') + '❌ ' + x.path + '\n   ' + x.reason + ' — ' + fmtBytes(x.size || 0) + (x.mod ? ' — ' + fmtDate(x.mod) : '') +
        (x.inGoogle ? '\n   ⚠️ une copie existe dans Google' : '') + (x.preview ? '\n   ' + x.preview : '');
      box.appendChild(d);
    }
    set(key, list.map((x) => x.id));
    const inG = list.filter((x) => x.inGoogle && x.id);
    $('iRemove').hidden = !inG.length;
    $('iRemove').textContent = `Retirer de Google les ${inG.length} copie(s) corrompue(s)…`;
    $('iCopy').hidden = !list.length;
  }
  async function runIntegrity(deep) {
    if (running) return;
    if (!cfg.source) { alert('Choisissez d\'abord le dossier OneDrive.'); return; }
    try { await msToken(); } catch (e) { alert(e.message); return; }
    if (deep && !confirm('Le contrôle complet lit chaque fichier (environ ' + '74 Go pour « Immobilier 2025-2026 »). Il prendra plusieurs heures, écran allumé. Continuer ?')) return;
    running = true; intMode = true; intT0 = Date.now(); render(); await keepAwake(true);
    $('iProgress').hidden = false; $('iBar').style.width = '2%'; $('iText').textContent = 'Lecture de OneDrive…';
    engine = MiroirEngine.createEngine({ fetch: (u, o) => fetch(u, o), xhrPut, msToken, gToken, gTokenQuiet, log, progress, journal, device: DEVICE });
    try {
      const r = await engine.integrity(cfg.source, { deep });
      $('iBar').style.width = '100%';
      $('iText').textContent = (r.stopped ? '⏸ Contrôle arrêté. ' : '') + `${deep ? 'Contrôle complet' : 'Contrôle rapide'} : ${r.suspects.length} fichier(s) corrompu(s) sur ${r.total}.` +
        (r.walkErrors.length ? ` ⚠️ ${r.walkErrors.length} dossier(s) illisible(s).` : '') + (r.googlePaused ? ' (Google non consulté : reconnectez Google pour savoir lesquels y sont copiés.)' : '');
      showSuspects(r.suspects, r.suspects.length ? 'Fichiers corrompus dans OneDrive (jamais copiés vers Google) :' : '✅ Aucun fichier corrompu trouvé.');
    } catch (e) { $('iText').textContent = '❌ ' + e.message; log('❌ ' + e.message); }
    running = false; intMode = false; await keepAwake(false); render();
  }
  async function removeCorrupt() {
    const inG = lastSuspects.filter((x) => x.inGoogle && x.id);
    if (!inG.length) return;
    const sample = inG.slice(0, 5).map((x) => '• ' + x.path).join('\n');
    if (!confirm(`Retirer de Google ces ${inG.length} copie(s) corrompue(s) ?\n\n${sample}${inG.length > 5 ? '\n…' : ''}\n\nElles iront dans la corbeille Google (récupérables 30 jours). OneDrive n'est pas touché.`)) return;
    if (!gTokenQuiet()) { alert('Reconnectez d\'abord Google (bouton Connecter), puis recommencez.'); return; }
    running = true; intMode = true; intT0 = Date.now(); render();
    engine = MiroirEngine.createEngine({ fetch: (u, o) => fetch(u, o), xhrPut, msToken, gToken, gTokenQuiet, log, progress, journal, device: DEVICE });
    try {
      const r = await engine.removeFromGoogle(cfg.source, inG.map((x) => x.id));
      const failedIds = new Set(r.failed.map((f) => f.id));
      for (const x of inG) if (!failedIds.has(x.id)) x.inGoogle = false;
      showSuspects(lastSuspects, $('iSummary').textContent);
      $('iBar').style.width = '100%';
      $('iText').textContent = `🗑️ ${r.removed} copie(s) corrompue(s) retirée(s) de Google (corbeille).` + (r.failed.length ? ` ❌ ${r.failed.length} échec(s) : voir le Journal, puis réessayez.` : '');
    } catch (e) {
      $('iText').textContent = e instanceof MiroirEngine.PauseError ? '⏸ Connexion Google expirée : touchez « Connecter » (Google), puis recommencez. Ce qui est fait reste fait.' : '❌ ' + e.message;
      log('❌ ' + e.message);
    }
    running = false; intMode = false; render();
  }

  /* ---------------- Affichage ---------------- */
  function render() {
    const ms = get(LS.ms, null), g = get(LS.g, null);
    $('msState').textContent = ms && ms.refresh ? '✅ connecté' : '— non connecté';
    const gOk = !!gTokenQuiet();
    $('gState').textContent = gOk ? '✅ connecté' + (g.email ? ' (' + g.email + ')' : '') : (g && g.access ? '⏳ à renouveler (se fait seul au lancement)' : '— non connecté');
    if (g && g.quota && g.quota.limit) $('gQuota').textContent = `Espace Google : ${fmtBytes(Number(g.quota.usage))} utilisés sur ${fmtBytes(Number(g.quota.limit))}`;
    $('srcPath').textContent = cfg.source ? (cfg.source.path || cfg.source.name) : 'Aucun dossier choisi';
    $('dstPath').textContent = 'Mon Drive › ' + MiroirEngine.TOP_NAME + ' › ' + (cfg.source ? cfg.source.name : '…');
    const last = get(LS.last, null);
    $('lastRun').textContent = last ? 'Dernière copie complète : ' + new Date(last.at).toLocaleString('fr-BE') + (last.failed ? ` (${last.failed} erreur(s))` : '') : 'Aucune copie terminée pour l\'instant.';
    $('btnRun').disabled = running; $('btnDry').disabled = running;
    $('iQuick').disabled = running; $('iDeep').disabled = running; $('iRemove').disabled = running;
    $('btnStop').hidden = !running;
    $('msClient').value = cfg.msClient; $('gClient').value = cfg.gClient;
    document.querySelectorAll('input[name=delMode]').forEach((r) => { r.checked = r.value === cfg.delMode; });
    $('redirect').textContent = REDIRECT;
  }

  function bind() {
    $('btnMs').onclick = msLogin;
    $('btnG').onclick = () => gLogin(false);
    $('gNeedBtn').onclick = () => { set(LS.resume, 'run'); gLogin(false); };
    $('btnPick').onclick = openPicker;
    $('pickBack').onclick = () => { pickStack.pop(); renderPicker(); };
    $('pickChoose').onclick = choosePicked;
    $('pickClose').onclick = () => { $('picker').hidden = true; };
    $('btnRun').onclick = () => start(false, false);
    $('btnDry').onclick = () => start(true, false);
    $('btnStop').onclick = () => { if (engine) { engine.stop(); log('Arrêt demandé : les fichiers en cours se terminent…'); } };
    $('iQuick').onclick = () => runIntegrity(false);
    $('iDeep').onclick = () => runIntegrity(true);
    $('iRemove').onclick = removeCorrupt;
    $('iCopy').onclick = async () => {
      const txt = lastSuspects.map((x) => `${x.path} — ${x.reason} — ${fmtBytes(x.size || 0)}${x.mod ? ' — ' + fmtDate(x.mod) : ''}${x.inGoogle ? ' — copie dans Google' : ''}`).join('\n');
      try { await navigator.clipboard.writeText(txt); $('iCopy').textContent = 'Liste copiée ✓'; setTimeout(() => { $('iCopy').textContent = 'Copier la liste'; }, 2000); }
      catch (e) { prompt('Liste à copier :', txt); }
    };
    $('btnSave').onclick = () => {
      cfg.msClient = $('msClient').value.trim(); cfg.gClient = $('gClient').value.trim();
      saveCfg(); render(); log('Réglages enregistrés');
    };
    document.querySelectorAll('input[name=delMode]').forEach((r) => { r.onchange = () => { cfg.delMode = r.value; saveCfg(); }; });
    $('btnCopy').onclick = async () => { try { await navigator.clipboard.writeText(REDIRECT); $('btnCopy').textContent = 'Copié ✓'; } catch (e) { prompt('Adresse à copier :', REDIRECT); } };
    $('btnLogout').onclick = () => { if (confirm('Oublier les connexions Microsoft et Google sur cet iPhone ?')) { del(LS.ms); del(LS.g); render(); } };
    window.addEventListener('beforeunload', (e) => { if (running) { e.preventDefault(); e.returnValue = ''; } });
  }

  window.addEventListener('DOMContentLoaded', async () => {
    $('ver').textContent = 'v' + APP_VERSION + ' · moteur v' + MiroirEngine.ENGINE_VERSION;
    bind();
    await handleRedirect();
    render();
    const resume = localStorage.getItem(LS.resume);
    if (resume && gTokenQuiet() && get(LS.ms, null)) {
      del(LS.resume);
      start(JSON.parse(resume) === 'dry', JSON.parse(resume) === 'run');
    }
  });
})();
