/**
 * Agenda Canin — partie « serveur » (Google Apps Script).
 *
 * Deux Google Sheets sont créés automatiquement dans le Drive :
 *  - « Agenda Canin — base » : les données brutes de l'appli (ne pas modifier à la main) ;
 *  - « Agenda Canin — cahier » : le fichier lisible (cours, livre des recettes…),
 *    remis à jour automatiquement, téléchargeable en Excel.
 * Chaque séance est aussi copiée dans Google Agenda.
 */

// Version de ce code. Le bouton « Mettre à jour » compare avec le dépôt public ci-dessous.
const VERSION = '2026.10.16';
const DEPOT_MAJ = 'https://raw.githubusercontent.com/EdenRaison/agenda-canin-versions/main/';

const TZ = 'Europe/Paris';
const FMT = "yyyy-MM-dd'T'HH:mm";

const TABLES = {
  Clients: ['id', 'nom', 'telephone', 'email', 'adresse', 'cp', 'ville', 'notes'],
  Chiens: ['id', 'clientId', 'nom', 'race', 'naissance', 'notes'],
  Forfaits: ['id', 'chienId', 'nbCours', 'libelle', 'date', 'note'],
  Seances: ['id', 'debut', 'duree', 'type', 'titre', 'statut', 'eventId', 'source', 'note'],
  Participations: ['id', 'seanceId', 'chienId', 'formule', 'statut', 'prix', 'mode', 'note'],
  Recettes: ['id', 'date', 'categorie', 'clientId', 'chienId', 'forfaitId', 'libelle', 'nbCours', 'mode', 'cheques', 'montant', 'note'],
};

// Périodes du livre des recettes, comme dans le fichier Excel d'origine
const PERIODES = [
  { mois: [2, 3, 4], nom: 'Février - Mars - Avril' },
  { mois: [5, 6, 7], nom: 'Mai - Juin - Juillet' },
  { mois: [8, 9, 10], nom: 'Août - Septembre - Octobre' },
  { mois: [11, 12, 1], nom: 'Novembre - Décembre - Janvier' },
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Agenda Canin')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ================= Fonctions appelées par l'appli ================= */

function getData() {
  return withLock_(() => {
    return pack_(readDb_());
  });
}

/**
 * Enregistre des changements envoyés par l'appli.
 * changes = { Table: [objets complets] }, deletes = { Table: [ids] }
 */
function save(changes, deletes) {
  return withLock_(() => {
    const db = readDb_();
    const touched = new Set();
    Object.keys(deletes || {}).forEach(t => {
      (deletes[t] || []).forEach(id => {
        const i = db[t].findIndex(r => r.id === id);
        if (i < 0) return;
        const r = db[t][i];
        if (t === 'Participations') touched.add(r.seanceId);
        if (t === 'Seances') { deleteEvent_(r.eventId); }
        if (r._row) db._deleted[t].push(r._row);
        db[t].splice(i, 1);
      });
    });
    Object.keys(changes || {}).forEach(t => {
      if (!TABLES[t]) throw new Error('Table inconnue : ' + t);
      (changes[t] || []).forEach(obj => {
        const cur = db[t].find(r => r.id === obj.id);
        const next = cur || { id: obj.id };
        TABLES[t].forEach(c => { if (c !== 'eventId' && obj[c] !== undefined) next[c] = obj[c] === null ? '' : String(obj[c]); });
        next._dirty = true;
        if (!cur) db[t].push(next);
        if (t === 'Seances') touched.add(next.id);
        if (t === 'Participations') touched.add(next.seanceId);
        if (t === 'Chiens' || t === 'Clients') {
          // le nom d'un chien ou d'un client a pu changer : on rafraîchit ses séances à venir
          const dogIds = t === 'Chiens' ? [next.id] : db.Chiens.filter(d => d.clientId === next.id).map(d => d.id);
          db.Participations.filter(p => dogIds.indexOf(p.chienId) >= 0).forEach(p => touched.add(p.seanceId));
        }
      });
    });
    const now = new Date();
    touched.forEach(id => {
      const s = db.Seances.find(x => x.id === id);
      // Les séances reprises des anciens fichiers (passées) ne sont pas recopiées dans Google Agenda
      if (s && (!s.source || toDate_(s.debut) > now)) syncEvent_(db, s);
    });
    flush_(db);
    majCahierSiBesoin_(db);
    return pack_(db);
  });
}

// Reconstruit le cahier si la dernière mise à jour date de plus d'une minute
function majCahierSiBesoin_(db) {
  const props = PropertiesService.getScriptProperties();
  const last = Number(props.getProperty('CAHIER_MAJ') || 0);
  if (Date.now() - last < 60000) return;
  try { rebuildCahier_(db); } catch (e) { /* le cahier sera refait à la prochaine mise à jour horaire */ }
}

/** Import unique de l'historique (fichier .json préparé à partir des fichiers Excel). */
function importHistorique(payload) {
  return withLock_(() => {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('IMPORT_FAIT')) throw new Error('L’historique a déjà été importé.');
    if (!payload || payload.format !== 'agenda-canin-import-1') throw new Error('Ce fichier n’est pas un fichier d’historique Agenda Canin.');
    const db = readDb_();
    Object.keys(TABLES).forEach(t => {
      const ids = {};
      db[t].forEach(r => { ids[r.id] = 1; });
      (payload.tables[t] || []).forEach(o => {
        if (ids[o.id]) return;
        const r = { _dirty: true };
        TABLES[t].forEach(c => { r[c] = o[c] === undefined || o[c] === null ? '' : String(o[c]); });
        db[t].push(r);
      });
    });
    flush_(db);
    props.setProperty('IMPORT_FAIT', new Date().toISOString());
    rebuildCahier_(db);
    return pack_(db);
  });
}

/** Applique un fichier de mise à jour préparé à partir de l'agenda papier (une seule fois par fichier). */
function appliquerMaj(payload) {
  if (!payload || payload.format !== 'agenda-canin-maj-1' || !payload.id) throw new Error('Ce fichier n’est pas un fichier de mise à jour Agenda Canin.');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('MAJ_' + payload.id)) throw new Error('Cette mise à jour a déjà été appliquée.');
  const out = save(payload.changes || {}, payload.deletes || {});
  props.setProperty('MAJ_' + payload.id, new Date().toISOString());
  return out;
}

/** Met à jour le cahier (Google Sheets lisible) et renvoie ses liens. */
function majCahier() {
  return withLock_(() => {
    const db = readDb_();
    rebuildCahier_(db);
    return liens_();
  });
}

/** Lancée toutes les heures (voir installer) : les cours passés deviennent « effectués » dans le cahier. */
function majAutomatique() {
  const db = readDb_();
  syncFromCalendar_(db);
  flush_(db);
  rebuildCahier_(db);
}

/** À lancer une fois à la main : crée les fichiers, demande les autorisations, programme la mise à jour horaire. */
function installer() {
  const db = readDb_();
  rebuildCahier_(db);
  CalendarApp.getDefaultCalendar().getName();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'majAutomatique') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('majAutomatique').timeBased().everyHours(1).create();
  Logger.log('Base : ' + base_().getUrl());
  Logger.log('Cahier : ' + cahier_().getUrl());
}

/* ================= Mise à jour de l'appli (bouton dans l'appli) ================= */
// L'appli lit la dernière version publiée sur un dépôt GitHub public qui ne contient que le code,
// remplace son propre code grâce à l'API Apps Script, puis republie le même lien.
// À activer une fois : script.google.com/home/usersettings → « API Google Apps Script » sur Activé.

/** Version installée + dernière version publiée (appelée à l'ouverture de l'appli). */
function infoVersion() {
  let derniere = null;
  try { derniere = JSON.parse(lireDepot_('version.json')); } catch (e) { /* pas de réseau : on réessaiera */ }
  return { actuelle: VERSION, derniere: derniere };
}

/** Installe la dernière version publiée, puis met à jour le lien de l'appli sur cette version. */
function mettreAJour() {
  const info = JSON.parse(lireDepot_('version.json'));
  if (info.version === VERSION) return { version: VERSION, deja: true, url: ScriptApp.getService().getUrl() };
  const code = lireDepot_('Code.gs'), index = lireDepot_('Index.html'), manifeste = lireDepot_('appsscript.json');
  if (code.indexOf('function doGet') < 0 || index.indexOf('</html>') < 0 || manifeste.indexOf('oauthScopes') < 0) {
    throw new Error('Les fichiers de la nouvelle version sont incomplets. Réessaie dans quelques minutes.');
  }
  const id = ScriptApp.getScriptId();
  const api = 'https://script.googleapis.com/v1/projects/' + id;
  // 1. Remplacer le code (en gardant d'éventuels autres fichiers du projet)
  const actuel = appelApi_('get', api + '/content');
  const fichiers = (actuel.files || []).filter(f => ['Code', 'Index', 'appsscript'].indexOf(f.name) < 0);
  fichiers.push({ name: 'appsscript', type: 'JSON', source: manifeste });
  fichiers.push({ name: 'Code', type: 'SERVER_JS', source: code });
  fichiers.push({ name: 'Index', type: 'HTML', source: index });
  appelApi_('put', api + '/content', { files: fichiers });
  // 2. Créer une version figée de ce code
  const v = appelApi_('post', api + '/versions', { description: 'Agenda Canin ' + info.version });
  // 3. Faire pointer le lien de l'appli (même adresse qu'avant) sur cette version
  const deploiement = idDeploiement_();
  appelApi_('put', api + '/deployments/' + deploiement, {
    deploymentConfig: { scriptId: id, versionNumber: v.versionNumber, manifestFileName: 'appsscript', description: 'Agenda Canin ' + info.version },
  });
  return { version: info.version, url: ScriptApp.getService().getUrl() };
}

function lireDepot_(nom) {
  const r = UrlFetchApp.fetch(DEPOT_MAJ + nom + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Impossible de lire la nouvelle version (' + nom + '). Réessaie plus tard.');
  return r.getContentText();
}

function appelApi_(methode, url, corps) {
  const options = { method: methode, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  if (corps) { options.contentType = 'application/json'; options.payload = JSON.stringify(corps); }
  const r = UrlFetchApp.fetch(url, options);
  const texte = r.getContentText();
  if (r.getResponseCode() >= 300) {
    if (texte.indexOf('Apps Script API') >= 0 || texte.indexOf('usersettings') >= 0) {
      throw new Error('Il faut d’abord autoriser les mises à jour : ouvre script.google.com/home/usersettings et mets « API Google Apps Script » sur Activé, puis réessaie.');
    }
    throw new Error('La mise à jour n’a pas pu se faire (erreur ' + r.getResponseCode() + '). L’appli actuelle continue de fonctionner normalement.');
  }
  return texte ? JSON.parse(texte) : {};
}

// L'adresse de l'appli ressemble à https://script.google.com/macros/s/IDENTIFIANT/exec
function idDeploiement_() {
  const morceaux = ScriptApp.getService().getUrl().split('/');
  const i = morceaux.indexOf('s');
  if (i < 0 || !morceaux[i + 1]) throw new Error('Ouvre l’appli depuis son lien habituel (celui qui finit par /exec) pour faire la mise à jour.');
  return morceaux[i + 1];
}

/* ================= Règles des cours (identiques à l'appli) ================= */

function statutPart_(p, s) {
  if (!s || s.statut === 'annule' || p.statut === 'annule') return 'annule';
  const fin = toDate_(s.debut).getTime() + (Number(s.duree) || 60) * 60000;
  return fin <= Date.now() ? 'fait' : 'avenir';
}

// Une séance individuelle de 30 min (ou moins) compte pour ½ cours du forfait
function poids_(s) { return s && s.type !== 'collectif' && (Number(s.duree) || 60) <= 30 ? 0.5 : 1; }
const PAYANT = { seance: 1, chiot: 1, detente: 1 };

// Répartit les cours « forfait » de chaque chien dans ses forfaits, dans l'ordre chronologique.
function allocation_(db) {
  const seances = index_(db.Seances);
  const res = {}; // participationId -> { forfaitId, n }
  const parDog = {};
  db.Participations.forEach(p => {
    const s = seances[p.seanceId];
    if (p.formule !== 'forfait' || statutPart_(p, s) === 'annule') return;
    (parDog[p.chienId] = parDog[p.chienId] || []).push(p);
  });
  Object.keys(parDog).forEach(dogId => {
    const list = parDog[dogId].sort((a, b) => seances[a.seanceId].debut.localeCompare(seances[b.seanceId].debut));
    const fs = db.Forfaits.filter(f => f.chienId === dogId).sort((a, b) => (a.date + a.id).localeCompare(b.date + b.id));
    let fi = 0, n = 0;
    list.forEach(p => {
      while (fi < fs.length && n >= (Number(fs[fi].nbCours) || 10)) { fi++; n = 0; }
      if (fi < fs.length) { n += poids_(seances[p.seanceId]); res[p.id] = { forfaitId: fs[fi].id, n: n }; } else res[p.id] = { forfaitId: '', n: 0 };
    });
  });
  return res;
}

/* ================= Cahier lisible (Google Sheets → Excel) ================= */

function rebuildCahier_(db) {
  PropertiesService.getScriptProperties().setProperty('CAHIER_MAJ', String(Date.now()));
  const ss = cahier_();
  const seances = index_(db.Seances), chiens = index_(db.Chiens), clients = index_(db.Clients);
  const alloc = allocation_(db);
  const dmy = s => s ? s.slice(8, 10) + '/' + s.slice(5, 7) + '/' + s.slice(0, 4) : '';
  const maj = 'Mis à jour automatiquement le ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy à HH:mm') + ' — ne pas modifier ici, tout se fait dans l’appli.';

  // 1. Cours au terrain : une ligne par forfait
  const rowsCours = db.Forfaits.slice().sort((a, b) => (a.date + a.id).localeCompare(b.date + b.id)).map(f => {
    const d = chiens[f.chienId] || {}, c = clients[d.clientId] || {};
    const parts = db.Participations.filter(p => alloc[p.id] && alloc[p.id].forfaitId === f.id)
      .map(p => ({ p: p, s: seances[p.seanceId] })).sort((a, b) => a.s.debut.localeCompare(b.s.debut));
    const ligne = x => dmy(x.s.debut) + (poids_(x.s) < 1 ? ' (½ h)' : '') + (statutPart_(x.p, x.s) === 'avenir' ? ' (prévu)' : '');
    const faits = parts.filter(x => statutPart_(x.p, x.s) === 'fait').reduce((a, x) => a + poids_(x.s), 0);
    const annules = db.Participations.filter(p => p.chienId === f.chienId && p.formule === 'forfait' && statutPart_(p, seances[p.seanceId]) === 'annule').length;
    return [c.nom || '', d.nom || '', d.race || '', dmy(d.naissance), f.nbCours + ' cours ' + (f.libelle || '').toLowerCase(), dmy(f.date),
      parts.filter(x => x.s.type === 'collectif').map(ligne).join('\n'),
      parts.filter(x => x.s.type !== 'collectif').map(ligne).join('\n'),
      faits, Math.max((Number(f.nbCours) || 10) - faits, 0), f.note || ''];
  });
  writeTab_(ss, 'Cours au terrain', 'Cours au terrain — forfaits', maj,
    ['Propriétaire', 'Chien', 'Race', 'Né(e) le', 'Forfait', 'Acheté le', 'Cours collectifs', 'Cours individuels', 'Effectués', 'Restants', 'Remarques'], rowsCours, [180, 110, 140, 90, 150, 90, 150, 150, 80, 80, 220]);

  // 2. Séances hors forfait (chiots, à la séance, offertes)
  const rowsHors = db.Participations.filter(p => p.formule !== 'forfait').map(p => ({ p: p, s: seances[p.seanceId] })).filter(x => x.s)
    .sort((a, b) => b.s.debut.localeCompare(a.s.debut)).map(x => {
      const d = chiens[x.p.chienId] || {}, c = clients[d.clientId] || {};
      const st = statutPart_(x.p, x.s);
      return [dmy(x.s.debut), x.s.source === 'import' ? '' : x.s.debut.slice(11, 16), d.nom || '', d.race || '', c.nom || '', x.s.type === 'collectif' ? (x.s.titre || 'Collectif') : 'Individuel',
        { seance: 'À la séance', chiot: 'École des chiots', detente: 'Détente', offert: 'Offert' }[x.p.formule] || x.p.formule, Number(x.p.prix) || '', x.p.mode || '', { fait: 'Effectué', avenir: 'Prévu', annule: 'Annulé' }[st]];
    });
  writeTab_(ss, 'Séances hors forfait', 'Séances hors forfait (chiots, détentes, à la séance, offertes)', maj,
    ['Date', 'Heure', 'Chien', 'Race', 'Propriétaire', 'Type', 'Formule', 'Prix (€)', 'Paiement', 'Statut'], rowsHors, [90, 60, 110, 140, 180, 120, 130, 70, 100, 90]);

  // 3. Livre des recettes, par période
  const recettes = recettesCompletes_(db, alloc, seances, chiens, clients);
  const blocs = {};
  recettes.forEach(r => {
    const m = Number(r.date.slice(5, 7)), y = Number(r.date.slice(0, 4));
    const pi = PERIODES.findIndex(p => p.mois.indexOf(m) >= 0);
    const annee = (pi === 3 && m === 1) ? y - 1 : y;
    const key = annee + '-' + pi;
    (blocs[key] = blocs[key] || { titre: PERIODES[pi].nom + ' ' + annee + (pi === 3 ? '/' + (annee + 1) : ''), rows: [] }).rows.push(r);
  });
  const livre = [];
  const totaux = [];
  Object.keys(blocs).sort().reverse().forEach(k => {
    const b = blocs[k];
    livre.push({ titre: b.titre });
    b.rows.sort((x, y) => y.date.localeCompare(x.date)).forEach(r => {
      const c = clients[r.clientId] || {}, d = chiens[r.chienId] || {};
      livre.push([dmy(r.date), c.nom || '', c.adresse || '', c.cp || '', c.ville || '', c.telephone || '', c.email || '', d.nom || '', d.race || '', dmy(d.naissance),
        r.categorie, r.libelle, r.mode, r.cheques, Number(r.montant) || 0, r.note || '']);
    });
    const tot = b.rows.reduce((a, r) => a + (Number(r.montant) || 0), 0);
    livre.push({ total: tot });
    const parCat = {};
    b.rows.forEach(r => { parCat[r.categorie] = (parCat[r.categorie] || 0) + (Number(r.montant) || 0); });
    totaux.push([b.titre, parCat['Forfait'] || 0, parCat['Cours'] || 0, parCat['École des chiots'] || 0, parCat['Cours collectif'] || 0, parCat['Détente'] || 0,
      (parCat['Friandises'] || 0) + (parCat['Harnais'] || 0), parCat['Autre'] || 0, tot]);
  });
  writeLivre_(ss, livre, maj);
  writeTab_(ss, 'Totaux', 'Totaux par période (€)', maj,
    ['Période', 'Forfaits', 'Cours à la séance', 'École des chiots', 'Cours collectifs', 'Détentes', 'Friandises / harnais', 'Autres', 'TOTAL'], totaux, [260, 90, 110, 110, 110, 90, 130, 80, 90]);

  // Nettoyage : ordre des onglets, suppression de la feuille vide d'origine
  ['Cours au terrain', 'Livre des recettes', 'Totaux', 'Séances hors forfait'].forEach((n, i) => { ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1); });
  ss.getSheets().forEach(s => { if (['Cours au terrain', 'Livre des recettes', 'Totaux', 'Séances hors forfait'].indexOf(s.getName()) < 0) ss.deleteSheet(s); });
}

// Recettes saisies + séances payées à l'unité (chiots, à la séance) effectuées
function recettesCompletes_(db, alloc, seances, chiens, clients) {
  const out = db.Recettes.slice();
  db.Participations.forEach(p => {
    if (!PAYANT[p.formule] || !(Number(p.prix) > 0)) return;
    const s = seances[p.seanceId];
    if (statutPart_(p, s) !== 'fait') return;
    const d = chiens[p.chienId] || {};
    out.push({ id: 'p-' + p.id, date: s.debut.slice(0, 10), categorie: { chiot: 'École des chiots', detente: 'Détente' }[p.formule] || (s.type === 'collectif' ? 'Cours collectif' : 'Cours'),
      clientId: d.clientId || '', chienId: p.chienId, libelle: { chiot: 'Cours chiot', detente: 'Détente' }[p.formule] || 'Cours à la séance' + (s.titre ? ' — ' + s.titre : ''),
      nbCours: '1', mode: p.mode || '', cheques: '', montant: p.prix, note: p.note || '' });
  });
  return out;
}

function writeTab_(ss, name, titre, maj, headers, rows, widths) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1).setValue(titre).setFontSize(14).setFontWeight('bold');
  sh.getRange(2, 1).setValue(maj).setFontColor('#6b7280').setFontStyle('italic');
  sh.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e1ece3').setWrap(true);
  if (rows.length) sh.getRange(5, 1, rows.length, headers.length).setValues(rows).setVerticalAlignment('top').setWrap(true);
  sh.setFrozenRows(4);
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function writeLivre_(ss, lignes, maj) {
  const headers = ['Date', 'Nom du client', 'Adresse', 'Code postal', 'Ville', 'Tél.', 'Mail', 'Nom du chien', 'Race', 'Date de naissance', 'Catégorie', 'Détail', 'Mode de paiement', 'N° chèque', 'Montant (€)', 'Remarques'];
  let sh = ss.getSheetByName('Livre des recettes');
  if (!sh) sh = ss.insertSheet('Livre des recettes');
  sh.clear();
  sh.getRange(1, 1).setValue('Livre chronologique des recettes').setFontSize(14).setFontWeight('bold');
  sh.getRange(2, 1).setValue(maj).setFontColor('#6b7280').setFontStyle('italic');
  sh.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e1ece3').setWrap(true);
  sh.setFrozenRows(4);
  const values = [], styles = [];
  lignes.forEach(l => {
    if (Array.isArray(l)) { values.push(l); styles.push(''); }
    else if (l.titre) { const r = new Array(headers.length).fill(''); r[0] = l.titre; values.push(r); styles.push('titre'); }
    else { const r = new Array(headers.length).fill(''); r[13] = 'TOTAL'; r[14] = l.total; values.push(r); styles.push('total'); values.push(new Array(headers.length).fill('')); styles.push(''); }
  });
  if (values.length) {
    sh.getRange(5, 1, values.length, headers.length).setValues(values).setVerticalAlignment('top');
    styles.forEach((st, i) => {
      if (st === 'titre') sh.getRange(5 + i, 1, 1, headers.length).setFontWeight('bold').setBackground('#f7ead2');
      if (st === 'total') sh.getRange(5 + i, 14, 1, 2).setFontWeight('bold');
    });
  }
  [90, 180, 180, 70, 130, 110, 190, 110, 130, 90, 110, 200, 110, 110, 90, 200].forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

/* ================= Google Agenda ================= */

function syncEvent_(db, s) {
  const cal = CalendarApp.getDefaultCalendar();
  const parts = db.Participations.filter(p => p.seanceId === s.id && p.statut !== 'annule');
  const ev = getEvent_(s.eventId);
  if (s.statut === 'annule' || !parts.length) {
    if (ev) { try { ev.deleteEvent(); } catch (e) {} }
    if (s.eventId) { s.eventId = ''; s._dirty = true; }
    return;
  }
  const chiens = index_(db.Chiens), clients = index_(db.Clients);
  const noms = parts.map(p => (chiens[p.chienId] || {}).nom || '?');
  const titre = s.type === 'collectif'
    ? '🐕 ' + (s.titre || 'Cours collectif') + ' — ' + noms.join(', ')
    : '🐕 ' + noms[0] + ' — ' + ((clients[(chiens[parts[0].chienId] || {}).clientId] || {}).nom || '') + (s.titre ? ' (' + s.titre + ')' : '');
  const desc = parts.map(p => {
    const d = chiens[p.chienId] || {}, c = clients[d.clientId] || {};
    return '• ' + d.nom + (d.race ? ' (' + d.race + ')' : '') + ' — ' + (c.nom || '') + (c.telephone ? ' — ' + c.telephone : '')
      + ({ forfait: '', seance: ' — à la séance', chiot: ' — chiot', detente: ' — détente', offert: ' — offert' }[p.formule] || '');
  }).join('\n') + '\n\nCréé par Agenda Canin';
  const start = toDate_(s.debut), end = new Date(start.getTime() + (Number(s.duree) || 60) * 60000);
  try {
    if (ev) { ev.setTime(start, end); ev.setTitle(titre); ev.setDescription(desc); }
    else { s.eventId = cal.createEvent(titre, start, end, { description: desc }).getId(); s._dirty = true; }
  } catch (e) { /* l'agenda n'est pas indispensable */ }
}

// Si une séance a été déplacée dans Google Agenda, l'appli suit.
// Si une séance À VENIR a été supprimée dans Google Agenda, elle passe en « annulée ».
function syncFromCalendar_(db) {
  const now = new Date();
  const from = new Date(now.getTime() - 30 * 864e5), to = new Date(now.getTime() + 400 * 864e5);
  let events;
  try { events = CalendarApp.getDefaultCalendar().getEvents(from, to); } catch (e) { return; }
  const byId = {};
  events.forEach(ev => { byId[ev.getId()] = ev; });
  db.Seances.forEach(s => {
    if (s.statut === 'annule' || !s.eventId) return;
    const start = toDate_(s.debut);
    if (start < from || start > to) return;
    const ev = byId[s.eventId];
    if (!ev) {
      if (start > now) { s.statut = 'annule'; s.eventId = ''; s.note = 'Supprimée dans Google Agenda'; s._dirty = true; }
      return;
    }
    const d = fmt_(ev.getStartTime());
    const du = String(Math.round((ev.getEndTime() - ev.getStartTime()) / 60000));
    if (d !== s.debut || du !== String(s.duree)) { s.debut = d; s.duree = du; s._dirty = true; }
  });
}

function getEvent_(eventId) {
  if (!eventId) return null;
  try { return CalendarApp.getDefaultCalendar().getEventById(eventId); } catch (e) { return null; }
}
function deleteEvent_(eventId) {
  const ev = getEvent_(eventId);
  if (ev) { try { ev.deleteEvent(); } catch (e) {} }
}

/* ================= Stockage (Google Sheets « base ») ================= */

function fichier_(prop, nom) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(prop);
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const ss = SpreadsheetApp.create(nom);
  props.setProperty(prop, ss.getId());
  return ss;
}
function base_() { return fichier_('BASE_ID', 'Agenda Canin — base (ne pas modifier)'); }
function cahier_() { return fichier_('CAHIER_ID', 'Agenda Canin — cahier'); }

function liens_() {
  const id = PropertiesService.getScriptProperties().getProperty('CAHIER_ID') || cahier_().getId();
  return { cahier: 'https://docs.google.com/spreadsheets/d/' + id + '/edit', excel: 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx' };
}

function sheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    const cols = TABLES[name];
    sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    ss.getSheets().forEach(s => { if (!TABLES[s.getName()] && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s); });
  }
  return sh;
}

function readDb_() {
  const ss = base_();
  const db = { _ss: ss, _deleted: {} };
  Object.keys(TABLES).forEach(t => {
    const cols = TABLES[t], sh = sheet_(ss, t), last = sh.getLastRow();
    db._deleted[t] = [];
    db[t] = last < 2 ? [] : sh.getRange(2, 1, last - 1, cols.length).getDisplayValues()
      .map((row, i) => { const o = { _row: i + 2 }; cols.forEach((c, j) => { o[c] = row[j]; }); return o; })
      .filter(o => o.id);
  });
  return db;
}

function flush_(db) {
  Object.keys(TABLES).forEach(t => {
    const cols = TABLES[t], sh = sheet_(db._ss, t);
    const line = r => cols.map(c => r[c] === undefined || r[c] === null ? '' : String(r[c]));
    db[t].filter(r => r._dirty && r._row).forEach(r => { sh.getRange(r._row, 1, 1, cols.length).setValues([line(r)]); r._dirty = false; });
    const add = db[t].filter(r => r._dirty && !r._row);
    if (add.length) {
      const start = sh.getLastRow() + 1;
      const rg = sh.getRange(start, 1, add.length, cols.length);
      rg.setNumberFormat('@');
      rg.setValues(add.map(line));
      add.forEach((r, i) => { r._row = start + i; r._dirty = false; });
    }
    db._deleted[t].sort((a, b) => b - a).forEach(row => sh.deleteRow(row));
    db._deleted[t] = [];
  });
}

function pack_(db) {
  const out = { meta: Object.assign(liens_(), { importFait: !!PropertiesService.getScriptProperties().getProperty('IMPORT_FAIT') }) };
  Object.keys(TABLES).forEach(t => {
    out[t] = db[t].map(r => { const o = {}; TABLES[t].forEach(c => { o[c] = r[c]; }); return o; });
  });
  return out;
}

/* ================= Outils ================= */

function index_(list) { const o = {}; list.forEach(x => { o[x.id] = x; }); return o; }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function toDate_(s) { return Utilities.parseDate(s, TZ, FMT); }
function fmt_(d) { return Utilities.formatDate(d, TZ, FMT); }
