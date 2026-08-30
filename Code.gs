/**
 * ============================================================================
 * Code.gs — Backend partagé (Google Sheets) pour l'app de collecte FS
 * ============================================================================
 * À coller dans l'éditeur Apps Script lié à un Google Sheets (voir README.md
 * fourni à part pour les étapes de déploiement pas-à-pas).
 *
 * Rôle de ce script :
 *  - Il reçoit les requêtes de l'appli (action=pull / action=push).
 *  - Il lit/écrit dans les feuilles du classeur.
 *  - Il filtre ce que chaque compte a le droit de voir/modifier (admin, chef
 *    d'équipe "user", ou gestionnaire de zone "zsmanager").
 *
 * Aucun secret n'est écrit en clair ici : ils sont générés et stockés dans
 * les "Propriétés du script" (Project Settings > Script properties), donc
 * jamais visibles dans ce fichier ni dans l'appli une fois configurés.
 * ============================================================================
 */

const SHEET_NAMES = {
  centres: 'Centres',
  utilisateurs: 'Utilisateurs',
  zsmanagers: 'ZS_Managers',
  parametres: 'Parametres',
  indicateurs: 'Donnees_Indicateurs',
  sources: 'Donnees_Sources',
  synthese: 'Donnees_Synthese',
  plan: 'Donnees_Plan',
  evaluateurs: 'Donnees_Evaluateurs',
  meta: 'Donnees_Meta',
  photos: 'Photos_Titres'
};

const SHEET_HEADERS = {
  centres: ['ID','Ordre','Nom','Type','Departement','Zone','Commune','Date','OverridesJSON','OverridesTouchedJSON','AdminSourcesJSON','AdminSourcesTouchedJSON','Supprime','ModifieLe'],
  utilisateurs: ['ID','Equipe','Identifiant','MotDePasse','CentreIDs','CreeParZone','Supprime','ModifieLe'],
  zsmanagers: ['ID','Zone','Identifiant','MotDePasse','Supprime','ModifieLe'],
  parametres: ['Cle','Valeur','ModifieLe'],
  indicateurs: ['CentreID','IndicateurID','Numerateur','Denominateur','ValeurOuiNon','Touche','ModifieLe'],
  sources: ['CentreID','SourceIndex','SIMR','DCM','DCN','D5','Touche','ModifieLe'],
  synthese: ['CentreID','Domaine','Ligne','PointFort','PointAmeliorer','Recommandation','ModifieLe'],
  plan: ['CentreID','Ligne','Domaine','IndicateurID','PointAmeliorer','Cause','Activite','Cout','ModifieLe'],
  evaluateurs: ['CentreID','Role','Nom','Qualification','Contact','ModifieLe'],
  meta: ['CentreID','Verrouille','Envoye','EnvoyeLe','ModifieLe'],
  photos: ['CentreID','PhotoKey','Titre','ModifieLe']
};

// Colonne(s) qui identifient une ligne de façon unique dans chaque feuille
// (utilisées pour savoir si on doit mettre à jour une ligne existante ou en
// ajouter une nouvelle).
const SHEET_KEYS = {
  centres: ['ID'],
  utilisateurs: ['ID'],
  zsmanagers: ['ID'],
  parametres: ['Cle'],
  indicateurs: ['CentreID','IndicateurID'],
  sources: ['CentreID','SourceIndex'],
  synthese: ['CentreID','Domaine','Ligne'],
  plan: ['CentreID','Ligne'],
  evaluateurs: ['CentreID','Role'],
  meta: ['CentreID'],
  photos: ['CentreID','PhotoKey']
};

/**
 * À exécuter UNE SEULE FOIS depuis l'éditeur Apps Script (menu "Exécuter" >
 * choisir "setup"), avant le premier déploiement. Crée les feuilles
 * manquantes avec leurs en-têtes, et génère un secret + une clé admin
 * aléatoires si ce n'est pas déjà fait. Les valeurs générées s'affichent
 * dans "Affichage > Journaux d'exécution" (Logs).
 */
function setup(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEET_NAMES).forEach(key => {
    const name = SHEET_NAMES[key];
    let sheet = ss.getSheetByName(name);
    if(!sheet) sheet = ss.insertSheet(name);
    if(sheet.getLastRow() === 0){
      sheet.appendRow(SHEET_HEADERS[key]);
      sheet.setFrozenRows(1);
    }
  });
  const props = PropertiesService.getScriptProperties();
  if(!props.getProperty('SECRET')){
    props.setProperty('SECRET', Utilities.getUuid().replace(/-/g,'').slice(0,16));
  }
  if(!props.getProperty('ADMIN_KEY')){
    props.setProperty('ADMIN_KEY', Utilities.getUuid().replace(/-/g,'').slice(0,20));
  }
  Logger.log('SECRET (à coller dans DEFAULT_BACKEND_SECRET) = ' + props.getProperty('SECRET'));
  Logger.log('ADMIN_KEY (à coller dans DEFAULT_BACKEND_ADMIN_KEY, admin uniquement) = ' + props.getProperty('ADMIN_KEY'));
}

/* ============================== Entrées HTTP ============================== */

function doGet(e){
  return handleAction_(e.parameter || {});
}
function doPost(e){
  let body = {};
  try{ body = JSON.parse(e.postData.contents); }catch(err){ /* body vide/invalide */ }
  return handleAction_(body);
}

function handleAction_(p){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(30000);
  }catch(err){
    return jsonOut_({ok:false, error:'Serveur occupé, réessayez.'});
  }
  try{
    if(p.action === 'pull') return jsonOut_(handlePull_(p));
    if(p.action === 'push') return jsonOut_(handlePush_(p));
    return jsonOut_({ok:false, error:'action inconnue'});
  } catch(err){
    return jsonOut_({ok:false, error:String(err)});
  } finally {
    lock.releaseLock();
  }
}

function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================== Authentification ============================== */
// Détermine QUI fait la demande et ce à quoi cette personne a droit.
// Le client (le navigateur) n'est jamais cru sur parole : tout est revérifié ici.
function resolveAuth_(p){
  const props = PropertiesService.getScriptProperties();
  const secret = props.getProperty('SECRET');
  if(!secret || p.secret !== secret) return {ok:false, error:'secret invalide'};

  if(p.role === 'admin'){
    const adminKey = props.getProperty('ADMIN_KEY');
    if(adminKey && p.akey === adminKey) return {ok:true, role:'admin'};
    return {ok:false, error:'clé admin invalide'};
  }
  if(p.role === 'user'){
    const users = getSheetRows_('utilisateurs');
    const u = users.find(r => r.Identifiant === p.uid && String(r.MotDePasse) === String(p.upass) && r.Supprime !== '1');
    if(!u) return {ok:false, error:'identifiants invalides'};
    return {ok:true, role:'user', centerIds: String(u.CentreIDs||'').split(',').filter(Boolean), user:u};
  }
  if(p.role === 'zsmanager'){
    const zs = getSheetRows_('zsmanagers');
    const z = zs.find(r => r.Identifiant === p.zid && String(r.MotDePasse) === String(p.zpass) && r.Supprime !== '1');
    if(!z) return {ok:false, error:'identifiants invalides'};
    return {ok:true, role:'zsmanager', zone: z.Zone, manager:z};
  }
  return {ok:false, error:'rôle inconnu ou manquant'};
}

/* ============================== PULL (lecture) ============================== */

function handlePull_(p){
  const auth = resolveAuth_(p);
  if(!auth.ok) return {error: auth.error};
  const since = Number(p.since || 0);

  const allCentres = getSheetRows_('centres');
  let accessibleCentres;
  if(auth.role === 'admin'){
    accessibleCentres = allCentres;
  } else if(auth.role === 'user'){
    accessibleCentres = allCentres.filter(c => auth.centerIds.includes(c.ID));
  } else { // zsmanager
    accessibleCentres = allCentres.filter(c => c.Zone === auth.zone);
  }
  const cIds = accessibleCentres.map(c => c.ID);

  const tables = {};
  tables.centres = filterSince_(accessibleCentres, since);

  if(auth.role === 'admin'){
    tables.utilisateurs = filterSince_(getSheetRows_('utilisateurs'), since);
    tables.zsmanagers = filterSince_(getSheetRows_('zsmanagers'), since);
    tables.parametres = filterSince_(getSheetRows_('parametres'), since);
  } else if(auth.role === 'user'){
    // Un compte "user" ne voit que sa propre fiche, jamais celle des autres.
    tables.utilisateurs = filterSince_(getSheetRows_('utilisateurs').filter(u => u.Identifiant === auth.user.Identifiant), since);
    // Paramètres non sensibles uniquement (ex. période évaluée) : AdminPass reste
    // réservé à l'administrateur et n'est jamais transmis aux équipes/zones.
    tables.parametres = filterSince_(getSheetRows_('parametres').filter(r => r.Cle !== 'AdminPass'), since);
  } else { // zsmanager
    tables.utilisateurs = filterSince_(getSheetRows_('utilisateurs').filter(u => u.CreeParZone === auth.zone), since);
    tables.zsmanagers = filterSince_(getSheetRows_('zsmanagers').filter(z => z.Identifiant === auth.manager.Identifiant), since);
    tables.parametres = filterSince_(getSheetRows_('parametres').filter(r => r.Cle !== 'AdminPass'), since);
  }

  ['indicateurs','sources','synthese','plan','evaluateurs','meta','photos'].forEach(key => {
    const rows = getSheetRows_(key).filter(r => cIds.includes(r.CentreID));
    tables[key] = filterSince_(rows, since);
  });

  return {ok:true, serverTime: Date.now(), tables};
}

function filterSince_(rows, since){
  if(!since) return rows;
  return rows.filter(r => Number(r.ModifieLe || 0) > since);
}

/* ============================== PUSH (écriture) ============================== */

function handlePush_(p){
  const auth = resolveAuth_(p);
  if(!auth.ok) return {error: auth.error};
  const updates = p.updates || {};
  const now = Date.now();

  const allowedCenterIds = (() => {
    if(auth.role === 'admin') return null; // null = toutes
    if(auth.role === 'user') return auth.centerIds;
    const allCentres = getSheetRows_('centres');
    return allCentres.filter(c => c.Zone === auth.zone).map(c => c.ID);
  })();
  const canTouchCenter = cid => allowedCenterIds === null || allowedCenterIds.includes(cid);

  // Réservé à l'admin : liste des FS et mot de passe admin global.
  ['centres','parametres'].forEach(key => {
    if(auth.role === 'admin' && updates[key] && updates[key].length){
      updates[key].forEach(row => upsertRow_(key, row, now));
    }
  });

  // Comptes utilisateurs : admin gère tout, un gestionnaire de zone ne peut
  // créer/modifier que les comptes qu'il a lui-même créés dans sa zone.
  if(updates.utilisateurs && updates.utilisateurs.length){
    updates.utilisateurs.forEach(row => {
      if(auth.role === 'admin') upsertRow_('utilisateurs', row, now);
      else if(auth.role === 'zsmanager' && row.CreeParZone === auth.zone) upsertRow_('utilisateurs', row, now);
    });
  }
  if(updates.zsmanagers && updates.zsmanagers.length && auth.role === 'admin'){
    updates.zsmanagers.forEach(row => upsertRow_('zsmanagers', row, now));
  }

  // Données de terrain : autorisé seulement pour les FS accessibles au compte.
  ['indicateurs','sources','synthese','plan','evaluateurs','meta','photos'].forEach(key => {
    if(updates[key] && updates[key].length){
      updates[key].forEach(row => {
        if(canTouchCenter(row.CentreID)) upsertRow_(key, row, now);
      });
    }
  });

  return {ok:true, serverTime: now};
}

/* ============================== Accès feuilles ============================== */

function getSheet_(tableKey){
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES[tableKey]);
}

function getSheetRows_(tableKey){
  const sheet = getSheet_(tableKey);
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];
  const headers = data[0];
  const rows = [];
  for(let i = 1; i < data.length; i++){
    const obj = {};
    headers.forEach((h, idx) => obj[h] = data[i][idx]);
    rows.push(obj);
  }
  return rows;
}

// Met à jour la ligne existante (repérée par sa clé) ou en ajoute une
// nouvelle. Renseigne toujours ModifieLe = maintenant, ce qui permet aux
// autres appareils de détecter le changement au prochain cycle de synchro.
function upsertRow_(tableKey, obj, now){
  const sheet = getSheet_(tableKey);
  const headers = SHEET_HEADERS[tableKey];
  const keyCols = SHEET_KEYS[tableKey];
  const data = sheet.getDataRange().getValues();
  const idxOf = {};
  headers.forEach((h, i) => idxOf[h] = i);

  let foundRow = -1;
  for(let i = 1; i < data.length; i++){
    const match = keyCols.every(k => String(data[i][idxOf[k]]) === String(obj[k]));
    if(match){ foundRow = i + 1; break; }
  }
  const rowValues = headers.map(h => h === 'ModifieLe' ? now : (obj[h] !== undefined ? obj[h] : ''));
  if(foundRow === -1){
    sheet.appendRow(rowValues);
  } else {
    sheet.getRange(foundRow, 1, 1, headers.length).setValues([rowValues]);
  }
}
