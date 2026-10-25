// =====================================================================================
// ⬆️ Mise à jour du code en un clic (menu 🏒 LNH Live › « Mettre à jour le code »)
// Va chercher bundle.json dans le dépôt GitHub (CODE.REPO), remplace les fichiers du projet
// (sauf appsscript.json), crée une nouvelle version et met à jour le déploiement de l'app (même lien).
// Une seule fois : activer l'API Apps Script (script.google.com/home/usersettings) et ajouter les
// autorisations script.projects / script.deployments dans appsscript.json.
// =====================================================================================
const CODE = {
  REPO_PROP: 'CODE_REPO',        // « propriétaire/dépôt » GitHub, demandé au premier usage
  BRANCH: 'main',
  BUNDLE: 'bundle.json',         // { version, files: [{ name, type: 'SERVER_JS' | 'HTML', source }] }
  API: 'https://script.googleapis.com/v1/projects/'
};

function updateCode() {
  const ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  let repo = props.getProperty(CODE.REPO_PROP);
  if (!repo) {
    const r = ui.prompt('⬆️ Mettre à jour le code', 'Dépôt GitHub (propriétaire/dépôt), ex. ulysse/lnh-fantasy :', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK || !/^[\w.-]+\/[\w.-]+$/.test(r.getResponseText().trim())) return;
    repo = r.getResponseText().trim();
    props.setProperty(CODE.REPO_PROP, repo);
  }
  try {
    const res = codeUpdate_(repo, latestSha_(repo));
    ui.alert('⬆️ Code mis à jour', `Version ${res.bundle} installée (${res.files.length} fichiers : ${res.files.join(', ')}).\n` +
      `Version Apps Script n° ${res.version}${res.deployments ? ` · app mise à jour (${res.deployments} déploiement${res.deployments > 1 ? 's' : ''}, même lien)` : ' · ⚠ aucun déploiement d’app trouvé'}.\n\n` +
      'Recharge le classeur (F5) pour le nouveau menu, puis fais une mise à jour des données.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('⬆️ Mise à jour impossible', codeHelp_(e.message), ui.ButtonSet.OK);
  }
}
// Bouton « ⬆️ Mettre à jour le code » de l'app (onglet Modèle) : même chose, sans fenêtre du classeur
function updateCodeFromApp() {
  const repo = PropertiesService.getScriptProperties().getProperty(CODE.REPO_PROP);
  if (!repo) return JSON.stringify({ ok: false, err: 'Dépôt GitHub pas encore choisi : fais la première mise à jour du code depuis le menu du classeur.' });
  try { const r = codeUpdate_(repo, latestSha_(repo)); r.ok = true; return JSON.stringify(r); }
  catch (e) { return JSON.stringify({ ok: false, err: codeHelp_(e.message) }); }
}
// Bouton ↻ de l'app : s'il y a une nouvelle version du code sur GitHub (version.txt ≠ version installée), on l'installe d'abord.
// La page lance ensuite la mise à jour des données, qui tourne déjà avec le nouveau code (le déploiement pointe dessus).
// Le dernier commit est lu directement chez GitHub (git « info/refs » : toujours à jour, sans limite de requêtes),
// puis les fichiers sont pris À CE COMMIT : jamais la vieille copie que GitHub garde 5 minutes en cache pour « main ».
function checkCodeFromApp() {
  const props = PropertiesService.getScriptProperties(), repo = props.getProperty(CODE.REPO_PROP);
  let url = ''; try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  if (!repo) return JSON.stringify({ updated: false, ver: props.getProperty('CODE_VER') || '', url: url });
  try {
    const sha = latestSha_(repo);
    let remote = '';
    if (sha) {
      if (sha === props.getProperty('CODE_SHA')) return JSON.stringify({ updated: false, ver: props.getProperty('CODE_VER') || '', url: url });
      const r = UrlFetchApp.fetch(`https://raw.githubusercontent.com/${repo}/${sha}/version.txt`, { muteHttpExceptions: true });
      remote = r.getResponseCode() === 200 ? r.getContentText().trim() : '';
    } else {                                                          // plan B : l'ancienne méthode (peut avoir 5 min de retard)
      const r = UrlFetchApp.fetch(`https://raw.githubusercontent.com/${repo}/${CODE.BRANCH}/version.txt?t=${Date.now()}`, { muteHttpExceptions: true });
      remote = r.getResponseCode() === 200 ? r.getContentText().trim() : '';
    }
    if (!remote || remote === props.getProperty('CODE_VER')) {
      if (sha && remote) props.setProperty('CODE_SHA', sha);         // même version, autre commit : rien à installer
      return JSON.stringify({ updated: false, ver: props.getProperty('CODE_VER') || '', url: url });
    }
    const u = codeUpdate_(repo, sha);
    return JSON.stringify({ updated: true, bundle: u.bundle, ver: u.bundle, url: url });
  } catch (e) { return JSON.stringify({ updated: false, ver: props.getProperty('CODE_VER') || '', url: url, err: codeHelp_(e.message) }); }
}
function latestSha_(repo) {
  try {
    const r = UrlFetchApp.fetch(`https://github.com/${repo}.git/info/refs?service=git-upload-pack`, { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return null;
    const m = r.getContentText().match(new RegExp('([0-9a-f]{40}) refs/heads/' + CODE.BRANCH + '\\b'));
    return m ? m[1] : null;
  } catch (e) { return null; }
}
// Appelée par la page au chargement : version installée (la page compare avec la sienne)
function codeVersion() {
  let url = ''; try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  return JSON.stringify({ ver: PropertiesService.getScriptProperties().getProperty('CODE_VER') || '', url: url });
}
// ?what=ver : version installée + chronométrage de la dernière mise à jour des données
function codeStatus_() {
  const p = PropertiesService.getScriptProperties();
  let t = null; try { t = JSON.parse(p.getProperty('LAST_TIMING') || 'null'); } catch (e) {}
  return JSON.stringify({ ok: true, ver: p.getProperty('CODE_VER') || '', sha: p.getProperty('CODE_SHA') || '', timing: t });
}
function codeChangeRepo() {
  PropertiesService.getScriptProperties().deleteProperty(CODE.REPO_PROP);
  updateCode();
}

function codeUpdate_(repo, sha) {
  // 1) le paquet de fichiers : raw.githubusercontent (pas de limite de requêtes ; ?t= pour éviter le cache),
  //    sinon l'API GitHub (limitée à 60 requêtes/heure, partagées avec tous les scripts Google : souvent épuisée)
  const t = Date.now(), ref = sha || CODE.BRANCH;
  let g = UrlFetchApp.fetch(`https://raw.githubusercontent.com/${repo}/${ref}/${CODE.BUNDLE}?t=${t}`,
    { muteHttpExceptions: true, headers: { 'Cache-Control': 'no-cache', 'User-Agent': 'lnh-fantasy-updater' } });
  if (g.getResponseCode() !== 200) g = UrlFetchApp.fetch(`https://api.github.com/repos/${repo}/contents/${CODE.BUNDLE}?ref=${ref}`,
    { muteHttpExceptions: true, headers: { Accept: 'application/vnd.github.raw', 'User-Agent': 'lnh-fantasy-updater' } });
  if (g.getResponseCode() !== 200) throw new Error(`GitHub ${g.getResponseCode()} sur ${repo}/${CODE.BUNDLE}`);
  const bundle = JSON.parse(g.getContentText());
  if (!bundle.files || !bundle.files.length || bundle.files.some(f => !f.name || !f.source || (f.type !== 'SERVER_JS' && f.type !== 'HTML')))
    throw new Error('bundle.json invalide');

  // 2) le contenu actuel du projet (on garde appsscript.json et tout fichier qui n'est pas dans le paquet)
  const id = ScriptApp.getScriptId(), auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  const api = (path, method, body) => {
    const r = UrlFetchApp.fetch(CODE.API + id + path, { method: method || 'get', headers: auth, muteHttpExceptions: true,
      contentType: 'application/json', payload: body ? JSON.stringify(body) : undefined });
    if (r.getResponseCode() >= 300) throw new Error(`API Apps Script ${r.getResponseCode()} : ${r.getContentText().slice(0, 300)}`);
    return JSON.parse(r.getContentText() || '{}');
  };
  const cur = api('/content');
  const byName = {};
  (cur.files || []).forEach(f => { byName[f.name] = f; });
  if (!byName.appsscript) throw new Error('appsscript.json introuvable dans le projet');
  bundle.files.forEach(f => { byName[f.name] = { name: f.name, type: f.type, source: f.source }; });
  const files = Object.keys(byName).map(n => ({ name: byName[n].name, type: byName[n].type, source: byName[n].source }));

  // 3) écrire, créer une version, pointer le(s) déploiement(s) d'app web dessus (même lien)
  api('/content', 'put', { files: files });
  const v = api('/versions', 'post', { description: 'Code ' + (bundle.version || '') + ' (mise à jour en un clic)' });
  let n = 0;
  (api('/deployments').deployments || []).forEach(d => {
    const c = d.deploymentConfig || {};
    if (!c.versionNumber) return;                                     // @HEAD (tests) : on n'y touche pas
    if (!(d.entryPoints || []).some(e => e.entryPointType === 'WEB_APP')) return;
    api('/deployments/' + d.deploymentId, 'put', { deploymentConfig: { scriptId: id, versionNumber: v.versionNumber,
      manifestFileName: 'appsscript', description: 'Code ' + (bundle.version || '') } });
    n++;
  });
  PropertiesService.getScriptProperties().setProperties({ CODE_VER: bundle.version || '', CODE_SHA: sha || '' });
  return { bundle: bundle.version || '?', files: bundle.files.map(f => f.name), version: v.versionNumber, deployments: n };
}

function codeHelp_(msg) {
  if (/^GitHub 404/.test(msg)) return msg + '\n\n→ Vérifie que bundle.json est bien dans le dépôt (branche main), que le dépôt est public, et son nom (menu « Changer de dépôt GitHub »).';
  if (/^GitHub/.test(msg)) return msg + '\n\n→ GitHub refuse pour l’instant (limite de requêtes). Réessaie dans quelques minutes.';
  if (/403/.test(msg) && /has not been used|disabled|SERVICE_DISABLED|User has not enabled/i.test(msg))
    return msg + '\n\n→ Active l’API Apps Script : script.google.com/home/usersettings, et dans le projet Google Cloud relié au script. Réessaie dans quelques minutes.';
  if (/insufficient|scope|PERMISSION_DENIED/i.test(msg))
    return msg + '\n\n→ Il manque les autorisations dans appsscript.json (script.projects et script.deployments). Ajoute-les, sauvegarde, puis relance : Google te demandera d’approuver.';
  return msg;
}
