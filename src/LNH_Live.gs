/**
 * 🏒 LNH LIVE — Fantasy Hockey (v7)
 * Récupère les stats réelles de la saison dans l'API publique de la LNH,
 * calcule les points fantasy selon le barème de la ligue et met à jour :
 *   - « LNH Live »            : tous les joueurs de la LNH, classés
 *   - chaque équipe, col. S:X : FanPts réels, MJ, FP/MJ réel, FP/MJ projeté (GAF), écart, tendance
 *   - « GAF - Player Values », col. AE:AH : MJ, FP/MJ réel, écart, tendance pour TOUS les joueurs
 *   - « Waivers », col. V:Y   : les mêmes infos pour les joueurs disponibles
 *     (+ tri possible par « FP/MJ réel » ou « Écart FP/MJ »)
 *   - « Leaderboard Live »    : classement des équipes + écart réel vs projeté
 *   - (optionnel) alignements synchronisés depuis la ligue ESPN → colonne « Joueur » de chaque équipe
 *
 * v7 : leaderboard « vérité » — fiche ESPN, PF/PA, match live, all-play, force réelle par semaine.
 * v6 : mode DIRECT — stats lues dans le centre de match LNH (voir « MODE DIRECT »).
 * v5 : Waivers calculés par le script, bouton « Update » rapide.
 * v4 : alignements jusqu'à 29 joueurs (pickups ESPN), installation robuste avec les « Tableaux ».
 * v3 : la mise en page (en-têtes, formats, couleurs) n'est faite qu'UNE fois ;
 *      les mises à jour suivantes n'écrivent que les données.
 *
 * Installation : Extensions > Apps Script > remplacer tout le code par celui-ci > Enregistrer,
 * puis menu « 🏒 LNH Live » > Mettre à jour maintenant.
 */

// ======================= CONFIGURATION =======================
const CONFIG = {
  SEASON: '20262027',   // mets '20252026' pour tester avec la saison passée
  GAME_TYPE: 2,         // 2 = saison régulière, 3 = séries, 1 = préparatoire
  // Mode DIRECT : stats lues dans le centre de match (buts à la seconde près)
  LIVE_MODE: true,      // false = revenir à l'ancienne API seulement (quelques minutes de délai)
  SEASON_START: '2026-09-01',   // date à partir de laquelle chercher les matchs de la saison
  CACHE_SHEET: 'LNH Cache',     // feuilles techniques (cachées) du mode direct
  GAMES_SHEET: 'LNH Parties',
  GAMES_PER_BATCH: 20,          // matchs terminés lus par paquet
  CATCHUP_SECONDS: 200,         // temps max de rattrapage par mise à jour (le reste passe à la suivante)
  // Barème de la ligue (tiré de la Cheat Sheet GAF)
  SCORING: {
    G: 3, A: 2, PPP: 0.5, SHP: 0.5, SOG: 0.2, HIT: 0.2, BLK: 0.4,
    W: 3, GA: -1, SV: 0.2, SO: 3, OTL: 1,
    HAT: 1              // boni par tour du chapeau (3 buts ou plus dans un match)
  },
  // Tendance : écart = FP/MJ réel − FP/MJ projeté
  SEUIL: 0.3,           // joueur : écart > +0,3 → Surperforme · < −0,3 → Sous-performe · sinon Conforme
  SEUIL_EQUIPE: 0.3,    // équipe : même règle sur l'écart cumulé
  LAST_ROW: 30,         // dernière ligne d'alignement dans chaque équipe (lignes 2 à 30 = 29 joueurs max)
  MIN_GP: 3,            // sous 3 matchs joués → « Petit échantillon » (et exclu du tri Waivers)
  LIVE_SHEET: 'LNH Live',
  BOARD_SHEET: 'LIVE LEADERBOARD',
  ALIAS_SHEET: 'Alias',
  GAF_SHEET: 'GAF - Player Values',
  WAIVER_SHEET: 'Waivers',
  SEARCH_SHEET: 'Player Search',
  NEWS_SHEET: 'News Tracker',
  NEWS_HOUR: 8,         // heure de mise à jour quotidienne des meilleurs marqueurs
  NEWS_TOP: 15,         // nombre de joueurs dans « meilleurs de la soirée »
  NEWS_MAX_LOG: 200,    // transactions conservées
  REFRESH_MINUTES: 30,  // fréquence de la mise à jour automatique (1, 5, 10, 15 ou 30)
  // Synchro des alignements ESPN (cookies : menu 🏒 LNH Live > Connecter ESPN — jamais écrits ici)
  ESPN_LEAGUE_ID: '1442383032',   // mets '' pour désactiver la synchro et gérer les alignements à la main
  ESPN_SEASON: 2027,              // ESPN désigne la saison 2026-27 par « 2027 »
  ESPN_SHEET: 'ESPN',
  ESPN_MIN_ROSTER: 10,            // sécurité : un alignement ESPN de moins de 10 joueurs est ignoré
  // Leaderboard « vérité » : force hebdomadaire estimée de chaque équipe
  GAMES_PER_WEEK: 3,              // matchs LNH par semaine pour un patineur (moyenne de saison ≈ 3)
  GOALIE_START_SHARE: 0.6,        // part des matchs d'un gardien quand il a trop peu de départs pour juger
  LINEUP: { F: 10, D: 6, UTIL: 2, G: 2 }   // alignement actif de la ligue
};
// =============================================================

const UP = '▲ Surperforme', DOWN = '▼ Sous-performe', OK = '≈ Conforme';
const GREEN = '#C6EFCE', RED = '#FFC7CE';
const LAYOUT_VERSION = 'v4';   // (leaderboard « vérité » : clé 'board2')
const WARN = [];   // avertissements affichés en haut du Leaderboard Live
let ESPN_DATA = null;      // réponse complète de la ligue ESPN (fiches, calendrier, pointage)
const ESPN_STATUS = {};    // nom normalisé → statut ESPN (OUT, INJURY_RESERVE, SUSPENSION…)
const ESPN_SHEET_OF = {};  // id d'équipe ESPN → nom de la feuille d'équipe
let TEAM_GP = {};          // équipe LNH → matchs joués (pour la part de départs des gardiens)

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏒 LNH Live')
    .addItem('Update (données seulement — rapide)', 'updateData')
    .addItem('Mise à jour complète (données + mise en page)', 'updateLive')
    .addSeparator()
    .addItem('Activer la mise à jour automatique', 'enableAutoRefresh')
    .addItem('Désactiver la mise à jour automatique', 'disableAutoRefresh')
    .addItem('🔎 Player Search', 'openSearch')
    .addItem('📱 Ouvrir l’app', 'openApp')
    .addItem('🧪 Backtest 2025-26 (modèle maison)', 'runBacktest')
    .addItem('🧪 Backtest par match (adversaire du soir)', 'runMatchupBacktest')
    .addItem('🔮 Régler le moteur de projection (4 saisons)', 'tuneProjection')
    .addItem('🔮 Prédire une saison…', 'predictSeasonPrompt')
    .addItem('🗞 Rafraîchir le News Tracker (meilleurs de la soirée)', 'newsDaily')
    .addSeparator()
    .addItem('Connecter ESPN (cookies espn_s2 et SWID)', 'connectEspn')
    .addItem('Déconnecter ESPN', 'disconnectEspn')
    .addItem('Diagnostic ESPN', 'diagnoseEspn')
    .addSeparator()
    .addItem('Refaire la mise en page (au besoin)', 'resetLayout')
    .addSeparator()
    .addItem('⬆️ Mettre à jour le code (GitHub)', 'updateCode')
    .addItem('⬆️ Changer de dépôt GitHub', 'codeChangeRepo')
    .addToUi();
}

// ---------- Points d'entrée ----------
// updateLive  : mise à jour complète (installe/vérifie la mise en page au besoin)
// updateData  : mise à jour RAPIDE — seulement les données (LNH, ESPN, équipes, Waivers, leaderboard)
let DATA_ONLY = false;
function updateLive() { DATA_ONLY = false; run_(); }
function updateData() { DATA_ONLY = true; run_(); }
// chronométrage fin (voir mark_ dans App.gs)
const M_ = l => { try { if (typeof mark_ === 'function') mark_(l); } catch (e) { /* rien */ } };

function run_() {
  WARN.length = 0;
  const prog = (pct, label) => { try { if (typeof setProgress_ === 'function') setProgress_(pct, label); } catch (e) { /* rien */ } };
  prog(3, 'Lecture des matchs de la LNH…');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const players = getPlayers_();
  prog(40, 'Stats LNH à jour — alignements ESPN…');
  const index = buildIndex_(players);
  TEAM_GP = {};
  players.forEach(p => {
    if (p.pos !== 'G' && p.team) TEAM_GP[p.team] = Math.max(TEAM_GP[p.team] || 0, p.gp || 0);
  });
  const aliases = readAliases_(ss);
  M_('ESPN : synchro des alignements');
  if (!DATA_ONLY) {
    const teamSheets = ss.getSheets().filter(isTeamSheet_);
    teamSheets.forEach(ensureTeamRows_);                   // formules jusqu'à LAST_ROW
    once_('rows' + CONFIG.LAST_ROW, () => extendRefs_(ss, teamSheets));
  }
  syncEspn_(ss, aliases);          // met à jour les alignements depuis ESPN (si connecté)
  prog(50, 'Mise à jour des feuilles…');
  writeLiveSheet_(ss, players);
  M_('Feuille GAF live');
  updateGafLive_(ss, index, aliases);
  if (!DATA_ONLY) { setupWaivers_(ss); setupSearch_(ss); newsSheet_(ss); }
  M_('Feuille Ballotage');
  writeWaivers_(ss);

  prog(60, 'Calcul des équipes…');
  const teams = [];
  const teamNames = [];
  // toutes les lectures d'abord, puis toutes les écritures : une lecture après une écriture force Google à
  // recalculer les formules du classeur, à chaque feuille
  const pre = teamSheets_(ss).map(sh => { M_('Lecture ' + sh.getName()); return [sh, sh.getRange(`B2:I${CONFIG.LAST_ROW}`).getValues()]; });
  pre.forEach(([sh, vals]) => { M_('Écriture ' + sh.getName()); teams.push(updateTeamSheet_(sh, index, aliases, vals)); teamNames.push(sh.getName()); });
  PropertiesService.getDocumentProperties().setProperty('TEAM_SHEETS', JSON.stringify(teamNames));
  M_('Leaderboard');
  writeLeaderboard_(ss, teams);
  prog(72, 'Leaderboard à jour — projection maison et simulation…');
  // App web : projection maison, stats avancées, calendrier, simulation (fichier App.gs)
  try { if (typeof buildAppSnapshot_ === 'function') buildAppSnapshot_(ss, teams, players, index, aliases); }
  catch (e) { console.log('App : ' + e.message); WARN.push('App : ' + e.message); }
  M_('Nouvelles du jour');
  maybeDailyNews_(ss);
  prog(100, 'Terminé');
}

// ---------- Mise en page unique ----------
// Chaque feuille n'est mise en forme qu'une fois (mémorisé dans les propriétés du document).
function once_(key, fn) {
  if (DATA_ONLY) return;
  const props = PropertiesService.getDocumentProperties();
  const k = `layout_${LAYOUT_VERSION}:${key}`;
  if (props.getProperty(k)) return;
  fn();
  props.setProperty(k, '1');
}
function resetLayout() {
  const props = PropertiesService.getDocumentProperties();
  props.getKeys().filter(k => k.indexOf('layout_') === 0).forEach(k => props.deleteProperty(k));
  updateLive();
}
// Une erreur de mise en forme (ex. colonne « typée » d'un Tableau) ne bloque jamais la mise à jour.
function style_(fn) {
  if (DATA_ONLY) return;
  try { fn(); SpreadsheetApp.flush(); } catch (e) { console.log('Mise en forme ignorée : ' + e.message); }
}
function fmt_(range, f) {
  for (let c = 0; c < range.getNumColumns(); c++) {
    style_(() => range.offset(0, c, range.getNumRows(), 1).setNumberFormat(f));
  }
}
function header_(range, values, color) {
  for (let c = 0; c < range.getNumColumns(); c++) {
    const cell = range.offset(0, c, 1, 1);
    try { cell.setValue(values[0][c]); SpreadsheetApp.flush(); }
    catch (e) { WARN.push(`En-tête « ${values[0][c]} » (${range.getSheet().getName()}) non écrit : ${e.message}`); }
    style_(() => cell.setFontWeight('bold').setBackground(color).setFontColor('#FFFFFF').setWrap(true));
  }
}
// Couleurs vert/rouge via mise en forme conditionnelle (posée une fois, suit les données toute seule)
function trendRules_(sh, diffRange, trendRange, seuil) {
  style_(() => {
    const R = () => SpreadsheetApp.newConditionalFormatRule();
    // retire les anciennes règles posées sur ces colonnes (ex. seuils de la v2)
    const cols = [diffRange.getColumn(), trendRange.getColumn()];
    const rules = sh.getConditionalFormatRules()
      .filter(r => !r.getRanges().some(rg => cols.indexOf(rg.getColumn()) >= 0 && rg.getNumColumns() === 1));
    rules.push(R().whenTextStartsWith('▲').setBackground(GREEN).setRanges([trendRange]).build());
    rules.push(R().whenTextStartsWith('▼').setBackground(RED).setRanges([trendRange]).build());
    rules.push(R().whenNumberGreaterThan(seuil).setBackground(GREEN).setRanges([diffRange]).build());
    rules.push(R().whenNumberLessThan(-seuil).setBackground(RED).setRanges([diffRange]).build());
    sh.setConditionalFormatRules(rules);
  });
}
// Efface les anciennes lignes sous les nouvelles données (sans toucher la mise en forme)
function clearBelow_(sh, firstEmptyRow, nCols) {
  const last = sh.getLastRow();
  if (last >= firstEmptyRow) sh.getRange(firstEmptyRow, 1, last - firstEmptyRow + 1, nCols).clearContent();
}

// ---------- API LNH (pages téléchargées en parallèle) ----------
function reportUrl_(path, start) {
  const sort = encodeURIComponent('[{"property":"playerId","direction":"ASC"}]');
  const exp = encodeURIComponent(`seasonId=${CONFIG.SEASON} and gameTypeId=${CONFIG.GAME_TYPE}`);
  return `https://api.nhle.com/stats/rest/en/${path}?isAggregate=true&start=${start}&limit=100&sort=${sort}&cayenneExp=${exp}`;
}
function parse_(res, path) {
  if (res.getResponseCode() !== 200) throw new Error(`API LNH (${path}) : code ${res.getResponseCode()}`);
  return JSON.parse(res.getContentText());
}
function fetchReports_(paths) {
  const firsts = UrlFetchApp.fetchAll(paths.map(p => ({ url: reportUrl_(p, 0), muteHttpExceptions: true })))
    .map((r, i) => parse_(r, paths[i]));
  const out = {}, more = [];
  paths.forEach((p, i) => {
    out[p] = (firsts[i].data || []).slice();
    for (let s = 100; s < (firsts[i].total || 0); s += 100) more.push({ p: p, s: s });
  });
  if (more.length) {
    UrlFetchApp.fetchAll(more.map(m => ({ url: reportUrl_(m.p, m.s), muteHttpExceptions: true })))
      .forEach((r, i) => (parse_(r, more[i].p).data || []).forEach(d => out[more[i].p].push(d)));
  }
  return out;
}

function fetchNhlPlayers_() {
  const S = CONFIG.SCORING;
  const rep = fetchReports_(['skater/summary', 'skater/realtime', 'goalie/summary']);
  const byId = {};

  rep['skater/summary'].forEach(r => {
    byId[r.playerId] = {
      id: r.playerId, name: r.skaterFullName, team: r.teamAbbrevs || '',
      pos: r.positionCode === 'L' ? 'LW' : r.positionCode === 'R' ? 'RW' : r.positionCode,
      gp: r.gamesPlayed || 0, g: r.goals || 0, a: r.assists || 0,
      ppp: r.ppPoints || 0, shp: r.shPoints || 0, sog: r.shots || 0,
      hit: 0, blk: 0, w: '', ga: '', sv: '', so: '', otl: '', hat: 0
    };
  });
  rep['skater/realtime'].forEach(r => {
    const p = byId[r.playerId];
    if (p) { p.hit = r.hits || 0; p.blk = r.blockedShots || 0; }
  });
  rep['goalie/summary'].forEach(r => {
    byId['G' + r.playerId] = {
      id: r.playerId, name: r.goalieFullName, team: r.teamAbbrevs || '', pos: 'G',
      gp: r.gamesPlayed || 0, g: '', a: '', ppp: '', shp: '', sog: '', hit: '', blk: '',
      w: r.wins || 0, ga: r.goalsAgainst || 0, sv: r.saves || 0,
      so: r.shutouts || 0, otl: r.otLosses || 0, hat: ''
    };
  });

  return scoreList_(Object.keys(byId).map(k => byId[k]));
}

// Calcule FanPts et FP/MJ selon le barème, puis trie
function scoreList_(list) {
  const S = CONFIG.SCORING;
  list.forEach(p => {
    const n = v => (typeof v === 'number' ? v : 0);
    p.fp = n(p.g) * S.G + n(p.a) * S.A + n(p.ppp) * S.PPP + n(p.shp) * S.SHP +
           n(p.sog) * S.SOG + n(p.hit) * S.HIT + n(p.blk) * S.BLK +
           n(p.w) * S.W + n(p.ga) * S.GA + n(p.sv) * S.SV + n(p.so) * S.SO + n(p.otl) * S.OTL + n(p.hat) * (S.HAT || 0);
    p.fpg = p.gp ? p.fp / p.gp : 0;
  });
  list.sort((x, y) => y.fp - x.fp);
  return list;
}

// ======================= MODE DIRECT (centre de match) =======================
// Principe anti double-compte :
//   • chaque match TERMINÉ est lu une seule fois et additionné dans « LNH Cache » (liste dans « LNH Parties ») ;
//   • chaque match EN COURS est relu à chaque update et ajouté par-dessus, sans être mis en cache.
//   Un match est donc soit en cache, soit en direct — jamais les deux.
// Vérification officielle : une fois par jour (sans match en cours), si l'ancienne API a exactement
// les mêmes matchs joués pour chaque équipe, ses totaux (corrigés par la LNH) remplacent le cache.
// Filet de sécurité : si le centre de match est inaccessible, on utilise l'ancienne API.
let SOURCE_NOTE = '';
const STAT_KEYS = ['gp', 'g', 'a', 'ppp', 'shp', 'sog', 'hit', 'blk', 'w', 'ga', 'sv', 'so', 'otl', 'hat'];
const CACHE_COLS = 4 + STAT_KEYS.length;
const WEB = 'https://api-web.nhle.com/v1';

function getPlayers_() {
  if (!CONFIG.LIVE_MODE) { SOURCE_NOTE = 'source : API stats LNH'; return fetchNhlPlayers_(); }
  try { return fetchLivePlayers_(); }
  catch (e) {
    WARN.push('Centre de match LNH inaccessible (' + e.message + ') — stats de l’API officielle utilisées');
    SOURCE_NOTE = 'source : API stats LNH (secours)';
    return fetchNhlPlayers_();
  }
}

// Les erreurs passagères du serveur LNH (5xx, 429 — ex. « 520 ») sont réessayées 3 fois avant d'abandonner.
function fetchJsonAll_(urls) {
  if (!urls.length) return [];
  const out = new Array(urls.length);
  let todo = urls.map((u, i) => i), lastCode = {};
  for (let attempt = 0; attempt < 4 && todo.length; attempt++) {
    if (attempt) Utilities.sleep(1500 * attempt);
    const res = UrlFetchApp.fetchAll(todo.map(i => ({ url: urls[i], muteHttpExceptions: true })));
    const retry = [];
    res.forEach((r, k) => {
      const i = todo[k], code = r.getResponseCode();
      if (code === 200) { out[i] = JSON.parse(r.getContentText()); return; }
      lastCode[i] = code;
      if (code >= 500 || code === 429) retry.push(i);
      else throw new Error(`code ${code} sur ${urls[i].split('/v1/')[1]}`);
    });
    todo = retry;
  }
  if (todo.length) throw new Error(`code ${lastCode[todo[0]]} sur ${urls[todo[0]].split('/v1/')[1]} (après 4 essais)`);
  return out;
}

function ymd_(d) { return Utilities.formatDate(d, 'America/Toronto', 'yyyy-MM-dd'); }

// Tous les matchs de la saison entre « from » et aujourd'hui (calendrier par semaine)
function scheduleGames_(from) {
  const today = ymd_(new Date());
  const games = [];
  let d = from, guard = 0;
  while (d && d <= today && guard++ < 60) {
    const j = fetchJsonAll_([`${WEB}/schedule/${d}`])[0];
    (j.gameWeek || []).forEach(day => (day.games || []).forEach(g => {
      if (String(g.season) === CONFIG.SEASON && g.gameType === CONFIG.GAME_TYPE && day.date <= today) {
        games.push({ id: g.id, date: day.date, state: g.gameState, away: g.awayTeam.id, home: g.homeTeam.id });
      }
    }));
    d = j.nextStartDate;
  }
  return games;
}

function blankLine_(id) {
  const o = { id: id, name: '', team: '', pos: '' };
  STAT_KEYS.forEach(k => { o[k] = 0; });
  return o;
}
function addLine_(map, l) {
  const t = map[l.id] || (map[l.id] = blankLine_(l.id));
  STAT_KEYS.forEach(k => { t[k] += l[k] || 0; });
  if (l.name) t.name = l.name;
  if (l.team) t.team = l.team;
  if (l.pos) t.pos = l.pos;
}

// Lit un match (boxscore + déroulement) → lignes de stats par joueur
function parseGame_(box, pbp) {
  const lines = {};
  const names = {};
  (pbp.rosterSpots || []).forEach(r => {
    names[r.playerId] = `${(r.firstName || {}).default || ''} ${(r.lastName || {}).default || ''}`.trim();
  });
  const played = toi => toi && toi !== '00:00';
  const posOf = p => (p === 'L' ? 'LW' : p === 'R' ? 'RW' : p);
  ['awayTeam', 'homeTeam'].forEach(side => {
    const abbrev = box[side].abbrev;
    const st = (box.playerByGameStats || {})[side] || {};
    ['forwards', 'defense'].forEach(grp => (st[grp] || []).forEach(p => {
      if (!played(p.toi)) return;
      const l = blankLine_(p.playerId);
      Object.assign(l, { name: names[p.playerId] || '', team: abbrev, pos: posOf(p.position), gp: 1,
        g: p.goals || 0, a: p.assists || 0, sog: p.sog || 0, hit: p.hits || 0, blk: p.blockedShots || 0,
        hat: (p.goals || 0) >= 3 ? 1 : 0 });
      lines[l.id] = l;
    }));
    const goalies = (st.goalies || []).filter(p => played(p.toi));
    goalies.forEach(p => {
      const l = blankLine_(p.playerId);
      Object.assign(l, { name: names[p.playerId] || '', team: abbrev, pos: 'G', gp: 1,
        ga: p.goalsAgainst || 0, sv: p.saves || 0,
        w: p.decision === 'W' ? 1 : 0, otl: p.decision === 'O' ? 1 : 0,
        so: p.decision === 'W' && !p.goalsAgainst && goalies.length === 1 ? 1 : 0 });
      lines[l.id] = l;
    });
  });

  // Points en avantage / désavantage numérique : situation du but + marqueur et passeurs
  const awayId = box.awayTeam.id;
  (pbp.plays || []).forEach(pl => {
    if (pl.typeDescKey !== 'goal') return;
    if (pl.periodDescriptor && pl.periodDescriptor.periodType === 'SO') return;   // tirs de barrage
    const sc = String(pl.situationCode || '');
    if (sc.length !== 4) return;
    const ag = +sc[0], as = +sc[1], hs = +sc[2], hg = +sc[3];
    if (as < 3 || hs < 3) return;                         // tir de pénalité, etc.
    const d = pl.details || {};
    const owner = d.eventOwnerTeamId != null ? d.eventOwnerTeamId : pl.eventOwnerTeamId;
    const awayEff = as - (ag === 0 ? 1 : 0), homeEff = hs - (hg === 0 ? 1 : 0);   // gardien retiré ≠ avantage
    const mine = owner === awayId ? awayEff : homeEff, opp = owner === awayId ? homeEff : awayEff;
    const key = mine > opp ? 'ppp' : mine < opp ? 'shp' : null;
    if (!key) return;
    [d.scoringPlayerId, d.assist1PlayerId, d.assist2PlayerId].forEach(id => {
      if (id && lines[id]) lines[id][key] += 1;
    });
  });
  return Object.keys(lines).map(k => lines[k]);
}

function readCache_(ss) {
  const sh = ss.getSheetByName(CONFIG.CACHE_SHEET);
  const map = {};
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, CACHE_COLS).getValues().forEach(r => {
      if (!r[0]) return;
      const l = { id: r[0], name: r[1], team: r[2], pos: r[3] };
      STAT_KEYS.forEach((k, i) => { l[k] = +r[4 + i] || 0; });
      map[l.id] = l;
    });
  }
  return map;
}
function writeCache_(ss, map) {
  const sh = techSheet_(ss, CONFIG.CACHE_SHEET, ['playerId', 'Joueur', 'Équipe', 'Pos'].concat(STAT_KEYS));
  const rows = Object.keys(map).map(k => map[k]).map(l => [l.id, l.name, l.team, l.pos].concat(STAT_KEYS.map(k => l[k])));
  sh.getRange(1, 1, 1, CACHE_COLS).setValues([['playerId', 'Joueur', 'Équipe', 'Pos'].concat(STAT_KEYS)]);
  if (rows.length) sh.getRange(2, 1, rows.length, CACHE_COLS).setValues(rows);
  clearBelow_(sh, rows.length + 2, CACHE_COLS);
}
function readGames_(ss) {
  const sh = ss.getSheetByName(CONFIG.GAMES_SHEET);
  const out = [];
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => {
    if (r[0]) out.push({ id: r[0], date: String(r[1]), away: r[2], home: r[3] });
  });
  return out;
}
function appendGames_(ss, games) {
  if (!games.length) return;
  const sh = techSheet_(ss, CONFIG.GAMES_SHEET, ['gameId', 'Date', 'Visiteur (id)', 'Local (id)', 'Traité le']);
  const now = stamp_();
  sh.getRange(sh.getLastRow() + 1, 1, games.length, 5)
    .setValues(games.map(g => [g.id, g.date, g.away, g.home, now]));
}
function techSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    try { sh.hideSheet(); } catch (e) { /* rien */ }
  }
  return sh;
}

// Noms manquants (joueurs jamais vus) : fiche du joueur
function fillNames_(map) {
  const ids = Object.keys(map).filter(k => !map[k].name).slice(0, 80);
  if (!ids.length) return;
  try {
    fetchJsonAll_(ids.map(id => `${WEB}/player/${id}/landing`)).forEach((j, i) => {
      map[ids[i]].name = `${(j.firstName || {}).default || ''} ${(j.lastName || {}).default || ''}`.trim();
    });
  } catch (e) { WARN.push('Noms de certains joueurs introuvables (' + e.message + ')'); }
}

function fetchLivePlayers_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  // Changement de saison ou de type de match → on repart d'un cache vide
  const sig = CONFIG.SEASON + '/' + CONFIG.GAME_TYPE + '/hat';   // « /hat » : relecture unique pour compter les tours du chapeau
  if (props.getProperty('LIVE_SIG') !== sig) {
    [CONFIG.CACHE_SHEET, CONFIG.GAMES_SHEET].forEach(n => {
      const sh = ss.getSheetByName(n);
      if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    });
    props.deleteProperty('LIVE_SCAN_FROM'); props.deleteProperty('LIVE_RECONCILED');
    props.setProperty('LIVE_SIG', sig);
  }
  const cache = readCache_(ss);
  const done = {};
  const doneGames = readGames_(ss);
  doneGames.forEach(g => { done[g.id] = true; });

  const t0 = Date.now();
  // Calendrier à relire : jamais après HIER (un match commencé le 29 peut finir après minuit le 30),
  // et une fois par jour, toute la saison (rattrape tout match qui aurait pu être manqué).
  const today = ymd_(new Date());
  const yesterday = ymd_(new Date(Date.now() - 86400000));
  let from = props.getProperty('LIVE_SCAN_FROM') || CONFIG.SEASON_START;
  if (from > yesterday) from = yesterday;
  if (props.getProperty('LIVE_FULLSCAN') !== today) { from = CONFIG.SEASON_START; props.setProperty('LIVE_FULLSCAN', today); }
  const games = scheduleGames_(from);
  const finals = games.filter(g => (g.state === 'FINAL' || g.state === 'OFF') && !done[g.id])
                      .sort((x, y) => (x.date < y.date ? -1 : 1));
  const lives = games.filter(g => g.state === 'LIVE' || g.state === 'CRIT');
  const readGame_ = list => {
    const urls = [];
    list.forEach(g => { urls.push(`${WEB}/gamecenter/${g.id}/boxscore`, `${WEB}/gamecenter/${g.id}/play-by-play`); });
    const data = fetchJsonAll_(urls);
    return list.map((g, i) => parseGame_(data[2 * i], data[2 * i + 1]));
  };

  // 1) matchs terminés → cache (une seule fois), par paquets, jusqu'à épuisement ou limite de temps
  const added = [];
  for (let k = 0; k < finals.length && (Date.now() - t0) / 1000 < CONFIG.CATCHUP_SECONDS; k += CONFIG.GAMES_PER_BATCH) {
    const batch = finals.slice(k, k + CONFIG.GAMES_PER_BATCH);
    readGame_(batch).forEach((lines, i) => {
      if (!lines.length) { WARN.push(`Match ${batch[i].id} : feuille de match vide, réessai à la prochaine mise à jour`); return; }
      lines.forEach(l => addLine_(cache, l));
      added.push(batch[i]);
    });
  }
  if (added.length) {
    fillNames_(cache);
    writeCache_(ss, cache);
    appendGames_(ss, added);
    added.forEach(g => { done[g.id] = true; doneGames.push(g); });
  }
  const liveParsed = readGame_(lives);
  // prochain balayage du calendrier : premier match pas encore traité (terminé, en cours ou à venir)
  // (un match d'avant-hier ou plus vieux jamais commencé = reporté : ignoré, il réapparaîtra à sa nouvelle date)
  const isFinal = g => g.state === 'FINAL' || g.state === 'OFF';
  const isLive = g => g.state === 'LIVE' || g.state === 'CRIT';
  const pending = games.filter(g => !done[g.id] && (isFinal(g) || isLive(g) || g.date >= yesterday))
                       .map(g => g.date).sort();
  props.setProperty('LIVE_SCAN_FROM', pending[0] || today);

  // 2) vérification officielle quotidienne : aucun match en cours et tous les matchs terminés traités
  const backlog = games.filter(g => isFinal(g) && !done[g.id]).length;
  if (!lives.length && !backlog) reconcile_(ss, props, cache, doneGames);

  // 3) total = cache + matchs en cours (jamais mis en cache)
  const total = {};
  Object.keys(cache).forEach(k => addLine_(total, cache[k]));
  liveParsed.forEach(lines => lines.forEach(l => addLine_(total, l)));
  fillNames_(total);

  SOURCE_NOTE = `source : centre de match (direct) · ${doneGames.length} match(s) terminé(s) comptés` +
    (lives.length ? ` + ${lives.length} en cours` : '') +
    (backlog ? ` · ⚠ ${backlog} match(s) terminé(s) pas encore lus (prochain update)` : '');
  if (backlog) WARN.push(`${backlog} match(s) terminé(s) pas encore lus — ils seront ajoutés au prochain update`);
  return scoreList_(Object.keys(total).map(k => total[k]).filter(l => l.name).map(l => {
    const o = Object.assign({}, l);
    if (o.pos === 'G') ['g', 'a', 'ppp', 'shp', 'sog', 'hit', 'blk', 'hat'].forEach(k => { o[k] = ''; });
    else ['w', 'ga', 'sv', 'so', 'otl'].forEach(k => { o[k] = ''; });
    return o;
  }));
}

function reconcile_(ss, props, cache, doneGames) {
  const today = ymd_(new Date());
  if (props.getProperty('LIVE_RECONCILED') === today) return;
  const hour = +Utilities.formatDate(new Date(), 'America/Toronto', 'H');
  if (hour < 6) return;
  try {
    // matchs joués par équipe : cache vs API officielle
    const mine = {};
    doneGames.forEach(g => { mine[g.away] = (mine[g.away] || 0) + 1; mine[g.home] = (mine[g.home] || 0) + 1; });
    const teams = fetchReports_(['team/summary'])['team/summary'];
    const same = teams.length > 0 && teams.every(t => (mine[t.teamId] || 0) === t.gamesPlayed);
    if (!same) return;   // l'API officielle n'a pas encore tous les matchs : on réessaiera
    const official = fetchNhlPlayers_();
    const fresh = {};
    official.forEach(p => {
      const l = blankLine_(p.id);
      Object.assign(l, { name: p.name, team: String(p.team).split(',').pop(), pos: p.pos });
      STAT_KEYS.forEach(k => { l[k] = typeof p[k] === 'number' ? p[k] : 0; });
      l.hat = cache[l.id] ? cache[l.id].hat || 0 : 0;   // l'API officielle ne donne pas les tours du chapeau
      fresh[l.id] = l;
    });
    Object.keys(cache).forEach(k => delete cache[k]);
    Object.keys(fresh).forEach(k => { cache[k] = fresh[k]; });
    writeCache_(ss, cache);
    props.setProperty('LIVE_RECONCILED', today);
  } catch (e) { console.log('Vérification officielle reportée : ' + e.message); }
}

// ---------- Correspondance des noms ----------
function norm_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[^a-z]/g, '');
}
function buildIndex_(players) {
  const idx = {};
  players.forEach(p => { (idx[norm_(p.name)] = idx[norm_(p.name)] || []).push(p); });
  return idx;
}
function readAliases_(ss) {
  const map = {};
  const sh = ss.getSheetByName(CONFIG.ALIAS_SHEET);
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([from, to]) => {
    if (from && to) map[norm_(from)] = norm_(to);
  });
  return map;
}
// Retrouve un joueur ; en cas d'homonymes (ex. deux Elias Pettersson), départage par position/équipe
function findPlayer_(name, nhlTeam, pos, index, aliases) {
  const key = norm_(name);
  const cands = index[key] || (aliases[key] && index[aliases[key]]) || [];
  if (cands.length <= 1) return cands[0] || null;
  const grp = p => (p === 'G' ? 'G' : p === 'D' ? 'D' : 'F');
  const byPos = cands.filter(c => grp(c.pos) === grp(pos));
  if (byPos.length === 1) return byPos[0];
  const pool = byPos.length ? byPos : cands;
  return pool.find(c => String(c.team).split(',').indexOf(nhlTeam) >= 0) || pool[0];
}

// ---------- Réel vs projeté ----------
// Écart = FP/MJ réel − FP/MJ projeté. Positif = le joueur produit plus que prévu.
function trendOf_(diff, seuil) {
  return diff > seuil ? UP : diff < -seuil ? DOWN : OK;
}
function compare_(p, proj) {
  if (!p || !p.gp) return { gp: 0, fpg: '', diff: '', trend: 'Pas encore joué' };
  const diff = typeof proj === 'number' ? p.fpg - proj : '';
  let trend = '';
  if (diff !== '') trend = p.gp < CONFIG.MIN_GP ? `Petit échantillon (${p.gp} MJ)` : trendOf_(diff, CONFIG.SEUIL);
  return { gp: p.gp, fpg: p.fpg, diff: diff, trend: trend };
}

// ---------- Alignements plus longs (pickups) ----------
// Recopie les formules de la ligne 2 (col. A et C:R) sur toute ligne 2..LAST_ROW qui n'en a pas.
function ensureTeamRows_(sh) {
  const L = CONFIG.LAST_ROW;
  try {
    const rcol = sh.getRange(2, 18, L - 1, 1).getFormulas().map(r => r[0]);   // col. R « Ligne GAF »
    if (rcol.every(f => f)) return;
    const a2 = sh.getRange(2, 1).getFormulaR1C1();
    const cr2 = sh.getRange(2, 3, 1, 16).getFormulasR1C1()[0];
    let r = 0;
    while (r < rcol.length) {
      if (rcol[r]) { r++; continue; }
      let e = r;
      while (e < rcol.length && !rcol[e]) e++;
      const n = e - r, row = r + 2;
      sh.getRange(row, 1, n, 1).setFormulasR1C1(Array(n).fill([a2]));
      sh.getRange(row, 3, n, 16).setFormulasR1C1(Array(n).fill(cr2));
      r = e;
    }
  } catch (e) {
    WARN.push(`« ${sh.getName()} » : impossible de prolonger les formules jusqu'à la ligne ${L} (${e.message})`);
  }
}

// Remplace, dans les autres feuilles, les références aux alignements (ex. 'Équipe'!G2:G22) par …:G30
function extendRefs_(ss, teamSheets) {
  const names = {};
  teamSheets.forEach(t => { names[t.getName()] = true; });
  const L = CONFIG.LAST_ROW;
  const re = /(?:'((?:[^']|'')+)'|([A-Za-z0-9_À-ÿ]+))!(\$?[A-Z]{1,3}\$?)2:(\$?[A-Z]{1,3}\$?)(\d+)/g;
  ss.getSheets().forEach(sh => {
    if (names[sh.getName()]) return;
    const last = sh.getLastRow(), lastC = sh.getLastColumn();
    if (!last || !lastC) return;
    const fs = sh.getRange(1, 1, last, lastC).getFormulas();
    for (let c = 0; c < lastC; c++) {
      let run = null;
      const flush = () => { if (run) { sh.getRange(run.r + 1, c + 1, run.f.length, 1).setFormulas(run.f.map(x => [x])); run = null; } };
      for (let r = 0; r < last; r++) {
        const f = fs[r][c];
        const nf = f ? f.replace(re, (m, q, u, c1, c2, end) => {
          const nm = q ? q.replace(/''/g, "'") : u;
          const e = +end;
          return names[nm] && e >= 20 && e < L ? m.slice(0, m.length - end.length) + L : m;
        }) : f;
        if (f && nf !== f) {
          if (run && run.r + run.f.length === r) run.f.push(nf); else { flush(); run = { r: r, f: [nf] }; }
        } else flush();
      }
      flush();
    }
  });
}

// ---------- Feuilles d'équipes ----------
// Mémorisé pendant toute l'exécution : l'en-tête d'une feuille ne change pas en cours de mise à jour
// (avant : A1:B1 relu ~70 fois par mise à jour, ~7 s)
const TEAM_SHEET_MEMO = {};
// Les feuilles d'équipe : d'abord la liste gardée à la dernière mise à jour (TEAM_SHEETS), vérifiée ;
// tour complet du classeur seulement si elle ne colle plus (feuille renommée, ajoutée ou supprimée) ou si on le demande.
let TEAM_SHEETS_LIST = null;
function teamSheets_(ss, full) {
  if (TEAM_SHEETS_LIST && !full) return TEAM_SHEETS_LIST;
  const all = ss.getSheets();
  let names = [];
  try { names = JSON.parse(PropertiesService.getDocumentProperties().getProperty('TEAM_SHEETS') || '[]'); } catch (e) { names = []; }
  const set = {}; names.forEach(n => { set[n] = true; });
  let list = !full && names.length ? all.filter(sh => set[sh.getName()] && isTeamSheet_(sh)) : null;
  if (!list || list.length !== names.length) list = all.filter(isTeamSheet_);
  return (TEAM_SHEETS_LIST = list);
}
function isTeamSheet_(sh) {
  const id = sh.getSheetId();
  if (TEAM_SHEET_MEMO[id] !== undefined) return TEAM_SHEET_MEMO[id];
  const h = sh.getRange('A1:B1').getValues()[0];
  return (TEAM_SHEET_MEMO[id] = h[0] === 'Choix Repêchage' && h[1] === 'Joueur');
}

function updateTeamSheet_(sh, index, aliases, vals) {   // vals : B2:I déjà lus (facultatif)
  const L = CONFIG.LAST_ROW;
  once_('team:' + sh.getSheetId(), () => {
    header_(sh.getRange('S1:X1'), [['FanPts réels (LNH)', 'MJ réels', 'FP/MJ réel', 'FP/MJ projeté (GAF)',
                                    'Écart FP/MJ (réel − proj.)', 'Tendance']], '#0B6E4F');
    fmt_(sh.getRange(`S2:S${L}`), '0.0');
    fmt_(sh.getRange(`U2:V${L}`), '0.00');
    fmt_(sh.getRange(`W2:W${L}`), '+0.00;-0.00;0.00');
    style_(() => sh.getRange(`W2:X${L}`).setBackground(null));   // enlève les couleurs fixes de la v2
    trendRules_(sh, sh.getRange(`W2:W${L}`), sh.getRange(`X2:X${L}`), CONFIG.SEUIL);
  });

  const data = vals ? vals.map(r => [r[0], r[1], r[2]]) : sh.getRange(`B2:D${L}`).getValues();   // Joueur, Équipe LNH, Pos
  saveRoster_(sh, data.map(r => r[0]));
  const proj = vals ? vals.map(r => r[7]) : sh.getRange(`I2:I${L}`).getValues().map(r => r[0]); // FP/Match projeté (GAF), colonne I
  const out = [];
  const t = { team: sh.getName(), total: 0, gp: 0, sumReal: 0, sumProj: 0, counted: 0,
              best: null, worst: null, missing: [], players: [] };

  data.forEach(([name, team, pos], i) => {
    const pj = typeof proj[i] === 'number' ? proj[i] : '';
    if (!name) { out.push(['', '', '', '', '', '']); return; }
    const p = findPlayer_(name, team, pos, index, aliases);
    const c = compare_(p, pj);
    t.players.push({ name: name, pos: String(pos || (p && p.pos) || ''), nhl: p ? p.team : String(team || ''),
                     gp: p ? p.gp : 0, fpg: p && p.gp ? p.fpg : 0, proj: pj === '' ? 0 : pj,
                     status: ESPN_STATUS[norm_(name)] || '' });
    if (!p) t.missing.push(name);
    else {
      t.total += p.fp; t.gp += p.gp;
      if (p.gp && pj !== '') {
        t.sumReal += p.fpg; t.sumProj += pj; t.counted++;
        if (p.gp >= CONFIG.MIN_GP) {
          if (!t.best || c.diff > t.best.diff) t.best = { name: name, diff: c.diff };
          if (!t.worst || c.diff < t.worst.diff) t.worst = { name: name, diff: c.diff };
        }
      }
    }
    out.push([p ? p.fp : 0, c.gp, c.fpg, pj, c.diff, c.trend]);
  });

  sh.getRange(`S2:X${L}`).setValues(out);
  sh.getRange('S1').setNote(t.missing.length
    ? 'Sans stats LNH (pas encore joué, ou nom différent — ajoute un alias vers le nom LNH) : ' + t.missing.join(', ')
    : '');
  return t;
}

// ---------- GAF : réel vs projeté pour tous les joueurs ----------
function updateGafLive_(ss, index, aliases) {
  const sh = ss.getSheetByName(CONFIG.GAF_SHEET);
  if (!sh) return;
  once_('gaf', () => {
    const ag1 = sh.getRange('AG1').getValue();   // déplace l'ancienne note de source si elle occupe AG1
    if (typeof ag1 === 'string' && ag1.indexOf('Source') === 0) {
      sh.getRange('AJ1').setValue(ag1);
      sh.getRange('AG1').clearContent();
    }
    header_(sh.getRange('AE1:AH1'), [['MJ réels', 'FP/MJ réel', 'Écart FP/MJ', 'Tendance']], '#0B6E4F');
  });
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const rows = sh.getRange(2, 4, n, 5).getValues();   // D:H = Joueur, Équipe, Pos, FanPts, FP/GP
  const out = rows.map(([name, team, pos, , proj]) => {
    if (!name) return ['', '', '', ''];
    const c = compare_(findPlayer_(name, team, pos, index, aliases), proj);
    return [c.gp, c.fpg, c.diff, c.trend];
  });
  sh.getRange(2, 31, n, 4).setValues(out);
}

// ---------- Waivers ----------
// La liste est calculée par le script (plus de formules croisées) :
// à chaque mise à jour, et dès que tu changes « Classer par » / « Position » ou un nom dans une équipe.
const WAIVER_FIRST = 7, WAIVER_ROWS = 400;
const WAIVER_OPTS = ['FanPts', 'VORP', 'FP/GP', 'Écart FP/MJ', 'FP/MJ réel', 'G', 'A', 'PPP', 'SOG', 'HIT', 'BLK', 'W', 'SV', 'SO'];

function waiverSheets_(ss) {
  const b = ss.getSheetByName(CONFIG.WAIVER_SHEET);
  const g = ss.getSheetByName(CONFIG.GAF_SHEET);
  if (!b || !g) { WARN.push(`Feuille « ${CONFIG.WAIVER_SHEET} » ou « ${CONFIG.GAF_SHEET} » introuvable.`); return null; }
  const top = b.getRange('A2:U6').getValues();            // une seule lecture : C2, C3, C4, E4, U6
  if (top[4][20] !== 'Ligne GAF') {
    WARN.push(`« ${CONFIG.WAIVER_SHEET} » : structure inattendue (U6 ≠ « Ligne GAF ») — liste non mise à jour.`);
    return null;
  }
  return { b: b, g: g, crit: top[0][2], pos: top[1][2], prev: top[2][2], note: top[2][4] };
}

// Mise en page (une fois) : colonnes V:Y, couleurs, liste déroulante ; retire les anciennes formules lourdes
function setupWaivers_(ss) {
  const w = waiverSheets_(ss);
  if (!w) return;
  const b = w.b, g = w.g;
  once_('waivers-script', () => {
    header_(b.getRange('V6:Y6'), [['MJ réels', 'FP/MJ réel', 'Écart FP/MJ (réel − proj.)', 'Tendance']], '#0B6E4F');
    fmt_(b.getRange(WAIVER_FIRST, 23, WAIVER_ROWS, 1), '0.00');
    fmt_(b.getRange(WAIVER_FIRST, 24, WAIVER_ROWS, 1), '+0.00;-0.00;0.00');
    trendRules_(b, b.getRange(WAIVER_FIRST, 24, WAIVER_ROWS, 1), b.getRange(WAIVER_FIRST, 25, WAIVER_ROWS, 1), CONFIG.SEUIL);
    style_(() => b.getRange('C2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(WAIVER_OPTS, true).build()));
    b.getRange('E2').setValue(`← critère de classement. « Écart FP/MJ » = joueurs disponibles qui produisent le plus au-dessus de leur projection (min. ${CONFIG.MIN_GP} MJ)`);
    b.getRange('E4').setValue(WAIVER_NOTE);
    const n = g.getLastRow() - 1;
    if (n > 0) g.getRange(2, 30, n, 1).clearContent();   // ancienne « Clé Ballotage » (col. AD), devenue inutile
  });
}

const WAIVER_NOTE = "Liste calculée par le script : tout joueur GAF absent des alignements. Elle se recalcule quand tu changes le tri, la position ou un alignement.";

// Optimisé : 1 lecture GAF + 1 écriture (au lieu d'une lecture par feuille d'équipe)
function writeWaivers_(ss) {
  const w = waiverSheets_(ss);
  if (!w) return;
  const b = w.b, g = w.g;

  const all = g.getRange(1, 1, g.getLastRow(), 34).getValues();   // A:AH, en-tête compris
  const hdr = all[0].slice(6, 27);                                  // G:AA
  const data = all.slice(1);
  // Joueurs pris : lus directement dans la colonne R (« Ligne GAF ») de chaque alignement.
  // Liste des équipes mémorisée à chaque update → 1 lecture par équipe seulement.
  const taken = {};
  const readTeam = sh => sh.getRange(2, 18, CONFIG.LAST_ROW - 1, 1).getValues()
    .forEach(r => { if (typeof r[0] === 'number') taken[r[0]] = true; });
  let names = [];
  try { names = JSON.parse(PropertiesService.getDocumentProperties().getProperty('TEAM_SHEETS') || '[]'); } catch (e) { names = []; }
  const sheets = names.map(n => ss.getSheetByName(n));
  if (names.length && sheets.every(Boolean)) sheets.forEach(readTeam);
  else ss.getSheets().filter(isTeamSheet_).forEach(readTeam);   // liste absente ou périmée : on parcourt tout
  const isTaken = i => taken[i + 1];

  const crit = String(w.crit || 'FanPts');
  const posSel = String(w.pos || 'Tous');
  let col;
  if (crit === 'FP/MJ réel') col = 31;                     // AF
  else if (crit === 'Écart FP/MJ') col = 32;               // AG
  else {
    const i = hdr.indexOf(crit);
    if (i < 0) WARN.push(`Waivers : critère « ${crit} » inconnu, tri par FanPts.`);
    col = i < 0 ? 6 : 6 + i;
  }
  const live = col === 31 || col === 32;
  const posOk = p => posSel === 'Tous' || posSel === p ||
    (posSel === 'F' && (p === 'C' || p === 'LW' || p === 'RW')) || (posSel === 'Patineurs' && p !== 'G');

  // Rang GAF selon les FanPts (équivalent de RANK)
  const rank = {};
  data.map(r => r[6]).filter(v => typeof v === 'number').sort((x, y) => y - x)
      .forEach((v, k) => { if (!(v in rank)) rank[v] = k + 1; });

  const cands = [];
  data.forEach((r, i) => {
    if (!r[3] || isTaken(i) || !posOk(r[5])) return;
    const v = r[col];
    if (typeof v !== 'number') return;
    if (live && !(typeof r[30] === 'number' && r[30] >= CONFIG.MIN_GP)) return;
    cands.push({ i: i, v: v });
  });
  cands.sort((x, y) => y.v - x.v || x.i - y.i);

  // n'écrit que les lignes utiles (+ efface celles de l'ancienne liste si elle était plus longue)
  const prev = typeof w.prev === 'number' ? Math.min(w.prev, WAIVER_ROWS) : WAIVER_ROWS;
  const nRows = Math.max(1, Math.min(WAIVER_ROWS, Math.max(prev, cands.length)));
  const z = v => (v === null || v === undefined ? '' : v);
  const out = [];
  for (let k = 0; k < nRows; k++) {
    const c = cands[k];
    if (!c) { out.push(Array(25).fill('')); continue; }
    const r = data[c.i];
    out.push([k + 1, r[3], r[4], r[5], r[6], r[8], r[7], z(rank[r[6]]), z(r[0]), z(r[10]), z(r[11]),
              z(r[14]), z(r[15]), z(r[16]), z(r[18]), z(r[19]), z(r[20]), z(r[22]), z(r[24]), z(r[25]),
              c.i + 1, z(r[30]), z(r[31]), z(r[32]), z(r[33])]);
  }
  b.getRange(WAIVER_FIRST, 1, nRows, 25).setValues(out);
  b.getRange('C4').setValue(cands.length);
  const note = !cands.length && live
    ? `Aucun joueur disponible n'a encore ${CONFIG.MIN_GP} matchs joués : ce tri se remplira quand la saison sera commencée.`
    : WAIVER_NOTE;
  if (note !== w.note) b.getRange('E4').setValue(note);
}

// ======================= PLAYER SEARCH =======================
// Choisis un joueur en B3 → le script trouve son équipe et t'y amène (ligne du joueur sélectionnée).
function setupSearch_(ss) {
  let sh = ss.getSheetByName(CONFIG.SEARCH_SHEET);
  if (sh && sh.getRange('A3').getValue() === 'Joueur :') return sh;
  if (!sh) sh = ss.insertSheet(CONFIG.SEARCH_SHEET, 1);
  sh.getRange('A1').setValue('🔎 PLAYER SEARCH');
  sh.getRange('A3').setValue('Joueur :');
  sh.getRange('C3').setValue('← choisis ou tape un nom (avec ou sans accents) : tu seras amené à son équipe');
  sh.getRange('A5:A14').setValues([['Statut'], ['Équipe fantasy'], ['Équipe LNH'], ['Pos'], ['FanPts projetés (GAF)'],
    ['FP/MJ projeté'], ['MJ réels'], ['FP/MJ réel'], ['Écart FP/MJ'], ['Tendance']]);
  const g = ss.getSheetByName(CONFIG.GAF_SHEET);
  const saveDraft = DATA_ONLY; DATA_ONLY = false;          // la mise en page ici doit toujours se faire
  if (g) style_(() => sh.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(g.getRange('D2:D2000'), true).setAllowInvalid(true).build()));
  style_(() => sh.getRange('A1').setFontSize(14).setFontWeight('bold'));
  style_(() => sh.getRange('A3:A14').setFontWeight('bold'));
  style_(() => sh.getRange('B3').setBackground('#FFF2CC').setFontWeight('bold').setFontColor('#0000FF'));
  style_(() => sh.getRange('C3').setFontColor('#808080').setFontStyle('italic'));
  style_(() => sh.setColumnWidth(1, 170));
  style_(() => sh.setColumnWidth(2, 240));
  DATA_ONLY = saveDraft;
  return sh;
}

function openSearch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = setupSearch_(ss);
  sh.activate();
  sh.getRange('B3').activate();
}

// Index des alignements (noms normalisés par équipe), mis à jour à chaque update et à chaque modif manuelle
function saveRoster_(sh, names) {
  try {
    PropertiesService.getDocumentProperties()
      .setProperty('ROSTER:' + sh.getName(), JSON.stringify(names.map(n => (n ? norm_(n) : ''))));
  } catch (e) { /* index facultatif */ }
}
function findInRosters_(ss, matchName, matchKey) {
  const props = PropertiesService.getDocumentProperties().getProperties();
  let teams = [];
  try { teams = JSON.parse(props.TEAM_SHEETS || '[]'); } catch (e) { teams = []; }
  const indexed = teams.length && teams.every(t => props['ROSTER:' + t]);
  if (indexed) {
    for (const t of teams) {
      const list = JSON.parse(props['ROSTER:' + t]);
      const i = list.findIndex(k => k && matchKey(k));
      if (i < 0) continue;
      const sheet = ss.getSheetByName(t);
      if (sheet && matchName(sheet.getRange(i + 2, 2).getValue())) return { sheet: sheet, row: i + 2 };
      break;   // index périmé : on vérifie tout
    }
    if (teams.every(t => JSON.parse(props['ROSTER:' + t]).every(k => !k || !matchKey(k)))) return null;
  }
  // secours : parcours des feuilles (index absent ou périmé)
  let found = null;
  ss.getSheets().some(t => {
    const v = t.getRange(1, 1, CONFIG.LAST_ROW, 2).getValues();
    if (v[0][0] !== 'Choix Repêchage' || v[0][1] !== 'Joueur') return false;
    saveRoster_(t, v.slice(1).map(r => r[1]));
    const i = v.findIndex((r, k) => k > 0 && r[1] && matchName(r[1]));
    if (i > 0) { found = { sheet: t, row: i + 1 }; return true; }
    return false;
  });
  return found;
}

function searchPlayer_(ss, sh, name) {
  const out = sh.getRange('B5:B14');
  if (!name) { out.clearContent(); return; }
  const key = norm_(name);
  const aliases = readAliases_(ss);
  const k2 = aliases[key];
  const match = n => { const k = norm_(n); return k === key || (k2 && k === k2) || aliases[k] === key; };

  // Infos GAF (projection + réel)
  let info = null;
  const g = ss.getSheetByName(CONFIG.GAF_SHEET);
  if (g) {
    const names = g.getRange(2, 4, g.getLastRow() - 1, 1).getValues();
    const idx = names.findIndex(r => r[0] && match(r[0]));
    if (idx >= 0) {
      const r = g.getRange(idx + 2, 1, 1, 34).getValues()[0];
      info = { team: r[4], pos: r[5], fp: r[6], fpg: r[7], gp: r[30], rfpg: r[31], diff: r[32], trend: r[33] };
    }
  }
  const infoCells = info ? [info.team, info.pos, info.fp, info.fpg, info.gp, info.rfpg, info.diff, info.trend]
                         : ['', '', '', '', '', '', '', ''];

  // Équipe fantasy : index des alignements mémorisé (aucune feuille à parcourir)
  let found = findInRosters_(ss, n => match(n), k => k === key || (k2 && k === k2) || aliases[k] === key);
  if (found) {
    out.setValues([['✔ Dans un alignement'], [found.sheet.getName()]].concat(infoCells.map(x => [x])));
    found.sheet.activate();
    found.sheet.getRange(found.row, 2).activate();
    return;
  }
  // Pas dans un alignement → Waivers
  const w = ss.getSheetByName(CONFIG.WAIVER_SHEET);
  let wRow = -1;
  if (w) {
    const list = w.getRange(WAIVER_FIRST, 2, WAIVER_ROWS, 1).getValues();
    wRow = list.findIndex(r => r[0] && match(r[0]));
  }
  const status = info ? (wRow >= 0 ? '🟢 Disponible (Waivers)' : '🟢 Disponible — hors de la liste Waivers affichée (change le tri ou la position)')
                      : '❓ Joueur introuvable (vérifie l’orthographe ou ajoute un alias)';
  out.setValues([[status], ['—']].concat(infoCells.map(x => [x])));
  if (w && wRow >= 0) { w.activate(); w.getRange(WAIVER_FIRST + wRow, 2).activate(); }
}

// Déclencheur simple : recalcule les Waivers quand on change le tri/la position ou un nom d'alignement
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    const col = e.range.getColumn(), row = e.range.getRow();
    const ss = e.source;
    if (sh.getName() === CONFIG.SEARCH_SHEET && row === 3 && col === 2) { searchPlayer_(ss, sh, e.range.getValue()); return; }
    if (sh.getName() === CONFIG.WAIVER_SHEET && row >= 2 && row <= 3 && col <= 3) { writeWaivers_(ss); return; }
    if (col <= 2 && e.range.getLastColumn() >= 2 && row >= 2 && row <= CONFIG.LAST_ROW && isTeamSheet_(sh)) {
      saveRoster_(sh, sh.getRange(2, 2, CONFIG.LAST_ROW - 1, 1).getValues().map(r => r[0]));
      SpreadsheetApp.flush();   // laisse la colonne « Ligne GAF » se recalculer
      writeWaivers_(ss);
    }
  } catch (err) { console.log('onEdit : ' + err.message); }
}

// ---------- LNH Live ----------
function writeLiveSheet_(ss, players) {
  const sh = ss.getSheetByName(CONFIG.LIVE_SHEET) || ss.insertSheet(CONFIG.LIVE_SHEET);
  const header = ['Rang', 'Joueur', 'Équipe', 'Pos', 'MJ', 'G', 'A', 'PPP', 'SHP', 'SOG', 'HIT', 'BLK',
                  'W', 'GA', 'SV', 'SO', 'OTL', 'FanPts réels', 'FP/Match', 'HAT'];
  once_('live', () => {
    header_(sh.getRange(1, 1, 1, header.length), [header], '#1F3864');
    fmt_(sh.getRange(2, 18, 1500, 2), '0.0');
    style_(() => sh.setFrozenRows(1));
  });
  const rows = players.map((p, i) => [i + 1, p.name, p.team, p.pos, p.gp, p.g, p.a, p.ppp, p.shp, p.sog,
                                      p.hit, p.blk, p.w, p.ga, p.sv, p.so, p.otl, p.fp, p.fpg, p.hat]);
  sh.getRange(1, header.length).setValue('HAT');          // colonne ajoutée (tours du chapeau)
  sh.getRange(1, header.length + 1).setValue('');         // ancienne position de « Mis à jour »
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  clearBelow_(sh, rows.length + 2, header.length);
  sh.getRange(1, header.length + 2).setValue('Mis à jour : ' + stamp_() + ' — saison ' + CONFIG.SEASON + ' — ' + SOURCE_NOTE);
}

// Séparateur d'arguments des formules écrites par script : testé une fois dans une cellule, puis mémorisé
function formulaSep_(ss, sh) {
  const props = PropertiesService.getDocumentProperties();
  const known = props.getProperty('FORMULA_SEP');
  if (known) return known;
  const cell = sh.getRange('Z1');
  let found = ',';
  for (const sep of [',', ';']) {
    try {
      cell.setFormula(`=HYPERLINK("#gid=0"${sep}"ok")`);
      SpreadsheetApp.flush();
      if (cell.getDisplayValue() === 'ok') { found = sep; break; }
    } catch (e) { /* essai suivant */ }
  }
  cell.clearContent();
  props.setProperty('FORMULA_SEP', found);
  return found;
}

// ---------- Leaderboard Live (« vérité ») ----------
// Classement officiel ESPN (fiche, PF, PA) + match de la semaine + all-play (chance) +
// force hebdomadaire estimée (meilleur alignement actif × matchs par semaine, blessés exclus).
const OUT_STATUSES = { OUT: 1, INJURY_RESERVE: 1, SUSPENSION: 1, IR: 1 };

// Points attendus par semaine pour un alignement optimal. mode 'real' = FP/MJ réels, 'proj' = GAF.
function teamPower_(t, mode) {
  const L = CONFIG.LINEUP, gpw = CONFIG.GAMES_PER_WEEK;
  const ok = t.players.filter(p => !OUT_STATUSES[p.status]);
  const val = p => mode === 'real' ? (p.gp ? p.fpg : 0) : (p.proj || 0);
  const isG = p => /\bG\b/.test(p.pos), isD = p => !isG(p) && /\bD\b/.test(p.pos);
  const by = (a, b) => val(b) - val(a);
  const F = ok.filter(p => !isG(p) && !isD(p)).sort(by), D = ok.filter(isD).sort(by);
  const util = F.slice(L.F).concat(D.slice(L.D)).sort(by).slice(0, L.UTIL);
  const sk = F.slice(0, L.F).concat(D.slice(0, L.D), util).reduce((s, p) => s + val(p), 0) * gpw;
  const g = ok.filter(isG).map(p => {
    const tg = TEAM_GP[p.nhl] || 0;
    const share = tg >= 3 && p.gp ? Math.min(1, p.gp / tg) : CONFIG.GOALIE_START_SHARE;
    return val(p) * share * gpw;
  }).sort((a, b) => b - a).slice(0, L.G).reduce((s, x) => s + x, 0);
  return sk + g;
}

// Fiches, PF/PA, match en cours et all-play à partir de la réponse ESPN
function espnStandings_() {
  const j = ESPN_DATA;
  if (!j || !j.teams) return null;
  const cur = (j.status && j.status.currentMatchupPeriod) || null;
  const byId = {};
  j.teams.forEach(t => {
    const o = (t.record && t.record.overall) || {};
    byId[t.id] = { id: t.id, espnName: (t.name || `${t.location || ''} ${t.nickname || ''}`).trim(),
                   w: o.wins || 0, l: o.losses || 0, tie: o.ties || 0,
                   pf: o.pointsFor || 0, pa: o.pointsAgainst || 0, seed: t.playoffSeed || '',
                   apW: 0, apL: 0, apT: 0, live: '', opp: '', oppLive: '', wkW: 0, wkL: 0, wkT: 0 };
  });
  const weeks = {};
  (j.schedule || []).forEach(m => {
    const side = s => s ? { id: s.teamId, pts: typeof s.totalPointsLive === 'number' ? s.totalPointsLive : (s.totalPoints || 0) } : null;
    const h = side(m.home), a = side(m.away);
    if (!h || !byId[h.id]) return;
    if (m.winner && m.winner !== 'UNDECIDED') {
      (weeks[m.matchupPeriodId] = weeks[m.matchupPeriodId] || []).push(h);
      if (a) weeks[m.matchupPeriodId].push(a);
    } else if (cur && m.matchupPeriodId === cur && a && byId[a.id]) {
      byId[h.id].live = h.pts; byId[h.id].opp = a.id; byId[h.id].oppLive = a.pts;
      byId[a.id].live = a.pts; byId[a.id].opp = h.id; byId[a.id].oppLive = h.pts;
    }
  });
  // all-play de la semaine en cours : pointage live de chaque équipe contre celui des 11 autres
  const liveIds = Object.keys(byId).filter(id => byId[id].live !== '');
  liveIds.forEach(x => liveIds.forEach(y => {
    if (x === y) return;
    const a = byId[x].live, b = byId[y].live;
    if (a > b) byId[x].wkW++; else if (a < b) byId[x].wkL++; else byId[x].wkT++;
  }));
  Object.keys(weeks).forEach(k => {
    const list = weeks[k];
    list.forEach(x => list.forEach(y => {
      if (x === y || !byId[x.id]) return;
      if (x.pts > y.pts) byId[x.id].apW++; else if (x.pts < y.pts) byId[x.id].apL++; else byId[x.id].apT++;
    }));
  });
  const set = (j.settings && j.settings.scheduleSettings) || {};
  return { byId: byId, playoffTeams: set.playoffTeamCount || '', weeksDone: Object.keys(weeks).length, current: cur };
}

function writeLeaderboard_(ss, teams) {
  const sh = ss.getSheetByName(CONFIG.BOARD_SHEET) || ss.insertSheet(CONFIG.BOARD_SHEET, 1);
  const header = ['Rang', 'Équipe', 'Fiche ESPN (V-D-N)', '% victoires', 'Points pour (PF)', 'Points contre (PA)',
                  'Cette semaine (live)', 'Adversaire (live)', 'All-play (vs toutes les équipes)', 'All-play cette semaine (live)', 'Chance cette semaine (live)', 'Chance (victoires de + / − que mérité)',
                  'Force réelle / sem.', 'Rang force réelle', 'Force projetée / sem.', 'Écart réel − proj. / sem.',
                  'Absents (blessés, suspendus)', 'Pire sous-performant', 'FanPts saison de l\'effectif actuel'];
  const N = header.length;
  once_('board4', () => {
    style_(() => sh.getRange('A1').setFontSize(14).setFontWeight('bold'));
    style_(() => sh.getRange('A2').setFontColor('#808080'));
    header_(sh.getRange(4, 1, 1, N), [header], '#1F3864');
    fmt_(sh.getRange(5, 4, 20, 1), '0.0%');
    fmt_(sh.getRange(5, 5, 20, 4), '0.0');
    fmt_(sh.getRange(5, 11, 20, 2), '+0.0;-0.0;0.0');
    fmt_(sh.getRange(5, 13, 20, 1), '0');
    fmt_(sh.getRange(5, 15, 20, 1), '0');
    fmt_(sh.getRange(5, 16, 20, 1), '+0.0;-0.0;0.0');
    fmt_(sh.getRange(5, 19, 20, 1), '0.0');
    style_(() => {
      const R = () => SpreadsheetApp.newConditionalFormatRule();
      const rules = sh.getConditionalFormatRules().filter(r => !r.getRanges().some(rg => rg.getRow() >= 4));
      [11, 12, 16].forEach(c => {
        const rg = sh.getRange(5, c, 20, 1);
        rules.push(R().whenNumberGreaterThan(0).setBackground(GREEN).setRanges([rg]).build());
        rules.push(R().whenNumberLessThan(0).setBackground(RED).setRanges([rg]).build());
      });
      sh.setConditionalFormatRules(rules);
    });
    style_(() => sh.setFrozenRows(4));
    style_(() => sh.autoResizeColumns(1, N));
  });
  sh.getRange(4, 1, 1, N).setValues([header]);   // en-têtes toujours à jour, même en « Update » rapide

  const st = espnStandings_();
  sh.getRange('A1').setValue('🏒 LEADERBOARD — classement officiel ESPN + force réelle des alignements');
  sh.getRange('A2').setValue('Mis à jour : ' + stamp_() + ' — saison ' + CONFIG.SEASON +
    (st ? ` · ${st.weeksDone} semaine(s) terminée(s)` + (st.playoffTeams ? ` · ${st.playoffTeams} équipes en séries` : '') : ' · ESPN non connecté : classement par force réelle') +
    ` · Force / sem. = meilleur alignement actif (${CONFIG.LINEUP.F} F, ${CONFIG.LINEUP.D} D, ${CONFIG.LINEUP.UTIL} UTIL, ${CONFIG.LINEUP.G} G) × ${CONFIG.GAMES_PER_WEEK} matchs, gardiens selon leur part de départs, blessés/suspendus exclus` +
    ' · All-play = ta fiche si tu affrontais les 11 autres équipes chaque semaine · Chance = victoires réelles − victoires méritées selon l\'all-play · Chance cette semaine : +0,7 = tu mènes alors que tu ne battrais que 3 équipes sur 11');
  sh.getRange('A3').setValue(WARN.length ? '⚠ ' + WARN.join('  ·  ') : '');
  style_(() => sh.getRange('A3').setFontColor('#C00000'));

  // relie chaque feuille d'équipe à sa fiche ESPN
  const sheetToId = {};
  Object.keys(ESPN_SHEET_OF).forEach(id => { sheetToId[ESPN_SHEET_OF[id]] = +id; });
  const nameOfId = id => ESPN_SHEET_OF[id] || (st && st.byId[id] ? st.byId[id].espnName : '');

  teams.forEach(t => {
    t.powReal = teamPower_(t, 'real');
    t.powProj = teamPower_(t, 'proj');
    t.out = t.players.filter(p => OUT_STATUSES[p.status]).map(p => p.name);
    t.espn = st ? st.byId[sheetToId[t.team]] || null : null;
  });
  const realRank = {};
  teams.slice().sort((a, b) => b.powReal - a.powReal).forEach((t, i) => { realRank[t.team] = i + 1; });

  const pct = e => { const g = e.w + e.l + e.tie; return g ? (e.w + e.tie / 2) / g : 0; };
  teams.sort((a, b) => {
    if (a.espn && b.espn) return (pct(b.espn) - pct(a.espn)) || (b.espn.w - a.espn.w) || (b.espn.pf - a.espn.pf);
    if (a.espn || b.espn) return a.espn ? -1 : 1;
    return b.powReal - a.powReal;
  });
  const fmt = x => x ? `${x.name} (${x.diff >= 0 ? '+' : ''}${x.diff.toFixed(2)})` : '';
  const rows = teams.map((t, i) => {
    const e = t.espn;
    const ap = e ? e.apW + e.apL + e.apT : 0;
    const apPct = ap ? (e.apW + e.apT / 2) / ap : '';
    const games = e ? e.w + e.l + e.tie : 0;
    // texte « 2 V – 1 D » : jamais interprété comme une date par Sheets
    const rec = (w, l, tie) => `${w} V – ${l} D` + (tie ? ` – ${tie} N` : '');
    const luck = e && ap ? Math.round((e.w + e.tie / 2 - apPct * games) * 10) / 10 : '';
    // chance de la semaine : résultat actuel contre ton adversaire (1 = en avance, ½ = égalité, 0 = en retard)
    // moins ta part de victoires contre toutes les équipes. > 0 = ton adversaire est plus facile que la moyenne.
    const wk = e ? e.wkW + e.wkL + e.wkT : 0;
    const weekLuck = e && e.live !== '' && e.oppLive !== '' && wk
      ? Math.round(((e.live > e.oppLive ? 1 : e.live < e.oppLive ? 0 : 0.5) - (e.wkW + e.wkT / 2) / wk) * 10) / 10 : '';
    return [i + 1, t.team,
            e ? rec(e.w, e.l, e.tie) : '', e ? pct(e) : '', e ? e.pf : '', e ? e.pa : '',
            e && e.live !== '' ? e.live : '', e && e.opp !== '' ? `${nameOfId(e.opp)} (${(+e.oppLive).toFixed(1)})` : '',
            e && ap ? rec(e.apW, e.apL, e.apT) : '',
            e && e.live !== '' && (e.wkW + e.wkL + e.wkT) ? rec(e.wkW, e.wkL, e.wkT) : '', weekLuck, luck,
            Math.round(t.powReal), realRank[t.team], Math.round(t.powProj), Math.round((t.powReal - t.powProj) * 10) / 10,
            t.out.join(', '), t.worst && t.worst.diff < 0 ? fmt(t.worst) : '', t.total];
  });
  if (rows.length) {
    sh.getRange(5, 1, rows.length, N).setValues(rows);
    // Nom d'équipe = lien vers son alignement (formule HYPERLINK). Si le Tableau refuse la formule,
    // les noms restent en texte et les liens vont dans une colonne « Aller à l'alignement ».
    try {
      const sep = formulaSep_(ss, sh);
      const link = (r, label) => {
        const t = ss.getSheetByName(r[1]);
        return t ? `=HYPERLINK("#gid=${t.getSheetId()}"${sep}"${String(label).replace(/"/g, '""')}")` : String(label);
      };
      const props = PropertiesService.getDocumentProperties();
      if (props.getProperty('LINKS_IN_M') !== '1') {
        sh.getRange(5, 2, rows.length, 1).setFormulas(rows.map(r => [link(r, r[1])]));
        SpreadsheetApp.flush();
        const bad = sh.getRange(5, 2, rows.length, 1).getDisplayValues().some(v => String(v[0]).charAt(0) === '#');
        if (bad) props.setProperty('LINKS_IN_M', '1');
      }
      if (props.getProperty('LINKS_IN_M') === '1') {
        sh.getRange(5, 2, rows.length, 1).setValues(rows.map(r => [r[1]]));
        sh.getRange(4, N + 1).setValue("Aller à l'alignement");
        sh.getRange(5, N + 1, rows.length, 1).setFormulas(rows.map(r => [link(r, '→ ' + r[1])]));
      }
    } catch (e) { WARN.push('Liens des équipes : ' + e.message); }
  }
  clearBelow_(sh, rows.length + 5, N + 1);
}

function stamp_() {
  return Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm');
}

// ======================= SYNCHRO ESPN =======================
function connectEspn() {
  const ui = SpreadsheetApp.getUi();
  const a = ui.prompt('Connecter ESPN (1/2)', 'Colle la valeur du cookie espn_s2 :', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK || !a.getResponseText().trim()) return;
  const b = ui.prompt('Connecter ESPN (2/2)', 'Colle la valeur du cookie SWID (avec les accolades { }) :', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK || !b.getResponseText().trim()) return;
  // Stockés dans TON compte Google (propriétés utilisateur), invisibles pour les autres éditeurs du fichier
  PropertiesService.getUserProperties().setProperties({
    ESPN_S2: a.getResponseText().trim(), ESPN_SWID: b.getResponseText().trim()
  });
  ui.alert('ESPN connecté. Lance « Mettre à jour maintenant » pour synchroniser les alignements.');
}
function disconnectEspn() {
  const up = PropertiesService.getUserProperties();
  up.deleteProperty('ESPN_S2'); up.deleteProperty('ESPN_SWID');
  SpreadsheetApp.getActive().toast('ESPN déconnecté.');
}

function normTeam_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

const ESPN_HOSTS = ['lm-api-reads.fantasy.espn.com', 'fantasy.espn.com'];

function espnUrl_(host) {
  return `https://${host}/apis/v3/games/fhl/seasons/${CONFIG.ESPN_SEASON}` +
         `/segments/0/leagues/${CONFIG.ESPN_LEAGUE_ID}` +
         '?view=mTeam&view=mRoster&view=mMatchupScore&view=mLiveScoring&view=mStatus&view=mSettings';
}
function espnOptions_(withCookies) {
  const up = PropertiesService.getUserProperties();
  const s2 = up.getProperty('ESPN_S2'), swid = up.getProperty('ESPN_SWID');
  const opts = { muteHttpExceptions: true, followRedirects: true,
                 headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' } };
  if (withCookies && s2 && swid) opts.headers.Cookie = `espn_s2=${s2}; SWID=${swid}`;
  return opts;
}

// Essaie chaque adresse ESPN (2 tentatives chacune) ; renvoie la première réponse valide
function fetchEspnTeams_() {
  const errors = [];
  for (const host of ESPN_HOSTS) {
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        const res = UrlFetchApp.fetch(espnUrl_(host), espnOptions_(true));
        const code = res.getResponseCode();
        if (code === 401 || code === 403) {
          errors.push(`${host} : accès refusé (${code}) — reconnecte ESPN, cookies absents ou expirés`);
          break;
        }
        if (code !== 200) { errors.push(`${host} : code ${code}`); break; }
        let json;
        try { json = JSON.parse(res.getContentText()); }
        catch (e) { errors.push(`${host} : réponse illisible (ligue privée ? reconnecte ESPN)`); break; }
        ESPN_DATA = json;
        return (json.teams || []).map(t => {
          const entries = ((t.roster || {}).entries || [])
            .map(e => (e.playerPoolEntry && e.playerPoolEntry.player) || null).filter(p => p && p.fullName);
          entries.forEach(p => { ESPN_STATUS[norm_(p.fullName)] = p.injuryStatus || (p.injured ? 'OUT' : ''); });
          return {
            id: t.id,
            name: (t.name || `${t.location || ''} ${t.nickname || ''}`).trim(),
            players: entries.map(p => p.fullName)
          };
        });
      } catch (e) {
        errors.push(`${host} (essai ${attempt}) : ${e.message.split(':')[0]}`);
        Utilities.sleep(1500);
      }
    }
  }
  throw new Error(errors.join(' | '));
}

// Menu « Diagnostic ESPN » : teste chaque adresse avec et sans cookies, et affiche le résultat
function diagnoseEspn() {
  const lines = [];
  const up = PropertiesService.getUserProperties();
  lines.push('Cookies enregistrés : ' + (up.getProperty('ESPN_S2') && up.getProperty('ESPN_SWID') ? 'oui' : 'NON'));
  lines.push('Ligue ' + CONFIG.ESPN_LEAGUE_ID + ', saison ' + CONFIG.ESPN_SEASON);
  ESPN_HOSTS.forEach(host => [true, false].forEach(ck => {
    let r;
    try {
      const res = UrlFetchApp.fetch(espnUrl_(host), espnOptions_(ck));
      const txt = res.getContentText();
      let info = '';
      try { const j = JSON.parse(txt); info = `${(j.teams || []).length} équipes`; } catch (e) { info = 'pas du JSON'; }
      r = `code ${res.getResponseCode()}, ${info}`;
    } catch (e) { r = 'ERREUR : ' + e.message.split(':')[0]; }
    lines.push(`${host} ${ck ? 'avec' : 'sans'} cookies → ${r}`);
  }));
  // Contrôle : Google peut-il joindre un autre site ?
  try { lines.push('Contrôle API LNH → code ' + UrlFetchApp.fetch('https://api.nhle.com/stats/rest/en/season', { muteHttpExceptions: true }).getResponseCode()); }
  catch (e) { lines.push('Contrôle API LNH → ERREUR : ' + e.message); }
  console.log(lines.join('\n'));
  SpreadsheetApp.getUi().alert('Diagnostic ESPN', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// Met la colonne « Joueur » d'une équipe à jour sans déplacer les joueurs déjà présents :
// les joueurs partis laissent leur ligne vide, les nouveaux prennent les lignes vides.
function applyRoster_(sh, players, aliases) {
  const key = n => { const k = norm_(n); return aliases[k] || k; };
  const rng = sh.getRange(`B2:B${CONFIG.LAST_ROW}`);
  const cur = rng.getValues().map(r => r[0]);
  const espnKeys = {};
  players.forEach(p => { espnKeys[key(p)] = true; });
  const next = cur.slice(), removed = [], added = [], overflow = [];
  cur.forEach((n, i) => { if (n && !espnKeys[key(n)]) { removed.push(n); next[i] = ''; } });
  const present = {};
  next.forEach(n => { if (n) present[key(n)] = true; });
  players.forEach(p => {
    if (present[key(p)]) return;
    const i = next.indexOf('');
    if (i < 0) { overflow.push(p); return; }
    next[i] = p; present[key(p)] = true; added.push(p);
  });
  if (removed.length || added.length) rng.setValues(next.map(v => [v]));
  const parts = [];
  if (added.length) parts.push('+ ' + added.join(', '));
  if (removed.length) parts.push('− ' + removed.join(', '));
  return { changed: parts.length > 0, text: parts.join('  ·  ') || 'Aucun changement', overflow: overflow,
           added: added, removed: removed };
}

function syncEspn_(ss, aliases) {
  if (!CONFIG.ESPN_LEAGUE_ID) return;
  const sh = ss.getSheetByName(CONFIG.ESPN_SHEET) || ss.insertSheet(CONFIG.ESPN_SHEET);
  const header = ['Équipe ESPN', 'Feuille liée (modifiable)', 'Joueurs ESPN', 'Changements (dernière synchro)', `Non placés (plus de ${CONFIG.LAST_ROW - 1} joueurs)`];
  once_('espn', () => {
    header_(sh.getRange(1, 1, 1, header.length), [header], '#1F3864');
    style_(() => sh.setFrozenRows(1));
    style_(() => sh.getRange('G1').setFontColor('#808080'));
  });

  // Liens déjà établis (ou écrits à la main) : Équipe ESPN → nom de feuille
  const links = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([t, f]) => { if (t && f) links[t] = String(f); });
  }

  let teams;
  M_('ESPN : téléchargement');
  try { teams = fetchEspnTeams_(); }
  catch (e) {
    sh.getRange('G1').setValue(`⚠ Synchro ESPN impossible (${stamp_()}) : ${e.message}. Les alignements n'ont pas été modifiés.`);
    WARN.push('Synchro ESPN impossible — voir la feuille ESPN');
    return;
  }
  if (!teams.length) {
    sh.getRange('G1').setValue(`⚠ ESPN n'a renvoyé aucune équipe (${stamp_()}). Vérifie ESPN_LEAGUE_ID / ESPN_SEASON.`);
    return;
  }

  M_('ESPN : repérer les feuilles d’équipe');
  let byKey = {};
  teamSheets_(ss).forEach(s => { byKey[normTeam_(s.getName())] = s; });
  // une équipe ESPN sans feuille connue → tour complet du classeur (nouvelle feuille, renommée…)
  if (teams.some(t => !(links[t.name] && ss.getSheetByName(links[t.name])) && !byKey[normTeam_(t.name)])) {
    byKey = {}; teamSheets_(ss, true).forEach(s => { byKey[normTeam_(s.getName())] = s; });
  }
  M_('ESPN : écriture dans les feuilles');

  // Une équipe qui change de nom sur ESPN garde son numéro : on retrouve sa feuille par ce numéro (gardé à chaque synchro),
  // sinon par élimination (une seule équipe ESPN et une seule feuille sans partenaire)
  const dp = PropertiesService.getDocumentProperties();
  let idMap = {}; try { idMap = JSON.parse(dp.getProperty('ESPN_ID_SHEET') || '{}'); } catch (e) { idMap = {}; }
  const sheetFor = t => {
    if (links[t.name] && ss.getSheetByName(links[t.name])) return ss.getSheetByName(links[t.name]);
    if (byKey[normTeam_(t.name)]) return byKey[normTeam_(t.name)];
    if (idMap[t.id] && ss.getSheetByName(idMap[t.id])) return ss.getSheetByName(idMap[t.id]);
    return null;
  };
  const found = {}; teams.forEach(t => { const s0 = sheetFor(t); if (s0) found[t.id] = s0; });
  const lonelyT = teams.filter(t => !found[t.id]);
  const usedS = {}; Object.keys(found).forEach(id => { usedS[found[id].getName()] = true; });
  const lonelyS = teamSheets_(ss).filter(s0 => !usedS[s0.getName()]);
  if (lonelyT.length === 1 && lonelyS.length === 1) { found[lonelyT[0].id] = lonelyS[0]; WARN.push(`Équipe ESPN « ${lonelyT[0].name} » reliée à la feuille « ${lonelyS[0].getName()} » (nouveau nom ?)`); }
  const newMap = {}; Object.keys(found).forEach(id => { newMap[id] = found[id].getName(); });
  try { dp.setProperty('ESPN_ID_SHEET', JSON.stringify(newMap)); } catch (e) { /* rien */ }

  let changed = false;
  const moves = [];   // pour le News Tracker
  const rows = teams.map(t => {
    const team = found[t.id] || null;
    if (!team || !isTeamSheet_(team)) {
      return [t.name, links[t.name] || '', t.players.length, '⚠ Aucune feuille trouvée : écris le nom exact de la feuille en colonne B', ''];
    }
    ESPN_SHEET_OF[t.id] = team.getName();
    if (t.players.length < CONFIG.ESPN_MIN_ROSTER) {
      return [t.name, team.getName(), t.players.length, '⚠ Alignement ESPN incomplet — ignoré par sécurité', ''];
    }
    M_('ESPN : ' + team.getName());
    const r = applyRoster_(team, t.players, aliases);
    changed = changed || r.changed;
    if (r.changed) moves.push({ team: team.getName(), added: r.added, removed: r.removed });
    return [t.name, team.getName(), t.players.length, r.text, r.overflow.join(', ')];
  });
  M_('ESPN : feuille ESPN');
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  clearBelow_(sh, rows.length + 2, header.length);
  sh.getRange('G1').setValue(`✔ Synchro ESPN : ${stamp_()}`);
  M_('ESPN : News Tracker + recalcul');
  if (moves.length) { try { logTransactions_(ss, moves); } catch (e) { WARN.push('News Tracker : ' + e.message); } }
  if (changed) SpreadsheetApp.flush();   // laisse les formules GAF se recalculer avant la suite
}

// ======================= NEWS TRACKER =======================
// Gauche (A:D) : transactions de la ligue, détectées à chaque synchro ESPN (échanges, ajouts/drops, réclamations).
// Droite (F:O) : meilleurs joueurs de la soirée précédente — mis à jour une fois par jour vers 8 h.
function newsSheet_(ss) {
  let sh = ss.getSheetByName(CONFIG.NEWS_SHEET);
  if (sh && sh.getRange('A4').getValue() === 'Date') return sh;
  if (!sh) sh = ss.insertSheet(CONFIG.NEWS_SHEET, 1);
  const save = DATA_ONLY; DATA_ONLY = false;                // mise en page toujours faite à la création
  sh.getRange('A1').setValue('🗞 NEWS TRACKER');
  sh.getRange('A3').setValue('Transactions récentes (ESPN)');
  sh.getRange('A4:D4').setValues([['Date', 'Type', 'Équipe(s)', 'Détail']]);
  sh.getRange('F3').setValue('Meilleurs de la soirée précédente');
  sh.getRange('F4:O4').setValues([['Rang', 'Joueur', 'Équipe LNH', 'Pos', 'G', 'A', 'PTS', 'Tirs', 'FanPts (soirée)', 'Équipe fantasy']]);
  style_(() => sh.getRange('A1').setFontSize(16).setFontWeight('bold'));
  style_(() => sh.getRange('A3').setFontSize(12).setFontWeight('bold'));
  style_(() => sh.getRange('F3').setFontSize(12).setFontWeight('bold'));
  ['A4:D4', 'F4:O4'].forEach(a => style_(() => sh.getRange(a).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF')));
  style_(() => sh.setFrozenRows(4));
  style_(() => { sh.setColumnWidth(1, 120); sh.setColumnWidth(2, 130); sh.setColumnWidth(3, 190); sh.setColumnWidth(4, 420);
                 sh.setColumnWidth(5, 20); sh.setColumnWidth(7, 170); sh.setColumnWidth(15, 170); });
  style_(() => sh.getRange('D5:D' + (5 + CONFIG.NEWS_MAX_LOG)).setWrap(true));
  DATA_ONLY = save;
  return sh;
}

// moves : [{ team, added: [noms], removed: [noms] }] issus d'une même synchro
function logTransactions_(ss, moves) {
  const key = n => norm_(n);
  const addedBy = {};                                   // joueur → équipe qui l'a ajouté
  moves.forEach(m => m.added.forEach(p => { addedBy[key(p)] = { team: m.team, name: p }; }));
  const used = {};                                      // joueurs déjà expliqués
  const flows = [];                                     // A → B (joueur passé d'une équipe à l'autre)
  moves.forEach(m => m.removed.forEach(p => {
    const to = addedBy[key(p)];
    if (to && to.team !== m.team) flows.push({ from: m.team, to: to.team, name: to.name });
  }));
  const rows = [];
  const when = stamp_();
  // Échanges : des joueurs dans les deux sens entre deux mêmes équipes
  const pairs = {};
  flows.forEach(f => { const k = [f.from, f.to].sort().join('|'); (pairs[k] = pairs[k] || []).push(f); });
  Object.keys(pairs).forEach(k => {
    const [a, b] = k.split('|');
    const aGets = pairs[k].filter(f => f.to === a).map(f => f.name);
    const bGets = pairs[k].filter(f => f.to === b).map(f => f.name);
    if (aGets.length && bGets.length) {
      rows.push([when, '🔁 Échange', `${a} ↔ ${b}`, `${a} reçoit ${aGets.join(', ')} · ${b} reçoit ${bGets.join(', ')}`]);
      pairs[k].forEach(f => { used[key(f.name)] = true; });
    } else {
      pairs[k].forEach(f => {
        rows.push([when, '📋 Réclamation', f.to, `${f.to} réclame ${f.name} (libéré par ${f.from})`]);
        used[key(f.name)] = true;
      });
    }
  });
  // Ajouts / drops restants, par équipe
  moves.forEach(m => {
    const add = m.added.filter(p => !used[key(p)]);
    const drop = m.removed.filter(p => !used[key(p)]);
    if (add.length && drop.length) rows.push([when, '🔄 Ajout / Drop', m.team, `drop ${drop.join(', ')} pour ajouter ${add.join(', ')}`]);
    else if (add.length) rows.push([when, '➕ Ajout', m.team, `ajoute ${add.join(', ')}`]);
    else if (drop.length) rows.push([when, '➖ Drop', m.team, `drop ${drop.join(', ')}`]);
  });
  if (!rows.length) return;

  const sh = newsSheet_(ss);
  const n = CONFIG.NEWS_MAX_LOG;
  const old = sh.getRange(5, 1, n, 4).getValues().filter(r => r[0] !== '');
  const all = rows.concat(old).slice(0, n);             // plus récentes en haut
  sh.getRange(5, 1, all.length, 4).setValues(all);
}

// Meilleurs joueurs de la soirée d'hier (matchs terminés), classés par FanPts de la soirée
function updateNews_(ss) {
  const sh = newsSheet_(ss);
  const day = ymd_(new Date(Date.now() - 86400000));
  const j = fetchJsonAll_([`${WEB}/schedule/${day}`])[0];
  const games = [];
  (j.gameWeek || []).forEach(d => {
    if (d.date !== day) return;
    (d.games || []).forEach(g => {
      if (g.gameType === CONFIG.GAME_TYPE && (g.gameState === 'FINAL' || g.gameState === 'OFF')) games.push(g.id);
    });
  });
  const tot = {};
  for (let k = 0; k < games.length; k += CONFIG.GAMES_PER_BATCH) {
    const ids = games.slice(k, k + CONFIG.GAMES_PER_BATCH);
    const urls = [];
    ids.forEach(id => urls.push(`${WEB}/gamecenter/${id}/boxscore`, `${WEB}/gamecenter/${id}/play-by-play`));
    const data = fetchJsonAll_(urls);
    ids.forEach((id, i) => parseGame_(data[2 * i], data[2 * i + 1]).forEach(l => addLine_(tot, l)));
  }
  fillNames_(tot);
  const list = scoreList_(Object.keys(tot).map(k => tot[k]).filter(l => l.name));

  // propriétaire fantasy (index des alignements)
  const owner = {};
  const props = PropertiesService.getDocumentProperties().getProperties();
  let teams = [];
  try { teams = JSON.parse(props.TEAM_SHEETS || '[]'); } catch (e) { teams = []; }
  const aliases = readAliases_(ss);
  teams.forEach(t => { try { JSON.parse(props['ROSTER:' + t] || '[]').forEach(k => {
    if (!k) return;
    owner[k] = t;
    if (aliases[k]) owner[aliases[k]] = t;               // ex. « Sam Montembeault » → « Samuel Montembeault »
  }); } catch (e) { /* rien */ } });
  const ownerOf = n => owner[norm_(n)] || owner[aliases[norm_(n)]] || 'Waivers (libre)';

  const top = list.slice(0, CONFIG.NEWS_TOP).map((p, i) => {
    const sk = p.pos !== 'G';
    return [i + 1, p.name, p.team, p.pos, sk ? p.g : '', sk ? p.a : '', sk ? p.g + p.a : '', sk ? p.sog : '', p.fp, ownerOf(p.name)];
  });
  const L = CONFIG.NEWS_TOP;
  const out = top.concat(Array(Math.max(0, L - top.length)).fill(Array(10).fill('')));
  sh.getRange(5, 6, L, 10).setValues(out);
  sh.getRange('F2').setValue(games.length
    ? `Soirée du ${day} · ${games.length} match(s) · mis à jour ${stamp_()}`
    : `Aucun match terminé le ${day} · vérifié ${stamp_()}`);
  PropertiesService.getDocumentProperties().setProperty('NEWS_DATE', ymd_(new Date()));
}

// Filet de sécurité : si le déclencheur de 8 h n'a pas tourné, le premier update après 8 h s'en charge
function maybeDailyNews_(ss) {
  const today = ymd_(new Date());
  const hour = +Utilities.formatDate(new Date(), 'America/Toronto', 'H');
  if (hour < CONFIG.NEWS_HOUR || PropertiesService.getDocumentProperties().getProperty('NEWS_DATE') === today) return;
  try { updateNews_(ss); } catch (e) { WARN.push('News Tracker : ' + e.message); }
}

// Déclencheur quotidien (8 h) et menu
function newsDaily() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  updateNews_(ss);
}

// ---------- Mise à jour automatique ----------
function enableAutoRefresh() {
  disableAutoRefresh();
  ScriptApp.newTrigger('updateData').timeBased().everyMinutes(CONFIG.REFRESH_MINUTES).create();
  ScriptApp.newTrigger('newsDaily').timeBased().everyDays(1).atHour(CONFIG.NEWS_HOUR).inTimezone('America/Toronto').create();
  SpreadsheetApp.getActive().toast(`Mise à jour automatique activée (toutes les ${CONFIG.REFRESH_MINUTES} min).`);
}
function disableAutoRefresh() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['updateLive', 'updateData', 'newsDaily'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
}
