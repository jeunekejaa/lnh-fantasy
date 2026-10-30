/**
 * 🏒 LNH LIVE — APP (fichier 2 du projet Apps Script)
 * Projection maison + stats avancées (MoneyPuck) + données de l'app web.
 *
 * À chaque mise à jour, run_() appelle buildAppSnapshot_() qui :
 *   1. lit les stats avancées MoneyPuck (1 fois par jour, gardées dans la feuille cachée « MP Cache ») ;
 *   2. calcule pour chaque joueur une « projection maison » (FP par match) stat par stat ;
 *   3. calcule un signal de chance (🟢 malchanceux / 🔴 chanceux / ⚪ normal) ;
 *   4. lit le calendrier LNH (matchs restants cette semaine, la semaine prochaine) ;
 *   5. simule le reste de la saison (chances de séries) ;
 *   6. range le tout dans la feuille cachée « APP Data », lue par l'app web (doGet).
 *
 * Données avancées : MoneyPuck.com (usage personnel, non commercial — source citée dans l'app).
 */

const APP = {
  MY_TEAM: 'dans mon inoxtag kaizen era',   // nom de TA feuille d'équipe
  DATA_SHEET: 'APP Data',
  MP_SHEET: 'MP Cache',
  SIMS: 2000,              // saisons simulées pour les chances de séries
  WEEK_SD: 24,             // écart-type d'un pointage hebdomadaire (mesuré sur les semaines 1 à 3)
  GAMES_LEFT_DEFAULT: 3,   // matchs par semaine quand le calendrier est inconnu
  INJURY_FACTOR: 0.94,     // part des matchs réellement joués (blessures, repos)
  GOALIE_BIAS: -0.5,       // GAF surestime les gardiens (optimum du backtest 2025-26)
  // Vitesse de confiance dans la saison en cours, stat par stat : k = « matchs de projection » mélangés
  // aux matchs réels. Petit k = on croit vite la saison en cours.
  // Valeurs réglées par le backtest 2025-26 (561 patineurs, 35 gardiens) :
  //   g = buts réels (sans MoneyPuck) ; gx = buts attendus (MoneyPuck) — le backtest avec « tirs × % habituel »
  //   donne k = 15 : les chances de marquer sont fiables bien plus vite que les buts eux-mêmes (k = 60).
  K: { sog: 15, hit: 12, blk: 40, g: 60, gx: 15, a: 80, ppp: 50, shp: 60, gstart: 100 },
  LUCK: 0.5,               // écart (FP/match) au-delà duquel un joueur est déclaré chanceux / malchanceux
  // Rôle des gardiens que le modèle ne peut pas deviner (blessure du partant, échange…) :
  // part des départs à venir, de 0 à 1. Mets à jour au besoin ; laisse vide sinon.
  // ex. { 'Nom du gardien': 0.85 }
  GOALIE_ROLE: {
    // ex. 'Nom du gardien': 0.85   (aucun pour l'instant : la part vient des départs réels)
  }
};

// ---------------- App web ----------------
// ?k=<clé> renvoie les données brutes (JSON, lecture seule) — pour que Claude puisse consulter l'app.
// Change la clé (ou mets '') pour couper cet accès.
const SHARE_KEY = 'HZwSztv5v7cUf41u';
function doGet(e) {
  if (e && e.parameter && e.parameter.k && SHARE_KEY && e.parameter.k === SHARE_KEY) {
    const what = e.parameter.what;
    if (what === 'pts' && e.parameter.n) {          // « Mon estimé » d'un joueur (total de points de la saison) ; pts vide = effacer
      const cur = getOverrides_()[e.parameter.n] || {};
      return ContentService.createTextOutput(saveOverride(e.parameter.n, Object.assign({}, cur, { pts: e.parameter.pts || '' }))).setMimeType(ContentService.MimeType.JSON);
    }
    if (what === 'start' && e.parameter.n) {        // « Partant confirmé le … » d'un gardien (d vide = effacer)
      const cur = getOverrides_()[e.parameter.n] || {};
      return ContentService.createTextOutput(saveOverride(e.parameter.n, Object.assign({}, cur, { start: e.parameter.d || '' }))).setMimeType(ContentService.MimeType.JSON);
    }
    if (what === 'runbt') {                          // relance le backtest 2025-26 (le même que le menu) et renvoie ses résultats
      const L = LockService.getScriptLock(); if (!L.tryLock(2000)) return ContentService.createTextOutput('{"err":"occupé"}');
      try { runBacktest(); } finally { L.releaseLock(); }
      return ContentService.createTextOutput(getBacktestData() || '{}').setMimeType(ContentService.MimeType.JSON);
    }
    if (what === 'ver') return ContentService.createTextOutput(codeStatus_()).setMimeType(ContentService.MimeType.JSON);
    if (what === 'update') return ContentService.createTextOutput(checkCodeFromApp()).setMimeType(ContentService.MimeType.JSON);
    if (what === 'refresh') return ContentService.createTextOutput(refreshLocked_(false)).setMimeType(ContentService.MimeType.JSON);
    const body = what === 'bt' && typeof getBacktestData === 'function' ? getBacktestData()
      : what === 'pj' && typeof getProjectionResult === 'function' ? getProjectionResult()
      : what === 'mp' ? getMpProbe_()
      : getAppData();
    return ContentService.createTextOutput(body || '{}').setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('Interface').evaluate()
    .setTitle('LNH Fantasy')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}
// Appelée par la page (google.script.run.getAppData)
function getAppData() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP.DATA_SHEET);
  if (!sh || sh.getLastRow() < 1) return '';
  return sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => r[0]).join('');
}
// Menu « 📱 Ouvrir l'app » : affiche le lien de l'app (une fois déployée)
function openApp() {
  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  const html = url
    ? `<p style="font-family:sans-serif">Ton app : <a href="${url}" target="_blank">ouvrir LNH Fantasy</a></p><p style="font-family:sans-serif;color:#666;font-size:12px">${url}</p>`
    : '<p style="font-family:sans-serif">L’app n’est pas encore déployée : Extensions › Apps Script › Déployer › Nouveau déploiement › Application Web.</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(140), '📱 LNH Fantasy');
}
// Progression d'une mise à jour (lue par l'app toutes les secondes pendant le chargement)
// Chronométrage de chaque étape d'une mise à jour (gardé dans LAST_TIMING, lu avec ?what=ver)
let TMARK = [];
function mark_(label) { TMARK.push([label, Date.now()]); }
function saveTiming_() {
  try {
    if (TMARK.length < 2) return;
    const out = { at: stamp_(), total: Math.round((TMARK[TMARK.length - 1][1] - TMARK[0][1]) / 100) / 10, steps: [] };
    for (let i = 0; i < TMARK.length - 1; i++) out.steps.push([TMARK[i][0], Math.round((TMARK[i + 1][1] - TMARK[i][1]) / 100) / 10]);
    PropertiesService.getScriptProperties().setProperty('LAST_TIMING', JSON.stringify(out));
  } catch (e) { /* rien */ }
}
function setProgress_(pct, label) {
  if (pct <= 3) TMARK = [];
  mark_(label);
  if (pct >= 100) saveTiming_();
  try { CacheService.getScriptCache().put('APP_PROGRESS', JSON.stringify({ pct: pct, label: label, t: Date.now() }), 600); } catch (e) { /* rien */ }
}
function getProgress() { return CacheService.getScriptCache().get('APP_PROGRESS') || ''; }
// Bouton « Mettre à jour » de l'app
function refreshFromApp() { refreshLocked_(true); return getAppData(); }
// Une seule mise à jour à la fois (bouton de l'app, ou ?what=refresh)
function refreshLocked_(wait) {
  const L = LockService.getScriptLock();
  if (!L.tryLock(wait ? 300000 : 2000)) return JSON.stringify({ ok: false, err: 'Une mise à jour est déjà en cours.' });
  try { updateData(); } finally { L.releaseLock(); }
  return codeStatus_();
}

function hiddenSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); try { sh.hideSheet(); } catch (e) {} }
  return sh;
}

// ---------------- MoneyPuck ----------------
// Garde, par joueur (id LNH), quelques stats avancées de la saison en cours, situation « all » et « 5on4 ».
function moneyPuck_(ss) {
  const sh = hiddenSheet_(ss, APP.MP_SHEET);
  const today = ymd_(new Date());
  const NC = 14;
  if (String(sh.getRange('A1').getDisplayValue()) === today && sh.getLastRow() > 2 && sh.getLastColumn() >= NC) {   // texte : évite que Sheets en fasse une date
    const map = {};
    sh.getRange(3, 1, sh.getLastRow() - 2, NC).getValues().forEach(r => {
      map[r[0]] = { gp: r[1], toi: r[2], ixg: r[3], onxg: r[4], ong: r[5], pptoi: r[6], gxg: r[7], gga: r[8],
                    a1: +r[9] || 0, a2: +r[10] || 0, att: +r[11] || 0, ppatt: +r[12] || 0, sogm: +r[13] || 0 };
    });
    return map;
  }
  const year = String(CONFIG.SEASON).slice(0, 4);
  const base = `https://moneypuck.com/moneypuck/playerData/seasonSummary/${year}/regular/`;
  const res = UrlFetchApp.fetchAll([base + 'skaters.csv', base + 'goalies.csv'].map(u => ({ url: u, muteHttpExceptions: true })));
  if (res[0].getResponseCode() !== 200) throw new Error('MoneyPuck indisponible (code ' + res[0].getResponseCode() + ')');
  const map = {};
  const csv = t => Utilities.parseCsv(t);
  const col = (h, names) => { for (const n of names) { const i = h.indexOf(n); if (i >= 0) return i; } return -1; };

  const sk = csv(res[0].getContentText());
  const h = sk[0];
  const c = {
    id: col(h, ['playerId']), sit: col(h, ['situation']), gp: col(h, ['games_played']), toi: col(h, ['icetime']),
    ixg: col(h, ['I_F_xGoals']), onxg: col(h, ['OnIce_F_xGoals']), ong: col(h, ['OnIce_F_goals']),
    a1: col(h, ['I_F_primaryAssists']), a2: col(h, ['I_F_secondaryAssists']), att: col(h, ['I_F_shotAttempts']), sogm: col(h, ['I_F_shotsOnGoal'])
  };
  const num = (r, i) => (i >= 0 ? +r[i] || 0 : 0);
  if (c.id < 0 || c.sit < 0 || c.ixg < 0) throw new Error('format MoneyPuck inattendu (colonnes introuvables)');
  for (let i = 1; i < sk.length; i++) {
    const r = sk[i], id = +r[c.id];
    if (!id) continue;
    const m = map[id] = map[id] || { gp: 0, toi: 0, ixg: 0, onxg: 0, ong: 0, pptoi: 0, gxg: '', gga: '', a1: 0, a2: 0, att: 0, ppatt: 0, sogm: 0 };
    if (r[c.sit] === 'all') {
      m.gp = +r[c.gp] || 0; m.toi = +r[c.toi] || 0; m.ixg = +r[c.ixg] || 0;
      m.onxg = num(r, c.onxg); m.ong = num(r, c.ong); m.a1 = num(r, c.a1); m.a2 = num(r, c.a2); m.att = num(r, c.att); m.sogm = num(r, c.sogm);
    } else if (r[c.sit] === '5on4') { m.pptoi = +r[c.toi] || 0; m.ppatt = num(r, c.att); }
  }
  if (res[1].getResponseCode() === 200) {
    const gk = csv(res[1].getContentText()), g = gk[0];
    const gi = { id: col(g, ['playerId']), sit: col(g, ['situation']), gp: col(g, ['games_played']),
                 xg: col(g, ['xGoals']), ga: col(g, ['goals']) };
    for (let i = 1; i < gk.length; i++) {
      const r = gk[i], id = +r[gi.id];
      if (!id || r[gi.sit] !== 'all') continue;
      map[id] = { gp: +r[gi.gp] || 0, toi: 0, ixg: 0, onxg: 0, ong: 0, pptoi: 0, gxg: +r[gi.xg] || 0, gga: +r[gi.ga] || 0, a1: 0, a2: 0, att: 0, ppatt: 0, sogm: 0 };
    }
  }
  const rows = Object.keys(map).map(id => { const m = map[id]; return [+id, m.gp, m.toi, m.ixg, m.onxg, m.ong, m.pptoi, m.gxg, m.gga, m.a1, m.a2, m.att, m.ppatt, m.sogm]; });
  sh.clearContents();
  sh.getRange('A1').setValue("'" + today);
  sh.getRange('A2').setValue('Source : MoneyPuck.com — playerId, MJ, temps de jeu (s), xG individuels, xG sur la glace, buts sur la glace, temps en AN 5c4 (s), xG contre (gardien), buts contre (gardien), passes primaires, passes secondaires, tentatives de tir, tentatives en AN 5c4, tirs cadrés');
  if (rows.length) sh.getRange(3, 1, rows.length, NC).setValues(rows);
  return map;
}

// ---------------- GAF : stats projetées par match ----------------
function gafStats_(ss) {
  const sh = ss.getSheetByName(CONFIG.GAF_SHEET);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const v = sh.getRange(1, 1, sh.getLastRow(), Math.min(sh.getLastColumn(), 27)).getValues();
  const h = v[0].map(x => String(x).trim());
  const i = n => h.indexOf(n);
  const I = { name: i('PLAYER'), pos: i('POS'), gp: i('GP'), fpg: i('FP/GP'), g: i('G'), a: i('A'), ppp: i('PPP'),
              shp: i('SHP'), sog: i('SOG'), hit: i('HIT'), blk: i('BLK'), w: i('W'), ga: i('GA'), sv: i('SV'), so: i('SO'), otl: i('OTL') };
  for (let r = 1; r < v.length; r++) {
    const row = v[r], name = row[I.name];
    if (!name) continue;
    const gp = +row[I.gp] || 0;
    const per = k => (I[k] >= 0 && gp ? (+row[I[k]] || 0) / gp : 0);
    out[norm_(name)] = { pos: String(row[I.pos] || ''), fpg: +row[I.fpg] || 0, gp: gp,
      g: per('g'), a: per('a'), ppp: per('ppp'), shp: per('shp'), sog: per('sog'), hit: per('hit'), blk: per('blk'),
      w: per('w'), ga: per('ga'), sv: per('sv'), so: per('so'), otl: per('otl') };
  }
  return out;
}

// ---------------- Projection maison ----------------
// Valeurs de remplacement quand un joueur n'a pas de projection GAF (joueur de profondeur)
const DEPTH_PRIOR = {
  F: { g: 0.09, a: 0.12, ppp: 0.03, shp: 0, sog: 1.4, hit: 1.1, blk: 0.45 },
  D: { g: 0.03, a: 0.15, ppp: 0.03, shp: 0, sog: 1.2, hit: 1.2, blk: 1.3 }
};
const grp_ = pos => { const p = String(pos || ''); return p === 'G' ? 'G' : (p.split(/[,\/ ]/).indexOf('D') >= 0 ? 'D' : 'F'); };
const fp_ = s => { const S = CONFIG.SCORING; return S.G * s.g + S.A * s.a + S.PPP * s.ppp + S.SHP * s.shp + S.SOG * s.sog + S.HIT * s.hit + S.BLK * s.blk; };
const blend_ = (prior, obs, n, k) => (k * prior + n * obs) / (k + n);

// Réglages appris par le backtest 2025-26 (🧪), sinon valeurs de base de APP.K
let MPG = null, TSTR = null, WKROS = null, CUR_TEAMS = null, WEEKST = null, STARTERS = null, TX_SEEN = {}, OPP_ADJ = null, BT_P = undefined, PJ_GK = undefined, TX_CACHE = [], GK_RECENT = {}, JOURNAL = null, ROLES = null;
function inSeason_() {
  if (BT_P === undefined) { BT_P = null; try { if (typeof btParams_ === 'function') BT_P = btParams_(); } catch (e) { /* rien */ } }
  return BT_P;
}
// p = ligne de stats LNH de la saison (ou null), prior = base d'avant-saison (GAF + moteur, ou null),
// mp = MoneyPuck de la saison (ou null), pj = moteur maison du joueur (ou null) + ratios de la ligue (lg)
function modelSkater_(p, prior, mp, pos, pj, lg, gafP) {
  const g = grp_(pos), pr = Object.assign({}, prior || DEPTH_PRIOR[g]);
  const n = p ? p.gp || 0 : 0;
  const rate = k => (n ? (+p[k] || 0) / n : 0);
  const B = inSeason_();
  const K = Object.assign({}, APP.K, B ? Object.assign({}, B.K, { gx: B.kxg, g: B.gsrc === 'tirs' ? B.K.g : B.kg }) : {});
  // Optimisation finale du backtest (seulement les leviers retenus) : k par position, k selon le nombre de matchs,
  // vitesse de confiance et correction selon la tendance du joueur (progression, chanceux, malchanceux, panne)
  const FN = B && B.final;
  const cls = FN && (FN.mult || FN.bias) ? trendClass_(p, gafP, pj && pj.ly, mp) : null;
  if (FN) {
    const kp = FN.kpos && FN.kpos[g];
    if (kp) Object.keys(kp).forEach(s => { K[s] = kp[s]; if (s === 'g' && B.gsrc === 'xg') K.gx = kp[s]; });
    const m = (FN.gamma && n ? Math.pow(Math.max(1, n) / 15, FN.gamma) : 1) * (cls && FN.mult && FN.mult[cls] != null ? FN.mult[cls] : 1)
            * (FN.alpha && pj && pj.kr ? Math.pow(pj.kr, FN.alpha) : 1);
    if (m !== 1) Object.keys(K).forEach(s => { K[s] *= m; });
  }
  // Joue blessé ? (retenu par le backtest) : temps de jeu des 3 derniers matchs ≤ 80 % de son temps d'avant, 8 matchs et plus
  const inj = FN && FN.inj ? playingInjured_(p, mp) : false;
  if (inj && FN.inj.m !== 1) Object.keys(K).forEach(s => { K[s] *= FN.inj.m; });
  // plancher : le backtest (5 matchs et plus) peut choisir k = 0 ; sur 1 à 4 matchs, ce serait croire le réel aveuglément
  Object.keys(K).forEach(s => { if (K[s] < 5) K[s] = 5; });
  const hasMp = mp && mp.gp > 0;
  // Buts : buts attendus (xG MoneyPuck) au lieu des buts réels ; sans MoneyPuck, tirs × % habituel si le backtest le préfère
  // (la source vient du backtest : buts attendus, tirs × % habituel ou buts réels — celle qui a le mieux prédit)
  let gObs = rate('g'), kG = K.g;
  const xgObs = hasMp ? mp.ixg / mp.gp : null;
  if (B && B.gsrc && B.gsrc !== 'xg') { if (B.gsrc === 'tirs' && pr.sog > 0) gObs = rate('sog') * pr.g / pr.sog; }
  else if (hasMp) { gObs = xgObs; kG = K.gx; }
  else if (B && B.gsrc === 'tirs' && pr.sog > 0) gObs = rate('sog') * pr.g / pr.sog;
  // Passes : passes secondaires ramenées vers ce que ses passes primaires laissent attendre (poids appris lam)
  let aObs = rate('a'), pppObs = rate('ppp');
  if (B && hasMp && B.lam != null && B.lam < 1 && mp.a1 + mp.a2 > 0 && n) {
    const r21 = (B.r21 && B.r21[g]) || (lg && lg[g] && lg[g].r21) || 0.8;
    aObs = Math.max(0, (+p.a || 0) - (1 - B.lam) * (mp.a2 - mp.a1 * r21) * n / mp.gp) / n;
  }
  // … et corrigées de la chance de l'équipe sur la glace : la moitié de l'écart buts réels / attendus, bornée à ±25 %,
  // avec au moins 6 buts attendus sur la glace (sur quelques matchs, ce ratio est très bruité).
  if (hasMp && mp.ong > 0 && mp.onxg >= 6) {
    const f = Math.max(0.8, Math.min(1.25, 1 + 0.5 * (mp.onxg / mp.ong - 1)));
    aObs *= f; pppObs *= f;
  }
  // Tirs : en partie à partir des tentatives de tir × son % de tirs cadrés habituel (part apprise rho)
  let sogObs = rate('sog');
  if (B && B.rho && hasMp && mp.att > 0) {
    const sa = (pj && pj.sa) || (lg && lg[g] && lg[g].sa) || 0.52;
    sogObs = (1 - B.rho) * rate('sog') + B.rho * mp.att / mp.gp * sa;
  }
  // Pts AN : la base d'avant-saison suit le rôle en AN de cette saison vs la saison passée (poids appris ppb)
  let ppRole = null;
  if (B && B.ppb && hasMp && pj && n >= 3) {
    ppRole = B.ppMode === 'toi' && pj.ppt != null ? (mp.pptoi / mp.gp + 20) / (pj.ppt + 20)
           : pj.ppa != null ? (mp.ppatt / mp.gp + 0.2) / (pj.ppa + 0.2) : null;
    if (ppRole != null) pr.ppp *= Math.max(0.6, Math.min(1.6, Math.pow(ppRole, B.ppb)));
  }
  const proj = {
    g: blend_(pr.g, gObs, n, kG), a: blend_(pr.a, aObs, n, K.a), ppp: blend_(pr.ppp, pppObs, n, K.ppp),
    shp: blend_(pr.shp, rate('shp'), n, K.shp), sog: blend_(pr.sog, sogObs, n, K.sog),
    hit: blend_(pr.hit, rate('hit'), n, K.hit), blk: blend_(pr.blk, rate('blk'), n, K.blk)
  };
  // compagnons de trio (attaquants) : nouveau trio meilleur ou pire que d'habitude → attaque ±10 % au plus
  const lq = g === 'F' && p && LINEQ && LINEQ[p.id] != null ? LINEQ[p.id] : 0;
  if (lq) ['g', 'a', 'ppp', 'shp', 'sog'].forEach(k => { proj[k] *= 1 + 0.1 * lq; });
  // correction selon la tendance (FP/match), répartie sur toutes les stats
  const bias = cls && FN && FN.bias && FN.bias[cls] ? FN.bias[cls] : 0;
  if (inj && FN.inj.b) { const f0 = fp_(proj), f = f0 > 0 ? Math.max(0, f0 + FN.inj.b) / f0 : 1; Object.keys(proj).forEach(k => { proj[k] *= f; }); }
  if (bias) { const f0 = fp_(proj), f = f0 > 0 ? Math.max(0, f0 + bias) / f0 : 1; Object.keys(proj).forEach(k => { proj[k] *= f; }); }
  // FP « mérités » cette saison (mêmes stats réelles, buts et passes corrigés de la chance)
  // (la chance compare toujours aux buts attendus MoneyPuck quand on les a)
  const deserved = n ? fp_({ g: xgObs != null ? xgObs : gObs, a: aObs, ppp: pppObs, shp: rate('shp'), sog: sogObs, hit: rate('hit'), blk: rate('blk') }) : null;
  const actual = n ? (p.fpg || 0) : null;
  return {
    model: fp_(proj), deserved: deserved, luck: n && hasMp ? actual - deserved : null,
    toi: hasMp && mp.gp ? mp.toi / mp.gp / 60 : null, pptoi: hasMp && mp.gp ? mp.pptoi / mp.gp / 60 : null,
    ixg: hasMp && mp.gp ? mp.ixg / mp.gp : null, ppRole: ppRole, rates: proj, trend: cls, inj: inj || undefined, lq: lq || 0
  };
}
// Même règle que le backtest : 3 derniers matchs (feuilles de présences) vs le reste de sa saison (MoneyPuck)
function playingInjured_(p, mp) {
  const r = ROLES && p && p.id != null ? ROLES[p.id] || ROLES[String(p.id)] : null, n = mp && mp.gp || 0;
  if (!r || r.gp < 3 || n < 8 || !mp.toi) return false;
  const before = (mp.toi - 3 * r.toi) / (n - 3);
  return before > 0 && r.toi <= 0.8 * before;
}
// Tendance d'un patineur (mêmes règles que le backtest et l'onglet Tendances) : son jeu (temps de jeu, tirs, buts attendus)
// vs sa saison passée (et les analystes pour les tirs), ses FP vs les analystes, sa chance (buts − buts attendus)
function trendClass_(p, G, ly, mp) {
  const n = p ? p.gp || 0 : 0;
  if (n < 5 || !G || !G.g && !G.sog) return null;
  const parts = [], add = (now, base, w) => { if (now != null && base > 0) parts.push({ r: now / base, w: w }); };
  const hasMp = mp && mp.gp > 0;
  if (ly && hasMp && ly.toi) add(mp.toi / mp.gp / 60, ly.toi, 0.25);
  add((+p.sog || 0) / n, ly && ly.sog && G.sog ? (ly.sog + G.sog) / 2 : ly && ly.sog ? ly.sog : G.sog, 0.3);
  if (ly && ly.xg > 0 && hasMp) add(mp.ixg / mp.gp, ly.xg, 0.3);
  if (!parts.length) return null;
  const ws = parts.reduce((t, q) => t + q.w, 0);
  const proc = Math.exp(parts.reduce((t, q) => t + q.w * Math.log(Math.max(0.2, Math.min(5, q.r))), 0) / ws) - 1;
  const fpG = fp_({ g: G.g || 0, a: G.a || 0, ppp: G.ppp || 0, shp: G.shp || 0, sog: G.sog || 0, hit: G.hit || 0, blk: G.blk || 0 });
  const perf = fpG > 0 ? (p.fpg || 0) / fpG - 1 : 0;
  const luck = hasMp ? CONFIG.SCORING.G * ((+p.g || 0) - mp.ixg * n / mp.gp) / n : 0;
  const z = Math.max(1, Math.min(1.6, 1.6 - (n - 5) * 0.06));
  if (perf >= 0.12 * z) return proc >= 0.08 * z ? 'progression' : 'chanceux';
  if (perf <= -0.12 * z) return proc <= -0.08 * z ? 'panne' : 'malchanceux';
  if (proc <= -0.15 * z && luck >= 0.4 * z) return 'panne';
  if (proc >= 0.15 * z && luck <= -0.4 * z) return 'progression';
  if (luck >= 0.7 * z) return 'chanceux';
  if (luck <= -0.7 * z) return 'malchanceux';
  return 'conforme';
}
// Gardien : FP par match (base d'avant-saison, mélangée à la saison) × part des départs de son équipe
// base = GAF (corrigé) combiné au moteur maison (FP/match + GSAx sur 3 saisons) ; k et biais appris par le backtest
function modelGoalie_(p, prior, mp, name, engFpg) {
  const n = p ? p.gp || 0 : 0, B = inSeason_();
  const gafF = prior && prior.fpg ? prior.fpg : null;
  let base, k;
  if (B && B.gk && (gafF != null || engFpg != null)) {
    if (PJ_GK === undefined) { PJ_GK = null; try { PJ_GK = typeof pjParams_ === 'function' ? pjParams_().gk || null : null; } catch (e) { /* rien */ } }
    const GK = PJ_GK || { gafW: 0.5, gafB: APP.GOALIE_BIAS };
    base = (gafF != null && engFpg != null ? GK.gafW * (gafF + GK.gafB) + (1 - GK.gafW) * engFpg
          : gafF != null ? gafF + GK.gafB : engFpg) + B.gk.b;
    k = B.gk.k;
  } else { base = (gafF != null ? gafF : 3.6) + APP.GOALIE_BIAS; k = APP.K.gstart; }
  const perStart = blend_(base, n ? p.fpg || 0 : 0, n, k);
  const tg = p ? TEAM_GP[p.team] || 0 : 0;
  const fixed = APP.GOALIE_ROLE[name];
  let share = typeof fixed === 'number' ? fixed : (tg >= 3 && n ? Math.min(1, n / tg) : CONFIG.GOALIE_START_SHARE);
  // Peu de matchs (retour de blessure, échange, rappel) : « matchs joués ÷ matchs de l'équipe » le prend pour un auxiliaire.
  // Sous 5 matchs, on se fie surtout à son rôle prévu (matchs projetés par les analystes ÷ matchs de la saison).
  if (typeof fixed !== 'number' && n < 5 && prior && prior.gp > 0) {
    const season = typeof pjGames_ === 'function' ? pjGames_(String(CONFIG.SEASON)) : 82;
    const sp = Math.min(0.85, prior.gp / season), w = n / 5;
    share = w * share + (1 - w) * sp;
  }
  // part des 10 derniers matchs de son équipe, mélangée à la part de la saison (seulement si le backtest l'a retenue)
  const wR = B && B.gkRecent, rec = p && GK_RECENT[p.id];
  if (wR && rec != null && typeof fixed !== 'number' && n >= 5) share = wR * rec + (1 - wR) * share;
  // Chance : buts accordés vs attendus (MoneyPuck). Positif = il a accordé MOINS que prévu → chanceux… ou très bon
  const gsax = mp && mp.gp ? (mp.gxg - mp.gga) / mp.gp : null;
  return { model: perStart, share: share, shareFixed: typeof fixed === 'number', gsax: gsax, luck: null, deserved: null, base: base, k: k };
}

function signal_(m, gp) {
  if (m.luck === null || m.luck === undefined || gp < 5) return '';
  if (m.luck <= -APP.LUCK) return 'malchanceux';
  if (m.luck >= APP.LUCK) return 'chanceux';
  return 'normal';
}

// ---------------- Calendrier ----------------
// Matchs restants cette semaine (lundi→dimanche, à partir d'aujourd'hui, matchs pas encore commencés)
// et matchs de la semaine prochaine, par équipe LNH.
// ---------- Le « jour fantasy » : celui d'ESPN, pas celui de l'horloge ----------
// Après minuit, ESPN reste sur la journée d'hier (la semaine du dimanche n'est pas finie chez lui) jusqu'à sa bascule
// de la nuit. On suit donc le jour de pointage d'ESPN (1 par jour) : repère connu id 27 = 2026-10-25, mis à jour ensuite.
let FNOW = null;
function fantasyNow_() {
  if (FNOW) return FNOW;
  const now = new Date(), real = ymd_(now);
  let fd = real;
  try {
    const id = ESPN_DATA && (ESPN_DATA.scoringPeriodId || (ESPN_DATA.status || {}).latestScoringPeriod);
    const props = PropertiesService.getScriptProperties();
    let map = null; try { map = JSON.parse(props.getProperty('ESPN_DAY') || 'null'); } catch (e) { map = null; }
    if (!map || !map.id) map = { id: 27, date: '2026-10-25' };
    if (id) {
      fd = addDays_(map.date, id - map.id);
      if (fd > real || fd < addDays_(real, -2)) fd = real;           // repère incohérent : on garde l'horloge
      else if (id !== map.id) props.setProperty('ESPN_DAY', JSON.stringify({ id: id, date: fd }));
    } else if (+Utilities.formatDate(now, 'America/Toronto', 'H') < 5) fd = addDays_(real, -1);   // ESPN muet : avant 5 h = hier
  } catch (e) { fd = real; }
  FNOW = fd === real ? now : new Date(fd + 'T16:00:00Z');            // midi à Toronto ce jour-là
  return FNOW;
}
function weekSchedule_() {
  const now = fantasyNow_();
  const dow = +Utilities.formatDate(now, 'America/Toronto', 'u');   // 1 = lundi … 7 = dimanche
  const today = ymd_(now);
  const sunday = ymd_(new Date(now.getTime() + (7 - dow) * 86400000));
  const nextMon = ymd_(new Date(now.getTime() + (8 - dow) * 86400000));
  const nextSun = ymd_(new Date(now.getTime() + (14 - dow) * 86400000));
  const j = fetchJsonAll_([`${WEB}/schedule/${today}`, `${WEB}/schedule/${nextMon}`]);
  const left = {}, next = {}, todayTeams = {}, liveTeams = {}, seen = {}, weekGames = [];
  const add = (m, t) => { m[t] = (m[t] || 0) + 1; };
  // Les deux calendriers (7 jours chacun, à partir d'aujourd'hui et du lundi suivant) se chevauchent :
  // on ne compte chaque match qu'une fois.
  j.forEach(w => (w.gameWeek || []).forEach(day => (day.games || []).forEach(g => {
    if (g.gameType !== CONFIG.GAME_TYPE || seen[g.id]) return;
    seen[g.id] = true;
    const teams = [g.awayTeam.abbrev, g.homeTeam.abbrev];
    if (day.date >= today && day.date <= sunday) weekGames.push({ id: g.id, date: day.date, a: g.awayTeam.abbrev, h: g.homeTeam.abbrev });
    if (day.date === today && (g.gameState === 'LIVE' || g.gameState === 'CRIT')) teams.forEach(t => { liveTeams[t] = true; });   // match en cours
    if (day.date >= today && day.date <= sunday && (g.gameState === 'FUT' || g.gameState === 'PRE')) {
      teams.forEach(t => add(left, t));
      if (day.date === today) teams.forEach(t => { todayTeams[t] = true; });
    }
    if (day.date >= nextMon && day.date <= nextSun) teams.forEach(t => add(next, t));
  })));
  // Toutes les semaines (lundi → dimanche) jusqu'à la fin de la saison : matchs par équipe (fiche de joueur, « par semaine »)
  let weeks = [], games = {};
  // Calendrier de la saison (38 requêtes) : gardé 1 h, par semaine (clé = lundi). Les matchs de CE soir viennent d'en haut, toujours frais.
  const monK = ymd_(new Date(now.getTime() - (dow - 1) * 86400000)), cacheW = CacheService.getScriptCache(), hitW = cacheW.get('WEEKS_' + monK);
  if (hitW) { try { const o = JSON.parse(hitW); weeks = o.w; games = o.g; } catch (e) { weeks = []; games = {}; } }
  if (!weeks.length) try {
    // lundi de cette semaine ; on remonte 8 semaines pour couvrir le début de saison (les semaines sans match de saison régulière sont ignorées)
    const mon = monK, urls = [];
    for (let k = -8; k < 30; k++) urls.push(`${WEB}/schedule/${ymd_(new Date(new Date(mon + 'T12:00:00Z').getTime() + k * 7 * 86400000))}`);
    fetchJsonAll_(urls).forEach((w, k) => {
      const start = urls[k].split('/schedule/')[1], end = ymd_(new Date(new Date(start + 'T12:00:00Z').getTime() + 6 * 86400000)), n = {};
      let any = false;
      (w.gameWeek || []).forEach(day => {
        if (day.date < start || day.date > end) return;
        (day.games || []).forEach(g => {
          if (g.gameType !== CONFIG.GAME_TYPE) return;
          any = true; [g.awayTeam.abbrev, g.homeTeam.abbrev].forEach(t => add(n, t));
          // calendrier de chaque équipe (fiche de joueur) : [date, adversaire, 1 = à domicile]
          const a = g.awayTeam.abbrev, h = g.homeTeam.abbrev;
          (games[a] = games[a] || []).push([day.date, h, 0]); (games[h] = games[h] || []).push([day.date, a, 1]);
        });
      });
      if (any) weeks.push({ s: start, e: end, n: n });
    });
    try { const js = JSON.stringify({ w: weeks, g: games }); if (weeks.length && js.length < 95000) cacheW.put('WEEKS_' + monK, js, 3600); } catch (e) { /* trop gros : pas de cache */ }
  } catch (e) { WARN.push('Calendrier des semaines : ' + e.message); }
  // matchs en cours : part du match qui reste à jouer (période et horloge), pour compter la fin du match dans les projections
  const liveLeft = {};
  if (Object.keys(liveTeams).length) try {
    const sc = fetchJsonAll_([`${WEB}/score/${today}`])[0] || {};
    (sc.games || []).forEach(g => {
      if (g.gameState !== 'LIVE' && g.gameState !== 'CRIT') return;
      const per = (g.periodDescriptor && g.periodDescriptor.number) || 1, sec = g.clock && g.clock.secondsRemaining != null ? g.clock.secondsRemaining : 600;
      const inter = g.clock && g.clock.inIntermission;
      const elapsed = Math.min(3600, (per - 1) * 1200 + (inter ? 1200 : 1200 - sec));
      const f = per > 3 ? 0 : Math.max(0, Math.round((1 - elapsed / 3600) * 100) / 100);
      [g.awayTeam && g.awayTeam.abbrev, g.homeTeam && g.homeTeam.abbrev].forEach(t => { if (t) liveLeft[t] = f; });
    });
  } catch (e) { /* facultatif */ }
  return { left: left, next: next, today: todayTeams, live: liveTeams, liveLeft: liveLeft, weekGames: weekGames, weekEnd: sunday, weeks: weeks, games: games };
}

// ---------------- Simulation du reste de la saison ----------------
// avail (facultatif) : part de sa valeur selon son statut. Sans avail : un blessé ne compte pas (cette semaine).
// Pour le reste de la saison, on suppose qu'il revient : même règle que l'app (« long terme ») — LBP/absent 60 %, suspendu 50 %, incertain 95 %.
const AVAIL_LONG = { INJURY_RESERVE: 0.6, IR: 0.6, OUT: 0.6, SUSPENSION: 0.5, DAY_TO_DAY: 0.95 };
function lineupPower_(players, val0, avail) {
  const L = CONFIG.LINEUP;
  const val = avail ? p => val0(p) * (avail[p.status] == null ? 1 : avail[p.status]) : val0;
  const ok = avail ? players : players.filter(p => !OUT_STATUSES[p.status]);
  const by = (a, b) => val(b) - val(a);
  const F = ok.filter(p => p.grp === 'F').sort(by), D = ok.filter(p => p.grp === 'D').sort(by);
  const util = F.slice(L.F).concat(D.slice(L.D)).sort(by).slice(0, L.UTIL);
  const sk = F.slice(0, L.F).concat(D.slice(0, L.D), util).reduce((s, p) => s + val(p), 0) * CONFIG.GAMES_PER_WEEK * APP.INJURY_FACTOR;
  const g = ok.filter(p => p.grp === 'G').map(p => val(p) * (p.share || CONFIG.GOALIE_START_SHARE) * CONFIG.GAMES_PER_WEEK)
    .sort((a, b) => b - a).slice(0, L.G).reduce((s, x) => s + x, 0);
  return sk + g;
}
function gauss_() { let u = 0, v = 0; while (!u) u = Math.random(); while (!v) v = Math.random(); return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v); }

// Matchs restants du calendrier ESPN : [{period, a: sheetName, b: sheetName, liveA, liveB, frac}]
function remainingSchedule_() {
  const j = ESPN_DATA;
  if (!j || !j.schedule) return { games: [], playoffTeams: 0 };
  const cur = (j.status && j.status.currentMatchupPeriod) || 0;
  const name = id => ESPN_SHEET_OF[id] || '';
  const dow = +Utilities.formatDate(fantasyNow_(), 'America/Toronto', 'u');
  const games = [];
  j.schedule.forEach(m => {
    if (!m.home || !m.away || m.matchupPeriodId < cur) return;
    if (m.matchupPeriodId === cur && m.winner && m.winner !== 'UNDECIDED') return;
    const pts = s => typeof s.totalPointsLive === 'number' ? s.totalPointsLive : (s.totalPoints || 0);
    const a = name(m.home.teamId), b = name(m.away.teamId);
    if (!a || !b) return;
    const isCur = m.matchupPeriodId === cur;
    games.push({ period: m.matchupPeriodId, a: a, b: b, liveA: isCur ? pts(m.home) : 0, liveB: isCur ? pts(m.away) : 0,
                 frac: isCur ? Math.max(0, (8 - dow) / 7) : 1, cur: isCur });
  });
  const set = (j.settings && j.settings.scheduleSettings) || {};
  return { games: games, playoffTeams: set.playoffTeamCount || 8, current: cur };
}

function simulate_(teamsOut, sched) {
  const names = teamsOut.map(t => t.team);
  const idx = {}; names.forEach((n, i) => { idx[n] = i; });
  const N = APP.SIMS, T = names.length;
  const po = new Array(T).fill(0), first = new Array(T).fill(0), winsSum = new Array(T).fill(0), rankSum = new Array(T).fill(0);
  for (let s = 0; s < N; s++) {
    const w = teamsOut.map(t => t.record ? t.record.w + t.record.tie / 2 : 0);
    const pf = teamsOut.map(t => t.record ? t.record.pf : 0);
    sched.games.forEach(g => {
      const ia = idx[g.a], ib = idx[g.b];
      if (ia === undefined || ib === undefined) return;
      const pw = t => g.cur && t.powerNow != null ? t.powerNow : (t.powerBy && t.powerBy[g.period] != null ? t.powerBy[g.period] : t.power);
      const pa = pw(teamsOut[ia]), pb = pw(teamsOut[ib]);
      const sa = g.liveA + pa * g.frac + gauss_() * APP.WEEK_SD * Math.sqrt(g.frac);
      const sb = g.liveB + pb * g.frac + gauss_() * APP.WEEK_SD * Math.sqrt(g.frac);
      if (sa > sb) w[ia]++; else w[ib]++;
      pf[ia] += sa; pf[ib] += sb;
    });
    const order = names.map((n, i) => i).sort((x, y) => (w[y] - w[x]) || (pf[y] - pf[x]));
    order.forEach((i, r) => {
      rankSum[i] += r + 1; winsSum[i] += w[i];
      if (r < sched.playoffTeams) po[i]++;
      if (r === 0) first[i]++;
    });
  }
  teamsOut.forEach((t, i) => {
    t.sim = { playoffs: po[i] / N, first: first[i] / N, wins: winsSum[i] / N, rank: rankSum[i] / N };
  });
}

// ---------------- Instantané pour l'app ----------------
function buildAppSnapshot_(ss, teams, players, index, aliases) {
  mark_('Équipes LNH actuelles');
  try { INJ = injuries_(); } catch (e) { INJ = null; WARN.push('Blessures (dates de retour) : ' + e.message); }
  try { CUR_TEAMS = currentTeams_(); } catch (e) { CUR_TEAMS = null; WARN.push('Équipes LNH actuelles : ' + e.message); }
  // présences sur la glace d'abord : le drapeau « joue blessé » (temps de jeu −20 % sur 3 matchs) en a besoin
  mark_('Présences sur la glace');
  try { ROLES = typeof iceRoles_ === 'function' ? iceRoles_() : null; } catch (e) { WARN.push('Présences sur la glace : ' + e.message); }
  try { LINEQ = lineQuality_(players); } catch (e) { LINEQ = {}; WARN.push('Compagnons de trio : ' + e.message); }
  let mp = {};
  setProgress_(76, 'Stats avancées MoneyPuck…');
  try { mp = moneyPuck_(ss); } catch (e) { WARN.push('Stats avancées MoneyPuck : ' + e.message); }
  mark_('Projections GAF');
  const gaf = gafStats_(ss);
  setProgress_(82, 'Calendrier et moteur de projection…');
  let sch = { left: {}, next: {}, today: {}, weekEnd: '', weeks: [] };
  try { sch = weekSchedule_(); } catch (e) { WARN.push('Calendrier : ' + e.message); }

  mark_('Gardiens récents + moteur');
  try { const Bq = inSeason_(); if (Bq && Bq.gkRecent) GK_RECENT = goalieRecentShares_(players, sch.games || {}); } catch (e) { WARN.push('Gardiens récents : ' + e.message); }
  // Projection maison d'avant-saison (moteur réglé sur 4 saisons), combinée à GAF
  let pjPri = {}, gafW = 0.5;
  try { if (typeof pjCurrentPriors_ === 'function') { pjPri = pjCurrentPriors_(); gafW = pjParams_().gafW; if (gafW == null) gafW = 0.5;
    const Bf = inSeason_(); if (Bf && Bf.final && Bf.final.w != null) gafW = Bf.final.w; } }   // poids réglé en cours de saison, si retenu
  catch (e) { WARN.push('Moteur de projection : ' + e.message); }
  const combine = (g, m) => {
    if (!g) return m ? Object.assign({ pos: '' }, m) : null;
    if (!m) return g;
    const o = { pos: g.pos, gp: g.gp };
    ['g', 'a', 'ppp', 'shp', 'sog', 'hit', 'blk', 'fpg'].forEach(k => { o[k] = gafW * (g[k] || 0) + (1 - gafW) * (m[k] || 0); });
    return o;
  };
  const keyOf = n => { const k = norm_(n); return aliases[k] || k; };
  const OVR = getOverrides_();
  const rostered = {};
  const enrich = (name, nhl, pos, status, p) => {
    const g = grp_(pos || (p && p.pos));
    const gafPrior = gaf[keyOf(name)] || gaf[norm_(name)] || null;
    const pjPrior = g !== 'G' && p ? pjPri[p.id] || null : null;
    const engG = g === 'G' && p && pjPri._gk && pjPri._gk[p.id] != null ? pjPri._gk[p.id] : null;
    const prior = g === 'G' ? gafPrior : combine(gafPrior, pjPrior);
    const m = g === 'G' ? modelGoalie_(p, prior, p ? mp[p.id] : null, name, engG)
                        : modelSkater_(p, prior, p ? mp[p.id] : null, g, pjPrior, pjPri._lg, gafPrior);
    const team = (CUR_TEAMS && ((p && CUR_TEAMS.id[p.id]) || CUR_TEAMS.nom[norm_(name)])) || (p && p.team) || nhl || '';   // équipe actuelle (échanges compris)
    const ov = OVR[name], st = ov && ov.out && ov.out >= ymd_(new Date()) ? 'OUT' : (status || '');
    const ij = INJ && (OUT_STATUSES[st] || st === 'DAY_TO_DAY') ? INJ[norm_(name)] || null : null;
    // retour estimé (ESPN) ; date dépassée mais toujours blessé → on suppose dans 3 jours
    // « Mon info » : total de points visé sur la saison (ex. McDavid 175) — sa projection pour le reste de la saison suit
    if (ov && ov.pts > 0 && g !== 'G' && p && m.rates) {
      const season = typeof pjGames_ === 'function' ? pjGames_(String(CONFIG.SEASON)) : 82;
      const left = Math.max(1, season - (TEAM_GP[team] || p.gp || 0)), cur = (+p.g || 0) + (+p.a || 0);
      const R = m.rates, pg = (R.g || 0) + (R.a || 0), want = Math.max(0, (ov.pts - cur) / left), S = CONFIG.SCORING;
      if (pg > 0) {
        const f = want / pg;
        m.model += (f - 1) * (S.G * (R.g || 0) + S.A * (R.a || 0) + S.PPP * (R.ppp || 0) + S.SHP * (R.shp || 0));
        ['g', 'a', 'ppp', 'shp'].forEach(k => { R[k] = (R[k] || 0) * f; });
      }
    }
    const today0 = ymd_(new Date()), ret = ij && ij.r ? (ij.r > today0 ? ij.r : addDays_(today0, 3)) : (ov && ov.out && ov.out >= today0 ? addDays_(ov.out, 1) : null);
    return {
      name: name, pos: pos || (p && p.pos) || '', grp: g, nhl: team, status: st,
      gp: p ? p.gp : 0, fpg: p && p.gp ? round_(p.fpg) : null, gaf: gafPrior ? round_(gafPrior.fpg) : null,
      pj: pjPrior ? round_(pjPrior.fpg) : engG != null ? round_(engG) : null,
      model: round_(m.model), share: m.share !== undefined ? round_(m.share) : null,
      shareFixed: !!m.shareFixed, inj: m.inj ? 1 : 0,
      perGame: round_(g === 'G' ? m.model * (m.share || CONFIG.GOALIE_START_SHARE) : m.model),   // FP attendus par match de son équipe
      deserved: m.deserved === null ? null : round_(m.deserved), luck: m.luck === null ? null : round_(m.luck),
      signal: g === 'G' ? '' : signal_(m, p ? p.gp : 0),
      toi: m.toi ? round_(m.toi) : null, pptoi: m.pptoi ? round_(m.pptoi) : null, gsax: m.gsax ? round_(m.gsax) : null,
      left: sch.left[team] || 0, next: sch.next[team] || 0, today: !!sch.today[team], live: !!(sch.live && sch.live[team]),
      liveLeft: sch.liveLeft && sch.liveLeft[team] != null ? sch.liveLeft[team] : null,
      id: p ? p.id : null,
      ret: ret, injT: ij ? ij.t : null, injC: ij ? ij.c : null, lq: m.lq || 0,
      x: playerSheet_(g, p, m, p ? mp[p.id] : null, gafPrior, pjPrior, team)
    };
  };

  mark_('Classement ESPN + modèle des joueurs');
  const st = espnStandings_();
  const sheetToId = {};
  Object.keys(ESPN_SHEET_OF).forEach(id => { sheetToId[ESPN_SHEET_OF[id]] = +id; });
  const teamsOut = teams.map(t => {
    const ps = t.players.map(x => {
      rostered[keyOf(x.name)] = true;
      const p = findPlayer_(x.name, x.nhl, x.pos, index, aliases);
      return enrich(x.name, x.nhl, x.pos, x.status, p);
    });
    const e = st ? st.byId[sheetToId[t.team]] : null;
    return {
      team: t.team, players: ps,
      record: e ? { w: e.w, l: e.l, tie: e.tie, pf: round_(e.pf), pa: round_(e.pa), apW: e.apW, apL: e.apL,
                    live: e.live === '' ? null : round_(e.live), opp: e.opp === '' ? null : (ESPN_SHEET_OF[e.opp] || ''),
                    oppLive: e.oppLive === '' ? null : round_(e.oppLive), wkW: e.wkW, wkL: e.wkL } : null,
      power: round_(lineupPower_(ps, p => p.model || 0, AVAIL_LONG)),        // remplacé par la moyenne semaine par semaine (powerByPeriod_)
      powerNow: round_(lineupPower_(ps, p => p.model || 0)),                  // cette semaine : les blessés ne jouent pas
      powerReal: round_(lineupPower_(ps, p => (p.gp ? p.fpg || 0 : 0)))
    };
  });

  // transactions de la ligue (14 jours) + sonde MoneyPuck (1 fois par jour)
  mark_('Transactions ESPN');
  try { TX_CACHE = espnTransactions_(); } catch (e) { TX_CACHE = []; WARN.push('Transactions ESPN : ' + e.message); }
  try { if (false) mpProbe_(); } catch (e) { /* sonde facultative */ }
  mark_('Stats de la semaine');
  try { WEEKST = weekStats_(sch.weekEnd); } catch (e) { WARN.push('Stats de la semaine : ' + e.message); }
  mark_('Force des équipes');
  try { TSTR = teamStrength_(); } catch (e) { WARN.push('Force des équipes : ' + e.message); }
  mark_('Duel ESPN jour par jour');
  try { WKROS = espnWeekRosters_(); } catch (e) { WARN.push('Duel ESPN jour par jour : ' + e.message); }
  mark_('Gardiens partants (DailyFaceoff)');
  try { MPG = mpGames_(sch.weekGames); } catch (e) { WARN.push('MoneyPuck (gardiens partants, chances de victoire) : ' + e.message); }
  try { STARTERS = dfoStarters_(); } catch (e) { WARN.push('Gardiens partants : ' + e.message); }
  mark_('Défense adverse');
  try { OPP_ADJ = oppAdjust_(); } catch (e) { WARN.push('Adversaires : ' + e.message); }
  mark_('Journal ESPN');
  try { JOURNAL = espnJournal_(); } catch (e) { WARN.push('Journal ESPN : ' + e.message); }
  setProgress_(90, 'Simulation de 2 000 saisons…');
  const sched = remainingSchedule_();
  try { powerByPeriod_(teamsOut, sched, sch.weekEnd); } catch (e) { WARN.push('Force par semaine : ' + e.message); }
  if (sched.games.length) simulate_(teamsOut, sched);

  // Joueurs libres : tous les joueurs LNH ayant joué cette saison, absents de tous les alignements
  // (l'onglet Ballotage en montre les meilleurs ; Chercher les couvre tous)
  mark_('Joueurs libres');
  const waivers = players.filter(p => p.gp >= 1 && !rostered[keyOf(p.name)])
    .map(p => enrich(p.name, p.team, p.pos, ESPN_STATUS[norm_(p.name)] || '', p))
    .sort((a, b) => b.perGame - a.perGame).slice(0, 900);

  let myMoves = [];
  try { myMoves = myMoves_(teamsOut, waivers); } catch (e) { WARN.push('Bilan de tes moves : ' + e.message); }

  const snap = {
    v: 1, updated: stamp_(), fday: ymd_(fantasyNow_()), names: ESPN_NAME_OF, season: CONFIG.SEASON, me: APP.MY_TEAM, weekEnd: sch.weekEnd, weeks: sch.weeks || [], teamGames: sch.games || {},
    playoffTeams: sched.playoffTeams || 8, currentPeriod: sched.current || null,
    config: { gpw: CONFIG.GAMES_PER_WEEK, lineup: CONFIG.LINEUP, sd: APP.WEEK_SD, injury: APP.INJURY_FACTOR,
              share: CONFIG.GOALIE_START_SHARE, sims: APP.SIMS, luck: APP.LUCK,
              k: inSeason_() ? Object.assign({}, APP.K, inSeason_().K, { gx: inSeason_().kxg, g: inSeason_().gsrc === 'tirs' ? inSeason_().K.g : inSeason_().kg, gstart: inSeason_().gk ? inSeason_().gk.k : APP.K.gstart }) : APP.K,
              bt: inSeason_() },
    espn: espnLeagueInfo_(), tx: TX_CACHE, txSeen: TX_SEEN, ovr: getOverrides_(), watch: getWatch_(), rookies: rookies_(), moves: myMoves, journal: JOURNAL, roles: ROLES, opp: OPP_ADJ, starters: STARTERS, mpg: MPG, wk: WEEKST, tstr: TSTR, wkr: WKROS,
    teams: teamsOut, waivers: waivers, schedule: sched.games, warn: WARN.slice(),
    sources: ['LNH (api-web.nhle.com)', 'ESPN Fantasy', 'GAF Aggregate Projections', 'MoneyPuck.com']
  };
  setProgress_(96, 'Enregistrement…');
  mark_('Écriture des données');
  const json = JSON.stringify(snap);
  const sh = hiddenSheet_(ss, APP.DATA_SHEET);
  const chunks = [];
  for (let i = 0; i < json.length; i += 40000) chunks.push([json.slice(i, i + 40000)]);
  sh.clearContents();
  sh.getRange(1, 1, chunks.length, 1).setValues(chunks);
}
// Fiche de joueur (onglets Équipes / Chercher / Ballotage) : stats de la saison, projection stat par stat, stats avancées
function playerSheet_(g, p, m, mp, gafPrior, pjPrior, team) {
  const r3 = x => (typeof x === 'number' && isFinite(x) ? Math.round(x * 1000) / 1000 : null);
  const season = typeof pjGames_ === 'function' ? pjGames_(String(CONFIG.SEASON)) : 82;
  const teamLeft = Math.max(0, season - (TEAM_GP[team] || (p ? p.gp : 0) || 0));
  const o = { tl: teamLeft };
  if (g === 'G') {
    o.s = p ? [p.w, p.ga, p.sv, p.so, p.otl].map(v => +v || 0) : null;
    o.base = r3(m.base); o.k = m.k;
    // Prévision stat par stat (par match joué) : GAF mélangé à la saison avec le même k que les FP,
    // puis ramenée au niveau réel de la ligue (2025-26) : les analystes donnent ~25,3 arrêts par match (réel 23,5),
    // 2,89 buts accordés (réel 2,74), 0,50 victoire (réel 0,47). On ne force PLUS cette ligne de stats à égaler les FP maison :
    // l'ancien forçage mettait l'écart sur les buts accordés et faisait tomber tous les % d'arrêts sous ,900.
    const GK = ['w', 'ga', 'sv', 'so', 'otl'], LG = [0.5, 2.9, 25.5, 0.06, 0.1], n = p ? p.gp || 0 : 0, S = CONFIG.SCORING;
    const gr = gafPrior && gafPrior.sv ? GK.map(k => gafPrior[k] || 0) : null;
    if (gr) o.gr = gr.map(r3);
    const pr = gr || LG, k = m.k || 100;
    const r = GK.map((key, i) => (k * pr[i] + n * (o.s ? o.s[i] / Math.max(1, n) : 0) * (n ? 1 : 0)) / (k + n));
    const CAL = [0.95, 0.96, 0.93, 0.85, 1];          // victoires, buts accordés, arrêts, blanchissages, défaites en prol.
    for (let i = 0; i < 5; i++) r[i] *= CAL[i];
    o.r = r.map(r3);
    if (mp && mp.gp) { o.xga = r3(mp.gxg); o.ga = mp.gga; o.mgp = mp.gp; }
    return o;
  }
  const K = ['g', 'a', 'ppp', 'shp', 'sog', 'hit', 'blk'];
  o.s = p ? K.map(k => +p[k] || 0) : null;
  o.r = m.rates ? K.map(k => r3(m.rates[k])) : null;
  if (gafPrior && gafPrior.g != null) o.gr = K.map(k => r3(gafPrior[k]));
  if (pjPrior && pjPrior.g != null) o.pr = K.map(k => r3(pjPrior[k]));
  if (pjPrior && pjPrior.ly) o.ly = pjPrior.ly;
  if (mp && mp.gp) o.adv = { gp: mp.gp, toi: Math.round(mp.toi), pptoi: Math.round(mp.pptoi), ixg: r3(mp.ixg), a1: mp.a1, a2: mp.a2,
                             att: mp.att, sogm: mp.sogm, ppatt: mp.ppatt, onxg: r3(mp.onxg), ong: mp.ong };
  return o;
}
// Fiche de joueur, vue « par semaine » : ses matchs de la saison, un par un (demandé par la page au besoin)
function playerGames(id, grp) {
  if (!id) return '[]';
  const cache = CacheService.getScriptCache(), key = 'PG_' + id;
  const hit = cache.get(key);
  if (hit) return hit;
  const exp = encodeURIComponent(`seasonId=${CONFIG.SEASON} and gameTypeId=${CONFIG.GAME_TYPE} and playerId=${id}`);
  const url = path => `https://api.nhle.com/stats/rest/en/${path}?isAggregate=false&isGame=true&start=0&limit=100&cayenneExp=${exp}`;
  const paths = grp === 'G' ? ['goalie/summary'] : ['skater/summary', 'skater/realtime'];
  const res = UrlFetchApp.fetchAll(paths.map(p => ({ url: url(p), muteHttpExceptions: true })))
    .map(r => (r.getResponseCode() === 200 ? JSON.parse(r.getContentText()).data || [] : []));
  let out;
  if (grp === 'G') {
    out = res[0].map(r => ({ d: String(r.gameDate || '').slice(0, 10), w: r.wins || 0, ga: r.goalsAgainst || 0, sv: r.saves || 0, so: r.shutouts || 0, otl: r.otLosses || 0 }));
  } else {
    const rt = {}; res[1].forEach(r => { rt[r.gameId] = r; });
    out = res[0].map(r => { const x = rt[r.gameId] || {};
      return { d: String(r.gameDate || '').slice(0, 10), g: r.goals || 0, a: r.assists || 0, ppp: r.ppPoints || 0, shp: r.shPoints || 0, sog: r.shots || 0,
               hit: x.hits || 0, blk: x.blockedShots || 0 }; });
  }
  const json = JSON.stringify(out.filter(g => g.d).sort((a, b) => (a.d < b.d ? -1 : 1)));
  try { cache.put(key, json, 1800); } catch (e) { /* rien */ }
  return json;
}
// Infos de ligue ESPN pour l'app : limite d'acquisitions par duel (et ton total), calendrier des séries
function espnLeagueInfo_() {
  const j = ESPN_DATA;
  if (!j) return null;
  const set = j.settings || {}, acq = set.acquisitionSettings || {}, sch = set.scheduleSettings || {};
  const cur = (j.status && j.status.currentMatchupPeriod) || null;
  let used = null, total = null;
  (j.teams || []).forEach(t => {
    if (ESPN_SHEET_OF[t.id] !== APP.MY_TEAM) return;
    const tc = t.transactionCounter || {};
    total = tc.acquisitions != null ? tc.acquisitions : null;
    const m = tc.matchupAcquisitionTotals || {};
    used = cur != null && m[cur] != null ? m[cur] : (cur != null ? 0 : null);
  });
  // périodes de duel → jours (scoring periods) : pour situer les semaines de séries
  const mp = sch.matchupPeriods || {};
  return { cur: cur, regCount: sch.matchupPeriodCount || null, poLen: sch.playoffMatchupPeriodLength || 1,
           poTeams: sch.playoffTeamCount || null, periods: Object.keys(mp).length ? mp : null,
           firstScoringDay: (j.status && j.status.firstScoringPeriod) || 1, curScoringDay: j.scoringPeriodId || (j.status && j.status.latestScoringPeriod) || null,
           // « matchupLimitPerScoringPeriod » : la limite est donnée par jour de duel (1 × 7 jours = 7 par semaine),
           // mais elles s'utilisent quand on veut dans la semaine (toutes le même jour si on veut)
           gMax: set.rosterSettings && set.rosterSettings.positionLimits && +set.rosterSettings.positionLimits['5'] > 0 ? +set.rosterSettings.positionLimits['5'] : null,
           acqLimit: acq.matchupAcquisitionLimit != null && acq.matchupAcquisitionLimit >= 0 ? acq.matchupAcquisitionLimit * (acq.matchupLimitPerScoringPeriod ? 7 : 1) : null,
           acqUsed: used, acqSeason: total, acqSeasonLimit: acq.acquisitionLimit != null && acq.acquisitionLimit >= 0 ? acq.acquisitionLimit : null };
}
// Part des départs de chaque gardien sur les 10 derniers matchs de son équipe (matchs un par un, gardés 6 h)
function goalieRecentShares_(players, teamGames) {
  const cache = CacheService.getScriptCache(), hit = cache.get('GK_RECENT');
  if (hit) return JSON.parse(hit);
  const gs = players.filter(p => p.pos === 'G' && p.gp >= 2), today = ymd_(new Date()), out = {};
  const exp = id => encodeURIComponent(`seasonId=${CONFIG.SEASON} and gameTypeId=${CONFIG.GAME_TYPE} and playerId=${id}`);
  const res = UrlFetchApp.fetchAll(gs.map(p => ({ url: `https://api.nhle.com/stats/rest/en/goalie/summary?isAggregate=false&isGame=true&start=0&limit=100&cayenneExp=${exp(p.id)}`, muteHttpExceptions: true })));
  res.forEach((r, i) => {
    if (r.getResponseCode() !== 200) return;
    const p = gs[i], rows = JSON.parse(r.getContentText()).data || [];
    const last10 = ((teamGames[p.team] || []).map(g => g[0]).filter(d => d < today)).slice(-10);
    if (last10.length < 5) return;
    const started = rows.filter(g => (g.gamesStarted || 0) > 0 && last10.indexOf(String(g.gameDate || '').slice(0, 10)) >= 0).length;
    out[p.id] = Math.round(started / last10.length * 100) / 100;
  });
  try { cache.put('GK_RECENT', JSON.stringify(out), 21600); } catch (e) { /* rien */ }
  return out;
}
// Journal de bord : ton alignement ESPN jour par jour (actifs, banc, points de chaque joueur), 28 derniers jours, gardé jusqu'au lendemain
function espnJournal_() {
  if (!ESPN_DATA || !CONFIG.ESPN_LEAGUE_ID) return null;
  const sh = hiddenSheet_(SpreadsheetApp.getActiveSpreadsheet(), 'APP Journal'), today = ymd_(new Date());
  if (sh.getLastRow() && String(sh.getRange(1, 1).getDisplayValue()) === today) {
    try { return JSON.parse(sh.getRange(1, 2, 1, Math.max(1, sh.getLastColumn() - 1)).getValues()[0].join('')); } catch (e) { /* on relit */ }
  }
  const cur = ESPN_DATA.scoringPeriodId || (ESPN_DATA.status && ESPN_DATA.status.latestScoringPeriod);
  const myId = Object.keys(ESPN_SHEET_OF).find(id => ESPN_SHEET_OF[id] === APP.MY_TEAM);
  if (!cur || !myId) return null;
  const base = `https://${ESPN_HOSTS[0]}/apis/v3/games/fhl/seasons/${CONFIG.ESPN_SEASON}/segments/0/leagues/${CONFIG.ESPN_LEAGUE_ID}?view=mBoxscore&view=mMatchupScore&scoringPeriodId=`;
  const ids = []; for (let d = cur - 1; d >= Math.max(1, cur - 28); d--) ids.push(d);
  const res = UrlFetchApp.fetchAll(ids.map(d => Object.assign({ url: base + d }, espnOptions_(true))));
  const days = [];
  res.forEach((r, i) => {
    if (r.getResponseCode() !== 200) return;
    let j; try { j = JSON.parse(r.getContentText()); } catch (e) { return; }
    const d = ids[i], date = ymd_(new Date(Date.now() - (cur - d) * 86400000));
    (j.schedule || []).forEach(m => ['home', 'away'].forEach(side => {
      const t = m[side]; if (!t || String(t.teamId) !== String(myId) || !t.rosterForCurrentScoringPeriod) return;
      const rows = (t.rosterForCurrentScoringPeriod.entries || []).map(e => {
        const pe = e.playerPoolEntry || {}, pl = pe.player || {}, el = pl.eligibleSlots || [];
        return [pl.fullName || '', Math.round((pe.appliedStatTotal || 0) * 10) / 10, e.lineupSlotId,
                el.indexOf(5) >= 0 ? 'G' : el.indexOf(4) >= 0 && !(el.indexOf(0) >= 0 || el.indexOf(1) >= 0 || el.indexOf(2) >= 0 || el.indexOf(3) >= 0) ? 'D' : 'F'];
      });
      if (rows.length && !days.find(x => x.d === date)) days.push({ d: date, rows: rows });
    }));
  });
  const out = { days: days.sort((a, b) => (a.d < b.d ? -1 : 1)), bench: [7], ir: [8] };
  const json = JSON.stringify(out), parts = []; for (let i = 0; i < json.length; i += 45000) parts.push(json.slice(i, i + 45000));
  sh.clearContents(); sh.getRange(1, 1, 1, parts.length + 1).setValues([["'" + today].concat(parts)]);
  return out;
}
// Stats de la semaine en cours (lundi → aujourd'hui) pour TOUS les joueurs : id LNH → [matchs, FP]. Pour le duel de la semaine.
function weekStats_(weekEnd) {
  if (!weekEnd) return null;
  const mon = ymd_(new Date(new Date(weekEnd + 'T12:00:00Z').getTime() - 6 * 86400000)), today = ymd_(new Date());
  const cache = CacheService.getScriptCache(), ck = 'WKST_' + mon, hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  const exp = encodeURIComponent(`gameTypeId=${CONFIG.GAME_TYPE} and gameDate>="${mon}" and gameDate<="${today} 23:59:59"`);
  const url = (path, start) => `https://api.nhle.com/stats/rest/en/${path}?isAggregate=true&isGame=true&start=${start}&limit=100&sort=${encodeURIComponent('[{"property":"playerId","direction":"ASC"}]')}&cayenneExp=${exp}`;
  const pages = (path, n) => { const L = []; for (let i = 0; i < n; i++) L.push({ url: url(path, i * 100), muteHttpExceptions: true, path: path }); return L; };
  const reqs = pages('skater/summary', 10).concat(pages('skater/realtime', 10), pages('goalie/summary', 2));
  const res = UrlFetchApp.fetchAll(reqs.map(r => ({ url: r.url, muteHttpExceptions: true })));
  const rows = { 'skater/summary': [], 'skater/realtime': [], 'goalie/summary': [] };
  res.forEach((r, i) => { if (r.getResponseCode() === 200) { try { rows[reqs[i].path] = rows[reqs[i].path].concat(JSON.parse(r.getContentText()).data || []); } catch (e) {} } });
  const S = CONFIG.SCORING, rt = {}, out = { from: mon, p: {} };
  rows['skater/realtime'].forEach(r => { rt[r.playerId] = r; });
  rows['skater/summary'].forEach(r => {
    const x = rt[r.playerId] || {};
    const fp = S.G * (r.goals || 0) + S.A * (r.assists || 0) + S.PPP * (r.ppPoints || 0) + S.SHP * (r.shPoints || 0) + S.SOG * (r.shots || 0) + S.HIT * (x.hits || 0) + S.BLK * (x.blockedShots || 0);
    out.p[r.playerId] = [r.gamesPlayed || 0, Math.round(fp * 10) / 10];
  });
  rows['goalie/summary'].forEach(r => {
    const fp = S.W * (r.wins || 0) + S.GA * (r.goalsAgainst || 0) + S.SV * (r.saves || 0) + S.SO * (r.shutouts || 0) + S.OTL * (r.otLosses || 0);
    out.p[r.playerId] = [r.gamesStarted != null ? r.gamesStarted : (r.gamesPlayed || 0), Math.round(fp * 10) / 10];
  });
  try { cache.put(ck, JSON.stringify(out), 600); } catch (e) {}
  return out;
}
// Force des équipes LNH (classement) : buts pour / contre par match → chances de victoire d'un match (gardiens)
function teamStrength_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('TSTR');
  if (hit) return JSON.parse(hit);
  const r = UrlFetchApp.fetch(`${WEB}/standings/now`, { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return null;
  const out = {};
  (JSON.parse(r.getContentText()).standings || []).forEach(t => {
    const ab = t.teamAbbrev && (t.teamAbbrev.default || t.teamAbbrev); if (!ab) return;
    out[ab] = { gp: t.gamesPlayed || 0, gf: t.goalFor || 0, ga: t.goalAgainst || 0 };
  });
  try { cache.put('TSTR', JSON.stringify(out), 3 * 3600); } catch (e) {}
  return out;
}
// Le duel de la semaine selon ESPN, jour par jour : pour chaque équipe fantasy, les points de chaque joueur
// PENDANT qu'il était dans l'alignement (ajouts et retraits compris) — actifs (comptés) vs banc.
// { per, days, cur, teams: { feuille: { idESPN: [nom, FP comptés, FP banc, jours actifs joués, jours banc joués, jours dans l'équipe] } } }
function espnWeekRosters_() {
  if (!ESPN_DATA || !CONFIG.ESPN_LEAGUE_ID) return null;
  const st = ESPN_DATA.status || {}, per = st.currentMatchupPeriod, mp = ((ESPN_DATA.settings || {}).scheduleSettings || {}).matchupPeriods || {};
  const cur = ESPN_DATA.scoringPeriodId || st.latestScoringPeriod;
  // matchupPeriods donne des semaines, pas des jours : les jours de la semaine = du lundi (heure de l'Est) jusqu'à aujourd'hui
  const dow = +Utilities.formatDate(fantasyNow_(), 'America/Toronto', 'u');   // 1 = lundi … 7 = dimanche (jour fantasy)
  const ids = []; for (let d = Math.max(1, cur - (dow - 1)); d <= cur; d++) ids.push(d);
  if (!per || !ids.length) return null;
  const cache = CacheService.getScriptCache(), ck = 'WKR_' + per + '_' + ids.join('-') + '_' + Object.keys(ESPN_SHEET_OF).length, hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  const base = `https://${ESPN_HOSTS[0]}/apis/v3/games/fhl/seasons/${CONFIG.ESPN_SEASON}/segments/0/leagues/${CONFIG.ESPN_LEAGUE_ID}?view=mBoxscore&view=mMatchupScore&scoringPeriodId=`;
  const res = UrlFetchApp.fetchAll(ids.map(d => Object.assign({ url: base + d }, espnOptions_(true))));
  const out = { per: per, days: ids, cur: cur, teams: {} };
  res.forEach((r, ri) => {
    if (r.getResponseCode() !== 200) return;
    let j; try { j = JSON.parse(r.getContentText()); } catch (e) { return; }
    const day = ids[ri];
    (j.schedule || []).forEach(m => {
      if (m.matchupPeriodId !== per) return;
      ['home', 'away'].forEach(side => {
        const t = m[side]; if (!t || !t.rosterForCurrentScoringPeriod) return;
        const sheet = ESPN_SHEET_OF[t.teamId]; if (!sheet) return;
        const T = out.teams[sheet] || (out.teams[sheet] = {});
        (t.rosterForCurrentScoringPeriod.entries || []).forEach(e => {
          const pe = e.playerPoolEntry || {}, pl = pe.player || {}, pts = Math.round((pe.appliedStatTotal || 0) * 10) / 10;
          if (e.lineupSlotId === 8) return;                                     // réserve des blessés : ne compte pas
          // [nom, FP comptés, FP au banc, jours actifs avec un match, jours au banc avec un match, tous les jours dans l'équipe]
          const x = T[e.playerId] || (T[e.playerId] = [pl.fullName || '', 0, 0, [], [], []]);
          const st = (pl.stats || []).find(z => z.scoringPeriodId === day && z.statSourceId === 0 && z.stats && Object.keys(z.stats).length);
          const played = !!st || pts !== 0;
          x[5].push(day);
          if (e.lineupSlotId === 7) { x[2] += pts; if (played) x[4].push(day); } else { x[1] += pts; if (played) x[3].push(day); }
        });
      });
    });
  });
  Object.keys(out.teams).forEach(k => Object.keys(out.teams[k]).forEach(id => { const x = out.teams[k][id]; x[1] = Math.round(x[1] * 10) / 10; x[2] = Math.round(x[2] * 10) / 10; }));
  try { cache.put(ck, JSON.stringify(out), 600); } catch (e) {}
  return out;
}
// Blessures LNH (ESPN) : type, date de retour estimée, note. { nomnormalisé: { r: 'AAAA-MM-JJ', t, s, c } } — gardé 1 h.
let INJ = null;
const addDays_ = (d, n) => ymd_(new Date(new Date(d + 'T12:00:00Z').getTime() + n * 86400000));
const dayDiff_ = (a, b) => Math.round((new Date(b + 'T12:00:00Z') - new Date(a + 'T12:00:00Z')) / 86400000);
function injuries_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('INJ');
  if (hit) return JSON.parse(hit);
  // site.web.api : site.api refuse les serveurs de Google (403)
  let r = UrlFetchApp.fetch('https://site.web.api.espn.com/apis/site/v2/sports/hockey/nhl/injuries', { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) r = UrlFetchApp.fetch('https://site.api.espn.com/apis/site/v2/sports/hockey/nhl/injuries', { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('ESPN ' + r.getResponseCode());
  const out = {};
  (JSON.parse(r.getContentText()).injuries || []).forEach(t => (t.injuries || []).forEach(i => {
    const a = i.athlete || {}, d = i.details || {}; if (!a.displayName) return;
    out[norm_(a.displayName)] = { r: d.returnDate ? String(d.returnDate).slice(0, 10) : '', t: d.type || '', s: (d.fantasyStatus && d.fantasyStatus.abbreviation) || i.status || '', c: String(i.shortComment || '').slice(0, 220) };
  }));
  try { cache.put('INJ', JSON.stringify(out), 3600); } catch (e) {}
  return out;
}
// Part d'une semaine (s → e) où un joueur est disponible, selon sa date de retour estimée
function availWeek_(p, s, e) {
  const out = OUT_STATUSES[p.status];
  if (!out && p.status !== 'DAY_TO_DAY') return 1;
  if (!p.ret) return out ? AVAIL_LONG[p.status] || 0.6 : 1;      // pas de date : règle générale
  if (p.ret <= s) return 1;
  if (p.ret > e) return 0;
  return (dayDiff_(p.ret, e) + 1) / 7;
}
// Force de chaque équipe semaine par semaine (période ESPN → FP attendus), blessés comptés selon leur retour
function powerByPeriod_(teamsOut, sched, weekEnd) {
  if (!weekEnd || !sched.current) return;
  const pers = {}; sched.games.forEach(g => { if (g.period > sched.current) pers[g.period] = 1; });
  teamsOut.forEach(t => {
    const by = {}; let sum = 0, n = 0;
    Object.keys(pers).forEach(k => {
      const per = +k, e = addDays_(weekEnd, 7 * (per - sched.current)), s = addDays_(e, -6);
      const wmap = new Map(t.players.map(p => [p, availWeek_(p, s, e)]));
      by[per] = round_(lineupPower_(t.players, p => (p.model || 0) * wmap.get(p), {}));
      sum += by[per]; n++;
    });
    t.powerBy = by;
    if (n) t.power = round_(sum / n);                                  // force moyenne du reste de la saison
  });
}
// MoneyPuck, match par match (d'ici dimanche) : chances de victoire de chaque équipe et probabilité de départ de chaque gardien.
// { 'AAAA-MM-JJ': { ÉQUIPE: { w: 0.48, g: { idLNH: 0.42, … } } } } — chaque match gardé 2 h.
function mpGames_(games) {
  if (!games || !games.length) return null;
  const cache = CacheService.getScriptCache(), out = {}, todo = [];
  games.forEach(g => { const hit = cache.get('MPG_' + g.id); if (hit) put_(out, g, JSON.parse(hit)); else todo.push(g); });
  if (todo.length) {
    const B = 'https://moneypuck.com/moneypuck/', reqs = [];
    todo.forEach(g => ['predictions/' + g.id + '.csv', 'start_predictions/' + g.id + 'A.csv', 'start_predictions/' + g.id + 'H.csv']
      .forEach(u => reqs.push({ url: B + u, muteHttpExceptions: true })));
    const res = UrlFetchApp.fetchAll(reqs);
    const csv = r => { if (r.getResponseCode() !== 200) return []; const L = r.getContentText().trim().split(/\r?\n/); if (L.length < 2 || !/,/.test(L[0])) return [];
      const H = L[0].split(','); return L.slice(1).map(l => { const c = l.split(','), o = {}; H.forEach((h, i) => { o[h] = c[i]; }); return o; }); };
    todo.forEach((g, i) => {
      const pr = csv(res[3 * i])[0], ga = csv(res[3 * i + 1]), gh = csv(res[3 * i + 2]);
      const hw = pr ? +pr.preGameMoneyPuckHomeWinPrediction || +pr.preGameHomeTeamWinOverallScore : NaN;
      const v = { hw: isFinite(hw) && hw > 0 ? Math.round(hw * 1000) / 1000 : null, ga: {}, gh: {} };
      ga.forEach(x => { if (x.goalie_id) v.ga[x.goalie_id] = Math.round(+x.final_start_probability * 1000) / 1000; });
      gh.forEach(x => { if (x.goalie_id) v.gh[x.goalie_id] = Math.round(+x.final_start_probability * 1000) / 1000; });
      if (v.hw != null || ga.length || gh.length) { try { cache.put('MPG_' + g.id, JSON.stringify(v), 7200); } catch (e) {} }
      put_(out, g, v);
    });
  }
  return out;
  function put_(o, g, v) {
    const d = o[g.date] || (o[g.date] = {});
    if (v.hw != null || Object.keys(v.ga).length) d[g.a] = { w: v.hw != null ? Math.round((1 - v.hw) * 1000) / 1000 : null, g: v.ga };
    if (v.hw != null || Object.keys(v.gh).length) d[g.h] = { w: v.hw, g: v.gh };
  }
}
// Équipe LNH actuelle de chaque joueur (alignements officiels) : un joueur échangé change d'équipe tout de suite,
// même s'il n'a pas encore joué pour sa nouvelle équipe. Gardé 1 h. { id: { idLNH: 'ABC' }, nom: { nomnormalisé: 'ABC' } }
const NHL_TEAMS = ['ANA','BOS','BUF','CGY','CAR','CHI','COL','CBJ','DAL','DET','EDM','FLA','LAK','MIN','MTL','NSH','NJD','NYI','NYR','OTT','PHI','PIT','SJS','SEA','STL','TBL','TOR','UTA','VAN','VGK','WSH','WPG'];
function currentTeams_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('CUR_TEAMS');
  if (hit) return JSON.parse(hit);
  const res = UrlFetchApp.fetchAll(NHL_TEAMS.map(t => ({ url: `${WEB}/roster/${t}/current`, muteHttpExceptions: true })));
  const out = { id: {}, nom: {} };
  let ok = 0;
  res.forEach((r, i) => {
    if (r.getResponseCode() !== 200) return;
    let j; try { j = JSON.parse(r.getContentText()); } catch (e) { return; }
    ok++;
    ['forwards', 'defensemen', 'goalies'].forEach(k => (j[k] || []).forEach(p => {
      out.id[p.id] = NHL_TEAMS[i];
      const n = `${(p.firstName || {}).default || ''} ${(p.lastName || {}).default || ''}`;
      out.nom[norm_(n)] = NHL_TEAMS[i];
    }));
  });
  if (ok < 25) return out;                       // lecture incomplète : on ne la garde pas en cache
  try { cache.put('CUR_TEAMS', JSON.stringify(out), 3600); } catch (e) {}
  return out;
}
// Gardiens partants (DailyFaceoff, page « starting goalies » : statut Confirmed / Likely / Unconfirmed), aujourd'hui et demain.
// Résultat : { 'aaaa-mm-jj': { ÉQUIPE: { n: nom, s: statut } } } — l'équipe vient du nom du gardien (cache LNH de la saison). Gardé 20 min.
function dfoStarters_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('DFO_STARTERS');
  if (hit) return JSON.parse(hit);
  const ss = SpreadsheetApp.getActiveSpreadsheet(), teamOf = {}, map = readCache_(ss);
  Object.keys(map).forEach(id => { const l = map[id]; if (l.pos === 'G' && l.name) teamOf[norm_(l.name)] = l.team; });
  const days = [ymd_(new Date()), ymd_(new Date(Date.now() + 86400000))], out = {};
  const UA = { 'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36', 'Accept': 'text/html' };
  const res = UrlFetchApp.fetchAll(days.map(d => ({ url: 'https://www.dailyfaceoff.com/starting-goalies/' + d, muteHttpExceptions: true, followRedirects: true, headers: UA })));
  res.forEach((r, i) => {
    if (r.getResponseCode() !== 200) return;
    const m = r.getContentText().match(/<script id="__NEXT_DATA__"[^>]*>([\s\S]*?)<\/script>/);
    if (!m) return;
    let j; try { j = JSON.parse(m[1]); } catch (e) { return; }
    const games = [], walk = o => { if (!o || typeof o !== 'object') return; if (o.homeGoalieName || o.awayGoalieName) { games.push(o); return; } Object.keys(o).forEach(k => walk(o[k])); };
    walk(j);
    const day = out[days[i]] = {};
    games.forEach(g => ['home', 'away'].forEach(side => {
      const nm = g[side + 'GoalieName'], st = g[side + 'NewsStrengthName'];
      if (!nm || !st) return;
      const tm = teamOf[norm_(nm)]; if (!tm) return;
      day[tm] = { n: nm, s: String(st) };
    }));
  });
  try { cache.put('DFO_STARTERS', JSON.stringify(out), 1200); } catch (e) {}
  return out;
}
// Adversaire du soir (retenu par le backtest par match) : défense adverse (buts et tirs accordés, ramenés vers la ligue avec 10 matchs)
// → multiplicateur de la partie offensive (buts, passes, pts AN, tirs) d'un patineur pour un match contre cette équipe.
function oppAdjust_() {
  const raw = PropertiesService.getDocumentProperties().getProperty('BT_MATCHUP');
  const B = raw ? JSON.parse(raw) : null;
  if (!B || !B.coef || !Object.keys(B.coef).length) return null;
  const cache = CacheService.getScriptCache(), hit = cache.get('OPP_ADJ2');
  if (hit) return JSON.parse(hit);
  const exp = encodeURIComponent(`seasonId=${CONFIG.SEASON} and gameTypeId=${CONFIG.GAME_TYPE}`);
  const res = UrlFetchApp.fetchAll([
    { url: `https://api.nhle.com/stats/rest/en/team/summary?isAggregate=false&isGame=false&start=0&limit=50&cayenneExp=${exp}`, muteHttpExceptions: true },
    { url: 'https://api.nhle.com/stats/rest/en/team', muteHttpExceptions: true }]);
  if (res[0].getResponseCode() !== 200 || res[1].getResponseCode() !== 200) return null;
  const rows = JSON.parse(res[0].getContentText()).data || [], tri = {};
  (JSON.parse(res[1].getContentText()).data || []).forEach(t => { tri[t.id] = t.triCode; });
  let gp = 0, ga = 0, sa = 0;
  rows.forEach(r => { gp += r.gamesPlayed || 0; ga += (r.goalsAgainstPerGame || 0) * (r.gamesPlayed || 0); sa += (r.shotsAgainstPerGame || 0) * (r.gamesPlayed || 0); });
  if (!gp) return null;
  const mGA = ga / gp, mSA = sa / gp, K = 10, sh = (x, n, m) => (x + m * K) / (n + K) / m - 1, out = { coef: B.coef, t: {} };
  rows.forEach(r => {
    const n = r.gamesPlayed || 0, ab = tri[r.teamId]; if (!ab || n < 5) return;
    const f = { ga: sh((r.goalsAgainstPerGame || 0) * n, n, mGA), sa: sh((r.shotsAgainstPerGame || 0) * n, n, mSA) };
    let m = 1; Object.keys(B.coef).forEach(k => { if (f[k] != null) m += B.coef[k] * f[k]; });
    out.t[ab] = Math.round(Math.max(0.8, Math.min(1.25, m)) * 1000) / 1000;
  });
  try { cache.put('OPP_ADJ2', JSON.stringify(out), 6 * 3600); } catch (e) {}
  return out;
}
// Présences sur la glace (feuilles de présences LNH, après chaque match) : qui est sur la 1re vague d'AN, avec qui il joue à 5 c. 5.
// 6 derniers matchs de chaque équipe (3 récents vs 3 d'avant → promotions et changements de duo). Résumés gardés dans « APP Glace ».
const ICE = { GAMES_PER_TEAM: 6, MAX_NEW: 40, SHEET: 'APP Glace' };
let ICE_STORE = {}, LINEQ = {};
// Compagnons de trio (retenu par le backtest par match 2025-26, coef 0,1) : qualité de ses 2 principaux coéquipiers à 5 c. 5
// au DERNIER match (FP d'attaque par match de la saison) vs la moyenne de ses coéquipiers sur toute la saison → −1 à +1
function lineQuality_(players) {
  const off = {};
  players.forEach(p => { if (p.pos !== 'G' && p.gp >= 5) off[p.id] = (3 * (+p.g || 0) + 2 * (+p.a || 0) + 0.5 * (+p.ppp || 0) + 0.5 * (+p.shp || 0) + 0.2 * (+p.sog || 0)) / p.gp; });
  const hist = {};
  Object.keys(ICE_STORE).map(Number).sort((a, b) => a - b).forEach(id => {
    const T = (ICE_STORE[id] || {}).teams || {};
    Object.keys(T).forEach(t => Object.keys(T[t].p || {}).forEach(pid => {
      const ms = (T[t].p[pid][3] || []).map(m => m[0]).filter(m => off[m] != null).slice(0, 2);
      if (ms.length === 2) (hist[pid] = hist[pid] || []).push((off[ms[0]] + off[ms[1]]) / 2);
    }));
  });
  const out = {};
  Object.keys(hist).forEach(pid => {
    const h = hist[pid]; if (h.length < 3) return;
    const avg = h.reduce((x, v) => x + v, 0) / h.length; if (avg <= 0) return;
    out[pid] = Math.round(Math.max(-1, Math.min(1, h[h.length - 1] / avg - 1)) * 100) / 100;
  });
  return out;
}
function iceRoles_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), games = readGames_(ss);
  if (!games.length) return null;
  const byTeam = {};
  games.forEach(g => [g.away, g.home].forEach(t => { (byTeam[t] = byTeam[t] || []).push(+g.id); }));
  const want = {};
  Object.keys(byTeam).forEach(t => byTeam[t].sort((a, b) => b - a).forEach(id => { want[id] = true; }));   // toute la saison (compagnons de trio)
  const sh = hiddenSheet_(ss, ICE.SHEET), store = {};
  if (sh.getLastRow()) sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(r => { if (r[0]) { try { store[+r[0]] = JSON.parse(r[1]); } catch (e) {} } });
  const todo = Object.keys(want).map(Number).filter(id => !store[id]).sort((a, b) => b - a).slice(0, ICE.MAX_NEW);
  for (let k = 0; k < todo.length; k += 10) {
    const batch = todo.slice(k, k + 10), urls = [];
    batch.forEach(id => { urls.push(`https://api.nhle.com/stats/rest/en/shiftcharts?cayenneExp=gameId=${id}`, `${WEB}/gamecenter/${id}/play-by-play`); });
    const res = UrlFetchApp.fetchAll(urls.map(u => ({ url: u, muteHttpExceptions: true })));
    batch.forEach((id, i) => {
      const a = res[2 * i], b = res[2 * i + 1];
      if (a.getResponseCode() !== 200 || b.getResponseCode() !== 200) return;
      try { store[id] = iceGame_(JSON.parse(a.getContentText()).data || [], JSON.parse(b.getContentText())); } catch (e) { /* match illisible : on réessaiera */ }
    });
  }
  // on ne garde que les matchs voulus
  const keep = Object.keys(want).filter(id => store[id]);
  sh.clearContents();
  if (keep.length) sh.getRange(1, 1, keep.length, 2).setValues(keep.map(id => [+id, JSON.stringify(store[id])]));
  ICE_STORE = {}; keep.forEach(id => { ICE_STORE[id] = store[id]; });
  // par équipe : 3 derniers matchs vs 3 d'avant
  const out = {}, names = {}, goalie = {};
  // matchs déjà gardés avant le correctif : on écarte les gardiens (positions du cache LNH de la saison)
  try { const cm = readCache_(ss); Object.keys(cm).forEach(id => { if (cm[id].pos === 'G') goalie[id] = true; }); } catch (e) {}
  Object.keys(byTeam).forEach(t => {
    const ids = byTeam[t].sort((a, b) => b - a).slice(0, ICE.GAMES_PER_TEAM).filter(id => store[id]);
    if (ids.length < 3) return;
    const agg = list => {
      const pl = {}; let teamPP = 0;
      list.forEach(id => {
        const G = store[id], side = G.teams[t]; if (!side) return;
        teamPP += side.pp;
        Object.keys(side.p).forEach(pid => {
          if (goalie[pid]) return;
          const r = side.p[pid], o = pl[pid] || (pl[pid] = { pp: 0, es: 0, toi: 0, gp: 0, m: {} });
          o.pp += r[1]; o.es += r[2]; o.toi += r[0]; o.gp++; names[pid] = r[4] || names[pid];
          (r[3] || []).forEach(([mid, sec]) => { if (!goalie[mid]) o.m[mid] = (o.m[mid] || 0) + sec; });
        });
      });
      const rank = Object.keys(pl).sort((x, y) => pl[y].pp - pl[x].pp);
      rank.forEach((pid, i) => { pl[pid].rank = pl[pid].pp > 0 ? i + 1 : 99; });
      return { pl: pl, teamPP: teamPP, n: list.length };
    };
    const now = agg(ids.slice(0, 3)), prev = agg(ids.slice(3, 6));
    Object.keys(now.pl).forEach(pid => {
      const a = now.pl[pid], b = prev.pl[pid];
      const mates = o => Object.keys(o.m).sort((x, y) => o.m[y] - o.m[x]).slice(0, 2);
      out[pid] = { pp: a.rank, ppPrev: b ? b.rank : null, ppm: Math.round(a.pp / Math.max(1, a.gp)), toi: Math.round(a.toi / Math.max(1, a.gp)),
                   share: now.teamPP ? Math.round(a.pp / now.teamPP * 100) / 100 : 0,
                   sharePrev: b && prev.teamPP ? Math.round(b.pp / prev.teamPP * 100) / 100 : (prev.n ? 0 : null),
                   mates: mates(a), matesPrev: b ? mates(b) : null, gp: a.gp };
    });
  });
  Object.keys(out).forEach(pid => { out[pid].mates = out[pid].mates.map(m => names[m] || m); if (out[pid].matesPrev) out[pid].matesPrev = out[pid].matesPrev.map(m => names[m] || m); });
  return out;
}
// Un match : par équipe, temps d'AN de l'équipe et, par joueur, [temps de jeu, temps en AN, temps à 5 c. 5, 3 meilleurs coéquipiers à 5 c. 5, nom]
function iceGame_(shifts, pbp) {
  const sec = t => { const m = String(t || '').split(':'); return m.length === 2 ? (+m[0]) * 60 + (+m[1]) : 0; };
  const awayId = pbp.awayTeam && pbp.awayTeam.id;
  // situation seconde par seconde (période × 1200 s ; prolongation comprise) à partir des codes de situation du déroulement
  const T = 1200 * 5, sit = new Array(T).fill('1551');
  const plays = (pbp.plays || []).filter(pl => pl.situationCode && pl.periodDescriptor && pl.periodDescriptor.periodType !== 'SO')
    .map(pl => ({ t: (pl.periodDescriptor.number - 1) * 1200 + sec(pl.timeInPeriod), c: String(pl.situationCode) })).sort((a, b) => a.t - b.t);
  for (let i = 0; i < plays.length; i++) { const end = i + 1 < plays.length ? plays[i + 1].t : T; for (let x = plays[i].t; x < Math.min(end, T); x++) sit[x] = plays[i].c; }
  const onIce = {}, info = {}, isG = {};
  (pbp.rosterSpots || []).forEach(r => { if (r.positionCode === 'G') isG[r.playerId] = true; });   // les gardiens ne sont pas des coéquipiers de trio
  shifts.filter(r => r.typeCode === 517 && r.period >= 1 && r.period <= 5 && !isG[r.playerId]).forEach(r => {
    const a = (r.period - 1) * 1200 + sec(r.startTime), b = (r.period - 1) * 1200 + sec(r.endTime);
    const tm = r.teamId; (onIce[tm] = onIce[tm] || []).push([r.playerId, a, Math.min(b, T)]);
    info[r.playerId] = (r.firstName || '') + ' ' + (r.lastName || '');
  });
  const out = { teams: {} };
  Object.keys(onIce).forEach(tm => {
    const isAway = +tm === awayId, perSec = Array.from({ length: T }, () => []);
    onIce[tm].forEach(([pid, a, b]) => { for (let x = a; x < b; x++) perSec[x].push(pid); });
    const P = {}; let teamPP = 0;
    const get = pid => P[pid] || (P[pid] = { toi: 0, pp: 0, es: 0, m: {} });
    for (let x = 0; x < T; x++) {
      const c = sit[x], ag = +c[0], as = +c[1], hs = +c[2], hg = +c[3];
      const my = isAway ? as : hs, opp = isAway ? hs : as, myG = isAway ? ag : hg, oppG = isAway ? hg : ag;
      const pp = my > opp && myG === 1 && oppG === 1, ev = my === 5 && opp === 5 && myG === 1 && oppG === 1;
      const L = perSec[x]; if (!L.length) continue;
      if (pp) teamPP++;
      L.forEach(pid => { const o = get(pid); o.toi++; if (pp) o.pp++; if (ev) o.es++; });
      if (ev) for (let i = 0; i < L.length; i++) for (let j = i + 1; j < L.length; j++) { get(L[i]).m[L[j]] = (get(L[i]).m[L[j]] || 0) + 1; get(L[j]).m[L[i]] = (get(L[j]).m[L[i]] || 0) + 1; }
    }
    const pl = {};
    Object.keys(P).forEach(pid => { const o = P[pid];
      pl[pid] = [o.toi, o.pp, o.es, Object.keys(o.m).sort((a, b) => o.m[b] - o.m[a]).slice(0, 3).map(m => [+m, o.m[m]]), info[pid] || ''];
    });
    out.teams[tm] = { pp: teamPP, p: pl };  // clé = id d'équipe LNH (comme « LNH Parties »)
  });
  return out;
}
function round_(x) { return typeof x === 'number' && isFinite(x) ? Math.round(x * 100) / 100 : null; }

// =====================================================================================
// LIGUE : transactions ESPN, corrections manuelles (« Mon info »),
// nouvelles des joueurs, sonde des gardiens partants MoneyPuck
// =====================================================================================
const LG = { TX_DAYS: 14, NAMES_PROP: 'ESPN_NAMES', OVR_PROP: 'MY_OVERRIDES' };

// Noms des joueurs ESPN (id → nom), gardés d'une mise à jour à l'autre pour nommer les joueurs libérés
function espnNames_() {
  const props = PropertiesService.getDocumentProperties();
  let map = {}; try { map = JSON.parse(props.getProperty(LG.NAMES_PROP) || '{}'); } catch (e) { map = {}; }
  ((ESPN_DATA && ESPN_DATA.teams) || []).forEach(t => (((t.roster || {}).entries) || []).forEach(e => {
    const p = e.playerPoolEntry && e.playerPoolEntry.player; if (p && p.id && p.fullName) map[p.id] = p.fullName;
  }));
  try { props.setProperty(LG.NAMES_PROP, JSON.stringify(map)); } catch (e) { /* trop gros : on garde en mémoire seulement */ }
  return map;
}
// Transactions des 14 derniers jours (échanges, ajouts, retraits) de TOUTE la ligue, de 2 sources :
//  1) ESPN (vue mTransactions2, jour par jour, avec un filtre de types — sans filtre, ESPN ne renvoie que tes transactions) ;
//     un échange accepté (TRADE_ACCEPT) ne contient souvent pas les joueurs : on les prend dans sa proposition liée.
//  2) le News Tracker du classeur (changements d'alignement détectés à chaque synchro ESPN) : rattrape ce qu'ESPN aurait raté.
function espnTransactions_() {
  const out = [], seen = {}, nk = n => norm_(n);
  const itemKey = i => i.kind + '|' + nk(i.name) + '|' + i.from + '|' + i.to;
  const push = t => {
    const fresh = t.items.filter(i => !seen[itemKey(i)]);
    if (!fresh.length) return;
    t.items.forEach(i => { seen[itemKey(i)] = true; });
    out.push(t);
  };
  try { espnTxApi_().forEach(push); } catch (e) { TX_SEEN['ESPN : erreur ' + e.message] = 1; }
  try { newsTrackerTx_().forEach(push); } catch (e) { TX_SEEN['News Tracker : erreur ' + e.message] = 1; }
  return out.sort((a, b) => (a.date < b.date ? 1 : -1));
}
function espnTxApi_() {
  if (!ESPN_DATA || !CONFIG.ESPN_LEAGUE_ID) return [];
  // un ajout fait après le début des matchs (ou une réclamation au ballotage) compte pour le jour SUIVANT : on lit aussi 2 jours d'avance
  const st = ESPN_DATA.status || {}, cur = Math.max(ESPN_DATA.scoringPeriodId || 0, st.latestScoringPeriod || 0, st.currentScoringPeriod || 0) || null;
  if (!cur) return [];
  const names = espnNames_(), team = id => ESPN_SHEET_OF[id] || ('Équipe ' + id);
  const base = `https://${ESPN_HOSTS[0]}/apis/v3/games/fhl/seasons/${CONFIG.ESPN_SEASON}/segments/0/leagues/${CONFIG.ESPN_LEAGUE_ID}?view=mTransactions2&scoringPeriodId=`;
  const filt = JSON.stringify({ transactions: { filterType: { value: ['FREEAGENT', 'WAIVER', 'WAIVER_ERROR', 'TRADE_ACCEPT', 'TRADE_UPHOLD', 'TRADE_PROPOSAL', 'ROSTER'] } } });
  const opt = () => { const o = espnOptions_(true); o.headers = Object.assign({}, o.headers, { 'x-fantasy-filter': filt }); return o; };
  const reqs = [];
  for (let d = cur + 2; d > Math.max(0, cur - LG.TX_DAYS); d--) reqs.push(Object.assign({ url: base + d }, opt()));
  const all = {};
  UrlFetchApp.fetchAll(reqs).forEach(r => {
    if (r.getResponseCode() !== 200) return;
    let j; try { j = JSON.parse(r.getContentText()); } catch (e) { return; }
    (j.transactions || []).forEach(t => { if (!all[t.id]) all[t.id] = t; });
  });
  // joueurs absents des alignements actuels (retirés depuis) : on demande leur nom à ESPN, une fois pour toutes
  const miss = {};
  Object.keys(all).forEach(id => (all[id].items || []).forEach(i => { if (i.playerId > 0 && !names[i.playerId]) miss[i.playerId] = 1; }));
  const missIds = Object.keys(miss).map(Number).slice(0, 200);
  if (missIds.length) {
    try {
      const r = UrlFetchApp.fetch(`https://${ESPN_HOSTS[0]}/apis/v3/games/fhl/seasons/${CONFIG.ESPN_SEASON}/players?view=players_wl`,
        { muteHttpExceptions: true, headers: { 'x-fantasy-filter': JSON.stringify({ filterIds: { value: missIds } }) } });
      if (r.getResponseCode() === 200) {
        [].concat(JSON.parse(r.getContentText()) || []).forEach(x => { const q = x.player || x; if (q && q.id && q.fullName) names[q.id] = q.fullName; });
        try { PropertiesService.getDocumentProperties().setProperty(LG.NAMES_PROP, JSON.stringify(names)); } catch (e) { /* rien */ }
      }
    } catch (e) { TX_SEEN['ESPN noms : ' + e.message] = 1; }
  }
  const ids = Object.keys(all), out = [];
  ids.forEach(id => { const t = all[id], k = String(t.type || '') + ' ' + String(t.status || ''); TX_SEEN[k] = (TX_SEEN[k] || 0) + 1; });
  const mapItems = list => (list || []).filter(i => i.type === 'TRADE' || i.type === 'ADD' || i.type === 'DROP')
    .map(i => ({ name: names[i.playerId] || ('Joueur #' + i.playerId), kind: i.type, from: i.fromTeamId > 0 ? team(i.fromTeamId) : '', to: i.toTeamId > 0 ? team(i.toTeamId) : '' }));
  const used = {};
  ids.forEach(id => {
    const t = all[id], type = String(t.type || ''), status = String(t.status || '');
    if (status !== 'EXECUTED' || /DECLINE|VETO|ERROR/.test(type)) return;
    if (!/TRADE_ACCEPT|TRADE_UPHOLD|TRADE_PROPOSAL|FREEAGENT|WAIVER|ROSTER/.test(type)) return;
    let items = mapItems(t.items), src = t;
    if (/TRADE/.test(type) && !items.some(i => i.kind === 'TRADE')) {
      // les joueurs sont dans la proposition liée (dans un sens ou dans l'autre)
      const rel = all[t.relatedTransactionId] || ids.map(x => all[x]).find(x => x.relatedTransactionId === t.id && (x.items || []).length);
      if (rel) { items = mapItems(rel.items); src = rel; }
    }
    if (!items.length) return;
    const sig = items.map(i => i.kind + i.name + i.from + i.to).sort().join('|');
    if (used[sig]) return; used[sig] = true;
    const isTrade = /TRADE/.test(type) || items.some(i => i.kind === 'TRADE');
    if (isTrade) TX_SEEN['→ échanges trouvés'] = (TX_SEEN['→ échanges trouvés'] || 0) + 1;
    out.push({ id: String(t.id), date: Utilities.formatDate(new Date(t.processDate || src.processDate || t.proposedDate || Date.now()), 'America/Toronto', 'yyyy-MM-dd HH:mm'),
               trade: isTrade, team: isTrade ? '' : (t.teamId > 0 ? team(t.teamId) : ''), items: items });
  });
  return out;
}
// Lignes « Transactions récentes » du News Tracker (A:D) : Date, Type, Équipe(s), Détail
function newsTrackerTx_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NEWS_SHEET);
  if (!sh || sh.getLastRow() < 5) return [];
  const since = ymd_(new Date(Date.now() - LG.TX_DAYS * 86400000)), out = [];
  const list = s => String(s || '').split(',').map(x => x.trim()).filter(Boolean);
  sh.getRange(5, 1, Math.min(CONFIG.NEWS_MAX_LOG || 200, sh.getLastRow() - 4), 4).getValues().forEach((r, n) => {
    if (!r[0]) return;
    const date = r[0] && typeof r[0].getTime === 'function' ? Utilities.formatDate(r[0], 'America/Toronto', 'yyyy-MM-dd HH:mm') : String(r[0]);
    if (date.slice(0, 10) < since) return;
    const type = String(r[1]), tm = String(r[2]), det = String(r[3]), items = [];
    let m;
    if (/Échange/.test(type) && (m = det.match(/^(.*) reçoit (.*) · (.*) reçoit (.*)$/))) {
      const a = m[1].trim(), b = m[3].trim();
      list(m[2]).forEach(x => items.push({ name: x, kind: 'TRADE', from: b, to: a }));
      list(m[4]).forEach(x => items.push({ name: x, kind: 'TRADE', from: a, to: b }));
    } else if (/Réclamation/.test(type) && (m = det.match(/réclame (.*) \(libéré par (.*)\)$/))) {
      items.push({ name: m[1].trim(), kind: 'ADD', from: '', to: tm }, { name: m[1].trim(), kind: 'DROP', from: m[2].trim(), to: '' });
    } else if ((m = det.match(/^drop (.*) pour ajouter (.*)$/))) {
      list(m[1]).forEach(x => items.push({ name: x, kind: 'DROP', from: tm, to: '' }));
      list(m[2]).forEach(x => items.push({ name: x, kind: 'ADD', from: '', to: tm }));
    } else if ((m = det.match(/^ajoute (.*)$/))) list(m[1]).forEach(x => items.push({ name: x, kind: 'ADD', from: '', to: tm }));
    else if ((m = det.match(/^drop (.*)$/))) list(m[1]).forEach(x => items.push({ name: x, kind: 'DROP', from: tm, to: '' }));
    if (!items.length) return;
    const trade = items.some(i => i.kind === 'TRADE');
    TX_SEEN['News Tracker ' + (trade ? 'échanges' : 'mouvements')] = (TX_SEEN['News Tracker ' + (trade ? 'échanges' : 'mouvements')] || 0) + 1;
    out.push({ id: 'nt' + n, date: date, trade: trade, team: trade ? '' : tm, items: items });
  });
  return out;
}

// Corrections manuelles (« Mon info ») : { nom: { out: 'aaaa-mm-jj' (absent jusqu'au), start: 'aaaa-mm-jj' (partant confirmé), role: 'pp1' | 'top' | 'down', note, t } }
function getOverrides_() {
  try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty(LG.OVR_PROP) || '{}'); } catch (e) { return {}; }
}
// Liste de surveillance (gardée sur le serveur : suit sur téléphone et ordi)
// Bilan de tes moves : chaque ajout, retrait ou échange de ton équipe est gardé (au-delà des 14 jours d'ESPN)
// avec la projection maison (FP/match) des joueurs AU MOMENT où l'app l'a vu passer — pour juger la décision plus tard.
function myMoves_(teamsOut, waivers) {
  const props = PropertiesService.getDocumentProperties();
  let list = []; try { list = JSON.parse(props.getProperty('MYMOVES2') || '[]'); } catch (e) { list = []; }
  // 2 sources : l'API ESPN (exacte : heure réelle, moves intermédiaires) et le News Tracker du classeur (approximatif :
  // heure de la synchro, plusieurs moves fusionnés). Quand ESPN répond, on ignore le News Tracker sur la période qu'ESPN couvre.
  const espnTx = (TX_CACHE || []).filter(t => !/^nt/.test(t.id));
  const espnFrom = espnTx.length ? espnTx.map(t => t.date).reduce((a, b) => (a < b ? a : b)) : null;
  if (espnFrom) list = list.filter(m => m.s !== 'nt' || m.d < espnFrom);
  list = list.filter(m => !/Joueur #/.test(m.k));            // nom inconnu à l'époque : on le reprend avec le bon nom
  const first = props.getProperty('MYMOVES_SINCE') || stamp_();
  if (!props.getProperty('MYMOVES_SINCE')) props.setProperty('MYMOVES_SINCE', first);
  const idx = {}; teamsOut.forEach(t => t.players.forEach(p => { idx[p.name] = p; })); waivers.forEach(p => { if (!idx[p.name]) idx[p.name] = p; });
  const have = {}; list.forEach(m => { have[m.k] = true; });
  const r2 = x => Math.round((+x || 0) * 100) / 100;
  const pl = n => { const p = idx[n]; return p ? { n: n, id: p.id || null, g: p.grp, m: r2(p.model) } : { n: n, id: null, g: null, m: null }; };
  let added = 0;
  (TX_CACHE || []).forEach(t => {
    if (espnFrom && /^nt/.test(t.id) && t.date >= espnFrom) return;
    const mine = t.items.filter(i => i.to === APP.MY_TEAM || i.from === APP.MY_TEAM);
    if (!mine.length) return;
    const ins = mine.filter(i => i.to === APP.MY_TEAM).map(i => i.name), outs = mine.filter(i => i.from === APP.MY_TEAM).map(i => i.name);
    // ESPN montre parfois un échange comme des retraits + ajouts au même moment : on ignore un move dont les ajoutés
    // sont les joueurs reçus dans un échange du même jour (déjà compté comme échange)
    const k = t.date + '|' + ins.slice().sort().join(',') + '|' + outs.slice().sort().join(',');
    if (have[k]) return;
    have[k] = true; added++;
    list.push({ k: k, s: /^nt/.test(t.id) ? 'nt' : 'espn', d: t.date, tr: !!t.trade, late: t.date < first, i: ins.map(pl), o: outs.map(pl) });
  });
  list.sort((a, b) => (a.d < b.d ? 1 : -1));
  list = list.slice(0, 45);
  if (added || espnFrom) { try { props.setProperty('MYMOVES2', JSON.stringify(list)); } catch (e) { WARN.push('Bilan de tes moves : trop gros pour être gardé'); } }
  return list;
}
// Recrues de la saison (définition LNH : filtre isRookie des stats officielles), gardées 12 h — onglet Prédictions
function rookies_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('ROOKIES');
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* on refait */ } }
  try {
    const exp = encodeURIComponent(`seasonId=${CONFIG.SEASON} and gameTypeId=${CONFIG.GAME_TYPE} and isRookie="1"`);
    const ids = [];
    fetchJsonAll_(['skater', 'goalie'].map(k => `https://api.nhle.com/stats/rest/en/${k}/summary?limit=400&cayenneExp=${exp}`))
      .forEach(j => ((j && j.data) || []).forEach(r => { if (r.playerId) ids.push(r.playerId); }));
    if (ids.length) cache.put('ROOKIES', JSON.stringify(ids), 43200);
    return ids;
  } catch (e) { WARN.push('Recrues : ' + e.message); return []; }
}
function getWatch_() { try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty('WATCH') || '[]'); } catch (e) { return []; } }
function saveWatch(list) {          // appelée par l'app
  const clean = (Array.isArray(list) ? list : []).map(String).filter(Boolean).slice(0, 60);
  PropertiesService.getDocumentProperties().setProperty('WATCH', JSON.stringify(clean));
  return JSON.stringify(clean);
}
function saveOverride(name, obj) {          // appelée par l'app (fiche de joueur)
  const all = getOverrides_();
  if (obj && obj.pts) obj.pts = Math.round(+obj.pts) || '';
  if (!obj || (!obj.out && !obj.start && !obj.role && !obj.note && !obj.pts)) delete all[name];
  else all[name] = Object.assign({}, obj, { t: ymd_(new Date()) });
  // ménage : on oublie les corrections expirées depuis plus de 7 jours
  const old = ymd_(new Date(Date.now() - 7 * 86400000));
  Object.keys(all).forEach(n => { const o = all[n]; if (o.pts) return; if ((o.out && o.out < old && !o.role && !o.note) || (o.start && o.start < old && !o.out && !o.role && !o.note)) delete all[n]; });
  PropertiesService.getDocumentProperties().setProperty(LG.OVR_PROP, JSON.stringify(all));
  return JSON.stringify(all);
}

// Nouvelles d'un joueur (fil public d'ESPN LNH, gardé 30 min) — demandé par la fiche au besoin
function playerNews(name) {
  const cache = CacheService.getScriptCache(), key = 'NEWS_FEED';
  let feed = cache.get(key);
  if (!feed) {
    const r = UrlFetchApp.fetch('https://site.api.espn.com/apis/site/v2/sports/hockey/nhl/news?limit=100', { muteHttpExceptions: true });
    feed = r.getResponseCode() === 200 ? r.getContentText() : '{}';
    try { cache.put(key, feed.length < 95000 ? feed : JSON.stringify({ articles: (JSON.parse(feed).articles || []).slice(0, 60) }), 1800); } catch (e) { /* rien */ }
  }
  let arts = []; try { arts = JSON.parse(feed).articles || []; } catch (e) { arts = []; }
  const n = norm_(name), last = norm_(String(name).split(' ').slice(-1)[0]);
  const hits = arts.filter(a => {
    const cats = (a.categories || []).some(c => c.athlete && norm_(c.athlete.description || c.description || '') === n || norm_(c.description || '') === n);
    const txt = norm_((a.headline || '') + ' ' + (a.description || ''));
    return cats || txt.indexOf(n) >= 0 || (last.length > 4 && txt.indexOf(last) >= 0 && txt.indexOf(norm_(String(name).split(' ')[0])) >= 0);
  }).slice(0, 6).map(a => ({ h: a.headline || '', d: a.description || '', t: String(a.published || '').slice(0, 16).replace('T', ' '),
                             u: a.links && a.links.web && a.links.web.href || '' }));
  return JSON.stringify(hits);
}

// Sonde MoneyPuck (1 fois par jour) : où sont les gardiens partants probables / confirmés ? (résultat lisible par Claude : ?what=mp)
function mpProbe_() {
  // Sonde v2 : où MoneyPuck (et DailyFaceoff) mettent-ils les gardiens partants ? Une fois par jour, résultat dans « MP Sonde »
  const props = PropertiesService.getDocumentProperties(), today = ymd_(new Date());
  if (props.getProperty('MP_PROBE_DATE2') === today) return;
  props.setProperty('MP_PROBE_DATE2', today);
  const UA = { 'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36', 'Accept': 'text/html,application/json,*/*' };
  const get = u => { try { const r = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true, headers: UA }); return { code: r.getResponseCode(), t: r.getContentText() }; } catch (e) { return { code: 0, t: '', err: e.message }; } };
  const out = { v: 2, date: today, mp: {}, dfo: {}, tries: [] };
  try {
    const h = get('https://moneypuck.com/index.html');
    out.mp.code = h.code; out.mp.size = h.t.length;
    // scripts « en ligne » (sans src) : c'est là que la page va chercher ses données
    out.mp.inline = (h.t.match(/<script(?![^>]*src)[^>]*>[\s\S]*?<\/script>/gi) || []).map(x => x.replace(/\s+/g, ' ').slice(0, 1500)).slice(0, 8);
    out.mp.strings = []; (h.t.match(/["'][^"'\s]+\.(?:csv|json|txt|htm|html|js)(?:\?[^"']*)?["']/gi) || []).forEach(x => { if (out.mp.strings.indexOf(x) < 0) out.mp.strings.push(x); });
    out.mp.iframes = h.t.match(/<iframe[^>]*>/gi) || [];
    out.mp.cloudflare = /challenge-platform|cf-chl/i.test(h.t);
  } catch (e) { out.mp.err = e.message; }
  // essais directs (noms de fichiers plausibles)
  ['https://moneypuck.com/moneypuck/startingGoalies.csv', 'https://moneypuck.com/startingGoalies.htm', 'https://moneypuck.com/moneypuck/predictions/startingGoalies.csv',
   'https://moneypuck.com/moneypuck/todaysGames.csv', 'https://moneypuck.com/moneypuck/tweets/todaysGames.csv'].forEach(u => {
    const r = get(u); out.tries.push({ url: u, code: r.code, size: r.t.length, goalie: (r.t.match(/goalie/gi) || []).length, head: r.t.slice(0, 160) });
  });
  // DailyFaceoff : page « starting goalies » (site Next.js : données dans __NEXT_DATA__)
  try {
    const d = get('https://www.dailyfaceoff.com/starting-goalies/');
    out.dfo.code = d.code; out.dfo.size = d.t.length; out.dfo.err = d.err;
    const nd = d.t.match(/<script id="__NEXT_DATA__"[^>]*>([\s\S]*?)<\/script>/);
    out.dfo.nextData = !!nd;
    if (nd) {
      const txt = nd[1];
      out.dfo.nextSize = txt.length;
      out.dfo.keys = (txt.match(/"[a-zA-Z]*(?:[Gg]oalie|[Ss]trength|[Cc]onfirm|[Ss]tatus)[a-zA-Z]*"/g) || []).filter((x, i, a) => a.indexOf(x) === i).slice(0, 40);
      const k = txt.search(/[Cc]onfirmed|[Ll]ikely|[Uu]nconfirmed/);
      out.dfo.sample = k >= 0 ? txt.slice(Math.max(0, k - 900), k + 600) : txt.slice(0, 1500);
    } else out.dfo.head = d.t.slice(0, 400);
  } catch (e) { out.dfo.err = e.message; }
  const sh = hiddenSheet_(SpreadsheetApp.getActiveSpreadsheet(), 'MP Sonde');
  sh.clearContents();
  const json = JSON.stringify(out), parts = []; for (let i = 0; i < json.length; i += 45000) parts.push(json.slice(i, i + 45000));
  sh.getRange(1, 1, 1, parts.length).setValues([parts]);
}
function getMpProbe_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MP Sonde');
  return sh && sh.getLastRow() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].join('') : '';
}
