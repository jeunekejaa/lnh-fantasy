/**
 * 🔮 MOTEUR DE PROJECTION MAISON (fichier 4 du projet Apps Script)
 *
 * Projette n'importe quelle saison à partir des 3 saisons précédentes, sans GAF :
 *   • taux par match de chaque stat (buts, passes, pts AN, pts DN, tirs, mises en échec, tirs bloqués),
 *     pondérés par récence (ex. 6/3/1) et par matchs joués ;
 *   • « régression vers la moyenne » : on ajoute R matchs d'un joueur moyen de sa position ;
 *   • âge : progression avant 26 ans, déclin après 29 ans ;
 *   • rôle : variation du temps de jeu (total et en avantage numérique) la saison précédente ;
 *   • matchs joués projetés (historique de blessures) ;
 *   • stats avancées MoneyPuck (chacune testée, gardée seulement si elle réduit l'erreur) :
 *       buts attendus (3 variantes, dont « avec rebonds mérités ») × finition de carrière,
 *       finition séparée sur les tirs de haute qualité, passes primaires vs secondaires,
 *       tentatives de tir, chance de l'équipe sur la glace ;
 *   • gardiens : FP par match pondérés + buts sauvés au-delà des attentes (GSAx) sur 3 saisons.
 * Tous les réglages sont choisis par un backtest sur 4 saisons (2022-23 à 2025-26).
 * Pour la saison en cours, la projection finale combine GAF et ce moteur (poids appris sur 2025-26).
 *
 * Menus : 🔮 Régler le moteur (backtest 4 saisons) · 🔮 Prédire une saison…
 * Données : API statistiques de la LNH (api.nhle.com) · MoneyPuck.com (usage personnel, non commercial).
 */
const PJ = {
  SEASONS: ['20192020', '20202021', '20212022', '20222023', '20232024', '20242025', '20252026'],
  TARGETS: ['20222023', '20232024', '20242025', '20252026'],
  GAMES: { '20192020': 70, '20202021': 56, '20262027': 84 },   // saisons écourtées ou allongées (autres : 82 ; 84 dès 2026-27)
  CACHE_SHEET: 'PJ Cache', RESULT_SHEET: 'PJ Résultat', SHEET: 'Projection',
  MIN_GP: 20,
  TIME_LIMIT: 300000,
  MIN_GAIN: 0.0005,                     // gain minimal (FP/match) pour garder une stat avancée                   // au-delà (ms), on arrête de chercher et on enregistre le meilleur trouvé
  GRID: {
    w: [[5, 4, 3], [6, 3, 1], [3, 2, 1], [1, 1, 1], [1, 0, 0]],
    R: [0, 10, 25, 50, 100],
    c1: [0, 0.01, 0.02, 0.03],          // progression par année sous 26 ans
    c2: [0, 0.01, 0.02, 0.03, 0.05],    // déclin par année au-delà de 29 ans
    beta: [0, 0.5, 1]                   // poids du changement de rôle (temps de jeu)
  },
  DEFAULT: { w: [6, 3, 1], R: 25, c1: 0.01, c2: 0.02, beta: 0.5, gafW: 0.6, Kf: null, aL: 0, xgv: 'xg', Kh: null, lam: 1, rho: 0, Ken: null },
  // Stats avancées MoneyPuck (saisons passées) — recherche « une option à la fois », répétée jusqu'à stabilité.
  // La 1re valeur de chaque liste = option désactivée.
  MP_GRID: {
    Kf: [null, 0, 10, 25, 50, 100, 1e6],      // finition : buts / buts attendus, ramenée vers 1 avec Kf buts attendus (1e6 = xG purs)
    xgv: ['xg', 'xgr', 'xgf'],                // buts attendus : standard, avec rebonds mérités, ajustés « en rafale »
    Kh: [null, 0, 5, 10, 25, 50, 1e6],         // finition séparée sur les tirs de haute qualité (null = une seule finition)
    aL: [0, 0.25, 0.5, 0.75, 1],              // part corrigée de la chance de l'équipe sur la glace (passes, pts AN)
    lam: [1, 0.75, 0.5, 0.25, 0],             // poids des passes secondaires (1 = comme les primaires)
    rho: [0, 0.25, 0.5, 0.75, 1],             // tirs projetés à partir des tentatives de tir (0 = tirs seulement)
    Ken: [null, 0, 25, 100, 300, 1000]        // points dans un filet désert projetés à part (null = non ; régression vers la ligue en matchs)
  },
  // Gardiens : poids des saisons, régression (matchs d'un gardien moyen), poids du GSAx, régression du GSAx
  G_GRID: { w: [[6, 3, 1], [5, 4, 3], [3, 2, 1], [1, 1, 1], [1, 0, 0]], Rg: [0, 10, 25, 50, 100, 200],
            gamma: [-0.5, 0, 0.5, 1, 1.5, 2, 3], Kg: [0, 25, 50, 100, 200] },
  G_DEFAULT: { w: [6, 3, 1], Rg: 50, gamma: 0, Kg: 50, gafW: 0.5, gafB: -0.5 }
};
const PJ_STATS = ['g', 'a', 'ppp', 'shp', 'sog', 'hit', 'blk'];
const PJ_OFF = { g: 1, a: 1, ppp: 1, sog: 1 };   // stats offensives (âge et rôle)

// ---------------- Données d'une saison (API LNH), gardées en cache pour les saisons terminées ----------------
function pjReportUrl_(path, season, start) {
  const sort = encodeURIComponent('[{"property":"playerId","direction":"ASC"}]');
  const exp = encodeURIComponent(`seasonId=${season} and gameTypeId=2`);
  return `https://api.nhle.com/stats/rest/en/${path}?isAggregate=true&start=${start}&limit=100&sort=${sort}&cayenneExp=${exp}`;
}
function pjReports_(season, paths) {
  const first = UrlFetchApp.fetchAll(paths.map(p => ({ url: pjReportUrl_(p, season, 0), muteHttpExceptions: true })))
    .map((r, i) => {
      if (r.getResponseCode() === 200) return JSON.parse(r.getContentText());
      if (paths[i] === 'skater/summary') throw new Error(`API LNH ${paths[i]} ${season} : code ${r.getResponseCode()}`);
      return { data: [], total: 0 };   // rapport secondaire indisponible : on continue sans
    });
  const out = {}, more = [];
  paths.forEach((p, i) => { out[p] = (first[i].data || []).slice(); for (let s = 100; s < (first[i].total || 0); s += 100) more.push({ p: p, s: s }); });
  for (let k = 0; k < more.length; k += 30) {
    const chunk = more.slice(k, k + 30);
    UrlFetchApp.fetchAll(chunk.map(m => ({ url: pjReportUrl_(m.p, season, m.s), muteHttpExceptions: true })))
      .forEach((r, i) => { if (r.getResponseCode() === 200) (JSON.parse(r.getContentText()).data || []).forEach(d => out[chunk[i].p].push(d)); });
  }
  return out;
}
function pjSeason_(season) {
  const cached = pjCacheGet_('v2_' + season);   // v2 : avec les points dans un filet désert
  if (cached) return cached;
  const rep = pjReports_(season, ['skater/summary', 'skater/realtime', 'skater/timeonice', 'skater/bios', 'goalie/summary']);
  const sk = {};
  rep['skater/summary'].forEach(r => {
    sk[r.playerId] = { n: r.skaterFullName, p: r.positionCode === 'D' ? 'D' : 'F', gp: r.gamesPlayed || 0,
      g: r.goals || 0, a: r.assists || 0, ppp: r.ppPoints || 0, shp: r.shPoints || 0, sog: r.shots || 0, hit: 0, blk: 0,
      toi: r.timeOnIcePerGame || 0, pptoi: 0, b: '' };
  });
  rep['skater/realtime'].forEach(r => { const s = sk[r.playerId]; if (s) { s.hit = r.hits || 0; s.blk = r.blockedShots || 0; s.eng = r.emptyNetGoals || 0; s.ena = r.emptyNetAssists || 0; } });
  rep['skater/timeonice'].forEach(r => { const s = sk[r.playerId]; if (s) { s.pptoi = r.ppTimeOnIcePerGame || 0; if (!s.toi) s.toi = r.timeOnIcePerGame || 0; } });
  rep['skater/bios'].forEach(r => { const s = sk[r.playerId]; if (s && r.birthDate) s.b = String(r.birthDate).slice(0, 10); });
  const gk = {};
  rep['goalie/summary'].forEach(r => {
    gk[r.playerId] = { n: r.goalieFullName, gp: r.gamesPlayed || 0, gs: r.gamesStarted || r.gamesPlayed || 0, w: r.wins || 0,
      ga: r.goalsAgainst || 0, sv: r.saves || 0, so: r.shutouts || 0, otl: r.otLosses || 0 };
  });
  const data = { sk: sk, gk: gk };
  if (season !== CONFIG.SEASON) pjCachePut_('v2_' + season, data);   // saison terminée : ne change plus
  return data;
}
// Cache : une ligne par clé (saison, MoneyPuck…), le JSON découpé en morceaux de 45 000 caractères.
// Toute erreur de la feuille (ex. « Feuille … introuvable ») est réessayée une fois, puis ignorée :
// le cache n'est qu'un raccourci, on relit alors les données à la source.
function pjSheetTry_(fn, fallback) {
  for (let t = 0; t < 2; t++) {
    try { return fn(); }
    catch (e) { PJ_CACHE_ERR.push(e.message); try { SpreadsheetApp.flush(); } catch (x) {} Utilities.sleep(800); }
  }
  return fallback;
}
const PJ_CACHE_ERR = [];
function pjCacheGet_(season) {
  return pjSheetTry_(() => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PJ.CACHE_SHEET);
    if (!sh || sh.getLastRow() < 1) return null;
    // on lit d'abord la colonne des clés, puis seulement la ligne voulue (la feuille peut peser plusieurs Mo)
    const keys = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    let i = -1;
    for (let k = keys.length - 1; k >= 0; k--) if (String(keys[k][0]) === season) { i = k; break; }
    if (i < 0) return null;
    const row = sh.getRange(i + 1, 2, 1, Math.max(1, sh.getLastColumn() - 1)).getValues()[0];
    try { return JSON.parse(row.join('')); } catch (e) { return null; }
  }, null);
}
function pjCachePut_(season, data) {
  const json = JSON.stringify(data), parts = [];
  for (let i = 0; i < json.length; i += 45000) parts.push(json.slice(i, i + 45000));
  pjSheetTry_(() => { hiddenSheet_(SpreadsheetApp.getActiveSpreadsheet(), PJ.CACHE_SHEET).appendRow([season].concat(parts)); return true; }, false);
}

// ---------------- MoneyPuck : saisons passées (gardées en cache) ----------------
// Patineurs : id LNH → tableau (champs dans PJ_MP), situation « all » + avantage numérique « 5on4 »
const PJ_MP = ['xg', 'onxg', 'ong', 'a1', 'a2', 'att', 'sogm', 'hdxg', 'hdg', 'xgr', 'xgf', 'pptoi', 'ppatt'];
const PJ_MP_COLS = { xg: 'I_F_xGoals', onxg: 'OnIce_F_xGoals', ong: 'OnIce_F_goals', a1: 'I_F_primaryAssists', a2: 'I_F_secondaryAssists',
  att: 'I_F_shotAttempts', sogm: 'I_F_shotsOnGoal', hdxg: 'I_F_highDangerxGoals', hdg: 'I_F_highDangerGoals',
  xgr: 'I_F_xGoals_with_earned_rebounds', xgf: 'I_F_flurryAdjustedxGoals' };
// Gardiens : id LNH → [buts attendus contre, buts contre, MJ] (situation « all »)
function pjMoneyPuck_(season) {
  const key = 'mp2_' + season, gkey = 'mpg_' + season;
  const cached = pjCacheGet_(key), gcached = pjCacheGet_(gkey);
  if (cached && gcached) return { sk: cached, gk: gcached };
  const base = `https://moneypuck.com/moneypuck/playerData/seasonSummary/${season.slice(0, 4)}/regular/`;
  const res = UrlFetchApp.fetchAll([base + 'skaters.csv', base + 'goalies.csv'].map(u => ({ url: u, muteHttpExceptions: true })));
  let sk = cached, gk = gcached;
  if (!sk && res[0].getResponseCode() === 200) {
    const v = Utilities.parseCsv(res[0].getContentText()), h = v[0], c = n => h.indexOf(n);
    const id = c('playerId'), sit = c('situation'), ice = c('icetime');
    const I = {}; Object.keys(PJ_MP_COLS).forEach(k => { I[k] = c(PJ_MP_COLS[k]); });
    if (id >= 0 && I.xg >= 0) {
      sk = { _miss: Object.keys(PJ_MP_COLS).filter(k => I[k] < 0) };
      const r2 = x => Math.round(x * 100) / 100;
      for (let i = 1; i < v.length; i++) {
        const row = v[i], pid = +row[id];
        if (!pid) continue;
        const o = sk[pid] || (sk[pid] = PJ_MP.map(() => 0));
        if (row[sit] === 'all') Object.keys(PJ_MP_COLS).forEach(k => { if (I[k] >= 0) o[PJ_MP.indexOf(k)] = r2(+row[I[k]] || 0); });
        else if (row[sit] === '5on4') { o[PJ_MP.indexOf('pptoi')] = Math.round(+row[ice] || 0); if (I.att >= 0) o[PJ_MP.indexOf('ppatt')] = +row[I.att] || 0; }
      }
      if (season !== CONFIG.SEASON) pjCachePut_(key, sk);
    }
  }
  if (!gk && res[1].getResponseCode() === 200) {
    const v = Utilities.parseCsv(res[1].getContentText()), h = v[0], c = n => h.indexOf(n);
    const I = { id: c('playerId'), sit: c('situation'), gp: c('games_played'), xg: c('xGoals'), ga: c('goals') };
    if (I.id >= 0 && I.xg >= 0 && I.ga >= 0) {
      gk = {};
      for (let i = 1; i < v.length; i++) {
        const row = v[i];
        if (row[I.sit] !== 'all' || !+row[I.id]) continue;
        gk[+row[I.id]] = [Math.round((+row[I.xg] || 0) * 100) / 100, +row[I.ga] || 0, +row[I.gp] || 0];
      }
      if (season !== CONFIG.SEASON) pjCachePut_(gkey, gk);
    }
  }
  return sk || gk ? { sk: sk || null, gk: gk || null } : null;
}
// Ajoute les stats MoneyPuck aux saisons chargées, tant qu'il reste du temps (sinon : relancer)
function pjAttachMp_(data, deadline) {
  const missing = [];
  Object.keys(data).forEach(s => {
    if (!data[s] || data[s].mp) return;
    if (Date.now() > deadline && !pjCacheGet_('mp2_' + s)) { missing.push(s); return; }
    try {
      const m = pjMoneyPuck_(s);
      data[s].mp = m && m.sk || null; data[s].mpg = m && m.gk || null;
      if (!data[s].mp) missing.push(s);
    } catch (e) { missing.push(s); }
  });
  return missing;
}
// Ligne MoneyPuck d'un patineur → objet
const pjMpObj_ = arr => { const o = {}; PJ_MP.forEach((k, i) => { o[k] = arr[i] || 0; }); return o; };

// ---------------- Projection ----------------
const pjGames_ = season => PJ.GAMES[season] || (+season.slice(0, 4) >= 2026 ? 84 : 82);
const pjPrev_ = (season, k) => { const y = +season.slice(0, 4) - k; return `${y}${y + 1}`; };
function pjAge_(birth, season) {
  if (!birth) return null;
  const b = new Date(birth + 'T00:00:00Z'), ref = new Date(season.slice(0, 4) + '-10-01T00:00:00Z');
  return (ref - b) / (365.25 * 86400000);
}
// Moyenne de la ligue par position (joueurs avec 20 matchs et plus la saison précédente)
function pjLeague_(seasonData) {
  const acc = { F: { gp: 0, a1: 0, a2: 0, sogm: 0, att: 0 }, D: { gp: 0, a1: 0, a2: 0, sogm: 0, att: 0 } };
  const mp = seasonData.mp || {};
  Object.keys(seasonData.sk).forEach(id => {
    const s = seasonData.sk[id];
    if (s.gp < PJ.MIN_GP) return;
    const a = acc[s.p]; a.gp += s.gp; PJ_STATS.forEach(k => { a[k] = (a[k] || 0) + s[k]; }); a.eng = (a.eng || 0) + (s.eng || 0); a.ena = (a.ena || 0) + (s.ena || 0);
    const m = mp[id];
    if (m) { const o = pjMpObj_(m); a.a1 += o.a1; a.a2 += o.a2; a.sogm += o.sogm; a.att += o.att; }
  });
  const out = {};
  ['F', 'D'].forEach(p => {
    out[p] = {}; PJ_STATS.forEach(k => { out[p][k] = acc[p].gp ? acc[p][k] / acc[p].gp : 0; });
    out[p].eng = acc[p].gp ? (acc[p].eng || 0) / acc[p].gp : 0; out[p].ena = acc[p].gp ? (acc[p].ena || 0) / acc[p].gp : 0;
    out[p].r21 = acc[p].a1 ? acc[p].a2 / acc[p].a1 : 0.8;      // passes secondaires par passe primaire (ligue)
    out[p].sa = acc[p].att ? acc[p].sogm / acc[p].att : 0.5;   // tirs cadrés par tentative (ligue)
  });
  return out;
}
// Projection de toutes les stats par match d'un joueur pour la saison `season`
// hist = [saison-1, saison-2, saison-3] (objets de stats ou null), lg = moyenne de la ligue de sa position
function pjProjectSkater_(hist, lg, age, P, games) {
  let den = P.R, wsum = 0, toiW = 0, ppW = 0, gpN = 0, wgp = 0;
  const num = {}, useEN = P.Ken != null;
  PJ_STATS.forEach(k => { num[k] = P.R * lg[k]; });
  // Option Ken : les points dans un filet désert sont projetés à part (ramenés vers la ligue avec Ken matchs), le reste comme avant
  let enG = 0, enA = 0, enGp = 0;
  if (useEN) { num.g -= P.R * (lg.eng || 0); num.a -= P.R * (lg.ena || 0); }
  // Finition « de carrière » (3 saisons) : buts / buts attendus, ramenée vers 1 avec Kf buts attendus fictifs.
  // Option Kh : finition séparée sur les tirs de haute qualité (ramenée vers 1 avec Kh) et sur les autres (Kf).
  const xv = P.xgv || 'xg', X_ = h => (h[xv] > 0 ? h[xv] : h.xg) || 0;
  let G = 0, X = 0, Gh = 0, Xh = 0, S = 0, A = 0;
  hist.forEach(h => {
    if (!h) return;
    if (h.xg > 0) { G += h.g; X += X_(h); if (h.hdxg > 0) { Gh += h.hdg; Xh += h.hdxg; } }
    if (h.att > 0) { S += h.sogm; A += h.att; }
  });
  const useX = P.Kf != null && X > 0, split = useX && P.Kh != null && Xh > 0 && X > Xh;
  const finH = split ? (Gh + P.Kh) / (Xh + P.Kh) : null;
  const finO = !useX ? null : split ? (G - Gh + P.Kf) / (X - Xh + P.Kf) : (G + P.Kf) / (X + P.Kf);
  const sa = (S + 150 * lg.sa) / (A + 150);   // tirs cadrés par tentative, propre au joueur (ramené vers la ligue)
  hist.forEach((h, i) => {
    if (!h || !h.gp || !P.w[i]) return;
    const w = P.w[i];
    const eff = Object.assign({}, h);
    if (useX && h.xg > 0) {                                                  // buts mérités par ses chances × sa finition
      const x = X_(h);
      eff.g = split ? h.hdxg * finH + Math.max(0, x - h.hdxg) * finO : x * finO;
    }
    if (P.lam != null && P.lam < 1 && h.a1 + h.a2 > 0)                       // passes secondaires ramenées vers ce que ses primaires laissent attendre
      eff.a = Math.max(0, h.a - (1 - P.lam) * (h.a2 - h.a1 * lg.r21));
    if (P.aL && h.ong >= 10 && h.onxg > 0) {                                 // chance de l'équipe quand il est sur la glace
      const f = Math.max(0.75, Math.min(1.33, 1 + P.aL * (h.onxg / h.ong - 1)));
      eff.a *= f; eff.ppp = h.ppp * (1 + (f - 1) / 2);
    }
    if (P.rho && h.att > 0) eff.sog = (1 - P.rho) * h.sog + P.rho * h.att * sa;   // tirs selon les tentatives
    if (useEN) { eff.g = Math.max(0, eff.g - (h.eng || 0)); eff.a = Math.max(0, eff.a - (h.ena || 0)); enG += w * (h.eng || 0); enA += w * (h.ena || 0); enGp += w * h.gp; }
    PJ_STATS.forEach(k => { num[k] += w * eff[k]; });
    den += w * h.gp; wsum += w;
    toiW += w * h.gp * (h.toi || 0); ppW += w * h.gp * (h.pptoi || 0); wgp += w * h.gp;
    gpN += w * Math.min(1, h.gp / pjGames_(h.season));
  });
  if (!wsum) return null;
  const rate = {};
  PJ_STATS.forEach(k => { rate[k] = num[k] / den; });
  // âge
  if (age) {
    const m = Math.max(0.7, Math.min(1.25, 1 + P.c1 * Math.max(0, 26 - age) - P.c2 * Math.max(0, age - 29)));
    Object.keys(PJ_OFF).forEach(k => { rate[k] *= m; });
  }
  // rôle : temps de jeu de la dernière saison vs moyenne pondérée
  const last = hist[0];
  if (P.beta && last && last.gp >= 10 && wgp) {
    const tw = toiW / wgp, pw = ppW / wgp;
    if (tw > 0 && last.toi > 0) { const f = Math.pow(last.toi / tw, P.beta); ['g', 'a', 'sog'].forEach(k => { rate[k] *= Math.max(0.8, Math.min(1.25, f)); }); }
    if (pw > 30 && last.pptoi > 0) { const f = Math.pow(last.pptoi / pw, P.beta); rate.ppp *= Math.max(0.7, Math.min(1.4, f)); }
  }
  if (useEN) { rate.g += (enG + P.Ken * (lg.eng || 0)) / (enGp + P.Ken); rate.a += (enA + P.Ken * (lg.ena || 0)) / (enGp + P.Ken); }
  // matchs joués projetés (part de la saison), ramenée vers une moyenne gpM avec un poids gpR (réglés par le backtest)
  const gpR = P.gpR == null ? 2 : P.gpR, gpM = P.gpM == null ? 0.88 : P.gpM;
  const share = (gpN + gpR * gpM) / (wsum + gpR);
  return { rate: rate, fpg: fp_(rate), gp: Math.round((games || 82) * Math.min(1, share)), sa: sa };
}
// Historique des 3 saisons précédentes de chaque joueur (préparé une fois par jeu de données)
const PJ_HCACHE = new WeakMap();
function pjHistory_(season, data) {
  let byData = PJ_HCACHE.get(data);
  if (!byData) { byData = {}; PJ_HCACHE.set(data, byData); }
  if (byData[season]) return byData[season];
  const lgData = data[pjPrev_(season, 1)];
  if (!lgData) return null;
  const ids = {};
  [1, 2, 3].forEach(k => { const d = data[pjPrev_(season, k)]; if (d) Object.keys(d.sk).forEach(id => { ids[id] = true; }); });
  const list = Object.keys(ids).map(id => {
    const hist = [1, 2, 3].map(k => {
      const d = data[pjPrev_(season, k)]; const h = d && d.sk[id];
      if (!h) return null;
      const m = d.mp && d.mp[id];
      return Object.assign({ season: pjPrev_(season, k) }, h, m ? pjMpObj_(m) : {});
    });
    const any = hist.find(h => h);
    return { id: id, hist: hist, any: any, age: pjAge_(any.b, season) };
  });
  return (byData[season] = { lg: pjLeague_(lgData), list: list });
}
function pjProjectSeason_(season, P, data) {
  const H = pjHistory_(season, data);
  if (!H) return {};
  const out = {}, games = pjGames_(season);
  H.list.forEach(e => {
    const pr = pjProjectSkater_(e.hist, H.lg[e.any.p], e.age, P, games);
    if (pr) out[e.id] = Object.assign(pr, { name: e.any.n, pos: e.any.p, age: e.age });
  });
  return out;
}
const pjActualFpg_ = s => (s && s.gp ? fp_(Object.fromEntries(PJ_STATS.map(k => [k, s[k] / s.gp]))) : null);

// ---------------- Gardiens ----------------
const pjGoalieFp_ = g => { const S = CONFIG.SCORING; return S.W * g.w + S.GA * g.ga + S.SV * g.sv + S.SO * g.so + S.OTL * g.otl; };
// FP par match projetés : moyenne pondérée des 3 saisons, ramenée vers la moyenne de la ligue (Rg matchs),
// + gamma × buts sauvés au-delà des attentes par match (GSAx MoneyPuck), ramenés vers 0 avec Kg matchs.
function pjProjectGoalies_(season, G, data) {
  const prev = data[pjPrev_(season, 1)];
  if (!prev) return {};
  let fs = 0, gs = 0;
  Object.keys(prev.gk).forEach(id => { const g = prev.gk[id]; if (g.gp >= PJ.MIN_GP) { fs += pjGoalieFp_(g); gs += g.gp; } });
  const lg = gs ? fs / gs : 3.5, ids = {}, out = {};
  [1, 2, 3].forEach(k => { const d = data[pjPrev_(season, k)]; if (d) Object.keys(d.gk).forEach(id => { ids[id] = true; }); });
  Object.keys(ids).forEach(id => {
    let num = G.Rg * lg, den = G.Rg, sx = 0, sw = 0, name = '', has = false;
    [1, 2, 3].forEach((k, i) => {
      const d = data[pjPrev_(season, k)], g = d && d.gk[id], w = G.w[i];
      if (!g || !g.gp || !w) return;
      has = true; name = name || g.n;
      num += w * pjGoalieFp_(g); den += w * g.gp;
      const m = d.mpg && d.mpg[id];
      if (m && m[2]) { sx += w * (m[0] - m[1]); sw += w * m[2]; }
    });
    if (!has || !den) return;
    out[id] = { fpg: num / den + (G.gamma && sw ? G.gamma * sx / (G.Kg + sw) : 0), name: name, lg: lg };
  });
  return out;
}

// ---------------- Réglage par backtest sur 4 saisons ----------------
function tuneProjection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Réglage du moteur : 7 saisons LNH + MoneyPuck, puis backtest (2 à 5 minutes)…', '🔮', 30);
  const t0 = Date.now(), late = () => Date.now() - t0 > PJ.TIME_LIMIT;
  const data = {};
  PJ.SEASONS.forEach(s => { data[s] = pjSeason_(s); });
  const mpMissing = pjAttachMp_(data, t0 + 200000);
  const mpCols = PJ.SEASONS.map(s => data[s].mp && data[s].mp._miss || []).reduce((a, m) => a.concat(m.filter(x => a.indexOf(x) < 0)), []);

  // jeux d'évaluation : joueurs avec 20 MJ et plus la saison cible ET un historique
  const evalSets = PJ.TARGETS.map(t => ({ t: t, act: data[t].sk }));
  let nCombos = 0;
  const score = P => {
    let err = 0, n = 0; nCombos++;
    evalSets.forEach(es => {
      const proj = pjProjectSeason_(es.t, P, data);
      Object.keys(proj).forEach(id => { const a = es.act[id]; if (!a || a.gp < PJ.MIN_GP) return; err += Math.abs(proj[id].fpg - pjActualFpg_(a)); n++; });
    });
    return { mae: err / Math.max(1, n), n: n };
  };
  const OFF = { Kf: null, xgv: 'xg', Kh: null, aL: 0, lam: 1, rho: 0, Ken: null };
  const r3 = x => Math.round(x * 1000) / 1000;
  // 1er temps : poids et R ; 2e temps : âge et rôle (sans MoneyPuck)
  let best = null;
  const tryP = (P, minGain) => { const s = score(P); if (!best || s.mae < best.mae - (minGain || 1e-7)) { best = Object.assign({ P: P }, s); return true; } return false; };
  PJ.GRID.w.forEach(w => PJ.GRID.R.forEach(R => tryP(Object.assign({ w: w, R: R, c1: 0.02, c2: 0.02, beta: 0 }, OFF))));
  PJ.GRID.c1.forEach(c1 => PJ.GRID.c2.forEach(c2 => PJ.GRID.beta.forEach(beta => tryP(Object.assign({}, best.P, { c1: c1, c2: c2, beta: beta })))));
  const sansMp = best.mae;
  // 3e temps : stats avancées MoneyPuck, une option à la fois, en boucle jusqu'à ce que plus rien ne s'améliore
  const passes = [];
  if (!mpMissing.length) {
    for (let pass = 0; pass < 3 && !late(); pass++) {
      let moved = false;
      Object.keys(PJ.MP_GRID).forEach(k => {
        if (late()) return;
        // une stat avancée n'est gardée que si elle réduit l'erreur d'au moins 0,0005 FP/match (sinon c'est du bruit)
        PJ.MP_GRID[k].forEach(v => { if (v !== best.P[k] && tryP(Object.assign({}, best.P, { [k]: v }), PJ.MIN_GAIN)) moved = true; });
      });
      passes.push(r3(best.mae));
      if (!moved) break;
    }
    // poids des saisons et régression revus avec les stats avancées
    if (!late()) PJ.GRID.w.forEach(w => PJ.GRID.R.forEach(R => tryP(Object.assign({}, best.P, { w: w, R: R }))));
  }
  // Stabilité des points dans un filet désert d'une saison à l'autre (joueurs avec 40 MJ et plus les deux saisons)
  const enStab = (() => { const xs = [], ys = [];
    for (let i = 1; i < PJ.SEASONS.length; i++) { const a = data[PJ.SEASONS[i - 1]].sk, b = data[PJ.SEASONS[i]].sk;
      Object.keys(b).forEach(id => { const u = a[id], v = b[id]; if (u && v && u.gp >= 40 && v.gp >= 40 && u.eng != null) { xs.push(((u.eng || 0) + (u.ena || 0)) / u.gp); ys.push(((v.eng || 0) + (v.ena || 0)) / v.gp); } }); }
    return { n: xs.length, corr: round_(pjCorr_(xs, ys)) }; })();
  const P = best.P;
  // Ablation : erreur quand on retire UNE option (combien chacune rapporte)
  const ablation = [];
  const LABEL = { Kf: 'Buts attendus × finition de carrière', xgv: 'Variante de buts attendus (rebonds / rafales)',
                  Kh: 'Finition séparée sur les tirs de haute qualité', aL: 'Chance de l’équipe sur la glace (passes)',
                  lam: 'Passes primaires vs secondaires', rho: 'Tentatives de tir → tirs', Ken: 'Points dans un filet désert projetés à part' };
  if (!mpMissing.length && !late()) {
    Object.keys(OFF).forEach(k => {
      const dep = k === 'Kf' ? { Kh: null, xgv: 'xg' } : {};
      const off = score(Object.assign({}, P, { [k]: OFF[k] }, dep)).mae;
      let on = { mae: Infinity, v: null };   // meilleure valeur de cette option, les autres restant au réglage final
      PJ.MP_GRID[k].forEach(v => {
        if (v === OFF[k]) return;
        const m = v === P[k] ? best.mae : score(Object.assign({}, P, { [k]: v }, k !== 'Kf' && P.Kf == null && (k === 'Kh' || k === 'xgv') ? { Kf: 10 } : {})).mae;
        if (m < on.mae) on = { mae: m, v: v };
      });
      ablation.push({ option: k, nom: LABEL[k], valeur: P[k], retenue: P[k] !== OFF[k], maeSans: r3(off), maeAvec: r3(on.mae), meilleure: on.v, gain: r3(off - on.mae) });
    });
    const none = score(Object.assign({}, P, OFF)).mae;
    ablation.push({ option: 'tout', nom: 'Toutes les stats avancées retenues', valeur: '', retenue: ablation.some(a => a.retenue),
                    maeSans: r3(none), maeAvec: r3(best.mae), gain: r3(none - best.mae) });
  }
  // 4e temps : matchs joués (erreur en MJ, joueurs qui ont joué la saison cible)
  let bestGp = null;
  [0, 0.5, 1, 2, 4, 8].forEach(gpR => [0.82, 0.85, 0.88, 0.92, 0.95].forEach(gpM => {
    let e = 0, n = 0;
    const Q = Object.assign({}, P, { gpR: gpR, gpM: gpM });
    evalSets.forEach(es => { const proj = pjProjectSeason_(es.t, Q, data);
      Object.keys(proj).forEach(id => { const a = es.act[id]; if (!a || a.gp < PJ.MIN_GP) return; e += Math.abs(proj[id].gp - a.gp); n++; }); });
    if (!bestGp || e / n < bestGp.mae) bestGp = { gpR: gpR, gpM: gpM, mae: e / n };
  }));
  P.gpR = bestGp.gpR; P.gpM = bestGp.gpM;

  // Détail par saison avec les meilleurs réglages + repères
  const perSeason = evalSets.map(es => {
    const proj = pjProjectSeason_(es.t, P, data), prev = data[pjPrev_(es.t, 1)].sk;
    let e = 0, eNaive = 0, eGp = 0, eTot = 0, ePts = 0, n = 0;
    const xs = [], ys = [];
    Object.keys(proj).forEach(id => {
      const a = es.act[id]; if (!a || a.gp < PJ.MIN_GP) return;
      const af = pjActualFpg_(a), pf = proj[id].fpg, p0 = prev[id] && prev[id].gp >= 10 ? pjActualFpg_(prev[id]) : pf;
      e += Math.abs(pf - af); eNaive += Math.abs(p0 - af); eGp += Math.abs(proj[id].gp - a.gp);
      eTot += Math.abs(pf * proj[id].gp - af * a.gp);
      ePts += Math.abs((proj[id].rate.g + proj[id].rate.a) * proj[id].gp - (a.g + a.a));
      xs.push(pf); ys.push(af); n++;
    });
    return { saison: es.t, n: n, mae: r3(e / n), naive: r3(eNaive / n), corr: round_(pjCorr_(xs, ys)),
             maeMJ: round_(eGp / n), maeFPsaison: round_(eTot / n), maePoints: round_(ePts / n) };
  });

  // Combinaison avec GAF sur 2025-26 (seule saison où on a GAF)
  let gafRes = null, gaf = null;
  try {
    gaf = btGaf_();
    const t = '20252026', proj = pjProjectSeason_(t, P, data), act = data[t].sk;
    const rows = [];
    Object.keys(proj).forEach(id => {
      const a = act[id], gp = gaf[btKey_(proj[id].name)];
      if (!a || a.gp < PJ.MIN_GP || !gp || gp.gfpg != null) return;
      rows.push({ m: proj[id].fpg, g: fp_(gp), a: pjActualFpg_(a) });
    });
    const mae = f => r3(rows.reduce((s, r) => s + Math.abs(f(r) - r.a), 0) / rows.length);
    let bw = 0, bm = Infinity;
    [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1].forEach(w => { const m = mae(r => w * r.g + (1 - w) * r.m); if (m < bm) { bm = m; bw = w; } });
    gafRes = { n: rows.length, gafSeul: mae(r => r.g), moteurSeul: mae(r => r.m), meilleurPoidsGaf: bw, combine: bm };
    P.gafW = bw;
  } catch (e) { gafRes = { erreur: e.message }; }

  // Gardiens : moteur FP/match + GSAx, réglé sur les 4 mêmes saisons, puis combiné à GAF sur 2025-26
  let gk = null;
  try { gk = pjTuneGoalies_(data, gaf, late); P.gk = gk.params; } catch (e) { gk = { erreur: e.message }; }

  const res = { params: P, mae: r3(best.mae), maeSansMoneyPuck: r3(sansMp), moneyPuckManquant: mpMissing, colonnesManquantes: mpCols,
                passes: passes, ablation: ablation, n: best.n, saisons: perSeason, gaf: gafRes, gardiens: gk,
                maeMJ: round_(bestGp.mae), filetDesert: enStab, erreursFeuille: PJ_CACHE_ERR.slice(0, 5), combos: nCombos, secondes: Math.round((Date.now() - t0) / 1000), coupe: late() };
  PropertiesService.getDocumentProperties().setProperty('PJ_PARAMS', JSON.stringify(P));
  pjWriteResult_(ss, res);
  ss.toast(mpMissing.length ? `Réglé sans MoneyPuck pour ${mpMissing.join(', ')} (temps écoulé) : relance le réglage pour les inclure.`
                            : `Moteur réglé : erreur ${res.mae} FP/match (sans stats avancées : ${res.maeSansMoneyPuck}) — ${res.secondes} s.`, '🔮', 12);
  return res;
}
// Gardiens qui ont joué 20 matchs et plus la saison cible
function pjTuneGoalies_(data, gaf, late) {
  const r3 = x => Math.round(x * 1000) / 1000;
  const sets = PJ.TARGETS.map(t => ({ t: t, act: data[t].gk }));
  const score = G => {
    let e = 0, n = 0;
    sets.forEach(es => {
      const proj = pjProjectGoalies_(es.t, G, data);
      Object.keys(proj).forEach(id => { const a = es.act[id]; if (!a || a.gp < PJ.MIN_GP) return; e += Math.abs(proj[id].fpg - pjGoalieFp_(a) / a.gp); n++; });
    });
    return { mae: e / Math.max(1, n), n: n };
  };
  let best = null;
  const GG = PJ.G_GRID;
  GG.w.forEach(w => GG.Rg.forEach(Rg => GG.gamma.forEach(gamma => (gamma ? GG.Kg : [50]).forEach(Kg => {
    if (late()) return;
    const G = { w: w, Rg: Rg, gamma: gamma, Kg: Kg }, s = score(G);
    if (!best || s.mae < best.mae - 1e-7) best = Object.assign({ G: G }, s);
  }))));
  const G = best.G, sansGsax = score(Object.assign({}, G, { gamma: 0 })).mae;
  // repères : saison précédente telle quelle, moyenne de la ligue
  let eN = 0, eL = 0, n = 0;
  sets.forEach(es => {
    const prev = data[pjPrev_(es.t, 1)].gk, proj = pjProjectGoalies_(es.t, G, data);
    Object.keys(proj).forEach(id => {
      const a = es.act[id]; if (!a || a.gp < PJ.MIN_GP) return;
      const af = pjGoalieFp_(a) / a.gp, p = prev[id];
      eN += Math.abs((p && p.gp >= 10 ? pjGoalieFp_(p) / p.gp : proj[id].lg) - af); eL += Math.abs(proj[id].lg - af); n++;
    });
  });
  const out = { n: best.n, mae: r3(best.mae), sansGsax: r3(sansGsax), saisonPrecedente: r3(eN / n), moyenneLigue: r3(eL / n) };
  // GAF 2025-26 : biais et poids optimaux
  G.gafW = PJ.G_DEFAULT.gafW; G.gafB = PJ.G_DEFAULT.gafB;
  if (gaf) {
    const t = '20252026', proj = pjProjectGoalies_(t, G, data), act = data[t].gk, rows = [];
    Object.keys(proj).forEach(id => {
      const a = act[id], g = gaf[btKey_(proj[id].name)];
      if (!a || a.gp < PJ.MIN_GP || !g || g.gfpg == null) return;
      rows.push({ m: proj[id].fpg, g: g.gfpg, a: pjGoalieFp_(a) / a.gp });
    });
    if (rows.length >= 8) {
      const mae = f => rows.reduce((s, r) => s + Math.abs(f(r) - r.a), 0) / rows.length;
      let bb = { m: Infinity };
      [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1].forEach(w => [0, -0.25, -0.5, -0.75, -1].forEach(b => {
        const m = mae(r => w * (r.g + b) + (1 - w) * r.m); if (m < bb.m) bb = { m: m, w: w, b: b };
      }));
      G.gafW = bb.w; G.gafB = bb.b;
      out.gaf2526 = { n: rows.length, gafSeul: r3(mae(r => r.g)), moteurSeul: r3(mae(r => r.m)), poidsGaf: bb.w, biaisGaf: bb.b, combine: r3(bb.m) };
    }
  }
  out.params = G;
  return out;
}
function pjCorr_(a, b) {
  const n = a.length; if (n < 3) return null;
  const ma = a.reduce((s, x) => s + x, 0) / n, mb = b.reduce((s, x) => s + x, 0) / n;
  let c = 0, va = 0, vb = 0; for (let i = 0; i < n; i++) { c += (a[i] - ma) * (b[i] - mb); va += (a[i] - ma) ** 2; vb += (b[i] - mb) ** 2; }
  return c / Math.sqrt(va * vb);
}
function pjParams_() {
  const s = PropertiesService.getDocumentProperties().getProperty('PJ_PARAMS');
  return Object.assign({}, PJ.DEFAULT, s ? JSON.parse(s) : {});
}
const pjUsesMp_ = P => P.Kf != null || P.aL || (P.lam != null && P.lam < 1) || P.rho || (P.gk && P.gk.gamma);
function pjWriteResult_(ss, res) {
  const sh = ss.getSheetByName(PJ.SHEET) || ss.insertSheet(PJ.SHEET);
  sh.clearContents();
  const P = res.params, rows = [], blank = ['', '', '', '', '', '', ''];
  const opt = (k, v) => ({ Kf: v == null ? 'non' : v >= 1e5 ? 'xG purs' : `ramenée vers 1 avec ${v} xG`, xgv: { xg: 'standard', xgr: 'avec rebonds mérités', xgf: 'ajustés en rafale' }[v],
    Kh: v == null ? 'non' : v >= 1e5 ? 'oui, sans régression' : `oui (ramenée avec ${v} xG)`, aL: `${Math.round(v * 100)} %`, lam: `poids ${v}`, rho: `${Math.round(v * 100)} % tentatives`, Ken: v == null ? 'non' : `oui (ramenés vers la ligue avec ${v} MJ)` }[k]);
  rows.push(['🔮 Moteur de projection maison — réglé sur 4 saisons', '', '', '', '', '', '']);
  rows.push([`Réglages : poids ${P.w.join('/')} · régression ${P.R} MJ · âge +${P.c1}/an avant 26, −${P.c2}/an après 29 · rôle ${P.beta} · MJ : ramenés vers ${Math.round((P.gpM || 0.88) * 100)} % de la saison avec un poids ${P.gpR}`, '', '', '', '', '', '']);
  rows.push([`${res.combos} combinaisons testées · ${res.secondes} s · erreur ${res.mae} FP/match (sans stats avancées : ${res.maeSansMoneyPuck})` +
            (res.moneyPuckManquant && res.moneyPuckManquant.length ? ` · ⚠ MoneyPuck manquant : ${res.moneyPuckManquant.join(', ')} (relancer)` : '') +
            (res.coupe ? ' · ⚠ temps limite atteint (relancer pour finir la recherche)' : '') +
            (res.erreursFeuille && res.erreursFeuille.length ? ` · ⚠ cache : ${res.erreursFeuille[0]}` : ''), '', '', '', '', '', '']);
  rows.push(blank);
  if (res.ablation && res.ablation.length) {
    rows.push(['Stat avancée testée', 'Retenue ?', 'Réglage choisi', 'Erreur sans elle', 'Meilleure erreur avec elle', 'Gain (FP/match)', 'Meilleur réglage testé']);
    res.ablation.forEach(a => rows.push([a.nom, a.retenue ? 'oui' : a.option === 'tout' ? 'aucune' : 'non (gain < ' + PJ.MIN_GAIN + ')', a.option === 'tout' ? '' : opt(a.option, P[a.option]),
                                         a.maeSans, a.maeAvec, a.gain, a.option === 'tout' ? '' : opt(a.option, a.meilleure)]));
    if (res.filetDesert) rows.push([`Points dans un filet désert : corrélation d’une saison à l’autre ${res.filetDesert.corr} (${res.filetDesert.n} paires de saisons) — près de 0 = surtout du hasard ; près de 1 = un vrai rôle (fins de match)`, '', '', '', '', '', '']);
    if (res.colonnesManquantes && res.colonnesManquantes.length) rows.push([`Colonnes MoneyPuck introuvables : ${res.colonnesManquantes.join(', ')}`, '', '', '', '', '', '']);
    rows.push(blank);
  }
  rows.push(['Saison prédite', 'Joueurs', 'Erreur FP/match', 'Repère : saison précédente telle quelle', 'Corrélation', 'Erreur MJ', 'Erreur points (B+P) sur la saison']);
  res.saisons.forEach(s => rows.push([s.saison, s.n, s.mae, s.naive, s.corr, s.maeMJ, s.maePoints]));
  rows.push(blank);
  if (res.gaf && !res.gaf.erreur) {
    rows.push(['2025-26 avec GAF', 'Joueurs', 'GAF seul', 'Moteur seul', 'Poids GAF optimal', 'Combiné', '']);
    rows.push(['', res.gaf.n, res.gaf.gafSeul, res.gaf.moteurSeul, res.gaf.meilleurPoidsGaf, res.gaf.combine, '']);
    rows.push(blank);
  }
  const g = res.gardiens;
  if (g && !g.erreur) {
    const G = g.params;
    rows.push(['Gardiens (FP par match, 4 saisons)', 'Gardiens', 'Moteur', 'Sans GSAx', 'Saison précédente telle quelle', 'Moyenne de la ligue', '']);
    rows.push([`poids ${G.w.join('/')} · régression ${G.Rg} MJ · GSAx × ${G.gamma} (ramené avec ${G.Kg} MJ)`, g.n, g.mae, g.sansGsax, g.saisonPrecedente, g.moyenneLigue, '']);
    if (g.gaf2526) {
      rows.push(['Gardiens 2025-26 avec GAF', 'Gardiens', 'GAF seul', 'Moteur seul', 'Poids GAF', 'Biais GAF', 'Combiné']);
      rows.push(['', g.gaf2526.n, g.gaf2526.gafSeul, g.gaf2526.moteurSeul, g.gaf2526.poidsGaf, g.gaf2526.biaisGaf, g.gaf2526.combine]);
    }
  } else if (g) rows.push([`Gardiens : ${g.erreur}`, '', '', '', '', '', '']);
  sh.getRange(1, 1, rows.length, 7).setValues(rows);
  const d = hiddenSheet_(ss, PJ.RESULT_SHEET); d.clearContents(); d.getRange(1, 1).setValue(JSON.stringify(res));
}
function getProjectionResult() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PJ.RESULT_SHEET);
  return sh && sh.getLastRow() ? String(sh.getRange(1, 1).getValue()) : '';
}

// ---------------- Prédire une saison (menu) ----------------
function predictSeasonPrompt() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('🔮 Prédire une saison', 'Saison à prédire (ex. 2024-2025, ou 2026-2027 pour la saison en cours) :', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const m = r.getResponseText().match(/(\d{4})\D*(\d{4})?/);
  if (!m) { ui.alert('Format attendu : 2024-2025'); return; }
  predictSeason(`${m[1]}${+m[1] + 1}`);
}
function predictSeason(season) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), P = pjParams_();
  const data = {};
  [season, pjPrev_(season, 1), pjPrev_(season, 2), pjPrev_(season, 3)].forEach(s => { try { data[s] = pjSeason_(s); } catch (e) { /* saison absente */ } });
  if (pjUsesMp_(P)) { const keep = data[season]; delete data[season]; pjAttachMp_(data, Date.now() + 200000); if (keep) data[season] = keep; }
  const proj = pjProjectSeason_(season, P, data), act = data[season] ? data[season].sk : {};
  const gproj = pjProjectGoalies_(season, P.gk || PJ.G_DEFAULT, data), gact = data[season] ? data[season].gk : {};
  const rows = Object.keys(proj).map(id => {
    const p = proj[id], a = act[id], af = pjActualFpg_(a);
    const pts = (p.rate.g + p.rate.a) * p.gp;
    return [p.name, p.pos, p.age ? Math.floor(p.age) : '', p.gp, round_(p.fpg), Math.round(p.fpg * p.gp), Math.round(pts),
            a ? a.gp : '', af == null ? '' : round_(af), a ? Math.round(af * a.gp) : '', a ? a.g + a.a : '', af == null ? '' : round_(af - p.fpg)];
  }).sort((x, y) => y[5] - x[5]);
  // gardiens : FP par match (la part des départs dépend de l'équipe, non projetée ici)
  Object.keys(gproj).map(id => {
    const p = gproj[id], a = gact[id], af = a && a.gp ? pjGoalieFp_(a) / a.gp : null;
    return [p.name, 'G', '', '', round_(p.fpg), '', '', a ? a.gp : '', af == null ? '' : round_(af), a ? Math.round(pjGoalieFp_(a)) : '', '', af == null ? '' : round_(af - p.fpg)];
  }).sort((x, y) => y[4] - x[4]).forEach(r => rows.push(r));
  const name = 'Prédiction ' + season.slice(0, 4) + '-' + season.slice(6, 8);
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, 12).setValues([['Joueur', 'Pos', 'Âge', 'MJ prédits', 'FP/match prédits', 'FP prédits (saison)', 'Points prédits (B+P)',
                                        'MJ réels', 'FP/match réels', 'FP réels', 'Points réels', 'Écart FP/match (réel − prédit)']]);
  if (rows.length) sh.getRange(2, 1, rows.length, 12).setValues(rows);
  sh.setFrozenRows(1);
  ss.toast(`${rows.length} joueurs projetés pour ${season}.`, '🔮', 8);
  return rows.length;
}

// ---------------- Pour l'app : projection maison de la saison en cours (1 fois par jour) ----------------
// id → { fpg, gp, stats par match, sa (tirs cadrés/tentative), ppa (tentatives en AN/match la saison passée), ppt (s en AN/match) }
// + _gk : id gardien → FP par match ; _lg : ratios de la ligue (passes secondaires/primaires, tirs/tentatives)
function pjCurrentPriors_() {
  const props = PropertiesService.getDocumentProperties(), today = ymd_(new Date());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (props.getProperty('PJ_CUR_DATE4') === today) {
    const sh = ss.getSheetByName('PJ Courant');
    if (sh && sh.getLastRow()) return JSON.parse(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].join(''));
  }
  const P = pjParams_(), data = {};
  [1, 2, 3].forEach(k => { const s = pjPrev_(CONFIG.SEASON, k); try { data[s] = pjSeason_(s); } catch (e) { /* rien */ } });
  pjAttachMp_(data, Date.now() + 120000);   // aussi utile à l'app (rôle en AN, tirs par tentative), même si le moteur n'en a pas besoin
  const proj = pjProjectSeason_(CONFIG.SEASON, P, data), out = {};
  const last = data[pjPrev_(CONFIG.SEASON, 1)] || { sk: {} }, lmp = last.mp || {};
  Object.keys(proj).forEach(id => {
    const p = proj[id], o = Object.assign({ fpg: p.fpg, gp: p.gp, sa: round_(p.sa) }, p.rate), h = last.sk[id], m = lmp[id];
    if (h && h.gp >= 20 && m) { const mo = pjMpObj_(m); o.ppa = round_(mo.ppatt / h.gp); o.ppt = Math.round(mo.pptoi / h.gp); }
    // la saison passée du joueur, par match (pour « processus vs résultats » : panne, malchance, chance, vraie progression)
    if (h && h.gp >= 15) {
      const ly = { gp: h.gp, toi: round_((h.toi || 0) / 60), sog: round_(h.sog / h.gp), pts: round_((h.g + h.a) / h.gp), g: round_(h.g / h.gp) };
      if (m) { const mo = pjMpObj_(m); Object.assign(ly, { xg: round_(mo.xg / h.gp), att: round_(mo.att / h.gp), onxg: round_(mo.onxg / h.gp), ong: round_(mo.ong / h.gp) }); }
      o.ly = ly;
    }
    out[id] = o;
  });
  // variabilité de son niveau d'une saison à l'autre ÷ celle d'un joueur typique (k individuel, si le backtest le retient)
  const cvs = {};
  Object.keys(out).forEach(id => {
    const v = Object.keys(data).map(k => data[k] && data[k].sk[id]).filter(h => h && h.gp >= 20).map(h => pjActualFpg_(h));
    if (v.length >= 2) { const m = v.reduce((t, x) => t + x, 0) / v.length; if (m > 0.5) cvs[id] = Math.sqrt(v.reduce((t, x) => t + (x - m) ** 2, 0) / (v.length - 1)) / m; }
  });
  const med = Object.values(cvs).sort((a, b) => a - b)[Math.floor(Object.keys(cvs).length / 2)] || 0.1;
  Object.keys(cvs).forEach(id => { out[id].kr = round_(Math.max(0.33, Math.min(3, cvs[id] / med))); });
  const gp = pjProjectGoalies_(CONFIG.SEASON, P.gk || PJ.G_DEFAULT, data);
  out._gk = {}; Object.keys(gp).forEach(id => { out._gk[id] = round_(gp[id].fpg); });
  const H = pjHistory_(CONFIG.SEASON, data);
  if (H) out._lg = { F: { r21: round_(H.lg.F.r21), sa: round_(H.lg.F.sa) }, D: { r21: round_(H.lg.D.r21), sa: round_(H.lg.D.sa) } };
  const json = JSON.stringify(out), parts = [];
  for (let i = 0; i < json.length; i += 45000) parts.push(json.slice(i, i + 45000));
  const sh = hiddenSheet_(ss, 'PJ Courant'); sh.clearContents(); sh.getRange(1, 1, 1, parts.length).setValues([parts]);
  props.setProperty('PJ_CUR_DATE4', today);
  return out;
}
