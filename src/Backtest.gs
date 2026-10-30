/**
 * 🏒 BACKTEST 2025-26 (fichier 3 du projet Apps Script)
 * « Si on avait utilisé le modèle maison l'an passé, à quel point aurait-il prédit la vraie vie ? »
 *
 * Méthode (aucune donnée du futur n'est utilisée pour prédire) :
 *   • base d'avant-saison de chaque joueur : GAF 2025-26 (feuille publique de Crome) combiné au moteur maison ;
 *   • stats de la saison jusqu'à une date (lues match par match comme le mode direct), à 4 dates :
 *     28 oct., 1er déc., 15 janv., 1er mars — pour voir si la bonne vitesse de confiance (k) change en cours de saison ;
 *   • on prédit la production par match pour le RESTE de la saison et on compare au réel (complet − jusqu'à la date).
 * Stats « avancées » refaites à partir du déroulement des matchs (play-by-play), testées une à une :
 *   • buts attendus maison (distance du tir + rebond), au lieu des buts réels ;
 *   • passes primaires vs secondaires ;
 *   • tentatives de tir (cadrés + ratés + bloqués) pour prédire les tirs ;
 *   • rôle en avantage numérique cette saison (temps en AN si la feuille de match le donne, sinon tentatives en AN).
 * Les réglages gagnants sont enregistrés et utilisés par l'app (avec les vraies stats MoneyPuck de la saison en cours).
 *
 * Lancer : menu 🏒 LNH Live › 🧪 Backtest 2025-26 (≈ 2 à 5 minutes ; régler le moteur 🔮 avant).
 * Résultats : feuille « Backtest » et dans l'app (onglet Modèle).
 */
const BT = {
  SEASON: '20252026',
  START: '2025-10-06',
  CUTOFFS: [
    { d: '2025-10-28', label: '28 oct.', minEarly: 5, minRest: 20 },
    { d: '2025-12-01', label: '1er déc.', minEarly: 12, minRest: 20 },
    { d: '2026-01-15', label: '15 janv.', minEarly: 20, minRest: 20 },
    { d: '2026-03-01', label: '1er mars', minEarly: 30, minRest: 12 }
  ],
  READ_LIMIT: 230000,          // temps max (ms) de lecture des matchs ; les dates non atteintes sont sautées
  SHEET: 'Backtest', DATA_SHEET: 'BT Data',
  GAF_URL: 'https://docs.google.com/spreadsheets/d/1_vb4HoUcpKt9rWISKSPb_LnFau3zwA9OTUhTfWGje40/gviz/tq?tqx=out:csv&sheet=' +
           encodeURIComponent('Player Values - Pts'),
  K_GRID: [0, 2, 4, 6, 8, 10, 12, 15, 20, 25, 30, 35, 40, 50, 60, 80, 100, 130, 160, 200, 300, 500, 1000],
  // options testées, stat par stat (1re valeur = option désactivée)
  OPT: { g: ['g', 'tirs', 'xg'], a: [1, 0.75, 0.5, 0.25, 0], sog: [0, 0.25, 0.5, 0.75, 1], ppp: [0, 0.25, 0.5, 0.75, 1] },
  G_K: [0, 5, 10, 15, 20, 30, 50, 75, 100, 150, 200, 300], G_B: [0.25, 0, -0.25, -0.5, -0.75, -1],
  MIN_GAIN: 0.0002,            // gain minimal (erreur par match de la stat) pour garder une option (réglage stat par stat)
  MIN_GAIN_FP: 0.0005          // gain minimal (FP par match) pour garder une option (réglage final)
};
const SK_STATS = ['g', 'a', 'ppp', 'shp', 'sog', 'hit', 'blk'];
const BT_OPT_NAME = { g: 'Buts : source', a: 'Passes secondaires : poids', sog: 'Tirs : part des tentatives', ppp: 'Pts AN : poids du rôle en AN' };
const BT_G_LABEL = { g: 'buts réels', tirs: 'tirs × % habituel', xg: 'buts attendus maison' };

function runBacktest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try { ss.toast('Backtest 2025-26 à 4 dates (2 à 5 minutes)…', '🧪', 30); } catch (e) { /* lancé depuis l'app web : pas de fenêtre */ }
  const t0 = Date.now(), warn = [];

  // 1) Bases d'avant-saison : GAF 2025-26 + moteur maison (3 saisons précédentes)
  const gaf = btGaf_();
  let eng = null;
  try { eng = btEngine_(); } catch (e) { warn.push('Moteur indisponible (' + e.message + ') : base = GAF seul'); }
  const P = eng ? eng.P : PJ.DEFAULT, gafW = P.gafW == null ? 0.6 : P.gafW, GK = (P.gk || PJ.G_DEFAULT);

  // 2) Saison 2025-26 jusqu'à chaque date
  const early = btEarlySnapshots_(t0);
  if (early.skipped.length) warn.push(`Temps de lecture dépassé : dates sautées (${early.skipped.join(', ')}) — relance le backtest (il reprend où il était)`);
  if (!early.ppToi) warn.push('Temps en AN absent des feuilles de match : rôle en AN mesuré par les tentatives de tir en AN');

  // 3) Saison complète 2025-26 (API officielle)
  const keep = CONFIG.SEASON;
  let full;
  try { CONFIG.SEASON = BT.SEASON; full = fetchNhlPlayers_(); } finally { CONFIG.SEASON = keep; }

  const mix = (g, m) => { const o = {}; SK_STATS.forEach(k => { o[k] = gafW * (g[k] || 0) + (1 - gafW) * (m[k] || 0); }); return o; };
  const rate = (o, k) => (o.gp ? (+o[k] || 0) / o.gp : 0);

  // 4) Pour chaque date : joueurs évalués, taux observés et réels
  const sets = early.snaps.map(sn => {
    const c = sn.cut, sk = [], gk = [];
    // buts attendus maison : taux de but par case (distance × rebond), mesuré sur toute la ligue jusqu'à la date
    const S = new Array(12).fill(0), Gb = new Array(12).fill(0), lgA = { F: [0, 0], D: [0, 0] };
    Object.keys(sn.ext).forEach(id => { const x = sn.ext[id]; for (let i = 0; i < 12; i++) { S[i] += x.b[i]; Gb[i] += x.gb[i]; } });
    const tab = S.map((s, i) => (Gb[i] + 0.5) / (s + 5));
    full.forEach(f => {
      const e = sn.map[f.id];
      if (!e) return;
      const isG = f.pos === 'G';
      const rest = {};
      ['gp'].concat(isG ? ['w', 'ga', 'sv', 'so', 'otl'] : SK_STATS).forEach(k => { rest[k] = (+f[k] || 0) - (+e[k] || 0); });
      const mE = isG ? Math.ceil(c.minEarly * 0.6) : c.minEarly, mR = isG ? Math.ceil(c.minRest * 0.6) : c.minRest;
      if (e.gp < mE || rest.gp < mR) return;
      const key = btKey_(f.name), gp = gaf[key] || null;
      if (isG) {
        const eg = eng && eng.gk[f.id] ? eng.gk[f.id].fpg : null, gg = gp && gp.gfpg != null ? gp.gfpg : null;
        const prior = gg != null && eg != null ? GK.gafW * (gg + GK.gafB) + (1 - GK.gafW) * eg : gg != null ? gg + GK.gafB : eg != null ? eg : 3.1;
        gk.push({ name: f.name, n: e.gp, early: btGfp_(e), rest: btGfp_(rest), prior: prior, gaf: gg != null ? gg : 3.6 });
        return;
      }
      const pos = f.pos === 'D' ? 'D' : 'F', x = sn.ext[f.id] || btBlankExt_();
      if (x.a1 + x.a2 > 0) { lgA[pos][0] += x.a1; lgA[pos][1] += x.a2; }
      const pj = eng && eng.sk[f.id] ? eng.sk[f.id].rate : null;
      const base = gp && pj ? mix(gp, pj) : gp || pj || DEPTH_PRIOR[pos];
      const prev = eng && eng.prevPP[f.id];
      sk.push({ id: f.id, name: f.name, pos: pos, n: e.gp, early: e, x: x, rest: rest, base: base, gafBase: gp || DEPTH_PRIOR[pos],
                sa: eng && eng.sk[f.id] ? eng.sk[f.id].sa : (eng && eng.lg ? eng.lg[pos].sa : 0.52),
                ixg: x.b.reduce((s, v, i) => s + v * tab[i], 0), prevPP: prev || null,
                ly: eng && eng.ly[f.id] || null, age: eng && eng.age[f.id] || null, gaf: gp || null, pjr: pj || null,
                kr: eng && eng.kr[f.id] || 1, inj: btInjured_(x, e.gp),
                act: Object.fromEntries(SK_STATS.map(k => [k, rate(rest, k)])) });
    });
    const r21 = { F: lgA.F[0] ? lgA.F[1] / lgA.F[0] : 0.8, D: lgA.D[0] ? lgA.D[1] / lgA.D[0] : 0.8 };
    return { cut: c, sk: sk, gk: gk, r21: r21, ppToi: early.ppToi, sn: sn };
  });
  if (!sets.length) throw new Error('aucune date évaluée');

  // 5) Réglage stat par stat : option × k, erreur moyenne sur les 4 dates (chaque date pèse autant)
  const mae = arr => arr.reduce((a, x) => a + Math.abs(x), 0) / Math.max(1, arr.length);
  const tune = {}, kByCut = {};
  SK_STATS.forEach(s => {
    const vals = BT.OPT[s] || [null];
    const byOpt = {};
    vals.forEach(v => {
      let best = null;
      BT.K_GRID.forEach(k => {
        const m = sets.reduce((a, st) => a + mae(st.sk.map(r => btPred_(r, s, v, k, st) - r.act[s])), 0) / sets.length;
        if (!best || m < best.mae) best = { k: k, mae: m };
      });
      byOpt[v] = best;
    });
    // option gardée seulement si elle bat l'option « désactivée » d'au moins MIN_GAIN
    let pick = vals[0];
    vals.forEach(v => { if (byOpt[v].mae < byOpt[pick].mae - (v === vals[0] ? 0 : BT.MIN_GAIN)) pick = v; });
    tune[s] = { opt: pick, k: byOpt[pick].k, mae: byOpt[pick].mae, off: byOpt[vals[0]], byOpt: byOpt };
    // k optimal à chaque date (avec l'option retenue) → la confiance dans la saison en cours change-t-elle avec le temps ?
    kByCut[s] = sets.map(st => {
      let b = null;
      BT.K_GRID.forEach(k => { const m = mae(st.sk.map(r => btPred_(r, s, pick, k, st) - r.act[s])); if (!b || m < b.mae) b = { k: k, mae: m }; });
      return b.k;
    });
  });
  // 5b) Réglage final sur ce qui compte : l'erreur en FP par match (toutes les stats ensemble), moyenne des dates.
  // Départ = réglages actuels de l'app ; on ajuste une stat à la fois (option × k), en boucle jusqu'à stabilité.
  // Une option n'est gardée que si elle bat l'option « désactivée » d'au moins MIN_GAIN_FP.
  const fpOf = o => fp_(o);
  const SC = CONFIG.SCORING, SW = { g: SC.G, a: SC.A, ppp: SC.PPP, shp: SC.SHP, sog: SC.SOG, hit: SC.HIT, blk: SC.BLK };
  const flat = [];
  sets.forEach((st, si) => st.sk.forEach(r => flat.push({ r: r, st: st, si: si })));
  const actFp = flat.map(f => fpOf(f.r.act)), nBy = sets.map(st => st.sk.length), contrib = {};
  const col = (s, v, k) => {
    const key = s + '|' + v + '|' + k;
    let a = contrib[key];
    if (!a) { a = new Array(flat.length); for (let i = 0; i < flat.length; i++) a[i] = SW[s] * btPred_(flat[i].r, s, v, k, flat[i].st); contrib[key] = a; }
    return a;
  };
  const objective = (C, only) => {
    const arrs = SK_STATS.map(s => col(s, C[s].v, C[s].k)), err = sets.map(() => 0);
    for (let i = 0; i < flat.length; i++) {
      const si = flat[i].si;
      if (only != null && si !== only) continue;
      let p = 0; for (let j = 0; j < arrs.length; j++) p += arrs[j][i];
      err[si] += Math.abs(p - actFp[i]);
    }
    return only != null ? err[only] / nBy[only] : err.reduce((a, e, si) => a + e / nBy[si], 0) / sets.length;
  };
  const appNow = { g: { v: 'xg', k: APP.K.gx }, a: { v: 1, k: APP.K.a }, ppp: { v: 0, k: APP.K.ppp }, shp: { v: null, k: APP.K.shp },
                   sog: { v: 0, k: APP.K.sog }, hit: { v: null, k: APP.K.hit }, blk: { v: null, k: APP.K.blk } };
  let C = JSON.parse(JSON.stringify(appNow));
  const bestFor = {};
  const scan = s => {
    const out = {};
    (BT.OPT[s] || [null]).forEach(v => {
      let b = null;
      BT.K_GRID.forEach(k => { const m = objective(Object.assign({}, C, { [s]: { v: v, k: k } })); if (!b || m < b.m) b = { k: k, m: m }; });
      out[v] = b;
    });
    return out;
  };
  for (let pass = 0; pass < 4; pass++) {
    let moved = false;
    SK_STATS.forEach(s => {
      const vals = BT.OPT[s] || [null], bf = scan(s);
      let pick = vals[0];
      vals.forEach(v => { if (v !== vals[0] && bf[v].m < bf[pick].m - BT.MIN_GAIN_FP) pick = v; });
      if (C[s].v !== pick || C[s].k !== bf[pick].k) moved = true;
      C[s] = { v: pick, k: bf[pick].k };
    });
    if (!moved) break;
  }
  SK_STATS.forEach(s => { bestFor[s] = scan(s); });
  const fpBest = objective(C), fpApp = objective(appNow);
  // k optimal à chaque date, en FP (les autres stats au réglage final) → la confiance dans la saison change-t-elle ?
  SK_STATS.forEach(s => {
    kByCut[s] = sets.map((st, si) => {
      let b = null;
      BT.K_GRID.forEach(k => { const m = objective(Object.assign({}, C, { [s]: { v: C[s].v, k: k } }), si); if (!b || m < b.m) b = { k: k, m: m }; });
      return b.k;
    });
    tune[s].opt = C[s].v; tune[s].k = C[s].k;
  });
  const kxg = bestFor.g.xg.k;

  // 6) Méthodes comparées, en FP par match, à chaque date
  const cfg = (over) => Object.assign(JSON.parse(JSON.stringify(C)), over || {});
  const OFFC = s => { const off = (BT.OPT[s] || [null])[0]; return { [s]: { v: off, k: bestFor[s][off].k } }; };
  const predFp = (r, st, C, useGafBase) => {
    const rr = useGafBase ? Object.assign({}, r, { base: r.gafBase }) : r;
    return fpOf(Object.fromEntries(SK_STATS.map(s => [s, btPred_(rr, s, C[s].v, C[s].k, st)])));
  };
  const methods = [
    ['GAF seul (avant-saison)', (r) => fpOf(r.gafBase)],
    ['GAF + moteur (avant-saison)', (r) => fpOf(r.base)],
    ['Saison en cours seule', (r) => fpOf(Object.fromEntries(SK_STATS.map(k => [k, rate(r.early, k)])))],
    ['Maison — réglages actuels de l’app', (r, st) => predFp(r, st, appNow)],
    ['Maison optimisée', (r, st) => predFp(r, st, cfg())],
    ['  … base GAF seul (sans moteur)', (r, st) => predFp(r, st, cfg(), true)],
    ...['g', 'a', 'sog', 'ppp'].map(s => btToggle_(s, C, bestFor, cfg, OFFC, predFp))
  ];
  const corr = (a, b) => { const n = a.length, ma = a.reduce((s, x) => s + x, 0) / n, mb = b.reduce((s, x) => s + x, 0) / n;
    let c = 0, va = 0, vb = 0; for (let i = 0; i < n; i++) { c += (a[i] - ma) * (b[i] - mb); va += (a[i] - ma) ** 2; vb += (b[i] - mb) ** 2; }
    return c / Math.sqrt(va * vb); };
  const r3 = x => Math.round(x * 1000) / 1000;
  sets.forEach(st => st.sk.forEach(r => { r.restFp = fpOf(r.act); }));
  const methodRes = methods.map(([name, f]) => ({
    name: name,
    parDate: sets.map(st => {
      const pred = st.sk.map(r => f(r, st)), act = st.sk.map(r => r.restFp), err = pred.map((p, i) => p - act[i]);
      return { mae: r3(mae(err)), bias: round_(err.reduce((a, x) => a + x, 0) / err.length), corr: round_(corr(pred, act)),
               within05: round_(err.filter(x => Math.abs(x) <= 0.5).length / err.length) };
    })
  }));

  // 7) Gardiens : FP par match, base GAF + moteur (corrigée), k et biais communs aux 4 dates
  let goalies = null;
  if (sets.every(st => st.gk.length >= 5)) {
    const gPred = (r, k, b, base) => (k * (base + b) + r.n * r.early) / (k + r.n);
    const pooled = (k, b, useGaf) => sets.reduce((a, st) => a + mae(st.gk.map(r => gPred(r, k, b, useGaf ? r.gaf : r.prior) - r.rest)), 0) / sets.length;
    let best = null;
    BT.G_K.forEach(k => BT.G_B.forEach(b => { const m = pooled(k, b, false); if (!best || m < best.mae) best = { k: k, b: b, mae: m }; }));
    goalies = {
      n: sets.map(st => st.gk.length), meilleur: { k: best.k, biais: best.b, mae: r3(best.mae) },
      appActuel: r3(pooled(APP.K.gstart, APP.GOALIE_BIAS, true)),
      baseSeule: r3(sets.reduce((a, st) => a + mae(st.gk.map(r => r.prior - r.rest)), 0) / sets.length),
      saisonSeule: r3(sets.reduce((a, st) => a + mae(st.gk.map(r => r.early - r.rest)), 0) / sets.length),
      kParDate: sets.map(st => { let b = null; BT.G_K.forEach(k => { const m = mae(st.gk.map(r => gPred(r, k, best.b, r.prior) - r.rest)); if (!b || m < b.mae) b = { k: k, mae: m }; }); return b.k; })
    };
  }

  // 7b) Tendances : la classification (progression / chanceux / malchanceux / panne), faite avec ce qu'on savait à chaque date,
  //     annonce-t-elle quelque chose que la Maison ignore ? Et faut-il croire la saison plus vite pour certains joueurs (dont les jeunes) ?
  const tendances = btTrends_(sets, C, fpOf, rate, mae);
  // 7c) Optimisation finale (FP/match, moyenne des 4 dates) : poids analystes/moteur en cours de saison, k par position,
  //     k qui varie avec le nombre de matchs, et correction par catégorie de tendance. Chaque levier n'est gardé
  //     que s'il fait gagner au moins BT_FINAL.MIN ET s'il aide à au moins 3 des 4 dates.
  try { const AG = btAges_(); sets.forEach(st => st.sk.forEach(r => { r.age = AG[r.id] || null; })); } catch (e) { warn.push('Âges introuvables : ' + e.message); }
  const final = btFinal_(sets, C, gafW, SW, mae);
  // 7d) Gardiens : la part des 10 derniers matchs prédit-elle mieux les départs à venir ? 7e) Volatilité : fourchettes calibrées
  const gShare = btGoalieShare_(sets);
  const vol = btVolatility_(sets, C, SW);

  // 8) Plus gros ratés du modèle optimisé à la 1re date (pour comprendre)
  const st0 = sets[0], CF = cfg();
  const misses = st0.sk.map(r => ({ name: r.name, pos: r.pos, pred: round_(predFp(r, st0, CF)), real: round_(r.restFp), base: round_(fpOf(r.base)),
                                    early: round_(fpOf(Object.fromEntries(SK_STATS.map(k => [k, rate(r.early, k)])))), err: 0 }))
                      .map(m => Object.assign(m, { err: round_(m.pred - m.real) }))
                      .sort((a, b) => Math.abs(b.err) - Math.abs(a.err)).slice(0, 15);

  // 9) Réglages pour l'app
  const params = { K: Object.fromEntries(SK_STATS.map(s => [s, tune[s].k])), gsrc: tune.g.opt, kxg: kxg, kg: bestFor.g.g.k,
                   lam: tune.a.opt, rho: tune.sog.opt, ppb: tune.ppp.opt, ppMode: early.ppToi ? 'toi' : 'att',
                   r21: { F: round_(sets[sets.length - 1].r21.F), D: round_(sets[sets.length - 1].r21.D) },
                   gk: goalies ? { k: goalies.meilleur.k, b: goalies.meilleur.biais } : null, date: ymd_(new Date()),
                   final: final.retenu, gkRecent: gShare && gShare.retenu ? gShare.w : null, vol: vol ? vol.choisi : null };
  PropertiesService.getDocumentProperties().setProperty('BT_PARAMS', JSON.stringify(params));

  const res = {
    season: BT.SEASON, dates: sets.map(st => st.cut.label), cutoffs: sets.map(st => st.cut.d),
    nSkaters: sets.map(st => st.sk.length), nGames: early.snaps.map(s => s.nGames), withGaf: sets[0].sk.filter(r => gaf[btKey_(r.name)]).length,
    moteur: !!eng, warn: warn, params: params,
    objectif: { maisonOptimisee: Math.round(fpBest * 10000) / 10000, appActuelle: Math.round(fpApp * 10000) / 10000 },
    options: SK_STATS.map(s => {
      const vals = BT.OPT[s] || [null], f4 = x => Math.round(x * 10000) / 10000;
      return { stat: s, option: BT_OPT_NAME[s] || '', choisi: C[s].v, k: C[s].k, fp: f4(bestFor[s][C[s].v].m),
               sansOption: BT.OPT[s] ? f4(bestFor[s][vals[0]].m) : null,
               parValeur: BT.OPT[s] ? vals.map(v => ({ v: v, k: bestFor[s][v].k, fp: f4(bestFor[s][v].m) })) : null,
               maeStat: Math.round(tune[s].byOpt[C[s].v] ? tune[s].byOpt[C[s].v].mae * 10000 : 0) / 10000,
               kParDate: kByCut[s], kApp: s === 'g' ? APP.K.gx : APP.K[s] };
    }),
    methods: methodRes, goalies: goalies, tendances: tendances, final: final, gardiensRecents: gShare, volatilite: vol, misses: misses, seconds: Math.round((Date.now() - t0) / 1000)
  };
  btWrite_(ss, res);
  try { ss.toast(`Backtest terminé (${sets.length} dates, ${res.seconds} s). Voir la feuille « Backtest ».`, '🧪', 10); } catch (e) { /* idem */ }
  return res;
}

// Ligne « avec / sans » une option : si l'option est retenue, on la retire ; sinon, on l'ajoute avec son meilleur réglage
// (sans ça, « sans X » donnerait exactement la même erreur que le modèle quand X n'est déjà pas retenu)
const BT_OPT_TXT = { g: 'buts attendus maison (au lieu des buts réels)', a: 'passes secondaires pondérées', sog: 'tentatives de tir', ppp: 'rôle en AN' };
function btToggle_(s, C, bestFor, cfg, OFFC, predFp) {
  const off = (BT.OPT[s] || [null])[0];
  if (C[s].v !== off) return [`  … sans ${s === 'g' ? (C.g.v === 'xg' ? 'buts attendus maison' : 'tirs × % habituel') + ' (buts réels à la place)' : BT_OPT_TXT[s]}`, (r, st) => predFp(r, st, cfg(OFFC(s)))];
  let bv = null; Object.keys(bestFor[s]).forEach(v => { if (String(v) !== String(off) && (!bv || bestFor[s][v].m < bestFor[s][bv].m)) bv = v; });
  if (s === 'g' && bestFor.g.tirs && bestFor.g.xg) bv = bestFor.g.xg.m <= bestFor.g.tirs.m ? 'xg' : 'tirs';
  const val = isNaN(+bv) ? bv : +bv, k = bestFor[s][bv].k;
  const lab = s === 'g' ? (val === 'xg' ? 'buts attendus maison' : 'tirs × % habituel') + ' au lieu des buts réels' : BT_OPT_TXT[s] + ` (réglage ${val})`;
  return [`  … avec ${lab} — non retenu`, (r, st) => predFp(r, st, cfg({ [s]: { v: val, k: k } }))];
}

// ---------------- Optimisation finale ----------------
const BT_FINAL = { w: [0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1], gamma: [-0.4, -0.3, -0.2, -0.1, 0, 0.1, 0.2, 0.3, 0.4],
                   mult: [0.5, 0.75, 1, 1.5, 2], bias: [-0.3, -0.2, -0.15, -0.1, -0.05, 0, 0.05, 0.1, 0.15, 0.2, 0.3], MIN: 0.0005 };
const BT_CLS = ['progression', 'chanceux', 'malchanceux', 'panne'];
function btFinal_(sets, C, w0, SW, mae) {
  const K = BT.K_GRID, mixW = (g, m, w) => { const o = {}; SK_STATS.forEach(k => { o[k] = w * (g[k] || 0) + (1 - w) * (m[k] || 0); }); return o; };
  const baseFor = (r, w) => { if (w === w0 || !r.gaf || !r.pjr) return r; r._b = r._b || {}; return r._b[w] || (r._b[w] = Object.assign({}, r, { base: mixW(r.gaf, r.pjr, w) })); };
  const neutral = () => ({ w: w0, gamma: 0, kpos: { F: {}, D: {} }, mult: {}, bias: {}, alpha: 0, injM: 1, injB: 0, ageB: 0, wpos: {}, yB: 0, toiB: 0 });
  // jeunes en progression : base d'avant-saison × (1 + yB × (26 − âge) ÷ 4) sur les stats offensives (buts, passes, pts AN, tirs)
  const OFF = ['g', 'a', 'ppp', 'sog'];
  // courbe d'âge (les deux sens) : base offensive × (1 + yB × (27 − âge) ÷ 5) — jeunes relevés, vétérans baissés
  // rôle : temps de jeu par match cette saison ÷ celui de la saison passée, à la puissance toiB, sur la base d'avant-saison
  const toiFor = (r, tb) => { if (!tb || !r.ly || !r.ly.toi || !r.x || !r.x.toi || !r.n) return r; const f = Math.max(0.6, Math.min(1.6, Math.pow((r.x.toi / r.n / 60) / r.ly.toi, tb)));
    const b = Object.assign({}, r.base); ['g', 'a', 'sog', 'hit', 'blk'].forEach(k => { b[k] = (b[k] || 0) * f; }); return Object.assign({}, r, { base: b }); };
  const youngFor = (r, yB) => { if (!yB || !r.age) return r; const a = Math.max(19, Math.min(36, r.age)), f = Math.max(0.7, 1 + yB * (27 - a) / 5), b = Object.assign({}, r.base); OFF.forEach(k => { b[k] = (b[k] || 0) * f; }); return Object.assign({}, r, { base: b }); };
  const clone = X => JSON.parse(JSON.stringify(X));
  const predR = (r, st, X) => {
    const rr = toiFor(youngFor(baseFor(r, X.wpos && X.wpos[r.pos] != null ? X.wpos[r.pos] : X.w), X.yB), X.toiB), cm = r.cls && X.mult[r.cls] != null ? X.mult[r.cls] : 1;
    // âge : k × (âge ÷ 27)^ageB — ageB > 0 = on croit plus vite la saison d'un jeune, plus lentement celle d'un vétéran
    const gf = (X.gamma ? Math.pow(Math.max(1, r.n) / 15, X.gamma) : 1) * (X.alpha ? Math.pow(r.kr || 1, X.alpha) : 1) * (r.inj ? X.injM : 1)
             * (X.ageB && r.age ? Math.pow(Math.max(18, Math.min(40, r.age)) / 27, X.ageB) : 1);
    let fp = 0;
    SK_STATS.forEach(s => { const kp = X.kpos[r.pos][s]; fp += SW[s] * btPred_(rr, s, C[s].v, (kp != null ? kp : C[s].k) * cm * gf, st); });
    return fp + (r.cls && X.bias[r.cls] ? X.bias[r.cls] : 0) + (r.inj ? X.injB : 0);
  };
  // half : null = tous les joueurs, 0/1 = une moitié (selon l'identifiant) — pour régler sur une moitié et mesurer sur l'autre
  const perDate = (X, half) => sets.map(st => mae(st.sk.filter(r => half == null || r.id % 2 === half).map(r => predR(r, st, X) - r.restFp)));
  const obj = (X, half) => { const d = perDate(X, half); return d.reduce((t, x) => t + x, 0) / d.length; };
  const fit = (half, only) => {
    let X = neutral(), best = obj(X, half);
    const tryX = Y => { const o = obj(Y, half); if (o < best - 1e-7) { X = Y; best = o; return true; } return false; };
    const on = id => !only || only.indexOf(id) >= 0;
    for (let pass = 0; pass < 2; pass++) {
      let moved = false;
      if (on('w')) BT_FINAL.w.forEach(v => { if (v !== X.w) { const Y = clone(X); Y.w = v; moved = tryX(Y) || moved; } });
      if (on('gamma')) BT_FINAL.gamma.forEach(v => { if (v !== X.gamma) { const Y = clone(X); Y.gamma = v; moved = tryX(Y) || moved; } });
      if (on('alpha')) [-1, -0.5, 0.5, 1].forEach(v => { if (v !== X.alpha) { const Y = clone(X); Y.alpha = v; moved = tryX(Y) || moved; } });
      if (on('age')) [-3, -2, -1, -0.5, 0.5, 1, 2, 3].forEach(v => { if (v !== X.ageB) { const Y = clone(X); Y.ageB = v; moved = tryX(Y) || moved; } });
      if (on('toi')) [0.25, 0.5, 0.75, 1].forEach(v => { if (v !== X.toiB) { const Y = clone(X); Y.toiB = v; moved = tryX(Y) || moved; } });
      if (on('young')) [0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.08, 0.1].forEach(v => { if (v !== X.yB) { const Y = clone(X); Y.yB = v; moved = tryX(Y) || moved; } });
      if (on('wpos')) ['F', 'D'].forEach(pos => BT_FINAL.w.forEach(v => { const cur = X.wpos[pos] != null ? X.wpos[pos] : X.w; if (v !== cur) { const Y = clone(X); Y.wpos[pos] = v; moved = tryX(Y) || moved; } }));
      if (on('inj')) { BT_FINAL.mult.forEach(v => { if (v !== X.injM) { const Y = clone(X); Y.injM = v; moved = tryX(Y) || moved; } }); BT_FINAL.bias.forEach(v => { if (v !== X.injB) { const Y = clone(X); Y.injB = v; moved = tryX(Y) || moved; } }); }
      if (on('kpos')) ['F', 'D'].forEach(pos => SK_STATS.forEach(s => K.forEach(k => {
        const cur = X.kpos[pos][s] != null ? X.kpos[pos][s] : C[s].k;
        if (k !== cur) { const Y = clone(X); Y.kpos[pos][s] = k; moved = tryX(Y) || moved; }
      })));
      BT_CLS.forEach(c => {
        if (on('mult')) BT_FINAL.mult.forEach(v => { if (v !== (X.mult[c] || 1)) { const Y = clone(X); Y.mult[c] = v; moved = tryX(Y) || moved; } });
        if (on('bias')) BT_FINAL.bias.forEach(v => { if (v !== (X.bias[c] || 0)) { const Y = clone(X); Y.bias[c] = v; moved = tryX(Y) || moved; } });
      });
      if (!moved) break;
    }
    return X;
  };
  const levers = [
    { id: 'w', nom: 'Poids analystes / moteur en cours de saison', off: Y => { Y.w = w0; }, val: Y => `analystes ${Math.round(Y.w * 100)} % (avant : ${Math.round(w0 * 100)} %)` },
    { id: 'gamma', nom: 'k qui varie avec le nombre de matchs', off: Y => { Y.gamma = 0; }, val: Y => `k × (matchs ÷ 15)^${Y.gamma}` },
    { id: 'kpos', nom: 'k différent pour attaquants et défenseurs', off: Y => { Y.kpos = { F: {}, D: {} }; }, val: Y => ['F', 'D'].map(p => p + ' : ' + (Object.keys(Y.kpos[p]).map(s => s + ' ' + Y.kpos[p][s]).join(', ') || '—')).join(' · ') },
    { id: 'mult', nom: 'Vitesse de confiance selon la tendance', off: Y => { Y.mult = {}; }, val: Y => Object.keys(Y.mult).map(c => c + ' ×' + Y.mult[c]).join(', ') || '—' },
    { id: 'alpha', nom: 'k individuel (selon la variabilité de son niveau d’une saison à l’autre)', off: Y => { Y.alpha = 0; }, val: Y => `k × (variabilité ÷ typique)^${Y.alpha}` },
    { id: 'inj', nom: 'Joueur qui joue blessé (temps de jeu −20 % sur 3 matchs)', off: Y => { Y.injM = 1; Y.injB = 0; }, val: Y => `k ×${Y.injM}, ${Y.injB > 0 ? '+' : ''}${Y.injB} FP/match` },
    { id: 'age', nom: 'k selon l’âge (jeunes : on croit la saison plus vite)', off: Y => { Y.ageB = 0; }, val: Y => `k × (âge ÷ 27)^${Y.ageB}` },
    { id: 'toi', nom: 'Rôle : temps de jeu cette saison vs saison passée', off: Y => { Y.toiB = 0; }, val: Y => `base × (TJ cette saison ÷ TJ saison passée)^${Y.toiB}` },
    { id: 'young', nom: 'Courbe d’âge : jeunes relevés, vétérans baissés (base d’avant-saison)', off: Y => { Y.yB = 0; }, val: Y => `base offensive × (1 + ${Y.yB} × (27 − âge) ÷ 5)` },
    { id: 'wpos', nom: 'Poids analystes / moteur différent pour attaquants et défenseurs', off: Y => { Y.wpos = {}; }, val: Y => ['F', 'D'].map(p => p + ' : analystes ' + Math.round((Y.wpos[p] != null ? Y.wpos[p] : Y.w) * 100) + ' %').join(' · ') },
    { id: 'bias', nom: 'Correction selon la tendance (FP/match)', off: Y => { Y.bias = {}; }, val: Y => Object.keys(Y.bias).filter(c => Y.bias[c]).map(c => c + ' ' + (Y.bias[c] > 0 ? '+' : '') + Y.bias[c]).join(', ') || '—' }
  ];
  // Validation croisée : chaque levier est réglé (seul) sur une moitié des joueurs et mesuré sur l'autre moitié, dans les deux sens
  const rows = levers.map(L => {
    const g = [0, 1].map(h => { const X = fit(h, [L.id]); return obj(neutral(), 1 - h) - obj(X, 1 - h); });
    const Xall = fit(null, [L.id]), dOff = perDate(neutral()), dOn = perDate(Xall);
    const wins = dOff.filter((x, i) => x > dOn[i] + 1e-6).length;
    const cv = (g[0] + g[1]) / 2, ok = cv >= BT_FINAL.MIN && g[0] > 0 && g[1] > 0 && wins >= Math.min(3, sets.length);
    return { id: L.id, nom: L.nom, reglage: L.val(Xall), gainHorsEchantillon: Math.round(cv * 10000) / 10000,
             moities: g.map(x => Math.round(x * 10000) / 10000).join(' / '), datesGagnees: wins + '/' + sets.length, retenu: ok };
  });
  // Diagnostic : le modèle (sans levier) se trompe-t-il dans un sens pour certains groupes ? (prédit − réel, FP et points par match)
  const X0 = neutral(), groups = [['21 ans et moins', r => r.age && r.age < 22], ['22 à 24 ans', r => r.age >= 22 && r.age < 25], ['25 à 28 ans', r => r.age >= 25 && r.age < 29],
    ['29 à 32 ans', r => r.age >= 29 && r.age < 33], ['33 ans et plus', r => r.age >= 33], ['Top 50 prédits', null], ['Tous', r => true]];
  const biaisAge = groups.map(([nom, f]) => {
    let n = 0, dFp = 0, pP = 0, aP = 0;
    sets.forEach(st => {
      let L = st.sk;
      if (!f) L = L.slice().sort((a, b) => predR(b, st, X0) - predR(a, st, X0)).slice(0, 50); else L = L.filter(f);
      L.forEach(r => { n++; dFp += predR(r, st, X0) - r.restFp; pP += btPred_(r, 'g', C.g.v, C.g.k, st) + btPred_(r, 'a', C.a.v, C.a.k, st); aP += (r.act.g || 0) + (r.act.a || 0); });
    });
    return { groupe: nom, n: n, biaisFp: n ? Math.round(dFp / n * 1000) / 1000 : null, ptsPredits: n ? Math.round(pP / n * 1000) / 1000 : null, ptsReels: n ? Math.round(aP / n * 1000) / 1000 : null };
  });
  const kept = rows.filter(r => r.retenu).map(r => r.id);
  const F = kept.length ? fit(null, kept) : neutral();
  const b0 = perDate(neutral()), b1 = perDate(F), avg = d => d.reduce((t, x) => t + x, 0) / d.length;
  return { biaisAge: biaisAge, avant: Math.round(avg(b0) * 10000) / 10000, apres: Math.round(avg(b1) * 10000) / 10000,
           avantParDate: b0.map(x => Math.round(x * 1000) / 1000), apresParDate: b1.map(x => Math.round(x * 1000) / 1000), leviers: rows,
           retenu: kept.length ? { w: kept.indexOf('w') >= 0 ? F.w : null, gamma: F.gamma || 0, kpos: F.kpos, mult: F.mult, bias: F.bias, alpha: F.alpha || 0,
                                   inj: kept.indexOf('inj') >= 0 ? { m: F.injM, b: F.injB } : null,
                                   ageB: kept.indexOf('age') >= 0 ? F.ageB : 0, wpos: kept.indexOf('wpos') >= 0 ? F.wpos : null, yB: kept.indexOf('young') >= 0 ? F.yB : 0, toiB: kept.indexOf('toi') >= 0 ? F.toiB : 0 } : null };
}

// Joueur qui joue blessé : temps de jeu moyen des 3 derniers matchs ≤ 80 % de sa moyenne d'avant (au moins 8 matchs)
function btInjured_(x, n) {
  const L = x && x.last; if (!L || L.length < 3 || n < 8 || !x.toi) return false;
  const l3 = L.slice(-3).reduce((t, g) => t + g[1], 0), before = (x.toi - l3) / (n - 3);
  return before > 0 && l3 / 3 <= 0.8 * before;
}
// Gardiens : part des départs à venir (segment jusqu'à la date suivante) prédite par la saison, les 10 derniers matchs, ou un mélange
function btGoalieShare_(sets) {
  const rows = [];
  for (let i = 0; i + 1 < sets.length; i++) {
    const sn = sets[i].sn, nx = sets[i + 1].sn, cut = sets[i].cut.d;
    Object.keys(sn.ext).forEach(id => {
      const e = sn.ext[id], en = nx.ext[id]; if (!e || !e.gs || e.gs.length < 3 || !e.gteam) return;
      const td = (sn.tdates || {})[e.gteam] || [], tn = ((nx.tdates || {})[e.gteam] || []).filter(d => d > cut);
      if (td.length < 8 || tn.length < 5) return;
      const l10 = td.slice(-10), season = e.gs.length / td.length, recent = e.gs.filter(d => l10.indexOf(d) >= 0).length / l10.length;
      const target = ((en && en.gs) || []).filter(d => d > cut).length / tn.length;
      rows.push({ id: +id, season: Math.min(1, season), recent: Math.min(1, recent), target: Math.min(1, target), wt: tn.length });
    });
  }
  if (rows.length < 20) return null;
  const W = [0, 0.25, 0.5, 0.75, 1];
  const err = (w, half) => { let e = 0, n = 0; rows.forEach(r => { if (half != null && r.id % 2 !== half) return; e += r.wt * Math.abs(w * r.recent + (1 - w) * r.season - r.target); n += r.wt; }); return e / Math.max(1, n); };
  const best = half => W.reduce((b, w) => (err(w, half) < err(b, half) ? w : b), 0);
  const cv = [0, 1].map(h => err(0, 1 - h) - err(best(h), 1 - h));
  const wAll = best(null), gain = (cv[0] + cv[1]) / 2;
  return { n: rows.length, parPoids: W.map(w => ({ w: w, erreur: Math.round(err(w) * 1000) / 1000 })), w: wAll,
           gainHorsEchantillon: Math.round(gain * 1000) / 1000, retenu: wAll > 0 && gain >= 0.01 && cv[0] > 0 && cv[1] > 0 };
}
// Volatilité : fourchettes sur les matchs à venir (segment suivant). Variance par match = phi × Σ (poids² × taux) (loi de Poisson élargie)
// + incertitude du modèle (sm FP/match). On choisit phi et sm pour que 80 % (et 50 %) des résultats tombent dans les fourchettes.
function btVolatility_(sets, C, SW) {
  const obs = [];
  for (let i = 0; i + 1 < sets.length; i++) {
    const st = sets[i], nx = sets[i + 1].sn;
    st.sk.forEach(r => {
      const a = nx.map[r.id], e = st.sn.map[r.id]; if (!a || !e) return;
      const n = (a.gp || 0) - (e.gp || 0); if (n < 3) return;
      const actual = SK_STATS.reduce((t, k) => t + SW[k] * ((+a[k] || 0) - (+e[k] || 0)), 0);
      const rates = SK_STATS.map(k => btPred_(r, k, C[k].v, C[k].k, st));
      const mean = n * SK_STATS.reduce((t, k, j) => t + SW[k] * rates[j], 0), pois = SK_STATS.reduce((t, k, j) => t + SW[k] * SW[k] * Math.max(0, rates[j]), 0);
      obs.push({ n: n, actual: actual, mean: mean, pois: pois });
    });
  }
  if (obs.length < 50) return null;
  const cover = (phi, sm) => { let c80 = 0, c50 = 0; obs.forEach(o => { const sd = Math.sqrt(o.n * phi * o.pois + (sm * o.n) ** 2), z = Math.abs(o.actual - o.mean) / sd; if (z <= 1.2816) c80++; if (z <= 0.6745) c50++; });
    return { c80: c80 / obs.length, c50: c50 / obs.length }; };
  let best = null;
  [0.6, 0.8, 1, 1.2, 1.5, 1.8, 2.2, 2.6, 3].forEach(phi => [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6].forEach(sm => {
    const c = cover(phi, sm), e = Math.abs(c.c80 - 0.8) + Math.abs(c.c50 - 0.5);
    if (!best || e < best.e) best = { phi: phi, sm: sm, e: e, c: c };
  }));
  const naive = cover(1, 0);
  return { n: obs.length, choisi: { phi: best.phi, sm: best.sm }, couverture80: Math.round(best.c.c80 * 100), couverture50: Math.round(best.c.c50 * 100),
           naif80: Math.round(naive.c80 * 100), naif50: Math.round(naive.c50 * 100) };
}

// ---------------- Tendances (backtest) ----------------
const BT_TREND = { mult: [0.1, 0.25, 0.5, 0.75, 1, 1.5, 2, 3], MIN_GAIN: 0.002 };
// Classe un joueur comme l'app (onglet Tendances), avec seulement ce qu'on savait à la date :
// son jeu (temps de jeu, tirs, buts attendus) vs sa saison passée (et les analystes pour les tirs), ses FP vs les analystes, sa chance (buts − xG)
function btTrendOf_(r, xn, fpOf, rate) {
  const n = r.n, e = r.early, x = r.x, ly = r.ly, G = r.gaf;
  if (n < 5 || !G) return null;
  const parts = [], add = (now, base, w) => { if (now != null && base > 0) parts.push({ r: now / base, w: w }); };
  if (ly && x.toi) add(x.toi / n / 60, ly.toi, 0.25);
  const sb = ly && G.sog ? (ly.sog + G.sog) / 2 : ly ? ly.sog : G.sog;
  add((+e.sog || 0) / n, sb, 0.3);
  if (ly && ly.xg > 0 && xn > 0) add(r.ixg / n / xn, ly.xg, 0.3);
  if (!parts.length) return null;
  const ws = parts.reduce((t, q) => t + q.w, 0);
  const proc = Math.exp(parts.reduce((t, q) => t + q.w * Math.log(Math.max(0.2, Math.min(5, q.r))), 0) / ws) - 1;
  const fpE = fpOf(Object.fromEntries(SK_STATS.map(k => [k, rate(e, k)]))), fpG = fpOf(G);
  const perf = fpG > 0 ? fpE / fpG - 1 : 0, luck = CONFIG.SCORING.G * ((+e.g || 0) - r.ixg) / n;
  const z = Math.max(1, Math.min(1.6, 1.6 - (n - 5) * 0.06));
  let cls = 'conforme';
  if (perf >= 0.12 * z) cls = proc >= 0.08 * z ? 'progression' : 'chanceux';
  else if (perf <= -0.12 * z) cls = proc <= -0.08 * z ? 'panne' : 'malchanceux';
  else if (proc <= -0.15 * z && luck >= 0.4 * z) cls = 'panne';
  else if (proc >= 0.15 * z && luck <= -0.4 * z) cls = 'progression';
  else if (luck >= 0.7 * z) cls = 'chanceux';
  else if (luck <= -0.7 * z) cls = 'malchanceux';
  return { cls: cls, proc: proc, perf: perf, luck: luck };
}
function btTrends_(sets, C, fpOf, rate, mae) {
  const CL = ['progression', 'chanceux', 'malchanceux', 'panne', 'conforme'];
  const pred = (r, st, m) => fpOf(Object.fromEntries(SK_STATS.map(s => [s, btPred_(r, s, C[s].v, C[s].k * m, st)])));
  sets.forEach(st => {
    // buts attendus maison → échelle MoneyPuck (pour comparer à la saison passée)
    let a = 0, b = 0;
    st.sk.forEach(r => { if (r.ly && r.ly.xg > 0) { a += r.ixg; b += r.ly.xg * r.n; } });
    const xn = b ? a / b : 1;
    st.sk.forEach(r => { r.tr = btTrendOf_(r, xn, fpOf, rate); r.cls = r.tr ? r.tr.cls : null; r.young = r.age != null && r.age < 23; r.p0 = pred(r, st, 1); });
  });
  // 1) Écart réel − Maison, par catégorie (positif = la Maison sous-estime)
  const byClass = CL.map(c => {
    const par = sets.map(st => { const L = st.sk.filter(r => r.cls === c); return { n: L.length, ecart: L.length ? round_(L.reduce((t, r) => t + r.restFp - r.p0, 0) / L.length) : null }; });
    const all = []; sets.forEach(st => st.sk.forEach(r => { if (r.cls === c) all.push(r.restFp - r.p0); }));
    const m = all.length ? all.reduce((t, x) => t + x, 0) / all.length : null;
    const sd = all.length > 1 ? Math.sqrt(all.reduce((t, x) => t + (x - m) ** 2, 0) / (all.length - 1)) : null;
    const an = []; sets.forEach(st => st.sk.forEach(r => { if (r.cls === c && r.gaf) an.push(r.restFp - fpOf(r.gaf)); }));
    return { cls: c, n: all.length, ecart: round_(m), erreurType: sd && all.length ? round_(sd / Math.sqrt(all.length)) : null,
             auDessus: all.length ? round_(all.filter(x => x > 0).length / all.length) : null,
             ecartAnalystes: an.length ? round_(an.reduce((t, x) => t + x, 0) / an.length) : null, parDate: par };
  });
  const young = []; sets.forEach(st => st.sk.forEach(r => { if (r.young) young.push(r.restFp - r.p0); }));
  return { parCategorie: byClass, jeunes: { n: young.length, ecart: young.length ? round_(young.reduce((t, x) => t + x, 0) / young.length) : null } };
}

// Prédiction d'une stat par match pour le reste de la saison
// r = joueur, s = stat, v = valeur de l'option de cette stat, k = matchs de « base » mélangés aux matchs réels
function btPred_(r, s, v, k, st) {
  const n = r.n, e = r.early, x = r.x, b = r.base;
  let prior = b[s] || 0, obs = n ? (+e[s] || 0) / n : 0;
  if (s === 'g') {
    if (v === 'tirs') obs = n ? (+e.sog || 0) / n * (b.sog > 0 ? b.g / b.sog : (r.pos === 'D' ? 0.05 : 0.1)) : 0;
    else if (v === 'xg') obs = n ? r.ixg / n : 0;
  } else if (s === 'a' && v != null && v < 1 && x.a1 + x.a2 > 0) {
    obs = n ? Math.max(0, (+e.a || 0) - (1 - v) * (x.a2 - x.a1 * st.r21[r.pos])) / n : 0;
  } else if (s === 'sog' && v) {
    obs = n ? ((1 - v) * (+e.sog || 0) + v * x.att * r.sa) / n : 0;
  } else if (s === 'ppp' && v && r.prevPP) {
    const role = st.ppToi ? (x.pptoi / n + 20) / (r.prevPP.ppt + 20) : (x.ppatt / n + 0.2) / (r.prevPP.ppa + 0.2);
    prior *= Math.max(0.6, Math.min(1.6, Math.pow(role, v)));
  }
  return (k * prior + n * obs) / (k + n);
}
const btGfp_ = o => (o.gp ? (3 * (+o.w || 0) - (+o.ga || 0) + 0.2 * (+o.sv || 0) + 3 * (+o.so || 0) + (+o.otl || 0)) / o.gp : 0);

// Moteur maison pour 2025-26, à partir de 2022-23 à 2024-25 seulement
function btEngine_() {
  const P = pjParams_(), data = {};
  ['20222023', '20232024', '20242025'].forEach(s => { data[s] = pjSeason_(s); });
  pjAttachMp_(data, Date.now() + 90000);
  const sk = pjProjectSeason_(BT.SEASON, P, data);
  const gk = pjProjectGoalies_(BT.SEASON, P.gk || PJ.G_DEFAULT, data);
  const H = pjHistory_(BT.SEASON, data);
  const last = data['20242025'], lmp = last.mp || {}, prevPP = {};
  Object.keys(last.sk).forEach(id => {
    const h = last.sk[id], m = lmp[id];
    if (h.gp >= 20 && m) { const o = pjMpObj_(m); prevPP[id] = { ppa: o.ppatt / h.gp, ppt: o.pptoi / h.gp }; }
  });
  // saison passée de chaque joueur (tendances : son jeu cette saison vs l'an passé) + âge
  const ly = {}, age = {};
  Object.keys(last.sk).forEach(id => {
    const h = last.sk[id], m = lmp[id];
    if (h.gp >= 15) ly[id] = { toi: (h.toi || 0) / 60, sog: h.sog / h.gp, xg: m ? pjMpObj_(m).xg / h.gp : null, onxg: m ? pjMpObj_(m).onxg / h.gp : null };
  });
  let cur = null; try { cur = pjSeason_(BT.SEASON); } catch (e) { /* rien */ }
  [cur].concat(Object.keys(data).map(k => data[k])).forEach(d => { if (d) Object.keys(d.sk).forEach(id => { if (!age[id] && d.sk[id].b) age[id] = pjAge_(d.sk[id].b, BT.SEASON); }); });
  // variabilité de son niveau d'une saison à l'autre (FP/match, saisons de 20 MJ et plus) ÷ celle d'un joueur typique
  const cvs = {}, kr = {};
  Object.keys(last.sk).forEach(id => {
    const v = Object.keys(data).map(k => data[k].sk[id]).filter(h => h && h.gp >= 20).map(h => pjActualFpg_(h));
    if (v.length >= 2) { const m = v.reduce((t, x) => t + x, 0) / v.length; if (m > 0.5) cvs[id] = Math.sqrt(v.reduce((t, x) => t + (x - m) ** 2, 0) / (v.length - 1)) / m; }
  });
  const med = Object.values(cvs).sort((a, b) => a - b)[Math.floor(Object.keys(cvs).length / 2)] || 0.1;
  Object.keys(cvs).forEach(id => { kr[id] = Math.max(0.33, Math.min(3, cvs[id] / med)); });
  return { P: P, sk: sk, gk: gk, lg: H ? H.lg : null, prevPP: prevPP, ly: ly, age: age, kr: kr };
}

// GAF 2025-26 (feuille publique de Crome) → stats projetées par match
// Âge de chaque patineur au 1er octobre 2025 (API stats LNH, fiches) — gardé 7 jours
function btAges_() {
  const hit = pjCacheGet_('bt_ages_20252026'); if (hit) return hit;
  const exp = encodeURIComponent('seasonId=20252026 and gameTypeId=2');
  const r = UrlFetchApp.fetch(`https://api.nhle.com/stats/rest/en/skater/bios?isAggregate=false&isGame=false&start=0&limit=-1&cayenneExp=${exp}`, { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('code ' + r.getResponseCode());
  const out = {}, ref = new Date('2025-10-01T12:00:00Z').getTime();
  (JSON.parse(r.getContentText()).data || []).forEach(x => { if (x.playerId && x.birthDate) out[x.playerId] = Math.round((ref - new Date(x.birthDate + 'T12:00:00Z').getTime()) / 3.15576e10 * 10) / 10; });
  pjCachePut_('bt_ages_20252026', out);
  return out;
}
function btGaf_() {
  const r = UrlFetchApp.fetch(BT.GAF_URL, { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Projections GAF 2025-26 introuvables (code ' + r.getResponseCode() + ')');
  const v = Utilities.parseCsv(r.getContentText());
  const h = v[0].map(x => String(x).trim()), i = n => h.indexOf(n);
  const I = { name: i('PLAYER'), pos: i('POS'), gp: i('GP'), g: i('G'), a: i('A'), ppp: i('PPP'), shp: i('SHP'), sog: i('SOG'),
              hit: i('HIT'), blk: i('BLK'), w: i('W'), ga: i('GA'), sv: i('SV'), so: i('SO'), otl: i('OTL') };
  const out = {};
  for (let k = 1; k < v.length; k++) {
    const row = v[k], name = row[I.name], gp = +row[I.gp] || 0;
    if (!name || !gp) continue;
    const per = s => (I[s] >= 0 ? (+row[I[s]] || 0) / gp : 0);
    const o = { g: per('g'), a: per('a'), ppp: per('ppp'), shp: per('shp'), sog: per('sog'), hit: per('hit'), blk: per('blk') };
    if (String(row[I.pos]) === 'G') o.gfpg = 3 * per('w') - per('ga') + 0.2 * per('sv') + 3 * per('so') + per('otl');
    out[btKey_(name)] = o;
  }
  return out;
}
// Noms GAF ≠ noms LNH : « Elias Pettersson (F) », « Mack Celebrini »…
const BT_ALIAS = { mackcelebrini: 'macklincelebrini', matthewboldy: 'mattboldy', mitchmarner: 'mitchellmarner' };
function btKey_(name) {
  const k = norm_(String(name).replace(/\s*\(.*?\)\s*/g, ' '));
  return BT_ALIAS[k] || k;
}

// ---------------- Lecture des matchs, avec un instantané à chaque date ----------------
const btBlankExt_ = () => ({ a1: 0, a2: 0, att: 0, ppatt: 0, pptoi: 0, toi: 0, b: new Array(12).fill(0), gb: new Array(12).fill(0) });
function btEarlySnapshots_(t0) {
  const last = BT.CUTOFFS[BT.CUTOFFS.length - 1].d, games = [];
  let d = BT.START, guard = 0;
  while (d && d <= last && guard++ < 40) {
    const j = fetchJsonAll_([`${WEB}/schedule/${d}`])[0];
    (j.gameWeek || []).forEach(day => (day.games || []).forEach(g => {
      if (String(g.season) === BT.SEASON && g.gameType === 2 && day.date <= last) games.push({ id: g.id, date: day.date });
    }));
    d = j.nextStartDate;
  }
  games.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.id - b.id));
  let map = {}, ext = {}, tdates = {};   // tdates : équipe → dates de ses matchs (gardiens : part des 10 derniers matchs)
  const snaps = [], skipped = [], seen = {}, gdate = {};
  games.forEach(g => { gdate[g.id] = g.date; });
  const flag = { ppToi: false };
  let i = 0, stop = false;
  // Instantanés déjà calculés (les matchs de 2025-26 ne changent plus) : on repart du dernier
  const cacheKey = c => 'bt3_' + c.d;   // bt3 : temps de jeu, 6 derniers matchs, départs des gardiens
  for (let j = 0; j < BT.CUTOFFS.length; j++) {
    const c = BT.CUTOFFS[j], got = typeof pjCacheGet_ === 'function' ? pjCacheGet_(cacheKey(c)) : null;
    if (!got) break;
    snaps.push({ cut: c, map: got.map, ext: got.ext, nGames: got.nGames, tdates: got.tdates || {} });
    if (got.ppToi) flag.ppToi = true;
  }
  if (snaps.length) {
    const lastSnap = snaps[snaps.length - 1];
    map = JSON.parse(JSON.stringify(lastSnap.map)); ext = JSON.parse(JSON.stringify(lastSnap.ext)); tdates = JSON.parse(JSON.stringify(lastSnap.tdates || {}));
    while (i < games.length && games[i].date <= lastSnap.cut.d) { seen[games[i].id] = true; i++; }
  }
  BT.CUTOFFS.slice(snaps.length).forEach(c => {
    if (stop) { skipped.push(c.label); return; }
    const seg = [];
    while (i < games.length && games[i].date <= c.d) { if (!seen[games[i].id]) { seen[games[i].id] = true; seg.push(games[i].id); } i++; }
    for (let k = 0; k < seg.length; k += 20) {
      if (Date.now() - t0 > BT.READ_LIMIT) { stop = true; break; }
      const batch = seg.slice(k, k + 20), urls = [];
      batch.forEach(id => { urls.push(`${WEB}/gamecenter/${id}/boxscore`, `${WEB}/gamecenter/${id}/play-by-play`); });
      const data = fetchJsonAll_(urls);
      batch.forEach((id, j) => {
        parseGame_(data[2 * j], data[2 * j + 1]).forEach(l => addLine_(map, l));
        btExtras_(data[2 * j], data[2 * j + 1], ext, flag, gdate[id], tdates);
      });
    }
    if (stop) { skipped.push(c.label); return; }
    const snap = { cut: c, map: JSON.parse(JSON.stringify(map)), ext: JSON.parse(JSON.stringify(ext)), nGames: Object.keys(seen).length, tdates: JSON.parse(JSON.stringify(tdates)) };
    snaps.push(snap);
    if (typeof pjCachePut_ === 'function') pjCachePut_(cacheKey(c), { map: snap.map, ext: snap.ext, nGames: snap.nGames, ppToi: flag.ppToi, tdates: snap.tdates });
  });
  return { snaps: snaps, skipped: skipped, ppToi: flag.ppToi };
}
// Stats « avancées » d'un match à partir du déroulement : passes 1re/2e, tentatives, tentatives en AN,
// tirs non bloqués par case (distance × rebond) pour les buts attendus maison, temps en AN (si fourni)
function btExtras_(box, pbp, ext, flag, date, tdates) {
  const E = id => ext[id] || (ext[id] = btBlankExt_());
  if (date && tdates) [box.awayTeam.abbrev, box.homeTeam.abbrev].forEach(t => { (tdates[t] = tdates[t] || []).push(date); });
  const sec = t => { const m = String(t || '').split(':'); return m.length === 2 ? (+m[0]) * 60 + (+m[1]) : 0; };
  ['awayTeam', 'homeTeam'].forEach(side => {
    const st = (box.playerByGameStats || {})[side] || {};
    ['forwards', 'defense'].forEach(g => (st[g] || []).forEach(p => {
      if (p.powerPlayToi != null && p.powerPlayToi !== '') { flag.ppToi = true; E(p.playerId).pptoi += sec(p.powerPlayToi); }
      if (p.toi && p.toi !== '00:00') {
        const e = E(p.playerId); e.toi += sec(p.toi);   // temps de jeu (tendances)
        e.last = (e.last || []).concat([[date || '', sec(p.toi), p.hits || 0, p.blockedShots || 0]]).slice(-6);   // 6 derniers matchs
      }
    }));
    // gardien partant (ou, à défaut, celui qui a joué le plus) : date du départ
    const gl = (st.goalies || []).filter(g => g.toi && g.toi !== '00:00');
    const starter = gl.find(g => g.starter === true) || gl.slice().sort((a, b) => sec(b.toi) - sec(a.toi))[0];
    if (starter && date) { const e = E(starter.playerId); e.gs = (e.gs || []).concat([date]); e.gteam = box[side].abbrev; }
  });
  const teamOf = {};
  (pbp.rosterSpots || []).forEach(r => { teamOf[r.playerId] = r.teamId; });
  const awayId = box.awayTeam.id;
  let lastSave = null;
  (pbp.plays || []).forEach(pl => {
    const type = pl.typeDescKey;
    if (type !== 'goal' && type !== 'shot-on-goal' && type !== 'missed-shot' && type !== 'blocked-shot') return;
    const pd = pl.periodDescriptor || {};
    if (pd.periodType === 'SO') return;
    const d = pl.details || {};
    if (type === 'goal') {
      if (d.assist1PlayerId) E(d.assist1PlayerId).a1++;
      if (d.assist2PlayerId) E(d.assist2PlayerId).a2++;
    }
    const shooter = type === 'goal' ? d.scoringPlayerId : d.shootingPlayerId, team = teamOf[shooter];
    if (!shooter || !team) return;
    const sc = String(pl.situationCode || ''), ok = sc.length === 4, isAway = team === awayId;
    const ag = +sc[0], as = +sc[1], hs = +sc[2], hg = +sc[3];
    const awayEff = as - (ag === 0 ? 1 : 0), homeEff = hs - (hg === 0 ? 1 : 0);
    const e = E(shooter);
    e.att++;
    if (ok && (isAway ? awayEff > homeEff : homeEff > awayEff)) e.ppatt++;
    const t = sec(pl.timeInPeriod), per = pd.number;
    const reb = !!(lastSave && lastSave.team === team && lastSave.per === per && t >= lastSave.t && t - lastSave.t <= 3);
    if (type === 'shot-on-goal') lastSave = { team: team, per: per, t: t };
    if (type === 'blocked-shot' || !ok || (isAway ? hg : ag) === 0) return;   // bloqué ou filet désert : hors buts attendus
    const x = +d.xCoord, y = +d.yCoord;
    if (isNaN(x) || isNaN(y)) return;
    const dist = d.zoneCode === 'O' ? Math.sqrt(Math.pow(89 - Math.abs(x), 2) + y * y) : 99;
    const di = dist < 10 ? 0 : dist < 20 ? 1 : dist < 30 ? 2 : dist < 40 ? 3 : dist < 55 ? 4 : 5;
    const bin = di * 2 + (reb ? 1 : 0);
    e.b[bin]++;
    if (type === 'goal') e.gb[bin]++;
  });
}

// ---------------- Feuille « Backtest » ----------------
function btWrite_(ss, res) {
  const sh = ss.getSheetByName(BT.SHEET) || ss.insertSheet(BT.SHEET);
  sh.clear();   // contenu + formats (sinon Sheets transforme « 28 oct. » en date et les erreurs en heures)
  const W = 2 + res.dates.length * 2 > 8 ? 2 + res.dates.length * 2 : 8, out = [];
  const row = a => { const r = a.slice(0, W); while (r.length < W) r.push(''); out.push(r); };
  const fmtV = (s, v) => (s === 'g' ? BT_G_LABEL[v] : v == null ? '' : v);
  row([`🧪 Backtest ${res.season} — prédire le reste de la saison à partir des stats jusqu'à une date`]);
  row([`Dates : ${res.dates.map((d, i) => `${d} (${res.nGames[i]} matchs, ${res.nSkaters[i]} patineurs)`).join(' · ')} · base : GAF ${res.moteur ? '+ moteur maison' : 'seul'} · ${res.seconds} s`]);
  res.warn.forEach(w => row(['⚠ ' + w]));
  row(['Note : les réglages du moteur et le poids de GAF ont été appris en partie sur 2025-26 ; l’avantage de « GAF + moteur » est donc un peu flatté.']);
  row([]);
  row(['Méthode — erreur moyenne (FP/match)'].concat(res.dates).concat(['Corrélation (1re date)', '% à ±0,5 (1re date)']));
  res.methods.forEach(m => row([m.name].concat(m.parDate.map(x => x.mae)).concat([m.parDate[0].corr, m.parDate[0].within05])));
  row([]);
  row([`Réglage final sur l’erreur en FP/match (moyenne des dates) : ${res.objectif.maisonOptimisee} (réglages actuels de l’app : ${res.objectif.appActuelle})`]);
  row(['Stat', 'Option testée', 'Retenu', 'k retenu', 'Erreur FP/match', 'Sans l’option', 'k actuel de l’app', 'Détail par valeur (k → erreur FP)']);
  res.options.forEach(o => row([o.stat, o.option, fmtV(o.stat, o.choisi), o.k, o.fp, o.sansOption == null ? '' : o.sansOption, o.kApp,
    o.parValeur ? o.parValeur.map(p => `${fmtV(o.stat, p.v)} : k ${p.k} → ${p.fp}`).join(' | ') : '']));
  row([]);
  row(['k optimal à chaque date (la confiance dans la saison change-t-elle ?)'].concat(res.dates));
  res.options.forEach(o => row([o.stat].concat(o.kParDate)));
  if (res.goalies) {
    const g = res.goalies;
    row(['gardiens (FP/match)'].concat(g.kParDate));
    row([]);
    row(['Gardiens (FP par match)', 'Gardiens par date', 'Base seule', 'Saison seule', 'App actuelle (GAF, k 100, −0,5)', 'Meilleur k', 'Meilleur biais', 'Erreur meilleure']);
    row(['', g.n.join(' / '), g.baseSeule, g.saisonSeule, g.appActuel, g.meilleur.k, g.meilleur.biais, g.meilleur.mae]);
  }
  if (res.tendances) {
    const T = res.tendances, LBL = { progression: '🟢 Vraie progression', chanceux: '🔵 Chanceux', malchanceux: '🟡 Malchanceux', panne: '🔴 En panne', conforme: 'Conforme' };
    row([]);
    row(['Tendances : la catégorie annonce-t-elle quelque chose que la Maison ignore ?']);
    row(['Catégorie (classée à chaque date)', 'Joueurs (4 dates)', 'Écart réel − Maison (FP/match)', '± erreur type', '% au-dessus de la Maison', 'Écart réel − analystes', 'Écart par date']);
    T.parCategorie.forEach(c => row([LBL[c.cls], c.n, c.ecart, c.erreurType, c.auDessus, c.ecartAnalystes, c.parDate.map(d => d.ecart == null ? '–' : d.ecart).join(' / ')]));
    row(['Jeunes (moins de 23 ans)', T.jeunes.n, T.jeunes.ecart]);
  }
  if (res.final) {
    const F = res.final;
    row([]);
    row([`Optimisation finale : erreur ${F.avant} → ${F.apres} FP/match (moyenne des dates) · par date : ${F.avantParDate.join(' / ')} → ${F.apresParDate.join(' / ')}`]);
    row(['Levier', 'Réglage trouvé (tous les joueurs)', 'Gain hors échantillon (FP/match)', 'Gain sur chaque moitié', 'Dates où il aide', 'Retenu ?']);
    F.leviers.forEach(l => row([l.nom, l.reglage, l.gainHorsEchantillon, l.moities, l.datesGagnees, l.retenu ? 'oui' : 'non']));
    row(['Gain hors échantillon = réglé sur la moitié des joueurs, mesuré sur l’autre moitié (et inversement) : le vrai gain attendu sur des joueurs que le modèle n’a jamais vus.']);
  }
  if (res.gardiensRecents) {
    const g = res.gardiensRecents;
    row([]);
    row([`Gardiens : prédire la part des départs à venir (${g.n} cas) — 0 = part de la saison, 1 = part des 10 derniers matchs`]);
    row(['Poids des 10 derniers matchs'].concat(g.parPoids.map(x => x.w)));
    row(['Erreur moyenne (part des départs)'].concat(g.parPoids.map(x => x.erreur)));
    row([`Gain hors échantillon : ${g.gainHorsEchantillon} — ${g.retenu ? 'RETENU (poids ' + g.w + ')' : 'non retenu'}`]);
  }
  if (res.volatilite) {
    const v = res.volatilite;
    row([]);
    row([`Volatilité (${v.n} cas) : fourchette à 80 % → ${v.couverture80} % des résultats dedans, à 50 % → ${v.couverture50} % (sans calibration : ${v.naif80} % et ${v.naif50} %). Réglage : phi ${v.choisi.phi}, incertitude du modèle ${v.choisi.sm} FP/match`]);
  }
  row([]);
  row([`Plus gros ratés à la 1re date (${res.dates[0]})`, 'Pos', 'Prédit', 'Réel (reste de saison)', 'Écart', 'Base d’avant-saison', 'Saison jusque-là']);
  res.misses.forEach(m => row([m.name, m.pos, m.pred, m.real, m.err, m.base, m.early]));
  const txt = out.map(r => r.map(v => (typeof v === 'string' && /^\d{1,2}(er)? [a-zéû]+\.?$/i.test(v) ? "'" + v : v)));
  sh.getRange(1, 1, out.length, W).setNumberFormat('General').setValues(txt);
  const json = JSON.stringify(res), parts = [];
  for (let i = 0; i < json.length; i += 45000) parts.push(json.slice(i, i + 45000));
  const dsh = hiddenSheet_(ss, BT.DATA_SHEET);
  dsh.clearContents();
  dsh.getRange(1, 1, 1, parts.length).setValues([parts]);
}

function getBacktestData() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BT.DATA_SHEET);
  return sh && sh.getLastRow() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].join('') : '';
}
function btParams_() {
  const s = PropertiesService.getDocumentProperties().getProperty('BT_PARAMS');
  return s ? JSON.parse(s) : null;
}

// =====================================================================================
// 🧪 Backtest PAR MATCH (adversaire du soir) — la projection d'un match doit-elle tenir compte de l'adversaire ?
// Pour chaque match de 2025-26 (à partir du 15 nov.), prédiction de base = moyenne FP/match du joueur AVANT ce match
// (ramenée vers la moyenne de sa position), puis on teste des multiplicateurs :
//   défense adverse (buts et tirs accordés), rythme (tirs pour + contre), discipline adverse (minutes de pénalité),
//   repos (son équipe / l'adversaire au 2e de deux soirs), domicile, gardien adverse (auxiliaire ? % d'arrêts).
// Validation croisée par moitiés de joueurs (comme le reste) : un facteur n'est retenu que s'il aide sur les DEUX moitiés.
// Lecture des matchs reprise là où elle s'est arrêtée (relancer si « incomplet »).
// =====================================================================================
const MU = {
  SEASON: '20252026', START: '2025-10-07', END: '2026-04-17', EVAL_FROM: '2025-11-15',
  READ_LIMIT: 230000, CHUNK: 150, SHEET: 'Backtest matchs', KBASE: 10, MIN_N: 10, MIN_GAIN: 0.005,
  W: { g: 3, a: 2, ppp: 0.5, shp: 0.5, sog: 0.2, hit: 0.2, blk: 0.4 },
  GRID: [-1, -0.8, -0.6, -0.4, -0.3, -0.2, -0.15, -0.1, -0.05, 0, 0.05, 0.1, 0.15, 0.2, 0.3, 0.4, 0.6, 0.8, 1]
};
const MU_F = [   // facteurs testés : [clé, libellé, s'applique à « off » (buts/passes/tirs) ou « def » (mises en échec/tirs bloqués)]
  ['ga', 'Défense adverse : buts accordés', 'off'], ['sa', 'Défense adverse : tirs accordés', 'off'],
  ['pace', 'Rythme des matchs de l’adversaire', 'both'], ['pim', 'Discipline adverse (pénalités)', 'off'],
  ['b2b', 'Son équipe au 2e de deux soirs', 'both'], ['ob2b', 'Adversaire au 2e de deux soirs', 'both'],
  ['home', 'À domicile', 'both'], ['bk', 'Gardien adverse auxiliaire', 'off'], ['gsv', 'Gardien adverse : % d’arrêts faible', 'off'],
  ['poss', 'Adversaire qui tire beaucoup (plus de jeu défensif)', 'def']
];
function runMatchupBacktest() {
  const t0 = Date.now(), ss = SpreadsheetApp.getActiveSpreadsheet();
  // 1) liste des matchs
  let games = pjCacheGet_('mu_games_' + MU.SEASON);
  if (!games) {
    games = []; let d = MU.START, guard = 0;
    while (d && d <= MU.END && guard++ < 40) {
      const j = fetchJsonAll_([`${WEB}/schedule/${d}`])[0];
      (j.gameWeek || []).forEach(day => (day.games || []).forEach(g => {
        if (String(g.season) === MU.SEASON && g.gameType === 2 && day.date <= MU.END) games.push([g.id, day.date]);
      }));
      d = j.nextStartDate;
    }
    games.sort((a, b) => (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : a[0] - b[0]));
    pjCachePut_('mu_games_' + MU.SEASON, games);
  }
  // 2) matchs lus par paquets de 150 (gardés en cache) : [id, date, équipes, patineurs, gardiens]
  const chunks = [], nCh = Math.ceil(games.length / MU.CHUNK);
  let incomplete = false;
  for (let c = 0; c < nCh; c++) {
    let got = pjCacheGet_('mu1_' + c);
    if (!got) {
      if (Date.now() - t0 > MU.READ_LIMIT) { incomplete = true; break; }
      const part = games.slice(c * MU.CHUNK, (c + 1) * MU.CHUNK), recs = [];
      for (let k = 0; k < part.length; k += 20) {
        const b = part.slice(k, k + 20), urls = [];
        b.forEach(g => urls.push(`${WEB}/gamecenter/${g[0]}/boxscore`, `${WEB}/gamecenter/${g[0]}/play-by-play`));
        const data = fetchJsonAll_(urls);
        b.forEach((g, j) => { try { recs.push(muGame_(g, data[2 * j], data[2 * j + 1])); } catch (e) { /* match illisible */ } });
      }
      got = recs;
      if (part.length && part[part.length - 1][1] < ymd_(new Date())) pjCachePut_('mu1_' + c, got);
    }
    chunks.push(got);
  }
  if (incomplete) {
    const done = chunks.reduce((s, x) => s + x.length, 0);
    muWrite_(ss, { incomplete: true, done: done, total: games.length, seconds: Math.round((Date.now() - t0) / 1000) });
    return;
  }
  const recs = [].concat.apply([], chunks);
  const res = muEvaluate_(recs);
  res.seconds = Math.round((Date.now() - t0) / 1000);
  muWrite_(ss, res);
  PropertiesService.getDocumentProperties().setProperty('BT_MATCHUP', JSON.stringify({ date: ymd_(new Date()), kept: res.kept, coef: res.coefKept }));
}
// Un match compact : [id, date, [awayId, homeId], [[pid, side(0/1), grp(F/D), fp_off, fp_def]...], team: [[sog, pim, ga]×2], goalies: [[pid, side, starter, sa, ga]...]]
function muGame_(g, box, pbp) {
  const lines = parseGame_(box, pbp), sideOf = {}, grpOf = {};
  const st = box.playerByGameStats || {}, T = [[0, 0, 0], [0, 0, 0]], G = [];
  ['awayTeam', 'homeTeam'].forEach((side, si) => {
    const s = st[side] || {};
    ['forwards', 'defense'].forEach(grp => (s[grp] || []).forEach(p => { sideOf[p.playerId] = si; grpOf[p.playerId] = grp === 'defense' ? 'D' : 'F'; T[si][1] += p.pim || 0; }));
    T[si][0] = (box[side] && box[side].sog) || 0;
    T[si][2] = box[si ? 'awayTeam' : 'homeTeam'] ? box[si ? 'awayTeam' : 'homeTeam'].score || 0 : 0;   // buts accordés
    (s.goalies || []).forEach(p => { if (p.toi && p.toi !== '00:00') G.push([p.playerId, si, p.starter ? 1 : 0, p.shotsAgainst || 0, p.goalsAgainst || 0]); });
  });
  const W = MU.W, P = [];
  lines.forEach(l => {
    if (l.pos === 'G' || sideOf[l.id] == null) return;
    const off = W.g * l.g + W.a * l.a + W.ppp * l.ppp + W.shp * l.shp + W.sog * l.sog, def = W.hit * l.hit + W.blk * l.blk;
    P.push([l.id, sideOf[l.id], grpOf[l.id], Math.round(off * 100) / 100, Math.round(def * 100) / 100]);
  });
  return [g[0], g[1], [box.awayTeam.id, box.homeTeam.id], P, T, G];
}
function muEvaluate_(recs) {
  recs.sort((a, b) => (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : a[0] - b[0]));
  const team = {}, pl = {}, gk = {}, lg = { gp: 0, sog: 0, pim: 0, ga: 0, sa: 0, sv: 0, off: { F: [0, 0], D: [0, 0] }, def: { F: [0, 0], D: [0, 0] } };
  const tm = id => team[id] || (team[id] = { gp: 0, sf: 0, sa: 0, ga: 0, pim: 0, last: null, starts: {} });
  const rows = [];   // [pid, base_off, base_def, features..., actual_off, actual_def]
  recs.forEach(r => {
    const [id, date, ids, P, T, G] = r;
    const eval_ = date >= MU.EVAL_FROM;
    // --- caractéristiques AVANT le match ---
    const feat = [0, 1].map(si => {
      const me = tm(ids[si]), op = tm(ids[1 - si]), L = lg.gp ? lg : null;
      if (!L || op.gp < 5) return null;
      const lgGA = lg.ga / lg.gp, lgS = lg.sog / lg.gp, lgPim = lg.pim / lg.gp, K = 10;
      const sh = (x, n, m) => (x + m * K) / (n + K) / m - 1;
      const opG = G.filter(q => q[1] === 1 - si && q[2] === 1)[0];
      let bk = 0, gsv = 0;
      if (opG) {
        const ss = op.starts, lead = Object.keys(ss).sort((a, b) => ss[b] - ss[a])[0];
        bk = lead && +lead !== opG[0] && (ss[lead] || 0) >= 5 ? 1 : 0;
        const g = gk[opG[0]] || { sa: 0, ga: 0 }, lsv = lg.sa ? 1 - lg.gaG / lg.sa : 0.9;
        const sv = (g.sa - g.ga + lsv * 300) / (g.sa + 300);
        gsv = (lsv - sv) * 100;   // + = gardien plus faible que la moyenne (en points de %)
      }
      const yest = d => { const x = new Date(d + 'T12:00:00Z'); x.setUTCDate(x.getUTCDate() - 1); return x.toISOString().slice(0, 10); };
      return { ga: sh(op.ga, op.gp, lgGA), sa: sh(op.sa, op.gp, lgS), pace: sh(op.sf + op.sa, op.gp, 2 * lgS), pim: sh(op.pim, op.gp, lgPim),
               b2b: me.last === yest(date) ? 1 : 0, ob2b: op.last === yest(date) ? 1 : 0, home: si, bk: bk, gsv: gsv, poss: sh(op.sf, op.gp, lgS) };
    });
    if (eval_) P.forEach(([pid, si, grp, off, def]) => {
      const f = feat[si], x = pl[pid];
      if (!f || !x || x.n < MU.MIN_N) return;
      const mo = lg.off[grp][1] ? lg.off[grp][0] / lg.off[grp][1] : 0, md = lg.def[grp][1] ? lg.def[grp][0] / lg.def[grp][1] : 0;
      rows.push({ pid: pid, bo: (x.off + mo * MU.KBASE) / (x.n + MU.KBASE), bd: (x.def + md * MU.KBASE) / (x.n + MU.KBASE), f: f, ao: off, ad: def });
    });
    // --- mise à jour APRÈS le match ---
    [0, 1].forEach(si => {
      const me = tm(ids[si]);
      me.gp++; me.sf += T[si][0]; me.sa += T[1 - si][0]; me.ga += T[si][2]; me.pim += T[si][1]; me.last = date;
      lg.gp++; lg.sog += T[si][0]; lg.pim += T[si][1]; lg.ga += T[si][2];
    });
    G.forEach(([pid, si, starter, sa, ga]) => {
      const g = gk[pid] || (gk[pid] = { sa: 0, ga: 0 }); g.sa += sa; g.ga += ga; lg.sa += sa; lg.gaG = (lg.gaG || 0) + ga;
      if (starter) { const s = tm(ids[si]).starts; s[pid] = (s[pid] || 0) + 1; }
    });
    P.forEach(([pid, si, grp, off, def]) => {
      const x = pl[pid] || (pl[pid] = { n: 0, off: 0, def: 0 }); x.n++; x.off += off; x.def += def;
      lg.off[grp][0] += off; lg.off[grp][1]++; lg.def[grp][0] += def; lg.def[grp][1]++;
    });
  });
  // --- erreurs --- (la prédiction est linéaire dans les coefficients : erreur = e0 + Σ coef × z ; calcul rapide)
  const half = r => (Math.floor(r.pid / 7) % 2), NF = MU_F.length;
  const prep = set => {
    const n = set.length, e0 = new Float64Array(n), Z = MU_F.map(() => new Float64Array(n));
    set.forEach((r, i) => { e0[i] = r.bo + r.bd - (r.ao + r.ad);
      MU_F.forEach(([k, , app], j) => { const v = r.f[k] || 0; Z[j][i] = (app !== 'def' ? r.bo * v : 0) + (app !== 'off' ? r.bd * v : 0); }); });
    return { n: n, e0: e0, Z: Z };
  };
  const err = (S, c) => {
    const act = []; MU_F.forEach(([k], j) => { if (c[k]) act.push([S.Z[j], c[k]]); });
    let se = 0, ae = 0;
    for (let i = 0; i < S.n; i++) { let e = S.e0[i]; for (let q = 0; q < act.length; q++) e += act[q][1] * act[q][0][i]; se += e * e; ae += e < 0 ? -e : e; }
    return { mse: se / S.n, mae: ae / S.n };
  };
  const H = [prep(rows.filter(r => !half(r))), prep(rows.filter(r => half(r)))], ALL = prep(rows);
  const fit = (set, keys) => {
    const c = {}; let best = err(set, c).mse;
    for (let pass = 0; pass < (keys.length > 1 ? 3 : 1); pass++) keys.forEach(k => {
      let bv = c[k] || 0;
      MU.GRID.forEach(v => { const t = Object.assign({}, c); t[k] = v; const e = err(set, t).mse; if (e < best - 1e-9) { best = e; bv = v; } });
      c[k] = bv;
    });
    return c;
  };
  const base = [err(H[0], {}), err(H[1], {})];
  // a) chaque facteur seul : réglé sur une moitié, mesuré sur l'autre
  const single = MU_F.map(([k, lbl]) => {
    const cA = fit(H[0], [k]), cB = fit(H[1], [k]);
    const gB = base[1].mse - err(H[1], cA).mse, gA = base[0].mse - err(H[0], cB).mse;
    const mB = base[1].mae - err(H[1], cA).mae, mA = base[0].mae - err(H[0], cB).mae;
    const cAll = fit(ALL, [k]);
    return { k: k, label: lbl, coefA: cA[k] || 0, coefB: cB[k] || 0, coef: cAll[k] || 0, gainMse: [round4_(gA), round4_(gB)], gainMae: [round4_(mA), round4_(mB)],
             retenu: gA > 0 && gB > 0 && Math.min(mA, mB) > -0.0005 && (gA + gB) / 2 >= MU.MIN_GAIN * 0.1 && Math.sign(cA[k] || 0) === Math.sign(cB[k] || 0) && (cA[k] || 0) !== 0 };
  });
  // b) tous les facteurs retenus ensemble
  const keep = single.filter(s => s.retenu).map(s => s.k);
  let joint = null, coefKept = {};
  if (keep.length) {
    const cA = fit(H[0], keep), cB = fit(H[1], keep);
    const gB = base[1].mae - err(H[1], cA).mae, gA = base[0].mae - err(H[0], cB).mae;
    coefKept = fit(ALL, keep);
    joint = { gainMae: [round4_(gA), round4_(gB)], gainMse: [round4_(base[0].mse - err(H[0], cB).mse), round4_(base[1].mse - err(H[1], cA).mse)], coef: coefKept };
  }
  // moyenne de la caractéristique (pour lire l'effet : coef × écart typique)
  const sd = k => { const v = rows.map(r => r.f[k]), m = v.reduce((s, x) => s + x, 0) / v.length; return Math.sqrt(v.reduce((s, x) => s + (x - m) * (x - m), 0) / v.length); };
  single.forEach(s => { s.ecartType = round4_(sd(s.k)); s.effet = Math.round(s.coef * (['b2b', 'ob2b', 'home', 'bk'].indexOf(s.k) >= 0 ? 1 : sd(s.k)) * 1000) / 10; });
  return { n: rows.length, nGames: recs.length, base: { mae: round4_((base[0].mae + base[1].mae) / 2), mse: round4_((base[0].mse + base[1].mse) / 2) },
           single: single, joint: joint, kept: keep, coefKept: coefKept };
}
function round4_(x) { return Math.round(x * 10000) / 10000; }
function muWrite_(ss, res) {
  const sh = ss.getSheetByName(MU.SHEET) || ss.insertSheet(MU.SHEET);
  sh.clear();
  const out = [], W = 8, row = a => { const r = a.slice(0, W); while (r.length < W) r.push(''); out.push(r.map(v => (typeof v === 'string' && /^\d/.test(v) ? "'" + v : v))); };
  row(['🧪 Backtest par match — tenir compte de l’adversaire du soir aide-t-il ? (saison 2025-26)']);
  if (res.incomplete) {
    row([`⏳ Lecture des matchs incomplète : ${res.done} / ${res.total} matchs lus (${res.seconds} s). Relance « 🧪 Backtest par match » : ça reprend où c’était rendu.`]);
  } else {
    row([`${res.nGames} matchs lus · ${res.n} matchs-joueurs évalués (à partir du ${MU.EVAL_FROM}, joueurs avec ≥ ${MU.MIN_N} matchs avant) · ${res.seconds} s`]);
    row([`Base : moyenne FP/match du joueur avant le match, ramenée vers sa position (k ${MU.KBASE}). Erreur moyenne de la base : ${res.base.mae} FP/match (carré moyen ${res.base.mse}). Les points en AN sont comptés.`]);
    row(['Validation : réglé sur une moitié des joueurs, mesuré sur l’autre (et l’inverse). Retenu seulement si ça aide sur les deux moitiés, avec le même sens d’effet.']);
    row([]);
    row(['Facteur', 'Effet (% de la projection, pour un écart typique)', 'Coef. (toute la saison)', 'Coef. moitié A / B', 'Gain erreur carrée A / B', 'Gain erreur moyenne A / B (FP/match)', 'Retenu']);
    res.single.forEach(s => row([s.label, (s.effet > 0 ? '+' : '') + s.effet + ' %', s.coef, s.coefA + ' / ' + s.coefB, s.gainMse.join(' / '), s.gainMae.join(' / '), s.retenu ? '✅ oui' : 'non']));
    row([]);
    if (res.joint) row([`Facteurs retenus ensemble : ${res.kept.join(', ')} · gain erreur moyenne ${res.joint.gainMae.join(' / ')} FP/match · gain erreur carrée ${res.joint.gainMse.join(' / ')} · coefficients ${JSON.stringify(res.joint.coef)}`]);
    else row(['Aucun facteur ne bat la base sur les deux moitiés : l’adversaire du soir n’améliore pas la projection d’un match.']);
    row(['Lecture : « Effet » = de combien la projection d’un match bouge pour un adversaire typiquement au-dessus de la moyenne (ou pour un oui, dans le cas des facteurs oui/non).']);
  }
  sh.getRange(1, 1, out.length, W).setValues(out);
  sh.setColumnWidth(1, 320);
}
function getMatchupBacktest() { return PropertiesService.getDocumentProperties().getProperty('BT_MATCHUP') || '{}'; }
