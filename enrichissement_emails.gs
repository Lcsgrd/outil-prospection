/**
 * ENRICHISSEMENT GRATUIT DES EMAILS — fonctionne dans N'IMPORTE QUEL Google Sheet
 * (Extensions > Apps Script, colle ce code, Ctrl + S, puis recharge le Google Sheet).
 *
 * Un menu « Prospection » apparaît en haut du Google Sheet : il travaille sur l'onglet affiché,
 * quel que soit le fichier ou le nom de l'onglet.
 *
 * Pour chaque ligne qui a un site web mais pas d'email, le script ouvre la page d'accueil du site
 * + ses pages "contact" et "mentions légales", et récupère l'adresse email affichée.
 *
 * Colonnes reconnues automatiquement en ligne 1 (ou ligne 2) :
 *   site web : "site_web", "Site web", "Site internet", "website"
 *   email    : "email", "Email", "Mail", "Mail client"
 * Le script ajoute une colonne "Site analysé le" pour ne jamais analyser deux fois le même site.
 *
 * Lance chercherEmailsSurSites() : environ 80 à 150 sites par passage (limite Google de 6 minutes).
 * Relance-la jusqu'à ce que le journal indique "plus rien à analyser".
 */

const ENR = {
  NOMS_SITE: ['site_web', 'site web', 'site internet', 'website', 'site'],
  NOMS_EMAIL: ['email', 'e-mail', 'mail', 'mail client', 'email agence'],
  COL_ANALYSE: 'Site analysé le',
  DUREE_MAX_MS: 5 * 60 * 1000,
  PAGES_MAX: 4,
};

/** Ajoute le menu « Prospection » à l'ouverture du fichier. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Prospection')
    .addItem('Chercher les emails sur les sites (onglet affiché)', 'chercherEmailsSurSites')
    .addSeparator()
    .addItem('Mode d\'emploi', 'aideEnrichissement')
    .addToUi();
}

function aideEnrichissement() {
  SpreadsheetApp.getUi().alert('Chercher les emails',
    '1. Affiche l\'onglet qui contient tes entreprises (ex. le CSV importé depuis l\'outil de recherche).\n' +
    '2. Il doit avoir une colonne "site_web" (ou "Site web") et une colonne "email" (ou "Email"), en ligne 1 ou 2.\n' +
    '3. Menu Prospection > Chercher les emails sur les sites.\n' +
    '4. Chaque passage traite 80 à 150 sites : relance jusqu\'au message "plus rien à analyser".\n\n' +
    'Petites villes : beaucoup d\'entreprises n\'ont pas de site, donc pas d\'email récupérable.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function chercherEmailsSurSites() {
  const sh = SpreadsheetApp.getActiveSheet();
  SpreadsheetApp.getActive().toast(`Analyse des sites de l'onglet « ${sh.getName()} »…`, 'Prospection', 10);
  const debut = Date.now();
  const { hr, cSite, cEmail, cAnalyse } = colonnesEnrichissement_(sh);
  const n = sh.getLastRow() - hr;
  if (n <= 0) return Logger.log('Aucune ligne.');
  const sites = sh.getRange(hr + 1, cSite, n, 1).getValues();
  const emails = sh.getRange(hr + 1, cEmail, n, 1).getValues();
  const analyses = sh.getRange(hr + 1, cAnalyse, n, 1).getValues();
  let vus = 0, trouves = 0, restants = 0;

  for (let i = 0; i < n; i++) {
    const site = String(sites[i][0] || '').trim();
    if (!site || String(emails[i][0] || '').trim() || analyses[i][0]) continue;
    if (Date.now() - debut > ENR.DUREE_MAX_MS) { restants++; continue; }
    const email = emailDepuisSite_(site);
    vus++;
    sh.getRange(hr + 1 + i, cAnalyse).setValue(new Date());
    if (email) { sh.getRange(hr + 1 + i, cEmail).setValue(email); trouves++; }
  }
  const msg = `${vus} site(s) analysé(s), ${trouves} email(s) trouvé(s).` +
    (restants ? ` Il en reste ${restants} : relance « Chercher les emails ».` : ' Plus rien à analyser.');
  Logger.log(msg);
  try { SpreadsheetApp.getActive().toast(msg, 'Prospection', 30); } catch (e) { /* lancé depuis l'éditeur */ }
}

function colonnesEnrichissement_(sh) {
  const lignes = sh.getRange(1, 1, Math.min(2, sh.getLastRow()), sh.getLastColumn()).getValues();
  const n = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
  for (let r = 0; r < lignes.length; r++) {
    const h = lignes[r].map(n);
    const cSite = h.findIndex(x => ENR.NOMS_SITE.includes(x)) + 1;
    const cEmail = h.findIndex(x => ENR.NOMS_EMAIL.includes(x)) + 1;
    if (cSite && cEmail) {
      let cAnalyse = h.indexOf(n(ENR.COL_ANALYSE)) + 1;
      if (!cAnalyse) {
        cAnalyse = sh.getLastColumn() + 1;
        if (cAnalyse > sh.getMaxColumns()) sh.insertColumnAfter(sh.getMaxColumns());
        sh.getRange(r + 1, cAnalyse).setValue(ENR.COL_ANALYSE).setFontWeight('bold');
      }
      return { hr: r + 1, cSite, cEmail, cAnalyse };
    }
  }
  throw new Error('Colonnes "site web" et "email" introuvables en ligne 1 ou 2 de cet onglet.');
}

/** Ouvre le site (accueil + pages contact / mentions légales) et renvoie l'email le plus pertinent. */
function emailDepuisSite_(site) {
  let url = site.startsWith('http') ? site : 'https://' + site;
  const domaine = (url.match(/^https?:\/\/(?:www\.)?([^\/:?#]+)/i) || [])[1] || '';
  // Sites de réseaux / annuaires : l'email trouvé serait celui du siège, pas celui de l'établissement
  if (/facebook\.|instagram\.|linkedin\.|pagesjaunes\.|google\.|leboncoin\.|seloger\./i.test(domaine)) return '';
  const html = lirePage_(url);
  if (!html) return '';
  const pages = [html];
  const liens = [...html.matchAll(/href=["']([^"'#]+)["']/gi)].map(m => m[1])
    .filter(h => /contact|mentions|legal|nous-joindre|coordonn|a-propos|qui-sommes/i.test(h))
    .map(h => h.startsWith("http") ? h : URL_(h, url))
    .filter((h, i, a) => h && a.indexOf(h) === i && h.includes(domaine))
    .slice(0, ENR.PAGES_MAX - 1);
  liens.forEach(l => { const p = lirePage_(l); if (p) pages.push(p); });
  const tous = pages.flatMap(extraireEmails_).filter((e, i, a) => a.indexOf(e) === i);
  if (!tous.length) return '';
  const memeDomaine = tous.filter(e => e.split('@')[1].replace(/^www\./, '') === domaine);
  const prefere = l => l.find(e => /^(contact|accueil|agence|info|bonjour|hello|commercial|secretariat)@/.test(e)) || l[0];
  return prefere(memeDomaine.length ? memeDomaine : tous);
}

function lirePage_(url) {
  try {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: false,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; prospection-selectra)' } });
    const type = String(r.getHeaders()['Content-Type'] || r.getHeaders()['content-type'] || '');
    return r.getResponseCode() < 400 && /html|text/i.test(type || 'html') ? r.getContentText().slice(0, 600000) : '';
  } catch (e) { return ''; }
}

function extraireEmails_(html) {
  const res = [];
  // Emails protégés par Cloudflare (data-cfemail)
  [...html.matchAll(/data-cfemail=["']([0-9a-f]+)["']/gi)].forEach(m => {
    const h = m[1], k = parseInt(h.slice(0, 2), 16);
    let s = ''; for (let i = 2; i < h.length; i += 2) s += String.fromCharCode(parseInt(h.slice(i, i + 2), 16) ^ k);
    res.push(s);
  });
  const texte = html.replace(/&#64;|&#x40;|\[at\]|\(at\)| arobase /gi, '@').replace(/&#46;|\[dot\]|\(dot\)/gi, '.');
  (texte.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/gi) || []).forEach(e => res.push(e));
  return res.map(e => e.toLowerCase().replace(/^mailto:/, ''))
    .filter(e => !/\.(png|jpe?g|gif|svg|webp|css|js)$/.test(e))
    .filter(e => !/(example|exemple|domain|sentry|wixpress|godaddy|wordpress|noreply|no-reply|ne-pas-repondre|votre|your|email@|nom@|prenom)/.test(e))
    .filter(e => e.length < 60);
}

/** Construit une URL absolue à partir d'un lien relatif (Apps Script n'a pas l'objet URL). */
function URL_(lien, base) {
  if (/^mailto:|^tel:|^javascript:/i.test(lien)) return '';
  const m = base.match(/^(https?:\/\/[^\/]+)(\/.*)?$/i);
  if (!m) return '';
  if (lien.startsWith('//')) return 'https:' + lien;
  if (lien.startsWith('/')) return m[1] + lien;
  const dossier = (m[2] || '/').replace(/[^\/]*$/, '');
  return m[1] + dossier + lien;
}
