/* Test de rendu dans Chromium : onglet Modèle de frais, six vues, saisie, captures d'écran. */
const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const file = 'file://' + path.resolve(process.argv[2] || '/home/user/OM_AirFund_CVC_gestion.html');
  const outDir = process.argv[3] || '/home/user/om_fees_tool/test/shots';
  require('fs').mkdirSync(outDir, { recursive: true });
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
  const errors = [];
  page.on('pageerror', e => errors.push('pageerror: ' + e.message));
  page.on('console', msg => { if (msg.type() === 'error') errors.push('console: ' + msg.text()); });
  page.on('dialog', d => d.accept());
  await page.goto(file + '#tab=fees-tool');
  await page.waitForSelector('[data-fee-mode]', { timeout: 15000 });
  const modes = ['recap', 'saisie', 'detail', 'provision', 'treso', 'breakeven'];
  const report = {};
  for (const k of modes) {
    await page.click('[data-fee-mode="' + k + '"]');
    await page.waitForTimeout(150);
    const pressed = await page.getAttribute('[data-fee-mode="' + k + '"]', 'aria-pressed');
    const cards = await page.$$eval('#content .card', els => els.length);
    const h = await page.$eval('#content', el => el.scrollHeight);
    report[k] = { pressed, cards, height: h };
    await page.screenshot({ path: path.join(outDir, 'fees_' + k + '.png'), fullPage: false });
  }
  // récap : total à 10 M€ avant / après saisie du nombre d'investisseurs (60 -> 160 : +1 500 € tenue du passif + KYC 0)
  await page.click('[data-fee-mode="recap"]'); await page.waitForTimeout(100);
  const totBefore = await page.$eval('#content table.tbl.fm.wide tr.total td:nth-child(5)', el => el.textContent.trim());
  await page.click('[data-fee-mode="saisie"]'); await page.waitForTimeout(100);
  await page.fill('[data-fv="params.nb_investisseurs"]', '160');
  await page.dispatchEvent('[data-fv="params.nb_investisseurs"]', 'change');
  await page.waitForTimeout(200);
  const ovClass = await page.getAttribute('[data-fv="params.nb_investisseurs"]', 'class');
  const barText = await page.$eval('#content .savebar .mode', el => el.textContent.trim());
  await page.click('[data-fee-mode="recap"]'); await page.waitForTimeout(100);
  const totAfter = await page.$eval('#content table.tbl.fm.wide tr.total td:nth-child(5)', el => el.textContent.trim());
  // persistance : rechargement de la page -> la saisie reste
  await page.reload(); await page.waitForSelector('[data-fee-mode]');
  await page.click('[data-fee-mode="saisie"]'); await page.waitForTimeout(100);
  const persisted = await page.inputValue('[data-fv="params.nb_investisseurs"]');
  // tranche : passer la ligne 6 (CAC) à 6 500 € puis vérifier le récap
  await page.fill('[data-fv="lines.5.tranches.0.valeur"]', '6500'); await page.dispatchEvent('[data-fv="lines.5.tranches.0.valeur"]', 'change'); await page.waitForTimeout(150);
  await page.click('[data-fee-mode="recap"]'); await page.waitForTimeout(100);
  const cacRow = await page.$eval('#content table.tbl.fm.wide tbody tr:nth-child(6) td:nth-child(4)', el => el.textContent.trim());
  // trésorerie : ajouter une facture, puis revenir au classeur
  await page.click('[data-fee-mode="treso"]'); await page.waitForTimeout(100);
  const nInvBefore = await page.$$eval('[data-inv-del]', els => els.length);
  await page.click('[data-inv-add]'); await page.waitForTimeout(150);
  const nInvAfter = await page.$$eval('[data-inv-del]', els => els.length);
  await page.screenshot({ path: path.join(outDir, 'fees_treso_full.png'), fullPage: true });
  await page.click('[data-fee-reset]'); await page.waitForTimeout(200);
  await page.click('[data-fee-mode="saisie"]'); await page.waitForTimeout(100);
  const afterReset = await page.inputValue('[data-fv="params.nb_investisseurs"]');
  await page.screenshot({ path: path.join(outDir, 'fees_saisie_full.png'), fullPage: true });
  // autre onglet : l'application reste fonctionnelle (vue d'ensemble, DIC)
  await page.goto(file + '#tab=overview'); await page.waitForTimeout(300);
  const ovCards = await page.$$eval('#content a.d', els => els.length);
  await page.goto(file + '#tab=kids-tool'); await page.waitForTimeout(300);
  const kidsH = await page.$eval('#content', el => el.scrollHeight);
  console.log(JSON.stringify({ report, totBefore, totAfter, ovClass, barText, persisted, cacRow, nInvBefore, nInvAfter, afterReset, ovCards, kidsH, errors }, null, 1));
  await browser.close();
  process.exit(errors.length ? 1 : 0);
})().catch(e => { console.error('SMOKE FAIL', e); process.exit(2); });
