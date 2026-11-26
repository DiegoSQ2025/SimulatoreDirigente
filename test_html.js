/* Suite jsdom per simulatore_dirigente.html — mese standard 26 GG, KPI netto = annuale/14, scenari + sanity. */
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('/tmp/node_modules/jsdom');

const html = fs.readFileSync(path.join(__dirname, 'simulatore_dirigente.html'), 'utf-8');
const dom = new JSDOM(html, { runScripts: 'dangerously', pretendToBeVisual: true });
const { window } = dom;
const { document } = window;

const ERRS = [];
const check = (nome, got, want, tol = 0.02) => {
  const ok = typeof want === 'number' ? Math.abs(got - want) <= tol : got === want;
  console.log(`  ${ok ? 'OK  ' : 'FAIL'} ${nome}: ${got}${ok ? '' : ' (atteso ' + want + ')'}`);
  if (!ok) ERRS.push(nome);
};

const eur2num = s => parseFloat(String(s).replace(/\./g, '').replace(',', '.'));
function rowValue(tableId, labelPart){
  const rows = [...document.getElementById(tableId).querySelectorAll('tr')];
  const r = rows.find(tr => tr.cells[0] && tr.cells[0].textContent.includes(labelPart));
  return r ? eur2num(r.cells[1].textContent) : NaN;
}
function cellValue(tableId, labelPart, col = 1){
  const rows = [...document.getElementById(tableId).querySelectorAll('tr')];
  const r = rows.find(tr => tr.cells[0] && tr.cells[0].textContent.includes(labelPart));
  return r ? eur2num(r.cells[col].textContent) : NaN;
}

function set(id, value){
  const el = document.getElementById(id);
  if (!el) throw new Error('input inesistente: ' + id);
  el.value = value;
  el.dispatchEvent(new window.Event('input', { bubbles: true }));
}
function get(id){ return document.getElementById(id); }
const kpi = id => eur2num(get(id).textContent);
const val = (t, part) => Math.abs(rowValue(t, part));

/* selezione iniziale */
check('numero enti regionali nel menu', get('regione').options.length, 21, 0);
check('default regione', get('regione').value, 'Lazio');
check('default comune', get('comune').value, 'Roma');
check('righe confronto minimi CCNL (header + 3 anni)', get('confronto').rows.length - 1, 3, 0);
check('label superminimo ad personam (SUP.CONTR + S.MINIMO)', /personam/i.test(get('supi').labels[0].textContent) && /SUP.CONTR/.test(get('supi').labels[0].textContent) && /S.MINIMO/.test(get('supi').labels[0].textContent) ? 1 : 0, 1, 0);
check('input giorni e SUP.CONTR rimossi (RAL = minimo + superminimo + SMM)', (get('ggRet') || get('ggFer') || get('ggFest') || get('supc')) ? 0 : 1, 1, 0);
check('menu anno mostra tabellare art. 5', get('anno').options[0].textContent.includes('4.660,00') ? 1 : 0, 1, 0);

/* scenario 1: default = tutti gli importi variabili a ZERO, solo minimo tabellare 4.660, mese pieno 26 GG, Lazio/Roma */
check('mensile: retribuzione 26 gg', rowValue('mensile', 'Retribuzione ('), 4660.00);
check('mensile: nessuna riga ferie/festivita', isNaN(rowValue('mensile', 'Ferie')) && isNaN(rowValue('mensile', 'Festivit\u00e0')) ? 1 : 0, 1, 0);
check('mensile: TOTALE COMPETENZE', rowValue('mensile', '= TOTALE COMPETENZE'), 4660.00);
check('mensile: IVS 9,19% (ROUND euro 4.660,00)', val('mensile', 'Contributo IVS 9,19%'), 428.25);
check('mensile: IVS agg. 1% (sotto soglia mensile)', val('mensile', 'IVS agg. 1%'), 0);
check('mensile: Fondo Mario Negri 2%', val('mensile', 'Fondo Mario Negri 2%'), 98.71);
check('mensile: FASDAC 1,87%', val('mensile', 'FASDAC 1,87%'), 71.59);
check('mensile: A. Pastore dirigente', val('mensile', 'A. Pastore quota dirigente'), 38.73);
check('mensile: formazione Mario Negri', val('mensile', 'Fondo formazione Mario Negri'), 10.83);
check('mensile: IRPEF lorda (annua/14)', val('mensile', 'IRPEF lorda'), 1223.36);
check('mensile: add. regionale annua/12', val('mensile', 'Addizionale regionale'), 140.87);
check('mensile: add. comunale annua/12', val('mensile', 'Addizionale comunale'), 43.48);
check('mensile: TOTALE TRATTENUTE', rowValue('mensile', '= TOTALE TRATTENUTE'), 2055.82);
check('mensile: NETTO DEL MESE (payslip, addizionali/12)', rowValue('mensile', '= NETTO DEL MESE'), 2604.18);
check('mensile: fringe benefit (info, default 0)', val('mensile', 'Fringe benefit'), 0);
check('kpi: netto mensile medio (annuale/14)', kpi('kpiNettoM'), 2748.42);
check('kpi: nota = netto annuale / 14', get('kpiNote').textContent.includes('38.477,92') && get('kpiNote').textContent.includes('14') ? 1 : 0, 1, 0);
check('kpi: totale trattenute del mese (payslip)', kpi('kpiTratt'), 2055.82);
check('kpi: RAL 14 mensilita', kpi('kpiRAL'), 65240.00);
check('kpi: netto annuo in busta', kpi('kpiNettoA'), 38477.92);
check('annuale: RME', rowValue('annuale', 'RME \u2014 retribuzione mensile'), 4660.00);
check('annuale: retr. base di contratto CCNL', rowValue('annuale', 'Retr. base di contratto'), 65240.00);
check('annuale: superminimo ad personam annuo (default 0)', rowValue('annuale', 'Superminimo ad personam'), 0);
check('annuale: RAL (base contratto + S.MINIMO + scatti)', rowValue('annuale', 'RAL \u2014 retribuzione annua effettiva'), 65240.00);
check('annuale: IVS annua', rowValue('annuale', 'IVS 9,19% annua'), 5995.56);
check('annuale: IVS agg. annua', rowValue('annuale', 'IVS agg. 1% annua'), 90.16);
check('annuale: Negri annuo', rowValue('annuale', 'Fondo Mario Negri 2% annuo'), 1184.49);
check('annuale: FASDAC annuo', rowValue('annuale', 'FASDAC 1,87% annuo'), 859.08);
check('annuale: Pastore annuo', rowValue('annuale', 'A. Pastore annuo'), 464.81);
check('annuale: RLD', rowValue('annuale', '= RLD'), 57969.79);
check('annuale: imponibile IRPEF', rowValue('annuale', 'Imponibile IRPEF annuo'), 57969.79);
check('annuale: IRPEF lorda', rowValue('annuale', 'IRPEF lorda annua'), 17127.01);
check('annuale: detrazione art. 13 (RLD < 75k)', rowValue('annuale', 'Detrazione lavoro dip. art. 13'), 1301.11);
check('annuale: IRPEF netta', rowValue('annuale', 'IRPEF NETTA annua'), 15825.90);
check('annuale: addizionale regionale Lazio', rowValue('annuale', 'Addizionale regionale annua'), 1690.39);
check('annuale: addizionale comunale Roma', rowValue('annuale', 'Addizionale comunale annua'), 521.73);
check('annuale: NETTO ANNUO in busta', rowValue('annuale', 'NETTO ANNUO in busta'), 38477.92);
check('annuale: netto medio per mensilita', rowValue('annuale', 'Netto medio per mensilit'), 2748.42);

/* scenario 1bis: replica del cedolino 03 (supi = AUM.RETR. 500 + S.MINIMO 6.489,99714 + Ad Pers. 857,14286
   = 7.847,14; autovettura in denaro 261,30; mese pieno 26 GG) */
['supi', 'scatti', 'fringe', 'fringeCash',
 'irpOv', 'regRes', 'comRes', 'acconto', 'arrotT', 'arrotN'].forEach(id => set(id, '0'));
set('smm', '350'); set('supi', '7847.14'); set('fringeCash', '261.30');
check('ced.03: retribuzione 26 gg 494,50538', rowValue('mensile', 'Retribuzione ('), 12857.14);
check('ced.03: indennit\u00e0 denaro (autovettura)', val('mensile', 'Indennit\u00e0 denaro'), 261.30);
check('ced.03: TOTALE COMPETENZE 13.118,44', rowValue('mensile', '= TOTALE COMPETENZE'), 13118.44);
check('ced.03: IVS 1.205,54 su base arrotondata 13.118,00', val('mensile', 'Contributo IVS 9,19%'), 1205.54);
check('ced.03: IVS agg. 1% 84,33', val('mensile', 'IVS agg. 1%'), 84.33);
check('ced.03: IRPEF lorda annua/14', val('mensile', 'IRPEF lorda'), 4665.87);
check('ced.03: add. regionale annua/12', val('mensile', 'Addizionale regionale'), 451.89);
check('ced.03: TOTALE TRATTENUTE 6.755,03', rowValue('mensile', '= TOTALE TRATTENUTE'), 6755.03);
check('ced.03: NETTO DEL MESE (payslip)', rowValue('mensile', '= NETTO DEL MESE'), 6363.41);
check('ced.03: kpi netto mensile medio (93.186,70/14)', kpi('kpiNettoM'), 6656.19);
check('ced.03: kpi RAL 179.999,96', kpi('kpiRAL'), 179999.96);
check('ced.03: kpi netto annuo 93.186,70', kpi('kpiNettoA'), 93186.70);
check('ced.03: RAL base di contratto 70.140,00', rowValue('annuale', 'Retr. base di contratto'), 70140.00);
check('ced.03: superminimo ad personam annuo 109.859,96', rowValue('annuale', 'Superminimo ad personam'), 109859.96);
check('ced.03: RAL effettiva 179.999,96', rowValue('annuale', 'RAL \u2014 retribuzione annua effettiva'), 179999.96);

/* scenario 2: ritorno ai default (retribuzione cedolino 01) + anno 2027 (minimo 4.920) */
['supi', 'scatti', 'fringe', 'fringeCash',
 'irpOv', 'regRes', 'comRes', 'acconto', 'arrotT', 'arrotN'].forEach(id => set(id, '0'));
set('smm', '350'); set('supi', '2974.28572'); set('fringe', '136.4'); set('fringeCash', '0');
set('anno', '2027');
check('2027: TOTALE COMPETENZE (RME 8.244,29)', rowValue('mensile', '= TOTALE COMPETENZE'), 8244.29);
check('2027: NETTO DEL MESE (payslip)', rowValue('mensile', '= NETTO DEL MESE'), 4224.43);
check('2027: TOTALE TRATTENUTE', val('mensile', '= TOTALE TRATTENUTE'), 4019.86);
check('2027: NETTO ANNUO in busta', rowValue('annuale', 'NETTO ANNUO in busta'), 60185.43);
check('2027: kpi netto mensile medio (60.185,43/14)', kpi('kpiNettoM'), 4298.96);

/* scenario 3: Lombardia/Milano, anno 2026, mese pieno standard */
set('anno', '2026'); set('regione', 'Lombardia'); set('comune', 'Milano');
check('Lombardia/Milano: NETTO DEL MESE (payslip)', rowValue('mensile', '= NETTO DEL MESE'), 4234.57);
check('Lombardia/Milano: add. regionale annua', rowValue('annuale', 'Addizionale regionale annua'), 1657.59);
check('Lombardia/Milano: add. comunale annua', rowValue('annuale', 'Addizionale comunale annua'), 811.23);
check('Lombardia/Milano: kpi netto mensile medio (60.040,90/14)', kpi('kpiNettoM'), 4288.64);
set('regione', 'Lazio'); set('comune', 'Roma');
check('ritorno Lazio/Roma: kpi netto mensile medio', kpi('kpiNettoM'), 4175.74);

/* sanity: nessun NaN/undefined/Infinity nelle tabelle e nei KPI */
let bad = 0;
document.querySelectorAll('#mensile tr, #annuale tr, #confronto tr').forEach(tr => {
  [...tr.cells].forEach(td => { if (/NaN|undefined|Infinity/.test(td.textContent)) bad++; });
});
check('nessun NaN/undefined/Infinity nelle tabelle', bad, 0, 0);
check('kpi netto non \u2014', get('kpiNettoM').textContent.includes('\u2014') ? 1 : 0, 0, 0);

/* accessibilita: ogni input/select ha una label esplicita o aria-label */
let noLabel = 0;
document.querySelectorAll('input,select').forEach(el => {
  if ((!el.labels || !el.labels.length) && !el.getAttribute('aria-label')) noLabel++;
});
check('input/select senza label o aria-label', noLabel, 0, 0);

/* marker dati MEF intatti */
check('marker dati MEF presenti', html.includes('/*__COMUNI_DATA__*/') || (html.includes('__COMUNI_DATA_START__') && html.includes('__COMUNI_DATA_END__')) ? 1 : 0, 1, 0);

console.log(ERRS.length === 0
  ? '\nESITO: TUTTI I TEST HTML PASSATI'
  : `\nESITO: ${ERRS.length} TEST FALLITI: ${ERRS.join(', ')}`);
process.exit(ERRS.length === 0 ? 0 : 1);
