/* Suite jsdom per simulatore_dirigente.html — preset Zucchetti + scenari addizionali + sanity. */
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

/* scenario 1: preset identico al cedolino Zucchetti (dirigente Liv. 5, GG 25+1+1, Lazio/Roma, override conguaglio) */
check('mensile: retribuzione 25 gg', rowValue('mensile', 'Retribuzione ('), 7677.20);
check('mensile: TOTALE COMPETENZE', rowValue('mensile', '= TOTALE COMPETENZE'), 8291.38);
check('mensile: IVS 9,19% (FLOOR euro 8.291,00)', val('mensile', 'Contributo IVS 9,19%'), 761.94);
check('mensile: IVS agg. 1%', val('mensile', 'IVS agg. 1%'), 36.06);
check('mensile: Fondo Mario Negri 2%', val('mensile', 'Fondo Mario Negri 2%'), 98.71);
check('mensile: FASDAC 1,87%', val('mensile', 'FASDAC 1,87%'), 71.59);
check('mensile: A. Pastore dirigente', val('mensile', 'A. Pastore quota dirigente'), 38.73);
check('mensile: formazione Mario Negri', val('mensile', 'Fondo formazione Mario Negri'), 10.83);
check('mensile: IRPEF lorda (override cedolino)', val('mensile', 'IRPEF lorda'), 2755.77);
check('mensile: add. regionale residuo', val('mensile', 'Addizionale regionale'), 253.95);
check('mensile: add. comunale residuo', val('mensile', 'Addizionale comunale'), 52.97);
check('mensile: acconto addizionali', val('mensile', 'Acconto addizionali'), 27.33);
check('mensile: arrotondamento trattenute', val('mensile', 'Arrotondamento trattenute'), 0.03);
check('mensile: TOTALE TRATTENUTE', rowValue('mensile', '= TOTALE TRATTENUTE'), 4107.91);
check('mensile: NETTO DEL MESE (cedolino)', rowValue('mensile', '= NETTO DEL MESE'), 4184.00);
check('kpi: netto del mese', kpi('kpiNettoM'), 4184.00);
check('kpi: totale trattenute', kpi('kpiTratt'), 4107.91);
check('kpi: RAL 14 mensilita', kpi('kpiRAL'), 111780.00);
check('annuale: RME', rowValue('annuale', 'RME \u2014 retribuzione mensile'), 7984.29);
check('annuale: RAL', rowValue('annuale', 'RAL \u2014 retribuzione annua lorda'), 111780.00);
check('annuale: IVS annua', rowValue('annuale', 'IVS 9,19% annua'), 10272.58);
check('annuale: IVS agg. annua', rowValue('annuale', 'IVS agg. 1% annua'), 555.56);
check('annuale: Negri annuo', rowValue('annuale', 'Fondo Mario Negri 2% annuo'), 1184.49);
check('annuale: FASDAC annuo', rowValue('annuale', 'FASDAC 1,87% annuo'), 859.08);
check('annuale: Pastore annuo', rowValue('annuale', 'A. Pastore annuo'), 464.81);
check('annuale: RLD', rowValue('annuale', '= RLD'), 99767.37);
check('annuale: imponibile IRPEF', rowValue('annuale', 'Imponibile IRPEF annuo'), 101404.17);
check('annuale: IRPEF lorda', rowValue('annuale', 'IRPEF lorda annua'), 35803.79);
check('annuale: detrazione art. 13 (RLD > 75k = 0)', rowValue('annuale', 'Detrazione lavoro dip. art. 13'), 0);
check('annuale: IRPEF netta', rowValue('annuale', 'IRPEF NETTA annua'), 35803.79);
check('annuale: addizionale regionale Lazio', rowValue('annuale', 'Addizionale regionale annua'), 3136.76);
check('annuale: addizionale comunale Roma', rowValue('annuale', 'Addizionale comunale annua'), 912.64);
check('annuale: NETTO ANNUO in busta', rowValue('annuale', 'NETTO ANNUO in busta'), 58460.33);
check('annuale: netto medio per mensilita', rowValue('annuale', 'Netto medio per mensilit'), 4175.74);

/* scenario 2: mese pulito senza conguagli (override a 0, GG come cedolino) */
['irpOv', 'regRes', 'comRes', 'acconto', 'arrotT', 'arrotN'].forEach(id => set(id, '0'));
check('pulito: TRATTENUTE', rowValue('mensile', '= TOTALE TRATTENUTE'), 3912.72);
check('pulito: IRPEF lorda annua/14', val('mensile', 'IRPEF lorda'), 2557.41);
check('pulito: add. regionale annua/12', val('mensile', 'Addizionale regionale'), 261.40);
check('pulito: add. comunale annua/12', val('mensile', 'Addizionale comunale'), 76.05);
check('pulito: NETTO DEL MESE', rowValue('mensile', '= NETTO DEL MESE'), 4378.66);
check('pulito: kpi netto del mese', kpi('kpiNettoM'), 4378.66);

/* scenario 3: mese pieno 26 gg senza fringe, senza conguagli */
set('ggRet', '26'); set('ggFer', '0'); set('ggFest', '0'); set('fringe', '0');
check('26 gg: TOTALE COMPETENZE = RME', rowValue('mensile', '= TOTALE COMPETENZE'), 7984.29);
check('26 gg: NETTO DEL MESE', rowValue('mensile', '= NETTO DEL MESE'), 4158.89);
check('26 gg: NETTO ANNUO in busta', rowValue('annuale', 'NETTO ANNUO in busta'), 59233.39);

/* scenario 4: Lombardia/Milano, mese pulito 25+1+1 */
set('ggRet', '25'); set('ggFer', '1'); set('ggFest', '1'); set('fringe', '136.4');
set('regione', 'Lombardia');
set('comune', 'Milano');
check('Lombardia/Milano: NETTO DEL MESE', rowValue('mensile', '= NETTO DEL MESE'), 4510.38);
check('Lombardia/Milano: add. regionale annua', rowValue('annuale', 'Addizionale regionale annua'), 1657.59);
check('Lombardia/Milano: add. comunale annua', rowValue('annuale', 'Addizionale comunale annua'), 811.23);

/* scenario 5: anno 2027 (minimo 4.920), Lazio/Roma */
set('regione', 'Lazio'); set('comune', 'Roma'); set('anno', '2027');
check('2027 pulito: NETTO ANNUO in busta', rowValue('annuale', 'NETTO ANNUO in busta'), 60185.43);
check('2027 pulito: NETTO DEL MESE', rowValue('mensile', '= NETTO DEL MESE'), 4509.21);
check('2027 pulito: TOTALE TRATTENUTE', val('mensile', '= TOTALE TRATTENUTE'), 4052.17);

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
