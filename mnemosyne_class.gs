/**
 * Мнемозина: результаты учеников в Google Таблице.
 * 1. Расширения → Apps Script. Удалите всё, вставьте этот код, сохраните.
 * 2. Выберите функцию setup и нажмите «Выполнить», разрешите доступ.
 * 3. Развернуть → Новое развёртывание → Веб-приложение. От имени «Я», доступ «Все».
 * 4. Адрес веб-приложения (оканчивается на /exec) вставьте в игре:
 *    Меню → Организатор → «Подключить таблицу». Игра выдаст ссылку для учеников.
 */
const HEAD = {
  'Итог': ['Ученик', 'Обновлено', 'Оценка из 10', 'Бонус', 'Корабль, %', 'Палуба 1, %', 'Палуба 2, %', 'Палуба 3, %', 'Палуба 4, %', 'Палуба 5, %', 'Палуба 6, %', 'Экзамен, %', 'Модель, %', 'Боссов из 6', 'Задач решено', 'Подсказок', 'Вставок кода', 'Мобов', 'Режим', 'Герой', 'Слабые темы'],
  'Задания': ['Ученик', 'Задание', 'Станция', 'Палуба', 'Лучший, %', 'Последний, %', 'Попыток', 'Верно', 'Подсказок', 'Когда'],
  'Ошибки': ['Ученик', 'Тема', 'Ошибок', 'Всего ответов', 'Обновлено']
};
const DARK = '#14385c', AMBER = '#ffd54a';

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  let sh = ss_().getSheetByName(name);
  if (!sh) { sh = ss_().insertSheet(name); const h = HEAD[name]; sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground(DARK).setFontColor('#ffffff').setWrap(true); sh.setFrozenRows(1); }
  return sh;
}

/* создаёт листы, формулы и графики; повторный запуск ничего не ломает */
function setup() {
  const ss = ss_();
  ['Итог', 'Задания', 'Ошибки'].forEach(sheet_);
  let d = ss.getSheetByName('Дашборд'); if (!d) d = ss.insertSheet('Дашборд', 0);
  d.clear(); d.getCharts().forEach(c => d.removeChart(c));
  d.setHiddenGridlines(true);
  d.getRange('A1').setValue('Мнемозина: итоги группы').setFontSize(18).setFontWeight('bold');
  const kpi = [['A3', '=COUNTA(\'Итог\'!A2:A)', 'учеников'], ['C3', '=IFERROR(ROUND(AVERAGE(\'Итог\'!C2:C),1),"—")', 'средняя оценка из 10'], ['E3', '=COUNTIF(\'Итог\'!N2:N,6)&" из "&COUNTA(\'Итог\'!A2:A)', 'победили всех боссов'], ['G3', '=COUNTIFS(\'Итог\'!A2:A,"<>",\'Итог\'!C2:C,"<"&$K$4)', 'нужна помощь']];
  kpi.forEach(([c, f, l]) => { const r = d.getRange(c); r.setFormula(f).setFontSize(26).setFontWeight('bold').setHorizontalAlignment('left'); r.offset(1, 0).setValue(l).setFontColor('#666666'); });
  d.getRange('G3').setFontColor('#d93025');
  d.getRange('J3:J4').setValues([['зелёный от'], ['красный ниже']]).setFontColor('#666666');
  d.getRange('K3:K4').setValues([[7], [5]]).setBackground(AMBER).setFontWeight('bold').setHorizontalAlignment('center');
  d.getRange('J5').setValue('жёлтые ячейки можно менять').setFontColor('#999999').setFontSize(9);
  /* данные для графиков */
  d.getRange('A40').setValue('Данные для графиков').setFontWeight('bold');
  d.getRange('A41:B41').setValues([['Ученик', 'Оценка']]);
  d.getRange('A42').setFormula('=IFERROR(SORT(FILTER({\'Итог\'!A2:A,\'Итог\'!C2:C},\'Итог\'!A2:A<>""),2,FALSE),"")');
  d.getRange('D41:E41').setValues([['Тема', 'Ошибок']]);
  d.getRange('D42').setFormula('=IFERROR(QUERY(\'Ошибки\'!A2:C,"select B, sum(C) where B is not null group by B order by sum(C) desc limit 12 label sum(C) \'\'",0),"")');
  d.getRange('G41:H41').setValues([['Блок', 'Средний %']]);
  const blocks = [['Палуба 1', 'F'], ['Палуба 2', 'G'], ['Палуба 3', 'H'], ['Палуба 4', 'I'], ['Палуба 5', 'J'], ['Палуба 6', 'K'], ['Экзамен', 'L'], ['Модель', 'M']];
  blocks.forEach(([n, col], i) => { d.getRange(42 + i, 7).setValue(n); d.getRange(42 + i, 8).setFormula('=IFERROR(ROUND(AVERAGE(\'Итог\'!' + col + '2:' + col + '),0),0)'); });
  d.getRange('A41:H41').setFontWeight('bold');
  const bar = (rng, row, col, title, sub, color) => d.insertChart(d.newChart().setChartType(Charts.ChartType.BAR).addRange(d.getRange(rng)).setNumHeaders(1)
    .setPosition(row, col, 0, 0).setOption('title', title).setOption('subtitle', sub).setOption('legend', { position: 'none' }).setOption('colors', [color])
    .setOption('width', 560).setOption('height', 340).setOption('hAxis', { minValue: 0 }).build());
  bar('A41:B101', 7, 1, 'Рейтинг по оценке', 'оценка из 10, сверху лучшие', '#14385c');
  bar('D41:E53', 7, 7, 'Где группа ошибается чаще всего', 'ошибки всей группы по темам', '#d9822b');
  bar('G41:H49', 24, 1, 'Средний результат по блокам', 'лучший балл по задачам блока, %', '#2f8f6a');
  d.setColumnWidths(1, 12, 95);
  /* итог: вывод словами и цвет оценки */
  const it = sheet_('Итог');
  it.getRange('V1').setFormula('={"Вывод";ARRAYFORMULA(IF(A2:A="","",IF(C2:C<INDIRECT("\'Дашборд\'!K4"),"нужна помощь",IF(C2:C>=INDIRECT("\'Дашборд\'!K3"),"усвоил(а) материал","есть пробелы"))))}');
  it.getRange('V1').setFontWeight('bold').setBackground(DARK).setFontColor('#ffffff');
  const c = it.getRange('C2:C');
  it.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(C2<>"",C2>=INDIRECT("\'Дашборд\'!K3"))').setBackground('#cdeccf').setRanges([c]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(C2<>"",C2<INDIRECT("\'Дашборд\'!K4"))').setBackground('#f6c7c3').setRanges([c]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=C2<>""').setBackground('#fff0b3').setRanges([c]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor('#d93025').setBold(true).setRanges([it.getRange('Q2:Q')]).build()]);
  /* темы: ученики по строкам, темы по столбцам, в клетке число ошибок */
  let t = ss.getSheetByName('Темы'); if (!t) t = ss.insertSheet('Темы');
  t.clear(); t.getRange('A1').setValue('Ошибки по темам').setFontSize(14).setFontWeight('bold');
  t.getRange('A2').setValue('Строки — ученики, столбцы — темы, в клетке число ошибок. Чем краснее, тем больше.').setFontColor('#666666');
  t.getRange('A4').setFormula('=IFERROR(QUERY(\'Ошибки\'!A1:C,"select A, sum(C) where A is not null group by A pivot B",1),"Ошибок пока нет")');
  t.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().setGradientMinpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0').setGradientMaxpointWithValue('#e06055', SpreadsheetApp.InterpolationType.PERCENTILE, '95').setRanges([t.getRange('B5:AZ300')]).build()]);
  t.getRange('A4:AZ4').setFontWeight('bold').setWrap(true);
  ss.setActiveSheet(d);
}

/* строки с одинаковым ключом заменяются, новые дописываются в конец */
function upsert_(name, rows, keyCols) {
  if (!rows.length) return;
  const sh = sheet_(name), n = HEAD[name].length, last = sh.getLastRow();
  const data = last > 1 ? sh.getRange(2, 1, last - 1, n).getValues() : [];
  const idx = {}; data.forEach((r, i) => idx[keyCols.map(c => String(r[c])).join('\u0001')] = i);
  const add = [];
  rows.forEach(r => { const k = keyCols.map(c => String(r[c])).join('\u0001'); if (k in idx) { if (idx[k] < data.length) data[idx[k]] = r; else add[idx[k] - data.length] = r; } else { idx[k] = data.length + add.length; add.push(r); } });
  if (data.length) sh.getRange(2, 1, data.length, n).setValues(data);
  if (add.length) sh.getRange(data.length + 2, 1, add.length, n).setValues(add);
}
const clean_ = (s, m) => String(s == null ? '' : s).replace(/^[=+\-@]/, "'$&").slice(0, m || 200);
const num_ = v => v == null || v === '' ? '' : Number(v) || 0;

function doGet(e) { return json_({ ok: true, title: ss_().getName() }); }
function doPost(e) {
  let p; try { p = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  if (!p || p.game !== 'mnemosyne' || !p.name) return json_({ ok: false, error: 'bad payload' });
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (!ss_().getSheetByName('Дашборд')) setup();
    const name = clean_(p.name, 60), now = new Date();
    if (p.type === 'join') {
      const sh = sheet_('Итог'), last = sh.getLastRow(), have = last > 1 && sh.getRange(2, 1, last - 1, 1).getValues().some(r => String(r[0]) === name);
      if (!have) sh.getRange(last + 1, 1, 1, HEAD['Итог'].length).setValues([[name, now, 0, '', 0, '', '', '', '', '', '', '', '', 0, 0, 0, 0, 0, '', '', '']]);
      return json_({ ok: true, title: ss_().getName() });
    }
    const s = p.sum || {}, d = s.decks || [];
    upsert_('Итог', [[name, now, num_(s.score) || 0, s.bonus ? 1 : '', num_(s.pct) || 0, ...[0, 1, 2, 3, 4, 5].map(i => num_(d[i])), num_(s.exam), num_(s.model), num_(s.boss) || 0, num_(s.tasks) || 0, num_(s.hints) || 0, num_(s.pastes) || 0, num_(s.kills) || 0, clean_(s.mode, 20), clean_(s.hero, 30), clean_(s.weak, 300)]], [0]);
    upsert_('Задания', (p.tasks || []).slice(0, 400).map(t => [name, clean_(t.id, 40), clean_(t.st, 60), clean_(t.deck, 20), num_(t.best) || 0, num_(t.last) || 0, num_(t.tries) || 0, num_(t.ok) || 0, num_(t.hints) || 0, t.t ? new Date(t.t) : '']), [0, 1]);
    upsert_('Ошибки', (p.errs || []).slice(0, 200).map(x => [name, clean_(x.topic, 80), num_(x.miss) || 0, num_(x.n) || 0, now]), [0, 1]);
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
