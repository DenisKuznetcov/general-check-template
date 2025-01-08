/**
 * Функция `run` предназначена для автоматизации настройки рабочего листа "Генпроверка".
 * На лист "Генпроверка" добавляются данные, начиная со столбца А, затем запускается эта функция при 
 * нажатии на кнопку "Добавить столбцы генпроверки" на листе "Settings"
 * Она добавляет новые столбцы, настраивает выпадающие списки, условное форматирование 
 * и формулы, используя данные с листа "Settings".
 */
function run() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Генпроверка");
  if (mainSheet === null) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Нет листа Генпроверка');
    return;
  }

  const settingsSheet = ss.getSheetByName("Settings");
  if (settingsSheet === null) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Нет листа Settings");
    return;
  }

  // Данные с листа "Settings"
  const settingsLastRow = settingsSheet.getLastRow();
  const valuesFromSettings = settingsSheet.getRange(`A2:A${settingsLastRow}`).getValues().filter(v => v[0] !== "");
  const backgroundsFromSettings = settingsSheet.getRange(`B2:B${settingsLastRow}`).getBackgrounds().filter(v => v[0] !== "");
  const noAngleFromSettings = settingsSheet.getRange(`D2:D${settingsLastRow}`).getValues().filter(v => v[0] !== "");

  const valuesAndBackgrounds = new Map();

  for (let i = 0; i < valuesFromSettings.length; i++) {
    valuesAndBackgrounds.set(valuesFromSettings[i][0], backgroundsFromSettings[i][0]);
  }

  // Работа с листом "Генпроверка"
  const initialLastColumn = mainSheet.getLastColumn();
  const initialLastRow = mainSheet.getLastRow();
  const initialRange = mainSheet.getRange(`A1:${_getColumnLetter(initialLastColumn)}${initialLastRow}`);
  mainSheet.clearConditionalFormatRules();

  const newRange = mainSheet.getRange(1, initialLastColumn + 1, 1, 12);
  const headers = ["", "Выставление доработок сслыка на лот", "описание", "Работа с фото", 
    "Ошибка расшифровки", "Ошибка мерджа", "Памятник", "ветеран", 
    "опечатка дата", "опечатка в имени", "Никнейм", "нет ракурса"];

  newRange.setValues([headers])
  .setBackgrounds([headers.map(_ => "#cddaf6")])
  .setFontWeights([headers.map(_ => "bold")]);

  // Простановка выпадающих списков
  const firstRange  = mainSheet.getRange(2, initialLastColumn + 1, initialLastRow, 1);
  let rule = SpreadsheetApp.newDataValidation().requireValueInList(valuesFromSettings.map(v => v[0])).build();
  firstRange.setDataValidation(rule);

  const photoWorkRange = mainSheet.getRange(2, initialLastColumn + 4, initialLastRow, 1);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(["да"]).build();
  photoWorkRange.setDataValidation(rule);

  const restRange = mainSheet.getRange(2, initialLastColumn + 5, initialLastRow, 7);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(["исправлено"]).build();
  restRange.setDataValidation(rule);

  const angleRange  = mainSheet.getRange(2, initialLastColumn + 12, initialLastRow, 1);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(noAngleFromSettings.map(v => v[0])).build();
  angleRange.setDataValidation(rule);

  const targetColumnLetter = _getColumnLetter(initialLastColumn + 1);
  addConditionalFormatting(initialRange, valuesAndBackgrounds, mainSheet, targetColumnLetter);
  if (mainSheet.getFilter() === null) {
    mainSheet.getDataRange().createFilter();
  }

  try {
    mainSheet.deleteRow(mainSheet.getLastRow() + 1);
  } catch (exc) {}
  

  // add
  mainSheet.insertRowAfter(mainSheet.getLastRow());
  const newLastCol = mainSheet.getLastColumn();
  const range = mainSheet.getRange(mainSheet.getLastRow() + 1, 1, 1, newLastCol);
  range.clear().clearDataValidations();
  let formula;
  for (let i = newLastCol - 8; i <= newLastCol; i++) {
    const columnLetter = String.fromCharCode(64 + i);
    formula = "=COUNTA(" + columnLetter + "2:" + columnLetter + initialLastRow + ")/ROWS(" + columnLetter + "2:" + columnLetter + initialLastRow + ")";
    range.getCell(1, i).setFormula(formula);
    range.getCell(1, i).setNumberFormat("0.00%");
  }
  
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(mainSheet);
}

/**
 * Функция `fullDeleteMainSheet` предназначена для полного удаления данных, правил форматирования, 
 * валидаций данных и примечаний на листе "Генпроверка".
 * Также удаляет существующий фильтр, если он есть.
 */
function fullDeleteMainSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("Генпроверка");
  if (mainSheet === null) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Нет листа Генпроверка');
    return;
  }

  mainSheet.clearConditionalFormatRules();
  mainSheet.getRange(`A1:${_getColumnLetter(mainSheet.getLastColumn())}${mainSheet.getLastRow()}`)
  .clear()
  .clearDataValidations()
  .clearNote();

  if (mainSheet.getFilter() !== null) {
    mainSheet.getFilter().remove();
  }
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(mainSheet);
}

/**
 * Функция `addConditionalFormatting` добавляет правила условного форматирования к диапазону на основе 
 * значений и цветов фона, указанных в параметре `valuesAndBackgrounds`.
 *
 * @param {Range} range - Диапазон, к которому будет применено условное форматирование.
 * @param {Map} valuesAndBackgrounds - Map, содержащий пары "значение" -> "цвет фона".
 * @param {Sheet} sheet - Лист, на котором будет применено условное форматирование.
 * @param {string} columnLetter - Буква колонки, на которую ссылается формула условного форматирования.
 */
function addConditionalFormatting(range, valuesAndBackgrounds, sheet, columnLetter) {
  const rules = sheet.getConditionalFormatRules();

  for (const [value, background] of valuesAndBackgrounds) {
    const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${columnLetter}1="${value}"`)
    .setBackground(background)
    .setRanges([range])
    .build();
    rules.push(rule);
  }
  
  sheet.setConditionalFormatRules(rules);
}

function _getColumnLetter(columnNumber) {
  let letter = "";
  while (columnNumber > 0) {
    let remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - remainder) / 26);
  }
  return letter;
}