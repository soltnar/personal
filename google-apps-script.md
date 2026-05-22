# Google Apps Script webhook для сайта

## Что делает
Принимает `POST` JSON от сайта и записывает:
- матрицу `Итого` на лист (по умолчанию `Импорт_Итого` или имя из `sheetName`)
- матрицы по группам на листы `Импорт_Кухня`, `Импорт_Зал`, `Импорт_Доставка`, `Импорт_Бар`

## Код Apps Script
```javascript
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents || "{}");
    var spreadsheetId = payload.spreadsheetId;
    if (!spreadsheetId) throw new Error("spreadsheetId is required");

    var ss = SpreadsheetApp.openById(spreadsheetId);

    writeMatrix_(ss, payload.sheetName || "Импорт_Итого", payload.matrix && payload.matrix.total);
    writeMatrix_(ss, "Импорт_Кухня", payload.matrix && payload.matrix.kitchen);
    writeMatrix_(ss, "Импорт_Зал", payload.matrix && payload.matrix.hall);
    writeMatrix_(ss, "Импорт_Доставка", payload.matrix && payload.matrix.delivery);
    writeMatrix_(ss, "Импорт_Бар", payload.matrix && payload.matrix.bar);

    return json_({ ok: true, message: "Updated" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function writeMatrix_(ss, sheetName, matrix) {
  if (!matrix || !matrix.length) return;
  var sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  sh.clear();
  sh.getRange(1, 1, matrix.length, matrix[0].length).setValues(matrix);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
```

## Как опубликовать
1. Открыть [script.google.com](https://script.google.com) и создать проект.
2. Вставить код, нажать `Deploy` -> `New deployment` -> `Web app`.
3. Execute as: `Me`; Who has access: `Anyone with the link`.
4. Скопировать `Web app URL` и вставить в поле `Webhook URL` на сайте.

## Важно
- В вашей Google таблице должны быть права редактирования у владельца Apps Script.
- Если структура финальной рабочей таблицы нестандартная (как на скриншоте с блоками), лучше сделать второй этап: точечный маппинг в конкретные ячейки.
