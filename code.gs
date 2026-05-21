const SHEETS = {
  STOCK: '在庫一覧',
  SHOPPING: '買い物リスト',
  HISTORY: '使用履歴',
  NOTICE: '告知管理',
  MENU: '献立',
  MENU_HISTORY: '献立履歴'
};

const HEADERS = {
  [SHEETS.STOCK]: ['ID', '食品名', '数量', '単位', '金額', '購入日', '期限', '最低在庫数', '状態', '管理タイプ','保存場所'],
  [SHEETS.SHOPPING]: ['ID', '食品名', 'メモ', '追加日', '購入済み'],
  [SHEETS.HISTORY]: ['ID', '使用日', '食品名', '使用量', '単位', 'メモ'],
  [SHEETS.NOTICE]: ['ID', '種類', '食品名', 'メッセージ', '確認済み', '作成日'],
  [SHEETS.MENU]: ['料理名', '必須食材', '任意食材', '必要量', 'ジャンル', '調理時間', '難易度', 'タグ', '調味料'],
  [SHEETS.MENU_HISTORY]: ['ID', '日付', '献立名', '使用食材', 'メモ']
};

function doGet() {
  initSheets();

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('冷蔵庫管理アプリ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetByNameOrThrow(name) {
  const sheet = getSS().getSheetByName(name);
  if (!sheet) throw new Error(`「${name}」シートが見つかりません。`);
  return sheet;
}

function initSheets() {
  const ss = getSS();

  Object.keys(HEADERS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS[name]);
    }
  });
}

function createId() {
  return Utilities.getUuid();
}

function todayText() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function toNumber(value, label, options) {
  const opts = options || {};
  if (value === '' || value === null || value === undefined) {
    if (opts.defaultValue !== undefined) return opts.defaultValue;
    throw new Error(`${label}を入力してください。`);
  }

  const num = Number(value);
  if (!isFinite(num)) throw new Error(`${label}は数値で入力してください。`);
  if (opts.min !== undefined && num < opts.min) {
    throw new Error(`${label}は${opts.min}以上で入力してください。`);
  }

  return num;
}

function sanitizeText(value, label, required) {
  const text = String(value || '').trim();
  if (required && !text) throw new Error(`${label}を入力してください。`);
  return text;
}

function normalizeDateText(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  }

  const text = String(value).trim().replace(/-/g, '/');
  const parts = text.split('/');
  if (parts.length !== 3) throw new Error(`日付の形式が不正です：${value}`);

  const y = Number(parts[0]);
  const m = Number(parts[1]);
  const d = Number(parts[2]);

  if (!y || !m || !d) throw new Error(`日付の形式が不正です：${value}`);

  return `${y}/${String(m).padStart(2, '0')}/${String(d).padStart(2, '0')}`;
}

function parseDateOnly(value) {
  if (!value) return null;

  const text = normalizeDateText(value);
  const parts = text.split('/').map(Number);
  const date = new Date(parts[0], parts[1] - 1, parts[2]);
  date.setHours(0, 0, 0, 0);

  return date;
}

function getStatus(quantity, minStock) {
  if (quantity === 0) return '在庫切れ';
  if (quantity <= minStock) return '在庫少なめ';
  return '通常';
}

function getDefaultStorageLocation(manageType) {
  const type = String(manageType || '通常').trim();

  if (type === '米') return '常温';
  if (type === '乾麺') return '常温';
  if (type === '調味料') return '常温';

  return '冷蔵';
}

function findRowById(sheet, id, columnCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, columnCount).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      return {
        rowNumber: i + 2,
        values: values[i]
      };
    }
  }

  return null;
}

function addFood(foodData) {
  initSheets();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheetByNameOrThrow(SHEETS.STOCK);

    const name = sanitizeText(foodData.name, '食品名', true);
    const quantity = toNumber(foodData.quantity, '数量', { min: 0 });
    const unit = sanitizeText(foodData.unit, '単位', true);
    const price = toNumber(foodData.price, '金額', { min: 0 });
    const purchaseDate = normalizeDateText(foodData.purchaseDate || todayText());
    const expiryDate = foodData.expiryDate ? normalizeDateText(foodData.expiryDate) : '';
    const minStock = toNumber(foodData.minStock, '最低在庫数', { min: 0, defaultValue: 0 });
    const status = getStatus(quantity, minStock);
    const manageType = sanitizeText(foodData.manageType || '通常', '管理タイプ', false);
    const storageLocation = sanitizeText(
  foodData.storageLocation || getDefaultStorageLocation(manageType),
  '保存場所',
  false
);

    sheet.appendRow([
  createId(),
  name,
  quantity,
  unit,
  price,
  purchaseDate,
  expiryDate,
  minStock,
  status,
  manageType,
  storageLocation
]);
    return '登録が完了しました';
  } finally {
    lock.releaseLock();
  }
}

function updateFood(foodId, foodData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheetByNameOrThrow(SHEETS.STOCK);
    const found = findRowById(sheet, foodId, 11);
    if (!found) throw new Error('対象の食品が見つかりませんでした。');

    const name = sanitizeText(foodData.name, '食品名', true);
    const quantity = toNumber(foodData.quantity, '数量', { min: 0 });
    const unit = sanitizeText(foodData.unit, '単位', true);
    const price = toNumber(foodData.price, '金額', { min: 0 });
    const purchaseDate = foodData.purchaseDate ? normalizeDateText(foodData.purchaseDate) : '';
    const expiryDate = foodData.expiryDate ? normalizeDateText(foodData.expiryDate) : '';
    const minStock = toNumber(foodData.minStock, '最低在庫数', { min: 0, defaultValue: 0 });
    const status = getStatus(quantity, minStock);
    const manageType = sanitizeText(foodData.manageType || '通常', '管理タイプ', false);
    const storageLocation = sanitizeText(
  foodData.storageLocation || getDefaultStorageLocation(manageType),
  '保存場所',
  false
);

    sheet.getRange(found.rowNumber, 2, 1, 10).setValues([[
  name,
  quantity,
  unit,
  price,
  purchaseDate,
  expiryDate,
  minStock,
  status,
  manageType,
  storageLocation
]]);

    return '更新しました';
  } finally {
    lock.releaseLock();
  }
}

function getFoods() {
  initSheets();

  const sheet = getSheetByNameOrThrow(SHEETS.STOCK);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  return sheet.getRange(2, 1, lastRow - 1, 11).getValues().map(row => ({
  id: row[0],
  name: row[1],
  quantity: row[2],
  unit: row[3],
  price: row[4],
  purchaseDate: formatDate(row[5]),
  expiryDate: formatDate(row[6]),
  minStock: row[7],
  status: row[8],
  manageType: row[9] || '通常',
  storageLocation: row[10] || getDefaultStorageLocation(row[9])
}));
}

function deleteFood(foodId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheetByNameOrThrow(SHEETS.STOCK);
    const found = findRowById(sheet, foodId, 1);
    if (!found) throw new Error('対象の食品が見つかりませんでした。');

    sheet.deleteRow(found.rowNumber);
    return '削除しました';
  } finally {
    lock.releaseLock();
  }
}

function useFood(foodId, useData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const stockSheet = getSheetByNameOrThrow(SHEETS.STOCK);
    const historySheet = getSheetByNameOrThrow(SHEETS.HISTORY);
    const found = findRowById(stockSheet, foodId, 10);

    if (!found) throw new Error('対象の食品が見つかりませんでした。');

    const row = found.values;
    const foodName = row[1];
    const currentQuantity = toNumber(row[2], '現在の在庫数', { min: 0 });
    const unit = row[3];
    const minStock = toNumber(row[7], '最低在庫数', { min: 0, defaultValue: 0 });
    const useQuantity = toNumber(useData.useQuantity, '使用量', { min: 0.000001 });

    if (useQuantity > currentQuantity) {
      throw new Error('使用量が現在の在庫数を超えています。');
    }

    const newQuantity = currentQuantity - useQuantity;
    const status = getStatus(newQuantity, minStock);

    stockSheet.getRange(found.rowNumber, 3, 1, 1).setValue(newQuantity);
    stockSheet.getRange(found.rowNumber, 9, 1, 1).setValue(status);

    historySheet.appendRow([
      createId(),
      todayText(),
      foodName,
      useQuantity,
      unit,
      sanitizeText(useData.memo, 'メモ', false)
    ]);

    return {
      message: '使用しました',
      foodName,
      newQuantity,
      unit
    };
  } finally {
    lock.releaseLock();
  }
}

function addShoppingItem(itemData) {
  initSheets();

  const sheet = getSheetByNameOrThrow(SHEETS.SHOPPING);
  const name = sanitizeText(itemData.name, '買うもの', true);
  const memo = sanitizeText(itemData.memo, 'メモ', false);

  sheet.appendRow([
    createId(),
    name,
    memo,
    todayText(),
    false
  ]);

  return '買い物リストに追加しました';
}

function getShoppingItems() {
  initSheets();

  const sheet = getSheetByNameOrThrow(SHEETS.SHOPPING);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  return sheet.getRange(2, 1, lastRow - 1, 5).getValues().map(row => ({
    id: row[0],
    name: row[1],
    memo: row[2],
    addedDate: formatDate(row[3]),
    purchased: row[4] === true || row[4] === 'TRUE'
  }));
}

function deleteShoppingItem(itemId) {
  const sheet = getSheetByNameOrThrow(SHEETS.SHOPPING);
  const found = findRowById(sheet, itemId, 1);
  if (!found) throw new Error('対象の買い物リストが見つかりませんでした。');

  sheet.deleteRow(found.rowNumber);
  return '削除しました';
}

function completeShoppingItem(itemId) {
  const sheet = getSheetByNameOrThrow(SHEETS.SHOPPING);
  const found = findRowById(sheet, itemId, 5);
  if (!found) throw new Error('対象の買い物リストが見つかりませんでした。');

  sheet.getRange(found.rowNumber, 5).setValue(true);

  return {
    message: '購入済みにしました',
    name: found.values[1],
    memo: found.values[2]
  };
}

function getUseHistory() {
  initSheets();

  const sheet = getSheetByNameOrThrow(SHEETS.HISTORY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  return sheet.getRange(2, 1, lastRow - 1, 6).getValues()
    .map(row => ({
      id: row[0],
      usedDate: formatDate(row[1]),
      name: row[2],
      quantity: row[3],
      unit: row[4],
      memo: row[5]
    }))
    .reverse();
}

function getHomeData() {
  initSheets();
  generateNotifications();

  const stockItems = getFoods();
  const shoppingItems = getShoppingItems();
  const useHistory = getUseHistory();
  const notifications = getNotifications(false);

  const activeShoppingItems = shoppingItems.filter(item => !item.purchased);

  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth();

  const monthlyTotal = stockItems.reduce((sum, item) => {
    if (!item.purchaseDate) return sum;

    const purchaseDate = parseDateOnly(item.purchaseDate);
    if (!purchaseDate) return sum;

    if (
      purchaseDate.getFullYear() === currentYear &&
      purchaseDate.getMonth() === currentMonth
    ) {
      return sum + (Number(item.price) || 0);
    }

    return sum;
  }, 0);

  return {
    stockCount: stockItems.length,
    shoppingCount: activeShoppingItems.length,
    monthlyTotal,
    historyCount: useHistory.length,
    stockPreview: stockItems.slice(0, 5),
    shoppingPreview: activeShoppingItems.slice(0, 5),
    notifications
  };
}

function generateNotifications() {
  const noticeSheet = getSheetByNameOrThrow(SHEETS.NOTICE);
  const foods = getFoods();

  const noticeRows = [];

 foods.forEach(food => {
  if (food.manageType === '調味料') {
    return;
  }

  const quantity = Number(food.quantity);
  const minStock = Number(food.minStock || 0);

  if (quantity === 0) {
    noticeRows.push(
      buildNoticeRow(
        '在庫切れ',
        food.name,
        `${food.name}が在庫切れです。買い足しがおすすめです。`
      )
    );
  } else if (quantity <= minStock) {
    noticeRows.push(
      buildNoticeRow(
        '在庫不足',
        food.name,
        `${food.name}が残り${food.quantity}${food.unit}です。買い足しがおすすめです。`
      )
    );
  }

const storageLocation = String(food.storageLocation || '冷蔵').trim();


if (isExpiredServer(food.expiryDate)) {
  noticeRows.push(
    buildNoticeRow(
      '期限切れ',
      food.name,
      `${food.name}の期限が切れています。確認してください。`
    )
  );
} else if (storageLocation !== '冷凍' && isExpirySoonServer(food.expiryDate)) {
  noticeRows.push(
    buildNoticeRow(
      '期限間近',
      food.name,
      `${food.name}の期限が近いです。早めに使いましょう。`
    )
  );
}
});

  const lastRow = noticeSheet.getLastRow();
  const existingValues = lastRow > 1
    ? noticeSheet.getRange(2, 1, lastRow - 1, 6).getValues()
    : [];

  const existingKeys = new Set(
    existingValues.map(row => noticeKey(row[1], row[2], row[3]))
  );

  noticeRows.forEach(row => {
    if (!existingKeys.has(noticeKey(row[1], row[2], row[3]))) {
      noticeSheet.appendRow(row);
    }
  });
}

function buildNoticeRow(type, foodName, message) {
  return [
    createId(),
    type,
    foodName,
    message,
    false,
    todayText()
  ];
}

function noticeKey(type, foodName, message) {
  return `${type}_${foodName}_${message}`;
}

function getNotifications(shouldGenerate) {
  initSheets();
  if (shouldGenerate !== false) generateNotifications();

  const sheet = getSheetByNameOrThrow(SHEETS.NOTICE);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  return sheet.getRange(2, 1, lastRow - 1, 6).getValues()
    .map(row => ({
      id: row[0],
      type: row[1],
      name: row[2],
      message: row[3],
      checked: row[4] === true || row[4] === 'TRUE',
      createdDate: formatDate(row[5])
    }))
    .filter(item => !item.checked)
    .reverse();
}

function confirmNotification(noticeId) {
  const sheet = getSheetByNameOrThrow(SHEETS.NOTICE);
  const found = findRowById(sheet, noticeId, 1);
  if (!found) throw new Error('対象の告知が見つかりませんでした。');

  sheet.getRange(found.rowNumber, 5).setValue(true);
  return '確認済みにしました';
}

function addShoppingFromNotification(noticeId) {
  const sheet = getSheetByNameOrThrow(SHEETS.NOTICE);
  const found = findRowById(sheet, noticeId, 6);
  if (!found) throw new Error('対象の告知が見つかりませんでした。');

  const type = found.values[1];
  const foodName = found.values[2];

  addShoppingItem({
    name: foodName,
    memo: `${type}の告知から追加`
  });

  sheet.getRange(found.rowNumber, 5).setValue(true);

  return '買い物リストに追加しました';
}

function isExpirySoonServer(expiryDateText) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const expiry = parseDateOnly(expiryDateText);
  if (!expiry) return false;

  const diffDays = Math.floor((expiry.getTime() - today.getTime()) / 86400000);
  return diffDays >= 0 && diffDays <= 3;
}

function isExpiredServer(expiryDateText) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const expiry = parseDateOnly(expiryDateText);
  if (!expiry) return false;

  return expiry.getTime() < today.getTime();
}

function formatDate(dateValue) {
  if (!dateValue) return '';
  return normalizeDateText(dateValue);
}

function getMenuData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("献立");

  if (!sheet) {
    throw new Error("献立シートが見つかりません");
  }

  const values = sheet.getDataRange().getValues();

  const menus = values.slice(1).map(row => {
    return {
     name: row[0],
     required: row[1] ? String(row[1]).split(",").map(item => item.trim()) : [],
     optional: row[2] ? String(row[2]).split(",").map(item => item.trim()) : [],
     requiredAmount: row[3] ? String(row[3]).split(",").map(item => item.trim()) : [],
     genre: String(row[4] || '').trim(),
     time: row[5],
     difficulty: row[6],
     tags: row[7] ? String(row[7]).split(",").map(tag => tag.trim()) : [],
     seasonings: row[8] ? String(row[8]).split(",").map(item => item.trim()) : []
    };
  });

  return menus;
}

function suggestMenusFromStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const stockSheet = ss.getSheetByName("在庫一覧");
  const menuSheet = ss.getSheetByName("献立");

  if (!stockSheet) {
    throw new Error("在庫一覧シートが見つかりません");
  }

  if (!menuSheet) {
    throw new Error("献立シートが見つかりません");
  }

  const stockValues = stockSheet.getDataRange().getValues();
  const menuValues = menuSheet.getDataRange().getValues();

  const stockItems = stockValues
  .slice(1)
  .filter(row => {
    const name = String(row[1] || '').trim();
    const quantity = Number(row[2]) || 0;
    const manageType = String(row[9] || '通常').trim();

    if (!name) return false;

    if (manageType === '調味料') {
      return quantity > 0;
    }

    return quantity > 0;
  })
  .map(row => String(row[1] || '').trim());

  const stockMap = {};

  stockValues.slice(1).forEach(row => {
    const name = String(row[1] || '').trim();
    if (!name) return;

    stockMap[name] = {
  quantity: Number(row[2]) || 0,
  unit: String(row[3] || '').trim(),
  expiryDate: row[6],
  manageType: row[9] || '通常',
  storageLocation: row[10] || getDefaultStorageLocation(row[9])
};
  });

  const menus = menuValues
    .slice(1)
    .filter(row => {
      const name = String(row[0] || '').trim();
      const required = String(row[1] || '').trim();

      return name !== '' && required !== '';
    })
    .map(row => {
      const name = String(row[0] || '').trim();
      const required = row[1] ? String(row[1]).split(",").map(item => item.trim()).filter(item => item !== "") : [];
      const optional = row[2] ? String(row[2]).split(",").map(item => item.trim()).filter(item => item !== "") : [];
      const requiredAmount = row[3] ? String(row[3]).split(",").map(item => item.trim()).filter(item => item !== "") : [];
      const parsedRequiredAmount = parseAmountText(row[3]).map(convertCookedRiceToRawRice);
      const genre = String(row[4] || '').trim();
      const time = row[5];
      const difficulty = row[6];
      const tags = row[7];
      const seasonings = row[8]
        ? String(row[8]).split(",").map(item => item.trim()).filter(item => item !== "")
        : [];

      const missing = required.filter(item => !stockItems.includes(item));
      const matchedOptional = optional.filter(item => stockItems.includes(item));
      const missingSeasonings = seasonings.filter(item => !stockItems.includes(item));

      const shortageByAmount = parsedRequiredAmount.filter(item => {
        const stock = stockMap[item.name];

        if (!stock) return true;
        if (stock.unit !== item.unit) return true;

        return stock.quantity < item.amount;
      });

      let score = 0;

      score += (required.length - missing.length) * 10;

      if (missing.length === 0 && shortageByAmount.length === 0 && missingSeasonings.length === 0) {
        score += 50;
      }

      score -= missing.length * 15;
      score -= shortageByAmount.length * 12;
      score -= missingSeasonings.length * 5;
      score += matchedOptional.length * 3;

      const cookingTime = Number(time) || 999;
      if (cookingTime <= 10) {
        score += 10;
      } else if (cookingTime <= 15) {
        score += 8;
      } else if (cookingTime <= 20) {
        score += 5;
      } else if (cookingTime <= 30) {
        score += 2;
      }

      let expiryScore = 0;

      required.forEach(itemName => {
  const stock = stockMap[itemName];

  if (stock) {
    if (stock.storageLocation === '冷凍') {
      return;
    }

    expiryScore += getExpiryUrgencyScore(stock.expiryDate);
  }
});

score += expiryScore * 2;

      const canMake =
        required.length > 0 &&
        missing.length === 0 &&
        shortageByAmount.length === 0 &&
        missingSeasonings.length === 0;

      return {
        name,
        required,
        optional,
        requiredAmount,
        parsedRequiredAmount,
        shortageByAmount,
        genre,
        time,
        difficulty,
        tags,
        seasonings,
        missingSeasonings,
        missing,
        canMake,
        expiryScore,
        recommendReason: buildRecommendReason({
          canMake,
          missingCount: missing.length,
          shortageCount: shortageByAmount.length,
          matchedOptionalCount: matchedOptional.length,
          cookingTime,
          expiryScore
        }),
        recommendTags: buildRecommendTags({
          canMake,
          missingCount: missing.length,
          shortageCount: shortageByAmount.length,
          matchedOptionalCount: matchedOptional.length,
          cookingTime,
          expiryScore
        }),
        score
      };
    });

  menus.sort((a, b) => b.score - a.score);

  return menus;
}

function parseAmountText(amountText) {
  if (!amountText) {
    return [];
  }

  return String(amountText).split(",").map(item => {
    const parts = item.split(":");

    return {
      name: parts[0] ? parts[0].trim() : "",
      amountText: parts[1] ? parts[1].trim() : "",
      amount: parts[1] ? Number(String(parts[1]).replace(/[^\d.]/g, "")) : 0,
      unit: parts[1] ? String(parts[1]).replace(/[\d.]/g, "").trim() : ""
    };
  }).filter(item => item.name);
}

const COOKED_RICE_GRAMS_PER_GO = 330;

function convertCookedRiceToRawRice(item) {
  if (!item) return item;

  const name = String(item.name || '').trim();
  const unit = String(item.unit || '').trim();
  const amount = Number(item.amount) || 0;

  if (name === 'ご飯' && unit === 'g') {
    const rawRiceAmount = amount / COOKED_RICE_GRAMS_PER_GO;

    return {
      name: '米',
      amountText: `${Math.ceil(rawRiceAmount * 100) / 100}合`,
      amount: rawRiceAmount,
      unit: '合',
      originalName: name,
      originalAmountText: item.amountText
    };
  }

  return item;
}

function addMissingItemsToShopping(items) {
  initSheets();

  if (!items || !Array.isArray(items)) {
    throw new Error('追加する食材がありません。');
  }

  items.forEach(item => {
    addShoppingItem({
      name: item.name,
      memo: item.memo || '献立提案から追加'
    });
  });

  return '買い物リストに追加しました';
}

function cookMenu(menu) {
  initSheets();

  if (!menu) {
    throw new Error('献立データがありません。');
  }

  const stockSheet = getSheetByNameOrThrow(SHEETS.STOCK);
  const menuHistorySheet = getSheetByNameOrThrow(SHEETS.MENU_HISTORY);

  const stockValues = stockSheet.getDataRange().getValues();
  const requiredItems = menu.parsedRequiredAmount || [];

  requiredItems.forEach(item => {
    for (let i = 1; i < stockValues.length; i++) {
      const row = stockValues[i];

      const stockName = String(row[1] || '').trim();
      const stockQuantity = Number(row[2]) || 0;
      const stockUnit = String(row[3] || '').trim();

      if (stockName === item.name && stockUnit === item.unit) {
        const newQuantity = Math.max(0, stockQuantity - item.amount);
        const minStock = Number(row[7]) || 0;
        const newStatus = getStatus(newQuantity, minStock);

        stockSheet.getRange(i + 1, 3).setValue(newQuantity);
        stockSheet.getRange(i + 1, 9).setValue(newStatus);

        break;
      }
    }
  });

  menuHistorySheet.appendRow([
    createId(),
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
    menu.name,
    menu.requiredAmount ? menu.requiredAmount.join(', ') : '',
    ''
  ]);

  return '献立を作成済みにしました';
}

function getMenuHistory() {
  initSheets();

  const sheet = getSheetByNameOrThrow(SHEETS.MENU_HISTORY);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  return sheet.getRange(2, 1, lastRow - 1, 5).getValues()
    .map(row => ({
      id: row[0],
      date: formatDateTime(row[1]),
      menuName: row[2],
      usedItems: row[3],
      memo: row[4]
    }))
    .reverse();
}

function formatDateTime(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  }

  return String(value);
}

function getMenuGenres() {
  initSheets();

  const sheet = getSheetByNameOrThrow(SHEETS.MENU);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 5, lastRow - 1, 1).getValues();

  const genres = values
    .map(row => String(row[0] || '').trim())
    .filter(genre => genre !== '');

  return [...new Set(genres)].sort();
}

function toggleMenuDetail(card) {
  card.classList.toggle('open');

  const arrow = card.querySelector('.menu-arrow');

  if (card.classList.contains('open')) {
    arrow.textContent = '▲';
  } else {
    arrow.textContent = '▼';
  }
}

function getExpiryUrgencyScore(expiryDateText) {
  if (!expiryDateText) return 0;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const expiry = parseDateOnly(expiryDateText);
  if (!expiry) return 0;

  const diffDays = Math.floor((expiry.getTime() - today.getTime()) / 86400000);

  if (diffDays < 0) return 0;
  if (diffDays === 0) return 10;
  if (diffDays === 1) return 8;
  if (diffDays === 2) return 6;
  if (diffDays === 3) return 4;

  return 0;
}

function buildRecommendReason(info) {
  const reasons = [];

  if (info.expiryScore > 0) {
    reasons.push('期限が近い食材を使えます');
  }

  if (info.canMake) {
    reasons.push('在庫だけで作れます');
  }

  if (info.cookingTime <= 15) {
    reasons.push('短時間で作れます');
  }

  if (info.matchedOptionalCount > 0) {
    reasons.push('任意食材も活用できます');
  }

  if (!info.canMake && (info.missingCount + info.shortageCount) <= 2) {
    reasons.push('あと少しで作れます');
  }

  if (reasons.length === 0) {
    reasons.push('在庫状況から候補に入りました');
  }

  return reasons.join('・');
}

function buildRecommendTags(info) {
  const tags = [];

  if (info.expiryScore > 0) {
    tags.push('期限優先');
  }

  if (info.canMake) {
    tags.push('在庫あり');
  }

  if (info.cookingTime <= 15) {
    tags.push('時短');
  }

  if (info.matchedOptionalCount > 0) {
    tags.push('任意食材あり');
  }

  if (!info.canMake && (info.missingCount + info.shortageCount) <= 2) {
    tags.push('あと少し');
  }

  if (tags.length === 0) {
    tags.push('候補');
  }

  return tags;
}

function registerDefaultSeasonings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('在庫一覧');

  if (!sheet) {
    throw new Error('在庫一覧シートが見つかりません');
  }

  const defaultSeasonings = [
    '塩', '砂糖', 'しょうゆ', 'みそ', '酢', 'みりん', '料理酒', 'サラダ油', 'ごま油', 'オリーブオイル', 'こしょう', '鶏ガラスープの素', '和風だし', 'コンソメ', 'めんつゆ', 'マヨネーズ', 'ケチャップ', 'ソース', 'ポン酢', 'チューブにんにく', 'チューブしょうが', 'カレー粉', '片栗粉', '小麦粉', 'ラー油', '七味唐辛子'
  ];

  const lastRow = sheet.getLastRow();

  let existingNames = [];

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

    existingNames = values
      .map(row => String(row[1]).trim())
      .filter(name => name !== '');
  }

  const now = new Date();
  const rowsToAdd = [];

  defaultSeasonings.forEach(name => {
    if (!existingNames.includes(name)) {
      rowsToAdd.push([
  createId(),
  name,
  1,
  'あり',
  '',
  now,
  '',
  1,
  '在庫あり',
  '調味料',
  '常温'
]);
    }
  });

  if (rowsToAdd.length > 0) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 11)
      .setValues(rowsToAdd);
  }

  return {
    addedCount: rowsToAdd.length,
    skippedCount: defaultSeasonings.length - rowsToAdd.length
  };
}

function toggleSeasoningStock(foodId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheetByNameOrThrow(SHEETS.STOCK);
    const found = findRowById(sheet, foodId, 11);

    if (!found) {
      throw new Error('対象の調味料が見つかりませんでした。');
    }

    const row = found.values;
    const manageType = String(row[9] || '').trim();

    if (manageType !== '調味料') {
      throw new Error('調味料ではない食品は切り替えできません。');
    }

    const currentQuantity = Number(row[2]) || 0;
    const newQuantity = currentQuantity > 0 ? 0 : 1;
    const newStatus = newQuantity > 0 ? '在庫あり' : '在庫なし';

    sheet.getRange(found.rowNumber, 3, 1, 1).setValue(newQuantity);
    sheet.getRange(found.rowNumber, 9, 1, 1).setValue(newStatus);

    return {
      message: newQuantity > 0 ? 'ありにしました' : 'なしにしました',
      quantity: newQuantity,
      status: newStatus
    };
  } finally {
    lock.releaseLock();
  }
}

function repairStorageLocations() {
  initSheets();

  const sheet = getSheetByNameOrThrow(SHEETS.STOCK);

  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0];

  let storageColumnIndex = headers.indexOf('保存場所') + 1;

  if (storageColumnIndex === 0) {
    storageColumnIndex = sheet.getLastColumn() + 1;
    sheet.getRange(1, storageColumnIndex).setValue('保存場所');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return {
      message: '在庫データがありません。',
      updatedCount: 0
    };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, Math.max(storageColumnIndex, 10)).getValues();
  const updates = [];

  values.forEach((row, index) => {
    const name = String(row[1] || '').trim();
    const manageType = String(row[9] || '通常').trim();
    const storageLocation = String(row[storageColumnIndex - 1] || '').trim();

    if (!name || storageLocation) return;

    updates.push({
      rowNumber: index + 2,
      value: getDefaultStorageLocation(manageType)
    });
  });

  updates.forEach(update => {
    sheet.getRange(update.rowNumber, storageColumnIndex).setValue(update.value);
  });

  return {
    message: `保存場所の自動補正が完了しました。更新：${updates.length}件`,
    updatedCount: updates.length
  };
}

function portionFreezeFood(foodId, portionData) {
  initSheets();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheetByNameOrThrow(SHEETS.STOCK);

    const found = findRowById(sheet, foodId, 11);

    if (!found) {
      throw new Error('対象の食品が見つかりませんでした。');
    }

    const row = found.values;

    const foodName = String(row[1] || '').trim();
    const currentQuantity = Number(row[2]) || 0;
    const unit = String(row[3] || '').trim();
    const purchaseDate = row[5] ? normalizeDateText(row[5]) : todayText();
    const minStock = Number(row[7]) || 0;
    const manageType = String(row[9] || '通常').trim();

    if (manageType === '調味料') {
      throw new Error('調味料は小分け冷凍の対象外です。');
    }

    const portionQuantity = toNumber(portionData.portionQuantity, '小分け量', { min: 0.000001 });
    const portionCount = toNumber(portionData.portionCount, '個数', { min: 1 });
    const portionExpiryDate = portionData.expiryDate
      ? normalizeDateText(portionData.expiryDate)
      : '';

    const totalPortionQuantity = portionQuantity * portionCount;

    if (totalPortionQuantity > currentQuantity) {
      throw new Error(
        `小分けする合計量が現在の在庫数を超えています。\n` +
        `現在：${currentQuantity}${unit}\n` +
        `小分け合計：${totalPortionQuantity}${unit}`
      );
    }

    const remainingQuantity = currentQuantity - totalPortionQuantity;
    const remainingStatus = getStatus(remainingQuantity, minStock);

    sheet.getRange(found.rowNumber, 3, 1, 1).setValue(remainingQuantity);
    sheet.getRange(found.rowNumber, 9, 1, 1).setValue(remainingStatus);

    const rowsToAdd = [];

    for (let i = 0; i < portionCount; i++) {
      rowsToAdd.push([
        createId(),
        foodName,
        portionQuantity,
        unit,
        '',
        purchaseDate,
        portionExpiryDate,
        0,
        getStatus(portionQuantity, 0),
        manageType,
        '冷凍'
      ]);
    }

    if (rowsToAdd.length > 0) {
      sheet
        .getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 11)
        .setValues(rowsToAdd);
    }

    return {
      message: '小分け冷凍しました',
      foodName,
      remainingQuantity,
      unit,
      addedCount: portionCount,
      portionQuantity,
      totalPortionQuantity
    };
  } finally {
    lock.releaseLock();
  }
}
