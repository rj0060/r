// منظومة عائلة أبو رجيلة الإلكترونية 2026
// نسخة محسنة للسرعة والأمان: قراءات جماعية، كاش للإحصائيات، وتصدير عند الطلب.

const SPREADSHEET_ID_PROPERTY = "FAMILY_SPREADSHEET_ID";
const SPREADSHEET_READY_PROPERTY_PREFIX = "FAMILY_SPREADSHEET_READY_";
const SPREADSHEET_NAME = "منظومة عائلة أبو رجيلة الإلكترونية";
const PARENT_SHEET_NAME = "الآباء";
const CHILDREN_SHEET_NAME = "الأبناء";
const TIMEZONE = "GMT+3";
const ADMIN_TOKEN_PREFIX = "admin_session_";
const ADMIN_TOKEN_TTL_SECONDS = 21600; // 6 ساعات، الحد الأعلى في CacheService.
const ADMIN_STATS_CACHE_KEY = "admin_stats_v2";
const ADMIN_STATS_CACHE_TTL_SECONDS = 120;
const ENABLE_SLOW_ID_SCAN_FALLBACK = false;
const APP_VERSION = "2026.09.25.1";

const PARENT_HEADERS = [
  "رقم الهوية",
  "الاسم الكامل",
  "تاريخ الميلاد",
  "العمر",
  "الحالة الاجتماعية",
  "مكان النزوح الحالي",
  "عدد أفراد الأسرة",
  "اسم الزوجة",
  "رقم هوية الزوجة",
  "تاريخ ميلاد الزوجة",
  "رقم الجوال",
  "رقم جوال بديل",
  "رقم الهاتف",
  "البنك",
  "رقم الحساب البنكي",
  "اسم صاحب الحساب"
];

const CHILDREN_HEADERS = [
  "رقم هوية رب الأسرة",
  "ملاحظات",
  "اسم الابن",
  "رقم هوية الابن",
  "تاريخ الميلاد",
  "الجنس",
  "التخصص الجامعي",
  "الجامعة",
  "السنة الدراسية الحالية"
];

function doGet(e) {
  return jsonResponse_({
    status: "success",
    message: "المنظومة تعمل بكفاءة",
    version: APP_VERSION
  });
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    const data = JSON.parse(body);
    const action = data.action;

    switch (action) {
      case "login":
        return loginUser(data.nationalId);

      case "adminLogin":
        return adminLogin(data.password);

      case "health":
        return healthCheck();

      case "updateParent":
        if (data.adminToken) requireAdminToken_(data.adminToken);
        return updateParentData(data.nationalId, data.fields);

      case "addChild":
        if (data.adminToken) requireAdminToken_(data.adminToken);
        return addChildData(data.parentId, data.childData);

      case "updateChild":
        if (data.adminToken) requireAdminToken_(data.adminToken);
        return updateChildData(data.parentId, data.oldChildId, data.childData);

      case "getAdminStats":
        requireAdminToken_(data.adminToken);
        return getAdminStats();

      case "exportData":
        requireAdminToken_(data.adminToken);
        return exportData(data.type);

      case "adminSearch":
        requireAdminToken_(data.adminToken);
        return loginUser(data.searchId);

      case "addFamily":
        requireAdminToken_(data.adminToken);
        return addFamilyData(data.familyData);

      case "prepareSheetsForSpeed":
        requireAdminToken_(data.adminToken);
        return prepareSheetsForSpeed();

      case "setupSpreadsheet":
        requireAdminToken_(data.adminToken);
        return setupSpreadsheet();

      default:
        return errorResponse("الإجراء غير معروف");
    }
  } catch (err) {
    return errorResponse(err && err.message ? err.message : "خطأ في الخادم");
  }
}

function loginUser(nationalId) {
  const userData = getUserData_(nationalId);
  return jsonResponse_(Object.assign({ status: "success" }, userData));
}

function updateParentData(nationalId, fields) {
  const lock = getWriteLock_();
  lock.waitLock(10000);

  try {
    const ss = getSpreadsheet_();
    const sheet = getSheet_(ss, 0, "شيت الآباء");
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return errorResponse("شيت الآباء فارغ");

    const searchId = validateId_(nationalId, "رقم هوية رب الأسرة غير صحيح");
    const rowIndex = findRowById_(sheet, searchId, 1);
    if (rowIndex === -1) return errorResponse("لم يتم العثور على السجل لتحديثه");

    const headers = getHeaders_(sheet);
    const headerMap = getHeaderMap_(headers);
    const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const updates = [];

    Object.keys(fields || {}).forEach(function(key) {
      const cleanKey = String(key).trim();
      const colIndex = headerMap[cleanKey];

      // لا نسمح بتعديل رقم هوية رب الأسرة من هذه الواجهة لأنه مفتاح البحث.
      if (colIndex === undefined || colIndex === 0) return;

      rowValues[colIndex] = fields[key] === null || fields[key] === undefined ? "" : fields[key];
      updates.push({ col: colIndex, key: cleanKey });
    });

    if (updates.length === 0) {
      return jsonResponse_({
        status: "success",
        message: "لا توجد حقول قابلة للتحديث"
      });
    }

    updates.sort(function(a, b) { return a.col - b.col; });

    updates.forEach(function(update) {
      if (update.key.indexOf("تاريخ") !== -1) {
        sheet.getRange(rowIndex, update.col + 1).setNumberFormat("@");
      }
    });

    setSparseRowValues_(sheet, rowIndex, rowValues, updates);
    invalidateAdminStatsCache_();

    return jsonResponse_({
      status: "success",
      message: "تم تحديث البيانات بنجاح",
      fields: fields
    });
  } finally {
    lock.releaseLock();
  }
}

function updateChildData(parentId, oldChildId, childData) {
  const lock = getWriteLock_();
  lock.waitLock(10000);

  try {
    const ss = getSpreadsheet_();
    const sheet = getSheet_(ss, 1, "شيت الأبناء");
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return errorResponse("شيت الأبناء فارغ");

    const targetParentId = validateId_(parentId, "رقم هوية رب الأسرة غير صحيح");
    const targetChildId = validateId_(oldChildId, "رقم هوية الابن القديم غير صحيح");
    const newChildId = validateId_(childData && childData.id, "رقم هوية الابن غير صحيح");
    const childName = String(childData && childData.name ? childData.name : "").trim();
    const birthDate = String(childData && childData.birthDate ? childData.birthDate : "").trim();

    if (!childName) return errorResponse("يرجى كتابة اسم الابن بالكامل");
    if (!birthDate) return errorResponse("يرجى تحديد تاريخ ميلاد الابن");

    const targetRows = findRowsByExactText_(sheet, 4, targetChildId);
    let rowIndex = -1;

    for (let i = 0; i < targetRows.length; i++) {
      const rowNumber = targetRows[i];
      const rowParentId = normalizeId_(sheet.getRange(rowNumber, 1).getDisplayValue());

      if (rowParentId === targetParentId) {
        rowIndex = rowNumber;
        break;
      }
    }

    if (rowIndex === -1) return errorResponse("لم يتم العثور على سجل الابن لتعديله");

    const duplicateRows = findRowsByExactText_(sheet, 4, newChildId);
    for (let j = 0; j < duplicateRows.length; j++) {
      if (duplicateRows[j] !== rowIndex) return errorResponse("رقم هوية الابن موجود مسبقاً ولا يمكن تكراره");
    }

    const rowValues = [[
      childName,
      newChildId,
      birthDate,
      String(childData.gender || ""),
      String(childData.major || ""),
      String(childData.university || ""),
      String(childData.year || "")
    ]];

    sheet.getRange(rowIndex, 4, 1, 2).setNumberFormat("@");
    sheet.getRange(rowIndex, 3, 1, 7).setValues(rowValues);
    invalidateAdminStatsCache_();

    return jsonResponse_({
      status: "success",
      message: "تم تحديث بيانات الابن بنجاح",
      child: getSerializedRow_(sheet, rowIndex)
    });
  } finally {
    lock.releaseLock();
  }
}

function addChildData(parentId, childData) {
  const lock = getWriteLock_();
  lock.waitLock(10000);

  try {
    const ss = getSpreadsheet_();
    const sheet = getSheet_(ss, 1, "شيت الأبناء");
    const lastRow = sheet.getLastRow();
    const parentNationalId = validateId_(parentId, "رقم هوية رب الأسرة غير صحيح");
    const childId = validateId_(childData && childData.id, "رقم هوية الابن غير صحيح");
    const childName = String(childData && childData.name ? childData.name : "").trim();
    const birthDate = String(childData && childData.birthDate ? childData.birthDate : "").trim();

    if (!childName) return errorResponse("يرجى كتابة اسم الابن بالكامل");
    if (!birthDate) return errorResponse("يرجى تحديد تاريخ ميلاد الابن");

    if (lastRow > 1 && findRowsByExactText_(sheet, 4, childId).length > 0) {
      return errorResponse("هذا الطفل مسجل مسبقاً في كشوفات العائلة الإلكترونية ولا يمكن تكراره");
    }

    const targetRow = lastRow + 1;
    const rowValues = [[
      parentNationalId,
      "",
      childName,
      childId,
      birthDate,
      String(childData.gender || ""),
      String(childData.major || ""),
      String(childData.university || ""),
      String(childData.year || "")
    ]];

    sheet.getRange(targetRow, 1, 1, 1).setNumberFormat("@");
    sheet.getRange(targetRow, 4, 1, 2).setNumberFormat("@");
    sheet.getRange(targetRow, 1, 1, rowValues[0].length).setValues(rowValues);
    invalidateAdminStatsCache_();

    return jsonResponse_({
      status: "success",
      message: "تم تسجيل الابن بنجاح",
      child: getSerializedRow_(sheet, targetRow)
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * ينشئ سجل عائلة جديداً من لوحة الإدارة فقط.
 * رقم هوية رب الأسرة هو المفتاح الرئيسي ولا يُسمح بتكراره.
 */
function addFamilyData(familyData) {
  const lock = getWriteLock_();
  lock.waitLock(10000);

  try {
    const data = familyData || {};
    const nationalId = validateId_(data.nationalId, "رقم هوية رب الأسرة يجب أن يتكون من 9 أرقام");
    const fullName = String(data.fullName || "").trim();
    const birthDate = String(data.birthDate || "").trim();

    if (fullName.length < 3) throw new Error("يرجى كتابة اسم رب الأسرة بالكامل");
    if (!birthDate) throw new Error("يرجى تحديد تاريخ ميلاد رب الأسرة");

    const ss = getSpreadsheet_();
    const sheet = getSheet_(ss, 0, "شيت الآباء");
    ensureHeaders_(sheet, PARENT_HEADERS);

    if (findRowById_(sheet, nationalId, 1) !== -1) {
      throw new Error("رقم الهوية مسجل مسبقاً ولا يمكن إنشاء عائلة مكررة");
    }

    const wifeIdRaw = normalizeId_(data.wifeId);
    if (wifeIdRaw && !/^\d{9}$/.test(wifeIdRaw)) {
      throw new Error("رقم هوية الزوجة يجب أن يتكون من 9 أرقام أو يُترك فارغاً");
    }

    const familyCountRaw = String(data.familyCount || "").trim();
    const familyCount = familyCountRaw === "" ? "" : Math.max(1, parseInt(familyCountRaw, 10) || 1);
    const targetRow = sheet.getLastRow() + 1;
    const rowValues = [[
      nationalId,
      fullName,
      birthDate,
      calculateAgeFromDate_(birthDate),
      String(data.status || "").trim(),
      String(data.displacement || "").trim(),
      familyCount,
      String(data.wifeName || "").trim(),
      wifeIdRaw,
      String(data.wifeBirth || "").trim(),
      String(data.phone || "").trim(),
      String(data.altPhone || "").trim(),
      String(data.telephone || "").trim(),
      String(data.bank || "").trim(),
      String(data.bankAccount || "").trim(),
      String(data.bankOwner || "").trim()
    ]];

    sheet.getRange(targetRow, 1).setNumberFormat("@");
    sheet.getRange(targetRow, 3).setNumberFormat("@");
    sheet.getRange(targetRow, 9, 1, 2).setNumberFormat("@");
    sheet.getRange(targetRow, 1, 1, PARENT_HEADERS.length).setValues(rowValues);
    invalidateAdminStatsCache_();

    return jsonResponse_({
      status: "success",
      message: "تم إنشاء سجل العائلة بنجاح",
      parent: getSerializedRow_(sheet, targetRow)
    });
  } finally {
    lock.releaseLock();
  }
}

function getAdminStats() {
  const cache = CacheService.getScriptCache();
  const cachedStats = cache.get(ADMIN_STATS_CACHE_KEY);

  if (cachedStats) {
    return jsonResponse_({
      status: "success",
      stats: JSON.parse(cachedStats),
      cached: true
    });
  }

  const ss = getSpreadsheet_();
  const parentSheet = getSheet_(ss, 0, "شيت الآباء");
  const childrenSheet = getSheet_(ss, 1, "شيت الأبناء");
  const parentData = parentSheet.getDataRange().getDisplayValues();
  const parentHeaders = normalizeHeaders_(parentData[0] || []);
  const childrenData = childrenSheet.getDataRange().getDisplayValues();
  const childrenHeaders = normalizeHeaders_(childrenData[0] || []);

  let totalFamilies = Math.max(parentData.length - 1, 0);
  let totalDisplaced = 0;
  let totalWidows = 0;
  let totalDivorced = 0;
  let totalMembersCount = 0;

  const statusCol = parentHeaders.indexOf("الحالة الاجتماعية");
  const displaceCol = parentHeaders.indexOf("مكان النزوح الحالي");
  const countCol = parentHeaders.indexOf("عدد أفراد الأسرة");

  for (let i = 1; i < parentData.length; i++) {
    const row = parentData[i];

    if (displaceCol !== -1 && String(row[displaceCol] || "").trim() !== "") totalDisplaced++;
    if (statusCol !== -1 && String(row[statusCol] || "").indexOf("أرمل") !== -1) totalWidows++;
    if (statusCol !== -1 && String(row[statusCol] || "").indexOf("مطلق") !== -1) totalDivorced++;

    const membersCount = parseInt(row[countCol], 10);
    if (countCol !== -1 && !isNaN(membersCount)) totalMembersCount += membersCount;
  }

  let totalUniStudents = 0;
  const majorCol = childrenHeaders.indexOf("التخصص الجامعي");

  for (let j = 1; j < childrenData.length; j++) {
    if (majorCol !== -1 && String(childrenData[j][majorCol] || "").trim() !== "") {
      totalUniStudents++;
    }
  }

  const stats = {
    totalFamilies: totalFamilies,
    totalMembers: totalMembersCount,
    totalDisplaced: totalDisplaced,
    totalWidows: totalWidows,
    totalDivorced: totalDivorced,
    totalUniStudents: totalUniStudents
  };

  cache.put(ADMIN_STATS_CACHE_KEY, JSON.stringify(stats), ADMIN_STATS_CACHE_TTL_SECONDS);

  return jsonResponse_({
    status: "success",
    stats: stats,
    cached: false
  });
}

function exportData(type) {
  const ss = getSpreadsheet_();
  const normalizedType = String(type || "").trim();
  const sheet = normalizedType === "children"
    ? getSheet_(ss, 1, "شيت الأبناء")
    : getSheet_(ss, 0, "شيت الآباء");

  return jsonResponse_({
    status: "success",
    data: serializeTable_(sheet)
  });
}

function adminLogin(password) {
  if (!verifyAdminPassword_(password)) {
    return errorResponse("كلمة المرور غير صحيحة");
  }

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(ADMIN_TOKEN_PREFIX + token, "1", ADMIN_TOKEN_TTL_SECONDS);

  return jsonResponse_({
    status: "success",
    token: token,
    expiresInSeconds: ADMIN_TOKEN_TTL_SECONDS
  });
}

function healthCheck() {
  return jsonResponse_({
    status: "success",
    version: APP_VERSION,
    timestamp: new Date().toISOString()
  });
}

function prepareSheetsForSpeed() {
  const lock = getWriteLock_();
  lock.waitLock(10000);

  try {
    const ss = getSpreadsheet_();
    const parentSheet = getSheet_(ss, 0, "شيت الآباء");
    const childrenSheet = getSheet_(ss, 1, "شيت الأبناء");

    ensureHeaders_(parentSheet, PARENT_HEADERS);
    ensureHeaders_(childrenSheet, CHILDREN_HEADERS);
    setColumnAsText_(parentSheet, 1);
    setColumnAsText_(parentSheet, 9);
    setColumnAsText_(childrenSheet, 1);
    setColumnAsText_(childrenSheet, 4);
    clearRuntimeCaches_();

    return jsonResponse_({
      status: "success",
      message: "تم تجهيز أعمدة الهوية للبحث السريع",
      version: APP_VERSION,
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl()
    });
  } finally {
    lock.releaseLock();
  }
}

function setupSpreadsheet() {
  const lock = getWriteLock_();
  lock.waitLock(10000);

  try {
    const ss = getSpreadsheet_(true);
    return jsonResponse_({
      status: "success",
      message: "تم تجهيز ملف Google Sheets تلقائياً",
      version: APP_VERSION,
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl()
    });
  } finally {
    lock.releaseLock();
  }
}

function getUserData_(nationalId) {
  const ss = getSpreadsheet_();
  const parentSheet = getSheet_(ss, 0, "شيت الآباء");
  const childrenSheet = getSheet_(ss, 1, "شيت الأبناء");
  const lastRowParent = parentSheet.getLastRow();
  const lastColParent = parentSheet.getLastColumn();

  if (lastRowParent <= 1 || lastColParent === 0) {
    throw new Error("شيت الآباء فارغ");
  }

  const searchId = validateId_(nationalId, "رقم الهوية غير صحيح");
  const userRowIndex = findRowById_(parentSheet, searchId, 1);

  if (userRowIndex === -1) {
    throw new Error("رقم الهوية غير مدرج بملفات العائلة");
  }

  const parent = getSerializedRow_(parentSheet, userRowIndex);
  const userChildren = [];
  const lastRowChildren = childrenSheet.getLastRow();
  const lastColChildren = childrenSheet.getLastColumn();

  if (lastRowChildren > 1 && lastColChildren > 0) {
    const childrenHeaders = getHeaders_(childrenSheet);
    const matchingRows = findRowsByExactText_(childrenSheet, 1, searchId);

    groupRowNumbers_(matchingRows).forEach(function(group) {
      const range = childrenSheet.getRange(group.start, 1, group.count, lastColChildren);
      const displays = range.getDisplayValues();

      for (let rowIndex = 0; rowIndex < displays.length; rowIndex++) {
        userChildren.push(serializeRow_(childrenHeaders, displays[rowIndex]));
      }
    });
  }

  return {
    parent: parent,
    children: userChildren
  };
}

function getSpreadsheet_(forceSetup) {
  const properties = PropertiesService.getScriptProperties();
  const savedId = properties.getProperty(SPREADSHEET_ID_PROPERTY);
  let spreadsheet = null;

  if (savedId) {
    try {
      spreadsheet = SpreadsheetApp.openById(savedId);
    } catch (err) {
      properties.deleteProperty(SPREADSHEET_ID_PROPERTY);
    }
  }

  if (!spreadsheet) {
    spreadsheet = getActiveSpreadsheet_();
  }

  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create(SPREADSHEET_NAME);
  }

  properties.setProperty(SPREADSHEET_ID_PROPERTY, spreadsheet.getId());

  const readyKey = SPREADSHEET_READY_PROPERTY_PREFIX + spreadsheet.getId() + "_" + APP_VERSION;
  if (forceSetup || properties.getProperty(readyKey) !== "1") {
    ensureSpreadsheetStructure_(spreadsheet);
    properties.setProperty(readyKey, "1");
  }

  return spreadsheet;
}

function getActiveSpreadsheet_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    return null;
  }
}

function getWriteLock_() {
  return LockService.getDocumentLock() || LockService.getScriptLock();
}

function getSheet_(ss, index, label) {
  const expectedName = index === 0 ? PARENT_SHEET_NAME : index === 1 ? CHILDREN_SHEET_NAME : "";
  let sheet = expectedName ? ss.getSheetByName(expectedName) : ss.getSheets()[index];

  if (!sheet) {
    ensureSpreadsheetStructure_(ss);
    sheet = expectedName ? ss.getSheetByName(expectedName) : ss.getSheets()[index];
  }

  if (!sheet) throw new Error(label + " غير موجود");
  return sheet;
}

function ensureSpreadsheetStructure_(ss) {
  const parentSheet = ensureNamedSheet_(ss, PARENT_SHEET_NAME, 0);
  const childrenSheet = ensureNamedSheet_(ss, CHILDREN_SHEET_NAME, 1);

  moveSheetToPosition_(ss, parentSheet, 1);
  moveSheetToPosition_(ss, childrenSheet, 2);

  ensureHeaders_(parentSheet, PARENT_HEADERS);
  ensureHeaders_(childrenSheet, CHILDREN_HEADERS);

  parentSheet.setFrozenRows(1);
  childrenSheet.setFrozenRows(1);
  setColumnAsText_(parentSheet, 1);
  setColumnAsText_(parentSheet, 9);
  setColumnAsText_(childrenSheet, 1);
  setColumnAsText_(childrenSheet, 4);
}

function ensureNamedSheet_(ss, sheetName, preferredIndex) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;

  const candidate = ss.getSheets()[preferredIndex];
  if (candidate && !isSystemSheetName_(candidate.getName())) {
    candidate.setName(sheetName);
    return candidate;
  }

  return ss.insertSheet(sheetName, preferredIndex);
}

function isSystemSheetName_(sheetName) {
  return sheetName === PARENT_SHEET_NAME || sheetName === CHILDREN_SHEET_NAME;
}

function moveSheetToPosition_(ss, sheet, position) {
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(position);
}

function ensureHeaders_(sheet, expectedHeaders) {
  const requiredCols = expectedHeaders.length;

  if (sheet.getMaxColumns() < requiredCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
  }

  const lastCol = Math.max(sheet.getLastColumn(), requiredCols);
  const firstRow = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const hasHeaders = firstRow.some(function(value) {
    return String(value || "").trim() !== "";
  });

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, requiredCols).setValues([expectedHeaders]);
    styleHeaderRow_(sheet, requiredCols);
    return;
  }

  const existing = {};
  firstRow.forEach(function(header) {
    const cleanHeader = String(header || "").trim();
    if (cleanHeader) existing[cleanHeader] = true;
  });

  const missingHeaders = expectedHeaders.filter(function(header) {
    return !existing[header];
  });

  if (missingHeaders.length > 0) {
    const startCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]);
    styleHeaderRow_(sheet, startCol + missingHeaders.length - 1);
  }
}

function styleHeaderRow_(sheet, columnsCount) {
  sheet.getRange(1, 1, 1, columnsCount)
    .setFontWeight("bold")
    .setBackground("#f4ebd9");
}

function getHeaders_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];

  const cache = CacheService.getScriptCache();
  const cacheKey = "headers_" + sheet.getSheetId() + "_" + lastCol;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const headers = normalizeHeaders_(sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]);
  cache.put(cacheKey, JSON.stringify(headers), 1800);
  return headers;
}

function getHeaderMap_(headers) {
  const map = {};
  headers.forEach(function(header, index) {
    if (header && map[header] === undefined) map[header] = index;
  });
  return map;
}

function normalizeHeaders_(headers) {
  return headers.map(function(header) {
    return String(header || "").trim();
  });
}

function findRowById_(sheet, nationalId, columnIndex) {
  const rows = findRowsByExactText_(sheet, columnIndex, nationalId);
  return rows.length ? rows[0] : -1;
}

function findRowsByExactText_(sheet, columnIndex, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const searchValue = normalizeId_(value);
  const range = sheet.getRange(2, columnIndex, lastRow - 1, 1);
  const textFinderRows = range
    .createTextFinder(searchValue)
    .matchEntireCell(true)
    .findAll()
    .map(function(cell) {
      return cell.getRow();
    });

  if (textFinderRows.length > 0) {
    return textFinderRows;
  }

  if (!ENABLE_SLOW_ID_SCAN_FALLBACK) {
    return [];
  }

  const displays = range.getDisplayValues();
  const rows = [];

  for (let i = 0; i < displays.length; i++) {
    if (normalizeId_(displays[i][0]) === searchValue) {
      rows.push(i + 2);
    }
  }

  return rows;
}

function getSerializedRow_(sheet, rowIndex) {
  const lastCol = sheet.getLastColumn();
  const headers = getHeaders_(sheet);
  const range = sheet.getRange(rowIndex, 1, 1, lastCol);
  const displays = range.getDisplayValues()[0];
  return serializeRow_(headers, displays);
}

function serializeRow_(headers, values, displays) {
  const obj = {};
  const displayValues = displays || values;

  headers.forEach(function(header, index) {
    if (!header) return;
    obj[header] = serializeCell_(values[index], displayValues[index]);
  });

  return obj;
}

function serializeTable_(sheet) {
  const range = sheet.getDataRange();
  const displays = range.getDisplayValues();

  return displays.map(function(row, rowIndex) {
    return row.map(function(value) {
      const text = value === null || value === undefined ? "" : String(value);
      return rowIndex === 0 ? text.trim() : text;
    });
  });
}

function groupRowNumbers_(rows) {
  if (rows.length === 0) return [];

  const groups = [];
  let start = rows[0];
  let previous = rows[0];

  for (let i = 1; i <= rows.length; i++) {
    const current = i < rows.length ? rows[i] : null;

    if (current === previous + 1) {
      previous = current;
      continue;
    }

    groups.push({
      start: start,
      count: previous - start + 1
    });

    if (current !== null) {
      start = current;
      previous = current;
    }
  }

  return groups;
}

function serializeCell_(value, displayValue) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, TIMEZONE, "yyyy-MM-dd");
  }

  if (displayValue !== null && displayValue !== undefined) {
    return displayValue;
  }

  return value === null || value === undefined ? "" : value;
}

function setSparseRowValues_(sheet, rowIndex, rowValues, updates) {
  let start = updates[0].col;
  let previous = updates[0].col;

  for (let i = 1; i <= updates.length; i++) {
    const current = i < updates.length ? updates[i].col : null;

    if (current === previous + 1) {
      previous = current;
      continue;
    }

    sheet.getRange(rowIndex, start + 1, 1, previous - start + 1)
      .setValues([rowValues.slice(start, previous + 1)]);

    if (current !== null) {
      start = current;
      previous = current;
    }
  }
}

function validateId_(value, message) {
  const id = normalizeId_(value);
  if (!/^\d{9}$/.test(id)) {
    throw new Error(message || "رقم الهوية يجب أن يتكون من 9 أرقام");
  }
  return id;
}

function normalizeId_(value) {
  return String(value === null || value === undefined ? "" : value).replace(/\D/g, "").trim();
}

function calculateAgeFromDate_(dateValue) {
  const parts = String(dateValue || "").split("-");
  if (parts.length !== 3) return "";

  const birth = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  if (isNaN(birth.getTime())) return "";

  const today = new Date();
  let age = today.getFullYear() - birth.getFullYear();
  const beforeBirthday = today.getMonth() < birth.getMonth() ||
    (today.getMonth() === birth.getMonth() && today.getDate() < birth.getDate());

  if (beforeBirthday) age--;
  return age >= 0 ? age : "";
}

function verifyAdminPassword_(password) {
  const configuredPassword = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD");
  const expectedPassword = configuredPassword || "Eng2026";
  return String(password || "") === expectedPassword;
}

function requireAdminToken_(token) {
  const cleanToken = String(token || "").trim();

  if (!cleanToken || !CacheService.getScriptCache().get(ADMIN_TOKEN_PREFIX + cleanToken)) {
    throw new Error("انتهت صلاحية جلسة الإدارة، يرجى تسجيل الدخول مرة أخرى");
  }

  CacheService.getScriptCache().put(ADMIN_TOKEN_PREFIX + cleanToken, "1", ADMIN_TOKEN_TTL_SECONDS);
}

function invalidateAdminStatsCache_() {
  CacheService.getScriptCache().remove(ADMIN_STATS_CACHE_KEY);
}

function clearRuntimeCaches_() {
  const cache = CacheService.getScriptCache();
  cache.remove(ADMIN_STATS_CACHE_KEY);
}

function setColumnAsText_(sheet, columnIndex) {
  const maxRows = sheet.getMaxRows();
  if (maxRows > 0) {
    sheet.getRange(1, columnIndex, maxRows, 1).setNumberFormat("@");
  }
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse(msg) {
  return jsonResponse_({
    status: "error",
    message: msg
  });
}
