/**
 * ============================================================================
 * نظام إدارة سنتر العباقرة التعليمي — المحرك السحابي الكامل (Google Apps Script)
 * الإصدار: المرحلة الأولى الكاملة (يعمل من script.google.com مباشرة وينشئ الشيت تلقائياً)
 * حساب مالك النظام الافتراضي: اسم المستخدم: mrk | كلمة المرور: 1510
 * ============================================================================
 * طريقة التركيب السريعة من الموبايل (بدون فتح تطبيق Google Sheets):
 * 1. افتح متصفح Chrome على موبايلك، فعّل (الموقع الإلكتروني المخصص للكمبيوتر / Desktop site)
 * 2. ادخل على الرابط: https://script.google.com ثم اضغط (مشروع جديد / New project)
 * 3. امسح أي كود موجود والصق هذا الكود بالكامل من أول سطر لآخر سطر ثم اضغط حفظ (Save).
 * 4. اضغط زر (تشغيل / Run) بالأعلى لتشغيل الدالة الأولى (تشغيل_لأول_مرة_إنشاء_الشيت) ووافق على الصلاحيات،
 *    سيقوم الكود تلقائياً بإنشاء ملف الشيت باسم "قاعدة بيانات سنتر العباقرة" وبداخله الـ 23 ورقة عمل كاملة!
 * 5. اضغط (نشر / Deploy) -> (نشر جديد / New deployment) -> اختر النوع: (تطبيق ويب / Web app)
 *    اجعل Execute as: Me (أنا) واجعل Who has access: Anyone (أي شخص) ثم اضغط Deploy.
 * 6. انسخ رابط الـ Web App وضعه في تبويب (7. الربط السحابي) داخل تطبيق الموبايل.
 * ============================================================================
 */

const SYSTEM_VERSION = "1.1.0-Phase1-Standalone";
const DATABASE_FILE_NAME = "قاعدة بيانات سنتر العباقرة";
const BACKUP_FOLDER_NAME = "نسخ_احتياطية_سنتر_العباقرة";
const PROP_SHEET_ID_KEY = "ABAKERA_SHEET_ID";

/**
 * الدالة الأولى في الملف (تظهر افتراضياً عند الضغط على زر Run / تشغيل في المحرر)
 * تقوم بإنشاء ملف الشيت وتوليد الـ 23 ورقة عمل وحساب الأدمن mrk / 1510 فوراً
 */
function تشغيل_لأول_مرة_إنشاء_الشيت() {
  const ss = initializeDatabaseIfNeeded();
  Logger.log("تم إنشاء وتجهيز قاعدة بيانات سنتر العباقرة (23 ورقة عمل) بنجاح!");
  Logger.log("رابط الشيت على جوجل درايف: " + ss.getUrl());
  return ss.getUrl();
}

/**
 * تعريف الـ 23 ورقة عمل وأعمدتها المعتمدة
 */
const SCHEMA = {
  "الإعدادات": ["المفتاح", "القيمة", "الوصف", "آخر_تحديث"],
  "الأعوام_الدراسية": ["كود_العام", "اسم_العام", "تاريخ_البداية", "تاريخ_النهاية", "نشط"],
  "الفصول_الدراسية": ["كود_الفصل", "كود_العام", "اسم_الفصل", "نشط"],
  "الفروع": ["كود_الفرع", "اسم_الفرع", "العنوان", "الهاتف", "نشط"],
  "القاعات": ["كود_القاعة", "كود_الفرع", "اسم_القاعة", "السعة", "نشط"],
  "المواد_الدراسية": ["كود_المادة", "اسم_المادة", "الاختصار", "نشط"],
  "الصفوف_الدراسية": ["كود_الصف", "اسم_الصف", "المرحلة", "نشط"],
  "أولياء_الأمور": ["كود_ولي_الأمر", "اسم_ولي_الأمر", "رقم_الواتساب", "صلة_القرابة", "ملاحظات"],
  "المدرسون": ["كود_المدرس", "اسم_المدرس", "كود_المادة", "رقم_الهاتف", "نسبة_المدرس", "نشط"],
  "الطلاب": ["كود_الطالب", "اسم_الطالب", "كود_الصف", "هاتف_الطالب", "كود_ولي_الأمر", "هاتف_ولي_الأمر", "تاريخ_التسجيل", "الحالة", "ملاحظات"],
  "المجموعات": ["كود_المجموعة", "اسم_المجموعة", "كود_الصف", "كود_المادة", "كود_المدرس", "كود_القاعة", "الأيام", "الموعد", "الاشتراك_الشهري", "نشط"],
  "اشتراكات_الطلاب": ["كود_الاشتراك", "كود_الدرس", "كود_الطالب", "كود_المجموعة", "كود_المادة", "كود_المدرس", "قيمة_الاشتراك", "تاريخ_الاشتراك", "الحالة"],
  "المحاضرات": ["كود_المحاضرة", "كود_المجموعة", "التاريخ", "الشهر", "عنوان_الحصة", "مسجل_بواسطة"],
  "الحضور_والغياب": ["كود_السجل", "كود_المحاضرة", "كود_المجموعة", "كود_الدرس", "كود_الطالب", "التاريخ", "الشهر", "الحالة", "وقت_التسجيل"],
  "الامتحانات": ["كود_الامتحان", "كود_المجموعة", "كود_المادة", "كود_المدرس", "اسم_الامتحان", "التاريخ", "الدرجة_النهائية", "درجة_النجاح"],
  "نتائج_الامتحانات": ["كود_النتيجة", "كود_الامتحان", "كود_الدرس", "كود_الطالب", "الدرجة", "التقدير", "ملاحظات"],
  "الاستحقاقات_الشهرية": ["كود_الاستحقاق", "كود_الدرس", "كود_الطالب", "كود_المجموعة", "الشهر", "المبلغ_المطلوب", "المبلغ_المسدد", "المتبقي", "الحالة", "تاريخ_التحديث"],
  "المدفوعات": ["رقم_الإيصال", "كود_الدرس", "كود_الطالب", "كود_المجموعة", "الشهر_المسدد", "المبلغ", "تاريخ_الدفع", "طريقة_الدفع", "المستلم", "ملاحظات"],
  "المصروفات": ["كود_المصروف", "الفئة", "البيان", "المبلغ", "التاريخ", "الشهر", "الصارف", "ملاحظات"],
  "الخزينة": ["كود_الحركة", "نوع_الحركة", "المرجع", "البيان", "وارد", "صادر", "الرصيد_اللحظي", "التاريخ", "المستخدم"],
  "تسويات_المدرسين": ["كود_التسوية", "كود_المدرس", "الشهر", "إجمالي_التحصيل", "نسبة_المدرس", "مستحق_المدرس", "نصيب_السنتر", "تاريخ_التسوية", "الحالة"],
  "المستخدمون": ["كود_المستخدم", "اسم_المستخدم", "كلمة_المرور", "الاسم_الكامل", "الدور", "نشط"],
  "سجل_العمليات": ["كود_العملية", "التاريخ_والوقت", "اسم_المستخدم", "الدور", "نوع_العملية", "التفاصيل"]
};

const TRANSACTION_TABLES = [
  "المحاضرات",
  "الحضور_والغياب",
  "الامتحانات",
  "نتائج_الامتحانات",
  "الاستحقاقات_الشهرية",
  "المدفوعات",
  "المصروفات",
  "الخزينة",
  "تسويات_المدرسين",
  "سجل_العمليات"
];

const FULL_RESET_TABLES = [
  "أولياء_الأمور",
  "المدرسون",
  "الطلاب",
  "المجموعات",
  "اشتراكات_الطلاب",
  "المحاضرات",
  "الحضور_والغياب",
  "الامتحانات",
  "نتائج_الامتحانات",
  "الاستحقاقات_الشهرية",
  "المدفوعات",
  "المصروفات",
  "الخزينة",
  "تسويات_المدرسين",
  "سجل_العمليات"
];

/**
 * دالة ذكية للوصول إلى ملف الشيت أو إنشائه تلقائياً على Google Drive
 * تعمل سواء فُتح الكود من script.google.com مباشرة أو من داخل شيت
 */
function getTargetSpreadsheet() {
  // 1. إذا كان الكود مربوطاً بشيت مفتوح بالفعل
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}

  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty(PROP_SHEET_ID_KEY);

  // 2. إذا كان الشيت قد أُنشئ سابقاً ومحفوظ رقم المعرف الخاص به
  if (savedId) {
    try {
      const existingById = SpreadsheetApp.openById(savedId);
      if (existingById) return existingById;
    } catch (e) {}
  }

  // 3. البحث في Google Drive عن شيت بنفس الاسم قبل إنشاء شيت جديد
  try {
    const files = DriveApp.getFilesByName(DATABASE_FILE_NAME);
    if (files.hasNext()) {
      const file = files.next();
      const foundSS = SpreadsheetApp.openById(file.getId());
      props.setProperty(PROP_SHEET_ID_KEY, foundSS.getId());
      return foundSS;
    }
  } catch (e) {}

  // 4. إنشاء ملف Google Sheet جديد تلقائياً على حساب المستخدم
  const createdSS = SpreadsheetApp.create(DATABASE_FILE_NAME);
  props.setProperty(PROP_SHEET_ID_KEY, createdSS.getId());
  return createdSS;
}

/**
 * استقبال طلبات GET
 */
function doGet(e) {
  try {
    const ss = initializeDatabaseIfNeeded();
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "ping";
    if (action === "ping") {
      return jsonResponse({
        status: "ok",
        version: SYSTEM_VERSION,
        center: "سنتر العباقرة التعليمي",
        sheetUrl: ss.getUrl(),
        sheetsCount: Object.keys(SCHEMA).length
      });
    }
    if (action === "pullAll") {
      return jsonResponse({
        status: "ok",
        sheetUrl: ss.getUrl(),
        data: readAllTables(ss)
      });
    }
    return jsonResponse({
      status: "ok",
      message: "تم إنشاء وتجهيز الشيت (23 ورقة عمل) بنجاح",
      sheetUrl: ss.getUrl()
    });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * استقبال طلبات POST (المزامنة، النسخ الاحتياطي، التفريغ، التهيئة)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const ss = initializeDatabaseIfNeeded();
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const action = body.action || "";

    if (action === "initDatabase") {
      return jsonResponse({
        status: "ok",
        message: "تم تهيئة جميع الـ 23 ورقة عمل بنجاح",
        sheetUrl: ss.getUrl(),
        data: readAllTables(ss)
      });
    }

    if (action === "syncSnapshot") {
      if (body.tables) {
        writeSnapshotTables(ss, body.tables);
      }
      logAction(ss, body.user || "mrk", body.role || "مالك النظام", "مزامنة سحابية", "تمت مزامنة البيانات مع Google Sheets");
      return jsonResponse({
        status: "ok",
        message: "تمت المزامنة السحابية بنجاح مع الـ 23 ورقة عمل",
        sheetUrl: ss.getUrl(),
        data: readAllTables(ss)
      });
    }

    if (action === "createBackup") {
      const backupInfo = createDriveBackup(ss, body.reason || "نسخة احتياطية يدوية");
      logAction(ss, body.user || "mrk", body.role || "مالك النظام", "نسخ احتياطي", "تم إنشاء نسخة احتياطية: " + backupInfo.name);
      return jsonResponse({
        status: "ok",
        message: "تم إنشاء النسخة الاحتياطية على Google Drive بنجاح",
        sheetUrl: ss.getUrl(),
        backup: backupInfo
      });
    }

    if (action === "clearTransactions") {
      const backupInfo = createDriveBackup(ss, "قبل_تفريغ_الحركات");
      clearSpecificTables(ss, TRANSACTION_TABLES);
      if (body.tables) {
        writeSnapshotTables(ss, body.tables);
      }
      logAction(ss, body.user || "mrk", body.role || "مالك النظام", "تفريغ الحركات", "تم تفريغ الحركات والمدفوعات مع الإبقاء على الطلاب وأكواد الدروس");
      return jsonResponse({
        status: "ok",
        message: "تم تفريغ الحركات والمدفوعات مع حفظ نسخة احتياطية وبقاء الطلاب وأكواد الدروس",
        sheetUrl: ss.getUrl(),
        backup: backupInfo,
        data: readAllTables(ss)
      });
    }

    if (action === "clearAll") {
      const backupInfo = createDriveBackup(ss, "قبل_التفريغ_الشامل");
      clearSpecificTables(ss, FULL_RESET_TABLES);
      ensureAdminUser(ss);
      if (body.tables) {
        writeSnapshotTables(ss, body.tables);
      }
      logAction(ss, body.user || "mrk", body.role || "مالك النظام", "تفريغ شامل", "تم التفريغ الشامل للبيانات مع حفظ نسخة احتياطية");
      return jsonResponse({
        status: "ok",
        message: "تم التفريغ الشامل للبيانات مع حفظ نسخة احتياطية",
        sheetUrl: ss.getUrl(),
        backup: backupInfo,
        data: readAllTables(ss)
      });
    }

    return jsonResponse({ status: "error", message: "أمر غير معروف: " + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * إنشاء وتهيئة كل الـ 23 ورقة عمل تلقائياً وحذف الورقة الفارغة الافتراضية
 */
function initializeDatabaseIfNeeded() {
  const ss = getTargetSpreadsheet();
  const sheetNames = Object.keys(SCHEMA);

  sheetNames.forEach(function(name) {
    let sh = ss.getSheetByName(name);
    const headers = SCHEMA[name];
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.setRightToLeft(true);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#1e3a8a")
        .setFontColor("#ffffff");
      sh.setFrozenRows(1);
    } else if (sh.getLastRow() === 0) {
      sh.setRightToLeft(true);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  });

  // حذف الورقة الافتراضية الفارغة (Sheet1 أو الورقة1) إن وجدت بعد إنشاء الـ 23 ورقة
  ["Sheet1", "الورقة1", "ورقة1"].forEach(function(defaultName) {
    const defSheet = ss.getSheetByName(defaultName);
    if (defSheet && ss.getSheets().length > 1) {
      try { ss.deleteSheet(defSheet); } catch (e) {}
    }
  });

  ensureAdminUser(ss);
  return ss;
}

/**
 * ضمان وجود حساب مالك النظام mrk / 1510 دائماً
 */
function ensureAdminUser(ss) {
  const sh = ss.getSheetByName("المستخدمون");
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) {
    sh.appendRow(["U001", "mrk", "1510", "محمد السعيد (مالك النظام)", "مالك النظام", "نعم"]);
    return;
  }
  const data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  let found = false;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === "mrk") {
      found = true;
      break;
    }
  }
  if (!found) {
    sh.appendRow(["U001", "mrk", "1510", "محمد السعيد (مالك النظام)", "مالك النظام", "نعم"]);
  }
}

/**
 * قراءة كل الجداول
 */
function readAllTables(ss) {
  const result = {};
  Object.keys(SCHEMA).forEach(function(tableName) {
    const sh = ss.getSheetByName(tableName);
    if (!sh || sh.getLastRow() <= 1) {
      result[tableName] = [];
      return;
    }
    const headers = SCHEMA[tableName];
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getDisplayValues();
    result[tableName] = rows.map(function(r) {
      const obj = {};
      headers.forEach(function(h, idx) {
        obj[h] = r[idx];
      });
      return obj;
    });
  });
  return result;
}

/**
 * كتابة الجداول دفعة واحدة (Batch Write)
 */
function writeSnapshotTables(ss, tablesObj) {
  Object.keys(tablesObj).forEach(function(tableName) {
    if (!SCHEMA[tableName]) return;
    const sh = ss.getSheetByName(tableName);
    if (!sh) return;
    const headers = SCHEMA[tableName];
    const rowsData = tablesObj[tableName] || [];

    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).clearContent();
    }

    if (rowsData.length > 0) {
      const matrix = rowsData.map(function(item) {
        return headers.map(function(h) {
          return item[h] !== undefined && item[h] !== null ? item[h] : "";
        });
      });
      sh.getRange(2, 1, matrix.length, headers.length).setValues(matrix);
    }
  });
  ensureAdminUser(ss);
}

/**
 * مسح محتوى قائمة محددة من الجداول
 */
function clearSpecificTables(ss, tableList) {
  tableList.forEach(function(tableName) {
    const sh = ss.getSheetByName(tableName);
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    }
  });
}

/**
 * إنشاء نسخة احتياطية كاملة على Google Drive
 */
function createDriveBackup(ss, reason) {
  const folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
  const timestamp = Utilities.formatDate(new Date(), "Africa/Cairo", "yyyy-MM-dd_HH-mm-ss");
  const backupName = "نسخة_سنتر_العباقرة_" + reason + "_" + timestamp;
  const file = DriveApp.getFileById(ss.getId()).makeCopy(backupName, folder);
  return {
    name: backupName,
    id: file.getId(),
    url: file.getUrl(),
    timestamp: timestamp
  };
}

/**
 * تسجيل حركة في سجل العمليات
 */
function logAction(ss, user, role, type, details) {
  const sh = ss.getSheetByName("سجل_العمليات");
  if (!sh) return;
  const nowStr = Utilities.formatDate(new Date(), "Africa/Cairo", "yyyy-MM-dd HH:mm:ss");
  const id = "LOG-" + new Date().getTime();
  sh.appendRow([id, nowStr, user, role, type, details]);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
