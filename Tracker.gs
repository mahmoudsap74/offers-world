// ============================================================
// Tracker.gs — نظام متابعة الباقات اليدوية (صفحة عبير)
// ✅ UPDATED: getTrackerSummary تدعم all=1 + month + year
// ============================================================

const TRACKER_SHEET = '📊 PackageLogs';
const TRACKER_PASSWORD = PropertiesService.getScriptProperties().getProperty('TRACKER_PASSWORD') || 'abeer123';

// ============================================================
// doGet — إضافة page=tracker
// ============================================================
// ملاحظة: أضف السطر ده في Client.gs داخل doGet:
//   if (page === 'tracker') {
//     return HtmlService.createHtmlOutputFromFile('tracker')
//       .setTitle('متابعة الباقات — عالم العروض')
//       .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
//   }

// ============================================================
// Tracker Actions — تُضاف لـ doGet في Client.gs
// ============================================================
// أضف ده في doGet بعد action === 'packages':
//
//   if (action === 'trackerPackages') {
//     return respond(getTrackerPackages());
//   }
//   if (action === 'trackerLogs') {
//     return respond(getTrackerLogs(e.parameter));
//   }
//   if (action === 'trackerSummary') {
//     return respond(getTrackerSummary(e.parameter.all, e.parameter.month, e.parameter.year));
//   }
//
// وأضف ده في doPost داخل الـ switch/if:
//   if (action === 'addTrackerLog')    return respond(addTrackerLog(data));
//   if (action === 'editTrackerLog' || action === 'updateTrackerLog') return respond(editTrackerLog(data));
//   if (action === 'deleteTrackerLog') return respond(deleteTrackerLog(data));
//   if (action === 'checkTrackerPass') return respond({ success: data.pass === TRACKER_PASSWORD });

// ============================================================
// جلب الباقات من Master Data (نفس الـ Packages sheet الموجودة)
// ============================================================
function getTrackerPackages() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PACKAGES);
    var data = sheet.getDataRange().getValues();
    var grouped = {};

    for (var i = 2; i < data.length; i++) {
      var row = data[i];
      var pkgId      = String(row[0] || '').trim();
      var company    = String(row[1] || '').trim();
      var pkgName    = String(row[2] || '').trim();
      var pkgPrice   = Number(row[3]) || 0;
      var status     = String(row[5] || '').trim();

      if (!pkgName || !company || status !== 'متاح') continue;

      if (!grouped[company]) grouped[company] = [];
      grouped[company].push({ id: pkgId, name: pkgName, price: pkgPrice });
    }

    Object.keys(grouped).forEach(function(co) {
      grouped[co].sort(function(a, b) { return a.price - b.price; });
    });

    return { success: true, data: grouped };
  } catch (err) {
    logError('getTrackerPackages', err.toString());
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// إضافة عملية جديدة
// ============================================================
function addTrackerLog(data) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateTrackerSheet(ss);

    var logId = 'LOG-' + Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyMMddHHmmss') + '-' + Math.floor(Math.random() * 100);
    var opDate = data.date ? data.date : Utilities.formatDate(new Date(), 'Africa/Cairo', 'dd/MM/yyyy');

    sheet.appendRow([
      logId,
      opDate,
      data.company     || '',
      data.packageId   || '',
      data.packageName || '',
      Number(data.price) || 0,
      data.source      || '',
      data.clientRef   || '',
      data.status      || 'تم',
      data.notes       || '',
      Utilities.formatDate(new Date(), 'Africa/Cairo', 'dd/MM/yyyy HH:mm:ss')
    ]);

    addLog('تسجيل باقة يدوي', logId, data.company + ' - ' + data.packageName);
    return { success: true, logId: logId };
  } catch (err) {
    logError('addTrackerLog', err.toString());
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// تعديل عملية موجودة — بحث بالصف المباشر (rowId)
// ============================================================
function editTrackerLog(data) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateTrackerSheet(ss);

    var row;
    var rowId = parseInt(data.rowId, 10);

    if (!isNaN(rowId) && rowId > 1) {
      row = rowId;
    } else {
      var found = sheet.getRange('A:A').createTextFinder(String(data.logId || '').trim()).matchEntireCell(true).findNext();
      if (!found) return { success: false, error: 'العملية غير موجودة' };
      row = found.getRow();
    }

    sheet.getRange(row, 3, 1, 8).setValues([[
      data.company     || '',
      data.packageId   || '',
      data.packageName || '',
      Number(data.price) || 0,
      data.source      || '',
      data.clientRef   || '',
      data.status      || 'تم',
      data.notes       || ''
    ]]);

    addLog('تعديل باقة يدوي', data.logId || ('row#' + row), data.packageName);
    return { success: true };
  } catch (err) {
    logError('editTrackerLog', err.toString());
    return { success: false, error: err.toString() };
  }
}

function updateTrackerLog(data) {
  return editTrackerLog(data);
}

// ============================================================
// حذف عملية — بحث بالصف المباشر (rowId)
// ============================================================
function deleteTrackerLog(data) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateTrackerSheet(ss);

    var rowToDelete;
    var rowId = parseInt(typeof data === 'object' ? data.rowId : null, 10);
    var logId = typeof data === 'object' ? (data.logId || '') : String(data || '');

    if (!isNaN(rowId) && rowId > 1) {
      rowToDelete = rowId;
    } else {
      var found = sheet.getRange('A:A').createTextFinder(logId.trim()).matchEntireCell(true).findNext();
      if (!found) return { success: false, error: 'العملية غير موجودة' };
      rowToDelete = found.getRow();
    }

    sheet.deleteRow(rowToDelete);
    addLog('حذف باقة يدوي', logId || ('row#' + rowToDelete), '');
    return { success: true };
  } catch (err) {
    logError('deleteTrackerLog', err.toString());
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// جلب السجلات مع فلترة
// ============================================================
function getTrackerLogs(params) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateTrackerSheet(ss);
    var data  = sheet.getDataRange().getValues();
    var logs  = [];

    var filterCompany = params.company || '';
    var filterStatus  = params.status  || '';
    var filterClient  = params.client  || '';

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;

      var log = {
        rowId:       i + 1,
        logId:       String(row[0]  || ''),
        date:        normalizeDate(row[1]),   // ✅ FIX: دعم Date Object
        company:     String(row[2]  || ''),
        packageId:   String(row[3]  || ''),
        packageName: String(row[4]  || ''),
        price:       Number(row[5]) || 0,
        source:      String(row[6]  || ''),
        clientRef:   String(row[7]  || ''),
        status:      String(row[8]  || ''),
        notes:       String(row[9]  || ''),
        createdAt:   normalizeDate(row[10])   // ✅ FIX: دعم Date Object
      };

      if (filterCompany && log.company !== filterCompany) continue;
      if (filterStatus  && log.status  !== filterStatus)  continue;
      if (filterClient  && !log.clientRef.toLowerCase().includes(filterClient.toLowerCase())) continue;

      logs.push(log);
    }

    logs.reverse();
    return { success: true, data: logs };
  } catch (err) {
    logError('getTrackerLogs', err.toString());
    return { success: false, error: err.toString(), data: [] };
  }
}

// ============================================================
// ✅ دالة مساعدة: تحويل التاريخ بأمان (Date Object أو نص)
// Google Sheets أحياناً يرجع التاريخ كـ Date Object مش String
// ============================================================
function normalizeDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Africa/Cairo', 'dd/MM/yyyy');
  }
  return String(val).trim().split(' ')[0]; // خذ التاريخ فقط بدون الوقت
}

// ============================================================
// ✅ FIXED: ملخص الفترة — يدعم all=1 + month + year
// الإصلاحات:
//   1. البارامتر الأول allFlag بدلاً من period
//   2. إصلاح مقارنة الشهر (parseInt)
//   3. ✅ إصلاح قراءة التاريخ (normalizeDate)
//   4. إضافة مفتاح filtered في الـ response
// ============================================================
function getTrackerSummary(allFlag, month, year) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateTrackerSheet(ss);
    var data  = sheet.getDataRange().getValues();

    var now      = new Date();
    var todayStr = Utilities.formatDate(now, 'Africa/Cairo', 'dd/MM/yyyy');

    // ✅ FIX: استخدام parseInt لتجنب مشكلة '7' vs '07'
    var showAll     = (allFlag === '1' || allFlag === true);
    var targetMonth = month ? parseInt(month, 10) : parseInt(Utilities.formatDate(now, 'Africa/Cairo', 'MM'), 10);
    var targetYear  = year  ? parseInt(year,  10) : parseInt(Utilities.formatDate(now, 'Africa/Cairo', 'yyyy'), 10);

    // مُجمِّعات اليوم
    var daily = { count: 0, revenue: 0, byCompany: {} };

    // مُجمِّعات الفترة المختارة (شهر أو كل البيانات)
    var filtered = {
      count: 0, revenue: 0,
      byCompany: {}, byPackage: {}, bySource: {}, dailyTrend: {}
    };

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;

      var dateStr = normalizeDate(row[1]);   // ✅ FIX: تحويل آمن (Date Object → dd/MM/yyyy)
      var company = String(row[2] || '').trim();
      var pkgName = String(row[4] || '').trim();
      var price   = Number(row[5]) || 0;
      var source  = String(row[6] || '').trim();
      var status  = String(row[8] || '').trim();

      if (status === 'ملغي') continue;

      // تقسيم التاريخ
      var parts = dateStr.split('/');  // [dd, MM, yyyy]
      if (parts.length < 3) continue;

      // ✅ FIX: parseInt لمنع مشكلة مقارنة '7' !== '07'
      var rowMonth = parseInt(parts[1], 10);
      var rowYear  = parseInt(parts[2], 10);

      // --- إحصائيات اليوم (دائمًا) ---
      if (dateStr === todayStr) {
        daily.count++;
        daily.revenue += price;
        daily.byCompany[company] = (daily.byCompany[company] || 0) + 1;
      }

      // --- إحصائيات الفترة ---
      var inPeriod = showAll
        ? true
        : (rowMonth === targetMonth && rowYear === targetYear);

      if (inPeriod) {
        filtered.count++;
        filtered.revenue += price;
        filtered.byCompany[company]                  = (filtered.byCompany[company]                  || 0) + 1;
        filtered.byPackage[pkgName]                  = (filtered.byPackage[pkgName]                  || 0) + 1;
        filtered.bySource[source || 'غير محدد']      = (filtered.bySource[source || 'غير محدد']      || 0) + 1;
        var dayKey = parts[0] + '/' + parts[1];
        filtered.dailyTrend[dayKey]                  = (filtered.dailyTrend[dayKey]                  || 0) + 1;
      }
    }

    // --- أعلى وأضعف شركة ---
    var compEntries = Object.keys(filtered.byCompany).map(function(k) {
      return { name: k, count: filtered.byCompany[k] };
    }).sort(function(a, b) { return b.count - a.count; });

    var topCompany    = compEntries.length > 0 ? compEntries[0]                        : { name: '', count: 0 };
    var bottomCompany = compEntries.length > 1 ? compEntries[compEntries.length - 1]   : { name: '', count: 0 };

    // --- أعلى وأضعف باقة ---
    var pkgEntries = Object.keys(filtered.byPackage).map(function(k) {
      return { name: k, count: filtered.byPackage[k] };
    }).sort(function(a, b) { return b.count - a.count; });

    var topPackage    = pkgEntries.length > 0 ? pkgEntries[0]                      : { name: '', count: 0 };
    var bottomPackage = pkgEntries.length > 1 ? pkgEntries[pkgEntries.length - 1]  : { name: '', count: 0 };

    // بناء كائن الفترة الكامل
    var periodData = {
      count:         filtered.count,
      revenue:       filtered.revenue,
      byCompany:     filtered.byCompany,
      byPackage:     filtered.byPackage,
      bySource:      filtered.bySource,
      dailyTrend:    filtered.dailyTrend,
      topCompany:    topCompany,
      bottomCompany: bottomCompany,
      topPackage:    topPackage,
      bottomPackage: bottomPackage
    };

    return {
      success: true,
      today: {
        date:      todayStr,
        count:     daily.count,
        revenue:   daily.revenue,
        byCompany: daily.byCompany
      },
      filtered: periodData,  // ✅ الجديد — يقرأه tracker.html
      monthly:  periodData   // للتوافق مع الإصدار القديم
    };

  } catch (err) {
    logError('getTrackerSummary', err.toString());
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// دالة مساعدة: جلب/إنشاء الـ Sheet
// ============================================================
function getOrCreateTrackerSheet(ss) {
  var sheet = ss.getSheetByName(TRACKER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TRACKER_SHEET);
    sheet.appendRow([
      'رقم العملية', 'تاريخ العملية', 'الشركة', 'رقم الباقة', 'اسم الباقة',
      'السعر', 'مصدر العملية', 'اسم/رقم العميل', 'الحالة', 'ملاحظات', 'وقت الإدخال'
    ]);
    sheet.getRange(1, 1, 1, 11).setBackground('#8b3cf7').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 11, 150);
  }
  return sheet;
}
