// فتح لوحة التحكم الرئيسية من القائمة العلويّة
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏥 نظام عيادة رونق سمايل')
    .addItem('اللوحة الرئيسية (التنقل)', 'showMainDashboard')
    .addItem('تسجيل مريض جديد', 'showPatientForm')
    .addItem('تسجيل دفعة جديدة', 'showPaymentForm')
    .addItem('كشف حساب مريض', 'showPatientReport')
    .addToUi();
}

// 1. دوال فتح النوافذ داخل Google Sheets
function showMainDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('MainDashboard').setWidth(600).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function showPatientForm() {
  const html = HtmlService.createHtmlOutputFromFile('PatientForm').setWidth(620).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function showPaymentForm() {
  const html = HtmlService.createHtmlOutputFromFile('PaymentForm').setWidth(600).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function showPatientReport() {
  const html = HtmlService.createHtmlOutputFromFile('PatientReport').setWidth(600).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

// 2. دوال جلب محتوى الصفحات الفرعية
function getPatientFormHtml() {
  return HtmlService.createHtmlOutputFromFile('PatientForm').getContent();
}

function getPaymentFormHtml() {
  return HtmlService.createHtmlOutputFromFile('PaymentForm').getContent();
}

function getMainDashboardHtml() {
  return HtmlService.createHtmlOutputFromFile('MainDashboard').getContent();
}

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'MainDashboard';
  
  var title = 'عيادة رونق سمايل';
  if (page === 'PatientForm') title = 'تسجيل مريض - عيادة رونق سمايل';
  if (page === 'PaymentForm') title = 'تسجيل دفعة - عيادة رونق سمايل';
  if (page === 'PatientReport') title = 'التقرير المالي للمريض - عيادة رونق سمايل';

  var htmlOutput = HtmlService.createTemplateFromFile(page).evaluate();
  
  return htmlOutput
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// 3. جلب قائمة الخدمات والأسعار من شيت Services
function getServicesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Services');
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data
    .filter(row => row[0] !== "")
    .map(row => ({ name: String(row[0]), price: Number(row[1]) || 0 }));
}

// 4. الفحص اللحظي لتكرار كود المريض (يعتمد كلياً على الكود القادم من الواجهة دون إضافة RS-)
function checkPatientIdExists(fullPatientId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Patients') || ss.getSheetByName('المرضى');
  if (!sheet || sheet.getLastRow() < 2) return false;
  
  const searchID = fullPatientId.trim().toUpperCase();
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(v => String(v).trim().toUpperCase());
  return ids.includes(searchID);
}

// 5. حفظ بيانات المريض كاملة في الشيت (تخزين الكود المختار بالكامل بدون أي تعديل أو دمج ثابت)
function savePatientData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Patients') || ss.getSheetByName('المرضى');
  
  if (!sheet) {
    sheet = ss.insertSheet('Patients');
    sheet.appendRow([
      "كود المريض", "اسم المريض", "الجنس", "رقم الموبايل", "العمر", 
      "أنواع العلاج", "الأمراض المزمنة", "الأدوية", "الحساسية", "مدخن", 
      "التكلفة الإجمالية", "نسبة الخصم", "المبلغ بعد الخصم", "الطبيب المتابع", "تاريخ التسجيل"
    ]);
  }

  // الكود الكامل المختار من واجهة المستخدم (مثال: R-1001 أو RS-1001 أو RZ-1001)
  const fullID = data.patientId.trim().toUpperCase();
  
  if (sheet.getLastRow() >= 2) {
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(v => String(v).trim().toUpperCase());
    if (ids.includes(fullID)) {
      throw new Error("كود المريض " + fullID + " مسجل مسبقاً في النظام!");
    }
  }

  const currentDate = new Date();

  sheet.appendRow([
    fullID,
    data.fullName,
    data.gender,
    "'" + data.phone,
    data.age,
    data.treatments.join(", "),
    data.diseases.join(", "),
    data.medicines,
    data.allergies,
    data.isSmoker,
    data.totalCost,
    data.discountPercent + "%",
    data.finalCost,
    data.doctor,
    currentDate
  ]);
  
  return "تم تسجيل المريض " + data.fullName + " بنجاح تحت كود (" + fullID + ")";
}

/**
 * دالة البحث المباشر للـ Dropdown
 */
function getPatientsListForDropdown(searchVal) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Patients") || ss.getSheetByName("المرضى");
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    // جلب أول 4 أعمدة: [0]=الكود، [1]=الاسم، [2]=الجنس، [3]=رقم الهاتف
    var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    var searchLower = String(searchVal).toLowerCase().trim();
    var results = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var pId = String(row[0] || "").trim();
      var pName = String(row[1] || "").trim();
      var pPhone = String(row[3] || "").trim(); // رقم الهاتف الخام كما هو بالظبط

      if (!pName) continue;

      var searchKey = pName + " (" + pId + ")";

      // المطابقة برقم الكود أو الاسم
      if (pName.toLowerCase().indexOf(searchLower) !== -1 || pId.toLowerCase().indexOf(searchLower) !== -1) {
        results.push({
          id: pId,
          name: pName,
          phone: pPhone, // إرجاع رقم الهاتف الخام
          searchKey: searchKey
        });
      }

      if (results.length >= 20) break; // اكتفاء بأول 10 نتائج للأداء
    }

    return results;
  } catch (e) {
    Logger.log("Error in getPatientsListForDropdown: " + e.message);
    return [];
  }
}

// 2. البحث عن التفاصيل المالية بالاسم أو الكود أو النص المركب
function getPatientFinancialDetails(searchQuery) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. الوصول لشيت المرضى والمدفوعات
    var patientsSheet = ss.getSheetByName("المرضى") || ss.getSheetByName("Patients");
    var paymentsSheet = ss.getSheetByName("المدفوعات") || ss.getSheetByName("Payments");

    if (!patientsSheet) {
      return { found: false, error: "لم يتم العثور على شيت المرضى" };
    }

    var patientsData = patientsSheet.getDataRange().getValues();
    var queryStr = String(searchQuery).trim().toLowerCase();

    var patientId = "";
    var patientName = "";
    var totalAmountRequired = 0; // المبلغ الصافي بعد الخصم
    var found = false;

    for (var i = 1; i < patientsData.length; i++) {
      var row = patientsData[i];
      var pId = String(row[0] || "").trim();    // العمود A: كود المريض
      var pName = String(row[1] || "").trim();  // العمود B: اسم المريض
      var comboKey = (pId + " - " + pName).toLowerCase();

      if (queryStr === comboKey || queryStr === pId.toLowerCase() || queryStr === pName.toLowerCase()) {
        found = true;
        patientId = pId;
        patientName = pName;

        // 🎯 جلب المبلغ المطلوب مباشرة من العمود N (row[13])
        totalAmountRequired = Number(row[12]) || 0;

        break;
      }
    }

    if (!found) {
      return { found: false };
    }

    // 2. حساب إجمالي ما تم دفعه سابقاً من شيت المدفوعات
    var totalPaidSoFar = 0;
    if (paymentsSheet) {
      var paymentsData = paymentsSheet.getDataRange().getValues();
      for (var j = 1; j < paymentsData.length; j++) {
        var payRow = paymentsData[j];
        var payPId = String(payRow[1] || "").trim(); // كود المريض في شيت المدفوعات (العمود B)
        
        if (payPId.toLowerCase() === patientId.toLowerCase()) {
          totalPaidSoFar += Number(payRow[3]) || 0;  // المبلغ المدفوع (العمود D)
        }
      }
    }

    // 3. حساب المتبقي
    var remainingBalance = totalAmountRequired - totalPaidSoFar;

    return {
      found: true,
      patientId: patientId,
      patientName: patientName,
      totalRequired: totalAmountRequired, // سيظهر القيمة المأخوذة مباشرة من row[13]
      totalPaidSoFar: totalPaidSoFar,
      remainingBalance: remainingBalance < 0 ? 0 : remainingBalance
    };

  } catch (err) {
    console.error("خطأ في جلب التفاصيل المالية:", err);
    return { found: false, error: err.message };
  }
}

// 7. تسجيل الدفعة الجديدة في شيت Payments
function saveAppointment(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Appointments');

    if (!sheet) {
      sheet = ss.insertSheet('Appointments');
      sheet.appendRow([
        "ID الموعد", 
        "كود المريض", 
        "اسم المريض", 
        "رقم الهاتف", 
        "تاريخ الموعد", 
        "وقت الموعد", 
        "نوع الخدمة", 
        "حالة الموعد", 
        "ملاحظات"
      ]);
    }

    var lastRow = sheet.getLastRow();
    var appointmentId = "APT-" + (1000 + lastRow);

    var pCode   = data.patientCode || "";
    var pName   = data.patientName || "";
    var service = data.serviceType || ""; 
    var notes   = data.notes || "";
    var pPhone  = data.patientPhone || "";
    var appDate = data.appointmentDate || "";
    var appTime = data.appointmentTime || "";

    sheet.appendRow([
      appointmentId,
      pCode,
      pName,
      pPhone,
      appDate,
      appTime,
      service,
      "قيد الانتظار",
      notes
    ]);

    // ✅ تخصيص رسالة النجاح لتشمل اسم المريض بوضوح
    return {
      success: true,
      message: "تم حجز الموعد بنجاح للمريض: " + pName
    };
  } catch (err) {
    return {
      success: false,
      message: "خطأ في السيرفر: " + err.toString()
    };
  }
}

function getPatientPaymentHistory(searchQuery) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var patientsSheet = ss.getSheetByName("المرضى") || ss.getSheetByName("Patients");
    var paymentsSheet = ss.getSheetByName("المدفوعات") || ss.getSheetByName("Payments");

    if (!patientsSheet) {
      return { success: false, message: "لم يتم العثور على شيت المرضى" };
    }

    var patientsData = patientsSheet.getDataRange().getValues();
    var queryStr = String(searchQuery).trim().toLowerCase();

    var patientInfo = null;

    function parseVal(val) {
      if (val === undefined || val === null) return 0;
      if (typeof val === 'number') return val;
      var cleaned = String(val).replace(/[^0-9.-]/g, '');
      return parseFloat(cleaned) || 0;
    }

    // 1. البحث عن المريض من شيت المرضى
    for (var i = 1; i < patientsData.length; i++) {
      var row = patientsData[i];
      var pId = String(row[0] || "").trim();    // العمود A (Index 0)
      var pName = String(row[1] || "").trim();  // العمود B (Index 1)
      var comboKey = (pId + " - " + pName).toLowerCase();

      if (queryStr === comboKey || queryStr === pId.toLowerCase() || queryStr === pName.toLowerCase()) {
        var netTotalCost = parseVal(row[12]); // الصافي بعد الخصم (العمود M - Index 12)

        patientInfo = {
          id: pId,
          name: pName,
          phone: row[3] || "", // 🎯 تعديل: جلب رقم الهاتف من العمود D (Index 3)
          totalCost: netTotalCost
        };
        break;
      }
    }

    if (!patientInfo) {
      return { success: false, message: "المريض غير مسجل بالسجلات" };
    }

    // 2. جلب الدفعات والتاريخ والملاحظات من شيت المدفوعات
    var patientPayments = [];
    var totalPaid = 0;
    var allNotesList = [];

    if (paymentsSheet) {
      var paymentsData = paymentsSheet.getDataRange().getValues();
      for (var j = 1; j < paymentsData.length; j++) {
        var payRow = paymentsData[j];
        var payPId = String(payRow[1] || "").trim(); // كود المريض في العمود B (Index 1)

        if (payPId.toLowerCase() === patientInfo.id.toLowerCase()) {
          var paidAmt = parseVal(payRow[3]); // المبلغ المدفوع بالعمود D (Index 3)
          totalPaid += paidAmt;

          // جلب رقم الإيصال من العمود A (Index 0)
          var receiptNumber = payRow[0] || j;

          // جلب تاريخ الدفعة من العمود F (Index 5)
          var rawDate = payRow[5];
          var formattedDate = "-";
          if (rawDate) {
            formattedDate = rawDate instanceof Date 
              ? Utilities.formatDate(rawDate, ss.getSpreadsheetTimeZone(), "yyyy/MM/dd")
              : String(rawDate);
          }

          // جلب الملاحظات من العمود H (Index 7)
          var noteVal = String(payRow[7] || "").trim();
          if (noteVal && noteVal !== "-") {
            allNotesList.push(noteVal);
          }

          patientPayments.push({
            receiptNo: receiptNumber,
            amount: paidAmt,
            date: formattedDate
          });
        }
      }
    }

    var remaining = patientInfo.totalCost - totalPaid;

    return {
      success: true,
      patient: patientInfo,
      payments: patientPayments,
      totalPaid: totalPaid,
      remaining: remaining < 0 ? 0 : remaining,
      notesSummary: allNotesList.length > 0 ? allNotesList.join(" | ") : "لا توجد ملاحظات"
    };

  } catch (err) {
    console.error("خطأ في جلب سجل المريض:", err);
    return { success: false, message: err.message };
  }

}

function getAppointmentsByDate(selectedDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Appointments');
    
    if (!sheet) {
      return { success: true, appointments: [] };
    }

    var data = sheet.getDataRange().getValues();
    var appointments = [];

    var cleanSelectedDate = String(selectedDate || "").trim().split('T')[0];

    // البدء من السطر الثاني (تجاوز الهيدر)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var rawDate = row[4]; // العمود E: تاريخ الموعد
      var formattedAppDate = "";

      if (rawDate instanceof Date) {
        var year = rawDate.getFullYear();
        var month = String(rawDate.getMonth() + 1).padStart(2, '0');
        var day = String(rawDate.getDate()).padStart(2, '0');
        formattedAppDate = year + '-' + month + '-' + day;
      } else {
        formattedAppDate = String(rawDate || "").trim().split('T')[0];
      }

      // معالجة الوقت وتحويله إلى نظام 12 ساعة (ص / م)
      var rawTime = row[5]; 
      var formattedTime = "";
      
      if (rawTime instanceof Date) {
        var h = rawTime.getHours();
        var m = rawTime.getMinutes();
        var period = h >= 12 ? 'م' : 'ص';
        h = h % 12;
        h = h ? h : 12;
        var minutesStr = String(m).padStart(2, '0');
        formattedTime = h + ':' + minutesStr + ' ' + period;
      } else {
        var timeStr = String(rawTime || "").trim();
        if (timeStr.includes(':')) {
          var parts = timeStr.split(':');
          var h = parseInt(parts[0], 10);
          var m = parseInt(parts[1], 10) || 0;
          if (!isNaN(h)) {
            var period = h >= 12 ? 'م' : 'ص';
            h = h % 12;
            h = h ? h : 12;
            var minutesStr = String(m).padStart(2, '0');
            formattedTime = h + ':' + minutesStr + ' ' + period;
          } else {
            formattedTime = timeStr;
          }
        } else {
          formattedTime = timeStr;
        }
      }

      if (!cleanSelectedDate || formattedAppDate === cleanSelectedDate) {
        appointments.push({
          rowId: i + 1,
          id: row[0],
          patientCode: row[1],
          patientName: row[2],
          phone: row[3],
          date: formattedAppDate,
          time: formattedTime || "00:00",
          serviceType: row[6],
          status: row[7] || "قيد الانتظار",
          notes: row[8]
        });
      }
    }

    // ملاحظة: بما أن الترتيب الحرفي قد يتأثر بوجود (ص/م)، يمكنك الاعتماد على وقت الشيت أو بقاء الفرز كما هو
    return { success: true, appointments: appointments };
  } catch (err) {
    return { success: false, appointments: [], message: err.toString() };
  }
}

// دالة لتحديث حالة الموعد مباشرة في العمود الثامن (حالة الموعد)
function updateAppointmentStatus(rowId, newStatus) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Appointments');
    
    if (!sheet) {
      return { success: false, message: "شيت المواعيد غير موجود" };
    }

    // العمود الثامن في الشيت هو "حالة الموعد" (Column H)
    sheet.getRange(rowId, 8).setValue(newStatus);

    return { success: true, message: "تم تحديث الحالة بنجاح" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

