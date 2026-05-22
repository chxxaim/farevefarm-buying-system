/**
 * Google Apps Script backend file (Code.gs)
 * สำหรับเชื่อมโยงระบบเว็บบอร์ดหน้าบ้าน HTML กับ Google Sheets ภายในองค์กร Farevefarm
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Farevefarm Internal Buying System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 1. ดึงข้อมูลนัดหมายประจําวันไปยังหน้า Web App HTML
function getAppointments() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Appointments");
  var range = sheet.getRange(4, 1, sheet.getLastRow() - 3, 9);
  var values = range.getValues();
  
  var data = [];
  for (var i = 0; i < values.length; i++) {
    data.push({
      id: values[i][0],
      date: Utilities.formatDate(new Date(values[i][1]), "GMT+7", "yyyy-MM-dd"),
      time: values[i][2],
      name: values[i][3],
      phone: values[i][4],
      itemCount: values[i][5],
      totalAmount: values[i][6],
      statusItem: values[i][7],
      statusPay: values[i][8]
    });
  }
  return data;
}

// 2. บันทึกเพิ่มรายชื่อคิวนัดหมายใหม่จากระบบหน้าบ้าน
function saveNewBooking(appData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Appointments");
  
  // เพิ่มแถวข้อมูลใหม่ลงใน Sheet Appointments ต่อท้ายแถวเดิม
  sheet.appendRow([
    appData.id,
    appData.date,
    appData.time,
    appData.name,
    appData.phone,
    0, // ตั้งต้นจำนวนชิ้นเป็น 0
    0, // ตั้งต้นยอดเงินรวมโอนเป็น 0
    "รอดำเนินการ",
    "รอดำเนินการ",
    "สร้างคิวนัดหมายผ่านหน้าเว็บ"
  ]);
  return true;
}

// 3. อัปเดตและบันทึกชุดข้อมูลรายข้อ Checklist ตรวจสอบรายชิ้น
function uploadChecklistToSheets(appointmentId, itemsArray, totalBuyingPrice) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var appSheet = ss.getSheetByName("Appointments");
  var itemSheet = ss.getSheetByName("Product_Items");
  
  // อัปเดตข้อมูลสรุปจำนวนชิ้นและยอดเงินในหน้านัดหมาย (Sheet 1)
  var appRows = appSheet.getLastRow();
  var appRange = appSheet.getRange(4, 1, appRows - 3, 1);
  var appIds = appRange.getValues();
  
  for (var i = 0; i < appIds.length; i++) {
    if (appIds[i][0] == appointmentId) {
      var rowNum = i + 4;
      appSheet.getRange(rowNum, 6).setValue(itemsArray.length); // จำนวนชิ้นที่ตรวจจริงหน้างาน
      appSheet.getRange(rowNum, 7).setValue(totalBuyingPrice);   // ยอดรวมโอนจริง
      appSheet.getRange(rowNum, 8).setValue("รับสินค้าแล้ว");
      appSheet.getRange(rowNum, 9).setValue("โอนเงินเรียบร้อย");
      break;
    }
  }
  
  // เพิ่มรายการสินค้าทุกชิ้นลงในชีตประเมินสภาพแยกชิ้น (Sheet 2)
  for (var j = 0; j < itemsArray.length; j++) {
    var item = itemsArray[j];
    var itemId = "ITEM-" + appointmentId.split("-")[2] + "-0" + (j+1);
    
    itemSheet.appendRow([
      itemId,
      appointmentId,
      item.cat,
      item.title,
      "ประเมินหน้างาน",
      item.score,
      item.cond,
      item.price,
      "พนักงานหน้างาน",
      new Date(),
      JSON.stringify(item)
    ]);
  }
  return true;
}
