// ============================================================
// ณดาคลีนิค — Google Apps Script
// Sheet ID : 1PrgRE2SULvivIT2DxmCOlrtJjFvvTHOeKexuKdjnhDY
// Sheet Name: patient
// Drive Folder ID: 13KirLLdz08S-0gBT-Jo4XZ7991Vhph6x
// ============================================================

const SHEET_ID   = "1PrgRE2SULvivIT2DxmCOlrtJjFvvTHOeKexuKdjnhDY";
const SHEET_NAME = "patient";
const FOLDER_ID  = "13KirLLdz08S-0gBT-Jo4XZ7991Vhph6x";

function buildResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// รองรับ CORS preflight
function doOptions(e) {
  return buildResponse({ status: "ok" });
}

function doPost(e) {
  try {
    const p = e.parameter;

    const name     = p.name     || "";
    const email    = p.email    || "";
    const phone    = p.phone    || "";
    const dob      = p.dob      || "";
    const age      = p.age      || "";
    const symptoms = p.symptoms || "";

    const folder    = DriveApp.getFolderById(FOLDER_ID);
    const timestamp = new Date().getTime();
    let imageUrl = "";
    let pdfUrl   = "";

    if (p.imageBase64) {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(p.imageBase64),
        p.imageMime || "image/jpeg",
        timestamp + "_" + (p.imageName || "photo.jpg")
      );
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageUrl = "https://drive.google.com/uc?id=" + file.getId();
    }

    if (p.pdfBase64) {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(p.pdfBase64),
        "application/pdf",
        timestamp + "_" + (p.pdfName || "document.pdf")
      );
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfUrl = "https://drive.google.com/file/d/" + file.getId() + "/view";
    }

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      const headers = [
        "ลำดับ", "วันที่ลงทะเบียน", "ชื่อ", "Email",
        "เบอร์โทรศัพท์", "วันเดือนปีเกิด", "อายุ (ปี)",
        "อาการป่วย", "URL รูปภาพ", "URL เอกสาร PDF"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight("bold")
           .setBackground("#7C3AED")
           .setFontColor("#ffffff");
    }

    const seq     = sheet.getLastRow();
    const dateStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");
    sheet.appendRow([seq, dateStr, name, email, phone, dob, age, symptoms, imageUrl, pdfUrl]);

    return buildResponse({ status: "success", message: "บันทึกข้อมูลสำเร็จ" });

  } catch (err) {
    return buildResponse({ status: "error", message: err.toString() });
  }
}

function doGet(e) {
  return buildResponse({ status: "ok", message: "ณดาคลีนิค API พร้อมใช้งาน" });
}

function testWrite() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  Logger.log("Sheet found: " + sheet);
  sheet.appendRow(["TEST", new Date(), "ทดสอบ", "test@test.com", "0812345678", "01/01/2533", "35", "ปวดหัว", "", ""]);
}
