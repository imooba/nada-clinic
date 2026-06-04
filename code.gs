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
    const action = p.action || "";
    
    // ==========================================
    // ระบบ Login ตรวจสอบสิทธิ์การเข้าใช้งาน
    // ==========================================
    if (action === "login") {
      const loginEmail = (p.email || "").trim().toLowerCase();
      const loginPass  = (p.password || "").trim();

      const ss = SpreadsheetApp.openById(SHEET_ID);
      let staffSheet = ss.getSheetByName("Staff");

      // ถ้ายังไม่มีชีต Staff ให้สร้างใหม่ และใส่รหัสผ่านแอดมินเริ่มต้นให้
      if (!staffSheet) {
        staffSheet = ss.insertSheet("Staff");
        staffSheet.appendRow(["Email", "Password", "ชื่อพนักงาน", "สิทธิ์"]);
        staffSheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#3B82F6").setFontColor("#ffffff");
        // สร้าง user เริ่มต้น (คุณสามารถไปแก้ใน Sheet ภายหลังได้)
        staffSheet.appendRow(["admin@gmail.com", "123456", "ผู้ดูแลระบบ", "Admin"]);
      }

      const data = staffSheet.getDataRange().getValues();
      
      // วนลูปเช็คว่า Email และ Password ตรงกับในระบบหรือไม่
      for (let i = 1; i < data.length; i++) {
        const sheetEmail = String(data[i][0]).trim().toLowerCase();
        const sheetPass  = String(data[i][1]).trim();
        
        if (sheetEmail === loginEmail && sheetPass === loginPass) {
          return buildResponse({ 
            status: "success", 
            name: data[i][2], 
            role: data[i][3],
            message: "เข้าสู่ระบบสำเร็จ" 
          });
        }
      }
      return buildResponse({ status: "error", message: "ไม่มีสิทธิ์เข้าถึง หรือรหัสผ่านไม่ถูกต้อง" });
    }

    // ==========================================
    // 1. เพิ่มฟีเจอร์: บันทึกประวัติการรักษา (History)
    // ==========================================
    if (action === "add_history") {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      
      // ตรวจสอบว่ามี Sheet ชื่อ History หรือยัง ถ้ายังให้สร้างใหม่
      let historySheet = ss.getSheetByName("History");
      if (!historySheet) {
        historySheet = ss.insertSheet("History");
        const headers = [
          "วันที่มารักษา", "รหัส HN", "ชื่อผู้ป่วย", "แพทย์ผู้รักษา", 
          "อาการ", "คำวินิจฉัยแพทย์", "ยาที่จ่าย", "สิทธิ์ที่เบิกจ่าย", "ราคา ค่ารักษา"
        ];
        historySheet.appendRow(headers);
        historySheet.getRange(1, 1, 1, headers.length)
                    .setFontWeight("bold")
                    .setBackground("#10B981") // สีเขียวแยกจากหน้า patient
                    .setFontColor("#ffffff");
      }

      // วันที่และเวลาปัจจุบัน
      const dateStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");
      
      // บันทึกข้อมูลลงตาราง History
      historySheet.appendRow([
        dateStr,
        p.hn || "",
        p.name || "",
        p.doctor || "",
        p.symptoms || "",
        p.diagnosis || "",
        p.medicine || "",
        p.right || "",
        p.cost || ""
      ]);

      return buildResponse({ status: "success", message: "บันทึกประวัติการรักษาสำเร็จ" });
    }

    // ==========================================
    // 2. ฟีเจอร์เดิม: อัปเดตข้อมูลผู้ป่วย
    // ==========================================
    if (action === "update") {
      const memberId = p.memberId;
      const ss    = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(SHEET_NAME);
      const data  = sheet.getDataRange().getValues();
      
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]).trim() === String(memberId).trim()) {
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex !== -1) {
        sheet.getRange(rowIndex, 4).setValue(p.name);
        sheet.getRange(rowIndex, 5).setValue(p.email);
        sheet.getRange(rowIndex, 6).setValue(p.phone);
        sheet.getRange(rowIndex, 7).setValue(p.dob);
        sheet.getRange(rowIndex, 8).setValue(p.age);
        sheet.getRange(rowIndex, 9).setValue(p.symptoms);
        sheet.getRange(rowIndex, 12).setValue(p.diagnosis);
        sheet.getRange(rowIndex, 13).setValue(p.nextAppointment);

        // --- ส่วนที่ต้องเพิ่มใหม่ (บันทึกคอลัมน์ 14 - 19) ---
        sheet.getRange(rowIndex, 14).setValue(p.idCard || "");
        sheet.getRange(rowIndex, 15).setValue(p.address || "");
        sheet.getRange(rowIndex, 16).setValue(p.udHospital || "");
        sheet.getRange(rowIndex, 17).setValue(p.udMedicine || "");
        sheet.getRange(rowIndex, 18).setValue(p.allergy || "");
        sheet.getRange(rowIndex, 19).setValue(p.otherNotes || "")
        // ------------------------------------------

        return buildResponse({ status: "success", message: "อัปเดตข้อมูลผู้ป่วยสำเร็จ" });
      } else {
        return buildResponse({ status: "error", message: "ไม่พบหมายเลขสมาชิกนี้ในระบบ" });
      }
    }
    // ==========================================
    // 3. ฟีเจอร์เดิม: ลงทะเบียนคนใหม่
    // ==========================================
    const name     = p.name     || "";
    const email    = p.email    || "";
    const phone    = p.phone    || "";
    const dob      = p.dob      || "";
    const age      = p.age      || "";
    const symptoms = p.symptoms || "";
    
    // --- รับตัวแปรใหม่ที่ส่งมาจากหน้าเว็บ ---
    const idCard     = p.idCard     || "";
    const address    = p.address    || "";
    const udHospital = p.udHospital || "";
    const udMedicine = p.udMedicine || "";
    const allergy    = p.allergy    || "";
    const otherNotes = p.otherNotes || "";

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    // หากยังไม่มีข้อมูลเลย ให้สร้าง Header ใหม่ครอบคลุมช่องใหม่ทั้งหมด
    if (sheet.getLastRow() === 0) {
      const headers = [
        "ลำดับ", "หมายเลขสมาชิก", "วันที่ลงทะเบียน", "ชื่อ", "Email",
        "เบอร์โทรศัพท์", "วันเดือนปีเกิด", "อายุ (ปี)",
        "อาการป่วย", "URL รูปภาพ", "URL เอกสาร PDF", "คำวินิจฉัยโรค", "วันนัดครั้งถัดไป",
        "เลขบัตรประชาชน", "ที่อยู่", "โรคประจำตัวรับยาที่", "ยาโรคประจำตัว", "ประวัติแพ้ยา/อาหาร", "อื่นๆ"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#7C3AED").setFontColor("#ffffff");
    }

    const data = sheet.getDataRange().getValues();

    if (phone !== "") {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][5]).trim() === phone.trim()) {
          return buildResponse({ status: "duplicate", memberId: data[i][1], message: "เบอร์โทรนี้มีประวัติในระบบแล้ว" });
        }
      }
    }

    let nextMemberId = 100000;
    if (data.length > 1) {
      let lastId = parseInt(data[data.length - 1][1], 10);
      if (!isNaN(lastId) && lastId >= 100000) {
        nextMemberId = lastId + 1;
      }
    }

    const folder    = DriveApp.getFolderById(FOLDER_ID);
    const timestamp = new Date().getTime();
    let imageUrl = "";
    let pdfUrl   = "";

    if (p.imageBase64) {
      const blob = Utilities.newBlob(Utilities.base64Decode(p.imageBase64), p.imageMime || "image/jpeg", timestamp + "_" + (p.imageName || "photo.jpg"));
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageUrl = "https://drive.google.com/uc?id=" + file.getId();
    }

    if (p.pdfBase64) {
      const blob = Utilities.newBlob(Utilities.base64Decode(p.pdfBase64), "application/pdf", timestamp + "_" + (p.pdfName || "document.pdf"));
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfUrl = "https://drive.google.com/file/d/" + file.getId() + "/view";
    }

    const seq     = sheet.getLastRow();
    const dateStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");
    
    // บันทึกข้อมูลเรียงตาม Header (เว้นช่อง 11,12 ไว้ให้ คำวินิจฉัย กับ วันนัด เหมือนเดิม ส่วนข้อมูลใหม่ไปต่อท้าย)
    sheet.appendRow([
      seq, nextMemberId, dateStr, name, email, phone, dob, age, symptoms, 
      imageUrl, pdfUrl, "", "", 
      idCard, address, udHospital, udMedicine, allergy, otherNotes
    ]);

    return buildResponse({ status: "success", memberId: nextMemberId, message: "บันทึกข้อมูลสำเร็จ" });
  } catch (err) {
    return buildResponse({ status: "error", message: err.toString() });
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    
    // ==========================================
    // 1. ดึงข้อมูลประวัติการรักษาของแต่ละบุคคล
    // ==========================================
    if (action === "get_history") {
      const hn = (e.parameter.hn || "").trim();
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const historySheet = ss.getSheetByName("History");
      
      // ถ้ายังไม่มี Sheet History ให้คืนค่าว่าง
      if (!historySheet || historySheet.getLastRow() <= 1) {
        return buildResponse({ status: "success", data: [] });
      }

      const data = historySheet.getDataRange().getValues();
      const results = [];

      // ค้นหาข้อมูลจากบรรทัดที่ 2 เป็นต้นไป (เช็ค HN ที่คอลัมน์ index 1)
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]).trim() === hn) {
          results.push({
            date: data[i][0],        // วันที่มารักษา
            hn: data[i][1],          // รหัส HN
            name: data[i][2],        // ชื่อผู้ป่วย
            doctor: data[i][3],      // แพทย์ผู้รักษา
            symptoms: data[i][4],    // อาการ
            diagnosis: data[i][5],   // คำวินิจฉัยแพทย์
            medicine: data[i][6],    // ยาที่จ่าย
            right: data[i][7],       // สิทธิ์ที่เบิกจ่าย
            cost: data[i][8]         // ราคา ค่ารักษา
          });
        }
      }
      return buildResponse({ status: "success", data: results });
    }

    // ==========================================
    // 2. ดึงข้อมูลผู้ป่วยทั้งหมด (สำหรับค้นหา)
    // ==========================================
    if (action === "search") {
      const q = (e.parameter.q || "").trim().toLowerCase();
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(SHEET_NAME);
      
      if (!sheet || sheet.getLastRow() <= 1) {
        return buildResponse({ status: "success", data: [] });
      }

      const data = sheet.getDataRange().getValues();
      const results = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const memberId = String(row[1]).toLowerCase();
        const name     = String(row[3]).toLowerCase();
        const phone    = String(row[5]).toLowerCase();

        if (q === "" || memberId.includes(q) || phone.includes(q) || name.includes(q)) {
          results.push({
            seq: row[0], memberId: row[1], regDate: row[2], name: row[3],
            email: row[4], phone: row[5], dob: row[6], age: row[7],
            symptoms: row[8], imageUrl: row[9], pdfUrl: row[10],
            diagnosis: row[11] || "-", nextAppointment: row[12] || "-",
            idCard: row[13] || "", address: row[14] || "",
            udHospital: row[15] || "", udMedicine: row[16] || "",
            allergy: row[17] || "", otherNotes: row[18] || ""
          });
        }
      }
      return buildResponse({ status: "success", data: results });
    }

    return buildResponse({ status: "ok", message: "API พร้อมใช้งาน" });
    
  } catch (err) {
    return buildResponse({ status: "error", message: err.toString() });
  }
}