/**
 * =========================================================
 * PRS - Partner Relation System (Backend Core)
 * =========================================================
 */

// ชื่อชีตต่างๆ ตาม Database Schema ที่เราออกแบบไว้
const SPREADSHEET_ID = '1kN53znVSlaYqVBZFCSiNsWeUE_kKvtU3sZNDLZkVN6s'; // เปลี่ยนเป็น ID ของ Sheet คุณ
const DB = {
  LEADS: 'DB_Leads',
  TRAINING: 'DB_Training',
  LOGS: 'DB_Logs',
  CONFIG_SYS: 'Config_System',
  CONFIG_DROP: 'Config_Dropdowns',
  PROMOTIONS: 'DB_Promotions'
};

/**
 * 1. Routing - โหลดหน้า index.html
 */
function doGet(e) {
  let template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('PRS - Recruitment Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 2. Helper - สำหรับดึงไฟล์ HTML ย่อย (Components)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 3. [สำคัญ] ฟังก์ชันสร้างฐานข้อมูลอัตโนมัติ (Run แค่ครั้งแรก หรือเมื่อมีการเพิ่ม Schema)
 */
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const schemas = {
    [DB.LEADS]: ['Lead_ID', 'Timestamp', 'RecordDate', 'FullName', 'Phone', 'Channel', 'JobType', 'Province', 'District', 'Age', 'StartWithin', 'ChatLink', 'ContactStatus', 'ContactResult', 'IsApplicant', 'TrainingStatus', 'P_Value', 'Promo', 'Admin', 'Notes'],
    [DB.TRAINING]: ['Train_ID', 'Lead_ID', 'FullName', 'TrainDate', 'TransferDate', 'Amount', 'RescheduleCount', 'Status'],
    [DB.LOGS]: ['Log_ID', 'Timestamp', 'UserEmail', 'Action', 'Module', 'Details'],
    [DB.CONFIG_SYS]: ['Admin_Email', 'Admin_Name', 'Role', 'AutoAssign_Status', 'Current_Workload'],
    [DB.CONFIG_DROP]: ['Category', 'Value', 'IsActive'],
    [DB.PROMOTIONS]: ['Promo_ID', 'Promo_Name', 'Start_Date', 'End_Date', 'Description', 'Status']
  };

  for (const [sheetName, headers] of Object.entries(schemas)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    const firstCell = sheet.getRange(1, 1).getValue();
    if (firstCell !== headers[0]) {
      if (firstCell !== "") {
        sheet.insertRowBefore(1); 
      }
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
           .setFontWeight("bold").setBackground("#34495e").setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }
  }
  
  const dropSheet = ss.getSheetByName(DB.CONFIG_DROP);
  if(dropSheet.getLastRow() === 1) {
    const mockDrops = [
      ['Channel', 'Facebook Page', 'TRUE'], ['Channel', 'Line OA', 'TRUE'],
      ['ContactResult', 'ติดต่อได้ - คุยเบื้องต้น', 'TRUE'], 
      ['ContactResult', 'Add Line เรียบร้อย', 'TRUE'], 
      ['ContactResult', 'ไม่รับสาย', 'TRUE']
    ];
    dropSheet.getRange(2, 1, mockDrops.length, 3).setValues(mockDrops);
  }

  const promoSheet = ss.getSheetByName(DB.PROMOTIONS);
  if(promoSheet.getLastRow() === 1) {
    const today = new Date();
    const promo1Start = Utilities.formatDate(new Date(today.getFullYear(), 0, 1), "GMT+7", "yyyy-MM-dd");
    const promo1End = Utilities.formatDate(new Date(today.getFullYear(), 11, 31), "GMT+7", "yyyy-MM-dd");
    
    const promo2Start = Utilities.formatDate(today, "GMT+7", "yyyy-MM-dd");
    const nextMonth = new Date(); nextMonth.setMonth(today.getMonth() + 1);
    const promo2End = Utilities.formatDate(nextMonth, "GMT+7", "yyyy-MM-dd");

    const lastMonth = new Date(); lastMonth.setMonth(today.getMonth() - 1);
    const yesterday = new Date(); yesterday.setDate(today.getDate() - 1);
    const promo3Start = Utilities.formatDate(lastMonth, "GMT+7", "yyyy-MM-dd");
    const promo3End = Utilities.formatDate(yesterday, "GMT+7", "yyyy-MM-dd");

    const mockPromos = [
      ['PRM-001', 'เพื่อนแนะนำเพื่อน (Active)', promo1Start, promo1End, 'แนะนำเพื่อนรับ 500', 'Active'],
      ['PRM-002', 'แคมเปญ Flash Sale (Active)', promo2Start, promo2End, 'โปรด่วน', 'Active'],
      ['PRM-003', 'โปรโมชั่นเก่า (Expired)', promo3Start, promo3End, 'หมดเวลาแล้ว', 'Active']
    ];
    promoSheet.getRange(2, 1, mockPromos.length, 6).setValues(mockPromos);
  }
}

/**
 * 4. API - โหลดข้อมูลเริ่มต้น (Leads & Dropdowns)
 */
function getInitialData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const leadSheet = ss.getSheetByName(DB.LEADS);
    const dropSheet = ss.getSheetByName(DB.CONFIG_DROP);
    const promoSheet = ss.getSheetByName(DB.PROMOTIONS);
    const adminSheet = ss.getSheetByName(DB.CONFIG_SYS);
    
    if (!leadSheet || !dropSheet || !promoSheet || !adminSheet) {
      throw new Error("ยังไม่ได้สร้างฐานข้อมูล กรุณารันฟังก์ชัน setupDatabase ใน Code.gs ก่อน");
    }

    // 1. โหลดข้อมูล Leads
    const leadData = leadSheet.getDataRange().getDisplayValues();
    const leadHeaders = leadData.shift(); 
    let leads = [];
    if (leadData.length > 0) {
      leads = leadData.map(row => {
        let obj = {};
        leadHeaders.forEach((header, index) => obj[header] = row[index]);
        return obj;
      }).reverse(); 
    }

    // 2. โหลดข้อมูล Dropdowns ทั่วไป
    const dropData = dropSheet.getDataRange().getDisplayValues();
    dropData.shift();
    let dropdowns = { Channel: [], Promo: [], ContactResult: [] };
    dropData.forEach(row => {
      if(row[2] === 'TRUE' && dropdowns[row[0]] && row[0] !== 'Promo') {
        dropdowns[row[0]].push(row[1]);
      }
    });

    // 3. โหลดและคัดกรอง โปรโมชั่น (เฉพาะที่ไม่หมดอายุ)
    const promoData = promoSheet.getDataRange().getValues();
    promoData.shift();

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let activePromos = [];
    promoData.forEach(row => {
      if (row[1] && row[5] !== 'Inactive') { 
        const startDate = new Date(row[2]);
        startDate.setHours(0, 0, 0, 0);
        
        const endDate = new Date(row[3]);
        endDate.setHours(23, 59, 59, 999);

        if (today >= startDate && today <= endDate) {
          activePromos.push(row[1]); 
        }
      }
    });
    dropdowns.Promo = activePromos;

    // 4. 🌟 โหลดข้อมูล Admins 🌟
    const adminData = adminSheet.getDataRange().getDisplayValues();
    adminData.shift();
    const admins = adminData.map(row => ({
      email: row[0],
      name: row[1] || row[0], // ถ้าไม่มีชื่อเล่น ให้แสดงอีเมลแทน
      role: row[2],
      auto: row[3] === 'TRUE' ? 'ON' : 'OFF'
    }));

    return JSON.stringify({ status: 'success', leads: leads, dropdowns: dropdowns, admins: admins });
    
  } catch (error) {
    throw new Error(error.message); 
  }
}

/**
 * 5. API - บันทึกข้อมูล Lead ใหม่จาก Form (บันทึกชื่อเล่นแอดมินอัตโนมัติ)
 */
function saveNewLead(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DB.LEADS);
    const lastRow = Math.max(sheet.getLastRow(), 1);
    
    // สร้างรหัส PRS-YYMM-0000
    const newId = "PRS-" + Utilities.formatDate(new Date(), "GMT+7", "yyMM") + "-" + String(lastRow).padStart(4, '0');
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    
    // ดึงอีเมลของ Google Account ปัจจุบันที่เข้าใช้งาน
    const currentUserEmail = Session.getActiveUser().getEmail() || "ระบบแอดมินอัตโนมัติ";
    let currentAdminName = currentUserEmail;

    // 🌟 ค้นหาชื่อเล่นจากตาราง Config_System 🌟
    const adminSheet = ss.getSheetByName(DB.CONFIG_SYS);
    const adminData = adminSheet.getDataRange().getValues();
    for (let i = 1; i < adminData.length; i++) {
      if (adminData[i][0] === currentUserEmail) {
        currentAdminName = adminData[i][1] || currentUserEmail; // ได้ชื่อเล่นมาแล้ว
        break;
      }
    }

    const rowData = [
      newId, timestamp, formData.basicInfo?.date || '', formData.basicInfo?.name || '',
      formData.basicInfo?.phone || '', formData.basicInfo?.channel || '', formData.basicInfo?.type || '',
      formData.personalDetails?.province || '', formData.personalDetails?.district || '', formData.personalDetails?.age || '',
      formData.personalDetails?.startWithin || '', formData.personalDetails?.chatLink || '', formData.personalDetails?.contactStatus || '',
      formData.personalDetails?.contactResult || '', formData.personalDetails?.isApplicant || '', formData.assessment?.trainingStatus || '',
      formData.pValue || '', '', currentAdminName, formData.assessment?.notes || '' // ใช้ currentAdminName บันทึกลงช่อง Admin
    ];

    sheet.appendRow(rowData);
    writeLog(currentUserEmail, "CREATE_LEAD", "Leads", `Created Lead ID: ${newId}`);
    
    return JSON.stringify({ status: 'success', id: newId });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * 6. API - อัปเดตข้อมูล Lead เบื้องต้น
 */
function updateLead(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DB.LEADS);
    const leadData = sheet.getDataRange().getValues();
    
    let targetRow = -1;
    for (let i = 1; i < leadData.length; i++) {
      if (leadData[i][0] === formData.leadId) {
        targetRow = i + 1; break;
      }
    }

    if (targetRow > -1) {
      const currentUserEmail = Session.getActiveUser().getEmail() || "ระบบแอดมินอัตโนมัติ";

      sheet.getRange(targetRow, 3).setValue(formData.basicInfo.date);
      sheet.getRange(targetRow, 4).setValue(formData.basicInfo.name);
      sheet.getRange(targetRow, 5).setValue(formData.basicInfo.phone);
      sheet.getRange(targetRow, 6).setValue(formData.basicInfo.channel);
      sheet.getRange(targetRow, 7).setValue(formData.basicInfo.type);
      sheet.getRange(targetRow, 8).setValue(formData.personalDetails.province);
      sheet.getRange(targetRow, 9).setValue(formData.personalDetails.district);
      sheet.getRange(targetRow, 10).setValue(formData.personalDetails.age);
      sheet.getRange(targetRow, 11).setValue(formData.personalDetails.startWithin);
      sheet.getRange(targetRow, 12).setValue(formData.personalDetails.chatLink);
      sheet.getRange(targetRow, 13).setValue(formData.personalDetails.contactStatus);
      sheet.getRange(targetRow, 14).setValue(formData.personalDetails.contactResult);
      sheet.getRange(targetRow, 15).setValue(formData.personalDetails.isApplicant);
      
      // เราจะไม่อัปเดตช่อง Admin ทับ เพื่อรักษาว่าใครเป็นคน "สร้าง" ข้อมูลคนแรกไว้
      
      // อัปเดต Notes เบื้องต้น
      if(formData.assessment && formData.assessment.notes !== undefined) {
        sheet.getRange(targetRow, 20).setValue(formData.assessment.notes);
      }

      writeLog(currentUserEmail, "UPDATE_LEAD", "Leads", `Updated Lead ID: ${formData.leadId}`);
      return JSON.stringify({ status: 'success', id: formData.leadId });
    } else {
      return JSON.stringify({ status: 'error', message: 'ไม่พบรหัสผู้สมัครนี้' });
    }
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * 7. API - อัปเดตข้อมูล P-Value และ Assessment
 */
function updatePValue(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DB.LEADS);
    const leadData = sheet.getDataRange().getValues();
    const headers = leadData[0]; // ดึงแถว Header มาเพื่อหาตำแหน่งคอลัมน์
    
    let targetRow = -1;
    for (let i = 1; i < leadData.length; i++) {
      if (leadData[i][0] === data.leadId) { 
        targetRow = i + 1; break;
      }
    }

    if (targetRow > -1) {
      // ใช้ชื่อ Header ในการหาตำแหน่งเพื่อป้องกันคอลัมน์เคลื่อนเวลาเราเพิ่ม/ลดคอลัมน์ในอนาคต
      const colTraining = headers.indexOf('TrainingStatus') + 1;
      const colPValue = headers.indexOf('P_Value') + 1;
      const colPromo = headers.indexOf('Promo') + 1;
      const colNotes = headers.indexOf('Notes') + 1;
      const colReferrer = headers.indexOf('Referrer') + 1;
      const colBgCheck = headers.indexOf('BackgroundCheck') + 1;

      // 🌟 ดึงค่า P-Value เดิมก่อนอัปเดต เพื่อไปบันทึกเป็นประวัติ
      let oldPValue = "";
      if (colPValue > 0) oldPValue = sheet.getRange(targetRow, colPValue).getValue();

      // บันทึกข้อมูลเฉพาะคอลัมน์ที่มีอยู่จริงในชีต
      if (colTraining > 0) sheet.getRange(targetRow, colTraining).setValue(data.trainingStatus);
      if (colPValue > 0) sheet.getRange(targetRow, colPValue).setValue(data.pValue);
      if (colPromo > 0) sheet.getRange(targetRow, colPromo).setValue(data.promo);
      if (colNotes > 0) sheet.getRange(targetRow, colNotes).setValue(data.notes); // โน้ตล่าสุดจะถูกเซฟทับตรงนี้
      
      if (colReferrer > 0) sheet.getRange(targetRow, colReferrer).setValue(data.referrer);
      if (colBgCheck > 0) sheet.getRange(targetRow, colBgCheck).setValue(data.backgroundCheck);

      // 🌟 บันทึกประวัติการเปลี่ยนแปลงลง DB_Logs
      const logDetails = `[${data.leadId}] เปลี่ยน P-Value: ${oldPValue || 'ไม่มี'} -> ${data.pValue || 'ไม่มี'} | Note: ${data.notes || '-'}`;
      writeLog(Session.getActiveUser().getEmail() || "Unknown", "UPDATE_PVALUE", "P-Value", logDetails);

      return JSON.stringify({ status: 'success', id: data.leadId });
    } else {
      return JSON.stringify({ status: 'error', message: 'ไม่พบรหัสผู้สมัครนี้' });
    }
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * 8. API - ดึงข้อมูลโปรโมชั่นทั้งหมด (Gantt Chart)
 */
function getAllPromotions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DB.PROMOTIONS);
    const data = sheet.getDataRange().getDisplayValues(); 
    const headers = data.shift();
    
    const promos = data.map(row => {
      let obj = {};
      headers.forEach((header, index) => obj[header] = row[index]);
      return obj;
    });
    
    return JSON.stringify({ status: 'success', data: promos });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * 9. API - บันทึก/อัปเดต โปรโมชั่น
 */
function savePromotion(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DB.PROMOTIONS);
    
    if (formData.promoId) {
      // Update
      const data = sheet.getDataRange().getValues();
      let targetRow = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === formData.promoId) {
          targetRow = i + 1; break;
        }
      }
      if (targetRow > -1) {
        sheet.getRange(targetRow, 2).setValue(formData.promoName);
        sheet.getRange(targetRow, 3).setValue(formData.startDate);
        sheet.getRange(targetRow, 4).setValue(formData.endDate);
        sheet.getRange(targetRow, 5).setValue(formData.description);
        sheet.getRange(targetRow, 6).setValue(formData.status);
        writeLog(Session.getActiveUser().getEmail() || "Unknown", "UPDATE_PROMO", "Promotions", `Updated Promo ID: ${formData.promoId}`);
        return JSON.stringify({ status: 'success', id: formData.promoId });
      } else {
        return JSON.stringify({ status: 'error', message: 'ไม่พบรหัสโปรโมชั่น' });
      }
    } else {
      // Create
      const lastRow = Math.max(sheet.getLastRow(), 1);
      const newId = "PRM-" + Utilities.formatDate(new Date(), "GMT+7", "yyMM") + "-" + String(lastRow).padStart(3, '0');
      
      const rowData = [
        newId, formData.promoName, formData.startDate, formData.endDate, formData.description, formData.status || 'Active'
      ];
      sheet.appendRow(rowData);
      writeLog(Session.getActiveUser().getEmail() || "Unknown", "CREATE_PROMO", "Promotions", `Created Promo ID: ${newId}`);
      return JSON.stringify({ status: 'success', id: newId });
    }
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * 10. API - ลบโปรโมชั่น
 */
function deletePromotion(promoId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DB.PROMOTIONS);
    const data = sheet.getDataRange().getValues();
    
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === promoId) {
        targetRow = i + 1; break;
      }
    }
    
    if (targetRow > -1) {
      sheet.deleteRow(targetRow);
      writeLog(Session.getActiveUser().getEmail() || "Unknown", "DELETE_PROMO", "Promotions", `Deleted Promo ID: ${promoId}`);
      return JSON.stringify({ status: 'success', message: 'ลบข้อมูลสำเร็จ' });
    } else {
      return JSON.stringify({ status: 'error', message: 'ไม่พบรหัสโปรโมชั่น' });
    }
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * 11. API - ลบข้อมูลผู้สมัคร (Lead)
 */
function deleteLeadRecord(leadId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DB.LEADS);
    const data = sheet.getDataRange().getValues();
    
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === leadId) {
        targetRow = i + 1; break;
      }
    }
    
    if (targetRow > -1) {
      sheet.deleteRow(targetRow);
      writeLog(Session.getActiveUser().getEmail() || "Unknown", "DELETE_LEAD", "Leads", `Deleted Lead ID: ${leadId}`);
      return JSON.stringify({ status: 'success', message: 'ลบข้อมูลผู้สมัครสำเร็จ' });
    } else {
      return JSON.stringify({ status: 'error', message: 'ไม่พบรหัสผู้สมัครในระบบ' });
    }
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * 12. API - ดึงประวัติ Logs ของ Lead
 */
function getLeadLogs(leadId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName(DB.LOGS);
    const data = logSheet.getDataRange().getDisplayValues();
    data.shift(); // ลบ Header
    
    // กรองเอาเฉพาะบรรทัดที่ Details มีข้อความตรงกับ Lead_ID
    const logs = data.filter(row => row[5].includes(leadId)).map(row => {
      return {
        Timestamp: row[1],
        UserEmail: row[2],
        Action: row[3],
        Details: row[5]
      };
    }).reverse(); // เอาล่าสุดขึ้นก่อน
    
    return JSON.stringify({ status: 'success', data: logs });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * 13. API - จัดการสิทธิ์แอดมินและชื่อเล่น (หน้า Settings)
 */
function saveAdminConfig(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DB.CONFIG_SYS);
    const adminData = sheet.getDataRange().getValues();
    
    let targetRow = -1;
    for (let i = 1; i < adminData.length; i++) {
      if (adminData[i][0] === data.email) {
        targetRow = i + 1;
        break;
      }
    }

    const isAuto = data.autoAssign === true || data.autoAssign === 'TRUE' ? 'TRUE' : 'FALSE';
    
    // โครงสร้าง: [Admin_Email, Admin_Name, Role, AutoAssign_Status]
    const rowValues = [data.email, data.name || data.email, data.role, isAuto];
    
    if (targetRow > -1) {
      sheet.getRange(targetRow, 1, 1, 4).setValues([rowValues]); // อัปเดตข้อมูล 4 คอลัมน์แรก
    } else {
      sheet.appendRow([...rowValues, 0]); // สร้างใหม่ (0 คือ workload เริ่มต้น ในคอลัมน์ที่ 5)
    }
    
    writeLog(Session.getActiveUser().getEmail(), "UPDATE_ADMIN", "Settings", `Set Admin: ${data.name || data.email} as ${data.role}`);
    return JSON.stringify({ status: 'success' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.toString() });
  }
}

/**
 * 14. API - ลบสิทธิ์แอดมิน (หน้า Settings)
 */
function deleteAdminConfig(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DB.CONFIG_SYS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    
    writeLog(Session.getActiveUser().getEmail(), "DELETE_ADMIN", "Settings", `Removed Admin: ${email}`);
    return JSON.stringify({ status: 'success' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.toString() });
  }
}

/**
 * 15. API - อัปเดตรายการ Dropdown (เพิ่ม/ลบ/จัดเรียงลำดับ)
 */
function updateDropdownConfig(category, items) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DB.CONFIG_DROP);
    let data = sheet.getDataRange().getValues();
    
    // ดึงข้อมูลที่ไม่ใช่หมวดหมู่เป้าหมายเก็บไว้
    let keptRows = [data[0]]; // เก็บ Header ไว้
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== category) {
        keptRows.push(data[i]);
      }
    }
    
    // นำรายการใหม่มาต่อท้าย (จัดเรียงตามลำดับที่ส่งมา)
    const newItemsRows = items.map(item => [category, item, 'TRUE']);
    const finalData = keptRows.concat(newItemsRows);
    
    // ล้างชีตแล้วเขียนใหม่ทั้งหมด เพื่อประสิทธิภาพและความถูกต้องของลำดับ
    sheet.clear();
    sheet.getRange(1, 1, finalData.length, 3).setValues(finalData);
    
    writeLog(Session.getActiveUser().getEmail(), "UPDATE_DROPDOWN", "Settings", `Updated list for ${category}`);
    return JSON.stringify({ status: 'success' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.toString() });
  }
}

/**
 * 16. Helper - เขียน History Log เพื่อเก็บประวัติการทำรายการทุกอย่างในระบบ
 */
function writeLog(email, action, module, details) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DB.LOGS);
    if (sheet) {
      const logId = "LOG-" + new Date().getTime();
      const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
      sheet.appendRow([logId, timestamp, email, action, module, details]);
    }
  } catch (e) {
    console.error("Failed to write log: ", e);
  }
}

/**
 * 17. API - จ่ายงานอัตโนมัติ (Auto-Assign) 1 เคส
 */
function autoAssignLead(leadId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const adminSheet = ss.getSheetByName(DB.CONFIG_SYS);
    const leadSheet = ss.getSheetByName(DB.LEADS);
    
    // 1. หาแอดมินที่เปิด Auto และงานน้อยสุด
    const adminData = adminSheet.getDataRange().getValues();
    let selectedAdmin = null;
    let minWorkload = Infinity;

    for (let i = 1; i < adminData.length; i++) {
      const email = adminData[i][0];
      const name = adminData[i][1];
      const isAuto = adminData[i][3];
      const workload = Number(adminData[i][4]) || 0;

      if (isAuto === 'TRUE' || isAuto === true) {
        if (workload < minWorkload) {
          minWorkload = workload;
          selectedAdmin = { email, name, rowIndex: i + 1, currentWorkload: workload };
        }
      }
    }

    if (!selectedAdmin) {
      return JSON.stringify({ status: 'error', message: 'ไม่มีแอดมินที่เปิดรับงานอัตโนมัติ (Auto-Assign) ในระบบ' });
    }

    // 2. อัปเดตตาราง Leads
    const leadData = leadSheet.getDataRange().getValues();
    const headers = leadData[0];
    const adminColIndex = headers.indexOf('Admin') + 1;
    
    let leadRow = -1;
    for (let i = 1; i < leadData.length; i++) {
      if (leadData[i][0] === leadId) {
        leadRow = i + 1;
        break;
      }
    }

    if (leadRow > -1 && adminColIndex > 0) {
      leadSheet.getRange(leadRow, adminColIndex).setValue(selectedAdmin.email);
      // 3. เพิ่ม Workload ให้แอดมินคนนั้น +1
      adminSheet.getRange(selectedAdmin.rowIndex, 5).setValue(selectedAdmin.currentWorkload + 1);
      writeLog(Session.getActiveUser().getEmail() || "System", "AUTO_ASSIGN", "Leads", `Assigned Lead ${leadId} to ${selectedAdmin.name}`);
      return JSON.stringify({ status: 'success', assignedName: selectedAdmin.name });
    }

    return JSON.stringify({ status: 'error', message: 'ไม่พบรหัสผู้สมัคร' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.toString() });
  }
}

/**
 * 18. API - เลือกจ่ายงานเอง (Manual Assign)
 */
function manualAssignLead(leadId, adminEmail) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const leadSheet = ss.getSheetByName(DB.LEADS);
    
    const leadData = leadSheet.getDataRange().getValues();
    const headers = leadData[0];
    const adminColIndex = headers.indexOf('Admin') + 1;
    
    let leadRow = -1;
    for (let i = 1; i < leadData.length; i++) {
      if (leadData[i][0] === leadId) {
        leadRow = i + 1;
        break;
      }
    }

    if (leadRow > -1 && adminColIndex > 0) {
      leadSheet.getRange(leadRow, adminColIndex).setValue(adminEmail);
      writeLog(Session.getActiveUser().getEmail() || "System", "MANUAL_ASSIGN", "Leads", `Assigned Lead ${leadId} to ${adminEmail}`);
      return JSON.stringify({ status: 'success' });
    }

    return JSON.stringify({ status: 'error', message: 'ไม่พบรหัสผู้สมัคร' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.toString() });
  }
}

/**
 * 19. API - สุ่มจ่ายงานทั้งหมด (Bulk Auto-Assign)
 */
function bulkAutoAssignLeads(leadIds) {
  try {
    if (!leadIds || leadIds.length === 0) return JSON.stringify({ status: 'success' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const adminSheet = ss.getSheetByName(DB.CONFIG_SYS);
    const leadSheet = ss.getSheetByName(DB.LEADS);

    // 1. โหลดแอดมินที่เปิด Auto
    let adminData = adminSheet.getDataRange().getValues();
    let availableAdmins = [];
    
    for (let i = 1; i < adminData.length; i++) {
      const email = adminData[i][0];
      const name = adminData[i][1];
      const isAuto = adminData[i][3];
      const workload = Number(adminData[i][4]) || 0;

      if (isAuto === 'TRUE' || isAuto === true) {
        availableAdmins.push({ email, name, rowIndex: i + 1, workload: workload });
      }
    }

    if (availableAdmins.length === 0) {
      return JSON.stringify({ status: 'error', message: 'ไม่มีแอดมินที่เปิดรับงานอัตโนมัติ (Auto-Assign)' });
    }

    // 2. โหลด Leads
    const leadData = leadSheet.getDataRange().getValues();
    const headers = leadData[0];
    const adminColIndex = headers.indexOf('Admin') + 1;

    if (adminColIndex === 0) return JSON.stringify({ status: 'error', message: 'ไม่พบคอลัมน์ Admin' });

    let assignCount = 0;

    // 3. วนลูปแจกงานให้แอดมินที่งานน้อยสุด ณ ขณะนั้น
    leadIds.forEach(leadId => {
      // เรียงให้คนที่งานน้อยสุดอยู่ลำดับแรกเสมอ
      availableAdmins.sort((a, b) => a.workload - b.workload);
      let selectedAdmin = availableAdmins[0];

      let leadRow = -1;
      for (let i = 1; i < leadData.length; i++) {
        if (leadData[i][0] === leadId) {
          leadRow = i + 1;
          break;
        }
      }

      if (leadRow > -1) {
        // บันทึกแอดมินลง Sheet
        leadSheet.getRange(leadRow, adminColIndex).setValue(selectedAdmin.email);
        // เพิ่มจำนวนงานให้แอดมินคนนี้ เพื่อให้คนถัดไปมีโอกาสรับงานบ้าง
        selectedAdmin.workload += 1;
        assignCount++;
      }
    });

    // 4. บันทึก Workload กลับลงแผ่นตั้งค่าแอดมิน
    availableAdmins.forEach(admin => {
      adminSheet.getRange(admin.rowIndex, 5).setValue(admin.workload);
    });

    writeLog(Session.getActiveUser().getEmail() || "System", "BULK_AUTO_ASSIGN", "Leads", `Bulk assigned ${assignCount} leads.`);
    
    return JSON.stringify({ status: 'success', assignedCount: assignCount });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.toString() });
  }
}

/**
 * 20. API - ดึงประวัติการอัปเดต P-Value ของ Lead
 */
function getPValueHistory(leadId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName(DB.LOGS);
    const data = logSheet.getDataRange().getDisplayValues();
    data.shift(); // ลบ Header
    
    // กรองเอาเฉพาะ Action === 'UPDATE_PVALUE' และ Details มี Lead_ID
    const logs = data.filter(row => row[3] === 'UPDATE_PVALUE' && row[5].includes(`[${leadId}]`)).map(row => {
      // สกัดเอาเฉพาะข้อความส่วนที่เป็นการเปลี่ยนแปลงและ Note
      const details = row[5].replace(`[${leadId}] `, '');
      return {
        Timestamp: row[1],
        UserEmail: row[2],
        Details: details
      };
    }).reverse(); // เอาล่าสุดขึ้นก่อน
    
    return JSON.stringify({ status: 'success', data: logs });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * =========================================================
 * 🌟 21. API - ดึงข้อมูลผู้ใช้งานปัจจุบัน (แทนการใช้ระบบ Login)
 * =========================================================
 */
function getCurrentSession() {
  try {
    // ดึงอีเมลจากบัญชี Google Workspace ที่ใช้เปิดหน้าต่างนี้
    const email = Session.getActiveUser().getEmail();
    
    if (!email) {
      return JSON.stringify({ 
        status: 'error', 
        message: 'ไม่พบข้อมูลบัญชีอีเมล ระบบอาจไม่ได้ถูกตั้งค่าให้เข้าถึงด้วย Google Account (ตั้งค่า Deploy ให้เป็น User accessing the web app)' 
      });
    }
    
    return JSON.stringify({ status: 'success', email: email });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}
