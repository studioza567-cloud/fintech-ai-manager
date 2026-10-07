/**
 * ==============================================================================
 * AI Subscription & Expense Manager - Google Apps Script Backend (Production Build)
 * ==============================================================================
 * 
 * ระบบ Backend อัตโนมัติสำหรับจัดการฐานข้อมูลใน Google Sheets และส่ง Gmail Notification
 * แจ้งเตือนรอบชำระ Subscription อัตโนมัติตาม Time-driven Trigger แม้ปิดหน้าเว็บ
 */

const SHEET_NAME = "Subscriptions";
const TIMEZONE = "Asia/Bangkok";

/**
 * 1. ฟังก์ชันตั้งค่าชีตครั้งแรก (สร้าง Header และจัดรูปแบบ)
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const headers = [
    "id",
    "merchant",
    "amount",
    "category",
    "recurrence",
    "lastPaidDate",
    "nextBillingDate",
    "reminderDays",
    "email",
    "active",
    "lastReminderSentDate",
    "createdAt",
    "updatedAt"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#0f172a")
    .setFontColor("#10b981");
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * 2. ฟังก์ชันตั้งค่า Time-driven Trigger อัตโนมัติ (รันทุกวันเวลา 08:00 น.)
 */
function setupTrigger() {
  // ลบ Trigger เดิมที่มีชื่อซ้ำ
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "checkSubscriptionReminders") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // สร้าง Trigger ใหม่ให้รันทุกวันเวลา 08:00 - 09:00 น.
  ScriptApp.newTrigger("checkSubscriptionReminders")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log("Trigger created successfully for checkSubscriptionReminders at 08:00 daily.");
}

/**
 * 3. รับคำขอ GET (Health Check / อ่านข้อมูล)
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "ping";

  if (action === "ping") {
    return jsonResponse({
      status: "success",
      message: "AI Subscription Manager Backend is online and ready.",
      timestamp: new Date().toISOString()
    });
  }

  if (action === "getSubscriptions") {
    const list = getAllSubscriptions();
    return jsonResponse({ status: "success", data: list });
  }

  return jsonResponse({ status: "error", message: "Unknown GET action." });
}

/**
 * 4. รับคำขอ POST จาก Frontend
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: "error", message: "No post data received." });
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    // 4.1 ส่งอีเมลทดสอบจริง (Send Test Email)
    if (action === "sendTestEmail") {
      const email = payload.email;
      if (!email || email.indexOf("@") === -1) {
        return jsonResponse({ status: "error", message: "Invalid email address." });
      }

      const subject = "Subscription Alert Test";
      const body = "ระบบ Gmail Notification ทำงานเรียบร้อยแล้ว\n\nThis is a test email from\nAI Subscription & Expense Manager.\n\nระบบพร้อมส่งการแจ้งเตือนรอบบิลอัตโนมัติไปยัง Gmail ของคุณเมื่อถึงกำหนด!";

      GmailApp.sendEmail(
        email,
        subject,
        body,
        {
          from: "subscription.manager.alert@gmail.com",
          name: "Subscription Manager"
        }
      );

      return jsonResponse({
        status: "success",
        message: "ส่งอีเมลทดสอบไปยัง " + email + " สำเร็จแล้ว!"
      });
    }

    // 4.2 บันทึก/ซิงค์รายการ Subscription ทั้งหมด (Sync Subscriptions)
    if (action === "syncSubscriptions") {
      const subscriptions = payload.subscriptions || [];
      const userEmail = payload.email || "";
      syncSubscriptionsToSheet(subscriptions, userEmail);
      return jsonResponse({
        status: "success",
        message: "ซิงค์ข้อมูล Subscription ไปยัง Google Sheets สำเร็จ (" + subscriptions.length + " รายการ)",
        count: subscriptions.length
      });
    }

    // 4.0 ดึงข้อมูล Subscription ทั้งหมด (Get Subscriptions via POST)
    if (action === "getSubscriptions") {
      const list = getAllSubscriptions();
      return jsonResponse({ status: "success", data: list });
    }

    // 4.3 อัปเดตเมื่อผู้ใช้กด Mark as Paid
    if (action === "markAsPaid") {
      const subId = payload.id;
      const todayStr = getTodayString();
      const updated = updateMarkAsPaidInSheet(subId, todayStr);
      return jsonResponse({
        status: updated ? "success" : "not_found",
        message: updated ? "บันทึกชำระเงินและเลื่อนรอบบิลในชีตสำเร็จ" : "ไม่พบ Subscription ID นี้"
      });
    }

    // 4.4 รีเซ็ตข้อมูลทั้งหมด (Reset Data)
    if (action === "resetData") {
      resetSheetData();
      return jsonResponse({
        status: "success",
        message: "ล้างข้อมูล Subscription ใน Google Sheets เรียบร้อยแล้ว"
      });
    }

    // 4.5 สั่งให้รันตรวจสอบและส่งอีเมลแจ้งเตือนทันที (สำหรับทดสอบ)
    if (action === "triggerCheck") {
      const sentCount = checkSubscriptionReminders();
      return jsonResponse({
        status: "success",
        message: "รันตรวจสอบรอบบิลสำเร็จ ส่งอีเมลไปแล้ว " + sentCount + " ฉบับ",
        sentCount: sentCount
      });
    }

    // 4.6 ลบรายการ Subscription ตาม ID (Delete Subscription)
    if (action === "deleteSubscription") {
      const subId = payload.id;
      const deleted = deleteSubscriptionInSheet(subId);
      return jsonResponse({
        status: deleted ? "success" : "not_found",
        message: deleted ? "ลบรายการใน Google Sheets สำเร็จ" : "ไม่พบ Subscription ID นี้ในชีต"
      });
    }

    return jsonResponse({ status: "error", message: "Unknown POST action: " + action });

  } catch (err) {
    Logger.log("Error in doPost: " + err.toString());
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * 5. ฟังก์ชันแกนหลัก: ตรวจสอบและส่งการแจ้งเตือน (Scheduler Core)
 * ทำงานอัตโนมัติวันละ 1 ครั้งตาม Time-driven Trigger
 */
function checkSubscriptionReminders() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 0; // มีแค่ Header

  const todayStr = getTodayString();
  const today = parseDateOnly(todayStr);

  let emailsSent = 0;

  // คอลัมน์ index:
  // 0:id, 1:merchant, 2:amount, 3:category, 4:recurrence, 5:lastPaidDate,
  // 6:nextBillingDate, 7:reminderDays, 8:email, 9:active, 10:lastReminderSentDate
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const subId = row[0];
    const merchant = row[1];
    const amount = Number(row[2]) || 0;
    const nextBillingDateStr = formatDateStr(row[6]);
    const reminderDays = parseInt(row[7], 10) || 1;
    const email = row[8];
    const active = (row[9] === true || row[9] === "true" || row[9] === 1);
    const lastReminderSentDate = formatDateStr(row[10]);

    if (!active || !nextBillingDateStr || !email) continue;

    const billingDate = parseDateOnly(nextBillingDateStr);
    const diffTime = billingDate.getTime() - today.getTime();
    const daysUntilBilling = Math.round(diffTime / (1000 * 60 * 60 * 24));

    // ตรวจสอบเงื่อนไขการส่งแจ้งเตือน:
    // 1. ถึงวันแจ้งเตือนล่วงหน้า (daysUntilBilling === reminderDays)
    // 2. หรือถึงวันครบกำหนดจ่ายพอดี (daysUntilBilling === 0)
    const shouldRemind = (daysUntilBilling === reminderDays) || (daysUntilBilling === 0);

    if (shouldRemind) {
      // 🛡️ ANTI-DUPLICATE: ตรวจสอบว่าวันนี้เคยส่งเตือนรายการนี้ไปแล้วหรือยัง
      if (lastReminderSentDate === todayStr) {
        Logger.log("Skipping duplicate reminder for " + merchant + " (already sent today " + todayStr + ")");
        continue;
      }

      // ส่งอีเมลจริง
      const success = sendSubscriptionAlertEmail(email, merchant, amount, nextBillingDateStr, daysUntilBilling);

      if (success) {
        // อัปเดต lastReminderSentDate และ updatedAt ใน Google Sheets
        sheet.getRange(i + 1, 11).setValue(todayStr);
        sheet.getRange(i + 1, 13).setValue(new Date().toISOString());
        emailsSent++;
        Logger.log("Sent reminder for " + merchant + " to " + email);
      }
    } else if (daysUntilBilling < 0) {
      // จัดการรายการที่เลยกำหนดชำระ (Overdue) โดยไม่ส่งเตือนซ้ำรอบเดิม
      Logger.log(merchant + " is overdue by " + Math.abs(daysUntilBilling) + " days.");
    }
  }

  return emailsSent;
}

/**
 * 6. ฟังก์ชันสร้างข้อความและส่ง Email ภาษาไทย
 */
function sendSubscriptionAlertEmail(email, merchant, amount, nextBillingDateStr, daysUntilBilling) {
  try {
    const thaiDate = formatThaiDate(nextBillingDateStr);
    let subject = "";
    let timeLabel = "";

    if (daysUntilBilling === 1) {
      subject = "Subscription Alert – " + merchant + " อีก 1 วันจะถึงวันจ่าย";
      timeLabel = "1 วัน (จะเรียกเก็บเงินพรุ่งนี้)";
    } else if (daysUntilBilling === 0) {
      subject = "Subscription Alert – " + merchant + " ถึงกำหนดจ่ายวันนี้";
      timeLabel = "ถึงกำหนดจ่ายวันนี้";
    } else {
      subject = "Subscription Alert – " + merchant + " อีก " + daysUntilBilling + " วันจะถึงวันจ่าย";
      timeLabel = daysUntilBilling + " วัน";
    }

    const body = 
      "สวัสดี\n\n" +
      "Subscription ของคุณกำลังจะถึงรอบเรียกเก็บเงิน\n\n" +
      "━━━━━━━━━━━━━━━━━━\n" +
      "รายการ: " + merchant + "\n" +
      "จำนวนเงิน: ฿" + amount.toLocaleString("th-TH", { minimumFractionDigits: 2 }) + "\n" +
      "วันเรียกเก็บ: " + thaiDate + "\n" +
      "เหลือ: " + timeLabel + "\n" +
      "━━━━━━━━━━━━━━━━━━\n\n" +
      "กรุณาตรวจสอบยอดเงินก่อนถึงรอบเรียกเก็บ\n\n" +
      "Subscription Alert\n" +
      "AI Subscription & Expense Manager";

    GmailApp.sendEmail(
      email,
      subject,
      body,
      {
        from: "subscription.manager.alert@gmail.com",
        name: "Subscription Manager"
      }
    );
    
    return true;
  } catch (e) {
    Logger.log("Failed to send email to " + email + ": " + e.toString());
    return false;
  }
}

/**
 * 7. Helpers สำหรับจัดการข้อมูลใน Google Sheets
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = setupSheet();
  }
  return sheet;
}

function getAllSubscriptions() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue; // ข้ามแถวที่ว่างเปล่า
    list.push({
      id: String(row[0]),
      merchant: String(row[1] || ""),
      amount: Number(row[2]) || 0,
      category: String(row[3] || "Subscription"),
      recurrence: String(row[4] || "monthly"),
      lastPaidDate: formatDateStr(row[5]),
      nextBillingDate: formatDateStr(row[6]),
      reminderDays: parseInt(row[7], 10) || 1,
      email: String(row[8] || ""),
      active: (row[9] === true || row[9] === "true" || row[9] === 1 || row[9] === "TRUE"),
      lastReminderSentDate: formatDateStr(row[10]),
      createdAt: row[11] ? String(row[11]) : "",
      updatedAt: row[12] ? String(row[12]) : ""
    });
  }
  return list;
}

function syncSubscriptionsToSheet(subscriptions, defaultEmail) {
  // 🛡️ SAFETY GUARD: ป้องกันการเผลอลบข้อมูลทั้งหมดหาก subscriptions เป็น array ว่าง
  // การล้างข้อมูลทั้งหมดต้องทำผ่าน action: "resetData" เท่านั้น
  if (!Array.isArray(subscriptions) || subscriptions.length === 0) {
    Logger.log("syncSubscriptionsToSheet skipped: empty subscriptions array.");
    return;
  }

  const sheet = getOrCreateSheet();
  const existingMap = {};
  const currentData = sheet.getDataRange().getValues();

  // เก็บข้อมูลเดิมเพื่อคงสถานะ lastReminderSentDate
  for (let i = 1; i < currentData.length; i++) {
    const rowId = String(currentData[i][0]);
    existingMap[rowId] = {
      lastReminderSentDate: formatDateStr(currentData[i][10]),
      createdAt: currentData[i][11]
    };
  }

  // ลบข้อมูลเดิมทั้งหมดยกเว้น Header
  if (currentData.length > 1) {
    sheet.deleteRows(2, currentData.length - 1);
  }

  const nowIso = new Date().toISOString();
  const rows = [];

  for (let s of subscriptions) {
    const id = String(s.id || "sub_" + Date.now());
    const existing = existingMap[id] || {};
    const row = [
      id,
      s.merchant || "Subscription",
      Number(s.amount) || 0,
      s.category || "Subscription",
      s.recurrence || "monthly",
      s.lastPaidDate || getTodayString(),
      s.nextBillingDate || "",
      parseInt(s.reminderDays, 10) || 1,
      s.email || defaultEmail || "",
      s.active !== false && s.active !== "false",
      s.lastReminderSentDate || existing.lastReminderSentDate || "",
      existing.createdAt || s.createdAt || nowIso,
      nowIso
    ];
    rows.push(row);
  }

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function deleteSubscriptionInSheet(subId) {
  if (!subId) return false;
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(subId).trim()) {
      sheet.deleteRow(i + 1);
      Logger.log("Deleted subscription from sheet: " + subId);
      return true;
    }
  }
  return false;
}

function updateMarkAsPaidInSheet(subId, todayStr) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(subId)) {
      const recurrence = data[i][4] || "monthly";
      const currentBillingDate = formatDateStr(data[i][6]) || todayStr;
      const newNextBilling = addRecurrenceDate(currentBillingDate, recurrence);

      sheet.getRange(i + 1, 6).setValue(todayStr); // lastPaidDate = today
      sheet.getRange(i + 1, 7).setValue(newNextBilling); // nextBillingDate = newNextBilling
      sheet.getRange(i + 1, 11).setValue(""); // reset lastReminderSentDate = null
      sheet.getRange(i + 1, 13).setValue(new Date().toISOString()); // updatedAt
      return true;
    }
  }
  return false;
}

function resetSheetData() {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  Logger.log("All subscriptions in sheet cleared.");
}

/**
 * 8. Utilities & Date Formatting
 */
function getTodayString() {
  return Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
}

function formatDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, TIMEZONE, "yyyy-MM-dd");
  }
  const str = String(val).trim();
  if (str.length >= 10 && str.charAt(4) === "-" && str.charAt(7) === "-") {
    return str.substring(0, 10);
  }
  return str;
}

function parseDateOnly(dateStr) {
  const parts = dateStr.split("-");
  return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
}

function formatThaiDate(dateStr) {
  if (!dateStr) return "-";
  const parts = dateStr.split("-");
  if (parts.length !== 3) return dateStr;
  const y = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10) - 1;
  const d = parseInt(parts[2], 10);

  const thaiMonths = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
  ];
  return d + " " + thaiMonths[m] + " " + (y + 543) + " (" + y + ")";
}

function addRecurrenceDate(dateStr, recurrence) {
  const parts = dateStr.split("-");
  const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));

  if (recurrence === "weekly") {
    d.setDate(d.getDate() + 7);
  } else if (recurrence === "yearly") {
    d.setFullYear(d.getFullYear() + 1);
  } else {
    // monthly default
    const originalDay = d.getDate();
    d.setMonth(d.getMonth() + 1);
    if (d.getDate() !== originalDay && d.getDate() < 5) {
      d.setDate(0);
    }
  }

  const ry = d.getFullYear();
  const rm = ("0" + (d.getMonth() + 1)).slice(-2);
  const rd = ("0" + d.getDate()).slice(-2);
  return ry + "-" + rm + "-" + rd;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function showAliases() {
  Logger.log(GmailApp.getAliases());
}
