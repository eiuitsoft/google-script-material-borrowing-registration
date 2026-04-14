/**
 * NotifyBot.gs - Thông báo bot khi có phiếu mới
 * Sheet: Yêu cầu mượn TTB
 *
 * HƯỚNG DẪN SỬ DỤNG:
 * 1. Copy file này vào Apps Script của sheet nguồn
 * 2. Thay BOT_WEBAPP_URL bằng URL thực của bot
 * 3. Setup trigger: onEdit hoặc onChange
 */

// ==================== CONFIG ====================

// Dùng chung webhook từ CONFIG nếu có (chạy cùng project),
// fallback sang URL mặc định nếu file được copy chạy độc lập.
const BOT_WEBAPP_URL =
  (typeof CONFIG !== "undefined" && CONFIG.N8N_WEBHOOK_URL) ||
  "https://n8n.eiu.vn/webhook/82bbf839-8656-4e50-b1a2-19fd896e43fa";

// Tên sheet cần theo dõi
const PHIEU_NHAP_SHEET = "Phiếu nhập";

// ==================== TRIGGER FUNCTION ====================

/**
 * Trigger khi có thay đổi trong sheet
 * Chỉ xử lý khi có row mới được thêm vào sheet "Phiếu nhập"
 * @param {Object} e - Event object
 */
function onEdit(e) {
  try {
    // Check event object
    if (!e || !e.range) {
      console.log("No event or range");
      return;
    }

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    // Chỉ xử lý sheet "Phiếu nhập"
    if (sheetName !== PHIEU_NHAP_SHEET) {
      console.log("Not Phieu nhap sheet, skip");
      return;
    }

    const row = e.range.getRow();

    // Skip header row (row 106)
    if (row <= 106) {
      console.log("Header row, skip");
      return;
    }

    // Lấy dữ liệu row
    const rowData = sheet.getRange(row, 1, 1, 14).getValues()[0];

    // Check nếu có ID_Dexuat (cột A) - tức là row đã có data đầy đủ
    if (!rowData[0]) {
      console.log("No ID_Dexuat, skip");
      return;
    }

    // Check column N (Cập nhật lúc) - nếu là row mới thì timestamp sẽ gần với hiện tại
    const capNhatLuc = rowData[13]; // Column N
    const now = new Date();

    if (capNhatLuc instanceof Date) {
      const diffMinutes = (now - capNhatLuc) / (1000 * 60);

      // Chỉ notify nếu tạo trong vòng 5 phút
      if (diffMinutes > 5) {
        console.log("Old row (created " + diffMinutes + " minutes ago), skip");
        return;
      }
    }

    // Parse data từ row
    const ticketData = {
      id: rowData[0], // A: ID_Dexuat
      nguoiDangKy: rowData[1], // B: Tên người đăng ký
      giangVien: rowData[2], // C: Giảng viên phụ trách
      phone: rowData[3], // D: Số điện thoại
      email: rowData[4], // E: Email
      maMH: rowData[5], // F: Mã môn học
      ngayHoc: formatDate(rowData[6]), // G: Ngày học
      gioHoc: rowData[7], // H: Giờ học
      phong: rowData[8], // I: Phòng/Lab
      loaiLab: rowData[9], // J: Loại lab
      ngayNhan: formatDate(rowData[10]), // K: Ngày nhận
      trangThai: rowData[11], // L: Trạng thái
    };

    console.log("Detected new ticket: " + ticketData.id);

    // Gọi bot webhook
    notifyBot(ticketData);
  } catch (err) {
    console.log("onEdit error: " + err.message);
  }
}

// ==================== NOTIFY BOT ====================

/**
 * Gửi notification đến bot
 * @param {Object} ticketData - Dữ liệu phiếu
 */
function notifyBot(ticketData) {
  try {
    // Check bot URL đã config chưa
    if (!BOT_WEBAPP_URL || BOT_WEBAPP_URL === "YOUR_BOT_WEBAPP_URL_HERE") {
      console.log("ERROR: BOT_WEBAPP_URL chưa được config!");
      console.log("Vui lòng cập nhật BOT_WEBAPP_URL ở đầu file NotifyBot.gs");
      return;
    }

    const payload = {
      action: "NEW_TICKET",
      source: "PHIEU_NHAP",
      timestamp: new Date().toISOString(),
      data: ticketData,
    };

    console.log("Calling bot webhook: " + BOT_WEBAPP_URL);
    console.log("Payload: " + JSON.stringify(payload));

    const response = UrlFetchApp.fetch(BOT_WEBAPP_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    console.log("Bot response code: " + responseCode);
    console.log("Bot response: " + responseText);

    if (responseCode === 200) {
      console.log("✅ Bot notified successfully");
    } else {
      console.log("⚠️ Bot returned non-200 status: " + responseCode);
    }
  } catch (err) {
    console.log("❌ notifyBot error: " + err.message);
  }
}

// ==================== HELPER FUNCTIONS ====================

/**
 * Format date để gửi cho bot
 * @param {Date|string} date - Date object hoặc string
 * @returns {string} - Formatted date (dd/MM/yyyy)
 */
function formatDate(date) {
  if (!date) return "";

  if (date instanceof Date) {
    return Utilities.formatDate(date, "GMT+7", "dd/MM/yyyy");
  }

  return date.toString();
}

// ==================== TEST FUNCTION ====================

/**
 * Test function - Gọi thủ công để test
 * Sẽ lấy row cuối cùng trong sheet và gửi notification
 * KHÔNG CHECK TIMESTAMP - chỉ dùng để test
 */
function testNotifyBot() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PHIEU_NHAP_SHEET);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 106) {
    console.log("No data rows to test");
    return;
  }

  const rowData = sheet.getRange(lastRow, 1, 1, 14).getValues()[0];

  // Check có ID không
  if (!rowData[0]) {
    console.log("Row " + lastRow + " has no ID, try another row");
    return;
  }

  const ticketData = {
    id: rowData[0],
    nguoiDangKy: rowData[1],
    giangVien: rowData[2],
    phone: rowData[3],
    email: rowData[4],
    maMH: rowData[5],
    ngayHoc: formatDate(rowData[6]),
    gioHoc: rowData[7],
    phong: rowData[8],
    loaiLab: rowData[9],
    ngayNhan: formatDate(rowData[10]),
    trangThai: rowData[11],
  };

  console.log("🧪 TEST MODE - Forcing notification for row " + lastRow);
  console.log("Ticket data: " + JSON.stringify(ticketData));

  // Gửi notification (bỏ qua check timestamp)
  notifyBot(ticketData);
}

/**
 * 🧪 TEST: Gọi bot trực tiếp KHÔNG QUA WEBHOOK
 * Giống như cách bot test với testHandleNewTicket
 * Chỉ dùng để test, không dùng trong production
 */
function testDirectCallBot() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PHIEU_NHAP_SHEET);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 106) {
    console.log("No data rows to test");
    return;
  }

  const rowData = sheet.getRange(lastRow, 1, 1, 14).getValues()[0];

  if (!rowData[0]) {
    console.log("Row " + lastRow + " has no ID");
    return;
  }

  const ticketData = {
    id: rowData[0],
    nguoiDangKy: rowData[1],
    giangVien: rowData[2],
    phone: rowData[3],
    email: rowData[4],
    maMH: rowData[5],
    ngayHoc: formatDate(rowData[6]),
    gioHoc: rowData[7],
    phong: rowData[8],
    loaiLab: rowData[9],
    ngayNhan: formatDate(rowData[10]),
    trangThai: rowData[11],
  };

  console.log("🧪 DIRECT CALL TEST - Row " + lastRow);
  console.log("Ticket: " + ticketData.id + " - " + ticketData.giangVien);
  console.log("");
  console.log("⚠️ LƯU Ý: Bạn phải copy function này sang BOT script");
  console.log("Và chạy nó từ BOT, không phải từ sheet nguồn!");
  console.log("");
  console.log("Copy code này sang DirectTest.gs của BOT:");
  console.log("handleNewTicketFromSource(" + JSON.stringify(ticketData) + ");");
}
