/** =========================
 * CONFIG
 * ========================= */
const CONFIG = {
  CATALOG_SHEET: "Danh mục",
  REQUEST_SHEET: "Đăng ký TB",
  N8N_WEBHOOK_URL: "https://n8n.eiu.vn/webhook-test/00bd269d-5ac0-4cb5-a4ea-e52b4e527a40",
  ALLOWED_DOMAIN: "eiu.edu.vn",
  SEND_USER_EMAIL: true,
  SEND_ADMIN_EMAIL: true,
  ADMIN_EMAIL: "",
  START_WITH_ROWS: 5,
};

let LOOKUP_CACHE = null;

const REQUEST_HEADERS_ATOZ = [
  "STT",
  "Tên thiết bị và vật tư",
  "Tên thương mại",
  "Loại",
  "Nước SX",
  "Hãng",
  "NCC",
  "ĐVT",
  "Số lượng",
  "Mã môn học",
  "Tên môn học",
  "Loại lab",
  "Phòng/Lab",
  "Ngày học",
  "Giờ học",
  "Tên người đăng ký",
  "Email",
  "Số điện thoại",
  "Ghi chú",
  "Thời gian gửi",
  "Ngày nhận",
  "Giờ nhận",
  "Tên người nhận",
  "Email người nhận",
  "Thời gian nhận",
  "Tên người trả",
  "Email người trả",
  "Thời gian trả",
  "Key",
  "ID_Dexuat",
  "SL Giao",
  "SL Trả",
  "Ngày trả",
  "Giờ trả",
  "Tên kỹ thuật/Bài thí nghiệm",
  "Giảng viên phụ trách",
  "Email Giảng viên phụ trách",
  "Học kỳ",
];

function doGet(e) {
  let userEmail = "";
  try {
    userEmail = Session.getActiveUser().getEmail();
  } catch (err) {
    // Bỏ qua lỗi nếu không lấy được email
  }

  let baseUrl = "";
  try {
    baseUrl = ScriptApp.getService().getUrl() || "";
  } catch (err) {
    baseUrl = "";
  }

  const indexTemplate = HtmlService.createTemplateFromFile("index");
  indexTemplate.userEmail = userEmail;
  indexTemplate.baseUrl = baseUrl;
  return indexTemplate
    .evaluate()
    .setTitle("Đăng ký Trang thiết bị")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function _ss() {
  return SpreadsheetApp.getActive();
}
function _sh(name) {
  return _ss().getSheetByName(name);
}
function _trim(s) {
  return s == null ? "" : ("" + s).trim();
}
function _uniq(a) {
  return [
    ...new Set(
      a.filter((v) => v !== null && v !== undefined && ("" + v).trim() !== ""),
    ),
  ];
}

function getLookups() {
  const sh = _sh(CONFIG.CATALOG_SHEET);
  if (!sh) throw new Error(`Không tìm thấy sheet "${CONFIG.CATALOG_SHEET}"`);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2)
    return {
      tenTB: [],
      tenThuongMai: [],
      loai: [],
      nuocSX: [],
      hang: [],
      ncc: [],
      dvt: [],
      loaiMonHoc: [],
      phong: [],
      loaiPLab: [],
      kyThuat: [],
    };
  const rng = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const col = (letter) =>
    ({
      A: 1,
      B: 2,
      C: 3,
      D: 4,
      E: 5,
      F: 6,
      G: 7,
      H: 8,
      I: 9,
      J: 10,
      K: 11,
      L: 12,
      M: 13,
      N: 14,
    })[letter];
  const tenTB = _uniq(rng.map((r) => _trim(r[col("B") - 1]))).sort();
  const tenThuongMai = _uniq(rng.map((r) => _trim(r[col("C") - 1]))).sort();
  const loai = _uniq(rng.map((r) => _trim(r[col("D") - 1]))).sort();
  const nuocSX = _uniq(rng.map((r) => _trim(r[col("E") - 1]))).sort();
  const hang = _uniq(rng.map((r) => _trim(r[col("F") - 1]))).sort();
  const ncc = _uniq(rng.map((r) => _trim(r[col("G") - 1]))).sort();
  const dvt = _uniq(rng.map((r) => _trim(r[col("H") - 1]))).sort();
  const loaiMonHoc = _uniq(rng.map((r) => _trim(r[col("I") - 1]))).sort();
  const phong = _uniq(rng.map((r) => _trim(r[col("M") - 1]))).sort();
  const loaiPLab = _uniq(rng.map((r) => _trim(r[col("L") - 1]))).sort();
  const kyThuat = _uniq(rng.map((r) => _trim(r[col("N") - 1]))).sort();
  return {
    tenTB,
    tenThuongMai,
    loai,
    nuocSX,
    hang,
    ncc,
    dvt,
    loaiMonHoc,
    phong,
    loaiPLab,
    kyThuat,
  };
}

function getNhanSu() {
  const sh = _sh("Nhân sự");
  if (!sh) return { giangVien: [], admins: [], nhanSuMap: {} };
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { giangVien: [], admins: [], nhanSuMap: {} };
  const rng = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  const giangVien = _uniq(rng.map((r) => _trim(r[1]))).sort();
  const admins = _uniq(
    rng.filter((r) => _trim(r[4]) === "Admin").map((r) => _trim(r[0])),
  ).filter((e) => /@/.test(e));

  const nhanSuMap = {};
  rng.forEach((r) => {
    const email = _trim(r[0]);
    if (email) {
      nhanSuMap[email] = {
        hoTen: _trim(r[1]),
        sdt: _trim(r[2]),
        chucDanh: _trim(r[3]),
        vaiTro: _trim(r[4]),
        hoatDong: _trim(r[5]),
      };
    }
  });
  return { giangVien, admins, nhanSuMap };
}

function getDanhMucData() {
  const sh = _sh(CONFIG.CATALOG_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const rng = sh.getRange(2, 1, lastRow - 1, 14).getValues();
  return rng
    .map((r) => ({
      tenTB: _trim(r[1]),
      tenThuongMai: _trim(r[2]),
      loai: _trim(r[3]),
      nuocSX: _trim(r[4]),
      hang: _trim(r[5]),
      ncc: _trim(r[6]),
      dvt: _trim(r[7]),
      loaiMonHoc: _trim(r[8]),
      maMonHoc: _trim(r[9]),
      tenMonHoc: _trim(r[10]),
      loaiLab: _trim(r[11]),
      phong: _trim(r[12]),
      kyThuat: _trim(r[13]),
    }))
    .filter((x) => x.tenTB);
}

function getUserInfoByEmail(email) {
  const nhanSu = getNhanSu();
  const userInfo = nhanSu.nhanSuMap[email] || {
    hoTen: "",
    sdt: "",
    chucDanh: "",
    vaiTro: "",
    hoatDong: "",
  };
  userInfo.email = email || "";
  return userInfo;
}

// ==========================================
// HÀM MỚI: LẤY DANH SÁCH ĐƠN VỊ VÀ KHO (CHỈ TRUE)
// ==========================================
function getDonViVaKho() {
  const ss = SpreadsheetApp.getActive();
  const shDonVi = ss.getSheetByName("Đơn vị cơ sở");
  const shKho = ss.getSheetByName("Kho");

  const donViList = [];
  if (shDonVi) {
    const data = shDonVi.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // Cột C (Trạng thái) là index 2
      if (data[i][2] === true || String(data[i][2]).toUpperCase() === "TRUE") {
        donViList.push({ code: _trim(data[i][0]), name: _trim(data[i][1]) });
      }
    }
  }

  const khoList = [];
  if (shKho) {
    const data = shKho.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // Cột D (Trạng thái) của sheet Kho là index 3
      if (data[i][3] === true || String(data[i][3]).toUpperCase() === "TRUE") {
        khoList.push({ code: _trim(data[i][0]), name: _trim(data[i][1]) });
      }
    }
  }

  return { donVi: donViList, kho: khoList };
}

function generateUniqueId() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let result = "";
  for (let i = 0; i < 10; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

function getN8nWebhookUrl_() {
  const props = PropertiesService.getScriptProperties();
  const configuredUrl = _trim(props.getProperty("N8N_WEBHOOK_URL") || "");
  if (configuredUrl) return configuredUrl;
  return CONFIG.N8N_WEBHOOK_URL;
}

function parseWebhookBody_(text) {
  if (!text) return null;
  let current = text;
  for (let i = 0; i < 3; i++) {
    if (typeof current !== "string") return current;
    const trimmed = current.trim();
    if (!trimmed) return null;
    try {
      current = JSON.parse(trimmed);
      continue;
    } catch (e) {
      return i === 0 ? null : current;
    }
  }
  try {
    return current;
  } catch (e) {
    return null;
  }
}

function toReadableMessage_(value) {
  if (value == null) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean")
    return String(value);
  if (typeof value === "object") {
    if (typeof value.message === "string") return value.message;
    try {
      return JSON.stringify(value);
    } catch (e) {
      return String(value);
    }
  }
  return String(value);
}

function extractNestedApiMessage_(messageText) {
  const raw = _trim(messageText || "");
  if (!raw) return "";

  const tryExtractFromObject = (obj) => {
    const nested =
      obj?.error?.message ||
      obj?.message ||
      obj?.error_description ||
      obj?.details ||
      "";
    return _trim(nested || "");
  };

  // Trường hợp message là JSON string trực tiếp.
  if (raw.startsWith("{") && raw.endsWith("}")) {
    try {
      const parsed = JSON.parse(raw);
      const nested = tryExtractFromObject(parsed);
      if (nested) return nested;
    } catch (e) {}
  }

  // Dạng thường gặp từ n8n: 400 - "{\"error\":{\"message\":\"...\"}}"
  // Bóc chuỗi JSON nhúng ở trong message.
  const firstBrace = raw.indexOf("{");
  const lastBrace = raw.lastIndexOf("}");
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    let embedded = raw.slice(firstBrace, lastBrace + 1);
    for (let i = 0; i < 3; i++) {
      try {
        const parsed = JSON.parse(embedded);
        const nested = tryExtractFromObject(parsed);
        if (nested) return nested;
      } catch (e) {
        embedded = embedded
          .replace(/\\"/g, '"')
          .replace(/\\\\/g, "\\")
          .replace(/^"+|"+$/g, "");
      }
    }
  }

  return raw;
}

function resolveAbpWebhookResult_(responseCode, responseText) {
  const fallbackHeaders = arguments[2] || {};
  const parsedRaw = parseWebhookBody_(responseText);
  const reparsedRaw =
    typeof parsedRaw === "string" ? parseWebhookBody_(parsedRaw) : parsedRaw;
  const parsedBody = Array.isArray(reparsedRaw)
    ? reparsedRaw[0] || null
    : reparsedRaw;
  const toBooleanOrNull = (v) => {
    if (typeof v === "boolean") return v;
    if (typeof v === "number") return v === 1 ? true : v === 0 ? false : null;
    if (typeof v === "string") {
      const t = v.trim().toLowerCase();
      if (t === "true" || t === "1") return true;
      if (t === "false" || t === "0") return false;
    }
    return null;
  };
  const successRaw =
    parsedBody && parsedBody.IsSuccess !== undefined
      ? parsedBody.IsSuccess
      : parsedBody && parsedBody.isSuccess !== undefined
        ? parsedBody.isSuccess
        : null;
  const abpSuccess = toBooleanOrNull(successRaw);
  const primaryMessage = parsedBody
    ? parsedBody.Message ||
      parsedBody.message ||
      (parsedBody.error && parsedBody.error.message) ||
      parsedBody.error ||
      parsedBody.title
    : responseText;
  const message = extractNestedApiMessage_(toReadableMessage_(primaryMessage));
  const headerMessage =
    _trim(
      fallbackHeaders["x-error-message"] ||
        fallbackHeaders["X-Error-Message"] ||
        fallbackHeaders["x-message"] ||
        fallbackHeaders["X-Message"] ||
        "",
    ) || "";
  const transactionNo =
    (parsedBody && (parsedBody.TransactionNo || parsedBody.transactionNo)) || "";
  const inventoryIssueId =
    (parsedBody &&
      (parsedBody.InventoryIssueId || parsedBody.inventoryIssueId)) ||
    "";

  const okByHttp = responseCode < 400;
  // Nếu API đã trả cờ thành công rõ ràng thì ưu tiên cờ đó.
  const ok = abpSuccess === null ? okByHttp : abpSuccess;

  return {
    ok,
    statusCode: responseCode,
    message: message || headerMessage,
    rawBody: responseText,
    data: parsedBody,
    abp: {
      isSuccess: abpSuccess,
      transactionNo: transactionNo,
      inventoryIssueId: inventoryIssueId,
    },
  };
}

function formatDateToken_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "ddMMyy");
}

function resolveTransactionPrefix_(transactionType, source) {
  const basePrefix = transactionType === "XUAT_NHAP_KHO" ? "XK" : "XKDC";
  // Đánh dấu nguồn để phân biệt phiếu đi từ Google Script.
  return source === "GOOGLE_SCRIPT" ? `${basePrefix}GS` : basePrefix;
}

function generateTransactionNo_(sheet, options) {
  const transactionDate = parseAnyDate(options.transactionDate) || new Date();
  const transactionType = _trim(options.transactionType) || "XUAT_NHAP_KHO";
  const source = _trim(options.source) || "GOOGLE_SCRIPT";
  const providedId = _trim(options.providedId);

  const dateToken = formatDateToken_(transactionDate);
  const prefix = resolveTransactionPrefix_(transactionType, source);
  const expectedPrefix = `${prefix}_${dateToken}_`;

  const values = sheet.getDataRange().getValues();
  const headers = values[0] || [];
  const idCol = headers.indexOf("ID_Dexuat");
  const dateCol = headers.indexOf("Ngày học");

  if (idCol < 0 || dateCol < 0) {
    throw new Error('Không tìm thấy cột "ID_Dexuat" hoặc "Ngày học" trong sheet.');
  }

  // Nếu có truyền id và đã tồn tại cùng ngày + cùng loại thì dùng lại mã hiện có.
  if (providedId) {
    for (let i = 1; i < values.length; i++) {
      const rowId = _trim(values[i][idCol]);
      if (rowId !== providedId) continue;

      const rowDate = parseAnyDate(values[i][dateCol]);
      if (!rowDate) continue;

      const sameDate = formatDateToken_(rowDate) === dateToken;
      const sameType = rowId.indexOf(expectedPrefix) === 0;
      if (sameDate && sameType) return rowId;
    }
  }

  let maxSeq = 0;
  for (let i = 1; i < values.length; i++) {
    const rowId = _trim(values[i][idCol]);
    if (!rowId || rowId.indexOf(expectedPrefix) !== 0) continue;

    const seqPart = rowId.slice(expectedPrefix.length);
    if (!/^\d{3}$/.test(seqPart)) continue;
    const seq = Number(seqPart);
    if (seq > maxSeq) maxSeq = seq;
  }

  const nextSeq = String(maxSeq + 1).padStart(3, "0");
  return `${prefix}_${dateToken}_${nextSeq}`;
}

function getLookupsForClient() {
  if (!LOOKUP_CACHE) LOOKUP_CACHE = getLookups();
  const {
    tenTB,
    tenThuongMai,
    loai,
    nuocSX,
    hang,
    ncc,
    dvt,
    loaiMonHoc,
    phong,
    loaiPLab,
    kyThuat,
  } = LOOKUP_CACHE;
  const nhanSu = getNhanSu();
  const danhMucData = getDanhMucData();
  var userEmail = "";
  try {
    userEmail = getUserEmail() || "";
  } catch (err) {
    userEmail = "";
  }
  const userInfo = getUserInfoByEmail(userEmail);

  // --- BỔ SUNG LẤY ĐƠN VỊ VÀ KHO ---
  const extraData = getDonViVaKho();

  return {
    tenTB,
    tenThuongMai,
    loai,
    nuocSX,
    hang,
    ncc,
    dvt,
    loaiMonHoc,
    phong,
    loaiPLab,
    kyThuat,
    giangVien: nhanSu.giangVien,
    danhMucData: danhMucData,
    userInfo: userInfo,
    donViCoSo: extraData.donVi, // Trả mảng {code, name} Đơn vị
    kho: extraData.kho, // Trả mảng {code, name} Kho
  };
}

function getRegistrationById(idDexuat) {
  if (!idDexuat || typeof idDexuat !== "string") {
    throw new Error("ID phiếu đăng ký không hợp lệ");
  }
  const sh = _sh(CONFIG.REQUEST_SHEET);
  if (!sh) throw new Error(`Không tìm thấy sheet "${CONFIG.REQUEST_SHEET}"`);

  const allRows = sh.getDataRange().getValues();
  const headers = allRows[0];
  const idDexuatCol = headers.indexOf("ID_Dexuat");

  if (idDexuatCol < 0) {
    throw new Error("Không tìm thấy cột ID_Dexuat trong sheet");
  }

  const matchingRows = [];
  for (let i = 1; i < allRows.length; i++) {
    if (_trim(String(allRows[i][idDexuatCol])) === _trim(idDexuat)) {
      matchingRows.push(allRows[i]);
    }
  }

  if (matchingRows.length === 0) {
    throw new Error(`Không tìm thấy phiếu đăng ký với ID: ${idDexuat}`);
  }

  const firstRow = matchingRows[0];
  const meta = {
    idDexuat: _trim(firstRow[headers.indexOf("ID_Dexuat")]),
    ngayHoc: firstRow[headers.indexOf("Ngày học")],
    gioHoc: _trim(firstRow[headers.indexOf("Giờ học")]),
    maHocPhan: _trim(firstRow[headers.indexOf("Mã môn học")]),
    hocPhan: _trim(firstRow[headers.indexOf("Tên môn học")]),
    loaiLab: _trim(firstRow[headers.indexOf("Loại lab")]),
    phongLab: _trim(firstRow[headers.indexOf("Phòng/Lab")]),
    giangVien: _trim(firstRow[headers.indexOf("Tên người đăng ký")]),
    email: _trim(firstRow[headers.indexOf("Email")]),
    soDienThoai: _trim(firstRow[headers.indexOf("Số điện thoại")]),
    ngayNhan: firstRow[headers.indexOf("Ngày nhận")],
    gioNhan: _trim(firstRow[headers.indexOf("Giờ nhận")]),
    ngayTra: firstRow[headers.indexOf("Ngày trả")],
    gioTra: _trim(firstRow[headers.indexOf("Giờ trả")]),
    giangVienPhuTrach: _trim(firstRow[headers.indexOf("Giảng viên phụ trách")]),
    emailGiangVienPhuTrach: _trim(
      firstRow[headers.indexOf("Email Giảng viên phụ trách")],
    ),
    hocKy: _trim(firstRow[headers.indexOf("Học kỳ")]),
  };

  if (meta.ngayHoc instanceof Date)
    meta.ngayHoc = Utilities.formatDate(
      meta.ngayHoc,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
  if (meta.ngayNhan instanceof Date)
    meta.ngayNhan = Utilities.formatDate(
      meta.ngayNhan,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
  if (meta.ngayTra instanceof Date)
    meta.ngayTra = Utilities.formatDate(
      meta.ngayTra,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );

  if (typeof firstRow[headers.indexOf("Giờ nhận")] === "number") {
    const timeSerial = firstRow[headers.indexOf("Giờ nhận")];
    const totalSeconds = Math.round(timeSerial * 24 * 60 * 60);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    meta.gioNhan = `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  }
  if (typeof firstRow[headers.indexOf("Giờ trả")] === "number") {
    const timeSerial = firstRow[headers.indexOf("Giờ trả")];
    const totalSeconds = Math.round(timeSerial * 24 * 60 * 60);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    meta.gioTra = `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  }

  const groupedByKyThuat = {};
  matchingRows.forEach((row) => {
    const kyThuat =
      _trim(row[headers.indexOf("Tên kỹ thuật/Bài thí nghiệm")]) ||
      "Không có tên kỹ thuật";
    if (!groupedByKyThuat[kyThuat]) groupedByKyThuat[kyThuat] = [];
    groupedByKyThuat[kyThuat].push({
      tenThietBi: _trim(row[headers.indexOf("Tên thiết bị và vật tư")]),
      tenThuongMai: _trim(row[headers.indexOf("Tên thương mại")]),
      loai: _trim(row[headers.indexOf("Loại")]),
      nuocSX: _trim(row[headers.indexOf("Nước SX")]),
      hang: _trim(row[headers.indexOf("Hãng")]),
      ncc: _trim(row[headers.indexOf("NCC")]),
      dvt: _trim(row[headers.indexOf("ĐVT")]),
      soLuong: Number(row[headers.indexOf("Số lượng")]),
      ghiChu: _trim(row[headers.indexOf("Ghi chú")]),
    });
  });

  const kyNangData = Object.keys(groupedByKyThuat).map((kyThuat) => ({
    kyThuat: kyThuat,
    rows: groupedByKyThuat[kyThuat].map((item) => ({
      tenTB: item.tenThietBi,
      tenThuongMai: item.tenThuongMai,
      loai: item.loai,
      nuoc: item.nuocSX,
      hang: item.hang,
      model: item.ncc,
      dvt: item.dvt,
      sl: item.soLuong,
      ghichu: item.ghiChu,
    })),
  }));

  return { meta: meta, kyNangData: kyNangData, soKyNang: kyNangData.length };
}

function getAdmins() {
  const nhanSu = getNhanSu();
  const list = Array.from(
    new Set([
      ...(nhanSu.admins || []),
      ...(CONFIG.ADMIN_EMAIL ? [CONFIG.ADMIN_EMAIL] : []),
    ]),
  );
  return { admins: list };
}

function getUserInfo() {
  return { email: Session.getActiveUser().getEmail() || "" };
}
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

function getCourses() {
  const sh = _sh(CONFIG.CATALOG_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const vals = sh.getRange(2, 10, lastRow - 1, 2).getValues();
  const courses = vals
    .map((r) => ({ code: _trim(r[0]), name: _trim(r[1]) }))
    .filter((x) => x.code && x.name);
  const uniqueCourses = [];
  const seen = new Set();
  courses.forEach((c) => {
    if (!seen.has(c.code)) {
      seen.add(c.code);
      uniqueCourses.push(c);
    }
  });
  return uniqueCourses;
}

function parseAnyDate(s) {
  if (!s) return null;
  if (s instanceof Date) return s;
  const t = String(s).trim();
  let m;
  if ((m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/))) {
    const year = +m[1];
    const month = +m[2] - 1;
    const day = +m[3];
    return new Date(year, month, day);
  }
  if ((m = t.match(/^(\d{2})\/(\d{2})\/(\d{4})$/))) {
    const day = +m[1];
    const month = +m[2] - 1;
    const year = +m[3];
    return new Date(year, month, day);
  }
  const dt = new Date(t);
  return isNaN(dt) ? null : dt;
}

function parseTimeToDate(timeStr) {
  if (timeStr === null || timeStr === undefined || timeStr === "") return "";
  const t = String(timeStr).trim();
  if (!t || t === "") return "";
  const m = t.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return "";
  const hour = parseInt(m[1], 10);
  const minute = parseInt(m[2], 10);
  const second = parseInt(m[3] || "0", 10);
  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return "";
  const timeSerial = (hour + minute / 60 + second / 3600) / 24;
  return timeSerial;
}

function fmtVN(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function getHeaderIndexMap(sh) {
  const header = sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(_trim);
  const map = {};
  header.forEach((name, i) => {
    if (name) map[name] = i;
  });
  return { map, headerLen: header.length };
}

function ensureRequestSheetHeaders_(sh) {
  const width = Math.max(sh.getLastColumn(), REQUEST_HEADERS_ATOZ.length);
  const currentHeader = sh
    .getRange(1, 1, 1, width)
    .getValues()[0]
    .map((v) => _trim(v || ""));
  const needRewrite = REQUEST_HEADERS_ATOZ.some(
    (h, i) => (currentHeader[i] || "") !== h,
  );
  if (needRewrite)
    sh.getRange(1, 1, 1, REQUEST_HEADERS_ATOZ.length).setValues([
      REQUEST_HEADERS_ATOZ,
    ]);
  sh.getRange("N:N").setNumberFormat("dd/MM/yyyy");
  sh.getRange("O:O").setNumberFormat("@");
  sh.getRange("T:T").setNumberFormat("dd/MM/yyyy HH:mm:ss");
  sh.getRange("U:U").setNumberFormat("dd/MM/yyyy");
  sh.getRange("V:V").setNumberFormat("HH:mm");
  sh.getRange("Y:Y").setNumberFormat("dd/MM/yyyy HH:mm:ss");
  sh.getRange("AB:AB").setNumberFormat("dd/MM/yyyy HH:mm:ss");
  sh.getRange("AG:AG").setNumberFormat("dd/MM/yyyy");
  sh.getRange("AH:AH").setNumberFormat("HH:mm");
}

function buildRequestRow_(stt, meta, it, now, submitUniqueId) {
  const ngayHoc = parseAnyDate(meta.ngayHoc);
  const ngayNhan = parseAnyDate(meta.ngayNhan);
  const gioHoc = _trim(meta.gioHoc);
  const gioNhan = parseTimeToDate(meta.gioNhan);
  const gioTra = parseTimeToDate(meta.gioTra);
  const ngayTra = parseAnyDate(meta.ngayTra);
  const rowUniqueId = generateUniqueId();
  const idDexuat = _trim(meta.idDexuat);
  return [
    stt,
    _trim(it.tenThietBi),
    _trim(it.tenThuongMai || ""),
    _trim(it.loai || ""),
    _trim(it.nuocSX || ""),
    _trim(it.hang || ""),
    _trim(it.ncc || ""),
    _trim(it.dvt || ""),
    Number(it.soLuong),
    _trim(meta.maHocPhan),
    _trim(meta.hocPhan),
    _trim(meta.loaiLab || ""),
    _trim(meta.phongLab),
    ngayHoc,
    gioHoc,
    _trim(meta.giangVien),
    _trim(meta.email),
    _trim(meta.soDienThoai || ""),
    _trim(it.ghiChu || ""),
    now,
    ngayNhan,
    gioNhan,
    "",
    "",
    "",
    "",
    "",
    "",
    rowUniqueId,
    idDexuat,
    Number(it.soLuong),
    "",
    ngayTra,
    gioTra,
    _trim(it.kyThuat || ""),
    _trim(meta.giangVienPhuTrach || ""),
    _trim(meta.emailGiangVienPhuTrach || ""),
    _trim(meta.hocKy || ""),
  ];
}

// Hàm chính xử lý lưu dữ liệu và bắn webhook N8n
function submitRegistration(payload) {
  if (!payload || !payload.meta || !Array.isArray(payload.items))
    throw new Error("Dữ liệu gửi lên không hợp lệ.");

  const user = getUserInfo();
  if (!user.email || !user.email.endsWith("@" + CONFIG.ALLOWED_DOMAIN))
    throw new Error("Bạn không có quyền gửi biểu mẫu này.");

  const sh = _sh(CONFIG.REQUEST_SHEET);
  if (!sh) throw new Error(`Không tìm thấy sheet "${CONFIG.REQUEST_SHEET}"`);
  ensureRequestSheetHeaders_(sh);

  const meta = payload.meta;
  const now = new Date();
  if (!meta.ngayHoc) throw new Error('Thiếu "Ngày học".');
  if (!meta.gioHoc) throw new Error('Thiếu "Giờ học".');

  [
    "maHocPhan",
    "hocPhan",
    "phongLab",
    "email",
    "soDienThoai",
    "giangVienPhuTrach",
    "emailGiangVienPhuTrach",
    "hocKy",
  ].forEach((k) => {
    if (!meta[k] || !String(meta[k]).trim())
      throw new Error(`Thiếu thông tin bắt buộc: ${k}`);
  });

  let settings = {};
  try {
    settings = getAdminSettings();
  } catch (e) {
    settings = {};
  }

  const requiredFields = String(settings.requiredFields || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
  const reqSet = new Set(requiredFields);

  if (reqSet.has("ngayNhan") && !meta.ngayNhan)
    throw new Error('Thiếu "Ngày nhận" (được cấu hình là bắt buộc).');
  if (reqSet.has("gioNhan") && !meta.gioNhan)
    throw new Error('Thiếu "Giờ nhận" (được cấu hình là bắt buộc).');

  if (meta.ngayNhan && meta.ngayHoc) {
    var dNhan = parseAnyDate(meta.ngayNhan);
    var dHoc = parseAnyDate(meta.ngayHoc);
    if (dNhan && dHoc && dNhan > dHoc)
      throw new Error('"Ngày nhận" phải bằng hoặc trước "Ngày học".');
  }

  if (meta.gioHoc) {
    var m = String(meta.gioHoc).match(
      /^(\s*)(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})(\s*)$/,
    );
    if (m) {
      var sHour = Number(m[2]),
        sMin = Number(m[3]);
      var eHour = Number(m[4]),
        eMin = Number(m[5]);
      var startMinutes = sHour * 60 + sMin;
      var endMinutes = eHour * 60 + eMin;
      if (endMinutes < startMinutes)
        throw new Error('"Giờ kết thúc" phải bằng hoặc sau "Giờ bắt đầu".');
    }
  }

  const transactionType = _trim(meta.transactionType) || "XUAT_NHAP_KHO";
  const sourceChannel = "GOOGLE_SCRIPT";
  let submitUniqueId;
  let rowsToDelete = [];
  if (payload.oldIdDexuat && typeof payload.oldIdDexuat === "string") {
    submitUniqueId = payload.oldIdDexuat;
    meta.idDexuat = generateTransactionNo_(sh, {
      providedId: payload.oldIdDexuat,
      transactionDate: meta.ngayHoc,
      transactionType: transactionType,
      source: sourceChannel,
    });

    const allRows = sh.getDataRange().getValues();
    const idDexuatCol = REQUEST_HEADERS_ATOZ.indexOf("ID_Dexuat");
    if (idDexuatCol >= 0) {
      for (let i = 1; i < allRows.length; i++) {
        const cellValue = _trim(String(allRows[i][idDexuatCol] || ""));
        const searchValue = _trim(String(payload.oldIdDexuat));
        if (cellValue === searchValue) rowsToDelete.push(i + 1);
      }
    }
  } else {
    submitUniqueId = generateUniqueId();
    meta.idDexuat = generateTransactionNo_(sh, {
      transactionDate: meta.ngayHoc,
      transactionType: transactionType,
      source: sourceChannel,
    });
  }

  const lastRow = sh.getLastRow();
  let nextSTT = 0;
  if (lastRow >= 2) {
    const lastA = Number(sh.getRange(lastRow, 1).getValue());
    nextSTT = Number.isFinite(lastA) ? lastA : lastRow - 1;
  }

  const rows = [];
  payload.items.forEach((it, i) => {
    if (!it.tenThietBi)
      throw new Error(`Dòng thiết bị #${i + 1} thiếu "Tên thiết bị".`);
    if (!it.kyThuat)
      throw new Error(
        `Dòng thiết bị #${i + 1} thiếu "Tên kỹ thuật/Bài thí nghiệm".`,
      );
    const soLuong = Number(it.soLuong);
    if (!(soLuong > 0))
      throw new Error(`Dòng thiết bị #${i + 1} "Số lượng" phải > 0.`);
    const stt = ++nextSTT;
    rows.push(buildRequestRow_(stt, meta, it, now, submitUniqueId));
  });

  // ==========================================
  // --- TÍCH HỢP BẮN WEBHOOK SANG N8N ---
  // ==========================================
  let webhookResult = {
    ok: false,
    statusCode: 0,
    message: "",
    rawBody: "",
  };

  try {
    const n8nWebhookUrl = getN8nWebhookUrl_();

    // Cấu trúc Data JSON đẩy sang n8n
    const n8nPayload = {
      event_type: payload.oldIdDexuat
        ? "update_registration"
        : "new_registration",
      timestamp: new Date().toISOString(),
      idDexuat: meta.idDexuat,
      transactionNo: meta.idDexuat,
      transactionType: transactionType,
      sourceChannel: sourceChannel,
      meta: meta, // Dữ liệu Đơn vị và Kho đã nằm trong `meta` nếu bạn gửi từ `index.html`
      items: payload.items,
      source: "EIU_Equipment_System",
      user: user.email,
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(n8nPayload),
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(n8nWebhookUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    const responseHeaders = response.getAllHeaders() || {};

    webhookResult = resolveAbpWebhookResult_(
      responseCode,
      responseText,
      responseHeaders,
    );
    
    if (!webhookResult.ok) {
      throw new Error(
        webhookResult.message || `Đã xảy ra lỗi trong quá trình đồng bộ dữ liệu lên hệ thống [eiu-asset.eiu.vn] ${responseCode}.`,
      );
    }
  } catch (err) {
    const errMsg = err && err.message ? err.message : String(err);
    const detailFromBody = extractNestedApiMessage_(webhookResult.rawBody || "");
    const resolvedMessage =
      webhookResult.message || detailFromBody || errMsg || "";
    webhookResult = {
      ok: false,
      statusCode: webhookResult.statusCode || 0,
      message: resolvedMessage,
      rawBody: webhookResult.rawBody || "",
      data: webhookResult.data || null,
    };
    Logger.log("❌ Lỗi khi gửi tới n8n: " + errMsg);
  }
  // ==========================================

  // Chỉ lưu phiếu và gửi thông báo khi đồng bộ ABP thành công.
  if (!webhookResult.ok) {
    const finalErrorMessage =
      webhookResult.message ||
      extractNestedApiMessage_(webhookResult.rawBody || "") ||
      "Đồng bộ API ABP thất bại, hệ thống không lưu phiếu.";
    throw new Error(
      finalErrorMessage,
    );
  }

  if (rowsToDelete.length > 0) {
    rowsToDelete
      .sort((a, b) => b - a)
      .forEach((rowNum) => sh.deleteRow(rowNum));
  }

  if (rows.length)
    sh.getRange(
      sh.getLastRow() + 1,
      1,
      rows.length,
      REQUEST_HEADERS_ATOZ.length,
    ).setValues(rows);

  if (!payload.oldIdDexuat) {
    try {
      const shPhieuNhap = _sh("Phiếu nhập");
      if (shPhieuNhap) {
        const lastRowPN = shPhieuNhap.getLastRow();
        shPhieuNhap.getRange(lastRowPN + 1, 1).setValue(meta.idDexuat);
        shPhieuNhap.getRange(lastRowPN + 1, 12).setValue("Mới");
      }
    } catch (e) {
      Logger.log("Lỗi ghi Phiếu nhập: " + e);
    }
  }

  if (CONFIG.SEND_USER_EMAIL && meta.email) {
    try {
      const ngayHocFormatted = meta.ngayHoc
        ? fmtVN(parseAnyDate(meta.ngayHoc))
        : "";
      const isEdit =
        payload.oldIdDexuat && typeof payload.oldIdDexuat === "string";
      const userSubject = isEdit
        ? `Điều chỉnh phiếu đăng ký thiết bị của ${meta.giangVien} - ${ngayHocFormatted} - ${meta.maHocPhan} - ${meta.idDexuat}`
        : `Xác nhận đăng ký trang thiết bị của ${meta.giangVien} - ${ngayHocFormatted} - ${meta.maHocPhan} - ${meta.idDexuat}`;
      GmailApp.sendEmail(
        meta.email,
        userSubject,
        buildEmailText(meta, payload.items),
        {
          name: "Đăng ký trang thiết bị",
          htmlBody: buildEmailHtml(meta, payload.items),
        },
      );
    } catch (e) {}
  }

  if (CONFIG.SEND_ADMIN_EMAIL) {
    const nhanSu = getNhanSu();
    const listAdmins = Array.from(
      new Set([
        ...(nhanSu.admins || []),
        ...(CONFIG.ADMIN_EMAIL ? [CONFIG.ADMIN_EMAIL] : []),
      ]),
    );
    listAdmins.forEach((adminEmail) => {
      if (!/@/.test(adminEmail)) return;
      try {
        const ngayHocFormatted = meta.ngayHoc
          ? fmtVN(parseAnyDate(meta.ngayHoc))
          : "";
        const isEdit =
          payload.oldIdDexuat && typeof payload.oldIdDexuat === "string";
        const adminSubject = isEdit
          ? `[ADMIN] Điều chỉnh phiếu đăng ký thiết bị của ${meta.giangVien} - ${ngayHocFormatted} - ${meta.maHocPhan} - ${meta.idDexuat}`
          : `[ADMIN] Có đăng ký trang thiết bị mới - ${meta.giangVien} - ${ngayHocFormatted} - ${meta.maHocPhan} - ${meta.idDexuat}`;
        GmailApp.sendEmail(
          adminEmail,
          adminSubject,
          buildEmailText(meta, payload.items),
          {
            name: "Đăng ký trang thiết bị",
            htmlBody: buildEmailHtml(meta, payload.items),
          },
        );
      } catch (e) {
        Logger.log("Lỗi gửi email admin: " + e);
      }
    });
  }

  return {
    ok: true,
    count: rows.length,
    idDexuat: meta.idDexuat,
    transactionNo:
      (webhookResult.abp && webhookResult.abp.transactionNo) || meta.idDexuat,
    webhook: webhookResult,
  };
}

function buildEmailText(meta, items) {
  const lines = [];
  lines.push("Bạn đã gửi phiếu đăng ký trang thiết bị thành công.\n");
  if (meta.idDexuat)
    lines.push(`📋 ID PHIẾU ĐĂNG KÝ: ${meta.maHocPhan} - ${meta.idDexuat}\n`);
  lines.push("=== 1. THÔNG TIN MÔN HỌC ===");
  if (meta.ngayHoc)
    lines.push(`Ngày học: ${fmtVN(parseAnyDate(meta.ngayHoc))}`);
  if (meta.gioHoc) lines.push(`Giờ học: ${meta.gioHoc}`);
  if (meta.hocKy) lines.push(`Học kỳ: ${meta.hocKy}`);
  if (meta.maHocPhan) lines.push(`Mã môn học: ${meta.maHocPhan}`);
  if (meta.hocPhan) lines.push(`Tên môn học: ${meta.hocPhan}`);
  if (meta.loaiLab) lines.push(`Loại lab: ${meta.loaiLab}`);
  if (meta.phongLab) lines.push(`Phòng/Lab: ${meta.phongLab}\n`);

  lines.push("=== 2. THÔNG TIN NGƯỜI ĐĂNG KÝ ===");
  if (meta.giangVien) lines.push(`Người đăng ký: ${meta.giangVien}`);
  if (meta.email) lines.push(`Email: ${meta.email}`);
  if (meta.soDienThoai) lines.push(`Số điện thoại: ${meta.soDienThoai}\n`);

  lines.push("=== 3. THÔNG TIN GIẢNG VIÊN PHỤ TRÁCH ===");
  if (meta.giangVienPhuTrach)
    lines.push(`Giảng viên phụ trách: ${meta.giangVienPhuTrach}`);
  if (meta.emailGiangVienPhuTrach)
    lines.push(`Email Giảng viên phụ trách: ${meta.emailGiangVienPhuTrach}\n`);

  lines.push("=== 4. THÔNG TIN NHẬN THIẾT BỊ ===");
  if (meta.ngayNhan)
    lines.push(`Ngày nhận: ${fmtVN(parseAnyDate(meta.ngayNhan))}`);
  if (meta.gioNhan) lines.push(`Giờ nhận: ${meta.gioNhan}`);
  if (meta.ngayTra)
    lines.push(`Ngày trả: ${fmtVN(parseAnyDate(meta.ngayTra))}`);
  if (meta.gioTra) lines.push(`Giờ trả: ${meta.gioTra}\n`);

  lines.push("=== DANH SÁCH CHI TIẾT THIẾT BỊ ===");
  const groupedByKyThuat = {};
  items.forEach((it) => {
    const key = it.kyThuat || "Không có tên kỹ thuật";
    if (!groupedByKyThuat[key]) groupedByKyThuat[key] = [];
    groupedByKyThuat[key].push(it);
  });

  Object.keys(groupedByKyThuat).forEach((kyThuat) => {
    lines.push(`\n▪ ${kyThuat}`);
    groupedByKyThuat[kyThuat].forEach((it, idx) => {
      lines.push(`  ${idx + 1}. ${it.tenThietBi}`);
      lines.push(`     - Tên thương mại: ${it.tenThuongMai || "N/A"}`);
      lines.push(`     - ĐVT: ${it.dvt || "N/A"}`);
      lines.push(`     - Số lượng: ${it.soLuong}`);
      if (it.ghiChu) lines.push(`     - Ghi chú: ${it.ghiChu}`);
    });
  });

  lines.push("\nTrân trọng,");
  lines.push("Khoa Điều Dưỡng - EIU");
  return lines.join("\n");
}

function buildEmailHtml(meta, items) {
  const esc = (s) =>
    String(s == null ? "" : s).replace(/[&<>"']/g, function (m) {
      return m === "&"
        ? "&amp;"
        : m === "<"
          ? "&lt;"
          : m === ">"
            ? "&gt;"
            : m === '"'
              ? "&quot;"
              : "&#39;";
    });

  const groupedByKyThuat = {};
  items.forEach((it) => {
    const key = it.kyThuat || "Không có tên kỹ thuật";
    if (!groupedByKyThuat[key]) groupedByKyThuat[key] = [];
    groupedByKyThuat[key].push(it);
  });

  let kyThuatTablesHtml = "";
  Object.keys(groupedByKyThuat).forEach((kyThuat) => {
    const kyThuatItems = groupedByKyThuat[kyThuat];
    const rows = kyThuatItems
      .map(
        (it, i) => `<tr>
      <td style="border:1px solid #e5e7eb;padding:6px;text-align:center;">${i + 1}</td>
      <td style="border:1px solid #e5e7eb;padding:6px;">${esc(it.tenThietBi)}</td>
      <td style="border:1px solid #e5e7eb;padding:6px;">${esc(it.tenThuongMai)}</td>
      <td style="border:1px solid #e5e7eb;padding:6px;text-align:center;">${esc(it.dvt)}</td>
      <td style="border:1px solid #e5e7eb;padding:6px;text-align:center;">${esc(it.soLuong)}</td>
      <td style="border:1px solid #e5e7eb;padding:6px;">${esc(it.ghiChu || "")}</td>
    </tr>`,
      )
      .join("");

    kyThuatTablesHtml += `<div style="margin-bottom:20px;">
        <h4 style="margin:12px 0 8px;font-size:14px;color:#2563eb;font-weight:bold;">▪ ${esc(kyThuat)}</h4>
        <table style="border-collapse:collapse;width:100%;border:1px solid #e5e7eb;">
          <thead>
            <tr>
              <th style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;width:5%;text-align:center;">#</th>
              <th style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;width:25%;">Tên thiết bị</th>
              <th style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;width:35%;">Tên thương mại</th>
              <th style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;width:10%;text-align:center;">ĐVT</th>
              <th style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;width:10%;text-align:center;">SL</th>
              <th style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;width:13%;">Ghi chú</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
  });

  return `<div style="font-family:system-ui,Arial;line-height:1.45;color:#111;">
    ${
      meta.idDexuat
        ? `<div style="text-align:center;margin:16px 0;padding:12px;background:#f0f9ff;border:1px solid #bfdbfe;border-radius:8px;">
             <strong style="font-size:16px;color:#059669;">ID PHIẾU ĐĂNG KÝ: 
               <span style="color:#2563eb;">${esc(meta.maHocPhan)}</span> - 
               <span style="color:#dc2626;">${esc(meta.idDexuat)}</span>
             </strong>
           </div>`
        : ""
    }
    
    <h3 style="margin:16px 0 8px;font-size:16px;color:#2563eb;border-bottom:2px solid #2563eb;padding-bottom:4px;">1. Thông tin môn học</h3>
    <table style="border-collapse:collapse;border:1px solid #e5e7eb;width:100%;margin-bottom:16px;">
      <tbody>
        ${meta.ngayHoc ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;width:30%;">Ngày học</td><td style="border:1px solid #e5e7eb;padding:6px;">${fmtVN(parseAnyDate(meta.ngayHoc))}</td></tr>` : ""}
        ${meta.gioHoc ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;">Giờ học</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.gioHoc)}</td></tr>` : ""}
        ${meta.hocKy ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;">Học kỳ</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.hocKy)}</td></tr>` : ""}
        ${meta.maHocPhan ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;">Mã môn học</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.maHocPhan)}</td></tr>` : ""}
        ${meta.hocPhan ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;">Tên môn học</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.hocPhan)}</td></tr>` : ""}
        ${meta.loaiLab ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;">Loại lab</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.loaiLab)}</td></tr>` : ""}
        ${meta.phongLab ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;">Phòng/Lab</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.phongLab)}</td></tr>` : ""}
      </tbody>
    </table>
    
    <h3 style="margin:16px 0 8px;font-size:16px;color:#2563eb;border-bottom:2px solid #2563eb;padding-bottom:4px;">2. Thông tin người đăng ký</h3>
    <table style="border-collapse:collapse;border:1px solid #e5e7eb;width:100%;margin-bottom:16px;">
      <tbody>
        ${meta.giangVien ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;width:30%;">Người đăng ký</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.giangVien)}</td></tr>` : ""}
        ${meta.email ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;">Email</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.email)}</td></tr>` : ""}
        ${meta.soDienThoai ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;">SĐT</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.soDienThoai)}</td></tr>` : ""}
      </tbody>
    </table>
    
    <h3 style="margin:16px 0 8px;font-size:16px;color:#2563eb;border-bottom:2px solid #2563eb;padding-bottom:4px;">3. Thông tin giảng viên phụ trách</h3>
    <table style="border-collapse:collapse;border:1px solid #e5e7eb;width:100%;margin-bottom:16px;">
      <tbody>
        ${meta.giangVienPhuTrach ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;width:30%;">Giảng viên phụ trách</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.giangVienPhuTrach)}</td></tr>` : ""}
        ${meta.emailGiangVienPhuTrach ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;">Email Giảng viên phụ trách</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.emailGiangVienPhuTrach)}</td></tr>` : ""}
      </tbody>
    </table>
    
    <h3 style="margin:16px 0 8px;font-size:16px;color:#2563eb;border-bottom:2px solid #2563eb;padding-bottom:4px;">4. Thông tin nhận thiết bị</h3>
    <table style="border-collapse:collapse;border:1px solid #e5e7eb;width:100%;margin-bottom:16px;">
      <tbody>
        ${meta.ngayNhan ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;width:30%;">Ngày nhận</td><td style="border:1px solid #e5e7eb;padding:6px;">${fmtVN(parseAnyDate(meta.ngayNhan))}</td></tr>` : ""}
        ${meta.gioNhan ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;">Giờ nhận</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.gioNhan)}</td></tr>` : ""}
        ${meta.ngayTra ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;">Ngày trả</td><td style="border:1px solid #e5e7eb;padding:6px;">${fmtVN(parseAnyDate(meta.ngayTra))}</td></tr>` : ""}
        ${meta.gioTra ? `<tr><td style="border:1px solid #e5e7eb;padding:6px;background:#f8fafc;">Giờ trả</td><td style="border:1px solid #e5e7eb;padding:6px;">${esc(meta.gioTra)}</td></tr>` : ""}
      </tbody>
    </table>
    
    <h3 style="margin:16px 0 12px;font-size:16px;color:#2563eb;border-bottom:2px solid #2563eb;padding-bottom:4px;">Danh sách chi tiết thiết bị</h3>
    ${kyThuatTablesHtml}
  </div>`;
}

function saveAdminSettings(settings) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties(settings);
  return { success: true, message: "Đã lưu cài đặt thành công" };
}

function getAdminSettings() {
  const props = PropertiesService.getScriptProperties();
  const settings = props.getProperties();
  const defaults = {
    formTitle: "PHIẾU ĐĂNG KÝ TRANG THIẾT BỊ THỰC HÀNH - KHOA ĐIỀU DƯỠNG",
    showLogo: "true",
    logoUrl: "https://drive.google.com/uc?id=1-hzToSdzLBTUQoeP0iwGOyCAZlMEGdcZ",
    defaultStartRows: "3",
    maxRows: "50",
    requiredFields: "ngayHoc,gioHoc,maHocPhan,phongLab,soDienThoai",
    hiddenFields: "",
    readonlyFields: "email,giangVien,hocPhan,loaiLab",
    emailSubject: "Xác nhận đăng ký trang thiết bị",
    emailSenderName: "Khoa Điều Dưỡng - EIU",
    sendUserEmail: "true",
    sendAdminEmail: "true",
    workingHoursStart: "07:00",
    workingHoursEnd: "17:00",
    timeInterval: "30",
    allowPastDates: "false",
    minDaysAdvance: "1",
    popupMessage:
      "Bạn đã đăng ký thành công. Vui lòng kiểm tra email xác nhận.",
    showSuccessPopup: "true",
    redirectAfterSubmit: "false",
    redirectUrl: "",
    catalogSheet: "Danh mục",
    requestSheet: "Đăng ký TB",
    staffSheet: "Nhân sự",
    allowedDomain: "eiu.edu.vn",
    autoFillUserInfo: "true",
    requireLogin: "true",
    sessionTimeout: "3600",
    maxLoginAttempts: "5",
    lockoutDuration: "900",
  };

  Object.keys(defaults).forEach((key) => {
    if (!settings[key]) settings[key] = defaults[key];
  });
  return settings;
}

function getSheetNames() {
  const ss = _ss();
  const sheets = ss.getSheets();
  return sheets.map((s) => s.getName());
}

function getRegistrationStats() {
  const sh = _sh(CONFIG.REQUEST_SHEET);
  if (!sh) return { total: 0, today: 0, thisWeek: 0, thisMonth: 0 };
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { total: 0, today: 0, thisWeek: 0, thisMonth: 0 };
  const data = sh.getRange(2, 20, lastRow - 1, 1).getValues();
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

  let todayCount = 0,
    weekCount = 0,
    monthCount = 0;
  data.forEach((row) => {
    const date = row[0];
    if (date instanceof Date) {
      if (date >= today) todayCount++;
      if (date >= weekAgo) weekCount++;
      if (date >= monthStart) monthCount++;
    }
  });
  return {
    total: lastRow - 1,
    today: todayCount,
    thisWeek: weekCount,
    thisMonth: monthCount,
  };
}

function getRecentRegistrations(limit = 10) {
  const sh = _sh(CONFIG.REQUEST_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;
  const data = sh.getRange(startRow, 1, numRows, 30).getValues();
  return data.reverse().map((row) => ({
    stt: row[0],
    tenThietBi: row[1],
    hocPhan: row[10],
    giangVien: row[15],
    email: row[16],
    thoiGian: row[19],
    idDexuat: row[29],
  }));
}

function exportSettingsToJSON() {
  const settings = getAdminSettings();
  return JSON.stringify(settings, null, 2);
}

function importSettingsFromJSON(jsonString) {
  try {
    const settings = JSON.parse(jsonString);
    saveAdminSettings(settings);
    return { success: true, message: "Đã import cài đặt thành công" };
  } catch (e) {
    return { success: false, message: "Lỗi: " + e.message };
  }
}

function resetToDefaultSettings() {
  const props = PropertiesService.getScriptProperties();
  props.deleteAllProperties();
  return { success: true, message: "Đã khôi phục cài đặt mặc định" };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function adminLogin(id, password) {
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty("adminId") || "admin";
  const storedPassword = props.getProperty("adminPassword") || "eiu@123456";

  if (id === storedId && password === storedPassword) {
    return { success: true, message: "Đăng nhập thành công" };
  } else {
    return { success: false, message: "ID hoặc mật khẩu không đúng" };
  }
}

function changeAdminPassword(oldPass, newPass) {
  const props = PropertiesService.getScriptProperties();
  const currentPassword = props.getProperty("adminPassword") || "eiu@123456";
  if (oldPass !== currentPassword) {
    return { success: false, message: "Mật khẩu cũ không đúng" };
  }
  props.setProperty("adminPassword", newPass);
  return { success: true, message: "Mật khẩu đã được thay đổi" };
}

function sendAdminPasswordToEmail(email) {
  const props = PropertiesService.getScriptProperties();
  const currentPassword = props.getProperty("adminPassword") || "eiu@123456";
  try {
    GmailApp.sendEmail(
      email,
      "Reset mật khẩu admin",
      `Mật khẩu hiện tại của bạn là: ${currentPassword}`,
    );
    return { success: true, message: "Mật khẩu đã được gửi qua email" };
  } catch (e) {
    return { success: false, message: "Không thể gửi email: " + e.message };
  }
}
