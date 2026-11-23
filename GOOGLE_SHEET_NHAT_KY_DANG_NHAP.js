/**
 * =========================================================================================
 * GOOGLE APPS SCRIPT - NHẬT KÝ THEO DÕI CÁC MÁY TÍNH ĐĂNG NHẬP & SỬ DỤNG HỆ THỐNG
 * Tác giả: VinhPH
 * =========================================================================================
 * HƯỚNG DẪN CÀI ĐẶT NHANH TRONG 30 GIÂY:
 * 1. Mở trang: https://sheets.google.com -> Tạo 1 trang tính mới (Đặt tên: "NHAT_KY_SU_DUNG_TOOL")
 * 2. Trên thanh menu, chọn "Tiện ích mở rộng" (Extensions) -> Chọn "Apps Script"
 * 3. Xóa sạch mọi mã có sẵn, copy toàn bộ nội dung file này dán vào -> Bấm Lưu (Ctrl + S)
 * 4. Bấm nút "Triển khai" (Deploy) ở góc trên bên phải -> Chọn "Tùy chọn triển khai mới" (New deployment):
 *    - Chọn loại: "Ứng dụng web" (Web app)
 *    - Mô tả: "Nhật ký truy cập hệ thống thuế"
 *    - Thực thi dưới dạng (Execute as): "Tôi" (Me)
 *    - Ai có quyền truy cập (Who has access): "Bất kỳ ai" (Anyone) -> [RẤT QUAN TRỌNG ĐỂ HỆ THỐNG GỬI ĐƯỢC]
 *    - Bấm "Triển khai" (Deploy) -> Cấp quyền cho Google nếu có hộp thoại hiện lên
 * 5. Sao chép "URL ứng dụng web" nhận được (có dạng: https://script.google.com/macros/s/AKfycb.../exec)
 *    và dán vào mục Cấu hình Google Sheet trên ứng dụng Web để bắt đầu ghi nhận tự động!
 * =========================================================================================
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var rawData = {};
    if (e && e.postData && e.postData.contents) {
      try {
        rawData = JSON.parse(e.postData.contents);
      } catch (err) {
        rawData = e.parameter || {};
      }
    } else if (e && e.parameter) {
      rawData = e.parameter;
    }
    
    return recordEntry(rawData);
  } catch (err) {
    return jsonResponse({status: "ERROR", message: String(err)});
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var rawData = (e && e.parameter) ? e.parameter : {};
    return recordEntry(rawData);
  } catch (err) {
    return jsonResponse({status: "ERROR", message: String(err)});
  } finally {
    lock.releaseLock();
  }
}

function recordEntry(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  setupSheetHeaderIfNeeded(sheet);
  
  var now = new Date();
  var timeStr = data.timestamp || Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  var deviceId = String(data.device_id || "Chưa rõ").trim();
  var clientName = String(data.client_name || "Khách truy cập").trim();
  var ip = String(data.ip || "127.0.0.1").trim();
  var osAndBrowser = String(data.user_agent || "Trình duyệt Web").trim();
  var platform = String(data.platform || "Web Online (Streamlit)").trim();
  var action = String(data.action || "Đăng nhập hệ thống").trim();
  var fileCount = data.file_count !== undefined ? String(data.file_count) : "-";
  var note = String(data.note || "Bình thường").trim();
  
  // Ghi thêm dòng mới vào bảng tính
  sheet.appendRow([
    timeStr,
    deviceId,
    clientName,
    ip,
    osAndBrowser,
    platform,
    action,
    fileCount,
    note
  ]);
  
  var lastRow = sheet.getLastRow();
  
  // Định dạng căn chỉnh và đường viền
  var rowRange = sheet.getRange(lastRow, 1, 1, 9);
  rowRange.setVerticalAlignment("middle");
  
  // Căn giữa các cột thông số ngắn
  sheet.getRange(lastRow, 1).setHorizontalAlignment("center");
  sheet.getRange(lastRow, 2).setHorizontalAlignment("center").setFontFamily("Consolas");
  sheet.getRange(lastRow, 4).setHorizontalAlignment("center").setFontFamily("Consolas");
  sheet.getRange(lastRow, 6).setHorizontalAlignment("center");
  sheet.getRange(lastRow, 7).setHorizontalAlignment("center");
  sheet.getRange(lastRow, 8).setHorizontalAlignment("center");
  
  // Màu sắc theo hành động
  var actionCell = sheet.getRange(lastRow, 7);
  if (action.indexOf("Bóc tách") !== -1 || action.indexOf("Xử lý") !== -1) {
    actionCell.setBackground("#DCFCE7").setFontColor("#166534").setFontWeight("bold");
  } else if (action.indexOf("Đăng nhập") !== -1) {
    actionCell.setBackground("#DBEAFE").setFontColor("#1E40AF").setFontWeight("bold");
  }
  
  return jsonResponse({
    status: "SUCCESS",
    message: "Đã cập nhật danh sách máy tính thành công vào hàng " + lastRow,
    timestamp: timeStr,
    device_id: deviceId
  });
}

function setupSheetHeaderIfNeeded(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = [
      "Thời gian truy cập (GMT+7)",
      "Mã thiết bị / Session ID",
      "Tên máy / Người dùng",
      "Địa chỉ IP",
      "Hệ điều hành & Trình duyệt",
      "Nền tảng sử dụng",
      "Hành động thực hiện",
      "Số lượng tệp",
      "Ghi chú chi tiết"
    ];
    sheet.appendRow(headers);
    
    var headerRange = sheet.getRange(1, 1, 1, 9);
    headerRange
      .setFontWeight("bold")
      .setBackground("#1E3A8A")
      .setFontColor("#FFFFFF")
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center");
      
    sheet.setRowHeight(1, 35);
    sheet.setFrozenRows(1);
    
    sheet.setColumnWidth(1, 175); // Thời gian
    sheet.setColumnWidth(2, 170); // Mã thiết bị
    sheet.setColumnWidth(3, 190); // Tên máy
    sheet.setColumnWidth(4, 140); // IP
    sheet.setColumnWidth(5, 230); // OS & Browser
    sheet.setColumnWidth(6, 170); // Nền tảng
    sheet.setColumnWidth(7, 180); // Hành động
    sheet.setColumnWidth(8, 110); // Số lượng tệp
    sheet.setColumnWidth(9, 220); // Ghi chú
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}