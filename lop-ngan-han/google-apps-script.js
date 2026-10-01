/**
 * =========================================================================
 * GOOGLE APPS SCRIPT - TỰ ĐỘNG LƯU ĐĂNG KÝ VÀO GOOGLE SHEETS & GỬI EMAIL
 * KHOA CÔNG NGHỆ THÔNG TIN - TRƯỜNG ĐẠI HỌC LẠC HỒNG (LHU)
 * =========================================================================
 * 
 * ⚠️ NGUYÊN NHÂN EMAIL CHƯA ĐẾN HỘP THƯ:
 * 1. Google yêu cầu CẤP QUYỀN GỬI EMAIL (Authorization) cho script. Nếu chưa
 *    bấm "Chạy" (Run) trong Apps Script một lần thì Google sẽ chặn gửi thư ngầm!
 * 2. Cần triển khai "Phiên bản mới" (New version) để Google nhận mã mới.
 * 
 * =========================================================================
 * HƯỚNG DẪN CÀI ĐẶT & KÍCH HOẠT QUYỀN GỬI EMAIL (CHỈ 1 PHÚT):
 * =========================================================================
 * 
 * BƯỚC 1: Dán toàn bộ nội dung file này vào Trình soạn thảo Apps Script. Bấm Lưu (Ctrl + S).
 * 
 * BƯỚC 2 (CỰC KỲ QUAN TRỌNG - KÍCH HOẠT QUYỀN GỬI EMAIL):
 *    - Nhìn lên thanh công cụ phía trên (cạnh nút "Gỡ lỗi" / "Chạy"):
 *      Chọn hàm: "testSendEmail".
 *    - Bấm nút "▶ Chạy" (Run).
 *    - Google sẽ hiển thị popup: "Cần được ủy quyền" (Authorization Required):
 *      -> Bấm "Xem lại quyền" (Review Permissions).
 *      -> Chọn tài khoản Google của bạn.
 *      -> Bấm "Nâng cao" (Advanced) ở góc dưới bên trái popup.
 *      -> Bấm "Đi tới... (không an toàn)" / "Go to ... (unsafe)".
 *      -> Bấm "Cho phép" (Allow) để cho phép quyền gửi email và ghi trang tính.
 *    - Ngay sau đó, bạn sẽ thấy Google gửi một email test vào chính hộp thư của bạn!
 * 
 * BƯỚC 3 (CẬP NHẬT BẢN TRIỂN KHAI MỚI):
 *    - Bấm nút "Triển khai" (Deploy) ở góc trên bên phải -> Chọn "Quản lý các bản triển khai" (Manage deployments).
 *    - Bấm icon "Cây bút chì" (Chỉnh sửa / Edit).
 *    - Ở dòng "Phiên bản" (Version): Bấm chọn "Phiên bản mới" (New version).
 *    - Bấm "Triển khai" (Deploy) -> Bấm "Xong" (Done).
 * 
 * Xong! Từ thời điểm này, bất kỳ ai đăng ký trên web sẽ nhận được email ngay lập tức!
 * =========================================================================
 */

/**
 * =========================================================================
 * HÀM TEST ĐỂ KÍCH HOẠT CẤP QUYỀN & KIỂM TRA GỬI EMAIL NGAY LẬP TỨC
 * =========================================================================
 */
function testSendEmail() {
  var currentAccount = Session.getActiveUser().getEmail();
  var myEmail = currentAccount || "phuc@lhu.edu.vn";

  Logger.log(">>> Tài khoản Google đang chạy script: " + currentAccount);
  Logger.log(">>> Đang gửi thử email test xác nhận đến: " + myEmail);

  try {
    sendConfirmationEmail({
      regCode: "LHU-" + Math.floor(100000 + Math.random() * 900000),
      name: "Thầy Phúc (Test Kích Hoạt)",
      phone: "0912345678",
      email: myEmail,
      course: "⭐ COMBO TRỌN BỘ 4 KHÓA HỌC - 1.990.000 VNĐ",
      price: "1.990.000 VNĐ",
      learning_mode: "Online Tương Tác Qua Zoom (Có video xem lại)",
      time: Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss")
    });

    Logger.log("✅ THÀNH CÔNG! Đã gửi email test từ phuc@lhu.edu.vn đến: " + myEmail);
    Logger.log("👉 Vui lòng mở hòm thư " + myEmail + " để kiểm tra thư xác nhận!");
  } catch (err) {
    Logger.log("❌ LỖI GỬI EMAIL: " + err.toString());
  }
}

/**
 * =========================================================================
 * HÀM XỬ LÝ NHẬN FORM TỪ TRANG WEB (doPost)
 * =========================================================================
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getActiveSheet();

    // 1. Tự động tạo hàng tiêu đề màu xanh LHU (#013D88) nếu trang tính đang trống
    if (sheet.getLastRow() === 0) {
      var headers = [
        "Thời Gian",
        "Mã Hồ Sơ",
        "Họ và Tên",
        "Số Điện Thoại (Zalo)",
        "Email Nhận Tài Liệu",
        "Khóa Học Lựa Chọn",
        "Học Phí Ưu Đãi",
        "Hình Thức Học",
        "Trạng Thái",
        "Ghi Chú"
      ];
      sheet.appendRow(headers);
      
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#013D88");
      headerRange.setFontColor("#FFFFFF");
      headerRange.setHorizontalAlignment("center");
      headerRange.setVerticalAlignment("middle");
      sheet.setRowHeight(1, 40);
      sheet.setFrozenRows(1);

      sheet.setColumnWidth(1, 170); // Thời gian
      sheet.setColumnWidth(2, 130); // Mã hồ sơ
      sheet.setColumnWidth(3, 200); // Họ tên
      sheet.setColumnWidth(4, 160); // SĐT
      sheet.setColumnWidth(5, 230); // Email
      sheet.setColumnWidth(6, 320); // Khóa học
      sheet.setColumnWidth(7, 150); // Học phí
      sheet.setColumnWidth(8, 200); // Hình thức
      sheet.setColumnWidth(9, 140); // Trạng thái
      sheet.setColumnWidth(10, 200); // Ghi chú
    }

    // 2. Nhận dữ liệu gửi lên
    var data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else {
      data = e.parameter || {};
    }

    var timestamp = data.time || Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    var regCode = "LHU-" + Math.floor(100000 + Math.random() * 900000);
    var name = data.name || data.fullname || "Học viên";
    var rawPhone = (data.phone || "").toString().replace(/['"]/g, '');
    var phone = "'" + rawPhone;
    var email = (data.email || "").toString().trim();
    var course = data.course || data.course_select || "Khóa học ngắn hạn thực chiến";
    var price = data.price || "1.990.000 VNĐ";
    var learning_mode = data.learning_mode || data.mode || "Online Tương Tác Qua Zoom";
    var status = "Mới đăng ký";
    var note = data.note || "Chờ liên hệ xác nhận trong 15p";

    // 3. Ghi dòng dữ liệu mới vào Google Sheet
    sheet.appendRow([
      timestamp,
      regCode,
      name,
      phone,
      email,
      course,
      price,
      learning_mode,
      status,
      note
    ]);

    // 4. GỬI EMAIL TỰ ĐỘNG XÁC NHẬN CHO HỌC VIÊN
    var emailSent = false;
    var emailError = "";

    if (email && email.indexOf("@") !== -1) {
      try {
        sendConfirmationEmail({
          regCode: regCode,
          name: name,
          phone: rawPhone,
          email: email,
          course: course,
          price: price,
          learning_mode: learning_mode,
          time: timestamp
        });
        emailSent = true;
      } catch (mailErr) {
        emailError = mailErr.toString();
        Logger.log("Lỗi gửi email: " + emailError);
      }
    }

    // 5. Trả kết quả JSON phản hồi
    return ContentService
      .createTextOutput(JSON.stringify({ 
        result: "success", 
        message: "Đăng ký thành công! Dữ liệu đã lưu vào Google Sheet" + (emailSent ? " và email xác nhận đã gửi thành công." : (emailError ? " (Lưu ý email: " + emailError + ")" : ".")),
        emailSent: emailSent,
        emailError: emailError,
        regCode: regCode
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ 
        result: "error", 
        error: error.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

/**
 * =========================================================================
 * HÀM GỬI EMAIL XÁC NHẬN (Hỗ trợ cả GmailApp và MailApp)
 * =========================================================================
 */
function sendConfirmationEmail(data) {
  // Loại bỏ triệt để emoji và ký tự 4-byte trong tiêu đề để KHÔNG bị lỗi icon  trên mọi thiết bị
  var rawCourse = data.course || "Khóa học thực chiến";
  var cleanCourse = rawCourse
    .replace(/[\p{Extended_Pictographic}\u2600-\u27BF]/gu, '')
    .replace(/\s+/g, ' ')
    .trim();

  var subject = "[LHU - Khoa CNTT] Xác nhận đăng ký giữ chỗ thành công: " + cleanCourse;
  var htmlBody = buildEmailTemplate(data);

  // Thử gửi bằng GmailApp để thư gửi từ tài khoản phuc@lhu.edu.vn và lưu vào Thư đã gửi
  try {
    GmailApp.sendEmail(data.email, subject, "", {
      htmlBody: htmlBody,
      name: "Khoa CNTT - ĐH Lạc Hồng",
      replyTo: "phuc@lhu.edu.vn"
    });
  } catch (gmailErr) {
    // Fallback sang MailApp nếu cần
    MailApp.sendEmail({
      to: data.email,
      subject: subject,
      htmlBody: htmlBody,
      name: "Khoa CNTT - ĐH Lạc Hồng",
      replyTo: "phuc@lhu.edu.vn"
    });
  }
}

/**
 * =========================================================================
 * TEMPLATE EMAIL HTML CHUẨN THƯƠNG HIỆU LHU - ASU - ABET (KHÔNG LỖI FONT)
 * =========================================================================
 */
function buildEmailTemplate(data) {
  var regCode = data.regCode || "LHU-" + Math.floor(100000 + Math.random() * 900000);
  var name = escapeHtml(data.name || "Học viên");
  var phone = escapeHtml(data.phone || "---");
  var email = escapeHtml(data.email || "---");
  
  // Làm sạch emoji trong tên khóa học
  var rawCourse = (data.course || "Khóa học ngắn hạn thực chiến")
    .replace(/[\p{Extended_Pictographic}\u2600-\u27BF]/gu, '')
    .replace(/\s+/g, ' ')
    .trim();
  var course = escapeHtml(rawCourse);

  var price = escapeHtml(data.price || "1.990.000 VNĐ");
  var learning_mode = escapeHtml(data.learning_mode || "Online Tương Tác Qua Zoom");
  var time = escapeHtml(data.time || "");

  return `
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <title>Xác nhận đăng ký khóa học</title>
</head>
<body style="margin: 0; padding: 15px; background-color: #f1f5f9; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #334155;">

  <table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color: #f1f5f9; padding: 15px 0;">
    <tr>
      <td align="center">
        <table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0" style="max-width: 620px; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 25px rgba(0,0,0,0.08); border: 1px solid #e2e8f0;">
          
          <!-- BRAND ACCENT BAR -->
          <tr>
            <td style="background: linear-gradient(90deg, #013d88 0%, #013d88 50%, #e0b341 50%, #e0b341 75%, #f37021 75%, #f37021 100%); height: 6px;"></td>
          </tr>

          <!-- HEADER -->
          <tr>
            <td style="background-color: #013d88; padding: 24px 28px; text-align: left;">
              <table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td style="vertical-align: middle;">
                    <div style="font-size: 11px; font-weight: 800; color: #e0b341; text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 4px;">
                      TRƯỜNG ĐẠI HỌC LẠC HỒNG (LHU)
                    </div>
                    <div style="font-size: 18px; font-weight: 900; color: #ffffff; text-transform: uppercase; letter-spacing: 0.5px; line-height: 1.3;">
                      KHOA CÔNG NGHỆ THÔNG TIN
                    </div>
                    <div style="margin-top: 8px;">
                      <span style="display: inline-block; background-color: #8d1d41; color: #ffffff; font-size: 9px; font-weight: 800; padding: 3px 8px; border-radius: 4px; text-transform: uppercase; margin-right: 6px;">
                        POWERED BY ASU (HOA KỲ)
                      </span>
                      <span style="display: inline-block; background-color: #f37021; color: #ffffff; font-size: 9px; font-weight: 800; padding: 3px 8px; border-radius: 4px; text-transform: uppercase;">
                        KIỂM ĐỊNH QUỐC TẾ ABET
                      </span>
                    </div>
                  </td>
                  <td align="right" style="vertical-align: middle; width: 50px;">
                    <div style="width: 48px; height: 48px; border-radius: 12px; background-color: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.25); text-align: center; line-height: 48px;">
                      <svg width="24" height="24" viewBox="0 0 24 24" fill="#ffffff" style="display: inline-block; vertical-align: middle;">
                        <path d="M12 3L1 9L12 15L21 10.09V17H23V9M5 13.18V17.18C5 19.94 8.13 22 12 22C15.87 22 19 19.94 19 17.18V13.18L12 17L5 13.18Z"/>
                      </svg>
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- HERO CONFIRMATION BANNER -->
          <tr>
            <td style="background: linear-gradient(180deg, #eff6ff 0%, #ffffff 100%); padding: 28px 28px 16px 28px; text-align: center; border-bottom: 1px solid #f1f5f9;">
              <div style="display: inline-block; width: 60px; height: 60px; border-radius: 50%; background-color: #dcfce7; line-height: 60px; text-align: center; margin-bottom: 12px; box-shadow: 0 4px 12px rgba(22,163,74,0.15);">
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#16a34a" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round" style="display: inline-block; vertical-align: middle;">
                  <polyline points="20 6 9 17 4 12"></polyline>
                </svg>
              </div>
              <h1 style="margin: 0 0 6px 0; font-size: 21px; font-weight: 900; color: #0f172a; text-transform: uppercase; letter-spacing: 0.3px;">
                Xác Nhận Đăng Ký Thành Công!
              </h1>
              <p style="margin: 0; font-size: 13.5px; color: #475569; line-height: 1.5; max-width: 460px; margin: 0 auto;">
                Cảm ơn bạn đã lựa chọn nâng cấp kỹ năng cùng <strong style="color: #013d88;">Khoa CNTT - ĐH Lạc Hồng</strong>. Thông tin giữ chỗ ưu đãi của bạn đã được ghi nhận vào hệ thống.
              </p>
            </td>
          </tr>

          <!-- DETAILS TABLE -->
          <tr>
            <td style="padding: 20px 28px;">
              <p style="font-size: 14.5px; color: #1e293b; line-height: 1.6; margin: 0 0 16px 0;">
                Kính gửi <strong>${name}</strong>,
              </p>
              <p style="font-size: 13.5px; color: #475569; line-height: 1.6; margin: 0 0 18px 0;">
                Ban Đào tạo xin gửi đến bạn thông tin tóm tắt hồ sơ đăng ký khóa học thực chiến ngắn hạn như sau:
              </p>

              <table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; margin-bottom: 22px;">
                <tr>
                  <td colspan="2" style="background-color: #013d88; padding: 11px 16px; color: #ffffff; font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px;">
                    <span style="display: inline-block; width: 7px; height: 7px; border-radius: 50%; background-color: #e0b341; margin-right: 6px; vertical-align: 1px;"></span> CHI TIẾT HỒ SƠ ĐĂNG KÝ
                  </td>
                </tr>
                <tr>
                  <td style="padding: 11px 16px; font-size: 13px; color: #64748b; font-weight: 600; border-bottom: 1px solid #e2e8f0; width: 38%;">
                    Mã hồ sơ:
                  </td>
                  <td style="padding: 11px 16px; font-size: 13px; color: #013d88; font-weight: 800; border-bottom: 1px solid #e2e8f0;">
                    #${regCode}
                  </td>
                </tr>
                <tr>
                  <td style="padding: 11px 16px; font-size: 13px; color: #64748b; font-weight: 600; border-bottom: 1px solid #e2e8f0;">
                    Họ và tên học viên:
                  </td>
                  <td style="padding: 11px 16px; font-size: 13.5px; color: #0f172a; font-weight: 700; border-bottom: 1px solid #e2e8f0;">
                    ${name}
                  </td>
                </tr>
                <tr>
                  <td style="padding: 11px 16px; font-size: 13px; color: #64748b; font-weight: 600; border-bottom: 1px solid #e2e8f0;">
                    Số điện thoại (Zalo):
                  </td>
                  <td style="padding: 11px 16px; font-size: 13px; color: #0f172a; font-weight: 700; border-bottom: 1px solid #e2e8f0;">
                    ${phone}
                  </td>
                </tr>
                <tr>
                  <td style="padding: 11px 16px; font-size: 13px; color: #64748b; font-weight: 600; border-bottom: 1px solid #e2e8f0;">
                    Khóa học đăng ký:
                  </td>
                  <td style="padding: 11px 16px; font-size: 13px; color: #b45309; font-weight: 800; border-bottom: 1px solid #e2e8f0;">
                    ${course}
                  </td>
                </tr>
                <tr>
                  <td style="padding: 11px 16px; font-size: 13px; color: #64748b; font-weight: 600; border-bottom: 1px solid #e2e8f0;">
                    Học phí ưu đãi:
                  </td>
                  <td style="padding: 11px 16px; font-size: 14.5px; color: #dc2626; font-weight: 900; border-bottom: 1px solid #e2e8f0;">
                    ${price}
                  </td>
                </tr>
                <tr>
                  <td style="padding: 11px 16px; font-size: 13px; color: #64748b; font-weight: 600; border-bottom: 1px solid #e2e8f0;">
                    Hình thức học:
                  </td>
                  <td style="padding: 11px 16px; font-size: 13px; color: #0f172a; font-weight: 700; border-bottom: 1px solid #e2e8f0;">
                    ${learning_mode}
                  </td>
                </tr>
                <tr>
                  <td style="padding: 11px 16px; font-size: 13px; color: #64748b; font-weight: 600;">
                    Trạng thái hồ sơ:
                  </td>
                  <td style="padding: 11px 16px; font-size: 12px;">
                    <span style="display: inline-block; background-color: #dcfce7; color: #15803d; font-weight: 800; padding: 4px 10px; border-radius: 6px; border: 1px solid #bbf7d0;">
                      ĐÃ GIỮ CHỖ THÀNH CÔNG
                    </span>
                  </td>
                </tr>
              </table>

              <!-- NEXT STEPS -->
              <div style="background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 12px; padding: 16px 18px; margin-bottom: 20px;">
                <div style="font-size: 12.5px; font-weight: 800; color: #166534; text-transform: uppercase; margin-bottom: 8px; letter-spacing: 0.5px;">
                  <span style="display: inline-block; width: 7px; height: 7px; border-radius: 50%; background-color: #16a34a; margin-right: 6px; vertical-align: 1px;"></span> 3 BƯỚC TIẾP THEO:
                </div>
                <div style="font-size: 12.5px; color: #1e293b; line-height: 1.6;">
                  <strong>1. Tư vấn viên liên hệ trong 15 phút:</strong> Chuyên viên của Khoa CNTT - LHU sẽ gọi điện hoặc kết bạn Zalo để xác nhận lớp học và gửi lịch khai giảng.<br>
                  <strong>2. Nhận tài liệu & phòng học:</strong> Bạn sẽ được cấp quyền truy cập tài liệu mẫu chuẩn ASU và link Zoom hoặc sơ đồ phòng Lab.<br>
                  <strong>3. Thực hành & Cấp chứng nhận:</strong> Tham gia 2 buổi (8 tiết) thực chiến cùng Giảng viên LHU & trợ lý AI Copilot.
                </div>
              </div>

              <!-- COMMITMENT NOTE -->
              <div style="background-color: #fffbeb; border: 1px solid #fde68a; border-radius: 10px; padding: 12px 16px; font-size: 12px; color: #92400e; line-height: 1.5; margin-bottom: 20px;">
                <strong>[CAM KẾT VÀNG]:</strong> Hoàn tiền 100% học phí nếu bạn không hài lòng về nội dung sau buổi học đầu tiên. Đội ngũ trợ giảng hỗ trợ kỹ thuật và kèm cặp 1-1 trọn đời.
              </div>

              <!-- HOTLINE CALLOUT -->
              <table role="presentation" width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-bottom: 6px;">
                <tr>
                  <td align="center">
                    <a href="tel:0912345678" style="display: inline-block; background-color: #013d88; color: #ffffff; text-decoration: none; font-size: 13.5px; font-weight: 800; padding: 11px 26px; border-radius: 8px; box-shadow: 0 4px 10px rgba(1,61,136,0.25);">
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="#ffffff" style="display: inline-block; vertical-align: -2px; margin-right: 6px;"><path d="M20 15.5c-1.25 0-2.45-.2-3.57-.57a1 1 0 0 0-1.02.24l-2.2 2.2a15.045 15.045 0 0 1-6.59-6.59l2.2-2.21a.96.96 0 0 0 .25-1A11.36 11.36 0 0 1 8.5 4c0-.55-.45-1-1-1H4c-.55 0-1 .45-1 1 0 9.39 7.61 17 17 17 .55 0 1-.45 1-1v-3.5c0-.55-.45-1-1-1zM19 12h2a9 9 0 0 0-9-9v2c3.87 0 7 3.13 7 7zm-4 0h2a5 5 0 0 0-5-5v2c1.66 0 3 1.34 3 3z"/></svg> Hotline Tư Vấn / Zalo: 0912 345 678
                    </a>
                  </td>
                </tr>
              </table>

            </td>
          </tr>

          <!-- FOOTER -->
          <tr>
            <td style="background-color: #0f172a; padding: 22px 28px; text-align: center; color: #94a3b8; font-size: 11px; line-height: 1.6;">
              <div style="font-weight: 800; color: #ffffff; font-size: 12.5px; margin-bottom: 5px;">
                KHOA CÔNG NGHỆ THÔNG TIN - TRƯỜNG ĐẠI HỌC LẠC HỒNG
              </div>
              <div style="margin-bottom: 8px;">
                Địa chỉ: Số 10, Huỳnh Văn Nghệ, P. Bửu Long, TP. Biên Hòa, Đồng Nai<br>
                Website: <a href="https://lhu.edu.vn" style="color: #38bdf8; text-decoration: none;">lhu.edu.vn</a> &bull; Email: <a href="mailto:phuc@lhu.edu.vn" style="color: #38bdf8; text-decoration: none;">phuc@lhu.edu.vn</a>
              </div>
              <div style="border-top: 1px solid #1e293b; padding-top: 10px; margin-top: 10px; color: #64748b; font-size: 10px;">
                Email này được gửi tự động từ hệ thống đăng ký khóa học Khoa CNTT - ĐH Lạc Hồng.<br>
                © 2026 Khoa CNTT - Đại học Lạc Hồng. Bảo lưu mọi quyền.
              </div>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>

</body>
</html>
  `;
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function doGet(e) {
  return ContentService
    .createTextOutput("Google Apps Script Web App - Khoa CNTT ĐH Lạc Hồng đang hoạt động bình thường!")
    .setMimeType(ContentService.MimeType.TEXT);
}
