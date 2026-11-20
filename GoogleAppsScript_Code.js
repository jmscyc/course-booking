/**
 * 霖霖的課程預約系統 - Google Apps Script (GAS) 後端程式碼
 * 
 * 部署說明：
 * 1. 在 Google Sheets 點擊「擴充功能」 ->「Apps Script」
 * 2. 將此檔案程式碼覆蓋原本的 Code.gs，點擊儲存 (Ctrl+S)
 * 3. 點擊右上角「部署」 -> 「新建部署」
 * 4. 種類選擇「網路應用程式 (Web App)」
 *    - 說明：課程預約與全功能管理 API
 *    - 執行身分：我 (您的 Google 帳戶)
 *    - 誰可以存取：任何人 (Anyone)
 * 5. 點擊部署，複製產生的「網路應用程式網址 (Web App URL)」
 * 6. 在 admin.html 右上角點擊「設定 API 網址」貼上即可！
 */

const SHEET_COURSES = "課程清單";
const SHEET_BOOKINGS = "預約記錄";
const SHEET_CONFIG = "系統設定";

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 建立「課程清單」表
  let courseSheet = ss.getSheetByName(SHEET_COURSES);
  if (!courseSheet) {
    courseSheet = ss.insertSheet(SHEET_COURSES);
    courseSheet.appendRow(["課程編號", "課程名稱", "授課教練", "日期", "開始時間", "結束時間", "人數上限", "目前報名數", "地點/教室", "備註說明", "狀態"]);
    courseSheet.getRange(1, 1, 1, 11).setBackground("#2563eb").setFontColor("#ffffff").setFontWeight("bold");
    
    courseSheet.appendRow(["C001", "核心肌群與皮拉提斯", "霖霖教練", "2026-11-25", "19:00", "20:00", 8, 2, "A教室", "請自備毛巾與水壺", "開啟報名"]);
    courseSheet.appendRow(["C002", "高強度間歇 HIIT 燃脂", "霖霖教練", "2026-11-26", "20:15", "21:15", 10, 5, "B教室", "適合有運動基礎學員", "開啟報名"]);
    courseSheet.appendRow(["C003", "瑜珈伸展與深層放鬆", "小雅教練", "2026-11-28", "10:00", "11:00", 6, 6, "A教室", "額滿釋出排隊中", "開啟報名"]);
  }

  // 2. 建立「預約記錄」表
  let bookingSheet = ss.getSheetByName(SHEET_BOOKINGS);
  if (!bookingSheet) {
    bookingSheet = ss.insertSheet(SHEET_BOOKINGS);
    bookingSheet.appendRow(["預約編號", "報名時間", "課程編號", "課程名稱", "學員姓名", "聯絡電話", "Line ID", "備註需求", "簽到狀態"]);
    bookingSheet.getRange(1, 1, 1, 9).setBackground("#059669").setFontColor("#ffffff").setFontWeight("bold");
    
    bookingSheet.appendRow(["BK-1001", "2026-11-20 10:15:00", "C001", "核心肌群與皮拉提斯", "王小明", "0912345678", "ming123", "第一次參加", "已預約"]);
    bookingSheet.appendRow(["BK-1002", "2026-11-20 11:30:00", "C001", "核心肌群與皮拉提斯", "林小美", "0987654321", "may_lin", "", "已預約"]);
  }

  // 3. 建立「系統設定」表
  let configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET_CONFIG);
    configSheet.appendRow(["設定項目", "設定值"]);
    configSheet.appendRow(["系統名稱", "霖霖的課程預約系統"]);
    configSheet.appendRow(["管理員密碼", "admin888"]);
    configSheet.appendRow(["預約提醒事項", "課程開始前 2 小時前取消不扣堂數，如額滿可隨時關注名額釋出。"]);
    configSheet.appendRow(["常用教室清單", "A教室 (瑜珈/皮拉提斯),B教室 (多功能有氧),重訓區 (槓鈴/機械器材),VIP私人教室,戶外場地"]);
    configSheet.getRange(1, 1, 1, 2).setBackground("#475569").setFontColor("#ffffff").setFontWeight("bold");
  }

  return "初始化完成！";
}

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const action = e.parameter.action || "getCourses";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "getCourses") {
      let courseSheet = ss.getSheetByName(SHEET_COURSES);
      let bookingSheet = ss.getSheetByName(SHEET_BOOKINGS);
      
      if (!courseSheet || !bookingSheet) {
        setupSheets();
        courseSheet = ss.getSheetByName(SHEET_COURSES);
        bookingSheet = ss.getSheetByName(SHEET_BOOKINGS);
      }

      const courses = getSheetData(courseSheet);
      const bookings = getSheetData(bookingSheet);

      const bookingCounts = {};
      bookings.forEach(b => {
        if (b["簽到狀態"] !== "已取消") {
          const cId = String(b["課程編號"]);
          bookingCounts[cId] = (bookingCounts[cId] || 0) + 1;
        }
      });

      const resultCourses = courses.map(c => {
        const cId = String(c["課程編號"]);
        const maxCap = Number(c["人數上限"]) || 0;
        const currentCount = bookingCounts[cId] || 0;
        const isFull = currentCount >= maxCap;
        const remaining = Math.max(0, maxCap - currentCount);

        return {
          id: cId,
          title: c["課程名稱"],
          instructor: c["授課教練"],
          date: formatDate(c["日期"]),
          startTime: c["開始時間"],
          endTime: c["結束時間"],
          capacity: maxCap,
          bookedCount: currentCount,
          remaining: remaining,
          location: c["地點/教室"],
          notes: c["備註說明"],
          status: c["狀態"] || "開啟報名",
          isFull: isFull
        };
      });

      return jsonResponse({ success: true, courses: resultCourses });
    }

    if (action === "getBookings") {
      const adminPass = e.parameter.password;
      const config = getConfigMap(ss);
      if (adminPass !== (config["管理員密碼"] || "admin888")) {
        return jsonResponse({ success: false, message: "管理員密碼錯誤" });
      }

      const bookings = getSheetData(ss.getSheetByName(SHEET_BOOKINGS));
      return jsonResponse({ success: true, bookings: bookings });
    }

    if (action === "getFullAdminData") {
      const adminPass = e.parameter.password;
      const config = getConfigMap(ss);
      if (adminPass !== (config["管理員密碼"] || "admin888")) {
        return jsonResponse({ success: false, message: "管理員密碼錯誤" });
      }

      let courseSheet = ss.getSheetByName(SHEET_COURSES);
      let bookingSheet = ss.getSheetByName(SHEET_BOOKINGS);
      if (!courseSheet || !bookingSheet) setupSheets();

      const courses = getSheetData(ss.getSheetByName(SHEET_COURSES));
      const bookings = getSheetData(ss.getSheetByName(SHEET_BOOKINGS));

      return jsonResponse({
        success: true,
        courses: courses,
        bookings: bookings,
        config: config
      });
    }

    return jsonResponse({ success: false, message: "無效請求" });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      data = e.parameter;
    }

    const action = data.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 學員預約報名
    if (action === "bookCourse") {
      const courseId = String(data.courseId);
      const studentName = String(data.name || "").trim();
      const studentPhone = String(data.phone || "").trim();
      const lineId = String(data.lineId || "").trim();
      const notes = String(data.notes || "").trim();

      if (!courseId || !studentName || !studentPhone) {
        return jsonResponse({ success: false, message: "請填寫完整資訊！" });
      }

      const courseSheet = ss.getSheetByName(SHEET_COURSES);
      const bookingSheet = ss.getSheetByName(SHEET_BOOKINGS);
      const courses = getSheetData(courseSheet);
      const targetCourse = courses.find(c => String(c["課程編號"]) === courseId);

      if (!targetCourse) return jsonResponse({ success: false, message: "課程不存在" });
      if (targetCourse["狀態"] !== "開啟報名") return jsonResponse({ success: false, message: "該課程已停止報名" });

      const bookings = getSheetData(bookingSheet);
      const isAlreadyBooked = bookings.some(b => 
        String(b["課程編號"]) === courseId && 
        String(b["聯絡電話"]).trim() === studentPhone && 
        b["簽到狀態"] !== "已取消"
      );
      if (isAlreadyBooked) return jsonResponse({ success: false, message: "您已經預約過這堂課程囉！" });

      const currentBooked = bookings.filter(b => String(b["課程編號"]) === courseId && b["簽到狀態"] !== "已取消").length;
      const maxCap = Number(targetCourse["人數上限"]) || 0;
      if (currentBooked >= maxCap) return jsonResponse({ success: false, message: "該課程已額滿！" });

      const bookingId = "BK-" + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMddHHmmss") + "-" + Math.floor(Math.random() * 899 + 100);
      const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");

      bookingSheet.appendRow([
        bookingId,
        nowStr,
        courseId,
        targetCourse["課程名稱"],
        studentName,
        studentPhone,
        lineId,
        notes,
        "已預約"
      ]);

      updateCourseCount(courseSheet, courseId, currentBooked + 1);

      return jsonResponse({
        success: true,
        message: "預約成功！",
        bookingId: bookingId,
        courseTitle: targetCourse["課程名稱"],
        dateTime: formatDate(targetCourse["日期"]) + " " + targetCourse["開始時間"]
      });
    }

    // 管理端身分驗證
    const adminPass = data.password;
    const config = getConfigMap(ss);
    if (adminPass !== (config["管理員密碼"] || "admin888")) {
      return jsonResponse({ success: false, message: "管理員密碼錯誤" });
    }

    // 管理端：新增或更新課程
    if (action === "saveCourse") {
      const courseSheet = ss.getSheetByName(SHEET_COURSES);
      const courseRows = courseSheet.getDataRange().getValues();
      const courseId = data.courseId; // 若有則為編輯，若無則為新增

      if (courseId) {
        // 更新現有課程
        let rowIdx = -1;
        for (let r = 1; r < courseRows.length; r++) {
          if (String(courseRows[r][0]) === courseId) {
            rowIdx = r + 1;
            break;
          }
        }
        if (rowIdx !== -1) {
          courseSheet.getRange(rowIdx, 2).setValue(data.title);
          courseSheet.getRange(rowIdx, 3).setValue(data.instructor);
          courseSheet.getRange(rowIdx, 4).setValue(data.date);
          courseSheet.getRange(rowIdx, 5).setValue(data.startTime);
          courseSheet.getRange(rowIdx, 6).setValue(data.endTime);
          courseSheet.getRange(rowIdx, 7).setValue(Number(data.capacity));
          courseSheet.getRange(rowIdx, 9).setValue(data.location);
          courseSheet.getRange(rowIdx, 10).setValue(data.notes || "");
          courseSheet.getRange(rowIdx, 11).setValue(data.status || "開啟報名");
          return jsonResponse({ success: true, message: "課程更新成功！" });
        }
      }

      // 新增課程
      const newId = "C" + Utilities.formatDate(new Date(), "GMT+8", "MMddHHmm");
      courseSheet.appendRow([
        newId,
        data.title,
        data.instructor || "霖霖教練",
        data.date,
        data.startTime,
        data.endTime,
        Number(data.capacity) || 10,
        0,
        data.location || "活動教室",
        data.notes || "",
        data.status || "開啟報名"
      ]);

      return jsonResponse({ success: true, message: "新課程已成功上架！", courseId: newId });
    }

    // 管理端：刪除課程
    if (action === "deleteCourse") {
      const courseSheet = ss.getSheetByName(SHEET_COURSES);
      const courseRows = courseSheet.getDataRange().getValues();
      const courseId = data.courseId;

      for (let r = 1; r < courseRows.length; r++) {
        if (String(courseRows[r][0]) === courseId) {
          courseSheet.deleteRow(r + 1);
          return jsonResponse({ success: true, message: "課程已刪除！" });
        }
      }
      return jsonResponse({ success: false, message: "找不到該課程編號" });
    }

    // 管理端：更新預約簽到狀態
    if (action === "updateBookingStatus") {
      const bookingSheet = ss.getSheetByName(SHEET_BOOKINGS);
      const bookingRows = bookingSheet.getDataRange().getValues();
      const bookingId = data.bookingId;
      const newStatus = data.newStatus;

      for (let r = 1; r < bookingRows.length; r++) {
        if (String(bookingRows[r][0]) === bookingId) {
          bookingSheet.getRange(r + 1, 9).setValue(newStatus);
          return jsonResponse({ success: true, message: "預約狀態已更新為：" + newStatus });
        }
      }
      return jsonResponse({ success: false, message: "找不到該預約記錄" });
    }

    // 管理端：更新系統全域設定
    if (action === "saveSettings") {
      const configSheet = ss.getSheetByName(SHEET_CONFIG);
      const rows = configSheet.getDataRange().getValues();
      
      const newSysName = data.systemName;
      const newNotice = data.notice;
      const newPass = data.newPassword;
      const newClassrooms = data.classrooms;

      const keysPresent = {};
      for (let r = 1; r < rows.length; r++) {
        const key = String(rows[r][0]).trim();
        keysPresent[key] = r + 1;
        if (key === "系統名稱" && newSysName) configSheet.getRange(r + 1, 2).setValue(newSysName);
        if (key === "預約提醒事項" && newNotice) configSheet.getRange(r + 1, 2).setValue(newNotice);
        if (key === "管理員密碼" && newPass) configSheet.getRange(r + 1, 2).setValue(newPass);
        if (key === "常用教室清單" && newClassrooms) configSheet.getRange(r + 1, 2).setValue(newClassrooms);
      }

      if (newClassrooms && !keysPresent["常用教室清單"]) {
        configSheet.appendRow(["常用教室清單", newClassrooms]);
      }

      return jsonResponse({ success: true, message: "系統設定已更新！" });
    }

    return jsonResponse({ success: false, message: "未支援的操作" });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function getSheetData(sheet) {
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const obj = {};
    headers.forEach((header, colIdx) => {
      obj[header] = row[colIdx];
    });
    data.push(obj);
  }
  return data;
}

function getConfigMap(ss) {
  const configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!configSheet) return {};
  const rows = configSheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) {
      map[String(rows[i][0]).trim()] = rows[i][1];
    }
  }
  return map;
}

function updateCourseCount(courseSheet, courseId, count) {
  const values = courseSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === courseId) {
      courseSheet.getRange(i + 1, 8).setValue(count);
      break;
    }
  }
}

function formatDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
  }
  return String(val);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
