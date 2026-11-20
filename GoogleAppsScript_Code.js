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
    courseSheet.appendRow(["課程編號", "課程名稱", "授課教練", "日期", "開始時間", "結束時間", "人數上限", "目前報名數", "地點/教室", "備註說明", "狀態", "虛擬報名數"]);
    courseSheet.getRange(1, 1, 1, 12).setBackground("#2563eb").setFontColor("#ffffff").setFontWeight("bold");
    
    const realCourses = [
      [
            "260428",
            "❷ 學習：頌缽入門",
            "霖霖",
            "2026-11-22",
            "08:30",
            "12:30",
            4,
            0,
            "高雄-近高美館",
            "想學頌缽，照顧自己 (早鳥價 $4000 / 定價 $5000)",
            "開啟報名"
      ],
      [
            "260430",
            "❷ 學習：初進階整日班",
            "霖霖",
            "2026-11-21",
            "09:00",
            "17:30",
            1,
            0,
            "台南安南醫院",
            "想學頌缽，照顧自己，幫助他人 (早鳥價 $9000 / 定價 $12000)",
            "開啟報名"
      ],
      [
            "260435",
            "❸ 學習：頌缽進階",
            "霖霖",
            "2027-01-24",
            "11:00",
            "15:00",
            8,
            0,
            "台北場地",
            "想進階技能，運用頌缽 (早鳥價 $4000 / 定價 $5000)",
            "開啟報名"
      ],
      [
            "260436",
            "❷ 學習：初進階整日班",
            "霖霖",
            "2027-01-17",
            "09:00",
            "18:00",
            8,
            0,
            "高雄-近高美館",
            "學習整套頌缽使用方式 (早鳥價 $8000 / 定價 $12000)",
            "開啟報名"
      ],
      [
            "260437",
            "❷ 學習：初進階整日班",
            "霖霖",
            "2026-12-27",
            "09:00",
            "18:00",
            8,
            0,
            "高雄-近高美館",
            "學習整套頌缽使用方式 (早鳥價 $8000 / 定價 $10000)",
            "開啟報名"
      ],
      [
            "260438",
            "❸ 學習：頌缽進階",
            "霖霖",
            "2026-11-22",
            "14:00",
            "18:00",
            7,
            0,
            "高雄-近高美館",
            "想進階技能，運用頌缽 (早鳥價 $4000 / 定價 $5000)",
            "開啟報名"
      ],
      [
            "260439",
            "❷ 學習：初進階整日班",
            "霖霖",
            "2027-01-23",
            "09:00",
            "17:30",
            5,
            0,
            "台北場地",
            "想學頌缽，照顧自己 (早鳥價 $8000 / 定價 $12000)",
            "開啟報名"
      ],
      [
            "260420",
            "❷ 學習：頌缽入門",
            "霖霖",
            "2026-10-18",
            "09:00",
            "13:00",
            10,
            0,
            "高雄-近高美館",
            "歷史課程 (早鳥價 $4000 / 定價 $5000)",
            "已結束"
      ],
      [
            "260423",
            "❷ 學習：頌缽入門",
            "霖霖",
            "2026-09-05",
            "09:00",
            "13:00",
            6,
            0,
            "台中市東區",
            "歷史課程 (早鳥價 $4000 / 定價 $5000)",
            "已結束"
      ],
      [
            "260418",
            "❷ 學習：頌缽入門",
            "霖霖",
            "2026-09-14",
            "09:00",
            "13:00",
            2,
            0,
            "高雄-近高美館",
            "歷史課程 (早鳥價 $4000 / 定價 $5000)",
            "已結束"
      ],
      [
            "260429",
            "❷ 學習：頌缽入門",
            "霖霖",
            "2026-10-11",
            "09:00",
            "13:00",
            4,
            0,
            "台北新北地區",
            "歷史課程 (早鳥價 $4000 / 定價 $5000)",
            "已結束"
      ],
      [
            "260421",
            "❸ 學習：頌缽進階",
            "霖霖",
            "2026-10-18",
            "14:00",
            "18:00",
            5,
            0,
            "高雄-近高美館",
            "歷史課程 (早鳥價 $4000 / 定價 $5000)",
            "已結束"
      ],
      [
            "260434",
            "❹ 團體頌缽音樂療癒",
            "霖霖",
            "2026-10-26",
            "19:00",
            "21:00",
            5,
            0,
            "高雄-近高美館",
            "歷史課程 (早鳥價 $800 / 定價 $1000)",
            "已結束"
      ],
      [
            "260431",
            "❷ 學習：頌缽入門",
            "霖霖",
            "2026-11-16",
            "09:00",
            "13:00",
            3,
            0,
            "台中私人包班",
            "歷史課程 (私人包班)",
            "已結束"
      ]
];
    realCourses.forEach(r => courseSheet.appendRow(r));
  }

  // 2. 建立「預約記錄」表
  let bookingSheet = ss.getSheetByName(SHEET_BOOKINGS);
  if (!bookingSheet) {
    bookingSheet = ss.insertSheet(SHEET_BOOKINGS);
    bookingSheet.appendRow(["預約編號", "報名時間", "課程編號", "課程名稱", "學員姓名", "聯絡電話", "Line ID", "備註需求", "簽到狀態"]);
    bookingSheet.getRange(1, 1, 1, 9).setBackground("#059669").setFontColor("#ffffff").setFontWeight("bold");
    
    const realBookings = [
      [
            "c4b7d942",
            "2026-08-06 14:40:44",
            "260420",
            "❷ 學習：頌缽入門",
            "林芸如林祐羽",
            "0954150015",
            "0954150015",
            "匯款末五碼: 42919 (早鳥價 $4000)",
            "已簽到"
      ],
      [
            "eeb8b647",
            "2026-08-21 22:37:17",
            "260428",
            "❷ 學習：頌缽入門",
            "黃玉雰",
            "0936407323",
            "Sabrina46",
            "匯款末五碼: 77454 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "be0ff703",
            "2026-08-21 23:24:15",
            "260428",
            "❷ 學習：頌缽入門",
            "龔湘珈",
            "0912294525",
            "0912294525",
            "Line 轉帳 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "cedeb1d6",
            "2026-09-25 14:19:06",
            "260428",
            "❷ 學習：頌缽入門",
            "蘇暉雯",
            "0982636628",
            "gladys0306",
            "匯款末五碼: 82580 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "964f647e",
            "2026-10-06 19:11:45",
            "260428",
            "❷ 學習：頌缽入門",
            "鍾美亭",
            "0931875265",
            "ed101562",
            "匯款末五碼: 03861 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "4511397f",
            "2026-10-19 22:31:02",
            "260438",
            "❸ 學習：頌缽進階",
            "龔湘珈",
            "0912294525",
            "湘珈",
            "匯款末五碼: 40892 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "dea0490e",
            "2026-10-19 22:32:06",
            "260438",
            "❸ 學習：頌缽進階",
            "黃玉雰",
            "0936407323",
            "Sabrina46",
            "匯款末五碼: 77454 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "07da1bfe",
            "2026-10-19 23:27:47",
            "260437",
            "❷ 學習：初進階整日班",
            "黃采靚",
            "0982612550",
            "采靚",
            "匯款末五碼: 51145 (早鳥價 $8000)",
            "已預約"
      ],
      [
            "9e7b450b",
            "2026-10-20 08:50:05",
            "260435",
            "❸ 學習：頌缽進階",
            "郭曉晴",
            "0988504928",
            "0978260217",
            "已匯款 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "d038e72d",
            "2026-10-20 09:39:20",
            "260438",
            "❸ 學習：頌缽進階",
            "林美伶",
            "0901358272",
            "luck123st@gmail.com",
            "匯款末五碼: 33453 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "4759c0e0",
            "2026-10-20 11:52:58",
            "260435",
            "❸ 學習：頌缽進階",
            "黃宜芝",
            "0938470388",
            "0938470388",
            "85390，要發薪水才有辦法匯款謝謝 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "f4da9546",
            "2026-10-20 17:55:16",
            "260437",
            "❷ 學習：初進階整日班",
            "龔榆善",
            "0938109513",
            "Christine Kung",
            "匯款末五碼: 40892 (早鳥價 $8000)",
            "已預約"
      ],
      [
            "c1fba9b8",
            "2026-10-20 17:58:00",
            "260437",
            "❷ 學習：初進階整日班",
            "龔湘珈",
            "0912294525",
            "湘珈",
            "匯款末五碼: 40892 (早鳥價 $8000)",
            "已預約"
      ],
      [
            "64880286",
            "2026-10-20 18:44:18",
            "260435",
            "❸ 學習：頌缽進階",
            "黃柔靜",
            "0919276049",
            "dollysheep0630",
            "匯款末五碼: 48570 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "13fb88c3",
            "2026-10-20 18:49:28",
            "260435",
            "❸ 學習：頌缽進階",
            "高廷瑜",
            "0908353114",
            "廷瑜",
            "匯款末五碼: 31888 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "7d191a05",
            "2026-10-20 20:54:32",
            "260438",
            "❸ 學習：頌缽進階",
            "許端容",
            "0918021609",
            "0918021609",
            "匯款末五碼: 057852 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "332b6593",
            "2026-10-20 23:02:42",
            "260435",
            "❸ 學習：頌缽進階",
            "曾馨儀",
            "0963031521",
            "lala8977",
            "匯款末五碼: 41042 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "bef1cbf0",
            "2026-10-21 12:15:44",
            "260435",
            "❸ 學習：頌缽進階",
            "柯盈如",
            "0989111361",
            "12345",
            "匯款末五碼: 28719 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "9c250b84",
            "2026-10-28 09:35:02",
            "260437",
            "❷ 學習：初進階整日班",
            "賴亭竹",
            "0912895272",
            "0610520",
            "匯款末五碼: 27653 (早鳥價 $8000)",
            "已預約"
      ],
      [
            "b8e9f5b8",
            "2026-10-28 09:36:47",
            "260437",
            "❷ 學習：初進階整日班",
            "賴宣含",
            "0926678885",
            "qq2535",
            "匯款末五碼: 27653 (早鳥價 $8000)",
            "已預約"
      ],
      [
            "0d227ecc",
            "2026-11-08 19:58:44",
            "260438",
            "❸ 學習：頌缽進階",
            "鍾美亭",
            "0931875265",
            "ed101562",
            "匯款末五碼: 98840 (早鳥價 $4000)",
            "已預約"
      ],
      [
            "04cf284f",
            "2026-8-6 00:00:00",
            "260423",
            "❷ 學習：頌缽入門",
            "黃婷纓",
            "0982114827",
            "h7038498",
            "匯款末五碼: 37055 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "849ec138",
            "2026-8-6 00:00:00",
            "260423",
            "❷ 學習：頌缽入門",
            "王禎均",
            "0958852104",
            "danyasandy",
            "匯款末五碼: 93400 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "54f60090",
            "2026-8-6 00:00:00",
            "260423",
            "❷ 學習：頌缽入門",
            "楊琬琳",
            "0963921247",
            "912543365",
            "匯款末五碼: 82704 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "8cec1bc8",
            "2026-8-14 00:00:00",
            "260423",
            "❷ 學習：頌缽入門",
            "鄭如潔",
            "0980221837",
            "darking1026@gmail.com",
            "匯款末五碼: 12737 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "0f76f89d",
            "2026-8-24 00:00:00",
            "260423",
            "❷ 學習：頌缽入門",
            "林沁倫",
            "0958352189",
            "bonniealun",
            "匯款末五碼: 3672 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "454f07a6",
            "2026-8-29 00:00:00",
            "260423",
            "❷ 學習：頌缽入門",
            "黃郁雯",
            "0987321235",
            "evn0613",
            "匯款末五碼: 90000 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "b11bd978",
            "2026-8-5 00:00:00",
            "260418",
            "❷ 學習：頌缽入門",
            "黃玉慈",
            "0980216410",
            "玉慈",
            "匯款末五碼: 97832 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "5f3997ef",
            "2026-8-5 00:00:00",
            "260418",
            "❷ 學習：頌缽入門",
            "郭香君",
            "0979090963",
            "香君",
            "匯款末五碼: 97832 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "d169f697",
            "2026-8-21 00:00:00",
            "260429",
            "❷ 學習：頌缽入門",
            "陳誱蔆",
            "0953306702",
            "lenglengchen",
            "匯款末五碼: 67114 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "31c8573f",
            "2026-8-22 00:00:00",
            "260429",
            "❷ 學習：頌缽入門",
            "高廷瑜",
            "0908353114",
            "廷瑜",
            "匯款末五碼: 31888 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "7ca468cb",
            "2026-8-30 00:00:00",
            "260429",
            "❷ 學習：頌缽入門",
            "曾馨儀",
            "0963031521",
            "lala8977",
            "匯款末五碼: 41042 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "07ce8906",
            "2026-9-2 00:00:00",
            "260429",
            "❷ 學習：頌缽入門",
            "黃宜芝",
            "0938470388",
            "938470388",
            "匯款末五碼: 85390 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "46475186",
            "2026-8-16 00:00:00",
            "260420",
            "❷ 學習：頌缽入門",
            "陳庭妤",
            "0963209360",
            "Anna0963",
            "匯款末五碼: 15671 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "ae20577b",
            "2026-8-16 00:00:00",
            "260420",
            "❷ 學習：頌缽入門",
            "謝紫薰",
            "0968623725",
            "e83912255",
            "匯款末五碼: 25671 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "5f8f42ab",
            "2026-8-16 00:00:00",
            "260420",
            "❷ 學習：頌缽入門",
            "雷妮妮",
            "0963209360",
            "lfenny",
            "匯款末五碼: 25671 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "4de7956d",
            "2026-8-19 00:00:00",
            "260420",
            "❷ 學習：頌缽入門",
            "黃柔靜",
            "0919276049",
            "dollysheep0630",
            "匯款末五碼: 77128 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "da1df761",
            "2026-8-20 00:00:00",
            "260420",
            "❷ 學習：頌缽入門",
            "郭芳瑜",
            "0965105619",
            "moment0129",
            "匯款末五碼: 78810 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "77694ec6",
            "2026-8-24 00:00:00",
            "260420",
            "❷ 學習：頌缽入門",
            "黃秋華",
            "0987996302",
            "987996302",
            "匯款末五碼: 0 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "fc51d13b",
            "2026-10-5 00:00:00",
            "260420",
            "❷ 學習：頌缽入門",
            "龔珊靚",
            "0918558227",
            "kung0918558227",
            "匯款末五碼: 50714 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "239db794",
            "2026-10-7 00:00:00",
            "260420",
            "❷ 學習：頌缽入門",
            "潘鈺林",
            "0976529678",
            "orangei3553",
            "匯款末五碼: 8204 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "c6039f71",
            "2026-10-8 00:00:00",
            "260420",
            "❷ 學習：頌缽入門",
            "王怡文",
            "0963139722",
            "yiwen58",
            "匯款末五碼: 32556 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "bfa7dc64",
            "2026-7-28 00:00:00",
            "260421",
            "❸ 學習：頌缽進階",
            "莊雅薇",
            "0922751230",
            "love520131411",
            "(待匯款） (早鳥價 $4000)",
            "已完成"
      ],
      [
            "4094cd98",
            "2026-7-29 00:00:00",
            "260421",
            "❸ 學習：頌缽進階",
            "莊淑媚",
            "0975330310",
            "hello0310",
            "匯款末五碼: 24351 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "2552c1ca",
            "2026-8-1 00:00:00",
            "260421",
            "❸ 學習：頌缽進階",
            "郭香君",
            "0979090963",
            "香君",
            "匯款末五碼: 97832 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "4be443e9",
            "2026-8-16 00:00:00",
            "260421",
            "❸ 學習：頌缽進階",
            "陳庭妤",
            "0963209360",
            "Anna0963",
            "匯款末五碼: 25671 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "b079fe94",
            "2026-8-24 00:00:00",
            "260421",
            "❸ 學習：頌缽進階",
            "黃秋華",
            "0987996302",
            "987996302",
            "匯款末五碼: 0 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "6a037a5f",
            "2026-10-9 00:00:00",
            "260434",
            "❹ 團體頌缽音樂療癒",
            "郭香君",
            "0979090963",
            "shine_kuo941022",
            "匯款末五碼: 97832 (早鳥價 $800)",
            "已完成"
      ],
      [
            "ba07c3e9",
            "2026-10-11 00:00:00",
            "260431",
            "❷ 學習：頌缽入門",
            "廖靜彗",
            "0972619011",
            "sagchacity",
            "匯款末五碼: 51774 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "9878dd24",
            "2026-10-11 00:00:00",
            "260431",
            "❷ 學習：頌缽入門",
            "劉盈初",
            "0963048287",
            "Liuyingchu",
            "匯款末五碼: 51774 (早鳥價 $4000)",
            "已完成"
      ],
      [
            "5a7435f5",
            "2026-10-11 00:00:00",
            "260431",
            "❷ 學習：頌缽入門",
            "吳婉卉",
            "0953699936",
            "helloparty",
            "匯款末五碼: 98806 (早鳥價 $4000)",
            "已完成"
      ]
];
    realBookings.forEach(r => bookingSheet.appendRow(r));
  }

  // 3. 建立「系統設定」表
  let configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET_CONFIG);
    configSheet.appendRow(["設定項目", "設定值"]);
    configSheet.appendRow(["系統名稱", "霖霖的頌缽課程預約系統"]);
    configSheet.appendRow(["管理員密碼", "admin888"]);
    configSheet.appendRow(["預約提醒事項", "課程開始前 2 小時前取消不扣堂數，如額滿可隨時關注名額釋出。"]);
    configSheet.appendRow(["常用教室清單", "高雄-近高美館,台南安南醫院,台北場地"]);
    configSheet.appendRow(["通知總開關", "開啟"]);
    configSheet.appendRow(["學員LINE通知開關", "開啟"]);
    configSheet.appendRow(["TelegramBotToken", ""]);
    configSheet.appendRow(["TelegramChatId", ""]);
    configSheet.appendRow(["LineChannelAccessToken", ""]);
    configSheet.appendRow(["LineNotifyToken", ""]);
    configSheet.appendRow(["常用課程名稱選單", "❷ 學習：頌缽入門\n❸ 學習：頌缽進階\n❷ 學習：初進階整日班\n❹ 團體頌缽音樂療癒"]);
    configSheet.appendRow(["常用課程說明選單", "想學頌缽，照顧自己 (早鳥價 $4000 / 定價 $5000)\n想進階技能，運用頌缽 (早鳥價 $4000 / 定價 $5000)\n學習整套頌缽使用方式 (早鳥價 $8000 / 定價 $12000)\n音頻深層放鬆，舒緩身心壓力 (早鳥價 $800 / 定價 $1000)"]);
    configSheet.appendRow(["前台預設字體大小", "normal"]);
    configSheet.getRange(1, 1, 1, 2).setBackground("#475569").setFontColor("#ffffff").setFontWeight("bold");
  }

  return "初始化完成！已匯入全部課程與學員報名資料！";
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
        const realCount = bookingCounts[cId] || 0;
        const virtualCount = Number(c["虛擬報名數"]) || 0;
        const totalDisplayBooked = realCount + virtualCount;
        const isFull = totalDisplayBooked >= maxCap;
        const remaining = Math.max(0, maxCap - totalDisplayBooked);

        return {
          id: cId,
          title: c["課程名稱"],
          instructor: c["授課教練"],
          date: formatDate(c["日期"]),
          startTime: c["開始時間"],
          endTime: c["結束時間"],
          capacity: maxCap,
          bookedCount: totalDisplayBooked,
          realBookedCount: realCount,
          virtualCount: virtualCount,
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
      const virtualCount = Number(targetCourse["虛擬報名數"]) || 0;
      const maxCap = Number(targetCourse["人數上限"]) || 0;
      if ((currentBooked + virtualCount) >= maxCap) return jsonResponse({ success: false, message: "該課程已額滿！" });

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
        "待確認"
      ]);

      updateCourseCount(courseSheet, courseId, currentBooked + 1);

      // 觸發管理者即時通知 (Telegram / LINE)
      const config = getConfigMap(ss);
      if (String(config["通知總開關"] || "開啟") === "開啟") {
        const adminMsg = "🔔【收到新學員課程報名】\n" +
          "────────────\n" +
          "• 課程：" + targetCourse["課程名稱"] + "\n" +
          "• 日期：" + formatDate(targetCourse["日期"]) + " " + targetCourse["開始時間"] + "\n" +
          "• 學員：" + studentName + "\n" +
          "• 電話：" + studentPhone + "\n" +
          "• LINE ID：" + (lineId || "無") + "\n" +
          "• 備註/匯款：" + (notes || "無") + "\n" +
          "• 預約編號：" + bookingId + "\n" +
          "────────────\n" +
          "請至管理後台核對匯款後，點擊「確認入帳 (報名成功)」！";
        
        sendAdminNotification(config, adminMsg);
      }

      return jsonResponse({
        success: true,
        message: "預約已送出！請稍候管理員核對匯款款項，核對無誤後將更新為「報名成功」並發送通知！",
        bookingId: bookingId,
        status: "待確認",
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
          courseSheet.getRange(rowIdx, 12).setValue(Number(data.virtualCount) || 0);
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
        data.status || "開啟報名",
        Number(data.virtualCount) || 0
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

    // 管理端：確認匯款收到 -> 變更為「報名成功」，並可發送 LINE 通知
    if (action === "confirmPayment") {
      const bookingSheet = ss.getSheetByName(SHEET_BOOKINGS);
      const bookingRows = bookingSheet.getDataRange().getValues();
      const bookingId = data.bookingId;
      const sendStudentLine = data.sendStudentLine !== false; // 預設依前端設定
      
      let foundRow = -1;
      let targetBooking = null;

      for (let r = 1; r < bookingRows.length; r++) {
        if (String(bookingRows[r][0]) === bookingId) {
          foundRow = r + 1;
          targetBooking = {
            id: bookingRows[r][0],
            time: bookingRows[r][1],
            courseId: bookingRows[r][2],
            courseTitle: bookingRows[r][3],
            name: bookingRows[r][4],
            phone: bookingRows[r][5],
            lineId: bookingRows[r][6],
            notes: bookingRows[r][7],
            status: bookingRows[r][8]
          };
          break;
        }
      }

      if (foundRow === -1) {
        return jsonResponse({ success: false, message: "找不到該預約記錄" });
      }

      bookingSheet.getRange(foundRow, 9).setValue("報名成功");

      // 檢查是否發送通知至學員 LINE
      let notificationResult = "未發送";
      const config = getConfigMap(ss);
      const isStudentLineNotifyEnabled = String(config["學員LINE通知開關"] || "開啟") === "開啟";

      if (sendStudentLine && isStudentLineNotifyEnabled && targetBooking.lineId) {
        const studentMsg = "🎉【" + (config["系統名稱"] || "課程預約系統") + "】報名成功通知\n" +
          "────────────\n" +
          targetBooking.name + " 您好，已收到您的款項！\n" +
          "• 課程名稱：" + targetBooking.courseTitle + "\n" +
          "• 預約編號：" + targetBooking.id + "\n" +
          "• 報名狀態：【報名成功】\n" +
          "期待在課堂上與您相遇，請準時出席！";
        
        // 若有設定 LINE Channel Access Token，嘗試以 LINE Messaging API 發送 (若 lineId 為 UserId)
        if (config["LineChannelAccessToken"]) {
          try {
            sendLinePushMessage(config["LineChannelAccessToken"], targetBooking.lineId, studentMsg);
            notificationResult = "已透過 LINE API 發送";
          } catch(e) {
            notificationResult = "LINE API 發送失敗: " + e.toString();
          }
        } else {
          notificationResult = "可使用後台一鍵喚起 LINE 快速傳送";
        }
      }

      return jsonResponse({
        success: true,
        message: "已確認匯款入帳！狀態已更新為「報名成功」！",
        bookingId: bookingId,
        newStatus: "報名成功",
        notificationResult: notificationResult
      });
    }

    // 管理端：發送測試通知 (Telegram 或 LINE)
    if (action === "testNotification") {
      const type = data.type; // "telegram" 或 "line"
      const testMsg = "🚀【" + (config["系統名稱"] || "課程預約系統") + "】通知測試成功！\n當有新學員報名或報名成功時，您將會即時收到通知！";
      let result = "";
      if (type === "telegram") {
        const tgToken = data.token || config["TelegramBotToken"];
        const tgChat = data.chatId || config["TelegramChatId"];
        if (!tgToken || !tgChat) return jsonResponse({ success: false, message: "請先填寫 Telegram Bot Token 與 Chat ID！" });
        result = sendTelegramMessage(tgToken, tgChat, testMsg);
      } else if (type === "line") {
        const lineToken = data.token || config["LineNotifyToken"];
        if (!lineToken) return jsonResponse({ success: false, message: "請先填寫 LINE Token！" });
        result = sendLineNotify(lineToken, testMsg);
      }
      return jsonResponse({ success: true, message: "測試通知已發出！", result: result });
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
      const notifyEnabled = data.notifyEnabled;
      const studentLineNotify = data.studentLineNotify;
      const tgBotToken = data.tgBotToken;
      const tgChatId = data.tgChatId;
      const lineChannelAccessToken = data.lineChannelAccessToken;
      const lineNotifyToken = data.lineNotifyToken;

      const updateOrAddConfig = (k, v) => {
        if (v === undefined || v === null) return;
        for (let r = 1; r < rows.length; r++) {
          if (String(rows[r][0]).trim() === k) {
            configSheet.getRange(r + 1, 2).setValue(String(v));
            return;
          }
        }
        configSheet.appendRow([k, String(v)]);
      };

      if (newSysName) updateOrAddConfig("系統名稱", newSysName);
      if (newNotice) updateOrAddConfig("預約提醒事項", newNotice);
      if (newPass) updateOrAddConfig("管理員密碼", newPass);
      if (newClassrooms) updateOrAddConfig("常用教室清單", newClassrooms);
      if (notifyEnabled) updateOrAddConfig("通知總開關", notifyEnabled);
      if (studentLineNotify) updateOrAddConfig("學員LINE通知開關", studentLineNotify);
      if (tgBotToken !== undefined) updateOrAddConfig("TelegramBotToken", tgBotToken);
      if (tgChatId !== undefined) updateOrAddConfig("TelegramChatId", tgChatId);
      if (lineChannelAccessToken !== undefined) updateOrAddConfig("LineChannelAccessToken", lineChannelAccessToken);
      if (lineNotifyToken !== undefined) updateOrAddConfig("LineNotifyToken", lineNotifyToken);
      if (data.courseTitles !== undefined) updateOrAddConfig("常用課程名稱選單", data.courseTitles);
      if (data.courseNotes !== undefined) updateOrAddConfig("常用課程說明選單", data.courseNotes);
      if (data.defaultFontSize !== undefined) updateOrAddConfig("前台預設字體大小", data.defaultFontSize);

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

// ── 通知輔助函式 ───────────────────────────────────────────────────

function sendAdminNotification(config, messageText) {
  // 1. 若有設定 Telegram
  if (config["TelegramBotToken"] && config["TelegramChatId"]) {
    try {
      sendTelegramMessage(config["TelegramBotToken"], config["TelegramChatId"], messageText);
    } catch(err) {
      console.warn("Telegram 通知發送失敗：", err);
    }
  }

  // 2. 若有設定 LINE Notify
  if (config["LineNotifyToken"]) {
    try {
      sendLineNotify(config["LineNotifyToken"], messageText);
    } catch(err) {
      console.warn("LINE Notify 發送失敗：", err);
    }
  }
}

function sendTelegramMessage(token, chatId, text) {
  const url = "https://api.telegram.org/bot" + token + "/sendMessage";
  const payload = {
    chat_id: chatId,
    text: text
  };
  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  return response.getContentText();
}

function sendLineNotify(token, message) {
  const url = "https://notify-api.line.me/api/notify";
  const response = UrlFetchApp.fetch(url, {
    method: "post",
    headers: { "Authorization": "Bearer " + token },
    payload: { message: "\n" + message },
    muteHttpExceptions: true
  });
  return response.getContentText();
}

function sendLinePushMessage(channelAccessToken, toUserId, text) {
  const url = "https://api.line.me/v2/bot/message/push";
  const payload = {
    to: toUserId,
    messages: [{ type: "text", text: text }]
  };
  const response = UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Authorization": "Bearer " + channelAccessToken,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  return response.getContentText();
}
