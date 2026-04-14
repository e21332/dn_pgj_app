function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = [
    {
      name: "設定",
      headers: ["管理者名稱", "管理者密碼", "更新日期"]
    },
    {
      name: "聯絡人",
      headers: ["聯絡人ID", "類型", "姓名", "聯絡人", "電話", "Email", "地址", "備註"]
    },
    {
      name: "樂器分類",
      headers: ["分類ID", "分類名稱", "備註"]
    },
    {
      name: "樂器庫存",
      headers: ["樂器ID", "分類ID", "名稱", "狀態", "月租金", "售價", "庫存數量", "備註"]
    },
    {
      name: "租用合約",
      headers: ["合約ID", "聯絡人ID", "樂器ID", "續租次數", "開始日", "結束日", "月租金", "繳款日", "合約狀態", "備註"]
    },
    {
      name: "押金記錄",
      headers: ["押金ID", "合約ID", "押金金額", "繳交日期", "退還日期", "退還狀態", "備註"]
    },
    {
      name: "租金繳款",
      headers: ["繳款ID", "合約ID", "期別", "應繳日", "實際繳款日", "金額", "狀態", "備註"]
    },
    {
      name: "銷售記錄",
      headers: ["銷售ID", "樂器ID", "聯絡人ID", "銷售日期", "售價", "備註"]
    },
    {
      name: "進貨記錄",
      headers: ["進貨ID", "樂器ID", "聯絡人ID", "進貨日期", "數量", "進貨成本", "備註"]
    },
    {
      name: "維修記錄",
      headers: ["維修ID", "樂器ID", "發生日期", "維修次數", "故障狀況", "維修費用", "維修狀態", "備註"]
    }
  ];

  sheets.forEach(({ name, headers }) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    // 只在第一列為空時才寫入標題
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground("#4a148c")
        .setFontColor("#ffffff")
        .setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
  });

  SpreadsheetApp.getUi().alert("✅ 所有工作表建立完成！");
}

function updateHeaderColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ["設定", "聯絡人", "樂器分類", "樂器庫存", "租用合約", "押金記錄", "租金繳款", "銷售記錄", "進貨記錄", "維修記錄"];

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;
    sheet.getRange(1, 1, 1, lastCol)
      .setBackground("#4a148c")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setFontSize(16);
  });

  SpreadsheetApp.getUi().alert("✅ 所有工作表標題列顏色已更新！");
}

function autoResizeAllHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const CHAR_WIDTH = 18; // 每個中文字的估算像素寬度
  const PADDING = 24;    // 左右留白

  ss.getSheets().forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    headers.forEach((header, i) => {
      const charCount = String(header).length;
      const width = charCount * CHAR_WIDTH + PADDING;
      sheet.setColumnWidth(i + 1, width);
    });
  });

  SpreadsheetApp.getUi().alert("✅ 所有工作表欄寬已依欄位名稱字數調整完成！");
}
