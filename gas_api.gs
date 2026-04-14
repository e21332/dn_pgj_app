// ===== 主入口 =====
function doGet(e) {
  const action = e.parameter.action;
  const params = e.parameter;
  let result;

  try {
    switch (action) {
      case "login":         result = login(params); break;
      case "dashboard":     result = getDashboard(); break;
      case "getCategories": result = getCategories(); break;
      case "saveCategory":  result = saveCategory(params); break;
      case "getInventory":  result = getInventory(params); break;
      case "saveInstrument":result = saveInstrument(params); break;
      case "getContacts":   result = getContacts(params); break;
      case "saveContact":   result = saveContact(params); break;
      case "getContracts":  result = getContracts(params); break;
      case "saveContract":  result = saveContract(params); break;
      case "endContract":   result = endContract(params); break;
      case "getDeposits":   result = getDeposits(); break;
      case "returnDeposit": result = returnDeposit(params); break;
      case "getPayments":   result = getPayments(params); break;
      case "savePayment":   result = savePayment(params); break;
      case "getSales":      result = getSales(params); break;
      case "saveSale":      result = saveSale(params); break;
      case "getPurchases":  result = getPurchases(params); break;
      case "savePurchase":  result = savePurchase(params); break;
      case "getRepairs":    result = getRepairs(params); break;
      case "saveRepair":    result = saveRepair(params); break;
      default:              result = { success: false, message: "未知的 action" };
    }
  } catch (err) {
    result = { success: false, message: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 工具函式 =====
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheet) {
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows
    .filter(r => r[0] !== "")
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
}

function generateId(prefix) {
  return prefix + "-" + new Date().getTime();
}

function formatDate(date) {
  if (!date) return "";
  const d = new Date(date);
  return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getDate()).padStart(2,"0")}`;
}

function today() {
  return new Date();
}

function diffDays(dateStr) {
  const d = new Date(dateStr);
  const now = today();
  d.setHours(0,0,0,0);
  now.setHours(0,0,0,0);
  return Math.floor((now - d) / 86400000);
}

// ===== 登入 =====
function login(params) {
  const sheet = getSheet("設定");
  const data = sheet.getDataRange().getValues();
  const storedPwd = data[1][1]; // 第2列第2欄為密碼
  return { success: params.password === String(storedPwd) };
}

// ===== 儀表板 =====
function getDashboard() {
  const now = today();
  const thisMonth = now.getMonth();
  const thisYear = now.getFullYear();

  const payments = sheetToObjects(getSheet("租金繳款"));
  const contracts = sheetToObjects(getSheet("租用合約"));
  const inventory = sheetToObjects(getSheet("樂器庫存"));
  const contacts = sheetToObjects(getSheet("聯絡人"));
  const repairs = sheetToObjects(getSheet("維修記錄"));
  const sales = sheetToObjects(getSheet("銷售記錄"));

  // 建立 ID 對照表
  const contactMap = {};
  contacts.forEach(c => contactMap[c["聯絡人ID"]] = c["姓名"]);
  const instrumentMap = {};
  inventory.forEach(i => instrumentMap[i["樂器ID"]] = i["名稱"]);
  const contractMap = {};
  contracts.forEach(c => contractMap[c["合約ID"]] = c);

  // 逾期未繳
  const overdue = payments
    .filter(p => p["狀態"] === "未繳" && diffDays(p["應繳日"]) > 0)
    .map(p => {
      const contract = contractMap[p["合約ID"]] || {};
      return {
        customer: contactMap[contract["聯絡人ID"]] || "-",
        instrument: instrumentMap[contract["樂器ID"]] || "-",
        days: diffDays(p["應繳日"]),
        amount: p["金額"]
      };
    });

  // 今日應繳
  const todayStr = now.toISOString().slice(0,10);
  const dueToday = payments
    .filter(p => p["狀態"] === "未繳" && String(p["應繳日"]).slice(0,10) === todayStr)
    .map(p => {
      const contract = contractMap[p["合約ID"]] || {};
      return {
        customer: contactMap[contract["聯絡人ID"]] || "-",
        instrument: instrumentMap[contract["樂器ID"]] || "-",
        amount: p["金額"]
      };
    });

  // 3天內租約到期
  const expiring = contracts
    .filter(c => c["合約狀態"] === "租用中")
    .filter(c => {
      const days = -diffDays(c["結束日"]);
      return days >= 0 && days <= 3;
    })
    .map(c => ({
      customer: contactMap[c["聯絡人ID"]] || "-",
      instrument: instrumentMap[c["樂器ID"]] || "-",
      days: -diffDays(c["結束日"]),
      renewCount: c["續租次數"] || "首次"
    }));

  // 維修中
  const repairing = repairs
    .filter(r => r["維修狀態"] === "維修中")
    .map(r => ({
      instrument: instrumentMap[r["樂器ID"]] || "-",
      date: formatDate(r["發生日期"]),
      issue: r["故障狀況"]
    }));

  // 即時統計
  const rentalIncome = payments
    .filter(p => p["狀態"] === "已繳清" && new Date(p["實際繳款日"]).getMonth() === thisMonth && new Date(p["實際繳款日"]).getFullYear() === thisYear)
    .reduce((sum, p) => sum + Number(p["金額"]), 0);

  const salesIncome = sales
    .filter(s => {
      const dateStr = String(s["銷售日期"]).replace(/\//g, "-").slice(0,7);
      const currentMonth = `${thisYear}-${String(thisMonth + 1).padStart(2, "0")}`;
      return dateStr === currentMonth;
    })
    .reduce((sum, s) => sum + Number(s["售價"]), 0);

  const renting = contracts.filter(c => c["合約狀態"] === "租用中").length;
  const stock = inventory.reduce((sum, i) => sum + Number(i["庫存數量"] || 0), 0);
  const repairCount = repairs.filter(r => r["維修狀態"] === "維修中").length;
  const newCustomers = contacts.filter(c => {
    // 以聯絡人ID的時間戳判斷本月新增
    const id = c["聯絡人ID"] || "";
    const ts = parseInt(id.split("-")[1]);
    if (!ts) return false;
    const d = new Date(ts);
    return d.getMonth() === thisMonth && d.getFullYear() === thisYear;
  }).length;

  // 上月統計
  const lastMonth = thisMonth === 0 ? 11 : thisMonth - 1;
  const lastYear = thisMonth === 0 ? thisYear - 1 : thisYear;

  const monthly = {
    rentalIncome: payments
      .filter(p => p["狀態"] === "已繳清" && new Date(p["實際繳款日"]).getMonth() === lastMonth && new Date(p["實際繳款日"]).getFullYear() === lastYear)
      .reduce((sum, p) => sum + Number(p["金額"]), 0),
    salesIncome: sales
      .filter(s => {
        const dateStr = String(s["銷售日期"]).replace(/\//g, "-").slice(0,7);
        const lastMonthStr = `${lastYear}-${String(lastMonth + 1).padStart(2, "0")}`;
        return dateStr === lastMonthStr;
      })
      .reduce((sum, s) => sum + Number(s["售價"]), 0),
    newCustomers: contacts.filter(c => {
      const id = c["聯絡人ID"] || "";
      const ts = parseInt(id.split("-")[1]);
      if (!ts) return false;
      const d = new Date(ts);
      return d.getMonth() === lastMonth && d.getFullYear() === lastYear;
    }).length,
    purchaseCost: sheetToObjects(getSheet("進貨記錄"))
      .filter(p => {
        const dateStr = String(p["進貨日期"]).replace(/\//g, "-").slice(0,7);
        const lastMonthStr = `${lastYear}-${String(lastMonth + 1).padStart(2, "0")}`;
        return dateStr === lastMonthStr;
      })
      .reduce((sum, p) => sum + Number(p["進貨成本"]), 0),
    repairCount: repairs
      .filter(r => new Date(r["發生日期"]).getMonth() === lastMonth && new Date(r["發生日期"]).getFullYear() === lastYear).length,
    newContracts: contracts
      .filter(c => {
        const id = c["合約ID"] || "";
        const ts = parseInt(id.split("-")[1]);
        if (!ts) return false;
        const d = new Date(ts);
        return d.getMonth() === lastMonth && d.getFullYear() === lastYear;
      }).length
  };

  return { success: true, overdue, dueToday, expiring, repairing, stats: { rentalIncome, salesIncome, renting, stock, repair: repairCount, newCustomers }, monthly };
}

// ===== 樂器分類 =====
function getCategories() {
  const data = sheetToObjects(getSheet("樂器分類"));
  return { success: true, data: data.map(r => ({ id: r["分類ID"], name: r["分類名稱"], note: r["備註"] })) };
}

function saveCategory(params) {
  const sheet = getSheet("樂器分類");
  const id = generateId("C");
  sheet.appendRow([id, params.name, params.note || ""]);
  return { success: true };
}

// ===== 樂器庫存 =====
function getInventory(params) {
  let data = sheetToObjects(getSheet("樂器庫存"));
  const categories = sheetToObjects(getSheet("樂器分類"));
  const catMap = {};
  categories.forEach(c => catMap[c["分類ID"]] = c["分類名稱"]);

  if (params.category) data = data.filter(r => r["分類ID"] === params.category);
  if (params.status)   data = data.filter(r => r["狀態"] === params.status);

  return {
    success: true,
    data: data.map(r => ({
      id: r["樂器ID"], category: catMap[r["分類ID"]] || r["分類ID"],
      name: r["名稱"], status: r["狀態"],
      rent: r["月租金"], price: r["售價"], stock: r["庫存數量"], note: r["備註"],
      canSell: Number(r["庫存數量"] || 0) > 0
    }))
  };
}

function saveInstrument(params) {
  const sheet = getSheet("樂器庫存");
  const id = generateId("I");
  sheet.appendRow([id, params.category, params.name, params.status, Number(params.rent)||0, Number(params.price)||0, Number(params.stock)||1, params.note||""]);
  return { success: true };
}

// ===== 聯絡人 =====
function getContacts(params) {
  let data = sheetToObjects(getSheet("聯絡人"));
  if (params.type) data = data.filter(r => String(r["類型"]).includes(params.type));
  if (params.name) data = data.filter(r => String(r["姓名"]).includes(params.name));
  return {
    success: true,
    data: data.map(r => ({
      id: r["聯絡人ID"], type: r["類型"], name: r["姓名"],
      person: r["聯絡人"], phone: r["電話"], email: r["Email"],
      address: r["地址"], note: r["備註"]
    }))
  };
}

function saveContact(params) {
  const sheet = getSheet("聯絡人");
  if (params.id) {
    // 編輯
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === params.id) {
        sheet.getRange(i+1, 1, 1, 8).setValues([[
          params.id, params.type, params.name, params.person||"",
          params.phone||"", params.email||"", params.address||"", params.note||""
        ]]);
        return { success: true };
      }
    }
  }
  // 新增
  const id = generateId("P");
  sheet.appendRow([id, params.type, params.name, params.person||"", params.phone||"", params.email||"", params.address||"", params.note||""]);
  return { success: true };
}

// ===== 租用合約 =====
function getContracts(params) {
  let data = sheetToObjects(getSheet("租用合約"));
  const contactMap = buildMap("聯絡人", "聯絡人ID", "姓名");
  const instrumentMap = buildMap("樂器庫存", "樂器ID", "名稱");

  if (params.status) data = data.filter(r => r["合約狀態"] === params.status);

  return {
    success: true,
    data: data.map(r => ({
      id: r["合約ID"],
      customer: contactMap[r["聯絡人ID"]] || "-",
      instrument: instrumentMap[r["樂器ID"]] || "-",
      renew: r["續租次數"] || "首次",
      start: formatDate(r["開始日"]),
      end: formatDate(r["結束日"]),
      rent: r["月租金"],
      dueDay: r["繳款日"],
      status: r["合約狀態"],
      note: r["備註"]
    }))
  };
}

function saveContract(params) {
  const sheet = getSheet("租用合約");
  const id = generateId("R");
  // 2026-04-10" to "2026/04/10"
  const formattedStart = params.start ? params.start.replace(/-/g, "/") : "";
  const formattedEnd = params.end ? params.end.replace(/-/g, "/") : "";
  sheet.appendRow([
    id, params.customer, params.instrument,
    params.renew || "首次", formattedStart, formattedEnd,
    Number(params.rent)||0, params.dueDay, "租用中", params.note||""
  ]);

  // 自動建立押金記錄
  if (Number(params.deposit) > 0) {
    const depSheet = getSheet("押金記錄");
    // 2026-04-10" to "2026/04/10"
    const formattedDepositDate = params.start ? params.start.replace(/-/g, "/") : "";
    depSheet.appendRow([generateId("DEP"), id, Number(params.deposit), formattedDepositDate, "", "持有中", ""]);
  }

  // 自動產生每月繳款期別
  generatePaymentSchedule(id, params.start, params.end, Number(params.rent)||0, Number(params.dueDay)||1);

  return { success: true };
}

function generatePaymentSchedule(contractId, startDate, endDate, rent, dueDay) {
  const sheet = getSheet("租金繳款");
  const start = new Date(startDate);
  const end = new Date(endDate);
  let period = 1;
  let current = new Date(start.getFullYear(), start.getMonth(), dueDay);
  if (current < start) current.setMonth(current.getMonth() + 1);

  while (current <= end) {
    const dueStr = `${current.getFullYear()}/${String(current.getMonth()+1).padStart(2,"0")}/${String(current.getDate()).padStart(2,"0")}`;
    sheet.appendRow([generateId("PAY"), contractId, `第${period}期`, dueStr, "", rent, "未繳", ""]);
    period++;
    current.setMonth(current.getMonth() + 1);
  }
}

function endContract(params) {
  const sheet = getSheet("租用合約");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === params.id) {
      sheet.getRange(i+1, 9).setValue("已結束");
      return { success: true };
    }
  }
  return { success: false };
}

// ===== 押金 =====
function getDeposits() {
  const data = sheetToObjects(getSheet("押金記錄"));
  const contractMap = buildObjectMap("租用合約", "合約ID");
  const contactMap = buildMap("聯絡人", "聯絡人ID", "姓名");

  return {
    success: true,
    data: data.map(r => {
      const contract = contractMap[r["合約ID"]] || {};
      return {
        id: r["押金ID"],
        customer: contactMap[contract["聯絡人ID"]] || "-",
        amount: r["押金金額"],
        paidDate: formatDate(r["繳交日期"]),
        returnDate: formatDate(r["退還日期"]),
        status: r["退還狀態"]
      };
    })
  };
}

function returnDeposit(params) {
  const sheet = getSheet("押金記錄");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === params.id) {
      sheet.getRange(i+1, 5).setValue(params.date);
      sheet.getRange(i+1, 6).setValue("已退還");
      return { success: true };
    }
  }
  return { success: false };
}

// ===== 租金繳款 =====
function getPayments(params) {
  let data = sheetToObjects(getSheet("租金繳款"));
  const contractMap = buildObjectMap("租用合約", "合約ID");
  const contactMap = buildMap("聯絡人", "聯絡人ID", "姓名");
  const instrumentMap = buildMap("樂器庫存", "樂器ID", "名稱");

  // 自動標記逾期
  data.forEach(p => {
    if (p["狀態"] === "未繳" && diffDays(p["應繳日"]) > 0) p["狀態"] = "逾期";
  });

  if (params.status) data = data.filter(r => r["狀態"] === params.status);

  return {
    success: true,
    data: data.map(r => {
      const contract = contractMap[r["合約ID"]] || {};
      return {
        id: r["繳款ID"],
        customer: contactMap[contract["聯絡人ID"]] || "-",
        instrument: instrumentMap[contract["樂器ID"]] || "-",
        period: r["期別"],
        dueDate: formatDate(r["應繳日"]),
        amount: r["金額"],
        status: r["狀態"]
      };
    })
  };
}

function savePayment(params) {
  const sheet = getSheet("租金繳款");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === params.id) {
      sheet.getRange(i+1, 5).setValue(params.date);
      sheet.getRange(i+1, 7).setValue("已繳清");
      return { success: true };
    }
  }
  return { success: false };
}

// ===== 銷售 =====
function getSales(params) {
  let data = sheetToObjects(getSheet("銷售記錄"));
  const instrumentMap = buildMap("樂器庫存", "樂器ID", "名稱");
  const contactMap = buildMap("聯絡人", "聯絡人ID", "姓名");

  if (params.month) {
    data = data.filter(r => {
      const dateStr = String(r["銷售日期"]);
      return dateStr.replace(/\//g, "-").slice(0,7) === params.month;
    });
  }

  return {
    success: true,
    data: data.map(r => ({
      id: r["銷售ID"],
      instrument: instrumentMap[r["樂器ID"]] || "-",
      customer: contactMap[r["聯絡人ID"]] || "",
      date: formatDate(r["銷售日期"]),
      price: r["售價"],
      note: r["備註"]
    }))
  };
}

function saveSale(params) {
  try {
    const sheet = getSheet("銷售記錄");
    const id = generateId("S");
    const now = new Date();
    const nextRow = sheet.getLastRow() + 1;

    sheet.getRange(nextRow, 1).setValue(id);
    sheet.getRange(nextRow, 2).setValue(params.instrument || "");
    sheet.getRange(nextRow, 3).setValue(params.customer || "");
    sheet.getRange(nextRow, 4).setValue(now).setNumberFormat("yyyy/MM/dd");
    sheet.getRange(nextRow, 5).setValue(Number(params.price) || 0);
    sheet.getRange(nextRow, 6).setValue(params.note || "");
    
    return {
      success: true,
      id: id,
      date: Utilities.formatDate(
        now,
        SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || "Asia/Taipei",
        "yyyy/MM/dd"
      )
    };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ===== 進貨 =====
function getPurchases(params) {
  let data = sheetToObjects(getSheet("進貨記錄"));
  const instrumentMap = buildMap("樂器庫存", "樂器ID", "名稱");
  const contactMap = buildMap("聯絡人", "聯絡人ID", "姓名");

  if (params.month) {
    data = data.filter(r => {
      const dateStr = String(r["進貨日期"]);
      return dateStr.replace(/\//g, "-").slice(0,7) === params.month;
    });
  }

  return {
    success: true,
    data: data.map(r => ({
      id: r["進貨ID"],
      instrument: instrumentMap[r["樂器ID"]] || "-",
      supplier: contactMap[r["聯絡人ID"]] || "-",
      date: formatDate(r["進貨日期"]),
      qty: r["數量"],
      cost: r["進貨成本"],
      note: r["備註"]
    }))
  };
}

function savePurchase(params) {
  const sheet = getSheet("進貨記錄");
  const id = generateId("PUR");
  const formattedDate = params.date ? params.date.replace(/-/g, "/") : "";
  sheet.appendRow([id, params.instrument, params.supplier||"", formattedDate, Number(params.qty)||1, Number(params.cost)||0, params.note||""]);

  // 更新庫存數量
  const invSheet = getSheet("樂器庫存");
  const invData = invSheet.getDataRange().getValues();
  for (let i = 1; i < invData.length; i++) {
    if (invData[i][0] === params.instrument) {
      const current = Number(invData[i][6]) || 0;
      invSheet.getRange(i+1, 7).setValue(current + Number(params.qty));
      break;
    }
  }
  return { success: true };
}

// ===== 維修 =====
function getRepairs(params) {
  let data = sheetToObjects(getSheet("維修記錄"));
  const instrumentMap = buildObjectMap("樂器庫存", "樂器ID");

  if (params.status) data = data.filter(r => r["維修狀態"] === params.status);

  return {
    success: true,
    data: data.map(r => {
      const inst = instrumentMap[r["樂器ID"]] || {};
      return {
        id: r["維修ID"],
        instrumentId: r["樂器ID"],
        instrument: inst["名稱"] || "-",
        date: formatDate(r["發生日期"]),
        count: r["維修次數"],
        issue: r["故障狀況"],
        cost: r["維修費用"],
        status: r["維修狀態"],
        note: r["備註"]
      };
    })
  };
}

function saveRepair(params) {
  const sheet = getSheet("維修記錄");
  // 2026-04-10" to "2026/04/10"
  const formattedDate = params.date ? params.date.replace(/-/g, "/") : "";

  if (params.id) {
    // 編輯
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === params.id) {
        sheet.getRange(i+1, 1, 1, 8).setValues([[
          params.id, params.instrument, formattedDate,
          data[i][3], params.issue, Number(params.cost)||0,
          params.status, params.note||""
        ]]);
        return { success: true };
      }
    }
  }

  // 新增：計算該樂器第幾次維修
  const all = sheetToObjects(sheet);
  const count = all.filter(r => r["樂器ID"] === params.instrument).length + 1;
  const id = generateId("REP");
  sheet.appendRow([id, params.instrument, formattedDate, count, params.issue, Number(params.cost)||0, params.status, params.note||""]);
  return { success: true };
}

// ===== 共用對照表工具 =====
function buildMap(sheetName, keyCol, valueCol) {
  const map = {};
  sheetToObjects(getSheet(sheetName)).forEach(r => map[r[keyCol]] = r[valueCol]);
  return map;
}

function buildObjectMap(sheetName, keyCol) {
  const map = {};
  sheetToObjects(getSheet(sheetName)).forEach(r => map[r[keyCol]] = r);
  return map;
}
