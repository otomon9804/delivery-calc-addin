/* global document, Excel, Office */

const SECRET = "DC2026SECRET";
const DAY_NAMES = ["日", "月", "火", "水", "木", "金", "土"];

const DEFAULT_CONFIG = {
  calcMethod: "曜日",
  order: { sheetName: "受注一覧", customerCodeCol: 1, productCodeCol: 6, makerCodeCol: 9, logisticsCol: 10, arrivalCol: 11, deliveryCol: 12 },
  sheetNames: { maker: "TBL_MAKER", customer: "TBL_CUSTOMER", warehouse: "TBL_WAREHOUSE", holiday: "TBL_HOLIDAY", itemIrr: "TBL_ITEM_IRREGULAR", makerIrr: "TBL_MAKER_IRREGULAR" }
};

let cfg = JSON.parse(JSON.stringify(DEFAULT_CONFIG));

async function sha256(message) {
  const msgBuffer = new TextEncoder().encode(message);
  const hashBuffer = await crypto.subtle.digest("SHA-256", msgBuffer);
  return Array.from(new Uint8Array(hashBuffer)).map(b => b.toString(16).padStart(2, "0")).join("");
}

async function generateKey(email, yyyymm) {
  const raw = await sha256(email.toLowerCase() + yyyymm + SECRET);
  const chars = raw.toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 12);
  return "DC-" + chars.slice(0, 4) + "-" + chars.slice(4, 8) + "-" + chars.slice(8, 12);
}

async function validateLicense(email, key) {
  const now = new Date();
  for (let delta = 0; delta <= 2; delta++) {
    const d = new Date(now.getFullYear(), now.getMonth() - delta, 1);
    const yyyymm = d.getFullYear() + String(d.getMonth() + 1).padStart(2, "0");
    if (key.toUpperCase() === await generateKey(email, yyyymm)) return true;
  }
  return false;
}

async function activateLicense() {
  const msg = document.getElementById("msg-license");
  const email = document.getElementById("license-email").value.trim();
  const key = document.getElementById("license-key").value.trim();
  if (!email || !key) { msg.textContent = "❌ メールアドレスとライセンスキーを入力してください"; return; }
  msg.textContent = "認証中...";
  try {
    if (await validateLicense(email, key)) {
      await saveLicense(email, key);
      const license = await loadLicense();
      showApp(license);
    } else {
      msg.textContent = "❌ ライセンスキーが無効です";
    }
  } catch(e) {
    msg.textContent = "❌ エラー: " + e.message;
  }
}

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    try {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("btn-activate").onclick = activateLicense;
      document.getElementById("btn-logout").onclick = logoutLicense;
      document.getElementById("btn-init-master").onclick = initMaster;
      document.getElementById("btn-calc-delivery").onclick = calcDelivery;
      document.getElementById("btn-save-settings").onclick = saveSettings;
      document.getElementById("btn-reset-settings").onclick = resetSettings;
      document.getElementById("btn-toggle-settings").onclick = toggleSettings;
      document.getElementById("btn-method-weekday").onclick = () => setCalcMethod("曜日");
      document.getElementById("btn-method-days").onclick = () => setCalcMethod("日数");
      await loadConfig();
      updateCalcMethodUI();
      const license = await loadLicense();
      if (license && await validateLicense(license.email, license.key)) {
        showApp(license);
      } else {
        showLicenseScreen();
      }
    } catch(e) {
      document.getElementById("sideload-msg").style.display = "block";
      document.getElementById("sideload-msg").textContent = "エラー：" + e.message;
    }
  }
});

function showLicenseScreen() {
  document.getElementById("license-screen").style.display = "block";
  document.getElementById("app-body").style.display = "none";
}

function showApp(license) {
  document.getElementById("license-screen").style.display = "none";
  document.getElementById("app-body").style.display = "block";
  document.getElementById("license-info").innerHTML = "✅ 認証済み：" + license.email + "<br>有効期限：" + license.expiry;
}

async function saveLicense(email, key) {
  await Excel.run(async (context) => {
    const props = context.workbook.properties.custom;
    try { props.getItemOrNullObject("DC_EMAIL").delete(); } catch(e) {}
    try { props.getItemOrNullObject("DC_KEY").delete(); } catch(e) {}
    await context.sync();
    props.add("DC_EMAIL", email);
    props.add("DC_KEY", key);
    await context.sync();
  });
}

async function loadLicense() {
  try {
    return await Excel.run(async (context) => {
      const props = context.workbook.properties.custom;
      const ep = props.getItemOrNullObject("DC_EMAIL");
      const kp = props.getItemOrNullObject("DC_KEY");
      ep.load("value"); kp.load("value");
      await context.sync();
      if (ep.isNullObject || kp.isNullObject) return null;
      const now = new Date();
      const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
      const expiry = now.getFullYear() + "/" + String(now.getMonth() + 1).padStart(2, "0") + "/" + String(lastDay.getDate()).padStart(2, "0");
      return { email: ep.value, key: kp.value, expiry };
    });
  } catch (e) { return null; }
}

async function logoutLicense() {
  await Excel.run(async (context) => {
    const props = context.workbook.properties.custom;
    props.getItemOrNullObject("DC_EMAIL").delete();
    props.getItemOrNullObject("DC_KEY").delete();
    await context.sync();
  });
  showLicenseScreen();
}

async function saveConfig() {
  try {
    await Excel.run(async (context) => {
      const props = context.workbook.properties.custom;
      const existing = props.getItemOrNullObject("DC_CONFIG");
      existing.load("value");
      await context.sync();
      if (!existing.isNullObject) { existing.delete(); await context.sync(); }
      props.add("DC_CONFIG", JSON.stringify(cfg));
      await context.sync();
    });
  } catch (e) { console.error("設定保存エラー", e); }
}

async function loadConfig() {
  try {
    await Excel.run(async (context) => {
      const props = context.workbook.properties.custom;
      const item = props.getItemOrNullObject("DC_CONFIG");
      item.load("value");
      await context.sync();
      if (!item.isNullObject && item.value) {
        const saved = JSON.parse(item.value);
        cfg.calcMethod = saved.calcMethod || DEFAULT_CONFIG.calcMethod;
        cfg.order = Object.assign(JSON.parse(JSON.stringify(DEFAULT_CONFIG.order)), saved.order || {});
        cfg.sheetNames = Object.assign(JSON.parse(JSON.stringify(DEFAULT_CONFIG.sheetNames)), saved.sheetNames || {});
      }
    });
  } catch (e) { cfg = JSON.parse(JSON.stringify(DEFAULT_CONFIG)); }
}

function updateCalcMethodUI() {
  const isWeekday = cfg.calcMethod === "曜日";
  document.getElementById("btn-method-weekday").classList.toggle("method-active", isWeekday);
  document.getElementById("btn-method-days").classList.toggle("method-active", !isWeekday);
  document.getElementById("method-label").textContent = "演算方法：" + cfg.calcMethod + "モード";
}

async function setCalcMethod(method) {
  cfg.calcMethod = method;
  updateCalcMethodUI();
  await saveConfig();
}

function toggleSettings() {
  const panel = document.getElementById("settings-panel");
  const btn = document.getElementById("btn-toggle-settings");
  const isHidden = panel.style.display === "none" || panel.style.display === "";
  panel.style.display = isHidden ? "block" : "none";
  btn.textContent = isHidden ? "▲ 設定を閉じる" : "⚙️ 列・シート設定を開く";
  if (isHidden) populateSettings();
}

function populateSettings() {
  const o = cfg.order, s = cfg.sheetNames;
  const fields = {
    "s-order-sheet": o.sheetName, "s-order-customer-col": o.customerCodeCol,
    "s-order-product-col": o.productCodeCol, "s-order-maker-col": o.makerCodeCol,
    "s-order-logistics-col": o.logisticsCol, "s-order-arrival-col": o.arrivalCol,
    "s-order-delivery-col": o.deliveryCol, "s-maker-sheet": s.maker,
    "s-customer-sheet": s.customer, "s-warehouse-sheet": s.warehouse,
    "s-holiday-sheet": s.holiday, "s-item-irr-sheet": s.itemIrr, "s-maker-irr-sheet": s.makerIrr
  };
  for (const [id, val] of Object.entries(fields)) { const el = document.getElementById(id); if (el) el.value = val; }
}

async function saveSettings() {
  cfg.order.sheetName = document.getElementById("s-order-sheet").value.trim();
  cfg.order.customerCodeCol = parseInt(document.getElementById("s-order-customer-col").value);
  cfg.order.productCodeCol = parseInt(document.getElementById("s-order-product-col").value);
  cfg.order.makerCodeCol = parseInt(document.getElementById("s-order-maker-col").value);
  cfg.order.logisticsCol = parseInt(document.getElementById("s-order-logistics-col").value);
  cfg.order.arrivalCol = parseInt(document.getElementById("s-order-arrival-col").value);
  cfg.order.deliveryCol = parseInt(document.getElementById("s-order-delivery-col").value);
  cfg.sheetNames.maker = document.getElementById("s-maker-sheet").value.trim();
  cfg.sheetNames.customer = document.getElementById("s-customer-sheet").value.trim();
  cfg.sheetNames.warehouse = document.getElementById("s-warehouse-sheet").value.trim();
  cfg.sheetNames.holiday = document.getElementById("s-holiday-sheet").value.trim();
  cfg.sheetNames.itemIrr = document.getElementById("s-item-irr-sheet").value.trim();
  cfg.sheetNames.makerIrr = document.getElementById("s-maker-irr-sheet").value.trim();
  await saveConfig();
  showMsg("msg-settings", "✅ 設定を保存しました");
}

async function resetSettings() {
  cfg = JSON.parse(JSON.stringify(DEFAULT_CONFIG));
  await saveConfig();
  populateSettings();
  updateCalcMethodUI();
  showMsg("msg-settings", "✅ デフォルトに戻しました");
}

function showMsg(id, text, ms = 3000) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = text;
  setTimeout(() => { el.textContent = ""; }, ms);
}

async function initMaster() {
  const msg = document.getElementById("msg-init");
  msg.textContent = "作成中...";
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const isWeekday = cfg.calcMethod === "曜日";
      const sn = cfg.sheetNames;
      const irrHeaders = isWeekday ? ["得意先コード","得意先名","コード","名称","月","火","水","木","金","土","日"] : ["得意先コード","得意先名","コード","名称","日数（営業日）"];
      const masterDefs = [
        { name: sn.maker, headers: isWeekday ? ["発注先コード","メーカー名称","受注締め時間","月→納品","火→納品","水→納品","木→納品","金→納品","旗日明け加算"] : ["発注先コード","メーカー名称","受注締め時間","リードタイム（営業日）","旗日明け加算"], data: [] },
        { name: sn.customer, headers: ["得意先コード","得意先名称","月","火","水","木","金","土","日"], data: [] },
        { name: sn.warehouse, headers: ["区分","月","火","水","木","金","土","日","祝日"], data: [["入荷","〇","〇","〇","〇","〇","〇","×","×"],["納品","〇","〇","〇","〇","〇","×","×","×"]] },
        { name: sn.holiday, headers: ["日付","祝日名"], data: [] },
        { name: sn.itemIrr, headers: irrHeaders, data: [] },
        { name: sn.makerIrr, headers: irrHeaders, data: [] }
      ];
      for (const master of masterDefs) {
        const existing = sheets.getItemOrNullObject(master.name);
        await context.sync();
        if (!existing.isNullObject) existing.delete();
        const sheet = sheets.add(master.name);
        const headerRange = sheet.getRangeByIndexes(0, 0, 1, master.headers.length);
        headerRange.values = [master.headers];
        headerRange.format.fill.color = "#217346";
        headerRange.format.font.color = "white";
        headerRange.format.font.bold = true;
        headerRange.format.columnWidth = 120;
        if (master.data.length > 0) sheet.getRangeByIndexes(1, 0, master.data.length, master.headers.length).values = master.data;
      }
      await context.sync();
      msg.textContent = "✅ マスタシートを作成しました！";
    });
  } catch (e) { msg.textContent = "❌ エラー：" + e.message; }
}

function excelDateToDate(serial) { return new Date(Math.floor(serial - 25569) * 86400 * 1000); }
function formatDate(date) { return date.getUTCFullYear() + "/" + String(date.getUTCMonth() + 1).padStart(2, "0") + "/" + String(date.getUTCDate()).padStart(2, "0"); }
async function loadHolidays(context) { const sheet = context.workbook.worksheets.getItem(cfg.sheetNames.holiday); const range = sheet.getUsedRange(); range.load("values"); await context.sync(); return range.values.slice(1).filter(r => r[0]).map(r => formatDate(excelDateToDate(r[0]))); }
async function loadMakers(context) { const sheet = context.workbook.worksheets.getItem(cfg.sheetNames.maker); const range = sheet.getUsedRange(); range.load("values"); await context.sync(); const isWeekday = cfg.calcMethod === "曜日"; return range.values.slice(1).filter(r => r[0]).map(r => { const maker = { code: String(r[0]), name: r[1], deadline: r[2], holidayAdd: isWeekday ? (Number(r[8]) || 0) : (Number(r[4]) || 0) }; if (isWeekday) maker.deliveryMap = { "月": r[3], "火": r[4], "水": r[5], "木": r[6], "金": r[7] }; else maker.days = Number(r[3]) || 1; return maker; }); }
async function loadCustomers(context) { const sheet = context.workbook.worksheets.getItem(cfg.sheetNames.customer); const range = sheet.getUsedRange(); range.load("values"); await context.sync(); return range.values.slice(1).filter(r => r[0]).map(r => ({ code: String(r[0]), name: r[1], deliverable: { "月": r[2]==="〇", "火": r[3]==="〇", "水": r[4]==="〇", "木": r[5]==="〇", "金": r[6]==="〇", "土": r[7]==="〇", "日": r[8]==="〇" } })); }
async function loadWarehouse(context) { const sheet = context.workbook.worksheets.getItem(cfg.sheetNames.warehouse); const range = sheet.getUsedRange(); range.load("values"); await context.sync(); const wh = {}; range.values.slice(1).forEach(r => { wh[r[0]] = { "月": r[1]==="〇", "火": r[2]==="〇", "水": r[3]==="〇", "木": r[4]==="〇", "金": r[5]==="〇", "土": r[6]==="〇", "日": r[7]==="〇", "祝日": r[8]==="〇" }; }); return wh; }
async function loadIrregulars(context, sheetName, keyField) { try { const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName); await context.sync(); if (sheet.isNullObject) return []; const range = sheet.getUsedRange(); range.load("values"); await context.sync(); const isWeekday = cfg.calcMethod === "曜日"; return range.values.slice(1).filter(r => r[0] && r[2]).map(r => { const entry = { customerCode: String(r[0]), [keyField]: String(r[2]) }; if (isWeekday) entry.deliverable = { "月": r[4]==="〇", "火": r[5]==="〇", "水": r[6]==="〇", "木": r[7]==="〇", "金": r[8]==="〇", "土": r[9]==="〇", "日": r[10]==="〇" }; else entry.days = Number(r[4]) || 1; return entry; }); } catch (e) { return []; } }
async function calcDelivery() { const msg = document.getElementById("msg-delivery"); const resultBox = document.getElementById("result-box"); const resultDetail = document.getElementById("result-detail"); resultBox.style.display = "none"; msg.innerHTML = "計算中..."; const now = new Date(); const o = cfg.order; try { await Excel.run(async (context) => { const holidays = await loadHolidays(context); const makers = await loadMakers(context); const customers = await loadCustomers(context); const warehouse = await loadWarehouse(context); const itemIrrs = await loadIrregulars(context, cfg.sheetNames.itemIrr, "productCode"); const makerIrrs = await loadIrregulars(context, cfg.sheetNames.makerIrr, "makerCode"); const orderSheet = context.workbook.worksheets.getItem(o.sheetName); const orderRange = orderSheet.getUsedRange(); orderRange.load("values"); await context.sync(); const rows = orderRange.values; let successCount = 0, errorCount = 0; const errorLog = []; for (let i = 1; i < rows.length; i++) { const row = rows[i]; const customerCode = String(row[o.customerCodeCol - 1] || "").trim(); if (!customerCode) continue; const productCode = String(row[o.productCodeCol - 1] || "").trim(); const makerCode = String(row[o.makerCodeCol - 1] || "").trim(); const logistics = String(row[o.logisticsCol - 1] || "").trim(); const customer = customers.find(c => c.code === customerCode); const maker = makers.find(m => m.code === makerCode); if (!customer) { errorLog.push((i+1)+"行目：得意先「"+customerCode+"」が見つかりません"); errorCount++; orderSheet.getRangeByIndexes(i,0,1,o.deliveryCol).format.fill.color="#FFCCCC"; continue; } if (logistics !== "在庫" && !maker) { errorLog.push((i+1)+"行目：発注先「"+makerCode+"」が見つかりません"); errorCount++; orderSheet.getRangeByIndexes(i,0,1,o.deliveryCol).format.fill.color="#FFCCCC"; continue; } try { const itemIrr = itemIrrs.find(e => e.customerCode === customerCode && e.productCode === productCode); const makerIrr = makerIrrs.find(e => e.customerCode === customerCode && e.makerCode === makerCode); const irr = itemIrr || makerIrr || null; let deliveryDate = null, arrivalDate = null; const isWeekdayMode = cfg.calcMethod === "曜日"; if (isWeekdayMode) { if (logistics === "在庫") deliveryDate = calcZaikoWeekday(now, customer, warehouse, holidays, irr); else if (logistics === "通過") { const r = calcTsuukaWeekday(now, maker, customer, warehouse, holidays, irr); deliveryDate = r.deliveryDate; arrivalDate = r.arrivalDate; } else if (logistics === "直送") deliveryDate = calcChokusoWeekday(now, maker, customer, holidays, irr); else throw new Error("物流形態「"+logistics+"」が不正です"); } else { if (logistics === "在庫") deliveryDate = calcZaikoDays(now, customer, warehouse, holidays, irr); else if (logistics === "通過") { const r = calcTsuukaDays(now, maker, customer, warehouse, holidays, irr); deliveryDate = r.deliveryDate; arrivalDate = r.arrivalDate; } else if (logistics === "直送") deliveryDate = calcChokusoDays(now, maker, customer, holidays, irr); else throw new Error("物流形態「"+logistics+"」が不正です"); } orderSheet.getRangeByIndexes(i,0,1,o.deliveryCol).format.fill.clear(); orderSheet.getRangeByIndexes(i,o.deliveryCol-1,1,1).values = [[formatDate(deliveryDate)]]; if (arrivalDate) orderSheet.getRangeByIndexes(i,o.arrivalCol-1,1,1).values = [[formatDate(arrivalDate)]]; successCount++; } catch (rowErr) { errorLog.push((i+1)+"行目："+rowErr.message); errorCount++; orderSheet.getRangeByIndexes(i,0,1,o.deliveryCol).format.fill.color="#FFCCCC"; } } await context.sync(); let html = "✅ "+successCount+"件完了"; if (errorCount > 0) html += "　❌ "+errorCount+"件エラー<br><br>"+errorLog.join("<br>"); resultDetail.innerHTML = html; resultBox.style.display = "block"; msg.innerHTML = ""; }); } catch (e) { msg.innerHTML = "❌ エラー：" + e.message; } }
function isBusinessDay(date, holidays) { const d = date.getUTCDay(); return d !== 0 && d !== 6 && !holidays.includes(formatDate(date)); }
function nextBusinessDay(date, holidays) { let next = new Date(date); next.setUTCDate(next.getUTCDate()+1); while (!isBusinessDay(next, holidays)) next.setUTCDate(next.getUTCDate()+1); return next; }
function addBusinessDays(date, days, holidays) { let d = new Date(date), added = 0; while (added < days) { d.setUTCDate(d.getUTCDate()+1); if (isBusinessDay(d, holidays)) added++; } return d; }
function getBaseDate(now, deadlineStr, holidays) { let base = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())); if (deadlineStr) { const parts = deadlineStr.toString().split(":"); const dh = Number(parts[0]), dm = Number(parts[1]||0); if (now.getHours() > dh || (now.getHours() === dh && now.getMinutes() >= dm)) base = nextBusinessDay(base, holidays); } return base; }
function getNextDeliveryDate(baseDate, deliveryDayName) { const target = DAY_NAMES.indexOf(deliveryDayName); let date = new Date(baseDate); date.setUTCDate(date.getUTCDate()+1); while (date.getUTCDay() !== target) date.setUTCDate(date.getUTCDate()+1); return date; }
function adjustForCustomer(date, customer, holidays, irr) { const deliverable = (irr && irr.deliverable) ? irr.deliverable : customer.deliverable; let d = new Date(date), count = 0; while (!deliverable[DAY_NAMES[d.getUTCDay()]] || holidays.includes(formatDate(d))) { d.setUTCDate(d.getUTCDate()+1); if (++count > 30) break; } return d; }
function adjustForWarehouse(date, warehouse, type, holidays) { let d = new Date(date), count = 0; while (true) { const dayName = DAY_NAMES[d.getUTCDay()]; const isHol = holidays.includes(formatDate(d)); if ((isHol && !warehouse[type]["祝日"]) || !warehouse[type][dayName]) { d.setUTCDate(d.getUTCDate()+1); if (++count > 30) break; } else break; } return d; }
function calcZaikoWeekday(now, customer, warehouse, holidays, irr) { const base = getBaseDate(now,"14:00",holidays); let ship = nextBusinessDay(base,holidays); ship = adjustForWarehouse(ship,warehouse,"納品",holidays); return adjustForCustomer(ship,customer,holidays,irr); }
function calcTsuukaWeekday(now, maker, customer, warehouse, holidays, irr) { const base = getBaseDate(now,maker.deadline,holidays); const dayName = DAY_NAMES[base.getUTCDay()]; let arrival = getNextDeliveryDate(base,maker.deliveryMap[dayName]); if (holidays.includes(formatDate(arrival))) { arrival.setUTCDate(arrival.getUTCDate()+1); if (maker.holidayAdd > 0) arrival.setUTCDate(arrival.getUTCDate()+maker.holidayAdd); } const confirmedArrival = new Date(arrival); let ship = adjustForWarehouse(arrival,warehouse,"納品",holidays); ship = nextBusinessDay(ship,holidays); return { deliveryDate: adjustForCustomer(ship,customer,holidays,irr), arrivalDate: confirmedArrival }; }
function calcChokusoWeekday(now, maker, customer, holidays, irr) { const base = getBaseDate(now,maker.deadline,holidays); const dayName = DAY_NAMES[base.getUTCDay()]; let delivery = getNextDeliveryDate(base,maker.deliveryMap[dayName]); if (holidays.includes(formatDate(delivery))) { delivery.setUTCDate(delivery.getUTCDate()+1); if (maker.holidayAdd > 0) delivery.setUTCDate(delivery.getUTCDate()+maker.holidayAdd); } return adjustForCustomer(delivery,customer,holidays,irr); }
function calcZaikoDays(now, customer, warehouse, holidays, irr) { const base = getBaseDate(now,"14:00",holidays); let ship = nextBusinessDay(base,holidays); ship = adjustForWarehouse(ship,warehouse,"納品",holidays); if (irr && irr.days !== undefined) ship = addBusinessDays(ship,irr.days,holidays); return adjustForCustomer(ship,customer,holidays,null); }
function calcTsuukaDays(now, maker, customer, warehouse, holidays, irr) { const base = getBaseDate(now,maker.deadline,holidays); const days = (irr && irr.days !== undefined) ? irr.days : maker.days; let arrival = addBusinessDays(base,days,holidays); if (maker.holidayAdd > 0 && holidays.includes(formatDate(arrival))) arrival = addBusinessDays(arrival,maker.holidayAdd,holidays); const confirmedArrival = new Date(arrival); let ship = adjustForWarehouse(arrival,warehouse,"納品",holidays); ship = nextBusinessDay(ship,holidays); return { deliveryDate: adjustForCustomer(ship,customer,holidays,null), arrivalDate: confirmedArrival }; }
function calcChokusoDays(now, maker, customer, holidays, irr) { const base = getBaseDate(now,maker.deadline,holidays); const days = (irr && irr.days !== undefined) ? irr.days : maker.days; let delivery = addBusinessDays(base,days,holidays); if (maker.holidayAdd > 0 && holidays.includes(formatDate(delivery))) delivery = addBusinessDays(delivery,maker.holidayAdd,holidays); return adjustForCustomer(delivery,customer,holidays,null); }
