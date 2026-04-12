/* global Excel */

/**
 * 最短納品日を算出するカスタム関数
 * @customfunction
 * @param {string} orderDate 受注日（例："2026/03/07"）
 * @param {string} makerCode メーカーコード（例："MK001"）
 * @param {string} customerCode 得意先コード（例："CS001"）
 * @returns {string} 最短納品日
 */
async function DELIVERY_DATE(orderDate, makerCode, customerCode) {
  try {
    return await Excel.run(async (context) => {

      // 祝日マスタ読み込み
      const holidays = await loadHolidaysForFunc(context);

      // メーカーマスタ読み込み
      const makerSheet = context.workbook.worksheets.getItem("TBL_MAKER");
      const makerRange = makerSheet.getUsedRange();
      makerRange.load("values");

      // 得意先マスタ読み込み
      const customerSheet = context.workbook.worksheets.getItem("TBL_CUSTOMER");
      const customerRange = customerSheet.getUsedRange();
      customerRange.load("values");

      await context.sync();

      // メーカー情報を検索
      const makerRows = makerRange.values;
      let maker = null;
      for (let i = 1; i < makerRows.length; i++) {
        if (makerRows[i][0] === makerCode) {
          maker = {
            name:       makerRows[i][1],
            type:       makerRows[i][2],
            leadDays:   Number(makerRows[i][3]),
            closedDays: makerRows[i][5],
          };
          break;
        }
      }
      if (!maker) return `エラー：メーカー「${makerCode}」が見つかりません`;

      // 得意先情報を検索
      const customerRows = customerRange.values;
      let customer = null;
      for (let i = 1; i < customerRows.length; i++) {
        if (customerRows[i][0] === customerCode) {
          customer = {
            deliveryDays: customerRows[i][3],
            deliveryLead: Number(customerRows[i][4]),
            closedDays:   customerRows[i][5],
          };
          break;
        }
      }
      if (!customer) return `エラー：得意先「${customerCode}」が見つかりません`;

      // 納期算出
      const dayNames = ["日", "月", "火", "水", "木", "金", "土"];
      const parts = orderDate.split("/");
      let currentDate = new Date(Date.UTC(
        Number(parts[0]),
        Number(parts[1]) - 1,
        Number(parts[2])
      ));

      // メーカーリードタイム加算
      let remainDays = maker.leadDays;
      while (remainDays > 0) {
        currentDate.setUTCDate(currentDate.getUTCDate() + 1);
        if (isBusinessDayFunc(currentDate, maker.closedDays, holidays)) {
          remainDays--;
        }
      }

      // 配送リードタイム加算
      let deliveryRemain = customer.deliveryLead;
      while (deliveryRemain > 0) {
        currentDate.setUTCDate(currentDate.getUTCDate() + 1);
        deliveryRemain--;
      }

      // 得意先配送可能曜日に合わせてスライド
      let count = 0;
      while (
        !customer.deliveryDays.includes(dayNames[currentDate.getUTCDay()]) ||
        !isBusinessDayFunc(currentDate, customer.closedDays, holidays)
      ) {
        currentDate.setUTCDate(currentDate.getUTCDate() + 1);
        count++;
        if (count > 30) break;
      }

      return formatDateFunc(currentDate);
    });

  } catch (error) {
    return "エラー：" + error.message;
  }
}

// 以下はfunctions.js用のヘルパー関数

async function loadHolidaysForFunc(context) {
  const sheet = context.workbook.worksheets.getItem("TBL_HOLIDAY");
  const range = sheet.getUsedRange();
  range.load("values");
  await context.sync();
  const rows = range.values;
  const holidays = [];
  for (let i = 1; i < rows.length; i++) {
    const date = excelDateToDateFunc(rows[i][0]);
    holidays.push(formatDateFunc(date));
  }
  return holidays;
}

function excelDateToDateFunc(serial) {
  const utc_days = Math.floor(serial - 25569);
  return new Date(utc_days * 86400 * 1000);
}

function formatDateFunc(date) {
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, "0");
  const d = String(date.getUTCDate()).padStart(2, "0");
  return `${y}/${m}/${d}`;
}

function isBusinessDayFunc(date, closedDays, holidays) {
  const dayNames = ["日", "月", "火", "水", "木", "金", "土"];
  const dayName = dayNames[date.getUTCDay()];
  if (closedDays.includes(dayName)) return false;
  if (holidays.includes(formatDateFunc(date))) return false;
  return true;
}

CustomFunctions.associate("DELIVERY_DATE", DELIVERY_DATE);