/******/ (function() { // webpackBootstrap
/*!************************************!*\
  !*** ./src/functions/functions.js ***!
  \************************************/
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/* global Excel */
/**
 * 最短納品日を算出するカスタム関数
 * @customfunction
 * @param {string} orderDate 受注日（例："2026/03/07"）
 * @param {string} makerCode メーカーコード（例："MK001"）
 * @param {string} customerCode 得意先コード（例："CS001"）
 * @returns {string} 最短納品日
 */
function DELIVERY_DATE(_x, _x2, _x3) {
  return _DELIVERY_DATE.apply(this, arguments);
} // 以下はfunctions.js用のヘルパー関数
function _DELIVERY_DATE() {
  _DELIVERY_DATE = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2(orderDate, makerCode, customerCode) {
    var _t;
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.p = _context2.n) {
        case 0:
          _context2.p = 0;
          _context2.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee(context) {
              var holidays, makerSheet, makerRange, customerSheet, customerRange, makerRows, maker, i, customerRows, customer, _i, dayNames, parts, currentDate, remainDays, deliveryRemain, count;
              return _regenerator().w(function (_context) {
                while (1) switch (_context.n) {
                  case 0:
                    _context.n = 1;
                    return loadHolidaysForFunc(context);
                  case 1:
                    holidays = _context.v;
                    // メーカーマスタ読み込み
                    makerSheet = context.workbook.worksheets.getItem("TBL_MAKER");
                    makerRange = makerSheet.getUsedRange();
                    makerRange.load("values");

                    // 得意先マスタ読み込み
                    customerSheet = context.workbook.worksheets.getItem("TBL_CUSTOMER");
                    customerRange = customerSheet.getUsedRange();
                    customerRange.load("values");
                    _context.n = 2;
                    return context.sync();
                  case 2:
                    // メーカー情報を検索
                    makerRows = makerRange.values;
                    maker = null;
                    i = 1;
                  case 3:
                    if (!(i < makerRows.length)) {
                      _context.n = 5;
                      break;
                    }
                    if (!(makerRows[i][0] === makerCode)) {
                      _context.n = 4;
                      break;
                    }
                    maker = {
                      name: makerRows[i][1],
                      type: makerRows[i][2],
                      leadDays: Number(makerRows[i][3]),
                      closedDays: makerRows[i][5]
                    };
                    return _context.a(3, 5);
                  case 4:
                    i++;
                    _context.n = 3;
                    break;
                  case 5:
                    if (maker) {
                      _context.n = 6;
                      break;
                    }
                    return _context.a(2, "\u30A8\u30E9\u30FC\uFF1A\u30E1\u30FC\u30AB\u30FC\u300C".concat(makerCode, "\u300D\u304C\u898B\u3064\u304B\u308A\u307E\u305B\u3093"));
                  case 6:
                    // 得意先情報を検索
                    customerRows = customerRange.values;
                    customer = null;
                    _i = 1;
                  case 7:
                    if (!(_i < customerRows.length)) {
                      _context.n = 9;
                      break;
                    }
                    if (!(customerRows[_i][0] === customerCode)) {
                      _context.n = 8;
                      break;
                    }
                    customer = {
                      deliveryDays: customerRows[_i][3],
                      deliveryLead: Number(customerRows[_i][4]),
                      closedDays: customerRows[_i][5]
                    };
                    return _context.a(3, 9);
                  case 8:
                    _i++;
                    _context.n = 7;
                    break;
                  case 9:
                    if (customer) {
                      _context.n = 10;
                      break;
                    }
                    return _context.a(2, "\u30A8\u30E9\u30FC\uFF1A\u5F97\u610F\u5148\u300C".concat(customerCode, "\u300D\u304C\u898B\u3064\u304B\u308A\u307E\u305B\u3093"));
                  case 10:
                    // 納期算出
                    dayNames = ["日", "月", "火", "水", "木", "金", "土"];
                    parts = orderDate.split("/");
                    currentDate = new Date(Date.UTC(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]))); // メーカーリードタイム加算
                    remainDays = maker.leadDays;
                    while (remainDays > 0) {
                      currentDate.setUTCDate(currentDate.getUTCDate() + 1);
                      if (isBusinessDayFunc(currentDate, maker.closedDays, holidays)) {
                        remainDays--;
                      }
                    }

                    // 配送リードタイム加算
                    deliveryRemain = customer.deliveryLead;
                    while (deliveryRemain > 0) {
                      currentDate.setUTCDate(currentDate.getUTCDate() + 1);
                      deliveryRemain--;
                    }

                    // 得意先配送可能曜日に合わせてスライド
                    count = 0;
                  case 11:
                    if (!(!customer.deliveryDays.includes(dayNames[currentDate.getUTCDay()]) || !isBusinessDayFunc(currentDate, customer.closedDays, holidays))) {
                      _context.n = 13;
                      break;
                    }
                    currentDate.setUTCDate(currentDate.getUTCDate() + 1);
                    count++;
                    if (!(count > 30)) {
                      _context.n = 12;
                      break;
                    }
                    return _context.a(3, 13);
                  case 12:
                    _context.n = 11;
                    break;
                  case 13:
                    return _context.a(2, formatDateFunc(currentDate));
                }
              }, _callee);
            }));
            return function (_x5) {
              return _ref.apply(this, arguments);
            };
          }());
        case 1:
          return _context2.a(2, _context2.v);
        case 2:
          _context2.p = 2;
          _t = _context2.v;
          return _context2.a(2, "エラー：" + _t.message);
      }
    }, _callee2, null, [[0, 2]]);
  }));
  return _DELIVERY_DATE.apply(this, arguments);
}
function loadHolidaysForFunc(_x4) {
  return _loadHolidaysForFunc.apply(this, arguments);
}
function _loadHolidaysForFunc() {
  _loadHolidaysForFunc = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3(context) {
    var sheet, range, rows, holidays, i, date;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.n) {
        case 0:
          sheet = context.workbook.worksheets.getItem("TBL_HOLIDAY");
          range = sheet.getUsedRange();
          range.load("values");
          _context3.n = 1;
          return context.sync();
        case 1:
          rows = range.values;
          holidays = [];
          for (i = 1; i < rows.length; i++) {
            date = excelDateToDateFunc(rows[i][0]);
            holidays.push(formatDateFunc(date));
          }
          return _context3.a(2, holidays);
      }
    }, _callee3);
  }));
  return _loadHolidaysForFunc.apply(this, arguments);
}
function excelDateToDateFunc(serial) {
  var utc_days = Math.floor(serial - 25569);
  return new Date(utc_days * 86400 * 1000);
}
function formatDateFunc(date) {
  var y = date.getUTCFullYear();
  var m = String(date.getUTCMonth() + 1).padStart(2, "0");
  var d = String(date.getUTCDate()).padStart(2, "0");
  return "".concat(y, "/").concat(m, "/").concat(d);
}
function isBusinessDayFunc(date, closedDays, holidays) {
  var dayNames = ["日", "月", "火", "水", "木", "金", "土"];
  var dayName = dayNames[date.getUTCDay()];
  if (closedDays.includes(dayName)) return false;
  if (holidays.includes(formatDateFunc(date))) return false;
  return true;
}
CustomFunctions.associate("DELIVERY_DATE", DELIVERY_DATE);
/******/ })()
;
//# sourceMappingURL=functions.js.map