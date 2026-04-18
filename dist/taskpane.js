/******/ (function() { // webpackBootstrap
/******/ 	// The require scope
/******/ 	var __webpack_require__ = {};
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function _regeneratorValues(e) { if (null != e) { var t = e["function" == typeof Symbol && Symbol.iterator || "@@iterator"], r = 0; if (t) return t.call(e); if ("function" == typeof e.next) return e; if (!isNaN(e.length)) return { next: function next() { return e && r >= e.length && (e = void 0), { value: e && e[r++], done: !e }; } }; } throw new TypeError(_typeof(e) + " is not iterable"); }
function _defineProperty(e, r, t) { return (r = _toPropertyKey(r)) in e ? Object.defineProperty(e, r, { value: t, enumerable: !0, configurable: !0, writable: !0 }) : e[r] = t, e; }
function _toPropertyKey(t) { var i = _toPrimitive(t, "string"); return "symbol" == _typeof(i) ? i : i + ""; }
function _toPrimitive(t, r) { if ("object" != _typeof(t) || !t) return t; var e = t[Symbol.toPrimitive]; if (void 0 !== e) { var i = e.call(t, r || "default"); if ("object" != _typeof(i)) return i; throw new TypeError("@@toPrimitive must return a primitive value."); } return ("string" === r ? String : Number)(t); }
function _slicedToArray(r, e) { return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray(r, e) || _nonIterableRest(); }
function _nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t.return && (u = t.return(), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function _arrayWithHoles(r) { if (Array.isArray(r)) return r; }
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/* global document, Excel, Office */

var SECRET = "DC2026SECRET";
var DAY_NAMES = ["日", "月", "火", "水", "木", "金", "土"];
var DEFAULT_CONFIG = {
  calcMethod: "曜日",
  order: {
    sheetName: "受注一覧",
    customerCodeCol: 1,
    productCodeCol: 6,
    makerCodeCol: 9,
    logisticsCol: 10,
    arrivalCol: 11,
    deliveryCol: 12,
    warehouseCol: 13
  },
  sheetNames: {
    maker: "TBL_MAKER",
    customer: "TBL_CUSTOMER",
    warehouse: "TBL_WAREHOUSE",
    holiday: "TBL_HOLIDAY",
    itemIrr: "TBL_ITEM_IRREGULAR",
    makerIrr: "TBL_MAKER_IRREGULAR"
  }
};
var cfg = JSON.parse(JSON.stringify(DEFAULT_CONFIG));
function sha256(_x) {
  return _sha.apply(this, arguments);
}
function _sha() {
  _sha = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2(message) {
    var msgBuffer, hashBuffer;
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.n) {
        case 0:
          msgBuffer = new TextEncoder().encode(message);
          _context2.n = 1;
          return crypto.subtle.digest("SHA-256", msgBuffer);
        case 1:
          hashBuffer = _context2.v;
          return _context2.a(2, Array.from(new Uint8Array(hashBuffer)).map(function (b) {
            return b.toString(16).padStart(2, "0");
          }).join(""));
      }
    }, _callee2);
  }));
  return _sha.apply(this, arguments);
}
function generateKey(_x2) {
  return _generateKey.apply(this, arguments);
}
function _generateKey() {
  _generateKey = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3(email) {
    var raw, chars;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.n) {
        case 0:
          _context3.n = 1;
          return sha256(email.toLowerCase() + SECRET);
        case 1:
          raw = _context3.v;
          chars = raw.toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 12);
          return _context3.a(2, "DC-" + chars.slice(0, 4) + "-" + chars.slice(4, 8) + "-" + chars.slice(8, 12));
      }
    }, _callee3);
  }));
  return _generateKey.apply(this, arguments);
}
function validateLicense(_x3, _x4) {
  return _validateLicense.apply(this, arguments);
}
function _validateLicense() {
  _validateLicense = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4(email, key) {
    var _t3, _t4;
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.n) {
        case 0:
          _t3 = key.toUpperCase();
          _context4.n = 1;
          return generateKey(email);
        case 1:
          _t4 = _context4.v;
          return _context4.a(2, _t3 === _t4);
      }
    }, _callee4);
  }));
  return _validateLicense.apply(this, arguments);
}
function activateLicense() {
  return _activateLicense.apply(this, arguments);
}
function _activateLicense() {
  _activateLicense = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5() {
    var msg, email, key, license, _t5;
    return _regenerator().w(function (_context5) {
      while (1) switch (_context5.p = _context5.n) {
        case 0:
          msg = document.getElementById("msg-license");
          email = document.getElementById("license-email").value.trim();
          key = document.getElementById("license-key").value.trim();
          if (!(!email || !key)) {
            _context5.n = 1;
            break;
          }
          msg.textContent = "❌ メールアドレスとライセンスキーを入力してください";
          return _context5.a(2);
        case 1:
          msg.textContent = "認証中...";
          _context5.p = 2;
          _context5.n = 3;
          return validateLicense(email, key);
        case 3:
          if (!_context5.v) {
            _context5.n = 6;
            break;
          }
          _context5.n = 4;
          return saveLicense(email, key);
        case 4:
          _context5.n = 5;
          return loadLicense();
        case 5:
          license = _context5.v;
          showApp(license);
          _context5.n = 7;
          break;
        case 6:
          msg.textContent = "❌ ライセンスキーが無効です";
        case 7:
          _context5.n = 9;
          break;
        case 8:
          _context5.p = 8;
          _t5 = _context5.v;
          msg.textContent = "❌ エラー: " + _t5.message;
        case 9:
          return _context5.a(2);
      }
    }, _callee5, null, [[2, 8]]);
  }));
  return _activateLicense.apply(this, arguments);
}
Office.onReady(/*#__PURE__*/function () {
  var _ref = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee(info) {
    var license, _t, _t2;
    return _regenerator().w(function (_context) {
      while (1) switch (_context.p = _context.n) {
        case 0:
          if (!(info.host === Office.HostType.Excel)) {
            _context.n = 9;
            break;
          }
          _context.p = 1;
          document.getElementById("sideload-msg").style.display = "none";
          document.getElementById("btn-activate").onclick = activateLicense;
          document.getElementById("btn-logout").onclick = logoutLicense;
          document.getElementById("btn-init-master").onclick = initMaster;
          document.getElementById("btn-calc-delivery").onclick = calcDelivery;
          document.getElementById("btn-save-settings").onclick = saveSettings;
          document.getElementById("btn-reset-settings").onclick = resetSettings;
          document.getElementById("btn-toggle-settings").onclick = toggleSettings;
          document.getElementById("btn-method-weekday").onclick = function () {
            cfg.calcMethod = "曜日";
            updateCalcMethodUI();
          };
          document.getElementById("btn-method-days").onclick = function () {
            cfg.calcMethod = "日数";
            updateCalcMethodUI();
          };
          _context.n = 2;
          return loadConfig();
        case 2:
          updateCalcMethodUI();
          _context.n = 3;
          return loadLicense();
        case 3:
          license = _context.v;
          _t = license;
          if (!_t) {
            _context.n = 5;
            break;
          }
          _context.n = 4;
          return validateLicense(license.email, license.key);
        case 4:
          _t = _context.v;
        case 5:
          if (!_t) {
            _context.n = 6;
            break;
          }
          showApp(license);
          _context.n = 7;
          break;
        case 6:
          showLicenseScreen();
        case 7:
          _context.n = 9;
          break;
        case 8:
          _context.p = 8;
          _t2 = _context.v;
          document.getElementById("sideload-msg").style.display = "block";
          document.getElementById("sideload-msg").textContent = "エラー：" + _t2.message;
        case 9:
          return _context.a(2);
      }
    }, _callee, null, [[1, 8]]);
  }));
  return function (_x5) {
    return _ref.apply(this, arguments);
  };
}());
function showLicenseScreen() {
  document.getElementById("license-screen").style.display = "block";
  document.getElementById("app-body").style.display = "none";
}
function showApp(license) {
  document.getElementById("license-screen").style.display = "none";
  document.getElementById("app-body").style.display = "block";
  document.getElementById("license-info").innerHTML = "✅ 認証済み：" + license.email + "<br>有効期限：" + license.expiry;
}
function saveLicense(_x6, _x7) {
  return _saveLicense.apply(this, arguments);
}
function _saveLicense() {
  _saveLicense = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee7(email, key) {
    return _regenerator().w(function (_context7) {
      while (1) switch (_context7.n) {
        case 0:
          _context7.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref2 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee6(context) {
              var props;
              return _regenerator().w(function (_context6) {
                while (1) switch (_context6.n) {
                  case 0:
                    props = context.workbook.properties.custom;
                    try {
                      props.getItemOrNullObject("DC_EMAIL").delete();
                    } catch (e) {}
                    try {
                      props.getItemOrNullObject("DC_KEY").delete();
                    } catch (e) {}
                    _context6.n = 1;
                    return context.sync();
                  case 1:
                    props.add("DC_EMAIL", email);
                    props.add("DC_KEY", key);
                    _context6.n = 2;
                    return context.sync();
                  case 2:
                    return _context6.a(2);
                }
              }, _callee6);
            }));
            return function (_x14) {
              return _ref2.apply(this, arguments);
            };
          }());
        case 1:
          return _context7.a(2);
      }
    }, _callee7);
  }));
  return _saveLicense.apply(this, arguments);
}
function loadLicense() {
  return _loadLicense.apply(this, arguments);
}
function _loadLicense() {
  _loadLicense = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee8() {
    var email, key, _t6;
    return _regenerator().w(function (_context8) {
      while (1) switch (_context8.p = _context8.n) {
        case 0:
          _context8.p = 0;
          email = localStorage.getItem("DC_EMAIL");
          key = localStorage.getItem("DC_KEY");
          if (!(!email || !key)) {
            _context8.n = 1;
            break;
          }
          return _context8.a(2, null);
        case 1:
          return _context8.a(2, {
            email: email,
            key: key,
            expiry: "無期限"
          });
        case 2:
          _context8.p = 2;
          _t6 = _context8.v;
          return _context8.a(2, null);
      }
    }, _callee8, null, [[0, 2]]);
  }));
  return _loadLicense.apply(this, arguments);
}
function logoutLicense() {
  return _logoutLicense.apply(this, arguments);
}
function _logoutLicense() {
  _logoutLicense = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee0() {
    return _regenerator().w(function (_context0) {
      while (1) switch (_context0.n) {
        case 0:
          _context0.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref3 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee9(context) {
              var props;
              return _regenerator().w(function (_context9) {
                while (1) switch (_context9.n) {
                  case 0:
                    props = context.workbook.properties.custom;
                    props.getItemOrNullObject("DC_EMAIL").delete();
                    props.getItemOrNullObject("DC_KEY").delete();
                    _context9.n = 1;
                    return context.sync();
                  case 1:
                    return _context9.a(2);
                }
              }, _callee9);
            }));
            return function (_x15) {
              return _ref3.apply(this, arguments);
            };
          }());
        case 1:
          showLicenseScreen();
        case 2:
          return _context0.a(2);
      }
    }, _callee0);
  }));
  return _logoutLicense.apply(this, arguments);
}
function saveConfig() {
  return _saveConfig.apply(this, arguments);
}
function _saveConfig() {
  _saveConfig = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee10() {
    var _t7;
    return _regenerator().w(function (_context10) {
      while (1) switch (_context10.p = _context10.n) {
        case 0:
          _context10.p = 0;
          _context10.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref4 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee1(context) {
              var props, d1, d2, _i2, _arr, _arr$_i, key, val, item;
              return _regenerator().w(function (_context1) {
                while (1) switch (_context1.n) {
                  case 0:
                    props = context.workbook.properties.custom;
                    d1 = JSON.stringify({
                      calcMethod: cfg.calcMethod,
                      order: cfg.order
                    });
                    d2 = JSON.stringify({
                      sheetNames: cfg.sheetNames
                    });
                    _i2 = 0, _arr = [["DC_CFG1", d1], ["DC_CFG2", d2]];
                  case 1:
                    if (!(_i2 < _arr.length)) {
                      _context1.n = 4;
                      break;
                    }
                    _arr$_i = _slicedToArray(_arr[_i2], 2), key = _arr$_i[0], val = _arr$_i[1];
                    item = props.getItemOrNullObject(key);
                    item.load("isNullObject");
                    _context1.n = 2;
                    return context.sync();
                  case 2:
                    if (!item.isNullObject) {
                      item.value = val;
                    } else {
                      props.add(key, val);
                    }
                  case 3:
                    _i2++;
                    _context1.n = 1;
                    break;
                  case 4:
                    _context1.n = 5;
                    return context.sync();
                  case 5:
                    return _context1.a(2);
                }
              }, _callee1);
            }));
            return function (_x16) {
              return _ref4.apply(this, arguments);
            };
          }());
        case 1:
          _context10.n = 3;
          break;
        case 2:
          _context10.p = 2;
          _t7 = _context10.v;
          console.error("設定保存エラー", _t7);
        case 3:
          return _context10.a(2);
      }
    }, _callee10, null, [[0, 2]]);
  }));
  return _saveConfig.apply(this, arguments);
}
function loadConfig() {
  return _loadConfig.apply(this, arguments);
}
function _loadConfig() {
  _loadConfig = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee12() {
    var _t8;
    return _regenerator().w(function (_context12) {
      while (1) switch (_context12.p = _context12.n) {
        case 0:
          _context12.p = 0;
          _context12.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref5 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee11(context) {
              var props, i1, i2, s1, s2;
              return _regenerator().w(function (_context11) {
                while (1) switch (_context11.n) {
                  case 0:
                    props = context.workbook.properties.custom;
                    i1 = props.getItemOrNullObject("DC_CFG1");
                    i2 = props.getItemOrNullObject("DC_CFG2");
                    i1.load(["isNullObject", "value"]);
                    i2.load(["isNullObject", "value"]);
                    _context11.n = 1;
                    return context.sync();
                  case 1:
                    if (!i1.isNullObject && i1.value) {
                      s1 = JSON.parse(i1.value);
                      cfg.calcMethod = s1.calcMethod || DEFAULT_CONFIG.calcMethod;
                      cfg.order = Object.assign(JSON.parse(JSON.stringify(DEFAULT_CONFIG.order)), s1.order || {});
                    }
                    if (!i2.isNullObject && i2.value) {
                      s2 = JSON.parse(i2.value);
                      cfg.sheetNames = Object.assign(JSON.parse(JSON.stringify(DEFAULT_CONFIG.sheetNames)), s2.sheetNames || {});
                    }
                  case 2:
                    return _context11.a(2);
                }
              }, _callee11);
            }));
            return function (_x17) {
              return _ref5.apply(this, arguments);
            };
          }());
        case 1:
          _context12.n = 3;
          break;
        case 2:
          _context12.p = 2;
          _t8 = _context12.v;
          cfg = JSON.parse(JSON.stringify(DEFAULT_CONFIG));
        case 3:
          return _context12.a(2);
      }
    }, _callee12, null, [[0, 2]]);
  }));
  return _loadConfig.apply(this, arguments);
}
function updateCalcMethodUI() {
  var isWeekday = cfg.calcMethod === "曜日";
  document.getElementById("btn-method-weekday").className = isWeekday ? "method-btn method-active" : "method-btn";
  document.getElementById("btn-method-days").className = isWeekday ? "method-btn" : "method-btn method-active";
  document.getElementById("method-label").textContent = "演算方法：" + cfg.calcMethod + "モード";
}
function setCalcMethod(_x8) {
  return _setCalcMethod.apply(this, arguments);
}
function _setCalcMethod() {
  _setCalcMethod = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee13(method) {
    return _regenerator().w(function (_context13) {
      while (1) switch (_context13.n) {
        case 0:
          cfg.calcMethod = method;
          updateCalcMethodUI();
          _context13.n = 1;
          return saveConfig();
        case 1:
          return _context13.a(2);
      }
    }, _callee13);
  }));
  return _setCalcMethod.apply(this, arguments);
}
function toggleSettings() {
  var panel = document.getElementById("settings-panel");
  var btn = document.getElementById("btn-toggle-settings");
  var isHidden = panel.style.display === "none" || panel.style.display === "";
  panel.style.display = isHidden ? "block" : "none";
  btn.textContent = isHidden ? "▲ 設定を閉じる" : "⚙️ 列・シート設定を開く";
  if (isHidden) populateSettings();
}
function populateSettings() {
  var o = cfg.order,
    s = cfg.sheetNames;
  var fields = {
    "s-order-sheet": o.sheetName,
    "s-order-customer-col": o.customerCodeCol,
    "s-order-product-col": o.productCodeCol,
    "s-order-maker-col": o.makerCodeCol,
    "s-order-logistics-col": o.logisticsCol,
    "s-order-warehouse-col": o.warehouseCol,
    "s-order-arrival-col": o.arrivalCol,
    "s-order-delivery-col": o.deliveryCol,
    "s-maker-sheet": s.maker,
    "s-customer-sheet": s.customer,
    "s-warehouse-sheet": s.warehouse,
    "s-holiday-sheet": s.holiday,
    "s-item-irr-sheet": s.itemIrr,
    "s-maker-irr-sheet": s.makerIrr
  };
  for (var _i = 0, _Object$entries = Object.entries(fields); _i < _Object$entries.length; _i++) {
    var _Object$entries$_i = _slicedToArray(_Object$entries[_i], 2),
      id = _Object$entries$_i[0],
      val = _Object$entries$_i[1];
    var el = document.getElementById(id);
    if (el) el.value = val;
  }
}
function saveSettings() {
  return _saveSettings.apply(this, arguments);
}
function _saveSettings() {
  _saveSettings = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee14() {
    return _regenerator().w(function (_context14) {
      while (1) switch (_context14.n) {
        case 0:
          cfg.order.sheetName = document.getElementById("s-order-sheet").value.trim();
          cfg.order.customerCodeCol = parseInt(document.getElementById("s-order-customer-col").value);
          cfg.order.productCodeCol = parseInt(document.getElementById("s-order-product-col").value);
          cfg.order.makerCodeCol = parseInt(document.getElementById("s-order-maker-col").value);
          cfg.order.logisticsCol = parseInt(document.getElementById("s-order-logistics-col").value);
          cfg.order.warehouseCol = parseInt(document.getElementById("s-order-warehouse-col").value);
          cfg.order.arrivalCol = parseInt(document.getElementById("s-order-arrival-col").value);
          cfg.order.deliveryCol = parseInt(document.getElementById("s-order-delivery-col").value);
          cfg.sheetNames.maker = document.getElementById("s-maker-sheet").value.trim();
          cfg.sheetNames.customer = document.getElementById("s-customer-sheet").value.trim();
          cfg.sheetNames.warehouse = document.getElementById("s-warehouse-sheet").value.trim();
          cfg.sheetNames.holiday = document.getElementById("s-holiday-sheet").value.trim();
          cfg.sheetNames.itemIrr = document.getElementById("s-item-irr-sheet").value.trim();
          cfg.sheetNames.makerIrr = document.getElementById("s-maker-irr-sheet").value.trim();
          _context14.n = 1;
          return saveConfig();
        case 1:
          showMsg("msg-settings", "✅ 設定を保存しました");
        case 2:
          return _context14.a(2);
      }
    }, _callee14);
  }));
  return _saveSettings.apply(this, arguments);
}
function resetSettings() {
  return _resetSettings.apply(this, arguments);
}
function _resetSettings() {
  _resetSettings = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee15() {
    return _regenerator().w(function (_context15) {
      while (1) switch (_context15.n) {
        case 0:
          cfg = JSON.parse(JSON.stringify(DEFAULT_CONFIG));
          _context15.n = 1;
          return saveConfig();
        case 1:
          populateSettings();
          updateCalcMethodUI();
          showMsg("msg-settings", "✅ デフォルトに戻しました");
        case 2:
          return _context15.a(2);
      }
    }, _callee15);
  }));
  return _resetSettings.apply(this, arguments);
}
function showMsg(id, text) {
  var ms = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : 3000;
  var el = document.getElementById(id);
  if (!el) return;
  el.textContent = text;
  setTimeout(function () {
    el.textContent = "";
  }, ms);
}
function initMaster() {
  return _initMaster.apply(this, arguments);
}
function _initMaster() {
  _initMaster = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee17() {
    var msg, _t9;
    return _regenerator().w(function (_context17) {
      while (1) switch (_context17.p = _context17.n) {
        case 0:
          msg = document.getElementById("msg-init");
          msg.textContent = "作成中...";
          _context17.p = 1;
          _context17.n = 2;
          return Excel.run(/*#__PURE__*/function () {
            var _ref6 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee16(context) {
              var sheets, isWeekday, sn, irrHeaders, masterDefs, _i3, _masterDefs, master, existing, sheet, headerRange;
              return _regenerator().w(function (_context16) {
                while (1) switch (_context16.n) {
                  case 0:
                    sheets = context.workbook.worksheets;
                    isWeekday = cfg.calcMethod === "曜日";
                    sn = cfg.sheetNames;
                    irrHeaders = isWeekday ? ["得意先コード", "得意先名", "コード", "名称", "月", "火", "水", "木", "金", "土", "日"] : ["得意先コード", "得意先名", "コード", "名称", "日数（営業日）"];
                    masterDefs = [{
                      name: sn.maker,
                      headers: isWeekday ? ["発注先コード", "メーカー名称", "受注締め時間", "月→納品", "火→納品", "水→納品", "木→納品", "金→納品", "旗日明け加算"] : ["発注先コード", "メーカー名称", "受注締め時間", "リードタイム（営業日）", "旗日明け加算"],
                      data: []
                    }, {
                      name: sn.customer,
                      headers: ["得意先コード", "得意先名称", "月", "火", "水", "木", "金", "土", "日"],
                      data: []
                    }, {
                      name: sn.warehouse,
                      headers: ["倉庫コード", "倉庫名", "区分", "月", "火", "水", "木", "金", "土", "日", "祝日", "締め時間"],
                      data: [["WH001", "本社倉庫", "入荷", "〇", "〇", "〇", "〇", "〇", "〇", "×", "×", "14:00"], ["WH001", "本社倉庫", "納品", "〇", "〇", "〇", "〇", "〇", "×", "×", "×", ""], ["WH002", "第二倉庫", "入荷", "〇", "〇", "〇", "〇", "〇", "〇", "×", "×", "15:00"], ["WH002", "第二倉庫", "納品", "〇", "〇", "〇", "〇", "〇", "×", "×", "×", ""]]
                    }, {
                      name: sn.holiday,
                      headers: ["日付", "祝日名"],
                      data: []
                    }, {
                      name: sn.itemIrr,
                      headers: irrHeaders,
                      data: []
                    }, {
                      name: sn.makerIrr,
                      headers: irrHeaders,
                      data: []
                    }];
                    _i3 = 0, _masterDefs = masterDefs;
                  case 1:
                    if (!(_i3 < _masterDefs.length)) {
                      _context16.n = 4;
                      break;
                    }
                    master = _masterDefs[_i3];
                    existing = sheets.getItemOrNullObject(master.name);
                    _context16.n = 2;
                    return context.sync();
                  case 2:
                    if (!existing.isNullObject) existing.delete();
                    sheet = sheets.add(master.name);
                    headerRange = sheet.getRangeByIndexes(0, 0, 1, master.headers.length);
                    headerRange.values = [master.headers];
                    headerRange.format.fill.color = "#217346";
                    headerRange.format.font.color = "white";
                    headerRange.format.font.bold = true;
                    headerRange.format.columnWidth = 120;
                    if (master.data.length > 0) sheet.getRangeByIndexes(1, 0, master.data.length, master.headers.length).values = master.data;
                  case 3:
                    _i3++;
                    _context16.n = 1;
                    break;
                  case 4:
                    _context16.n = 5;
                    return context.sync();
                  case 5:
                    msg.textContent = "✅ マスタシートを作成しました！";
                  case 6:
                    return _context16.a(2);
                }
              }, _callee16);
            }));
            return function (_x18) {
              return _ref6.apply(this, arguments);
            };
          }());
        case 2:
          _context17.n = 4;
          break;
        case 3:
          _context17.p = 3;
          _t9 = _context17.v;
          msg.textContent = "❌ エラー：" + _t9.message;
        case 4:
          return _context17.a(2);
      }
    }, _callee17, null, [[1, 3]]);
  }));
  return _initMaster.apply(this, arguments);
}
function excelDateToDate(serial) {
  return new Date(Math.floor(serial - 25569) * 86400 * 1000);
}
function formatDate(date) {
  return date.getUTCFullYear() + "/" + String(date.getUTCMonth() + 1).padStart(2, "0") + "/" + String(date.getUTCDate()).padStart(2, "0");
}
function loadHolidays(_x9) {
  return _loadHolidays.apply(this, arguments);
}
function _loadHolidays() {
  _loadHolidays = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee18(context) {
    var sheet, range;
    return _regenerator().w(function (_context18) {
      while (1) switch (_context18.n) {
        case 0:
          sheet = context.workbook.worksheets.getItem(cfg.sheetNames.holiday);
          range = sheet.getUsedRange();
          range.load("values");
          _context18.n = 1;
          return context.sync();
        case 1:
          return _context18.a(2, range.values.slice(1).filter(function (r) {
            return r[0];
          }).map(function (r) {
            return formatDate(excelDateToDate(r[0]));
          }));
      }
    }, _callee18);
  }));
  return _loadHolidays.apply(this, arguments);
}
function loadMakers(_x0) {
  return _loadMakers.apply(this, arguments);
}
function _loadMakers() {
  _loadMakers = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee19(context) {
    var sheet, range, isWeekday;
    return _regenerator().w(function (_context19) {
      while (1) switch (_context19.n) {
        case 0:
          sheet = context.workbook.worksheets.getItem(cfg.sheetNames.maker);
          range = sheet.getUsedRange();
          range.load("values");
          _context19.n = 1;
          return context.sync();
        case 1:
          isWeekday = cfg.calcMethod === "曜日";
          return _context19.a(2, range.values.slice(1).filter(function (r) {
            return r[0];
          }).map(function (r) {
            var maker = {
              code: String(r[0]),
              name: r[1],
              deadline: r[2],
              holidayAdd: isWeekday ? Number(r[8]) || 0 : Number(r[4]) || 0
            };
            if (isWeekday) maker.deliveryMap = {
              "月": r[3],
              "火": r[4],
              "水": r[5],
              "木": r[6],
              "金": r[7]
            };else maker.days = Number(r[3]) || 1;
            return maker;
          }));
      }
    }, _callee19);
  }));
  return _loadMakers.apply(this, arguments);
}
function loadCustomers(_x1) {
  return _loadCustomers.apply(this, arguments);
}
function _loadCustomers() {
  _loadCustomers = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee20(context) {
    var sheet, range;
    return _regenerator().w(function (_context20) {
      while (1) switch (_context20.n) {
        case 0:
          sheet = context.workbook.worksheets.getItem(cfg.sheetNames.customer);
          range = sheet.getUsedRange();
          range.load("values");
          _context20.n = 1;
          return context.sync();
        case 1:
          return _context20.a(2, range.values.slice(1).filter(function (r) {
            return r[0];
          }).map(function (r) {
            return {
              code: String(r[0]),
              name: r[1],
              deliverable: {
                "月": r[2] === "〇",
                "火": r[3] === "〇",
                "水": r[4] === "〇",
                "木": r[5] === "〇",
                "金": r[6] === "〇",
                "土": r[7] === "〇",
                "日": r[8] === "〇"
              }
            };
          }));
      }
    }, _callee20);
  }));
  return _loadCustomers.apply(this, arguments);
}
function loadWarehouse(_x10) {
  return _loadWarehouse.apply(this, arguments);
}
function _loadWarehouse() {
  _loadWarehouse = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee21(context) {
    var sheet, range, wh, defaultCode;
    return _regenerator().w(function (_context21) {
      while (1) switch (_context21.n) {
        case 0:
          sheet = context.workbook.worksheets.getItem(cfg.sheetNames.warehouse);
          range = sheet.getUsedRange();
          range.load("values");
          _context21.n = 1;
          return context.sync();
        case 1:
          wh = {};
          defaultCode = null;
          range.values.slice(1).forEach(function (r) {
            var code = String(r[0] || "").trim();
            var kubun = String(r[2] || "").trim();
            if (!code || !kubun) return;
            if (!defaultCode) defaultCode = code;
            if (!wh[code]) wh[code] = {
              name: r[1],
              入荷: {},
              納品: {}
            };
            wh[code][kubun] = {
              "月": r[3] === "〇",
              "火": r[4] === "〇",
              "水": r[5] === "〇",
              "木": r[6] === "〇",
              "金": r[7] === "〇",
              "土": r[8] === "〇",
              "日": r[9] === "〇",
              "祝日": r[10] === "〇",
              "締め時間": String(r[11] || "").trim()
            };
          });
          wh.__default = defaultCode;
          return _context21.a(2, wh);
      }
    }, _callee21);
  }));
  return _loadWarehouse.apply(this, arguments);
}
function loadIrregulars(_x11, _x12, _x13) {
  return _loadIrregulars.apply(this, arguments);
}
function _loadIrregulars() {
  _loadIrregulars = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee22(context, sheetName, keyField) {
    var sheet, range, isWeekday, _t0;
    return _regenerator().w(function (_context22) {
      while (1) switch (_context22.p = _context22.n) {
        case 0:
          _context22.p = 0;
          sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
          _context22.n = 1;
          return context.sync();
        case 1:
          if (!sheet.isNullObject) {
            _context22.n = 2;
            break;
          }
          return _context22.a(2, []);
        case 2:
          range = sheet.getUsedRange();
          range.load("values");
          _context22.n = 3;
          return context.sync();
        case 3:
          isWeekday = cfg.calcMethod === "曜日";
          return _context22.a(2, range.values.slice(1).filter(function (r) {
            return r[0] && r[2];
          }).map(function (r) {
            var entry = _defineProperty({
              customerCode: String(r[0])
            }, keyField, String(r[2]));
            if (isWeekday) {
              entry.deliveryMap = {};
              var orderDay = String(r[4] || "").trim();
              var deliveryDay = String(r[5] || "").trim();
              if (orderDay && deliveryDay) entry.deliveryMap[orderDay] = deliveryDay;
            } else entry.days = Number(r[4]) || 1;
            return entry;
          }));
        case 4:
          _context22.p = 4;
          _t0 = _context22.v;
          return _context22.a(2, []);
      }
    }, _callee22, null, [[0, 4]]);
  }));
  return _loadIrregulars.apply(this, arguments);
}
function calcDelivery() {
  return _calcDelivery.apply(this, arguments);
}
function _calcDelivery() {
  _calcDelivery = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee24() {
    var msg, resultBox, resultDetail, now, o, _t10;
    return _regenerator().w(function (_context25) {
      while (1) switch (_context25.p = _context25.n) {
        case 0:
          msg = document.getElementById("msg-delivery");
          resultBox = document.getElementById("result-box");
          resultDetail = document.getElementById("result-detail");
          resultBox.style.display = "none";
          msg.innerHTML = "計算中...";
          now = new Date();
          o = cfg.order;
          _context25.p = 1;
          _context25.n = 2;
          return Excel.run(/*#__PURE__*/function () {
            var _ref7 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee23(context) {
              var holidays, makers, customers, warehouse, itemIrrs, makerIrrs, orderSheet, orderRange, rows, successCount, errorCount, errorLog, _loop, _ret, i, html;
              return _regenerator().w(function (_context24) {
                while (1) switch (_context24.n) {
                  case 0:
                    _context24.n = 1;
                    return loadHolidays(context);
                  case 1:
                    holidays = _context24.v;
                    _context24.n = 2;
                    return loadMakers(context);
                  case 2:
                    makers = _context24.v;
                    _context24.n = 3;
                    return loadCustomers(context);
                  case 3:
                    customers = _context24.v;
                    _context24.n = 4;
                    return loadWarehouse(context);
                  case 4:
                    warehouse = _context24.v;
                    _context24.n = 5;
                    return loadIrregulars(context, cfg.sheetNames.itemIrr, "productCode");
                  case 5:
                    itemIrrs = _context24.v;
                    _context24.n = 6;
                    return loadIrregulars(context, cfg.sheetNames.makerIrr, "makerCode");
                  case 6:
                    makerIrrs = _context24.v;
                    orderSheet = context.workbook.worksheets.getItem(o.sheetName);
                    orderRange = orderSheet.getUsedRange();
                    orderRange.load("values");
                    _context24.n = 7;
                    return context.sync();
                  case 7:
                    rows = orderRange.values;
                    successCount = 0, errorCount = 0;
                    errorLog = [];
                    _loop = /*#__PURE__*/_regenerator().m(function _loop() {
                      var row, customerCode, productCode, makerCode, logistics, warehouseCode, whCode, wh, customer, maker, itemIrr, makerIrr, irr, deliveryDate, arrivalDate, isWeekdayMode, r, _r, _t1;
                      return _regenerator().w(function (_context23) {
                        while (1) switch (_context23.p = _context23.n) {
                          case 0:
                            row = rows[i];
                            customerCode = String(row[o.customerCodeCol - 1] || "").trim();
                            if (customerCode) {
                              _context23.n = 1;
                              break;
                            }
                            return _context23.a(2, 0);
                          case 1:
                            productCode = String(row[o.productCodeCol - 1] || "").trim();
                            makerCode = String(row[o.makerCodeCol - 1] || "").trim();
                            logistics = String(row[o.logisticsCol - 1] || "").trim();
                            warehouseCode = o.warehouseCol ? String(row[o.warehouseCol - 1] || "").trim() : "";
                            whCode = warehouseCode && warehouse[warehouseCode] ? warehouseCode : warehouse.__default;
                            wh = warehouse[whCode] || {};
                            customer = customers.find(function (c) {
                              return c.code === customerCode;
                            });
                            maker = makers.find(function (m) {
                              return m.code === makerCode;
                            });
                            if (customer) {
                              _context23.n = 2;
                              break;
                            }
                            errorLog.push(i + 1 + "行目：得意先「" + customerCode + "」が見つかりません");
                            errorCount++;
                            orderSheet.getRangeByIndexes(i, 0, 1, o.deliveryCol).format.fill.color = "#FFCCCC";
                            return _context23.a(2, 0);
                          case 2:
                            if (!(logistics !== "在庫" && !maker)) {
                              _context23.n = 3;
                              break;
                            }
                            errorLog.push(i + 1 + "行目：発注先「" + makerCode + "」が見つかりません");
                            errorCount++;
                            orderSheet.getRangeByIndexes(i, 0, 1, o.deliveryCol).format.fill.color = "#FFCCCC";
                            return _context23.a(2, 0);
                          case 3:
                            _context23.p = 3;
                            itemIrr = itemIrrs.find(function (e) {
                              return e.customerCode === customerCode && e.productCode === productCode;
                            });
                            makerIrr = makerIrrs.find(function (e) {
                              return e.customerCode === customerCode && e.makerCode === makerCode;
                            });
                            irr = itemIrr || makerIrr || null;
                            deliveryDate = null, arrivalDate = null;
                            isWeekdayMode = cfg.calcMethod === "曜日";
                            if (!isWeekdayMode) {
                              _context23.n = 8;
                              break;
                            }
                            if (!(logistics === "在庫")) {
                              _context23.n = 4;
                              break;
                            }
                            deliveryDate = calcZaikoWeekday(now, customer, wh, holidays, irr);
                            _context23.n = 7;
                            break;
                          case 4:
                            if (!(logistics === "通過")) {
                              _context23.n = 5;
                              break;
                            }
                            r = calcTsuukaWeekday(now, maker, customer, wh, holidays, irr);
                            deliveryDate = r.deliveryDate;
                            arrivalDate = r.arrivalDate;
                            _context23.n = 7;
                            break;
                          case 5:
                            if (!(logistics === "直送")) {
                              _context23.n = 6;
                              break;
                            }
                            deliveryDate = calcChokusoWeekday(now, maker, customer, holidays, irr);
                            _context23.n = 7;
                            break;
                          case 6:
                            throw new Error("物流形態「" + logistics + "」が不正です");
                          case 7:
                            _context23.n = 12;
                            break;
                          case 8:
                            if (!(logistics === "在庫")) {
                              _context23.n = 9;
                              break;
                            }
                            deliveryDate = calcZaikoDays(now, customer, wh, holidays, irr);
                            _context23.n = 12;
                            break;
                          case 9:
                            if (!(logistics === "通過")) {
                              _context23.n = 10;
                              break;
                            }
                            _r = calcTsuukaDays(now, maker, customer, wh, holidays, irr);
                            deliveryDate = _r.deliveryDate;
                            arrivalDate = _r.arrivalDate;
                            _context23.n = 12;
                            break;
                          case 10:
                            if (!(logistics === "直送")) {
                              _context23.n = 11;
                              break;
                            }
                            deliveryDate = calcChokusoDays(now, maker, customer, holidays, irr);
                            _context23.n = 12;
                            break;
                          case 11:
                            throw new Error("物流形態「" + logistics + "」が不正です");
                          case 12:
                            orderSheet.getRangeByIndexes(i, 0, 1, o.deliveryCol).format.fill.clear();
                            orderSheet.getRangeByIndexes(i, o.deliveryCol - 1, 1, 1).values = [[formatDate(deliveryDate)]];
                            if (arrivalDate) orderSheet.getRangeByIndexes(i, o.arrivalCol - 1, 1, 1).values = [[formatDate(arrivalDate)]];
                            successCount++;
                            _context23.n = 14;
                            break;
                          case 13:
                            _context23.p = 13;
                            _t1 = _context23.v;
                            errorLog.push(i + 1 + "行目：" + _t1.message);
                            errorCount++;
                            orderSheet.getRangeByIndexes(i, 0, 1, o.deliveryCol).format.fill.color = "#FFCCCC";
                          case 14:
                            return _context23.a(2);
                        }
                      }, _loop, null, [[3, 13]]);
                    });
                    i = 1;
                  case 8:
                    if (!(i < rows.length)) {
                      _context24.n = 11;
                      break;
                    }
                    return _context24.d(_regeneratorValues(_loop()), 9);
                  case 9:
                    _ret = _context24.v;
                    if (!(_ret === 0)) {
                      _context24.n = 10;
                      break;
                    }
                    return _context24.a(3, 10);
                  case 10:
                    i++;
                    _context24.n = 8;
                    break;
                  case 11:
                    _context24.n = 12;
                    return context.sync();
                  case 12:
                    html = "✅ " + successCount + "件完了";
                    if (errorCount > 0) html += "　❌ " + errorCount + "件エラー<br><br>" + errorLog.join("<br>");
                    resultDetail.innerHTML = html;
                    resultBox.style.display = "block";
                    msg.innerHTML = "";
                  case 13:
                    return _context24.a(2);
                }
              }, _callee23);
            }));
            return function (_x19) {
              return _ref7.apply(this, arguments);
            };
          }());
        case 2:
          _context25.n = 4;
          break;
        case 3:
          _context25.p = 3;
          _t10 = _context25.v;
          msg.innerHTML = "❌ エラー：" + _t10.message;
        case 4:
          return _context25.a(2);
      }
    }, _callee24, null, [[1, 3]]);
  }));
  return _calcDelivery.apply(this, arguments);
}
function isBusinessDay(date, holidays) {
  var d = date.getUTCDay();
  return d !== 0 && d !== 6 && !holidays.includes(formatDate(date));
}
function calcChokusoWeekday(now, maker, customer, holidays, irr) {
  var base = getBaseDate(now, maker.deadline, holidays);
  var dayName = DAY_NAMES[base.getUTCDay()];
  var deliveryDayName = irr && irr.deliveryMap && irr.deliveryMap[dayName] ? irr.deliveryMap[dayName] : maker.deliveryMap[dayName];
  var delivery = getNextDeliveryDate(base, deliveryDayName);
  if (holidays.includes(formatDate(delivery))) {
    delivery.setUTCDate(delivery.getUTCDate() + 1);
    if (maker.holidayAdd > 0) delivery.setUTCDate(delivery.getUTCDate() + maker.holidayAdd);
    while (holidays.includes(formatDate(delivery))) delivery.setUTCDate(delivery.getUTCDate() + 1);
  }
  return adjustForCustomer(delivery, customer, holidays, irr);
}
function calcZaikoDays(now, customer, warehouse, holidays, irr) {
  var base = getBaseDate(now, warehouse["入荷"] && warehouse["入荷"]["締め時間"] || "14:00", holidays);
  var ship = nextBusinessDay(base, holidays);
  ship = adjustForWarehouse(ship, warehouse["納品"] || {}, holidays);
  if (irr && irr.days !== undefined) ship = addBusinessDays(ship, irr.days, holidays);
  return adjustForCustomer(ship, customer, holidays, irr);
}
function calcTsuukaDays(now, maker, customer, warehouse, holidays, irr) {
  var base = getBaseDate(now, maker.deadline, holidays);
  var days = irr && irr.days !== undefined ? irr.days : maker.days;
  var arrival = addBusinessDays(base, days, holidays);
  if (maker.holidayAdd > 0 && holidays.includes(formatDate(arrival))) arrival = addBusinessDays(arrival, maker.holidayAdd, holidays);
  var confirmedArrival = new Date(arrival);
  var ship = adjustForWarehouse(arrival, warehouse, "納品", holidays);
  ship = nextBusinessDay(ship, holidays);
  return {
    deliveryDate: adjustForCustomer(ship, customer, holidays, irr),
    arrivalDate: confirmedArrival
  };
}
function calcChokusoDays(now, maker, customer, holidays, irr) {
  var base = getBaseDate(now, maker.deadline, holidays);
  var days = irr && irr.days !== undefined ? irr.days : maker.days;
  var delivery = addBusinessDays(base, days, holidays);
  if (maker.holidayAdd > 0 && holidays.includes(formatDate(delivery))) delivery = addBusinessDays(delivery, maker.holidayAdd, holidays);
  return adjustForCustomer(delivery, customer, holidays, irr);
}
function calcZaikoWeekday(now, customer, warehouse, holidays, irr) {
  var base = getBaseDate(now, warehouse["入荷"] && warehouse["入荷"]["締め時間"] || "14:00", holidays);
  var ship = nextBusinessDay(base, holidays);
  ship = adjustForWarehouse(ship, warehouse["納品"] || {}, holidays);
  return adjustForCustomer(ship, customer, holidays, irr);
}
function calcTsuukaWeekday(now, maker, customer, warehouse, holidays, irr) {
  var base = getBaseDate(now, maker.deadline, holidays);
  var dayName = DAY_NAMES[base.getUTCDay()];
  var deliveryDayName2 = irr && irr.deliveryMap && irr.deliveryMap[dayName] ? irr.deliveryMap[dayName] : maker.deliveryMap[dayName];
  var arrival = getNextDeliveryDate(base, deliveryDayName2);
  if (holidays.includes(formatDate(arrival))) {
    arrival.setUTCDate(arrival.getUTCDate() + 1);
    if (maker.holidayAdd > 0) arrival.setUTCDate(arrival.getUTCDate() + maker.holidayAdd);
    while (holidays.includes(formatDate(arrival))) arrival.setUTCDate(arrival.getUTCDate() + 1);
  }
  var confirmedArrival = new Date(arrival);
  var ship = adjustForWarehouse(arrival, warehouse["納品"] || {}, holidays);
  ship = nextBusinessDay(ship, holidays);
  return {
    deliveryDate: adjustForCustomer(ship, customer, holidays, irr),
    arrivalDate: confirmedArrival
  };
}
function nextBusinessDay(date, holidays) {
  var next = new Date(date);
  next.setUTCDate(next.getUTCDate() + 1);
  var count = 0;
  while (!isBusinessDay(next, holidays)) {
    next.setUTCDate(next.getUTCDate() + 1);
    if (++count > 30) throw new Error("営業日が見つかりません（休日設定を確認してください）");
  }
  return next;
}
function addBusinessDays(date, days, holidays) {
  var d = new Date(date),
    added = 0,
    count = 0;
  while (added < days) {
    d.setUTCDate(d.getUTCDate() + 1);
    if (isBusinessDay(d, holidays)) added++;
    if (++count > 365) throw new Error("営業日加算が上限を超えました");
  }
  return d;
}
function getBaseDate(now, deadlineStr, holidays) {
  var base = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
  if (deadlineStr) {
    var parts = deadlineStr.toString().split(":");
    var dh = Number(parts[0]),
      dm = Number(parts[1] || 0);
    if (now.getHours() > dh || now.getHours() === dh && now.getMinutes() >= dm) base = nextBusinessDay(base, holidays);
  }
  return base;
}
function getNextDeliveryDate(baseDate, deliveryDayName) {
  var target = DAY_NAMES.indexOf(deliveryDayName);
  if (target === -1) throw new Error("納品曜日が未設定のメーカーがあります：「" + deliveryDayName + "」");
  var date = new Date(baseDate);
  date.setUTCDate(date.getUTCDate() + 1);
  var count = 0;
  while (date.getUTCDay() !== target) {
    date.setUTCDate(date.getUTCDate() + 1);
    if (++count > 14) throw new Error("納品曜日が見つかりません");
  }
  return date;
}
function adjustForCustomer(date, customer, holidays, irr) {
  var deliverable = irr && irr.deliverable ? irr.deliverable : customer.deliverable;
  var d = new Date(date),
    count = 0;
  while (!deliverable[DAY_NAMES[d.getUTCDay()]] || holidays.includes(formatDate(d))) {
    d.setUTCDate(d.getUTCDate() + 1);
    if (++count > 30) break;
  }
  return d;
}
function adjustForWarehouse(date, kubun, holidays) {
  var d = new Date(date),
    count = 0;
  while (true) {
    var dayName = DAY_NAMES[d.getUTCDay()];
    var isHol = holidays.includes(formatDate(d));
    if (isHol && !kubun["祝日"] || !kubun[dayName]) {
      d.setUTCDate(d.getUTCDate() + 1);
      if (++count > 30) break;
    } else break;
  }
  return d;
}
}();
// This entry needs to be wrapped in an IIFE because it needs to be in strict mode.
!function() {
"use strict";
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
// Module
var code = "<!DOCTYPE html>\n<html>\n<head>\n  <meta charset=\"UTF-8\" />\n  <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\n  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n  <title>DeliveryCalc</title>\n  <style>\n    * { box-sizing: border-box; margin: 0; padding: 0; }\n    body { font-family: \"Helvetica Neue\", Arial, sans-serif; font-size: 14px; color: #333; padding: 12px; background: #f5f5f5; }\n    #sideload-msg { display: none; }\n\n    #license-screen { background: white; border-radius: 8px; padding: 24px; text-align: center; }\n    .lock-icon { font-size: 48px; margin-bottom: 16px; }\n    .license-title { font-size: 18px; font-weight: bold; color: #217346; margin-bottom: 8px; }\n    .license-sub { font-size: 12px; color: #666; margin-bottom: 20px; }\n    .form-group { margin-bottom: 12px; text-align: left; }\n    .form-group label { display: block; font-size: 12px; color: #555; margin-bottom: 4px; }\n    .form-group input { width: 100%; padding: 8px 10px; border: 1px solid #ddd; border-radius: 4px; font-size: 13px; }\n    .btn-primary { width: 100%; padding: 10px; background: #217346; color: white; border: none; border-radius: 4px; font-size: 14px; font-weight: bold; cursor: pointer; margin-top: 8px; }\n    .btn-primary:hover { background: #1a5c38; }\n    .purchase-link { margin-top: 20px; padding: 14px; background: #f0f7f4; border-radius: 6px; }\n    .purchase-link p { font-size: 12px; color: #555; margin-bottom: 10px; }\n    .purchase-link a { display: inline-block; background: #217346; color: white; padding: 10px 24px; border-radius: 20px; font-size: 13px; font-weight: bold; text-decoration: none; }\n    .purchase-note { font-size: 11px; color: #777; margin-top: 8px; }\n    .msg { font-size: 12px; margin-top: 8px; color: #c00; min-height: 16px; }\n\n    #app-body { display: none; }\n\n    .method-bar { background: white; border-radius: 8px; padding: 12px 14px; margin-bottom: 10px; display: flex; align-items: center; gap: 10px; }\n    .method-label { font-size: 12px; color: #555; flex: 1; }\n    .method-btns { display: flex; gap: 4px; }\n    .method-btn { padding: 5px 14px; border: 1px solid #ddd; border-radius: 20px; font-size: 12px; background: white; cursor: pointer; color: #555; }\n    .method-btn.method-active { background: #217346; color: white; border-color: #217346; font-weight: bold; }\n\n    .card { background: white; border-radius: 8px; padding: 16px; margin-bottom: 10px; }\n    .card-title { font-size: 13px; font-weight: bold; color: #217346; margin-bottom: 12px; }\n    .btn-action { width: 100%; padding: 10px; background: #217346; color: white; border: none; border-radius: 4px; font-size: 13px; font-weight: bold; cursor: pointer; }\n    .btn-action:hover { background: #1a5c38; }\n    .btn-sub { width: 100%; padding: 8px; background: white; color: #217346; border: 1px solid #217346; border-radius: 4px; font-size: 12px; cursor: pointer; margin-top: 6px; }\n    .btn-sub:hover { background: #f0f7f4; }\n    .result-box { margin-top: 10px; padding: 10px; background: #f0f7f4; border-radius: 4px; font-size: 12px; line-height: 1.6; }\n    .info-text { font-size: 11px; color: #777; line-height: 1.5; margin-bottom: 10px; }\n\n    #settings-panel { display: none; margin-top: 12px; border-top: 1px solid #eee; padding-top: 12px; }\n    .s-section { margin-bottom: 14px; }\n    .s-section-title { font-size: 11px; font-weight: bold; color: #217346; margin-bottom: 8px; padding-bottom: 4px; border-bottom: 1px solid #e8e8e8; }\n    .s-row { display: flex; align-items: center; margin-bottom: 6px; gap: 8px; }\n    .s-row label { font-size: 11px; color: #555; width: 110px; flex-shrink: 0; }\n    .s-row input { flex: 1; padding: 4px 6px; border: 1px solid #ddd; border-radius: 3px; font-size: 12px; }\n    .s-note { font-size: 11px; color: #999; margin-bottom: 8px; }\n    .btn-s-save { padding: 7px 16px; background: #217346; color: white; border: none; border-radius: 4px; font-size: 12px; cursor: pointer; }\n    .btn-s-reset { padding: 7px 14px; background: white; color: #999; border: 1px solid #ddd; border-radius: 4px; font-size: 12px; cursor: pointer; margin-left: 6px; }\n    .msg-s { font-size: 11px; color: #217346; min-height: 14px; margin-top: 6px; }\n\n    .license-info-text { font-size: 11px; color: #555; line-height: 1.6; }\n    .btn-logout { padding: 6px 12px; background: white; color: #c00; border: 1px solid #c00; border-radius: 4px; font-size: 11px; cursor: pointer; margin-top: 8px; }\n    .btn-logout:hover { background: #fff0f0; }\n  </style>\n<" + "script src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\n</head>\n<body>\n  <div id=\"sideload-msg\"></div>\n\n  <!-- ライセンス画面 -->\n  <div id=\"license-screen\">\n    <div class=\"lock-icon\">🔐</div>\n    <div class=\"license-title\">DeliveryCalc</div>\n    <div class=\"license-sub\">ライセンスを認証してください</div>\n    <div class=\"form-group\">\n      <label>メールアドレス</label>\n      <input type=\"email\" id=\"license-email\" placeholder=\"例：your@email.com\" />\n    </div>\n    <div class=\"form-group\">\n      <label>ライセンスキー</label>\n      <input type=\"text\" id=\"license-key\" placeholder=\"例：DC-XXXX-XXXX-XXXX\" />\n    </div>\n    <button class=\"btn-primary\" id=\"btn-activate\">認証する</button>\n    <div class=\"msg\" id=\"msg-license\"></div>\n    <div class=\"purchase-link\">\n      <p>ライセンスをお持ちでない方はこちら</p>\n      <a href=\"https://buy.stripe.com/dRmdRb7Wk3vRfec2wZd7q01\" target=\"_blank\">🛒 今すぐ購入 ¥4,500/月</a>\n      <div class=\"purchase-note\">3ヶ月間有効・クレジットカード決済</div>\n    </div>\n  </div>\n\n  <!-- アプリ本体 -->\n  <div id=\"app-body\">\n\n    <!-- 演算方法トグル -->\n    <div class=\"method-bar\">\n      <span class=\"method-label\" id=\"method-label\">納期演算方法</span>\n      <div class=\"method-btns\">\n        <button class=\"method-btn\" id=\"btn-method-weekday\">配送曜日</button>\n        <button class=\"method-btn\" id=\"btn-method-days\">日数</button>\n      </div>\n    </div>\n\n    <!-- 納期一括算出 -->\n    <div class=\"card\">\n      <div class=\"card-title\">📦 納期を一括算出する</div>\n      <div class=\"info-text\">受注一覧シートの全行を一括処理します。</div>\n      <button class=\"btn-action\" id=\"btn-calc-delivery\">納期を一括算出する</button>\n      <div class=\"msg\" id=\"msg-delivery\"></div>\n      <div class=\"result-box\" id=\"result-box\" style=\"display:none;\">\n        <div id=\"result-detail\"></div>\n      </div>\n    </div>\n\n    <!-- マスタ管理 -->\n    <div class=\"card\">\n      <div class=\"card-title\">⚙️ マスタ管理</div>\n      <div class=\"info-text\">マスタシートを初期化します（既存データは削除）。<br>イレギュラーシートも同時に作成されます。</div>\n      <button class=\"btn-action\" id=\"btn-init-master\">マスタシートを初期化する</button>\n      <div class=\"msg\" id=\"msg-init\"></div>\n      <button class=\"btn-sub\" id=\"btn-toggle-settings\">⚙️ 列・シート設定を開く</button>\n\n      <!-- 設定パネル -->\n      <div id=\"settings-panel\">\n        <div class=\"s-section\">\n          <div class=\"s-section-title\">納期回答シート設定</div>\n          <div class=\"s-note\">列番号：A=1、B=2、C=3…</div>\n          <div class=\"s-row\"><label>シート名</label><input type=\"text\" id=\"s-order-sheet\" /></div>\n          <div class=\"s-row\"><label>得意先コード列</label><input type=\"number\" id=\"s-order-customer-col\" min=\"1\" /></div>\n          <div class=\"s-row\"><label>商品コード列</label><input type=\"number\" id=\"s-order-product-col\" min=\"1\" /></div>\n          <div class=\"s-row\"><label>発注先コード列</label><input type=\"number\" id=\"s-order-maker-col\" min=\"1\" /></div>\n          <div class=\"s-row\"><label>物流形態列</label><input type=\"number\" id=\"s-order-logistics-col\" min=\"1\" /></div>\n          <div class=\"s-row\"><label>倉庫コード列</label><input type=\"number\" id=\"s-order-warehouse-col\" min=\"1\" /></div>\n          <div class=\"s-row\"><label>入荷日出力列</label><input type=\"number\" id=\"s-order-arrival-col\" min=\"1\" /></div>\n          <div class=\"s-row\"><label>納期回答出力列</label><input type=\"number\" id=\"s-order-delivery-col\" min=\"1\" /></div>\n        </div>\n        <div class=\"s-section\">\n          <div class=\"s-section-title\">マスタシート名設定</div>\n          <div class=\"s-row\"><label>仕入先情報</label><input type=\"text\" id=\"s-maker-sheet\" /></div>\n          <div class=\"s-row\"><label>得意先情報</label><input type=\"text\" id=\"s-customer-sheet\" /></div>\n          <div class=\"s-row\"><label>倉庫情報</label><input type=\"text\" id=\"s-warehouse-sheet\" /></div>\n          <div class=\"s-row\"><label>祝日情報</label><input type=\"text\" id=\"s-holiday-sheet\" /></div>\n          <div class=\"s-row\"><label>商品イレギュラー</label><input type=\"text\" id=\"s-item-irr-sheet\" /></div>\n          <div class=\"s-row\"><label>仕入先イレギュラー</label><input type=\"text\" id=\"s-maker-irr-sheet\" /></div>\n        </div>\n        <div>\n          <button class=\"btn-s-save\" id=\"btn-save-settings\">保存</button>\n          <button class=\"btn-s-reset\" id=\"btn-reset-settings\">デフォルトに戻す</button>\n          <div class=\"msg-s\" id=\"msg-settings\"></div>\n        </div>\n      </div>\n    </div>\n\n    <!-- ライセンス情報 -->\n    <div class=\"card\">\n      <div class=\"card-title\">🔑 ライセンス情報</div>\n      <div class=\"license-info-text\" id=\"license-info\"></div>\n      <button class=\"btn-logout\" id=\"btn-logout\">ライセンスを解除する</button>\n    </div>\n\n  </div>\n  \n<" + "script>\nwindow.onerror = function(msg, src, line) {\n  document.getElementById('sideload-msg').style.display = 'block';\n  document.getElementById('sideload-msg').style.color = 'red';\n  document.getElementById('sideload-msg').textContent = 'エラー: ' + msg + ' (' + line + '行)';\n};\ndocument.addEventListener('DOMContentLoaded', function() {\n  document.getElementById('btn-activate').addEventListener('click', function() {\n    var msg = document.getElementById('msg-license');\n    msg.textContent = '認証中...';\n    setTimeout(function() {\n      if (typeof activateLicense === 'function') {\n        activateLicense();\n      } else {\n        msg.textContent = 'Office未初期化 - window.onerrorを確認';\n      }\n    }, 500);\n  });\n});\n<" + "/script></body>\n</html>";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map