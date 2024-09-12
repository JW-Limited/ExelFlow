//xlFlow Interpreter v0.12
// © JW Limited. All rights reserved.
// Author: Joe Valentino Lengefeld
// 
// This software, known as the "xlFlow Interpreter," is proprietary to JW Limited. 
// The code, in whole or in part, may not be copied, modified, distributed, or used 
// in any manner without prior written consent from JW Limited. 
// 
// The "xlFlow Interpreter" allows for the manipulation and transformation of Excel 
// files using a custom scripting engine, providing functionalities such as filtering, 
// transforming, modulating, and exporting data in various formats.
//
// DISCLAIMER:
// This software is provided "as is," without any express or implied warranties, 
// including but not limited to the implied warranties of merchantability and fitness 
// for a particular purpose. In no event shall JW Limited or its contributors be 
// held liable for any damages arising in any way from the use of this software.


"use strict";

function _extends() {
  return (
    (_extends = Object.assign
      ? Object.assign.bind()
      : function (n) {
          for (var e = 1; e < arguments.length; e++) {
            var t = arguments[e];
            for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]);
          }
          return n;
        }),
    _extends.apply(null, arguments)
  );
}
function _createForOfIteratorHelperLoose(r, e) {
  var t =
    ("undefined" != typeof Symbol && r[Symbol.iterator]) || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (
    Array.isArray(r) ||
    (t = _unsupportedIterableToArray(r)) ||
    (e && r && "number" == typeof r.length)
  ) {
    t && (r = t);
    var o = 0;
    return function () {
      return o >= r.length ? { done: !0 } : { done: !1, value: r[o++] };
    };
  }
  throw new TypeError(
    "Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."
  );
}
function _unsupportedIterableToArray(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return (
      "Object" === t && r.constructor && (t = r.constructor.name),
      "Map" === t || "Set" === t
        ? Array.from(r)
        : "Arguments" === t ||
          /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t)
        ? _arrayLikeToArray(r, a)
        : void 0
    );
  }
}
function _arrayLikeToArray(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function _regeneratorRuntime() {
  "use strict";
  /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/facebook/regenerator/blob/main/LICENSE */ _regeneratorRuntime =
    function _regeneratorRuntime() {
      return e;
    };
  var t,
    e = {},
    r = Object.prototype,
    n = r.hasOwnProperty,
    o =
      Object.defineProperty ||
      function (t, e, r) {
        t[e] = r.value;
      },
    i = "function" == typeof Symbol ? Symbol : {},
    a = i.iterator || "@@iterator",
    c = i.asyncIterator || "@@asyncIterator",
    u = i.toStringTag || "@@toStringTag";
  function define(t, e, r) {
    return (
      Object.defineProperty(t, e, {
        value: r,
        enumerable: !0,
        configurable: !0,
        writable: !0
      }),
      t[e]
    );
  }
  try {
    define({}, "");
  } catch (t) {
    define = function define(t, e, r) {
      return (t[e] = r);
    };
  }
  function wrap(t, e, r, n) {
    var i = e && e.prototype instanceof Generator ? e : Generator,
      a = Object.create(i.prototype),
      c = new Context(n || []);
    return o(a, "_invoke", { value: makeInvokeMethod(t, r, c) }), a;
  }
  function tryCatch(t, e, r) {
    try {
      return { type: "normal", arg: t.call(e, r) };
    } catch (t) {
      return { type: "throw", arg: t };
    }
  }
  e.wrap = wrap;
  var h = "suspendedStart",
    l = "suspendedYield",
    f = "executing",
    s = "completed",
    y = {};
  function Generator() {}
  function GeneratorFunction() {}
  function GeneratorFunctionPrototype() {}
  var p = {};
  define(p, a, function () {
    return this;
  });
  var d = Object.getPrototypeOf,
    v = d && d(d(values([])));
  v && v !== r && n.call(v, a) && (p = v);
  var g =
    (GeneratorFunctionPrototype.prototype =
    Generator.prototype =
      Object.create(p));
  function defineIteratorMethods(t) {
    ["next", "throw", "return"].forEach(function (e) {
      define(t, e, function (t) {
        return this._invoke(e, t);
      });
    });
  }
  function AsyncIterator(t, e) {
    function invoke(r, o, i, a) {
      var c = tryCatch(t[r], t, o);
      if ("throw" !== c.type) {
        var u = c.arg,
          h = u.value;
        return h && "object" == typeof h && n.call(h, "__await")
          ? e.resolve(h.__await).then(
              function (t) {
                invoke("next", t, i, a);
              },
              function (t) {
                invoke("throw", t, i, a);
              }
            )
          : e.resolve(h).then(
              function (t) {
                (u.value = t), i(u);
              },
              function (t) {
                return invoke("throw", t, i, a);
              }
            );
      }
      a(c.arg);
    }
    var r;
    o(this, "_invoke", {
      value: function value(t, n) {
        function callInvokeWithMethodAndArg() {
          return new e(function (e, r) {
            invoke(t, n, e, r);
          });
        }
        return (r = r
          ? r.then(callInvokeWithMethodAndArg, callInvokeWithMethodAndArg)
          : callInvokeWithMethodAndArg());
      }
    });
  }
  function makeInvokeMethod(e, r, n) {
    var o = h;
    return function (i, a) {
      if (o === f) throw Error("Generator is already running");
      if (o === s) {
        if ("throw" === i) throw a;
        return { value: t, done: !0 };
      }
      for (n.method = i, n.arg = a; ; ) {
        var c = n.delegate;
        if (c) {
          var u = maybeInvokeDelegate(c, n);
          if (u) {
            if (u === y) continue;
            return u;
          }
        }
        if ("next" === n.method) n.sent = n._sent = n.arg;
        else if ("throw" === n.method) {
          if (o === h) throw ((o = s), n.arg);
          n.dispatchException(n.arg);
        } else "return" === n.method && n.abrupt("return", n.arg);
        o = f;
        var p = tryCatch(e, r, n);
        if ("normal" === p.type) {
          if (((o = n.done ? s : l), p.arg === y)) continue;
          return { value: p.arg, done: n.done };
        }
        "throw" === p.type && ((o = s), (n.method = "throw"), (n.arg = p.arg));
      }
    };
  }
  function maybeInvokeDelegate(e, r) {
    var n = r.method,
      o = e.iterator[n];
    if (o === t)
      return (
        (r.delegate = null),
        ("throw" === n &&
          e.iterator["return"] &&
          ((r.method = "return"),
          (r.arg = t),
          maybeInvokeDelegate(e, r),
          "throw" === r.method)) ||
          ("return" !== n &&
            ((r.method = "throw"),
            (r.arg = new TypeError(
              "The iterator does not provide a '" + n + "' method"
            )))),
        y
      );
    var i = tryCatch(o, e.iterator, r.arg);
    if ("throw" === i.type)
      return (r.method = "throw"), (r.arg = i.arg), (r.delegate = null), y;
    var a = i.arg;
    return a
      ? a.done
        ? ((r[e.resultName] = a.value),
          (r.next = e.nextLoc),
          "return" !== r.method && ((r.method = "next"), (r.arg = t)),
          (r.delegate = null),
          y)
        : a
      : ((r.method = "throw"),
        (r.arg = new TypeError("iterator result is not an object")),
        (r.delegate = null),
        y);
  }
  function pushTryEntry(t) {
    var e = { tryLoc: t[0] };
    1 in t && (e.catchLoc = t[1]),
      2 in t && ((e.finallyLoc = t[2]), (e.afterLoc = t[3])),
      this.tryEntries.push(e);
  }
  function resetTryEntry(t) {
    var e = t.completion || {};
    (e.type = "normal"), delete e.arg, (t.completion = e);
  }
  function Context(t) {
    (this.tryEntries = [{ tryLoc: "root" }]),
      t.forEach(pushTryEntry, this),
      this.reset(!0);
  }
  function values(e) {
    if (e || "" === e) {
      var r = e[a];
      if (r) return r.call(e);
      if ("function" == typeof e.next) return e;
      if (!isNaN(e.length)) {
        var o = -1,
          i = function next() {
            for (; ++o < e.length; )
              if (n.call(e, o))
                return (next.value = e[o]), (next.done = !1), next;
            return (next.value = t), (next.done = !0), next;
          };
        return (i.next = i);
      }
    }
    throw new TypeError(typeof e + " is not iterable");
  }
  return (
    (GeneratorFunction.prototype = GeneratorFunctionPrototype),
    o(g, "constructor", {
      value: GeneratorFunctionPrototype,
      configurable: !0
    }),
    o(GeneratorFunctionPrototype, "constructor", {
      value: GeneratorFunction,
      configurable: !0
    }),
    (GeneratorFunction.displayName = define(
      GeneratorFunctionPrototype,
      u,
      "GeneratorFunction"
    )),
    (e.isGeneratorFunction = function (t) {
      var e = "function" == typeof t && t.constructor;
      return (
        !!e &&
        (e === GeneratorFunction ||
          "GeneratorFunction" === (e.displayName || e.name))
      );
    }),
    (e.mark = function (t) {
      return (
        Object.setPrototypeOf
          ? Object.setPrototypeOf(t, GeneratorFunctionPrototype)
          : ((t.__proto__ = GeneratorFunctionPrototype),
            define(t, u, "GeneratorFunction")),
        (t.prototype = Object.create(g)),
        t
      );
    }),
    (e.awrap = function (t) {
      return { __await: t };
    }),
    defineIteratorMethods(AsyncIterator.prototype),
    define(AsyncIterator.prototype, c, function () {
      return this;
    }),
    (e.AsyncIterator = AsyncIterator),
    (e.async = function (t, r, n, o, i) {
      void 0 === i && (i = Promise);
      var a = new AsyncIterator(wrap(t, r, n, o), i);
      return e.isGeneratorFunction(r)
        ? a
        : a.next().then(function (t) {
            return t.done ? t.value : a.next();
          });
    }),
    defineIteratorMethods(g),
    define(g, u, "Generator"),
    define(g, a, function () {
      return this;
    }),
    define(g, "toString", function () {
      return "[object Generator]";
    }),
    (e.keys = function (t) {
      var e = Object(t),
        r = [];
      for (var n in e) r.push(n);
      return (
        r.reverse(),
        function next() {
          for (; r.length; ) {
            var t = r.pop();
            if (t in e) return (next.value = t), (next.done = !1), next;
          }
          return (next.done = !0), next;
        }
      );
    }),
    (e.values = values),
    (Context.prototype = {
      constructor: Context,
      reset: function reset(e) {
        if (
          ((this.prev = 0),
          (this.next = 0),
          (this.sent = this._sent = t),
          (this.done = !1),
          (this.delegate = null),
          (this.method = "next"),
          (this.arg = t),
          this.tryEntries.forEach(resetTryEntry),
          !e)
        )
          for (var r in this)
            "t" === r.charAt(0) &&
              n.call(this, r) &&
              !isNaN(+r.slice(1)) &&
              (this[r] = t);
      },
      stop: function stop() {
        this.done = !0;
        var t = this.tryEntries[0].completion;
        if ("throw" === t.type) throw t.arg;
        return this.rval;
      },
      dispatchException: function dispatchException(e) {
        if (this.done) throw e;
        var r = this;
        function handle(n, o) {
          return (
            (a.type = "throw"),
            (a.arg = e),
            (r.next = n),
            o && ((r.method = "next"), (r.arg = t)),
            !!o
          );
        }
        for (var o = this.tryEntries.length - 1; o >= 0; --o) {
          var i = this.tryEntries[o],
            a = i.completion;
          if ("root" === i.tryLoc) return handle("end");
          if (i.tryLoc <= this.prev) {
            var c = n.call(i, "catchLoc"),
              u = n.call(i, "finallyLoc");
            if (c && u) {
              if (this.prev < i.catchLoc) return handle(i.catchLoc, !0);
              if (this.prev < i.finallyLoc) return handle(i.finallyLoc);
            } else if (c) {
              if (this.prev < i.catchLoc) return handle(i.catchLoc, !0);
            } else {
              if (!u) throw Error("try statement without catch or finally");
              if (this.prev < i.finallyLoc) return handle(i.finallyLoc);
            }
          }
        }
      },
      abrupt: function abrupt(t, e) {
        for (var r = this.tryEntries.length - 1; r >= 0; --r) {
          var o = this.tryEntries[r];
          if (
            o.tryLoc <= this.prev &&
            n.call(o, "finallyLoc") &&
            this.prev < o.finallyLoc
          ) {
            var i = o;
            break;
          }
        }
        i &&
          ("break" === t || "continue" === t) &&
          i.tryLoc <= e &&
          e <= i.finallyLoc &&
          (i = null);
        var a = i ? i.completion : {};
        return (
          (a.type = t),
          (a.arg = e),
          i
            ? ((this.method = "next"), (this.next = i.finallyLoc), y)
            : this.complete(a)
        );
      },
      complete: function complete(t, e) {
        if ("throw" === t.type) throw t.arg;
        return (
          "break" === t.type || "continue" === t.type
            ? (this.next = t.arg)
            : "return" === t.type
            ? ((this.rval = this.arg = t.arg),
              (this.method = "return"),
              (this.next = "end"))
            : "normal" === t.type && e && (this.next = e),
          y
        );
      },
      finish: function finish(t) {
        for (var e = this.tryEntries.length - 1; e >= 0; --e) {
          var r = this.tryEntries[e];
          if (r.finallyLoc === t)
            return this.complete(r.completion, r.afterLoc), resetTryEntry(r), y;
        }
      },
      catch: function _catch(t) {
        for (var e = this.tryEntries.length - 1; e >= 0; --e) {
          var r = this.tryEntries[e];
          if (r.tryLoc === t) {
            var n = r.completion;
            if ("throw" === n.type) {
              var o = n.arg;
              resetTryEntry(r);
            }
            return o;
          }
        }
        throw Error("illegal catch attempt");
      },
      delegateYield: function delegateYield(e, r, n) {
        return (
          (this.delegate = { iterator: values(e), resultName: r, nextLoc: n }),
          "next" === this.method && (this.arg = t),
          y
        );
      }
    }),
    e
  );
}
function asyncGeneratorStep(n, t, e, r, o, a, c) {
  try {
    var i = n[a](c),
      u = i.value;
  } catch (n) {
    return void e(n);
  }
  i.done ? t(u) : Promise.resolve(u).then(r, o);
}
function _asyncToGenerator(n) {
  return function () {
    var t = this,
      e = arguments;
    return new Promise(function (r, o) {
      var a = n.apply(t, e);
      function _next(n) {
        asyncGeneratorStep(a, r, o, _next, _throw, "next", n);
      }
      function _throw(n) {
        asyncGeneratorStep(a, r, o, _next, _throw, "throw", n);
      }
      _next(void 0);
    });
  };
}
// ExcelFlow Interpreter
// Requires xlsx.full.min.js to be loaded
var ExcelFlowInterpreter = /*#__PURE__*/ (function () {
  function ExcelFlowInterpreter() {
    this.builtInFunctions = {
      sum: function sum(args) {
        return args.reduce(function (a, b) {
          return a + b;
        }, 0);
      },
      avg: function avg(args) {
        return (
          args.reduce(function (a, b) {
            return a + b;
          }, 0) / args.length
        );
      },
      min: function min(args) {
        return Math.min.apply(Math, args);
      },
      max: function max(args) {
        return Math.max.apply(Math, args);
      },
      len: function len(args) {
        return args.length;
      },
      if: function _if(args) {
        return args[0] ? args[1] : args[2];
      }
      // Add more built-in functions as needed
    };
    this.workbook = null;
    this.variables = {};
    this.currentSheet = null;
  }
  var _proto = ExcelFlowInterpreter.prototype;
  _proto.execute = /*#__PURE__*/ (function () {
    var _execute = _asyncToGenerator(
      /*#__PURE__*/ _regeneratorRuntime().mark(function _callee(script) {
        var lines, i;
        return _regeneratorRuntime().wrap(
          function _callee$(_context) {
            while (1)
              switch ((_context.prev = _context.next)) {
                case 0:
                  lines = script
                    .split("\n")
                    .map(function (line) {
                      return line.trim();
                    })
                    .filter(function (line) {
                      return line && !line.startsWith("#");
                    });
                  i = 0;
                case 2:
                  if (!(i < lines.length)) {
                    _context.next = 15;
                    break;
                  }
                  this.currentLine = i + 1;
                  _context.prev = 4;
                  _context.next = 7;
                  return this.executeLine(lines[i]);
                case 7:
                  _context.next = 12;
                  break;
                case 9:
                  _context.prev = 9;
                  _context.t0 = _context["catch"](4);
                  throw new Error(
                    "Error on line " +
                      this.currentLine +
                      ": " +
                      _context.t0.message +
                      "\nLine content: " +
                      lines[i]
                  );
                case 12:
                  i++;
                  _context.next = 2;
                  break;
                case 15:
                case "end":
                  return _context.stop();
              }
          },
          _callee,
          this,
          [[4, 9]]
        );
      })
    );
    function execute(_x) {
      return _execute.apply(this, arguments);
    }
    return execute;
  })();
  _proto.executeLine = /*#__PURE__*/ (function () {
    var _executeLine = _asyncToGenerator(
      /*#__PURE__*/ _regeneratorRuntime().mark(function _callee2(line) {
        var _line$split, command, args;
        return _regeneratorRuntime().wrap(
          function _callee2$(_context2) {
            while (1)
              switch ((_context2.prev = _context2.next)) {
                case 0:
                  (_line$split = line.split(" ")),
                    (command = _line$split[0]),
                    (args = _line$split.slice(1));
                  _context2.t0 = command;
                  _context2.next =
                    _context2.t0 === "load"
                      ? 4
                      : _context2.t0 === "filter"
                      ? 7
                      : _context2.t0 === "transform"
                      ? 9
                      : _context2.t0 === "modulate"
                      ? 11
                      : _context2.t0 === "export"
                      ? 13
                      : 16;
                  break;
                case 4:
                  _context2.next = 6;
                  return this.loadFile(args[0].replace(/"/g, ""));
                case 6:
                  return _context2.abrupt("break", 21);
                case 7:
                  this.applyFilter(args.join(" "));
                  return _context2.abrupt("break", 21);
                case 9:
                  this.applyTransform(args.join(" "));
                  return _context2.abrupt("break", 21);
                case 11:
                  this.applyModulate(args.join(" "));
                  return _context2.abrupt("break", 21);
                case 13:
                  _context2.next = 15;
                  return this.exportFile(args.join(" "));
                case 15:
                  return _context2.abrupt("break", 21);
                case 16:
                  if (!line.includes("=")) {
                    _context2.next = 20;
                    break;
                  }
                  this.assignVariable(line);
                  _context2.next = 21;
                  break;
                case 20:
                  throw new Error("Unknown command: " + command);
                case 21:
                case "end":
                  return _context2.stop();
              }
          },
          _callee2,
          this
        );
      })
    );
    function executeLine(_x2) {
      return _executeLine.apply(this, arguments);
    }
    return executeLine;
  })();
  _proto.loadFile = /*#__PURE__*/ (function () {
    var _loadFile = _asyncToGenerator(
      /*#__PURE__*/ _regeneratorRuntime().mark(function _callee3(filename) {
        var response, data;
        return _regeneratorRuntime().wrap(
          function _callee3$(_context3) {
            while (1)
              switch ((_context3.prev = _context3.next)) {
                case 0:
                  _context3.next = 2;
                  return fetch(filename);
                case 2:
                  response = _context3.sent;
                  _context3.next = 5;
                  return response.arrayBuffer();
                case 5:
                  data = _context3.sent;
                  this.workbook = XLSX.read(data, {
                    type: "array"
                  });
                  this.currentSheet =
                    this.workbook.Sheets[this.workbook.SheetNames[0]];
                case 8:
                case "end":
                  return _context3.stop();
              }
          },
          _callee3,
          this
        );
      })
    );
    function loadFile(_x3) {
      return _loadFile.apply(this, arguments);
    }
    return loadFile;
  })();
  _proto.applyFilter = function applyFilter(condition) {
    var range = XLSX.utils.decode_range(this.currentSheet["!ref"]);
    for (var row = range.s.r; row <= range.e.r; row++) {
      if (!this.evaluateCondition(condition, row)) {
        for (var col = range.s.c; col <= range.e.c; col++) {
          var cellAddress = XLSX.utils.encode_cell({
            r: row,
            c: col
          });
          delete this.currentSheet[cellAddress];
        }
      }
    }
    this.currentSheet["!ref"] = XLSX.utils.encode_range(this.getActualRange());
  };
  _proto.applyTransform = function applyTransform(operation) {
    var _operation$split = operation.split(" "),
      subCommand = _operation$split[0],
      args = _operation$split.slice(1);
    switch (subCommand) {
      case "add_column":
        this.addColumn(args[0], args[1]);
        break;
      case "map":
        this.mapColumn(args[0], args.slice(1).join(" "));
        break;
      default:
        throw new Error(
          "Unknown transform operation: " +
            subCommand +
            ". Valid operations are 'add_column' and 'map'."
        );
    }
  };
  _proto.applyModulate = function applyModulate(func) {
    var _func$split = func.split("|>"),
      column = _func$split[0],
      operation = _func$split[1];
    var range = this.getColumnRange(column.trim());
    for (var row = range.s.r; row <= range.e.r; row++) {
      var cellAddress = XLSX.utils.encode_cell({
        r: row,
        c: range.s.c
      });
      var cell = this.currentSheet[cellAddress];
      if (cell) {
        cell.v = this.evaluateFunction(operation.trim(), cell.v);
        cell.w = cell.v.toString();
      }
    }
  };
  _proto.exportFile = /*#__PURE__*/ (function () {
    var _exportFile = _asyncToGenerator(
      /*#__PURE__*/ _regeneratorRuntime().mark(function _callee4(options) {
        var exportOptions, wb, _iterator, _step, sheet, ws;
        return _regeneratorRuntime().wrap(
          function _callee4$(_context4) {
            while (1)
              switch ((_context4.prev = _context4.next)) {
                case 0:
                  exportOptions = this.parseExportOptions(options);
                  wb = XLSX.utils.book_new();
                  for (
                    _iterator = _createForOfIteratorHelperLoose(
                      exportOptions.sheets,
                      true
                    );
                    !(_step = _iterator()).done;

                  ) {
                    sheet = _step.value;
                    ws = XLSX.utils.aoa_to_sheet(
                      this.getSheetData(sheet.range)
                    );
                    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
                  }
                  XLSX.writeFile(
                    wb,
                    exportOptions.filename,
                    _extends(
                      {
                        bookType: "csv"
                      },
                      exportOptions.options
                    )
                  );
                case 4:
                case "end":
                  return _context4.stop();
              }
          },
          _callee4,
          this
        );
      })
    );
    function exportFile(_x4) {
      return _exportFile.apply(this, arguments);
    }
    return exportFile;
  })();
  _proto.assignVariable = function assignVariable(assignment) {
    var _assignment$split$map = assignment.split("=").map(function (s) {
        return s.trim();
      }),
      varName = _assignment$split$map[0],
      value = _assignment$split$map[1];
    this.variables[varName] = this.evaluateExpression(value);
  };
  _proto.evaluateCondition = function evaluateCondition(condition, row) {
    var tokens = this.tokenizeCondition(condition);
    return this.parseCondition(tokens, row);
  };
  _proto.tokenizeCondition = function tokenizeCondition(condition) {
    var regex =
      /(\|\||&&|==|!=|>=|<=|>|<|\(|\)|!|\$?\w+(?::\w+)?|\d+(\.\d+)?|'[^']*'|"[^"]*")/g;
    return condition.match(regex);
  };
  _proto.parseCondition = function parseCondition(tokens, row) {
    var output = [];
    var operators = [];
    var precedence = {
      "||": 1,
      "&&": 2,
      "==": 3,
      "!=": 3,
      ">": 4,
      "<": 4,
      ">=": 4,
      "<=": 4,
      "!": 5
    };
    for (
      var _iterator2 = _createForOfIteratorHelperLoose(tokens, true), _step2;
      !(_step2 = _iterator2()).done;

    ) {
      var token = _step2.value;
      if (token === "(") {
        operators.push(token);
      } else if (token === ")") {
        while (operators.length && operators[operators.length - 1] !== "(") {
          output.push(operators.pop());
        }
        operators.pop(); // Remove the '('
      } else if (token in precedence) {
        while (
          operators.length &&
          precedence[operators[operators.length - 1]] >= precedence[token]
        ) {
          output.push(operators.pop());
        }
        operators.push(token);
      } else {
        output.push(this.evaluateToken(token, row));
      }
    }
    while (operators.length) {
      output.push(operators.pop());
    }
    return this.evaluateRPN(output);
  };
  _proto.evaluateToken = function evaluateToken(token, row) {
    if (token.startsWith("$")) {
      return this.variables[token];
    } else if (/^\d+(\.\d+)?$/.test(token)) {
      return parseFloat(token);
    } else if (/^['"].*['"]$/.test(token)) {
      return token.slice(1, -1);
    } else {
      return this.getCellValue(token, row);
    }
  };
  _proto.formatDate = function formatDate(date, format) {
    var d = new Date(date);
    var formatTokens = {
      YYYY: d.getFullYear(),
      MM: String(d.getMonth() + 1).padStart(2, "0"),
      DD: String(d.getDate()).padStart(2, "0"),
      HH: String(d.getHours()).padStart(2, "0"),
      mm: String(d.getMinutes()).padStart(2, "0"),
      ss: String(d.getSeconds()).padStart(2, "0")
    };
    return format.replace(/YYYY|MM|DD|HH|mm|ss/g, function (match) {
      return formatTokens[match];
    });
  };
  _proto.evaluateRPN = function evaluateRPN(tokens) {
    var stack = [];
    for (
      var _iterator3 = _createForOfIteratorHelperLoose(tokens, true), _step3;
      !(_step3 = _iterator3()).done;

    ) {
      var token = _step3.value;
      if (typeof token === "number" || Array.isArray(token)) {
        stack.push(token);
      } else if (typeof token === "string") {
        if (token in this.builtInFunctions) {
          var args = stack.pop();
          stack.push(this.builtInFunctions[token](args));
        } else {
          var b = stack.pop();
          var a = stack.pop();
          switch (token) {
            case "+":
              stack.push(this.add(a, b));
              break;
            case "-":
              stack.push(this.subtract(a, b));
              break;
            case "*":
              stack.push(this.multiply(a, b));
              break;
            case "/":
              stack.push(this.divide(a, b));
              break;
            case "^":
              stack.push(this.power(a, b));
              break;
            case "==":
              stack.push(this.equal(a, b));
              break;
            case "!=":
              stack.push(this.notEqual(a, b));
              break;
            case ">":
              stack.push(this.greaterThan(a, b));
              break;
            case "<":
              stack.push(this.lessThan(a, b));
              break;
            case ">=":
              stack.push(this.greaterThanOrEqual(a, b));
              break;
            case "<=":
              stack.push(this.lessThanOrEqual(a, b));
              break;
            case "&&":
              stack.push(this.and(a, b));
              break;
            case "||":
              stack.push(this.or(a, b));
              break;
            default:
              throw new Error("Unknown operator: " + token);
          }
        }
      }
    }
    return stack[0];
  };

  // Helper functions for mathematical operations
  _proto.add = function add(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return val + b[i];
      });
    }
    return a + b;
  };
  _proto.subtract = function subtract(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return val - b[i];
      });
    }
    return a - b;
  };
  _proto.multiply = function multiply(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return val * b[i];
      });
    }
    return a * b;
  };
  _proto.divide = function divide(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return val / b[i];
      });
    }
    return a / b;
  };
  _proto.power = function power(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return Math.pow(val, b[i]);
      });
    }
    return Math.pow(a, b);
  };

  // Helper functions for comparison operations
  _proto.equal = function equal(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return val == b[i];
      });
    }
    return a == b;
  };
  _proto.notEqual = function notEqual(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return val != b[i];
      });
    }
    return a != b;
  };
  _proto.greaterThan = function greaterThan(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return val > b[i];
      });
    }
    return a > b;
  };
  _proto.lessThan = function lessThan(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return val < b[i];
      });
    }
    return a < b;
  };
  _proto.greaterThanOrEqual = function greaterThanOrEqual(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return val >= b[i];
      });
    }
    return a >= b;
  };
  _proto.lessThanOrEqual = function lessThanOrEqual(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return val <= b[i];
      });
    }
    return a <= b;
  };

  // Helper functions for logical operations
  _proto.and = function and(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return val && b[i];
      });
    }
    return a && b;
  };
  _proto.or = function or(a, b) {
    if (Array.isArray(a) && Array.isArray(b)) {
      return a.map(function (val, i) {
        return val || b[i];
      });
    }
    return a || b;
  };
  _proto.tokenizeExpression = function tokenizeExpression(expr) {
    var regex =
      /(\$\w+|[A-Z]+\d+(?::[A-Z]+\d+)?|\d+(?:\.\d+)?|[-+*/()^]|[<>=!]=?|&&|\|\||[a-zA-Z_]\w*(?=\()|\S)/g;
    return expr.match(regex);
  };
  _proto.evaluateExpression = function evaluateExpression(expr) {
    try {
      var tokens = this.tokenizeExpression(expr);
      return this.parseExpression(tokens);
    } catch (error) {
      throw new Error(
        'Error evaluating expression "' + expr + '": ' + error.message
      );
    }
  };
  _proto.getCellValue = function getCellValue(cellAddress, row) {
    if (row === void 0) {
      row = null;
    }
    if (row !== null) {
      var col = XLSX.utils.decode_col(cellAddress);
      cellAddress = XLSX.utils.encode_cell({
        r: row,
        c: col
      });
    }
    var cell = this.currentSheet[cellAddress];
    if (!cell) {
      throw new Error(
        "Cell " + cellAddress + " not found in the current sheet."
      );
    }
    return cell.v;
  };
  _proto.getRangeValues = function getRangeValues(range) {
    var rangeObj = XLSX.utils.decode_range(range);
    var values = [];
    for (var row = rangeObj.s.r; row <= rangeObj.e.r; row++) {
      for (var col = rangeObj.s.c; col <= rangeObj.e.c; col++) {
        var cellAddress = XLSX.utils.encode_cell({
          r: row,
          c: col
        });
        values.push(this.getCellValue(cellAddress));
      }
    }
    return values;
  };
  _proto.addColumn = function addColumn(column, header) {
    var range = XLSX.utils.decode_range(this.currentSheet["!ref"]);
    var newColIndex = range.e.c + 1;
    this.currentSheet[
      XLSX.utils.encode_cell({
        r: range.s.r,
        c: newColIndex
      })
    ] = {
      t: "s",
      v: header
    };
    range.e.c++;
    this.currentSheet["!ref"] = XLSX.utils.encode_range(range);
  };
  _proto.mapColumn = function mapColumn(column, operation) {
    var range = this.getColumnRange(column);
    for (var row = range.s.r + 1; row <= range.e.r; row++) {
      var cellAddress = XLSX.utils.encode_cell({
        r: row,
        c: range.s.c
      });
      var result = this.evaluateMapOperation(operation, row);
      this.currentSheet[cellAddress] = {
        t: "s",
        v: result
      };
    }
  };
  _proto.evaluateMapOperation = function evaluateMapOperation(operation, row) {
    if (operation.startsWith("pattern_match")) {
      return this.evaluatePatternMatch(operation, row);
    } else {
      return this.evaluateExpression(operation);
    }
  };
  _proto.getColumnRange = function getColumnRange(column) {
    var range = XLSX.utils.decode_range(this.currentSheet["!ref"]);
    var col = XLSX.utils.decode_col(column);
    return {
      s: {
        r: range.s.r,
        c: col
      },
      e: {
        r: range.e.r,
        c: col
      }
    };
  };
  _proto.getActualRange = function getActualRange() {
    var range = XLSX.utils.decode_range(this.currentSheet["!ref"]);
    var minRow = range.e.r,
      maxRow = range.s.r,
      minCol = range.e.c,
      maxCol = range.s.c;
    for (var row = range.s.r; row <= range.e.r; row++) {
      for (var col = range.s.c; col <= range.e.c; col++) {
        var cellAddress = XLSX.utils.encode_cell({
          r: row,
          c: col
        });
        if (this.currentSheet[cellAddress]) {
          minRow = Math.min(minRow, row);
          maxRow = Math.max(maxRow, row);
          minCol = Math.min(minCol, col);
          maxCol = Math.max(maxCol, col);
        }
      }
    }
    return {
      s: {
        r: minRow,
        c: minCol
      },
      e: {
        r: maxRow,
        c: maxCol
      }
    };
  };
  _proto.getSheetData = function getSheetData(range) {
    var rangeObj = XLSX.utils.decode_range(range);
    var data = [];
    for (var row = rangeObj.s.r; row <= rangeObj.e.r; row++) {
      var rowData = [];
      for (var col = rangeObj.s.c; col <= rangeObj.e.c; col++) {
        var cellAddress = XLSX.utils.encode_cell({
          r: row,
          c: col
        });
        rowData.push(this.getCellValue(cellAddress));
      }
      data.push(rowData);
    }
    return data;
  };
  _proto.parseExportOptions = function parseExportOptions(options) {
    var parsedOptions = JSON.parse(options.replace(/'/g, '"'));
    if (
      !parsedOptions.filename ||
      !parsedOptions.sheets ||
      !Array.isArray(parsedOptions.sheets)
    ) {
      throw new Error(
        "Invalid export options: filename and sheets array are required"
      );
    }
    parsedOptions.sheets = parsedOptions.sheets.map(function (sheet) {
      if (!sheet.name || !sheet.range) {
        throw new Error("Each sheet must have a name and range");
      }
      return {
        name: sheet.name,
        range: XLSX.utils.decode_range(sheet.range)
      };
    });
    parsedOptions.options = parsedOptions.options || {};
    parsedOptions.options.delimiter = parsedOptions.options.delimiter || ",";
    parsedOptions.options.include_headers =
      parsedOptions.options.include_headers !== false;
    parsedOptions.options.date_format =
      parsedOptions.options.date_format || "YYYY-MM-DD";
    return parsedOptions;
  };
  _proto.parseExpression = function parseExpression(tokens) {
    var output = [];
    var operators = [];
    var precedence = {
      "^": 4,
      "*": 3,
      "/": 3,
      "+": 2,
      "-": 2,
      "==": 1,
      "!=": 1,
      ">": 1,
      "<": 1,
      ">=": 1,
      "<=": 1,
      "&&": 0,
      "||": 0
    };
    for (var i = 0; i < tokens.length; i++) {
      var token = tokens[i];
      if (token === "(") {
        operators.push(token);
      } else if (token === ")") {
        while (operators.length && operators[operators.length - 1] !== "(") {
          output.push(operators.pop());
        }
        operators.pop(); // Remove the '('
        // Check if the '(' was preceded by a function name
        if (
          operators.length &&
          /^[a-zA-Z_]\w*$/.test(operators[operators.length - 1])
        ) {
          output.push(operators.pop());
        }
      } else if (token in precedence) {
        while (
          operators.length &&
          precedence[operators[operators.length - 1]] >= precedence[token]
        ) {
          output.push(operators.pop());
        }
        operators.push(token);
      } else if (/^[a-zA-Z_]\w*$/.test(token) && tokens[i + 1] === "(") {
        // Function call
        operators.push(token);
      } else {
        output.push(this.evaluateToken(token));
      }
    }
    while (operators.length) {
      output.push(operators.pop());
    }
    return this.evaluateRPN(output);
  };
  _proto.evaluatePatternMatch = function evaluatePatternMatch(operation, row) {
    var _operation$split2 = operation.split(" "),
      column = _operation$split2[1],
      casesAndDefault = _operation$split2.slice(2);
    var value = this.getCellValue(column, row);
    for (var i = 0; i < casesAndDefault.length - 1; i += 2) {
      var pattern = new RegExp(casesAndDefault[i].replace(/^\/|\/$/g, ""));
      if (pattern.test(value)) {
        return this.evaluateExpression(
          casesAndDefault[i + 1].replace("=>", "").trim()
        );
      }
    }

    // Handle default case
    var defaultCase = casesAndDefault[casesAndDefault.length - 1];
    return defaultCase.startsWith("_")
      ? this.evaluateExpression(defaultCase.replace("=>", "").trim())
      : "Other";
  };
  return ExcelFlowInterpreter;
})();
