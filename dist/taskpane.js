/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/taskpane/ConditionalFormats.ts":
/*!********************************************!*\
  !*** ./src/taskpane/ConditionalFormats.ts ***!
  \********************************************/
/***/ (function(__unused_webpack_module, exports) {



Object.defineProperty(exports, "__esModule", ({
  value: true
}));
exports.JsonEnumToCellValueOperator = exports.CellValueOperatorToJsonEnum = exports.enumCellValueOperator = exports.enumConditionalFormatTextOperator = exports.enumConditionalFormatType = void 0;
var enumConditionalFormatType;
(function (enumConditionalFormatType) {
  enumConditionalFormatType["ColorScale"] = "ColorScale";
  enumConditionalFormatType["ContainsText"] = "ConstainsText";
  enumConditionalFormatType["CellValue"] = "CellValue";
  enumConditionalFormatType["Custom"] = "Custom";
})(enumConditionalFormatType = exports.enumConditionalFormatType || (exports.enumConditionalFormatType = {}));
var enumConditionalFormatTextOperator;
(function (enumConditionalFormatTextOperator) {
  enumConditionalFormatTextOperator["Contains"] = "Contains";
  enumConditionalFormatTextOperator["NotContains"] = "NotContains";
})(enumConditionalFormatTextOperator = exports.enumConditionalFormatTextOperator || (exports.enumConditionalFormatTextOperator = {}));
var enumCellValueOperator;
(function (enumCellValueOperator) {
  enumCellValueOperator["LT"] = "LessThan";
  enumCellValueOperator["GT"] = "GreaterThan";
  enumCellValueOperator["EQ"] = "Equal";
  enumCellValueOperator["BETWEEN"] = "Between";
})(enumCellValueOperator = exports.enumCellValueOperator || (exports.enumCellValueOperator = {}));
function CellValueOperatorToJsonEnum(op) {
  var retval = enumCellValueOperator.EQ;
  switch (op) {
    case Excel.ConditionalCellValueOperator.equalTo:
      retval = enumCellValueOperator.EQ;
      break;
    case Excel.ConditionalCellValueOperator.between:
      retval = enumCellValueOperator.BETWEEN;
      break;
    case Excel.ConditionalCellValueOperator.greaterThan:
      retval = enumCellValueOperator.GT;
      break;
    case Excel.ConditionalCellValueOperator.lessThan:
      retval = enumCellValueOperator.LT;
      break;
    default:
      retval = enumCellValueOperator.EQ;
      break;
  }
  return retval;
}
exports.CellValueOperatorToJsonEnum = CellValueOperatorToJsonEnum;
function JsonEnumToCellValueOperator(op) {
  var retval = Excel.ConditionalCellValueOperator.equalTo;
  switch (op) {
    case enumCellValueOperator.EQ:
      retval = Excel.ConditionalCellValueOperator.equalTo;
      break;
    case enumCellValueOperator.BETWEEN:
      retval = Excel.ConditionalCellValueOperator.between;
      break;
    case enumCellValueOperator.GT:
      retval = Excel.ConditionalCellValueOperator.greaterThan;
      break;
    case enumCellValueOperator.LT:
      retval = Excel.ConditionalCellValueOperator.lessThan;
      break;
    default:
      retval = Excel.ConditionalCellValueOperator.equalTo;
      break;
  }
  return retval;
}
exports.JsonEnumToCellValueOperator = JsonEnumToCellValueOperator;

/***/ }),

/***/ "./src/taskpane/jsonConfigUtils.ts":
/*!*****************************************!*\
  !*** ./src/taskpane/jsonConfigUtils.ts ***!
  \*****************************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {



Object.defineProperty(exports, "__esModule", ({
  value: true
}));
exports.JsonConfigUtils = void 0;
var ConditionalFormats_1 = __webpack_require__(/*! ./ConditionalFormats */ "./src/taskpane/ConditionalFormats.ts");
var JsonConfigUtils = /** @class */function () {
  function JsonConfigUtils() {
    this.columns = [];
    this.conditionalFormats = [];
    this.filterDefinitions = [];
    // this.json = value;
  }

  JsonConfigUtils.prototype.getValue = function () {
    return JSON.stringify(this, null, 2);
  };
  //#region Column Definitions
  JsonConfigUtils.prototype.addColumn = function (col) {
    this.columns.push(col);
    return col;
  };
  JsonConfigUtils.prototype.addColumnByName = function (columnName, isMandatory, horizontalAlignment, verticalAlignment, columnWidth, indentLevel, style, numberFormat, visible, autosizeColumn, searchFor, replaceWith) {
    if (visible === void 0) {
      visible = true;
    }
    if (autosizeColumn === void 0) {
      autosizeColumn = false;
    }
    if (searchFor === void 0) {
      searchFor = "";
    }
    if (replaceWith === void 0) {
      replaceWith = "";
    }
    var col = {
      columnName: columnName,
      isMandatory: isMandatory,
      horizontalAlignment: horizontalAlignment,
      verticalAlignment: verticalAlignment,
      columnWidth: columnWidth,
      indentLevel: indentLevel,
      style: style,
      numberFormat: numberFormat,
      visible: visible,
      autosizeColumn: autosizeColumn,
      searchFor: searchFor,
      replaceWith: replaceWith
    };
    this.columns.push(col);
    return col;
  };
  JsonConfigUtils.prototype.convertColumnDefinitionsToJson = function () {
    // Convert the array of ColumnDefinition objects to a JSON string
    var json = JSON.stringify(this.columns);
    return json;
  };
  JsonConfigUtils.prototype.convertToHorizontalAlignment = function (value) {
    switch (value) {
      case "Center":
        return Excel.HorizontalAlignment.center;
        break;
      case "Left":
        return Excel.HorizontalAlignment.left;
        break;
      case "Right":
        return Excel.HorizontalAlignment.right;
        break;
      case "Justify":
        return Excel.HorizontalAlignment.justify;
        break;
      case "General":
        return Excel.HorizontalAlignment.general;
        break;
      default:
        return Excel.HorizontalAlignment.general;
        break;
    }
  };
  JsonConfigUtils.prototype.convertToVerticalAlignment = function (value) {
    switch (value) {
      case "Bottom":
        return Excel.VerticalAlignment.bottom;
      case "Center":
        return Excel.VerticalAlignment.center;
      case "Distributed":
        return Excel.VerticalAlignment.distributed;
      case "Justify":
        return Excel.VerticalAlignment.justify;
      case "Top":
        return Excel.VerticalAlignment.top;
      default:
        return Excel.VerticalAlignment.top;
    }
  };
  //#endregion
  //#region ConditionalFormats Definitions
  JsonConfigUtils.prototype.addConditionalFormat = function (cf) {
    this.conditionalFormats.push(cf);
  };
  JsonConfigUtils.prototype.addConditionalFormatByName = function (friendlyName, columnName, countOccurrences, warnIfOccurrencesGTZero, type, style, fontColor, fillColor, colorScaleColorMinimum, colorScaleColorMaximum, containsTextSearch, containsTextOperator, cellValueFormula1, cellValueFormula2, cellValueOperator, customFormula) {
    var cf = {
      FriendlyName: friendlyName,
      ColumnName: columnName,
      CountOccurrences: countOccurrences,
      WarnIfOccurrencesGTZero: warnIfOccurrencesGTZero,
      Type: type,
      Style: style,
      FontColor: fontColor,
      FillColor: fillColor,
      ColorScaleColorMinimum: colorScaleColorMinimum,
      ColorScaleColorMaximum: colorScaleColorMaximum,
      ContainsTextSearch: containsTextSearch,
      ContainsTextOperator: containsTextOperator,
      CellValueFormula1: cellValueFormula1,
      CellValueFormula2: cellValueFormula2,
      CellValueOperator: cellValueOperator,
      CustomFormula: customFormula
    };
    this.conditionalFormats.push(cf);
  };
  JsonConfigUtils.prototype.addConditionalFormatColorScale = function (friendlyName, columnName, countOccurrences, warnIfOccurrencesGTZero, style, fontColor, fillColor, colorScaleColorMinimum, colorScaleColorMaximum) {
    var cf = {
      FriendlyName: friendlyName,
      ColumnName: columnName,
      CountOccurrences: countOccurrences,
      WarnIfOccurrencesGTZero: warnIfOccurrencesGTZero,
      Type: ConditionalFormats_1.enumConditionalFormatType.ColorScale,
      Style: style,
      FontColor: fontColor,
      FillColor: fillColor,
      ColorScaleColorMinimum: colorScaleColorMinimum,
      ColorScaleColorMaximum: colorScaleColorMaximum,
      ContainsTextSearch: null,
      ContainsTextOperator: null,
      CellValueFormula1: null,
      CellValueFormula2: null,
      CellValueOperator: null,
      CustomFormula: null
    };
    this.conditionalFormats.push(cf);
  };
  JsonConfigUtils.prototype.addConditionalFormatContainsText = function (friendlyName, columnName, countOccurrences, warnIfOccurrencesGTZero, style, fontColor, fillColor, containsTextSearch, containsTextOperator) {
    var cf = {
      FriendlyName: friendlyName,
      ColumnName: columnName,
      CountOccurrences: countOccurrences,
      WarnIfOccurrencesGTZero: warnIfOccurrencesGTZero,
      Type: ConditionalFormats_1.enumConditionalFormatType.ContainsText,
      Style: style,
      FontColor: fontColor,
      FillColor: fillColor,
      ColorScaleColorMinimum: null,
      ColorScaleColorMaximum: null,
      ContainsTextSearch: containsTextSearch,
      ContainsTextOperator: containsTextOperator,
      CellValueFormula1: null,
      CellValueFormula2: null,
      CellValueOperator: null,
      CustomFormula: null
    };
    this.conditionalFormats.push(cf);
  };
  JsonConfigUtils.prototype.addConditionalFormatCellValue = function (friendlyName, ColumnName, countOccurrences, warnIfOccurrencesGTZero, style, fontColor, fillColor, cellValueFormula1, cellValueFormula2, cellValueOperator) {
    var cf = {
      FriendlyName: friendlyName,
      ColumnName: ColumnName,
      CountOccurrences: countOccurrences,
      WarnIfOccurrencesGTZero: warnIfOccurrencesGTZero,
      Type: ConditionalFormats_1.enumConditionalFormatType.CellValue,
      Style: style,
      FontColor: fontColor,
      FillColor: fillColor,
      ColorScaleColorMinimum: null,
      ColorScaleColorMaximum: null,
      ContainsTextSearch: null,
      ContainsTextOperator: null,
      CellValueFormula1: cellValueFormula1,
      CellValueFormula2: cellValueFormula2,
      CellValueOperator: cellValueOperator,
      CustomFormula: null
    };
    this.conditionalFormats.push(cf);
  };
  JsonConfigUtils.prototype.addConditionalFormatCustom = function (friendlyName, columnName, countOccurrences, warnIfOccurrencesGTZero, style, fontColor, fillColor, customFormula) {
    var cf = {
      FriendlyName: friendlyName,
      ColumnName: columnName,
      CountOccurrences: countOccurrences,
      WarnIfOccurrencesGTZero: warnIfOccurrencesGTZero,
      Type: ConditionalFormats_1.enumConditionalFormatType.Custom,
      Style: style,
      FontColor: fontColor,
      FillColor: fillColor,
      ColorScaleColorMinimum: null,
      ColorScaleColorMaximum: null,
      ContainsTextSearch: null,
      ContainsTextOperator: null,
      CellValueFormula1: null,
      CellValueFormula2: null,
      CellValueOperator: null,
      CustomFormula: customFormula
    };
    this.conditionalFormats.push(cf);
  };
  //#endregion
  //#region Add filter Condition 
  JsonConfigUtils.prototype.addFilterCondition = function (columnName, value, key) {
    var newFilter = {
      FilterActiveByUIOnly: false,
      FilterKey: key,
      FilterValue: value,
      ColumnName: columnName,
      FriendlyName: "".concat(columnName, " filter ").concat(value)
    };
    this.filterDefinitions.push(newFilter);
    return newFilter;
  };
  return JsonConfigUtils;
}();
exports.JsonConfigUtils = JsonConfigUtils;

/***/ }),

/***/ "./src/taskpane/myConsole.ts":
/*!***********************************!*\
  !*** ./src/taskpane/myConsole.ts ***!
  \***********************************/
/***/ (function(__unused_webpack_module, exports) {



Object.defineProperty(exports, "__esModule", ({
  value: true
}));
exports.myConsole = void 0;
var myConsole = /** @class */function () {
  function myConsole() {
    myConsole.count = 0;
  }
  myConsole.addCounter = function () {
    myConsole.count++;
  };
  myConsole.log = function (message) {
    console.log(message);
    myConsole.addCounter();
    myConsole.addRow(myConsole.count.toString(), message);
  };
  myConsole.addRow = function (first, second) {
    var myConsole = document.querySelector('#myConsole');
    var newRow = document.createElement('div');
    newRow.classList.add('row');
    var firstCell = document.createElement('div');
    firstCell.classList.add('col-2');
    firstCell.innerHTML = "<small> ".concat(first, "</small>");
    var secondCell = document.createElement('div');
    secondCell.classList.add('col-10');
    secondCell.innerHTML = "<small>".concat(second, "</small>");
    newRow.appendChild(firstCell);
    newRow.appendChild(secondCell);
    myConsole.appendChild(newRow);
  };
  myConsole.reset = function () {
    myConsole.count = 0;
    var consoleDiv = document.querySelector('#myConsole');
    consoleDiv.innerHTML = "";
  };
  myConsole.count = 0;
  return myConsole;
}();
exports.myConsole = myConsole;

/***/ }),

/***/ "./src/taskpane/taskpane.ts":
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.ts ***!
  \**********************************/
/***/ (function(module, exports, __webpack_require__) {



/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
var __awaiter = this && this.__awaiter || function (thisArg, _arguments, P, generator) {
  function adopt(value) {
    return value instanceof P ? value : new P(function (resolve) {
      resolve(value);
    });
  }
  return new (P || (P = Promise))(function (resolve, reject) {
    function fulfilled(value) {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    }
    function rejected(value) {
      try {
        step(generator["throw"](value));
      } catch (e) {
        reject(e);
      }
    }
    function step(result) {
      result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected);
    }
    step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
};
var __generator = this && this.__generator || function (thisArg, body) {
  var _ = {
      label: 0,
      sent: function () {
        if (t[0] & 1) throw t[1];
        return t[1];
      },
      trys: [],
      ops: []
    },
    f,
    y,
    t,
    g;
  return g = {
    next: verb(0),
    "throw": verb(1),
    "return": verb(2)
  }, typeof Symbol === "function" && (g[Symbol.iterator] = function () {
    return this;
  }), g;
  function verb(n) {
    return function (v) {
      return step([n, v]);
    };
  }
  function step(op) {
    if (f) throw new TypeError("Generator is already executing.");
    while (g && (g = 0, op[0] && (_ = 0)), _) try {
      if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
      if (y = 0, t) op = [op[0] & 2, t.value];
      switch (op[0]) {
        case 0:
        case 1:
          t = op;
          break;
        case 4:
          _.label++;
          return {
            value: op[1],
            done: false
          };
        case 5:
          _.label++;
          y = op[1];
          op = [0];
          continue;
        case 7:
          op = _.ops.pop();
          _.trys.pop();
          continue;
        default:
          if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
            _ = 0;
            continue;
          }
          if (op[0] === 3 && (!t || op[1] > t[0] && op[1] < t[3])) {
            _.label = op[1];
            break;
          }
          if (op[0] === 6 && _.label < t[1]) {
            _.label = t[1];
            t = op;
            break;
          }
          if (t && _.label < t[2]) {
            _.label = t[2];
            _.ops.push(op);
            break;
          }
          if (t[2]) _.ops.pop();
          _.trys.pop();
          continue;
      }
      op = body.call(thisArg, _);
    } catch (e) {
      op = [6, e];
      y = 0;
    } finally {
      f = t = 0;
    }
    if (op[0] & 5) throw op[1];
    return {
      value: op[0] ? op[1] : void 0,
      done: true
    };
  }
};
Object.defineProperty(exports, "__esModule", ({
  value: true
}));
exports.run = exports.testConfig = exports.testJsonFile = exports.createConfig = void 0;
var ConditionalFormats_1 = __webpack_require__(/*! ./ConditionalFormats */ "./src/taskpane/ConditionalFormats.ts");
var jsonConfigUtils_1 = __webpack_require__(/*! ./jsonConfigUtils */ "./src/taskpane/jsonConfigUtils.ts");
var myConsole_1 = __webpack_require__(/*! ./myConsole */ "./src/taskpane/myConsole.ts");
/* global console, document, Excel, Office */
module.exports = ctx;
var ctx;
var tbl;
var sheet;
var tblRange;
var jsonConfigUtils;
//#region Constants
// Organizer table styles and checkbox
var organizerTableStyle = "TableStyleLight10";
var attendeeTableStyle = "TableStyleLight13";
var organizerTabColor = "#FFA500"; //orange
var attendeeTabColor = "#ADD8E6"; //light blue
//#endregion
//#region JSON properties of Callog
var jsonLog = "\n  [  \n    {    \n      \"columnName\": \"ModifiedDate\",   \n      \"isMandatory\": \"true\", \n      \"horizontalAlignment\": \"Center\",    \n      \"verticalAlignment\": \"Bottom\",    \n      \"columnWidth\": 180,    \n      \"indentLevel\": 1,    \n      \"style\": \"Neutral\",\n      \"numberFormat\": \"MM/dd/yyyy HH:mm:ss\"  \n    },  \n    {    \n      \"columnName\": \"Age\",  \n      \"isMandatory\": \"false\",  \n      \"horizontalAlignment\": \"center\",    \n      \"verticalAlignment\": \"middle\",    \n      \"columnWidth\": 80,    \n      \"indentLevel\": 1,    \n      \"style\": \"italic\",  \n      \"numberFormat\": \"MM/dd/yyyy HH:mm:ss\"  \n    }\n  ]";
//#endregion
//#region Properties
var _totalTblRows = -1;
function totalTblRows() {
  return __awaiter(this, void 0, void 0, function () {
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          if (!(_totalTblRows <= 0)) return [3 /*break*/, 2];
          tblRange.load(["rowCount"]);
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          _totalTblRows = tblRange.rowCount;
          _a.label = 2;
        case 2:
          return [2 /*return*/, _totalTblRows];
      }
    });
  });
}
// dropdown Type of CDL log
function typeCDL() {
  var selectElement = document.getElementById('typeCDL');
  return selectElement.value;
}
function isOrganizer() {
  return __awaiter(this, void 0, void 0, function () {
    var checkbox;
    return __generator(this, function (_a) {
      checkbox = document.getElementById("isOrganizer");
      return [2 /*return*/, checkbox.checked];
    });
  });
}
function warn1KRows() {
  return __awaiter(this, void 0, void 0, function () {
    var checkbox;
    return __generator(this, function (_a) {
      checkbox = document.getElementById("warn1KRows");
      return [2 /*return*/, checkbox.checked];
    });
  });
}
function hideLessRelevants() {
  return __awaiter(this, void 0, void 0, function () {
    var checkbox;
    return __generator(this, function (_a) {
      checkbox = document.getElementById("hideLessRelevants");
      return [2 /*return*/, checkbox.checked];
    });
  });
}
//#endregion
//#region Helper methods
function showSpinner(show) {
  return __awaiter(this, void 0, void 0, function () {
    var element;
    return __generator(this, function (_a) {
      element = document.getElementById("spinner");
      if (show) {
        element.classList.remove("invisible");
      } else {
        element.classList.add("invisible");
      }
      return [2 /*return*/];
    });
  });
}

function AddMessage(message) {
  var p = document.getElementById("statusMessage");
  p.textContent = message;
  myConsole_1.myConsole.log(message);
}
var enumTypeAnalysis;
(function (enumTypeAnalysis) {
  enumTypeAnalysis[enumTypeAnalysis["Warning"] = 0] = "Warning";
  enumTypeAnalysis[enumTypeAnalysis["Action"] = 1] = "Action";
  enumTypeAnalysis[enumTypeAnalysis["Danger"] = 2] = "Danger";
  enumTypeAnalysis[enumTypeAnalysis["Success"] = 3] = "Success";
})(enumTypeAnalysis || (enumTypeAnalysis = {}));
function addAnalysisInfo(title, badge, message, smallfooter, typeanalysis) {
  return __awaiter(this, void 0, void 0, function () {
    var analysisDiv, aElement, divElement, h5Element, spanElement, pElement, smallElement;
    return __generator(this, function (_a) {
      analysisDiv = document.getElementById("analysisInfo");
      aElement = document.createElement("a");
      switch (typeanalysis) {
        case enumTypeAnalysis.Warning:
          aElement.classList.add("list-group-item", "list-group-item-action", "list-group-item-warning");
          break;
        case enumTypeAnalysis.Action:
          aElement.classList.add("list-group-item", "list-group-item-action");
          break;
        case enumTypeAnalysis.Danger:
          aElement.classList.add("list-group-item", "list-group-item-action", "list-group-item-danger");
          break;
        case enumTypeAnalysis.Success:
          aElement.classList.add("list-group-item", "list-group-item-action", "list-group-item-success");
          break;
      }
      divElement = document.createElement("div");
      divElement.classList.add("d-flex", "w-100", "justify-content-between");
      h5Element = document.createElement("h5");
      h5Element.classList.add("mb-1");
      h5Element.innerText = title;
      spanElement = document.createElement("span");
      spanElement.classList.add("badge", "badge-primary", "badge-pill");
      if (badge == 0) {
        spanElement.classList.add("invisible");
      }
      spanElement.innerText = badge.toString();
      divElement.appendChild(h5Element);
      divElement.appendChild(spanElement);
      pElement = document.createElement("p");
      pElement.classList.add("mb-1");
      pElement.innerText = message;
      smallElement = document.createElement("small");
      smallElement.innerText = smallfooter;
      // Add the child elements to the <a> element
      aElement.appendChild(divElement);
      aElement.appendChild(pElement);
      aElement.appendChild(smallElement);
      // Add the <a> element to the analysis div
      analysisDiv.appendChild(aElement);
      AddMessage(message);
      return [2 /*return*/];
    });
  });
}

function resetAnalysisInfo() {
  return __awaiter(this, void 0, void 0, function () {
    var analysisInfoDiv;
    return __generator(this, function (_a) {
      analysisInfoDiv = document.getElementById('analysisInfo');
      analysisInfoDiv.innerHTML = '';
      return [2 /*return*/];
    });
  });
}

function freezeColumns(columnName) {
  return __awaiter(this, void 0, void 0, function () {
    var column, error_1;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          _a.trys.push([0, 2,, 3]);
          AddMessage("Freezing upto column ".concat(columnName, " (3 hardcoded for now)"));
          column = tbl.columns.getItem(columnName);
          sheet.freezePanes.freezeColumns(column.index);
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          AddMessage("Column ".concat(columnName, " Frozen successfully"));
          return [3 /*break*/, 3];
        case 2:
          error_1 = _a.sent();
          AddMessage("Error freezing column ".concat(columnName, ": ").concat(error_1));
          return [3 /*break*/, 3];
        case 3:
          return [2 /*return*/];
      }
    });
  });
}
//#endregion
//#region Init OfficeJS
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    jsonConfigUtils = new jsonConfigUtils_1.JsonConfigUtils();
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("run").onclick = run;
    document.getElementById("createConfig").onclick = createConfig;
    document.getElementById("testConfig").onclick = testConfig;
    document.getElementById("testJsonFile").onclick = testJsonFile;
    $(function () {
      $('[data-toggle="tooltip"]').tooltip();
    });
    resetAnalysisInfo();
  }
});
//#endregion
//#region WorkSheet Custom Properties
function resetCustomProperties() {
  return __awaiter(this, void 0, void 0, function () {
    var key, customProperty;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          sheet.customProperties.load();
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          // Iterate over all custom properties and delete them
          for (key in sheet.customProperties.items) {
            if (sheet.customProperties.items.hasOwnProperty(key)) {
              customProperty = sheet.customProperties.items[key];
              customProperty.delete();
            }
          }
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}

function addCustomProperty(key, value) {
  return __awaiter(this, void 0, void 0, function () {
    var cp;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          sheet.load(["customProperties"]);
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          sheet.customProperties.add(key, value);
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _a.sent();
          sheet.customProperties.load(["key", "value", "type"]);
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _a.sent();
          cp = sheet.customProperties.getItemOrNullObject(key);
          return [4 /*yield*/, ctx.sync()];
        case 4:
          _a.sent();
          return [2 /*return*/, cp];
      }
    });
  });
}
function GetCustomPropertyValue(key) {
  var _a;
  return __awaiter(this, void 0, void 0, function () {
    var cp;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          sheet.load(["customProperties"]);
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _b.sent();
          sheet.customProperties.load(["items"]);
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _b.sent();
          cp = sheet.customProperties.getItemOrNullObject(key);
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _b.sent();
          cp.load(["key", "value", "type"]);
          return [4 /*yield*/, ctx.sync()];
        case 4:
          _b.sent();
          return [2 /*return*/, (_a = cp.value) !== null && _a !== void 0 ? _a : ""];
      }
    });
  });
}
//#endregion
//#region Create Table
function CreateTable(context, keepFormats) {
  if (keepFormats === void 0) {
    keepFormats = false;
  }
  return __awaiter(this, void 0, void 0, function () {
    var tblCount, range, _isOrganizer, error_2;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          _a.trys.push([0, 15,, 16]);
          sheet.load(["tables"]);
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          tblCount = sheet.tables.getCount();
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _a.sent();
          if (tblCount.value > 1) {
            addAnalysisInfo("CreateTable Error", tblCount.value, "There is more than 1 table in the current worksheet. This is not supported. ", "None, or 1 table only supported", enumTypeAnalysis.Danger);
            return [2 /*return*/, false];
          }
          if (!(tblCount.value == 1)) return [3 /*break*/, 8];
          tbl = sheet.tables.getItemAt(0);
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _a.sent();
          tblRange = tbl.getRange();
          return [4 /*yield*/, ctx.sync()];
        case 4:
          _a.sent();
          if (keepFormats) {
            AddMessage("1 table found in current worksheet. Keeping formats");
            return [2 /*return*/, true];
          }
          AddMessage("1 table found in current worksheet. Clearing formats");
          tblRange.clear("Formats");
          return [4 /*yield*/, ctx.sync()];
        case 5:
          _a.sent();
          tbl.convertToRange();
          return [4 /*yield*/, ctx.sync()];
        case 6:
          _a.sent();
          resetCustomProperties();
          return [4 /*yield*/, context.sync()];
        case 7:
          _a.sent();
          AddMessage("Table cleared!");
          _a.label = 8;
        case 8:
          range = sheet.getUsedRange();
          range.load("address");
          return [4 /*yield*/, ctx.sync()];
        case 9:
          _a.sent();
          tbl = sheet.tables.add(range, true /* hasHeaders */);
          return [4 /*yield*/, ctx.sync()];
        case 10:
          _a.sent();
          tblRange = tbl.getRange();
          return [4 /*yield*/, context.sync()];
        case 11:
          _a.sent();
          if (keepFormats) return [2 /*return*/, true]; //just create table and leave
          tblRange.clear("Formats");
          resetCustomProperties();
          return [4 /*yield*/, context.sync()];
        case 12:
          _a.sent();
          return [4 /*yield*/, isOrganizer()];
        case 13:
          _isOrganizer = _a.sent();
          if (_isOrganizer) {
            tbl.style = organizerTableStyle;
            sheet.tabColor = organizerTabColor;
            addCustomProperty("Organizer", "true");
          } else {
            tbl.style = attendeeTableStyle;
            sheet.tabColor = attendeeTabColor;
            addCustomProperty("Organizer", "false");
          }
          tbl.load('tableStyle');
          tbl.columns.load();
          tblRange = tbl.getRange();
          return [4 /*yield*/, context.sync()];
        case 14:
          _a.sent();
          AddMessage("Table creation succeeded");
          return [2 /*return*/, true];
        case 15:
          error_2 = _a.sent();
          console.error(error_2);
          addAnalysisInfo("create Table", 0, "Error creating table ".concat(error_2), "Create Table", enumTypeAnalysis.Danger);
          return [2 /*return*/, false];
        case 16:
          return [2 /*return*/];
      }
    });
  });
}
//#endregion 
//#region Filters methods
function ClearFilters(ColumnName, value) {
  return __awaiter(this, void 0, void 0, function () {
    return __generator(this, function (_a) {
      tbl.clearFilters();
      AddMessage("Table filters cleared");
      return [2 /*return*/];
    });
  });
}

function SetFilter(ColumnName, value) {
  return __awaiter(this, void 0, void 0, function () {
    var columnFilter;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          tbl.columns.load();
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          columnFilter = tbl.columns.getItemOrNullObject(ColumnName).filter;
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _a.sent();
          columnFilter.apply({
            filterOn: Excel.FilterOn.values,
            values: [value]
          });
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _a.sent();
          AddMessage("Column ".concat(ColumnName, " Filtered for value ").concat(value));
          return [2 /*return*/];
      }
    });
  });
}

function FilterIgnorable(value) {
  return __awaiter(this, void 0, void 0, function () {
    var ignorableFilter;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          tbl.columns.load();
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          ignorableFilter = tbl.columns.getItemOrNullObject("Ignorable").filter;
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _a.sent();
          ignorableFilter.apply({
            filterOn: Excel.FilterOn.values,
            values: [value]
          });
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _a.sent();
          AddMessage("Cells filtered. ");
          return [2 /*return*/];
      }
    });
  });
}
//#endregion
//#region FormatDate (Legacy to remove)
function FormatDateColumn(context, columnName) {
  return __awaiter(this, void 0, void 0, function () {
    var col, colRange, criteria, rangeWidth;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          col = tbl.columns.getItemOrNullObject(columnName);
          colRange = col.getDataBodyRange();
          return [4 /*yield*/, context.sync()];
        case 1:
          _a.sent();
          criteria = {
            completeMatch: false,
            matchCase: true /* Ignore case when comparing strings. */
          };

          colRange.replaceAll("Z", "", criteria);
          return [4 /*yield*/, context.sync()];
        case 2:
          _a.sent();
          // Apply horizontal alignment as "center", vertical alignment as "bottom" and wrap text as "false"
          colRange.format.horizontalAlignment = "Center";
          colRange.format.verticalAlignment = "Bottom";
          colRange.format.wrapText = false;
          rangeWidth = colRange.getEntireColumn();
          rangeWidth.load("format");
          return [4 /*yield*/, context.sync()];
        case 3:
          _a.sent();
          rangeWidth.format.autofitColumns();
          return [4 /*yield*/, context.sync()];
        case 4:
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}
//#endregion
//#region Highlights methods
function HighlightIgnorable() {
  return __awaiter(this, void 0, void 0, function () {
    var colRange, conditionalFormat;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          colRange = tbl.columns.getItemOrNullObject("Ignorable").getDataBodyRange();
          conditionalFormat = colRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
          conditionalFormat.textComparison.format.font.color = "blue";
          conditionalFormat.textComparison.format.fill.color = "#ADD8E6";
          conditionalFormat.textComparison.rule = {
            operator: Excel.ConditionalTextOperator.contains,
            text: "TRUE"
          };
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          return [4 /*yield*/, CountFilterOccurrences(conditionalFormat.getRange())];
        case 2:
          return [2 /*return*/, _a.sent()];
      }
    });
  });
}
function HighlightApptSequence(context) {
  return __awaiter(this, void 0, void 0, function () {
    var col, range, conditionalFormat, criteria;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          col = tbl.columns.getItemOrNullObject("ApptSequence");
          range = col.getDataBodyRange();
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
          criteria = {
            minimum: {
              formula: null,
              type: Excel.ConditionalFormatColorCriterionType.lowestValue,
              color: "white"
            },
            // midpoint: { formula: "50", type: Excel.ConditionalFormatColorCriterionType.percent, color: "yellow" },
            maximum: {
              formula: null,
              type: Excel.ConditionalFormatColorCriterionType.highestValue,
              color: "green"
            }
          };
          conditionalFormat.colorScale.criteria = criteria;
          return [4 /*yield*/, context.sync()];
        case 2:
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}

function HighlightCRA(context) {
  return __awaiter(this, void 0, void 0, function () {
    var colRange, conditionalFormat;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          colRange = tbl.columns.getItemOrNullObject("Client").getRange();
          conditionalFormat = colRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
          // Color the font of every cell containing "Delayed".
          conditionalFormat.textComparison.format.font.color = "red";
          conditionalFormat.textComparison.rule = {
            operator: Excel.ConditionalTextOperator.contains,
            text: "CRA:CalendarRepairAssistant"
          };
          return [4 /*yield*/, context.sync()];
        case 1:
          _a.sent();
          return [4 /*yield*/, CountFilterOccurrences(conditionalFormat.getRange())];
        case 2:
          return [2 /*return*/, _a.sent()];
      }
    });
  });
}
function HighLightCreates(context) {
  return __awaiter(this, void 0, void 0, function () {
    var colRange, conditionalFormat, r, n;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          colRange = tbl.columns.getItemOrNullObject("Trigger").getRange();
          conditionalFormat = colRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
          // Color the font of every cell containing "Delayed".
          conditionalFormat.textComparison.format.fill.color = "Green";
          conditionalFormat.textComparison.rule = {
            operator: Excel.ConditionalTextOperator.contains,
            text: "Create"
          };
          return [4 /*yield*/, context.sync()];
        case 1:
          _a.sent();
          r = conditionalFormat.getRange();
          return [4 /*yield*/, CountFilterOccurrences(r)];
        case 2:
          n = _a.sent();
          return [2 /*return*/, n];
      }
    });
  });
}
//#endregion
//#region Count filter Occurrences
function CountFilterOccurrences(filterRange) {
  return __awaiter(this, void 0, void 0, function () {
    var affectedRange, rowCount;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          // var visibleTableRange: Excel.RangeView = tbl.getDataBodyRange().getVisibleView();
          // visibleTableRange.load(["rowCount"]); await ctx.sync();
          // var rows:number = visibleTableRange.rowCount;
          // console.log(`rows Filtered  %d`, rows);
          return [4 /*yield*/, ctx.sync()];
        case 1:
          // var visibleTableRange: Excel.RangeView = tbl.getDataBodyRange().getVisibleView();
          // visibleTableRange.load(["rowCount"]); await ctx.sync();
          // var rows:number = visibleTableRange.rowCount;
          // console.log(`rows Filtered  %d`, rows);
          _a.sent();
          return [4 /*yield*/, filterRange.getIntersectionOrNullObject(tbl.getRange())];
        case 2:
          affectedRange = _a.sent();
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _a.sent();
          affectedRange.load(["rowCount"]);
          return [4 /*yield*/, ctx.sync()];
        case 4:
          _a.sent();
          rowCount = affectedRange ? affectedRange.rowCount : 0;
          return [2 /*return*/, rowCount];
      }
    });
  });
}
//#endregion
//#region json Methods
function isJSONString(str) {
  try {
    var jsonObj = JSON.parse(str);
    return typeof jsonObj === "object" && jsonObj !== null;
  } catch (e) {
    return false;
  }
}
function applyJsonConfig(json, hideLessRelevants) {
  if (hideLessRelevants === void 0) {
    hideLessRelevants = false;
  }
  return __awaiter(this, void 0, void 0, function () {
    var jsonArray, retval, _a, _b, _c;
    return __generator(this, function (_d) {
      switch (_d.label) {
        case 0:
          retval = true;
          if (typeof json === "string") {
            if (!isJSONString(json)) {
              throw new Error("Invalid JSON string");
            }
            jsonArray = JSON.parse(json);
          } else if (typeof json === "object") {
            jsonArray = json;
          } else {
            throw new Error("Parameter must be a string or an object");
          }
          if (typeof jsonArray != "object") {
            throw new Error("Input JSON is not an Object.");
          }
          _a = retval;
          if (!_a) return [3 /*break*/, 2];
          return [4 /*yield*/, applyJsonColDefinitions(jsonArray, hideLessRelevants)];
        case 1:
          _a = _d.sent();
          _d.label = 2;
        case 2:
          retval = _a;
          _b = retval;
          if (!_b) return [3 /*break*/, 4];
          return [4 /*yield*/, applyJSONHighlights(jsonArray)];
        case 3:
          _b = _d.sent();
          _d.label = 4;
        case 4:
          retval = _b;
          _c = retval;
          if (!_c) return [3 /*break*/, 6];
          return [4 /*yield*/, applyJSONFilters(jsonArray)];
        case 5:
          _c = _d.sent();
          _d.label = 6;
        case 6:
          retval = _c;
          AddMessage("JSON Config Applied successfully!");
          return [2 /*return*/, retval];
      }
    });
  });
}
function applyJsonColDefinitions(jsonArray, hideLessRelevants) {
  if (hideLessRelevants === void 0) {
    hideLessRelevants = false;
  }
  return __awaiter(this, void 0, void 0, function () {
    var _i, _a, element, tblCol, tblColRange, tblColFormat, criteria, error_3;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          _b.trys.push([0, 9,, 10]);
          _i = 0, _a = jsonArray.columns;
          _b.label = 1;
        case 1:
          if (!(_i < _a.length)) return [3 /*break*/, 8];
          element = _a[_i];
          if (element.columnName == undefined || element.columnName == "") {
            AddMessage("Skipping json element as it is undefined ColumnName");
            return [3 /*break*/, 7];
          }
          tblCol = tbl.columns.getItemOrNullObject(element.columnName);
          tblCol.load(["isNullObject"]);
          ctx.trackedObjects.add([tblCol]);
          ;
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _b.sent();
          if (tblCol.isNullObject) {
            AddMessage("Column Name does not exist: ".concat(element.columnName));
            if (element.isMandatory !== undefined && element.isMandatory !== "" && element.isMandatory == "true") {
              AddMessage("isMandatory: ".concat(element.isMandatory, " Continuing to next column..."));
            }
            return [3 /*break*/, 7];
          }
          tblColRange = tblCol.getDataBodyRange();
          tblColRange.load(["format"]);
          tblColFormat = tblColRange.format;
          tblColFormat.load(["horizontalAlignment", "verticalAlignment"]);
          ctx.trackedObjects.add([tblColRange]);
          ;
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _b.sent();
          AddMessage("Column Name: ".concat(element.columnName));
          if (element.visible !== undefined && element.visible !== null) {
            AddMessage("Visible: ".concat(element.visible));
            if (!element.visible && hideLessRelevants) {
              tblColRange.columnHidden = true;
              return [3 /*break*/, 7]; //optimization: if column is set to become invisible skip remaining formatting
            }
            // await ctx.sync();
          }

          if (element.style !== undefined && element.style !== "") {
            //style must be the first prop to set as it overrides all the below props
            AddMessage("Style: ".concat(element.style));
            tblColRange.style = element.style;
            // await ctx.sync();
          }

          if (!(element.searchFor !== undefined && element.searchFor !== "")) return [3 /*break*/, 5];
          //Replacing values before applying remaining styles (dates come with Z )
          AddMessage("SearchFor/ReplaceWith: ".concat(element.searchFor, " / ").concat(element.ReplaceWith));
          criteria = {
            completeMatch: false,
            matchCase: true /* Ignore case when comparing strings. */
          };

          tblColRange.replaceAll(element.searchFor, element.replaceWith, criteria);
          return [4 /*yield*/, ctx.sync()];
        case 4:
          _b.sent();
          _b.label = 5;
        case 5:
          if (element.horizontalAlignment !== undefined && element.horizontalAlignment !== "") {
            AddMessage("Horizontal Alignment: ".concat(element.horizontalAlignment));
            tblColRange.format.horizontalAlignment = jsonConfigUtils.convertToHorizontalAlignment(element.horizontalAlignment);
          }
          if (element.verticalAlignment !== undefined && element.verticalAlignment !== "") {
            AddMessage("Vertical Alignment: ".concat(element.verticalAlignment));
            tblColRange.format.verticalAlignment = jsonConfigUtils.convertToVerticalAlignment(element.verticalAlignment);
          }
          if (element.columnWidth !== undefined && element.columnWidth !== null) {
            AddMessage("Column Width: ".concat(element.columnWidth));
            tblColRange.format.columnWidth = element.columnWidth;
            // await ctx.sync();
          }

          if (element.indentLevel !== undefined && element.indentLevel !== null) {
            AddMessage("Indent Level: ".concat(element.indentLevel));
            tblColRange.format.indentLevel = element.indentLevel;
            // await ctx.sync();
          }

          if (element.numberFormat !== undefined && element.numberFormat !== "") {
            AddMessage("Style: ".concat(element.numberFormat));
            tblColRange.numberFormat = element.numberFormat;
            // await ctx.sync();
          }

          if (element.autosizeColumn !== undefined && element.autosizeColumn !== null) {
            AddMessage("autosizeColumn: ".concat(element.autosizeColumn));
            if (element.autosizeColumn === "true") tblColRange.format.autofitColumns();
            // await ctx.sync();
          }

          return [4 /*yield*/, ctx.sync()];
        case 6:
          _b.sent();
          AddMessage("removing tracked objects for ".concat(element.columnName));
          ctx.trackedObjects.remove([tblCol, tblColRange]);
          ;
          _b.label = 7;
        case 7:
          _i++;
          return [3 /*break*/, 1];
        case 8:
          return [2 /*return*/, true];
        case 9:
          error_3 = _b.sent();
          addAnalysisInfo("columnName", 0, "Error traversing JSON array: ".concat(error_3), "ValidateJSONStruct", enumTypeAnalysis.Danger);
          console.error("Error traversing JSON array: ".concat(error_3));
          return [2 /*return*/, false];
        case 10:
          return [2 /*return*/];
      }
    });
  });
}

function applyJSONHighlights(jsonArray) {
  return __awaiter(this, void 0, void 0, function () {
    var retval, _i, _a, element, e, _b;
    return __generator(this, function (_c) {
      switch (_c.label) {
        case 0:
          retval = true;
          AddMessage(jsonArray);
          _i = 0, _a = jsonArray.conditionalFormats;
          _c.label = 1;
        case 1:
          if (!(_i < _a.length)) return [3 /*break*/, 11];
          element = _a[_i];
          e = element;
          _b = e.Type;
          switch (_b) {
            case ConditionalFormats_1.enumConditionalFormatType.ColorScale:
              return [3 /*break*/, 2];
            case ConditionalFormats_1.enumConditionalFormatType.ContainsText:
              return [3 /*break*/, 4];
            case ConditionalFormats_1.enumConditionalFormatType.CellValue:
              return [3 /*break*/, 6];
            case ConditionalFormats_1.enumConditionalFormatType.Custom:
              return [3 /*break*/, 8];
          }
          return [3 /*break*/, 9];
        case 2:
          return [4 /*yield*/, createConditionalFormatColorScale(e)];
        case 3:
          _c.sent();
          return [3 /*break*/, 10];
        case 4:
          return [4 /*yield*/, createConditionalFormatContainsText(e)];
        case 5:
          _c.sent();
          return [3 /*break*/, 10];
        case 6:
          return [4 /*yield*/, createConditionalFormatCellValue(e)];
        case 7:
          _c.sent();
          return [3 /*break*/, 10];
        case 8:
          return [3 /*break*/, 10];
        case 9:
          return [3 /*break*/, 10];
        case 10:
          _i++;
          return [3 /*break*/, 1];
        case 11:
          return [2 /*return*/, retval];
      }
    });
  });
}
function applyJSONFilters(jsonArray) {
  return __awaiter(this, void 0, void 0, function () {
    var retval, _i, _a, element, e;
    return __generator(this, function (_b) {
      retval = true;
      AddMessage(jsonArray);
      for (_i = 0, _a = jsonArray.filterDefinitions; _i < _a.length; _i++) {
        element = _a[_i];
        e = element;
        if (e.FilterActiveByUIOnly == false) {
          createColumnFilter(e);
        }
      }
      return [2 /*return*/, retval];
    });
  });
}
function createConditionalFormatColorScale(cf) {
  return __awaiter(this, void 0, void 0, function () {
    var retval, col, r, excelCF;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          AddMessage("Formatting column ".concat(cf.FriendlyName));
          retval = true;
          col = tbl.columns.getItemOrNullObject(cf.ColumnName);
          col.load(["isNullObject"]);
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          if (col.isNullObject) {
            return [2 /*return*/, false];
          }
          r = col.getDataBodyRange();
          excelCF = r.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
          excelCF.load(["colorScale"]);
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _a.sent();
          excelCF.colorScale.criteria.maximum.color = cf.ColorScaleColorMaximum;
          excelCF.colorScale.criteria.minimum.color = cf.ColorScaleColorMinimum;
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _a.sent();
          return [2 /*return*/, retval];
      }
    });
  });
}
function createConditionalFormatContainsText(cf) {
  return __awaiter(this, void 0, void 0, function () {
    var retval, col, r, excelCF;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          AddMessage("Creating Conditional Format ".concat(cf.FriendlyName));
          retval = true;
          col = tbl.columns.getItemOrNullObject(cf.ColumnName);
          col.load(["isNullObject"]);
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          if (col.isNullObject) {
            return [2 /*return*/, false];
          }
          r = col.getDataBodyRange();
          excelCF = r.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
          excelCF.load(["textComparison", "format", "format/fill", "format/font"]);
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _a.sent();
          excelCF.textComparison.rule = {
            operator: Excel.ConditionalTextOperator.contains,
            text: cf.ContainsTextSearch
          };
          if (cf.FillColor !== undefined && cf.FillColor !== null && cf.FillColor.toLowerCase() !== "null") {
            excelCF.textComparison.format.fill.color = cf.FillColor;
          }
          if (cf.FontColor !== undefined && cf.FontColor !== null && cf.FontColor.toLowerCase() !== "null") {
            excelCF.textComparison.format.font.color = cf.FontColor;
          }
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _a.sent();
          return [2 /*return*/, retval];
      }
    });
  });
}
function createConditionalFormatCellValue(cf) {
  return __awaiter(this, void 0, void 0, function () {
    var retval, col, r, excelCF;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          AddMessage("Creating Conditional Format ".concat(cf.FriendlyName));
          retval = true;
          col = tbl.columns.getItemOrNullObject(cf.ColumnName);
          col.load(["isNullObject"]);
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          if (col.isNullObject) {
            return [2 /*return*/, false];
          }
          r = col.getDataBodyRange();
          excelCF = r.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
          excelCF.load(["cellValue", "format", "format/fill", "format/font"]);
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _a.sent();
          excelCF.cellValue.rule = {
            formula1: cf.CellValueFormula1,
            formula2: cf.CellValueFormula2,
            operator: (0, ConditionalFormats_1.JsonEnumToCellValueOperator)(cf.CellValueOperator)
          };
          if (cf.FillColor !== undefined && cf.FillColor !== null && cf.FillColor.toLowerCase() !== "null") {
            excelCF.cellValue.format.fill.color = cf.FillColor;
          }
          if (cf.FontColor !== undefined && cf.FontColor !== null && cf.FontColor.toLowerCase() !== "null") {
            excelCF.cellValue.format.font.color = cf.FontColor;
          }
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _a.sent();
          return [2 /*return*/, retval];
      }
    });
  });
}
function createColumnFilter(f) {
  return __awaiter(this, void 0, void 0, function () {
    var retval, col, r, excelFilter;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          AddMessage("Creating Filter ".concat(f.FriendlyName));
          retval = true;
          col = tbl.columns.getItemOrNullObject(f.ColumnName);
          col.load(["isNullObject"]);
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _a.sent();
          if (col.isNullObject) {
            return [2 /*return*/, false];
          }
          r = col.getDataBodyRange();
          excelFilter = col.filter;
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _a.sent();
          excelFilter.apply({
            filterOn: Excel.FilterOn.values,
            values: [f.FilterValue]
          });
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _a.sent();
          AddMessage("Filter ".concat(f.FriendlyName, " created"));
          return [2 /*return*/, retval];
      }
    });
  });
}
function createColumnDefinitionsFromTable(jsonConfigUtils) {
  return __awaiter(this, void 0, void 0, function () {
    var columns, _i, _a, column, headerCell, isVisible, r, name_1, format, horizontalAlignment, verticalAlignment, columnWidth, indentLevel, style, numberFormat, autosizeColumn, visible, colDef;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          columns = tbl.columns.load(['name', 'values/format', 'values/horizontalAlignment', 'values/verticalAlignment', 'values/columnWidth', 'values/indentLevel', 'values/style', 'values/numberFormat', 'values/autosizeColumn']);
          // Synchronize with the document
          return [4 /*yield*/, ctx.sync()];
        case 1:
          // Synchronize with the document
          _b.sent();
          _i = 0, _a = columns.items;
          _b.label = 2;
        case 2:
          if (!(_i < _a.length)) return [3 /*break*/, 6];
          column = _a[_i];
          headerCell = column.getRange().getCell(0, 0);
          headerCell.load(["columnHidden"]);
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _b.sent();
          isVisible = headerCell.columnHidden;
          r = column.getDataBodyRange();
          r.load(['style', 'numberFormat', 'columnHidden']);
          r.format.load(['format', 'horizontalAlignment', 'verticalAlignment', 'columnWidth', 'indentLevel', 'style', 'numberFormat']);
          return [4 /*yield*/, ctx.sync()];
        case 4:
          _b.sent();
          name_1 = column.name;
          format = r.format;
          horizontalAlignment = r.format.horizontalAlignment;
          verticalAlignment = r.format.verticalAlignment;
          columnWidth = r.format.columnWidth;
          indentLevel = r.format.indentLevel;
          style = r.style;
          numberFormat = r.numberFormat[0].toString();
          autosizeColumn = false;
          visible = !headerCell.columnHidden;
          colDef = {
            columnName: name_1,
            isMandatory: true,
            horizontalAlignment: horizontalAlignment,
            verticalAlignment: verticalAlignment,
            columnWidth: columnWidth,
            indentLevel: indentLevel,
            style: style,
            numberFormat: numberFormat,
            visible: visible,
            autosizeColumn: false,
            searchFor: "",
            replaceWith: ""
          };
          jsonConfigUtils.addColumn(colDef);
          AddMessage("Adding column ".concat(name_1, "/format ").concat(numberFormat, "/ style ").concat(style));
          _b.label = 5;
        case 5:
          _i++;
          return [3 /*break*/, 2];
        case 6:
          return [2 /*return*/, jsonConfigUtils];
      }
    });
  });
}
function createFiltersFromTable(jsonConfig) {
  var _a;
  return __awaiter(this, void 0, void 0, function () {
    var conditionalFormats, results, _i, _b, column, filter, c, randomNumber, columnName, key, values, criterion2, correctedString;
    return __generator(this, function (_c) {
      switch (_c.label) {
        case 0:
          ctx.trackedObjects.add(tbl);
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _c.sent();
          tblRange.conditionalFormats.load();
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _c.sent();
          conditionalFormats = tblRange.conditionalFormats;
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _c.sent();
          results = [];
          _i = 0, _b = tbl.columns.items;
          _c.label = 4;
        case 4:
          if (!(_i < _b.length)) return [3 /*break*/, 8];
          column = _b[_i];
          column.load(["name", "filter"]);
          return [4 /*yield*/, ctx.sync()];
        case 5:
          _c.sent();
          AddMessage("Checking for filter on column ".concat(column.name));
          filter = column.filter;
          filter.load(["criteria"]);
          return [4 /*yield*/, ctx.sync()];
        case 6:
          _c.sent();
          c = filter.criteria;
          if (c !== null) {
            randomNumber = Math.ceil(Math.random() * 999999);
            columnName = column.name;
            key = columnName + randomNumber.toString();
            if (filter.criteria.filterOn == Excel.FilterOn.custom) {
              criterion2 = (_a = filter.criteria.criterion2) !== null && _a !== void 0 ? _a : '';
              correctedString = criterion2.toString() !== '' ? ", ".concat(criterion2.toString()) : '';
              values = filter.criteria.criterion1.toString() + correctedString;
            } else {
              values = filter.criteria.values.join(", ");
            }
            jsonConfig.addFilterCondition(columnName, values, key);
            AddMessage("Add filter ".concat(columnName, " / ").concat(values, " / ").concat(key));
          }
          _c.label = 7;
        case 7:
          _i++;
          return [3 /*break*/, 4];
        case 8:
          return [2 /*return*/];
      }
    });
  });
}

function createConditionalFormatsFromTable(jsonConfig) {
  return __awaiter(this, void 0, void 0, function () {
    var conditionalFormats, results, _i, _a, column, cfs, _b, _c, cf, randomNumber, _d;
    return __generator(this, function (_e) {
      switch (_e.label) {
        case 0:
          ctx.trackedObjects.add(tbl);
          return [4 /*yield*/, ctx.sync()];
        case 1:
          _e.sent();
          tblRange.conditionalFormats.load();
          return [4 /*yield*/, ctx.sync()];
        case 2:
          _e.sent();
          conditionalFormats = tblRange.conditionalFormats;
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _e.sent();
          results = [];
          _i = 0, _a = tbl.columns.items;
          _e.label = 4;
        case 4:
          if (!(_i < _a.length)) return [3 /*break*/, 18];
          column = _a[_i];
          cfs = column.getDataBodyRange().conditionalFormats;
          cfs.load("items");
          return [4 /*yield*/, ctx.sync()];
        case 5:
          _e.sent();
          _b = 0, _c = cfs.items;
          _e.label = 6;
        case 6:
          if (!(_b < _c.length)) return [3 /*break*/, 17];
          cf = _c[_b];
          randomNumber = Math.ceil(Math.random() * 999999);
          AddMessage("Column ".concat(column.name, " has CF of type ").concat(cf.type));
          _d = cf.type;
          switch (_d) {
            case Excel.ConditionalFormatType.cellValue:
              return [3 /*break*/, 7];
            case Excel.ConditionalFormatType.containsText:
              return [3 /*break*/, 9];
            case Excel.ConditionalFormatType.custom:
              return [3 /*break*/, 11];
            case Excel.ConditionalFormatType.colorScale:
              return [3 /*break*/, 12];
            case Excel.ConditionalFormatType.dataBar:
              return [3 /*break*/, 14];
          }
          return [3 /*break*/, 15];
        case 7:
          cf.cellValue.load(["rule", "format", "format/font", "format/fill"]);
          return [4 /*yield*/, ctx.sync()];
        case 8:
          _e.sent();
          jsonConfig.addConditionalFormatCellValue("".concat(column.name, " Cell Value ").concat(cf.cellValue.rule.formula1), column.name, false, false, "Normal", cf.cellValue.format.font.color, cf.cellValue.format.fill.color, cf.cellValue.rule.formula1, cf.cellValue.rule.formula2, (0, ConditionalFormats_1.CellValueOperatorToJsonEnum)(cf.cellValue.rule.operator));
          return [3 /*break*/, 16];
        case 9:
          cf.textComparison.load(["rule", "format", "format/font", "format/fill"]);
          return [4 /*yield*/, ctx.sync()];
        case 10:
          _e.sent();
          jsonConfig.addConditionalFormatContainsText("".concat(column.name, " Contains Text ").concat(cf.textComparison.rule.text), column.name, false, false, "Normal", cf.textComparison.format.fill.color, cf.textComparison.format.font.color, cf.textComparison.rule.text, cf.textComparison.rule.operator == Excel.ConditionalTextOperator.contains ? ConditionalFormats_1.enumConditionalFormatTextOperator.Contains : ConditionalFormats_1.enumConditionalFormatTextOperator.NotContains);
          return [3 /*break*/, 16];
        case 11:
          return [3 /*break*/, 16];
        case 12:
          cf.colorScale.load(["rule", "format", "format/font", "format/fill", "criteria"]);
          return [4 /*yield*/, ctx.sync()];
        case 13:
          _e.sent();
          jsonConfig.addConditionalFormatColorScale("".concat(column.name, " Color Scale ").concat(cf.colorScale.criteria.maximum.color, " Down to ").concat(cf.colorScale.criteria.minimum.color), column.name, false, false, "Normal", null, null, cf.colorScale.criteria.minimum.color, cf.colorScale.criteria.maximum.color);
          return [3 /*break*/, 16];
        case 14:
          return [3 /*break*/, 16];
        case 15:
          return [3 /*break*/, 16];
        case 16:
          _b++;
          return [3 /*break*/, 6];
        case 17:
          _i++;
          return [3 /*break*/, 4];
        case 18:
          return [2 /*return*/, results];
      }
    });
  });
}
function getJsonData() {
  return __awaiter(this, void 0, void 0, function () {
    var jsonType, response, _a, jsonData;
    return __generator(this, function (_b) {
      switch (_b.label) {
        case 0:
          return [4 /*yield*/, typeCDL()];
        case 1:
          jsonType = _b.sent();
          _a = jsonType;
          switch (_a) {
            case "rave-diag-log":
              return [3 /*break*/, 2];
            case "exo-cdl":
              return [3 /*break*/, 4];
            case "kusto-graph":
              return [3 /*break*/, 6];
            case "kusto-entityevent":
              return [3 /*break*/, 8];
          }
          return [3 /*break*/, 10];
        case 2:
          return [4 /*yield*/, fetch("./RaveCDLconfig.json")];
        case 3:
          response = _b.sent();
          return [3 /*break*/, 12];
        case 4:
          return [4 /*yield*/, fetch("./EXOCDLconfig.json")];
        case 5:
          response = _b.sent();
          return [3 /*break*/, 12];
        case 6:
          return [4 /*yield*/, fetch("./KustoGraphdb.json")];
        case 7:
          response = _b.sent();
          return [3 /*break*/, 12];
        case 8:
          return [4 /*yield*/, fetch("./KustoCalendarEntityEvent.json")];
        case 9:
          response = _b.sent();
          return [3 /*break*/, 12];
        case 10:
          return [4 /*yield*/, fetch("./RaveCDLconfig.json")];
        case 11:
          response = _b.sent();
          return [3 /*break*/, 12];
        case 12:
          AddMessage("Retrieving JSON data for key ".concat(jsonType, " / url=").concat(response.url, " /  status=").concat(response.status, " / statusText=").concat(response.statusText, " / type=").concat(response.type, " / ").concat(response.ok, " / ").concat(response.redirected, " / ").concat(response.body, " / ").concat(response.bodyUsed, " / ").concat(response.headers, " / ").concat(response.trailer, " "));
          return [4 /*yield*/, response.json()];
        case 13:
          jsonData = _b.sent();
          return [2 /*return*/, jsonData];
      }
    });
  });
}
//#endregion
//#region Analysys
function PerformAnalysis(context) {
  return __awaiter(this, void 0, void 0, function () {
    return __generator(this, function (_a) {
      return [2 /*return*/, null];
    });
  });
}
function CheckNumberOfRows() {
  return __awaiter(this, void 0, void 0, function () {
    var rowCount;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          return [4 /*yield*/, CheckNumberOfRows];
        case 1:
          if (!_a.sent()) return [2 /*return*/];
          return [4 /*yield*/, totalTblRows()];
        case 2:
          rowCount = _a.sent();
          if (rowCount >= 950) {
            addAnalysisInfo("Row count number", rowCount, "Number of rows is very close(or above) the Diag Limit of 1000Rows", "CheckNumberOfRows", enumTypeAnalysis.Warning);
            AddMessage("Number of rows is very close to the Diag Limit of 1000Rows returned($tblRange.rowCount)");
          }
          return [4 /*yield*/, ctx.sync()];
        case 3:
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}
//#endregion
//#region Config section
function createConfig() {
  return __awaiter(this, void 0, void 0, function () {
    var _this = this;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          myConsole_1.myConsole.reset();
          return [4 /*yield*/, Excel.run(function (context) {
            return __awaiter(_this, void 0, void 0, function () {
              var validTable;
              return __generator(this, function (_a) {
                switch (_a.label) {
                  case 0:
                    ctx = context;
                    sheet = context.workbook.worksheets.getActiveWorksheet();
                    tbl = sheet.tables.getItemOrNullObject("CDL");
                    tblRange = tbl.getRange();
                    jsonConfigUtils = new jsonConfigUtils_1.JsonConfigUtils();
                    return [4 /*yield*/, context.sync()];
                  case 1:
                    _a.sent();
                    return [4 /*yield*/, CreateTable(context, true)];
                  case 2:
                    validTable = _a.sent();
                    if (!validTable) {
                      return [2 /*return*/];
                    }

                    return [4 /*yield*/, createColumnDefinitionsFromTable(jsonConfigUtils)];
                  case 3:
                    _a.sent();
                    return [4 /*yield*/, createConditionalFormatsFromTable(jsonConfigUtils)];
                  case 4:
                    _a.sent();
                    return [4 /*yield*/, createFiltersFromTable(jsonConfigUtils)];
                  case 5:
                    _a.sent();
                    document.getElementById("jsonConfig").textContent = jsonConfigUtils.getValue();
                    AddMessage("JSON config creation successfull!");
                    return [2 /*return*/];
                }
              });
            });
          })];

        case 1:
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}

exports.createConfig = createConfig;
function testJsonFile() {
  return __awaiter(this, void 0, void 0, function () {
    var _this = this;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          return [4 /*yield*/, Excel.run(function (context) {
            return __awaiter(_this, void 0, void 0, function () {
              var validTable, tempJson, hideCols, isTableValid;
              return __generator(this, function (_a) {
                switch (_a.label) {
                  case 0:
                    ctx = context;
                    sheet = context.workbook.worksheets.getActiveWorksheet();
                    jsonConfigUtils = new jsonConfigUtils_1.JsonConfigUtils();
                    return [4 /*yield*/, ctx.sync()];
                  case 1:
                    _a.sent();
                    return [4 /*yield*/, CreateTable(context, true)];
                  case 2:
                    validTable = _a.sent();
                    if (!validTable) {
                      return [2 /*return*/];
                    }

                    myConsole_1.myConsole.reset();
                    document.getElementById("jsonConfig").textContent = "";
                    return [4 /*yield*/, getJsonData()];
                  case 3:
                    tempJson = _a.sent();
                    document.getElementById("jsonConfig").textContent = JSON.stringify(tempJson, null, 2);
                    return [4 /*yield*/, hideLessRelevants()];
                  case 4:
                    hideCols = _a.sent();
                    return [4 /*yield*/, applyJsonConfig(tempJson, hideCols)];
                  case 5:
                    isTableValid = _a.sent();
                    AddMessage("Test JSON file succeeded!");
                    return [2 /*return*/];
                }
              });
            });
          })];

        case 1:
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}

exports.testJsonFile = testJsonFile;
function testConfig() {
  return __awaiter(this, void 0, void 0, function () {
    var _this = this;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          return [4 /*yield*/, Excel.run(function (context) {
            return __awaiter(_this, void 0, void 0, function () {
              var validTable, textbox, tempJson, hideCols, isTableValid;
              return __generator(this, function (_a) {
                switch (_a.label) {
                  case 0:
                    ctx = context;
                    sheet = context.workbook.worksheets.getActiveWorksheet();
                    jsonConfigUtils = new jsonConfigUtils_1.JsonConfigUtils();
                    return [4 /*yield*/, ctx.sync()];
                  case 1:
                    _a.sent();
                    return [4 /*yield*/, CreateTable(context, true)];
                  case 2:
                    validTable = _a.sent();
                    if (!validTable) {
                      return [2 /*return*/];
                    }

                    myConsole_1.myConsole.reset();
                    textbox = document.getElementById("jsonConfig");
                    tempJson = textbox.value;
                    return [4 /*yield*/, hideLessRelevants()];
                  case 3:
                    hideCols = _a.sent();
                    return [4 /*yield*/, applyJsonConfig(tempJson, hideCols)];
                  case 4:
                    isTableValid = _a.sent();
                    AddMessage("Test JSON configuration succeeded!");
                    return [2 /*return*/];
                }
              });
            });
          })];

        case 1:
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}

exports.testConfig = testConfig;
//#endregion
//Main Function
function run() {
  return __awaiter(this, void 0, void 0, function () {
    var error_4;
    var _this = this;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          _a.trys.push([0, 2,, 3]);
          return [4 /*yield*/, Excel.run(function (context) {
            return __awaiter(_this, void 0, void 0, function () {
              var validTable, isTableValid, urlCDLVideo;
              return __generator(this, function (_a) {
                switch (_a.label) {
                  case 0:
                    /**
                     * Insert your Excel code here
                     */
                    myConsole_1.myConsole.reset();
                    return [4 /*yield*/, showSpinner(true)];
                  case 1:
                    _a.sent();
                    return [4 /*yield*/, resetAnalysisInfo()];
                  case 2:
                    _a.sent();
                    return [4 /*yield*/, AddMessage("Starting Processing")];
                  case 3:
                    _a.sent();
                    ctx = context;
                    sheet = context.workbook.worksheets.getActiveWorksheet();
                    jsonConfigUtils = new jsonConfigUtils_1.JsonConfigUtils();
                    // tbl = sheet.tables.getItemOrNullObject("CDL");
                    // tblRange = tbl.getRange();
                    // tbl.clearFilters();
                    sheet.getRange().clear("Formats");
                    sheet.getRange().conditionalFormats.clearAll();
                    sheet.freezePanes.unfreeze();
                    return [4 /*yield*/, context.sync()];
                  case 4:
                    _a.sent();
                    return [4 /*yield*/, CreateTable(context)];
                  case 5:
                    validTable = _a.sent();
                    if (!validTable) {
                      return [2 /*return*/];
                    }

                    AddMessage("Create Table Done");
                    return [4 /*yield*/, getJsonData()];
                  case 6:
                    jsonLog = _a.sent();
                    return [4 /*yield*/, applyJsonConfig(jsonLog)];
                  case 7:
                    isTableValid = _a.sent();
                    if (!isTableValid) {
                      addAnalysisInfo("CDL Invalid", 0, "CDL Structure is invalid (check previous exceptions)", "CDLInvalid", enumTypeAnalysis.Danger);
                      showSpinner(false);
                      return [2 /*return*/];
                    }
                    //format section
                    return [4 /*yield*/, freezeColumns("Ignorable")];
                  case 8:
                    //format section
                    _a.sent();
                    //await PerformAnalysis(context);
                    AddMessage("Processing done.");
                    showSpinner(false);
                    urlCDLVideo = "https://msit.microsoftstream.com/video/4221a4ff-0400-9fb2-4805-f1eb0f28f09b";
                    addAnalysisInfo("Success", 0, "Process executed successfully, check the video on CDL analysis ".concat(urlCDLVideo, " "), "success", enumTypeAnalysis.Success);
                    return [2 /*return*/];
                }
              });
            });
          })];

        case 1:
          _a.sent();
          return [3 /*break*/, 3];
        case 2:
          error_4 = _a.sent();
          showSpinner(false);
          console.error(error_4);
          AddMessage(error_4);
          addAnalysisInfo("Error", 0, error_4, "Run/Catch", enumTypeAnalysis.Danger);
          return [3 /*break*/, 3];
        case 3:
          return [2 /*return*/];
      }
    });
  });
}

exports.run = run;
// function CellValueOperatorToJsonEnum(operator: string): enumCellValueOperator {
//   //throw new Error('Function not implemented.');
//   return enumCellValueOperator.EQ;
// }

/***/ }),

/***/ "./src/taskpane/taskpane.html":
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../node_modules/html-loader/dist/runtime/getUrl.js */ "./node_modules/html-loader/dist/runtime/getUrl.js");
/* harmony import */ var _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0__);
// Imports

var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
// Module
var ___HTML_LOADER_REPLACEMENT_0___ = _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default()(___HTML_LOADER_IMPORT_0___);
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\r\n<!-- This file shows how to design a first-run page that provides a welcome screen to the user about the features of the add-in. -->\r\n\r\n<!DOCTYPE html>\r\n<html>\r\n\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\r\n    <title>Contoso Task Pane Add-in</title>\r\n\r\n    <!-- Office JavaScript API -->\r\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js\"><" + "/script>\r\n\r\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\r\n    <!-- <link rel=\"stylesheet\" href=\"https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css\"/> -->\r\n\r\n    <link href=\"https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css\" rel=\"stylesheet\">\r\n    <" + "script src=\"https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/js/bootstrap.bundle.min.js\"><" + "/script>\r\n    <!-- Bootstrap CSS\r\n    <link rel=\"stylesheet\" href=\"https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css\"> -->\r\n\r\n    <!-- jQuery and Bootstrap JS -->\r\n    <" + "script src=\"https://code.jquery.com/jquery-3.5.1.slim.min.js\"\r\n        integrity=\"sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj\"\r\n        crossorigin=\"anonymous\"><" + "/script>\r\n\r\n\r\n    <link rel=\"stylesheet\" href=\"https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.4/font/bootstrap-icons.css\">\r\n                                   \r\n\r\n\r\n    <!-- Template styles -->\r\n    <link href=\"" + ___HTML_LOADER_REPLACEMENT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n</head>\r\n\r\n<body class=\"bg-dark\">\r\n\r\n    <!-- Offcanvas Sidebar -->\r\n    <div class=\"offcanvas offcanvas-start\" id=\"timeline\">\r\n        <div class=\"offcanvas-header\">\r\n        <h2 class=\"offcanvas-title\">Console</h2>\r\n        <button type=\"button\" class=\"btn-close text-reset\" data-bs-dismiss=\"offcanvas\">.</button>\r\n        </div>\r\n        <div class=\"offcanvas-body\">\r\n            <div class=\"form-group container\">\r\n                <div class=\"form-group\">\r\n                    <div class=\"card\">\r\n                        <div class=\"card-header\">\r\n                          <h5 id=\"timeline-subject\" class=\"card-title\">[Subject]</h5>\r\n                          <p id=\"timeline-date\" class=\"card-text\">[Date]</p>\r\n                        </div>\r\n                        <div class=\"card-body\">\r\n                          <div id=\"timeline-body\">\r\n                            <!-- Timeline rows will be dynamically populated here -->\r\n                          </div>\r\n                          <!-- Other card body elements can be added here -->\r\n                        </div>\r\n                    </div>\r\n                </div>\r\n            </div>\r\n        </div>\r\n    </div>\r\n\r\n\r\n    <!-- Offcanvas Sidebar -->\r\n    <div class=\"offcanvas offcanvas-start\" id=\"console\">\r\n        <div class=\"offcanvas-header\">\r\n        <h2 class=\"offcanvas-title\">Console</h2>\r\n        <button type=\"button\" class=\"btn-close text-reset\" data-bs-dismiss=\"offcanvas\">.</button>\r\n        </div>\r\n        <div class=\"offcanvas-body\">\r\n            <div class=\"form-group container\">\r\n                <div class=\"form-group\">\r\n                    <div id=\"myConsole\" class=\"container\">\r\n                        <div class=\"row header\">\r\n                            <div class=\"col-2 text-xs\">#</div>\r\n                            <div class=\"col-10 text-xs\">First</div>\r\n                        </div>\r\n                    </div>\r\n                      \r\n                </div>\r\n            </div>\r\n        </div>\r\n    </div>\r\n\r\n\r\n    <div class=\"offcanvas offcanvas-start\" id=\"demo\">\r\n        <div class=\"offcanvas-header\">\r\n        <h2 class=\"offcanvas-title\">JSON Config</h2>\r\n        <button type=\"button\" class=\"btn-close text-reset\" data-bs-dismiss=\"offcanvas\">.</button>\r\n        </div>\r\n        <div class=\"offcanvas-body\">\r\n            <div class=\"form-group container\">\r\n                <div class=\"form-group\">\r\n                    <button type=\"button\" class=\"btn btn-primary\" id=\"createConfig\" data-toggle=\"tooltip\" data-placement=\"top\" title=\"Create JSON from table\"><i class=\"bi bi-arrow-bar-down\"></i></button>\r\n                    <button type=\"button\" class=\"btn btn-primary\" id=\"testConfig\" data-toggle=\"tooltip\" data-placement=\"top\" title=\"Test json Format on current table\"><i class=\"bi bi-arrow-bar-up\"></i></button>\r\n                    <button type=\"button\" class=\"btn btn-primary\" id=\"testJsonFile\" data-toggle=\"tooltip\" data-placement=\"top\" title=\"test json file against current table\"><i class=\"bi bi-file-arrow-down\"></i></button>\r\n                </div>\r\n                <div class=\"form-group\">\r\n                    <textarea class=\"form-control\" rows=\"15\" placeholder=\"Json Config\" id=\"jsonConfig\"></textarea>\r\n                </div>\r\n            </div>\r\n        </div>\r\n    </div>\r\n  \r\n\r\n\r\n    <div class=\"offcanvas offcanvas-start\" id=\"video\">\r\n        <div class=\"offcanvas-header\">\r\n        <h2 class=\"offcanvas-title\">Video Troubleshoot Help(POC)</h2>\r\n        <button type=\"button\" class=\"btn-close text-reset\" data-bs-dismiss=\"offcanvas\"></button>\r\n        </div>\r\n        <div class=\"offcanvas-body\">\r\n            <div class=\"container\">\r\n                <div class=\"card-header\">\r\n                    Video Sharing Calendars\r\n                </div>\r\n                <div class=\"card-body \">\r\n                    <div class=\"embed-responsive embed-responsive-21by9\">\r\n                        <iframe class=\"embed-responsive-item\" width=\"560\" height=\"315\" src=\"https://www.youtube.com/embed/7C38muJjnyc\"\r\n                            title=\"YouTube video player\" frameborder=\"0\"\r\n                            allow=\"accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share\"\r\n                            allowfullscreen></iframe>\r\n                    </div>\r\n                </div>\r\n            </div>\r\n        </div>\r\n    </div>\r\n  \r\n \r\n    <div class=\"bg-light rounded p-3 card\">\r\n        <div class=\"card-header\">\r\n            CDL Check provides compreehensive formatting of CDL logs as well as basic troubleshooting help\r\n            This is the AddingTimeline Branch deployed\r\n        </div>\r\n        <div class=\"card-body\">\r\n            <form>\r\n                <div class=\"form-group\" id=\"sideload-msg\">\r\n                    <H2>officeJS is not loaded, Addin will not run!</H2>\r\n                </div>\r\n\r\n                <div class=\"form-group\">\r\n                    <!-- Spacer     -->\r\n                </div>\r\n\r\n                <div class=\"form-group\">\r\n                    <label for=\"typeCDL\">Choose Type of CDL Log</label>\r\n                    <select class=\"form-control\" id=\"typeCDL\">\r\n                        <option value=\"rave-diag-log\">RAVE Diag Log (Default)</option>\r\n                        <option value=\"exo-cdl\">RAW CDL from Get-CalendarDiagnosticObjects</option>\r\n                        <option value=\"kusto-graph\">Kusto - Graph db</option>\r\n                        <option value=\"kusto-entityevent\">Kusto - Calendar Entity Event</option>\r\n                    </select>\r\n                </div>\r\n                \r\n                <div class=\"form-group\">\r\n\r\n                </div>\r\n\r\n                <div class=\"form-group\">\r\n                    <div class=\"form-check form-switch\">\r\n                        <input type=\"checkbox\" class=\"form-check-input\" role=\"switch\" id=\"isOrganizer\">\r\n                        <label class=\"form-check-label\" for=\"isOrganizer\">This is an Organizer CDL</label>\r\n                    </div>\r\n                </div>\r\n                <div class=\"form-group form-switch d-none\">\r\n                    <input type=\"checkbox\" class=\"form-check-input\" role=\"switch\" id=\"warn1KRows\">\r\n                    <label class=\"form-check-label\" for=\"warn1KRows\">Warn if Rows are close to 1K</label>\r\n                </div>\r\n                <div class=\"form-group form-switch\">\r\n                    <input type=\"checkbox\" class=\"form-check-input\" role=\"switch\" id=\"hideLessRelevants\">\r\n                    <label class=\"form-check-label\" for=\"hideLessRelevants\">Hide less relevant columns</label>\r\n                </div>\r\n\r\n                <div class=\"form-group\">\r\n                    <button type=\"button\" class=\"btn btn-primary\" id=\"run\"><i class=\"bi bi-calendar-week\"></i></button>\r\n                    <button class=\"btn btn-primary\" type=\"button\" data-bs-toggle=\"offcanvas\" data-bs-target=\"#demo\"><i class=\"bi bi-filetype-json\"></i></button>\r\n                    <button class=\"btn btn-primary\" type=\"button\" data-bs-toggle=\"offcanvas\" data-bs-target=\"#video\"><i class=\"bi bi-film\"></i></button>\r\n                    <button class=\"btn btn-primary\" type=\"button\" data-bs-toggle=\"offcanvas\" data-bs-target=\"#console\"><i class=\"bi bi-list-check\"></i></button>\r\n                    <!-- <button class=\"btn btn-primary\" type=\"button\" data-bs-toggle=\"offcanvas\" data-bs-target=\"#timeline\"><i class=\"bi bi-watch\"></i></button> -->\r\n                    \r\n                    <div class=\"spinner-border text-primary invisible\" role=\"status\" id=\"spinner\">\r\n                        <!-- <span class=\"sr-only\">Loading...</span> -->\r\n                    </div>\r\n                    \r\n                </div>\r\n                <div class=\"form-group\">\r\n                    <p class=\"mb-1 small\" id=\"statusMessage\">Status Message</p>\r\n                </div>\r\n                <div class=\"list-group\" id=\"analysisInfo\">\r\n                    <!-- Template of data to be shown on analysisInfo -->\r\n                    <!-- <a href=\"#\" class=\"list-group-item list-group-item-action list-group-item-warning\">\r\n                        <div class=\"d-flex w-100 justify-content-between\">\r\n                            <h5 class=\"mb-1\">Row limit</h5>\r\n                            <span class=\"badge badge-primary badge-pill\">1002</span>\r\n                        </div>\r\n                        <p class=\"mb-1\">If rows returned are close to 1K, this can be a limitation of Rave diagnostics and could mean dataset returned is not complete. To avoid this, ask for raw CDL to the customer, and format directly with the option RAW</p>\r\n                        <small>Get-CalendarDiagnosticObjects</small>\r\n                    </a> -->\r\n                </div>\r\n            </form>\r\n        </div>\r\n        <div class=\"card-footer\">\r\n            <small>Exchange Calendaring team</small>\r\n        </div>\r\n        \r\n    </div>\r\n</body>\r\n\r\n</html>\r\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);

/***/ }),

/***/ "./node_modules/html-loader/dist/runtime/getUrl.js":
/*!*********************************************************!*\
  !*** ./node_modules/html-loader/dist/runtime/getUrl.js ***!
  \*********************************************************/
/***/ (function(module) {



module.exports = function (url, options) {
  if (!options) {
    // eslint-disable-next-line no-param-reassign
    options = {};
  }

  if (!url) {
    return url;
  } // eslint-disable-next-line no-underscore-dangle, no-param-reassign


  url = String(url.__esModule ? url.default : url);

  if (options.hash) {
    // eslint-disable-next-line no-param-reassign
    url += options.hash;
  }

  if (options.maybeNeedQuotes && /[\t\n\f\r "'=<>`]/.test(url)) {
    return "\"".concat(url, "\"");
  }

  return url;
};

/***/ }),

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "4f424550f2dc5a27a461.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	!function() {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = function(module) {
/******/ 			var getter = module && module.__esModule ?
/******/ 				function() { return module['default']; } :
/******/ 				function() { return module; };
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	!function() {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = function(exports, definition) {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
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
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript)
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) scriptUrl = scripts[scripts.length - 1].src
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = document.baseURI || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module is referenced by other modules so it can't be inlined
/******/ 	__webpack_require__("./src/taskpane/taskpane.ts");
/******/ 	var __webpack_exports__ = __webpack_require__("./src/taskpane/taskpane.html");
/******/ 	
/******/ })()
;
//# sourceMappingURL=taskpane.js.map