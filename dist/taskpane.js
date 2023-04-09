/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/taskpane/taskpane.ts":
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.ts ***!
  \**********************************/
/***/ (function(__unused_webpack_module, exports) {



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
exports.run = void 0;
/* global console, document, Excel, Office */
var _ctx;
var tbl;
var sheet;
var tblRange;
//#region Properties
var _totalTblRows;
function totalTblRows() {
  return __awaiter(this, void 0, void 0, function () {
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          if (!(_totalTblRows <= 0)) return [3 /*break*/, 2];
          tblRange.load(["rowCount"]);
          return [4 /*yield*/, _ctx.sync()];
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
function typeCDL() {
  var selectElement = document.getElementById('typeCDL');
  return selectElement.value;
}
var organizerTableStyle = "TableStyleLight10";
var attendeeTableStyle = "TableStyleLight11";
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
}
var enumTypeAnalysis;
(function (enumTypeAnalysis) {
  enumTypeAnalysis[enumTypeAnalysis["Warning"] = 0] = "Warning";
  enumTypeAnalysis[enumTypeAnalysis["Action"] = 1] = "Action";
  enumTypeAnalysis[enumTypeAnalysis["Danger"] = 2] = "Danger";
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
      }
      divElement = document.createElement("div");
      divElement.classList.add("d-flex", "w-100", "justify-content-between");
      h5Element = document.createElement("h5");
      h5Element.classList.add("mb-1");
      h5Element.innerText = title;
      spanElement = document.createElement("span");
      spanElement.classList.add("badge", "badge-primary", "badge-pill");
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
//#endregion
//#region Init OfficeJS
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    resetAnalysisInfo();
  }
});
function ClearAllTables() {
  return __awaiter(this, void 0, void 0, function () {
    var error_1;
    var _this = this;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          _a.trys.push([0, 2,, 3]);
          return [4 /*yield*/, Excel.run(function (context) {
            return __awaiter(_this, void 0, void 0, function () {
              var sheet, tables;
              return __generator(this, function (_a) {
                switch (_a.label) {
                  case 0:
                    sheet = context.workbook.worksheets.getActiveWorksheet();
                    tables = sheet.tables;
                    // Load the items property of the tables object
                    tables.load("items");
                    // Synchronize the document state by executing the queued commands
                    return [4 /*yield*/, context.sync()];
                  case 1:
                    // Synchronize the document state by executing the queued commands
                    _a.sent();
                    // Loop through each table and remove its formatting
                    tables.items.forEach(function (table) {
                      table.getRange().clear("Formats");
                    });
                    // Synchronize the document state by executing the queued commands
                    return [4 /*yield*/, context.sync()];
                  case 2:
                    // Synchronize the document state by executing the queued commands
                    _a.sent();
                    return [2 /*return*/];
                }
              });
            });
          })];

        case 1:
          _a.sent();
          return [3 /*break*/, 3];
        case 2:
          error_1 = _a.sent();
          console.error(error_1);
          return [3 /*break*/, 3];
        case 3:
          return [2 /*return*/];
      }
    });
  });
}
//#endregion
function UnTableIfExistsAlready(context) {
  return __awaiter(this, void 0, void 0, function () {
    return __generator(this, function (_a) {
      return [2 /*return*/];
    });
  });
}

function CreateTable(context) {
  return __awaiter(this, void 0, void 0, function () {
    var _isOrganizer, range, tbl_1, tblRange_1, error_2;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          _a.trys.push([0, 8,, 9]);
          return [4 /*yield*/, ClearAllTables()];
        case 1:
          _a.sent();
          return [4 /*yield*/, isOrganizer()];
        case 2:
          _isOrganizer = _a.sent();
          range = sheet.getUsedRange();
          range.load("address");
          tbl_1 = sheet.tables.getItemOrNullObject("CDL");
          tblRange_1 = tbl_1.getRange();
          return [4 /*yield*/, context.sync()];
        case 3:
          _a.sent();
          if (!tbl_1.isNullObject) return [3 /*break*/, 5];
          tbl_1 = sheet.tables.add(range, true /* hasHeaders */);
          tbl_1.name = "CDL";
          tblRange_1 = tbl_1.getRange();
          return [4 /*yield*/, context.sync()];
        case 4:
          _a.sent();
          _a.label = 5;
        case 5:
          if (_isOrganizer) {
            tbl_1.style = organizerTableStyle;
          } else {
            tbl_1.style = attendeeTableStyle;
          }
          tbl_1.load('tableStyle');
          tbl_1.columns.load();
          tblRange_1 = tbl_1.getRange();
          return [4 /*yield*/, context.sync()];
        case 6:
          _a.sent();
          // // Update table style 
          // //Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" 
          // //through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11"
          // tbl.style = "TableStyleLight10";
          // tbl.load('tableStyle');
          return [4 /*yield*/, context.sync()];
        case 7:
          // // Update table style 
          // //Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" 
          // //through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11"
          // tbl.style = "TableStyleLight10";
          // tbl.load('tableStyle');
          _a.sent();
          return [3 /*break*/, 9];
        case 8:
          error_2 = _a.sent();
          console.error(error_2);
          AddMessage(error_2);
          return [3 /*break*/, 9];
        case 9:
          return [2 /*return*/];
      }
    });
  });
}

function FormatCells(context) {
  return __awaiter(this, void 0, void 0, function () {
    var sheet, columnsA, columnsEF, columnC, columnE, columnG, columnsHI, columnsKM, columnP, columnQ, columnN, columnY, columnW;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          sheet = context.workbook.worksheets.getActiveWorksheet();
          columnsA = sheet.getRange("A:A");
          columnsA.format.horizontalAlignment = "Center";
          columnsA.format.verticalAlignment = "Bottom";
          columnsA.columnWidth = 6;
          columnsA.format.autofitColumns();
          columnsEF = sheet.getRange("E:F");
          columnsEF.format.autofitColumns();
          columnsEF.format.indentLevel = 1;
          columnC = sheet.getRange("C:C");
          columnC.columnWidth = 8;
          columnE = sheet.getRange("E:E");
          columnE.columnWidth = 18;
          columnG = sheet.getRange("G:G");
          columnG.format.autofitColumns();
          columnG.columnWidth = 25;
          columnsHI = sheet.getRange("H:I");
          columnsHI.format.horizontalAlignment = "Center";
          columnsHI.format.verticalAlignment = "Bottom";
          columnsHI.columnWidth = 6;
          columnsKM = sheet.getRange("K:M");
          columnsKM.format.indentLevel = 1;
          columnP = sheet.getRange("P:P");
          columnP.format.columnWidth = 6.33;
          columnQ = sheet.getRange("Q:Q");
          columnQ.columnWidth = 11.67;
          columnN = sheet.getRange("N:N");
          columnN.format.horizontalAlignment = "Center";
          columnN.format.verticalAlignment = "Bottom";
          columnN.columnWidth = 10;
          columnY = sheet.getRange("Y:Y");
          columnY.style = "Neutral";
          columnY.columnWidth = 10.56;
          columnW = sheet.getRange("W:W");
          columnW.format.indentLevel = 1;
          columnW.columnWidth = 11.22;
          return [4 /*yield*/, context.sync()];
        case 1:
          _a.sent();
          // Done
          console.log("Cells formatted.");
          return [2 /*return*/];
      }
    });
  });
}

function FilterIgnorable(context) {
  return __awaiter(this, void 0, void 0, function () {
    var ignorableFilter;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          // let sheet = context.workbook.worksheets.getActiveWorksheet();
          // let CDLTable = sheet.tables.getItem("CDL");
          // Queue a command to apply a filter on the Category column.
          tbl.columns.load();
          return [4 /*yield*/, context.sync()];
        case 1:
          _a.sent();
          ignorableFilter = tbl.columns.getItemOrNullObject("Ignorable").filter;
          return [4 /*yield*/, context.sync()];
        case 2:
          _a.sent();
          ignorableFilter.apply({
            filterOn: Excel.FilterOn.values,
            values: ["FALSE"]
          });
          return [4 /*yield*/, context.sync()];
        case 3:
          _a.sent();
          console.log("Cells filtered.");
          return [2 /*return*/];
      }
    });
  });
}

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

function HighlighIgnorable(context) {
  return __awaiter(this, void 0, void 0, function () {
    var ignorableFilter, isFilterNull, range;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          ignorableFilter = tbl.columns.getItemOrNullObject("Ignorable").filter;
          ignorableFilter.clear();
          return [4 /*yield*/, context.sync()];
        case 1:
          _a.sent();
          isFilterNull = ignorableFilter.isNullObject;
          ignorableFilter.apply({
            filterOn: Excel.FilterOn.values,
            values: ["TRUE"]
          });
          return [4 /*yield*/, context.sync()];
        case 2:
          _a.sent();
          range = tbl.getDataBodyRange();
          tbl.load("address");
          return [4 /*yield*/, context.sync()];
        case 3:
          _a.sent();
          range.format.font.color = "blue";
          // range.format.fill.tintAndShade = 0.399975585192419;
          // column.format.font.style = "20% - Accent5";
          // Clear the filter
          ignorableFilter.clear();
          return [4 /*yield*/, context.sync()];
        case 4:
          _a.sent();
          return [2 /*return*/];
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
        case 1:
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}
//format columns SentRepresentingEmailAddress	ResponsibleUserName	SenderEmailAddress
function FormatRawFrom(context) {
  return __awaiter(this, void 0, void 0, function () {
    var colRepresentingEmailAddress, colResponsibleUserName, colSenderEmailAddress;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          colRepresentingEmailAddress = tbl.columns.getItemOrNullObject("SentRepresentingEmailAddress").getRange();
          colResponsibleUserName = tbl.columns.getItemOrNullObject("ResponsibleUserName").getRange();
          colSenderEmailAddress = tbl.columns.getItemOrNullObject("SenderEmailAddress").getRange();
          colRepresentingEmailAddress.format.horizontalAlignment = "Center";
          colRepresentingEmailAddress.format.verticalAlignment = "Bottom";
          colRepresentingEmailAddress.format.indentLevel = 1;
          colRepresentingEmailAddress.format.autofitColumns();
          return [4 /*yield*/, context.sync()];
        case 1:
          _a.sent();
          colResponsibleUserName.format.horizontalAlignment = "Center";
          colResponsibleUserName.format.verticalAlignment = "Bottom";
          colResponsibleUserName.format.indentLevel = 1;
          colResponsibleUserName.format.autofitColumns();
          return [4 /*yield*/, context.sync()];
        case 2:
          _a.sent();
          colSenderEmailAddress.format.horizontalAlignment = "Center";
          colSenderEmailAddress.format.verticalAlignment = "Bottom";
          colSenderEmailAddress.format.autofitColumns();
          colSenderEmailAddress.format.indentLevel = 1;
          return [4 /*yield*/, context.sync()];
        case 3:
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
          return [2 /*return*/];
      }
    });
  });
}

function HighLightCreates(context) {
  return __awaiter(this, void 0, void 0, function () {
    var colRange, conditionalFormat;
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
          return [2 /*return*/];
      }
    });
  });
}
//Main Function
function run() {
  return __awaiter(this, void 0, void 0, function () {
    var error_3;
    var _this = this;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          _a.trys.push([0, 2,, 3]);
          return [4 /*yield*/, Excel.run(function (context) {
            return __awaiter(_this, void 0, void 0, function () {
              return __generator(this, function (_a) {
                switch (_a.label) {
                  case 0:
                    /**
                     * Insert your Excel code here
                     */
                    return [4 /*yield*/, showSpinner(true)];
                  case 1:
                    /**
                     * Insert your Excel code here
                     */
                    _a.sent();
                    return [4 /*yield*/, AddMessage("Starting Processing")];
                  case 2:
                    _a.sent();
                    _ctx = context;
                    sheet = context.workbook.worksheets.getActiveWorksheet();
                    tbl = sheet.tables.getItemOrNullObject("CDL");
                    return [4 /*yield*/, context.sync()];
                  case 3:
                    _a.sent();
                    return [4 /*yield*/, CreateTable(context).then(function () {
                      AddMessage("Create Table Done");
                    })];
                  case 4:
                    _a.sent();
                    return [4 /*yield*/, FormatCells(context).then(function () {
                      AddMessage("Format cells Done");
                    })];
                  case 5:
                    _a.sent();
                    return [4 /*yield*/, HighlighIgnorable(context).then(function () {
                      AddMessage("Highlight Ignorable Done");
                    })];
                  case 6:
                    _a.sent();
                    return [4 /*yield*/, HighlightApptSequence(context).then(function () {
                      AddMessage("Highlight  Done");
                    })];
                  case 7:
                    _a.sent();
                    return [4 /*yield*/, HighlightCRA(context).then(function () {
                      AddMessage("Highlight CRA Done");
                    })];
                  case 8:
                    _a.sent();
                    return [4 /*yield*/, HighLightCreates(context).then(function () {
                      AddMessage("Highlight Create Done");
                    })];
                  case 9:
                    _a.sent();
                    return [4 /*yield*/, FormatRawFrom(context).then(function () {
                      AddMessage("Format raw from Done");
                    })];
                  case 10:
                    _a.sent();
                    return [4 /*yield*/, FormatDateColumn(context, "ModifiedDate").then(function () {
                      AddMessage("Format ModifiedDate Done");
                    })];
                  case 11:
                    _a.sent(); //ModifiedDate
                    return [4 /*yield*/, FormatDateColumn(context, "StartTime").then(function () {
                      AddMessage("Create StartTime Done");
                    })];
                  case 12:
                    _a.sent(); //StartTime
                    return [4 /*yield*/, FormatDateColumn(context, "EndTime").then(function () {
                      AddMessage("Create End Done");
                    })];
                  case 13:
                    _a.sent(); //EndTime
                    return [4 /*yield*/, FilterIgnorable(context).then(function () {
                      AddMessage("Filter Ignorable Done");
                    })];
                  case 14:
                    _a.sent();
                    return [4 /*yield*/, context.sync()];
                  case 15:
                    _a.sent();
                    return [4 /*yield*/, PerformAnalysis(context).then(function () {
                      AddMessage("Perform Analysis Done");
                    })];
                  case 16:
                    _a.sent();
                    console.log("Processing done.");
                    AddMessage("Done!");
                    showSpinner(false);
                    return [2 /*return*/];
                }
              });
            });
          })];

        case 1:
          _a.sent();
          return [3 /*break*/, 3];
        case 2:
          error_3 = _a.sent();
          showSpinner(false);
          console.error(error_3);
          AddMessage(error_3);
          return [3 /*break*/, 3];
        case 3:
          return [2 /*return*/];
      }
    });
  });
}

exports.run = run;
function PerformAnalysis(context) {
  return __awaiter(this, void 0, void 0, function () {
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          return [4 /*yield*/, CheckNumberOfRows(context)];
        case 1:
          _a.sent();
          return [4 /*yield*/, context.sync()];
        case 2:
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}

function CheckNumberOfRows(context) {
  return __awaiter(this, void 0, void 0, function () {
    var rowCount;
    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          return [4 /*yield*/, totalTblRows()];
        case 1:
          rowCount = _a.sent();
          if (rowCount >= 950) {
            addAnalysisInfo("Row count number", rowCount, "Number of rows is very close(or above) the Diag Limit of 1000Rows", "CheckNumberOfRows", enumTypeAnalysis.Warning);
            AddMessage("Number of rows is very close to the Diag Limit of 1000Rows returned($tblRange.rowCount)");
          }
          return [4 /*yield*/, context.sync()];
        case 2:
          _a.sent();
          return [2 /*return*/];
      }
    });
  });
}

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
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\r\n<!-- This file shows how to design a first-run page that provides a welcome screen to the user about the features of the add-in. -->\r\n\r\n<!DOCTYPE html>\r\n<html>\r\n\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\r\n    <title>Contoso Task Pane Add-in</title>\r\n\r\n    <!-- Office JavaScript API -->\r\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js\"><" + "/script>\r\n\r\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\r\n    <!-- <link rel=\"stylesheet\" href=\"https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css\"/> -->\r\n\r\n\r\n    <!-- Bootstrap CSS -->\r\n    <link rel=\"stylesheet\" href=\"https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css\">\r\n\r\n    <!-- jQuery and Bootstrap JS -->\r\n    <" + "script src=\"https://code.jquery.com/jquery-3.5.1.slim.min.js\"\r\n        integrity=\"sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj\"\r\n        crossorigin=\"anonymous\"><" + "/script>\r\n\r\n    <!-- Template styles -->\r\n    <link href=\"" + ___HTML_LOADER_REPLACEMENT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n</head>\r\n\r\n<body class=\"bg-dark\">\r\n \r\n    <div class=\"bg-light rounded p-3 card\">\r\n        <div class=\"card-header\">\r\n            CDL Check provides compreehensive formatting of CDL logs as well as basic troubleshooting help\r\n        </div>\r\n        <div class=\"card-body\">\r\n            <form>\r\n                <div class=\"form-group\" id=\"sideload-msg\">\r\n                    <H2>officeJS is not loaded, Addin will not run!</H2>\r\n                </div>\r\n\r\n                <div class=\"form-group\">\r\n                    <!-- Spacer     -->\r\n                </div>\r\n\r\n                <div class=\"form-group\">\r\n                    <label for=\"typeCDL\">Choose Type of CDL Log</label>\r\n                    <select class=\"form-control\" id=\"typeCDL\">\r\n                        <option value=\"rave-diag-log\">RAVE Diag Log (Default)</option>\r\n                        <option value=\"raw-cdl\">RAW CDL from Get-CalendarDiagnosticObjects</option>\r\n                        <option value=\"onprem-exchange-cdl\">ONPrem Exchange CDL (future development)</option>\r\n                    </select>\r\n                </div>\r\n                \r\n                <div class=\"form-group form-check\">\r\n                <input type=\"checkbox\" class=\"form-check-input\" id=\"isOrganizer\">\r\n                <label class=\"form-check-label\" for=\"isOrganizer\">This is an Organizer CDL</label>\r\n                </div>\r\n                <div class=\"form-group form-check\">\r\n                    <input type=\"checkbox\" class=\"form-check-input\" id=\"warn1KRows\">\r\n                    <label class=\"form-check-label\" for=\"warn1KRows\">Warn if Rows are close to 1K</label>\r\n                </div>\r\n        \r\n                <div class=\"form-group\">\r\n                    <button type=\"button\" class=\"btn btn-primary\" id=\"run\">Run CDL Checks</button>\r\n                    <div class=\"spinner-border text-primary invisible\" role=\"status\" id=\"spinner\">\r\n                        <span class=\"sr-only\">Loading...</span>\r\n                    </div>\r\n                </div>\r\n                <div class=\"form-group\">\r\n                    <p class=\"mb-1 small\" id=\"statusMessage\">Status Message</p>\r\n                </div>\r\n                <div class=\"list-group\" id=\"analysisInfo\">\r\n                    <a href=\"#\" class=\"list-group-item list-group-item-action list-group-item-warning\">\r\n                        <div class=\"d-flex w-100 justify-content-between\">\r\n                            <h5 class=\"mb-1\">Row limit</h5>\r\n                            <span class=\"badge badge-primary badge-pill\">1002</span>\r\n                        </div>\r\n                        <p class=\"mb-1\">If rows returned are close to 1K, this can be a limitation of Rave diagnostics and could mean dataset returned is not complete. To avoid this, ask for raw CDL to the customer, and format directly with the option RAW</p>\r\n                        <small>Get-CalendarDiagnosticObjects</small>\r\n                    </a>\r\n                    <a href=\"#\" class=\"list-group-item list-group-item-action \">\r\n                    <div class=\"d-flex w-100 justify-content-between\">\r\n                        <h5 class=\"mb-1\">Created events</h5>\r\n                        <span class=\"badge badge-primary badge-pill\">2</span>\r\n                    </div>\r\n                    <p class=\"mb-1\">There are 2 events of type CREATED, which means that Exchange has created new calendar items to send to others</p>\r\n                    <small class=\"text-muted\">Donec id elit non mi porta.</small>\r\n                    </a>\r\n                    <a href=\"#\" class=\"list-group-item list-group-item-action list-group-item-danger\">\r\n                    <div class=\"d-flex w-100 justify-content-between\">\r\n                        <h5 class=\"mb-1\">Exception registered</h5>\r\n                        <span class=\"badge badge-primary badge-pill\">1</span>\r\n                    </div>\r\n                    <p class=\"mb-1\">[{\r\n                        \"resource\": \"/c:/MS/Rusilva/ExcelCalLogs/ExcelCalLogs/tsconfig.json\",\r\n                        \"owner\": \"_generated_diagnostic_collection_name_#0\",\r\n                        \"code\": {\r\n                            \"value\": \"typescript-config/consistent-casing\",\r\n                            \"target\": {\r\n                                \"$mid\": 1,\r\n                                \"path\": \"/docs/user-guide/hints/hint-typescript-config/consistent-casing/\",\r\n                                \"scheme\": \"https\",\r\n                                \"authority\": \"webhint.io\"\r\n                            }\r\n                        },\r\n                        \"severity\": 4,\r\n                        \"message\": \"The compiler option \\\"forceConsistentCasingInFileNames\\\" should be enabled to reduce issues when working with different OSes.\",\r\n                        \"source\": \"Microsoft Edge Tools\",\r\n                        \"startLineNumber\": 2,\r\n                        \"startColumn\": 4,\r\n                        \"endLineNumber\": 2,\r\n                        \"endColumn\": 19\r\n                    }][{\r\n            \"resource\": \"/c:/MS/Rusilva/ExcelCalLogs/ExcelCalLogs/tsconfig.json\",\r\n            \"owner\": \"_generated_diagnostic_collection_name_#0\",\r\n            \"code\": {\r\n                \"value\": \"typescript-config/consistent-casing\",\r\n                \"target\": {\r\n                    \"$mid\": 1,\r\n                    \"path\": \"/docs/user-guide/hints/hint-typescript-config/consistent-casing/\",\r\n                    \"scheme\": \"https\",\r\n                    \"authority\": \"webhint.io\"\r\n                }\r\n            },\r\n            \"severity\": 4,\r\n            \"message\": \"The compiler option \\\"forceConsistentCasingInFileNames\\\" should be enabled to reduce issues when working with different OSes.\",\r\n            \"source\": \"Microsoft Edge Tools\",\r\n            \"startLineNumber\": 2,\r\n            \"startColumn\": 4,\r\n            \"endLineNumber\": 2,\r\n            \"endColumn\": 19\r\n        }]</p>\r\n                    <small class=\"text-muted\">This will need a fix in the code.</small>\r\n                    </a>\r\n                </div>\r\n            </form>\r\n        </div>\r\n        <div class=\"card-footer\">\r\n            footer\r\n        </div>\r\n\r\n\r\n        <div class=\"card-header\">\r\n            testing multimedia\r\n        </div>\r\n        <div class=\"card-body \">\r\n            <div class=\"embed-responsive embed-responsive-21by9\">\r\n                <iframe class=\"embed-responsive-item\" width=\"560\" height=\"315\" src=\"https://www.youtube.com/embed/7C38muJjnyc\"\r\n                    title=\"YouTube video player\" frameborder=\"0\"\r\n                    allow=\"accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share\"\r\n                    allowfullscreen></iframe>\r\n            </div>\r\n        </div>\r\n\r\n    </div>\r\n</body>\r\n\r\n</html>\r\n";
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