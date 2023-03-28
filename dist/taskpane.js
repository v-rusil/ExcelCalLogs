/******/ (function() { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ "./src/taskpane/taskpane.ts":
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.ts ***!
  \**********************************/
/***/ (function() {

throw new Error("Module build failed (from ./node_modules/ts-loader/index.js):\nError: TypeScript emitted no output for C:\\MS\\Rusilva\\ExcelCalLogs\\ExcelCalLogs\\src\\taskpane\\taskpane.ts.\n    at makeSourceMapAndFinish (C:\\MS\\Rusilva\\ExcelCalLogs\\ExcelCalLogs\\node_modules\\ts-loader\\dist\\index.js:52:18)\n    at successLoader (C:\\MS\\Rusilva\\ExcelCalLogs\\ExcelCalLogs\\node_modules\\ts-loader\\dist\\index.js:39:5)\n    at Object.loader (C:\\MS\\Rusilva\\ExcelCalLogs\\ExcelCalLogs\\node_modules\\ts-loader\\dist\\index.js:22:5)");

/***/ }),

/***/ "./src/taskpane/taskpane.html":
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
/***/ (function(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../node_modules/html-loader/dist/runtime/getUrl.js */ "./node_modules/html-loader/dist/runtime/getUrl.js");
/* harmony import */ var _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0__);
// Imports

var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_1___ = new URL(/* asset import */ __webpack_require__(/*! ../../assets/logo-filled.png */ "./assets/logo-filled.png"), __webpack_require__.b);
// Module
var ___HTML_LOADER_REPLACEMENT_0___ = _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default()(___HTML_LOADER_IMPORT_0___);
var ___HTML_LOADER_REPLACEMENT_1___ = _node_modules_html_loader_dist_runtime_getUrl_js__WEBPACK_IMPORTED_MODULE_0___default()(___HTML_LOADER_IMPORT_1___);
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\r\n<!-- This file shows how to design a first-run page that provides a welcome screen to the user about the features of the add-in. -->\r\n\r\n<!DOCTYPE html>\r\n<html>\r\n\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\r\n    <title>Contoso Task Pane Add-in</title>\r\n\r\n    <!-- Office JavaScript API -->\r\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js\"><" + "/script>\r\n\r\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\r\n    <link rel=\"stylesheet\" href=\"https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css\"/>\r\n\r\n    <!-- Template styles -->\r\n    <link href=\"" + ___HTML_LOADER_REPLACEMENT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n</head>\r\n\r\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\r\n    <header class=\"ms-welcome__header ms-bgColor-neutralLighter\">\r\n        <img width=\"90\" height=\"90\" src=\"" + ___HTML_LOADER_REPLACEMENT_1___ + "\" alt=\"Contoso\" title=\"Contoso\" />\r\n        <h1 class=\"ms-font-su\">CDL Formatter</h1>\r\n        <h2 class=\"ms-font-xs\">POC version</h2>\r\n    </header>\r\n    <section id=\"sideload-msg\" class=\"ms-welcome__main\">\r\n        <h2 class=\"ms-font-xl\">Please sideload your add-in to see app body.</h2>\r\n    </section>\r\n    <main id=\"app-body\" class=\"ms-welcome__main\" style=\"display: none;\">\r\n        <h2 class=\"ms-font-xl\"> This is pre Beta CDL Logs Formatter </h2>\r\n        <ul class=\"ms-List ms-welcome__features\">\r\n            <li class=\"ms-ListItem\">\r\n                <i class=\"ms-Icon ms-Icon--Ribbon ms-font-xl\"></i>\r\n                <span class=\"ms-font-m\">Inspect and visual format CDLLogs </span>\r\n            </li>\r\n            <li class=\"ms-ListItem\">\r\n                <i class=\"ms-Icon ms-Icon--Unlock ms-font-xl\"></i>\r\n                <span class=\"ms-font-m\">Video training on usage:</span>\r\n            </li>\r\n            <li class=\"ms-ListItem\">\r\n                <i class=\"ms-Icon ms-Icon--Design ms-font-xl\"></i>\r\n                <span class=\"ms-font-m\">formatter by rusilva@microsoft.com</span>\r\n            </li>\r\n        </ul>\r\n        <p class=\"ms-font-l\">Get the CDLLogs into A1 Cell, then click <b>Run</b>.</p>\r\n        <div role=\"button\" id=\"run\" class=\"ms-welcome__action ms-Button ms-Button--hero ms-font-xl\">\r\n            <span class=\"ms-Button-label\">Run</span>\r\n        </div>\r\n    </main>\r\n</body>\r\n\r\n</html>\r\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);

/***/ }),

/***/ "./node_modules/html-loader/dist/runtime/getUrl.js":
/*!*********************************************************!*\
  !*** ./node_modules/html-loader/dist/runtime/getUrl.js ***!
  \*********************************************************/
/***/ (function(module) {

"use strict";


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

/***/ "./assets/logo-filled.png":
/*!********************************!*\
  !*** ./assets/logo-filled.png ***!
  \********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "assets/logo-filled.png";

/***/ }),

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
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
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
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
/******/ 	// This entry module doesn't tell about it's top-level declarations so it can't be inlined
/******/ 	__webpack_require__("./src/taskpane/taskpane.ts");
/******/ 	var __webpack_exports__ = __webpack_require__("./src/taskpane/taskpane.html");
/******/ 	
/******/ })()
;
//# sourceMappingURL=taskpane.js.map