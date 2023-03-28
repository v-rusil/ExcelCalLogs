/******/ (function() { // webpackBootstrap
/******/ 	// The require scope
/******/ 	var __webpack_require__ = {};
/******/ 	
/************************************************************************/
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
/************************************************************************/
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/commands/commands.ts ***!
  \**********************************/
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global global, Office, self, window */
Office.onReady(function () {
  // If needed, Office.js is ready to be called
});
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event) {
  var message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true
  };
  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
function getGlobal() {
  return typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : typeof __webpack_require__.g !== "undefined" ? __webpack_require__.g : undefined;
}
var g = getGlobal();
// The add-in command functions need to be available in global scope
g.action = action;
/******/ })()
;
//# sourceMappingURL=commands.js.map