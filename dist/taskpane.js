!function(){"use strict";var e={51081:function(e,t){Object.defineProperty(t,"__esModule",{value:!0}),t.JsonConfigUtils=void 0;var n=function(){function e(){this.columns=[]}return e.prototype.getValue=function(){return this.json},e.prototype.addColumn=function(e){return this.columns.push(e),e},e.prototype.addColumnByName=function(e,t,n,r,o,i,a,s,c,l){void 0===c&&(c=!0),void 0===l&&(l=!1);var u={columnName:e,isMandatory:t,horizontalAlignment:n,verticalAlignment:r,columnWidth:o,indentLevel:i,style:a,numberFormat:s,visible:c,autosizeColumn:l};return this.columns.push(u),u},e.prototype.convertColumnDefinitionsToJson=function(){return JSON.stringify(this.columns)},e.prototype.convertToHorizontalAlignment=function(e,t){switch(e){case"Center":t.format.horizontalAlignment=Excel.HorizontalAlignment.center;break;case"Left":t.format.horizontalAlignment=Excel.HorizontalAlignment.left;break;case"Right":t.format.horizontalAlignment=Excel.HorizontalAlignment.right;break;case"Justify":t.format.horizontalAlignment=Excel.HorizontalAlignment.justify;break;default:t.format.horizontalAlignment=Excel.HorizontalAlignment.general}},e.prototype.converttoverticalalignment=function(e){switch(e){case"center":return Excel.VerticalAlignment.bottom;case"left":return Excel.VerticalAlignment.center;case"right":return Excel.VerticalAlignment.distributed;case"justify":return Excel.VerticalAlignment.justify;default:return Excel.VerticalAlignment.top}},e}();t.JsonConfigUtils=n},88555:function(e,t,n){var r=this&&this.__awaiter||function(e,t,n,r){return new(n||(n=Promise))((function(o,i){function a(e){try{c(r.next(e))}catch(e){i(e)}}function s(e){try{c(r.throw(e))}catch(e){i(e)}}function c(e){var t;e.done?o(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(a,s)}c((r=r.apply(e,t||[])).next())}))},o=this&&this.__generator||function(e,t){var n,r,o,i,a={label:0,sent:function(){if(1&o[0])throw o[1];return o[1]},trys:[],ops:[]};return i={next:s(0),throw:s(1),return:s(2)},"function"==typeof Symbol&&(i[Symbol.iterator]=function(){return this}),i;function s(s){return function(c){return function(s){if(n)throw new TypeError("Generator is already executing.");for(;i&&(i=0,s[0]&&(a=0)),a;)try{if(n=1,r&&(o=2&s[0]?r.return:s[0]?r.throw||((o=r.return)&&o.call(r),0):r.next)&&!(o=o.call(r,s[1])).done)return o;switch(r=0,o&&(s=[2&s[0],o.value]),s[0]){case 0:case 1:o=s;break;case 4:return a.label++,{value:s[1],done:!1};case 5:a.label++,r=s[1],s=[0];continue;case 7:s=a.ops.pop(),a.trys.pop();continue;default:if(!((o=(o=a.trys).length>0&&o[o.length-1])||6!==s[0]&&2!==s[0])){a=0;continue}if(3===s[0]&&(!o||s[1]>o[0]&&s[1]<o[3])){a.label=s[1];break}if(6===s[0]&&a.label<o[1]){a.label=o[1],o=s;break}if(o&&a.label<o[2]){a.label=o[2],a.ops.push(s);break}o[2]&&a.ops.pop(),a.trys.pop();continue}s=t.call(e,a)}catch(e){s=[6,e],r=0}finally{n=o=0}if(5&s[0])throw s[1];return{value:s[0]?s[1]:void 0,done:!0}}([s,c])}}};Object.defineProperty(t,"__esModule",{value:!0}),t.run=t.testConfig=t.testJsonFile=t.createConfig=void 0;var i,a,s,c,l=n(51081),u=-1;function m(){return r(this,void 0,void 0,(function(){return o(this,(function(e){switch(e.label){case 0:return u<=0?(c.load(["rowCount"]),[4,i.sync()]):[3,2];case 1:e.sent(),u=c.rowCount,e.label=2;case 2:return[2,u]}}))}))}var d,f="TableStyleLight10",g="TableStyleLight13";function h(){return r(this,void 0,void 0,(function(){return o(this,(function(e){return[2,document.getElementById("isOrganizer").checked]}))}))}function v(e){return r(this,void 0,void 0,(function(){var t;return o(this,(function(n){return t=document.getElementById("spinner"),e?t.classList.remove("invisible"):t.classList.add("invisible"),[2]}))}))}function b(e){document.getElementById("statusMessage").textContent=e}function y(e,t,n,i,a){return r(this,void 0,void 0,(function(){var r,s,c,l,u,m,f;return o(this,(function(o){switch(r=document.getElementById("analysisInfo"),s=document.createElement("a"),a){case d.Warning:s.classList.add("list-group-item","list-group-item-action","list-group-item-warning");break;case d.Action:s.classList.add("list-group-item","list-group-item-action");break;case d.Danger:s.classList.add("list-group-item","list-group-item-action","list-group-item-danger");break;case d.Success:s.classList.add("list-group-item","list-group-item-action","list-group-item-success")}return(c=document.createElement("div")).classList.add("d-flex","w-100","justify-content-between"),(l=document.createElement("h5")).classList.add("mb-1"),l.innerText=e,(u=document.createElement("span")).classList.add("badge","badge-primary","badge-pill"),0==t&&u.classList.add("invisible"),u.innerText=t.toString(),c.appendChild(l),c.appendChild(u),(m=document.createElement("p")).classList.add("mb-1"),m.innerText=n,(f=document.createElement("small")).innerText=i,s.appendChild(c),s.appendChild(m),s.appendChild(f),r.appendChild(s),[2]}))}))}function p(){return r(this,void 0,void 0,(function(){return o(this,(function(e){return document.getElementById("analysisInfo").innerHTML="",[2]}))}))}function C(e,t){return void 0===t&&(t=!1),r(this,void 0,void 0,(function(){var n,r,i,a,c;return o(this,(function(o){switch(o.label){case 0:return o.trys.push([0,7,,8]),(n=s.getUsedRange()).load("address"),r=s.tables.getItemOrNullObject("CDL"),i=r.getRange(),[4,e.sync().catch((function(e){y("error",0,e,"create table",d.Danger),console.log(e)}))];case 1:return o.sent(),r.isNullObject?((r=s.tables.add(n,!0)).name="CDL",i=r.getRange(),[4,e.sync()]):[3,3];case 2:o.sent(),o.label=3;case 3:return t?[2]:(i.clear("Formats"),[4,e.sync()]);case 4:return o.sent(),[4,h()];case 5:return a=o.sent(),r.style=a?f:g,r.load("tableStyle"),r.columns.load(),i=r.getRange(),[4,e.sync()];case 6:return o.sent(),[3,8];case 7:return c=o.sent(),y("create Table",0,"Error creating table ".concat(c),"Create Table",d.Danger),console.error(c),b(c),[3,8];case 8:return[2]}}))}))}function w(e){return r(this,void 0,void 0,(function(){var t;return o(this,(function(n){switch(n.label){case 0:return a.columns.load(),[4,i.sync()];case 1:return n.sent(),t=a.columns.getItemOrNullObject("Ignorable").filter,[4,i.sync()];case 2:return n.sent(),t.apply({filterOn:Excel.FilterOn.values,values:[e]}),[4,i.sync()];case 3:return n.sent(),console.log("Cells filtered. "),[2]}}))}))}function x(){return r(this,void 0,void 0,(function(){var e,t;return o(this,(function(n){switch(n.label){case 0:return e=a.columns.getItemOrNullObject("Ignorable").getDataBodyRange(),(t=e.conditionalFormats.add(Excel.ConditionalFormatType.containsText)).textComparison.format.font.color="blue",t.textComparison.format.fill.color="#ADD8E6",t.textComparison.rule={operator:Excel.ConditionalTextOperator.contains,text:"TRUE"},[4,i.sync()];case 1:return n.sent(),[4,E(t.getRange())];case 2:return[2,n.sent()]}}))}))}function A(e){return r(this,void 0,void 0,(function(){var t,n,r,s;return o(this,(function(o){switch(o.label){case 0:return t=a.columns.getItemOrNullObject("ApptSequence"),n=t.getDataBodyRange(),[4,i.sync()];case 1:return o.sent(),r=n.conditionalFormats.add(Excel.ConditionalFormatType.colorScale),s={minimum:{formula:null,type:Excel.ConditionalFormatColorCriterionType.lowestValue,color:"white"},maximum:{formula:null,type:Excel.ConditionalFormatColorCriterionType.highestValue,color:"green"}},r.colorScale.criteria=s,[4,e.sync()];case 2:return o.sent(),[2]}}))}))}function E(e){return r(this,void 0,void 0,(function(){var t;return o(this,(function(n){switch(n.label){case 0:return[4,i.sync()];case 1:return n.sent(),[4,e.getIntersectionOrNullObject(a.getRange())];case 2:return(t=n.sent()).load(["rowCount"]),[4,i.sync()];case 3:return n.sent(),[2,t?t.rowCount:0]}}))}))}function O(e){return r(this,void 0,void 0,(function(){var t,n;return o(this,(function(r){switch(r.label){case 0:return t=a.columns.getItemOrNullObject("Client").getRange(),(n=t.conditionalFormats.add(Excel.ConditionalFormatType.containsText)).textComparison.format.font.color="red",n.textComparison.rule={operator:Excel.ConditionalTextOperator.contains,text:"CRA:CalendarRepairAssistant"},[4,e.sync()];case 1:return r.sent(),[4,E(n.getRange())];case 2:return[2,r.sent()]}}))}))}function z(e){return r(this,void 0,void 0,(function(){var t,n;return o(this,(function(r){switch(r.label){case 0:return t=a.columns.getItemOrNullObject("Trigger").getRange(),(n=t.conditionalFormats.add(Excel.ConditionalFormatType.containsText)).textComparison.format.fill.color="Green",n.textComparison.rule={operator:Excel.ConditionalTextOperator.contains,text:"Create"},[4,e.sync()];case 1:return r.sent(),[4,E(n.getRange())];case 2:return[2,r.sent()]}}))}))}function N(e,t){return void 0===t&&(t=!1),r(this,void 0,void 0,(function(){var n,r,s,c,l,u,m;return o(this,(function(o){switch(o.label){case 0:if("string"==typeof e){if(!function(e){try{var t=JSON.parse(e);return"object"==typeof t&&null!==t}catch(e){return!1}}(e))throw new Error("Invalid JSON string");n=JSON.parse(e)}else{if("object"!=typeof e)throw new Error("Parameter must be a string or an object");n=e}if(!Array.isArray(n))throw new Error("Input JSON is not an array.");o.label=1;case 1:o.trys.push([1,23,,24]),r=0,s=n,o.label=2;case 2:return r<s.length?null==(c=s[r]).columnName||""==c.columnName?(console.log("Skipping json element as it is undefined ColumnName"),[3,21]):(l=a.columns.getItemOrNullObject(c.columnName),u=l.getDataBodyRange(),i.trackedObjects.add([l,u]),[4,i.sync()]):[3,22];case 3:if(o.sent(),l.isNullObject)return void 0!==c.isMandatory&&""!==c.isMandatory&&"false"==c.isMandatory?(console.log("isMandatory: ".concat(c.isMandatory)),[3,21]):(console.log("Column Name does not exist: ".concat(c.columnName)),y("columnName",0,"Column Name does not exist: ".concat(c.columnName),"ValidateJSONStruct",d.Danger),[2,!1]);if(console.log("Column Name: ".concat(c.columnName)),void 0===c.horizontalAlignment||""===c.horizontalAlignment)return[3,5];switch(console.log("Horizontal Alignment: ".concat(c.horizontalAlignment)),c.horizontalAlignment){case"Center":u.format.horizontalAlignment=Excel.HorizontalAlignment.center;break;case"Left":u.format.horizontalAlignment=Excel.HorizontalAlignment.left;break;case"Right":u.format.horizontalAlignment=Excel.HorizontalAlignment.right;break;case"Justify":u.format.horizontalAlignment=Excel.HorizontalAlignment.justify;break;default:u.format.horizontalAlignment=Excel.HorizontalAlignment.general}return[4,i.sync()];case 4:o.sent(),o.label=5;case 5:return void 0===c.verticalAlignment||""===c.verticalAlignment?[3,7]:(console.log("Vertical Alignment: ".concat(c.verticalAlignment)),[4,i.sync()]);case 6:o.sent(),o.label=7;case 7:return void 0===c.columnWidth||null===c.columnWidth?[3,9]:(console.log("Column Width: ".concat(c.columnWidth)),u.format.columnWidth=c.columnWidth,[4,i.sync()]);case 8:o.sent(),o.label=9;case 9:return void 0===c.indentLevel||null===c.indentLevel?[3,11]:(console.log("Ident Level: ".concat(c.indentLevel)),u.format.indentLevel=c.indentLevel,[4,i.sync()]);case 10:o.sent(),o.label=11;case 11:return void 0===c.style||""===c.style?[3,13]:(console.log("Style: ".concat(c.style)),u.style=c.style,[4,i.sync()]);case 12:o.sent(),o.label=13;case 13:return void 0===c.numberFormat||""===c.numberFormat?[3,15]:(console.log("Style: ".concat(c.numberFormat)),u.numberFormat=c.numberFormat,[4,i.sync()]);case 14:o.sent(),o.label=15;case 15:return void 0===c.visible||null===c.visible?[3,17]:(console.log("Visible: ".concat(c.visible)),!c.visible&&t&&(u.columnHidden=!0),[4,i.sync()]);case 16:o.sent(),o.label=17;case 17:return void 0===c.autosizeColumn||null===c.autosizeColumn?[3,19]:(console.log("autosizeColumn: ".concat(c.autosizeColumn)),"true"===c.autosizeColumn&&u.format.autofitColumns(),[4,i.sync()]);case 18:o.sent(),o.label=19;case 19:return[4,i.sync()];case 20:o.sent(),console.log("removing tracked objects for ".concat(c.columnName)),i.trackedObjects.remove([l,u]),o.label=21;case 21:return r++,[3,2];case 22:return[2,!0];case 23:return m=o.sent(),y("columnName",0,"Error traversing JSON array: ".concat(m),"ValidateJSONStruct",d.Danger),console.error("Error traversing JSON array: ".concat(m)),[2,!1];case 24:return[2]}}))}))}function L(){return r(this,void 0,void 0,(function(){var e,t,n,r,s,c,u,m,d,f,g,h,v,b,y;return o(this,(function(o){switch(o.label){case 0:return e=new l.JsonConfigUtils,t=a.columns.load(["name","values/format","values/horizontalAlignment","values/verticalAlignment","values/columnWidth","values/indentLevel","values/style","values/numberFormat","values/autosizeColumn"]),[4,i.sync()];case 1:o.sent(),n=0,r=t.items,o.label=2;case 2:return n<r.length?(s=r[n],(c=s.getRange().getCell(0,0)).load(["columnHidden"]),[4,i.sync()]):[3,6];case 3:return o.sent(),c.columnHidden,(u=s.getDataBodyRange()).load(["style","numberFormat"]),u.format.load(["format","horizontalAlignment","verticalAlignment","columnWidth","indentLevel","style","numberFormat"]),[4,i.sync()];case 4:o.sent(),m=s.name,u.format,d=u.format.horizontalAlignment,f=u.format.verticalAlignment,g=u.format.columnWidth,h=u.format.indentLevel,v=u.style,b=u.numberFormat[0].toString(),y={columnName:m,isMandatory:!0,horizontalAlignment:d,verticalAlignment:f,columnWidth:g,indentLevel:h,style:v,numberFormat:b,visible:!0,autosizeColumn:!1},e.addColumn(y),o.label=5;case 5:return n++,[3,2];case 6:return[2,e]}}))}))}function I(e){return r(this,void 0,void 0,(function(){return o(this,(function(t){switch(t.label){case 0:return[4,p()];case 1:return t.sent(),[4,S()];case 2:return t.sent(),[4,e.sync()];case 3:return t.sent(),[2]}}))}))}function S(){return r(this,void 0,void 0,(function(){var e;return o(this,(function(t){switch(t.label){case 0:return S?[4,m()]:[2];case 1:return(e=t.sent())>=950&&(y("Row count number",e,"Number of rows is very close(or above) the Diag Limit of 1000Rows","CheckNumberOfRows",d.Warning),b("Number of rows is very close to the Diag Limit of 1000Rows returned($tblRange.rowCount)")),[4,i.sync()];case 2:return t.sent(),[2]}}))}))}function j(e){return r(this,void 0,void 0,(function(){return o(this,(function(t){switch(t.label){case 0:return a.columns.getItem(e),s.freezePanes.freezeRows(1),s.freezePanes.freezeColumns(3),[4,i.sync()];case 1:return t.sent(),[2]}}))}))}function D(){return r(this,void 0,void 0,(function(){return o(this,(function(e){switch(e.label){case 0:return[4,fetch("./RaveCDLconfig.json")];case 1:return[4,e.sent().json()];case 2:return[2,e.sent()]}}))}))}function R(){return r(this,void 0,void 0,(function(){var e=this;return o(this,(function(t){switch(t.label){case 0:return[4,Excel.run((function(t){return r(e,void 0,void 0,(function(){var e;return o(this,(function(n){switch(n.label){case 0:return i=t,s=t.workbook.worksheets.getActiveWorksheet(),a=s.tables.getItemOrNullObject("CDL"),c=a.getRange(),[4,t.sync()];case 1:return n.sent(),[4,C(t,!0)];case 2:return n.sent(),[4,L()];case 3:return e=n.sent(),document.getElementById("jsonConfig").textContent=e.convertColumnDefinitionsToJson(),[2]}}))}))}))];case 1:return t.sent(),[2]}}))}))}function k(){return r(this,void 0,void 0,(function(){var e;return o(this,(function(t){switch(t.label){case 0:return[4,D()];case 1:return e=t.sent(),document.getElementById("jsonConfig").textContent=JSON.stringify(e),[4,N(e)];case 2:return t.sent(),[2]}}))}))}function F(){return r(this,void 0,void 0,(function(){return o(this,(function(e){switch(e.label){case 0:return[4,N(document.getElementById("jsonConfig").textContent)];case 1:return e.sent(),[2]}}))}))}function T(){return r(this,void 0,void 0,(function(){var e,t=this;return o(this,(function(n){switch(n.label){case 0:return n.trys.push([0,2,,3]),[4,Excel.run((function(e){return r(t,void 0,void 0,(function(){return o(this,(function(t){switch(t.label){case 0:return[4,v(!0)];case 1:return t.sent(),[4,p()];case 2:return t.sent(),[4,b("Starting Processing")];case 3:return t.sent(),i=e,s=e.workbook.worksheets.getActiveWorksheet(),a=s.tables.getItemOrNullObject("CDL"),c=a.getRange(),new l.JsonConfigUtils,s.getRange().clear("Formats"),s.getRange().conditionalFormats.clearAll(),a.clearFilters(),s.freezePanes.unfreeze(),[4,e.sync()];case 4:return t.sent(),[4,C(e).then((function(){b("Create Table Done")}))];case 5:return t.sent(),[4,D()];case 6:return[4,N(t.sent())];case 7:return t.sent()?[4,j("Ignorable")]:(y("CDL Invalid",0,"CDL Structure is invalid (check previous exceptions)","CDLInvalid",d.Danger),v(!1),[2]);case 8:return t.sent(),[4,x().then((function(){b("Highlight Ignorable Done")}))];case 9:return t.sent(),[4,A(e).then((function(){b("Highlight  Done")}))];case 10:return t.sent(),[4,O(e).then((function(){b("Highlight CRA Done")}))];case 11:return t.sent(),[4,z(e).then((function(){b("Highlight Create Done")}))];case 12:return t.sent(),[4,w("FALSE").then((function(){b("Filter Ignorable Done")}))];case 13:return t.sent(),[4,e.sync()];case 14:return t.sent(),[4,I(e).then((function(){b("Perform Analysis Done")}))];case 15:return t.sent(),console.log("Processing done."),b("Done!"),v(!1),y("Success",0,"Process executed successfully, check the video on CDL analysis ".concat("https://msit.microsoftstream.com/video/4221a4ff-0400-9fb2-4805-f1eb0f28f09b"," "),"success",d.Success),[2]}}))}))}))];case 1:return n.sent(),[3,3];case 2:return e=n.sent(),v(!1),console.error(e),b(e),y("Error",0,e,"Run/Catch",d.Danger),[3,3];case 3:return[2]}}))}))}!function(e){e[e.Warning=0]="Warning",e[e.Action=1]="Action",e[e.Danger=2]="Danger",e[e.Success=3]="Success"}(d||(d={})),Office.onReady((function(e){e.host===Office.HostType.Excel&&(document.getElementById("sideload-msg").style.display="none",document.getElementById("run").onclick=T,document.getElementById("createConfig").onclick=R,document.getElementById("testConfig").onclick=F,document.getElementById("testJsonFile").onclick=k,$((function(){$('[data-toggle="tooltip"]').tooltip()})),p())})),t.createConfig=R,t.testJsonFile=k,t.testConfig=F,t.run=T},93823:function(e,t,n){var r=n(27091),o=n.n(r),i=new URL(n(60806),n.b);o()(i)},27091:function(e){e.exports=function(e,t){return t||(t={}),e?(e=String(e.__esModule?e.default:e),t.hash&&(e+=t.hash),t.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(e)?'"'.concat(e,'"'):e):e}},60806:function(e,t,n){e.exports=n.p+"4f424550f2dc5a27a461.css"}},t={};function n(r){var o=t[r];if(void 0!==o)return o.exports;var i=t[r]={exports:{}};return e[r].call(i.exports,i,i.exports,n),i.exports}n.m=e,n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,{a:t}),t},n.d=function(e,t){for(var r in t)n.o(t,r)&&!n.o(e,r)&&Object.defineProperty(e,r,{enumerable:!0,get:t[r]})},n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;n.g.importScripts&&(e=n.g.location+"");var t=n.g.document;if(!e&&t&&(t.currentScript&&(e=t.currentScript.src),!e)){var r=t.getElementsByTagName("script");r.length&&(e=r[r.length-1].src)}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),n.p=e}(),n.b=document.baseURI||self.location.href,n(88555),n(93823)}();
//# sourceMappingURL=taskpane.js.map