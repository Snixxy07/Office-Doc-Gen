/*! For license information please see c10fbed9fa3d995696a7.js.LICENSE.txt */
function _typeof(e){return _typeof="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(e){return typeof e}:function(e){return e&&"function"==typeof Symbol&&e.constructor===Symbol&&e!==Symbol.prototype?"symbol":typeof e},_typeof(e)}function _regeneratorRuntime(){"use strict";_regeneratorRuntime=function(){return t};var e,t={},r=Object.prototype,n=r.hasOwnProperty,o=Object.defineProperty||function(e,t,r){e[t]=r.value},a="function"==typeof Symbol?Symbol:{},c=a.iterator||"@@iterator",i=a.asyncIterator||"@@asyncIterator",u=a.toStringTag||"@@toStringTag";function l(e,t,r){return Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}),e[t]}try{l({},"")}catch(e){l=function(e,t,r){return e[t]=r}}function s(e,t,r,n){var a=t&&t.prototype instanceof g?t:g,c=Object.create(a.prototype),i=new S(n||[]);return o(c,"_invoke",{value:L(e,r,i)}),c}function p(e,t,r){try{return{type:"normal",arg:e.call(t,r)}}catch(e){return{type:"throw",arg:e}}}t.wrap=s;var f="suspendedStart",d="suspendedYield",m="executing",h="completed",y={};function g(){}function v(){}function x(){}var b={};l(b,c,(function(){return this}));var _=Object.getPrototypeOf,w=_&&_(_(N([])));w&&w!==r&&n.call(w,c)&&(b=w);var E=x.prototype=g.prototype=Object.create(b);function F(e){["next","throw","return"].forEach((function(t){l(e,t,(function(e){return this._invoke(t,e)}))}))}function T(e,t){function r(o,a,c,i){var u=p(e[o],e,a);if("throw"!==u.type){var l=u.arg,s=l.value;return s&&"object"==_typeof(s)&&n.call(s,"__await")?t.resolve(s.__await).then((function(e){r("next",e,c,i)}),(function(e){r("throw",e,c,i)})):t.resolve(s).then((function(e){l.value=e,c(l)}),(function(e){return r("throw",e,c,i)}))}i(u.arg)}var a;o(this,"_invoke",{value:function(e,n){function o(){return new t((function(t,o){r(e,n,t,o)}))}return a=a?a.then(o,o):o()}})}function L(t,r,n){var o=f;return function(a,c){if(o===m)throw Error("Generator is already running");if(o===h){if("throw"===a)throw c;return{value:e,done:!0}}for(n.method=a,n.arg=c;;){var i=n.delegate;if(i){var u=D(i,n);if(u){if(u===y)continue;return u}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if(o===f)throw o=h,n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);o=m;var l=p(t,r,n);if("normal"===l.type){if(o=n.done?h:d,l.arg===y)continue;return{value:l.arg,done:n.done}}"throw"===l.type&&(o=h,n.method="throw",n.arg=l.arg)}}}function D(t,r){var n=r.method,o=t.iterator[n];if(o===e)return r.delegate=null,"throw"===n&&t.iterator.return&&(r.method="return",r.arg=e,D(t,r),"throw"===r.method)||"return"!==n&&(r.method="throw",r.arg=new TypeError("The iterator does not provide a '"+n+"' method")),y;var a=p(o,t.iterator,r.arg);if("throw"===a.type)return r.method="throw",r.arg=a.arg,r.delegate=null,y;var c=a.arg;return c?c.done?(r[t.resultName]=c.value,r.next=t.nextLoc,"return"!==r.method&&(r.method="next",r.arg=e),r.delegate=null,y):c:(r.method="throw",r.arg=new TypeError("iterator result is not an object"),r.delegate=null,y)}function I(e){var t={tryLoc:e[0]};1 in e&&(t.catchLoc=e[1]),2 in e&&(t.finallyLoc=e[2],t.afterLoc=e[3]),this.tryEntries.push(t)}function k(e){var t=e.completion||{};t.type="normal",delete t.arg,e.completion=t}function S(e){this.tryEntries=[{tryLoc:"root"}],e.forEach(I,this),this.reset(!0)}function N(t){if(t||""===t){var r=t[c];if(r)return r.call(t);if("function"==typeof t.next)return t;if(!isNaN(t.length)){var o=-1,a=function r(){for(;++o<t.length;)if(n.call(t,o))return r.value=t[o],r.done=!1,r;return r.value=e,r.done=!0,r};return a.next=a}}throw new TypeError(_typeof(t)+" is not iterable")}return v.prototype=x,o(E,"constructor",{value:x,configurable:!0}),o(x,"constructor",{value:v,configurable:!0}),v.displayName=l(x,u,"GeneratorFunction"),t.isGeneratorFunction=function(e){var t="function"==typeof e&&e.constructor;return!!t&&(t===v||"GeneratorFunction"===(t.displayName||t.name))},t.mark=function(e){return Object.setPrototypeOf?Object.setPrototypeOf(e,x):(e.__proto__=x,l(e,u,"GeneratorFunction")),e.prototype=Object.create(E),e},t.awrap=function(e){return{__await:e}},F(T.prototype),l(T.prototype,i,(function(){return this})),t.AsyncIterator=T,t.async=function(e,r,n,o,a){void 0===a&&(a=Promise);var c=new T(s(e,r,n,o),a);return t.isGeneratorFunction(r)?c:c.next().then((function(e){return e.done?e.value:c.next()}))},F(E),l(E,u,"Generator"),l(E,c,(function(){return this})),l(E,"toString",(function(){return"[object Generator]"})),t.keys=function(e){var t=Object(e),r=[];for(var n in t)r.push(n);return r.reverse(),function e(){for(;r.length;){var n=r.pop();if(n in t)return e.value=n,e.done=!1,e}return e.done=!0,e}},t.values=N,S.prototype={constructor:S,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=e,this.done=!1,this.delegate=null,this.method="next",this.arg=e,this.tryEntries.forEach(k),!t)for(var r in this)"t"===r.charAt(0)&&n.call(this,r)&&!isNaN(+r.slice(1))&&(this[r]=e)},stop:function(){this.done=!0;var e=this.tryEntries[0].completion;if("throw"===e.type)throw e.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var r=this;function o(n,o){return i.type="throw",i.arg=t,r.next=n,o&&(r.method="next",r.arg=e),!!o}for(var a=this.tryEntries.length-1;a>=0;--a){var c=this.tryEntries[a],i=c.completion;if("root"===c.tryLoc)return o("end");if(c.tryLoc<=this.prev){var u=n.call(c,"catchLoc"),l=n.call(c,"finallyLoc");if(u&&l){if(this.prev<c.catchLoc)return o(c.catchLoc,!0);if(this.prev<c.finallyLoc)return o(c.finallyLoc)}else if(u){if(this.prev<c.catchLoc)return o(c.catchLoc,!0)}else{if(!l)throw Error("try statement without catch or finally");if(this.prev<c.finallyLoc)return o(c.finallyLoc)}}}},abrupt:function(e,t){for(var r=this.tryEntries.length-1;r>=0;--r){var o=this.tryEntries[r];if(o.tryLoc<=this.prev&&n.call(o,"finallyLoc")&&this.prev<o.finallyLoc){var a=o;break}}a&&("break"===e||"continue"===e)&&a.tryLoc<=t&&t<=a.finallyLoc&&(a=null);var c=a?a.completion:{};return c.type=e,c.arg=t,a?(this.method="next",this.next=a.finallyLoc,y):this.complete(c)},complete:function(e,t){if("throw"===e.type)throw e.arg;return"break"===e.type||"continue"===e.type?this.next=e.arg:"return"===e.type?(this.rval=this.arg=e.arg,this.method="return",this.next="end"):"normal"===e.type&&t&&(this.next=t),y},finish:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var r=this.tryEntries[t];if(r.finallyLoc===e)return this.complete(r.completion,r.afterLoc),k(r),y}},catch:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var r=this.tryEntries[t];if(r.tryLoc===e){var n=r.completion;if("throw"===n.type){var o=n.arg;k(r)}return o}}throw Error("illegal catch attempt")},delegateYield:function(t,r,n){return this.delegate={iterator:N(t),resultName:r,nextLoc:n},"next"===this.method&&(this.arg=e),y}},t}function asyncGeneratorStep(e,t,r,n,o,a,c){try{var i=e[a](c),u=i.value}catch(e){return void r(e)}i.done?t(u):Promise.resolve(u).then(n,o)}function _asyncToGenerator(e){return function(){var t=this,r=arguments;return new Promise((function(n,o){var a=e.apply(t,r);function c(e){asyncGeneratorStep(a,n,o,c,i,"next",e)}function i(e){asyncGeneratorStep(a,n,o,c,i,"throw",e)}c(void 0)}))}}function _slicedToArray(e,t){return _arrayWithHoles(e)||_iterableToArrayLimit(e,t)||_unsupportedIterableToArray(e,t)||_nonIterableRest()}function _nonIterableRest(){throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")}function _unsupportedIterableToArray(e,t){if(e){if("string"==typeof e)return _arrayLikeToArray(e,t);var r={}.toString.call(e).slice(8,-1);return"Object"===r&&e.constructor&&(r=e.constructor.name),"Map"===r||"Set"===r?Array.from(e):"Arguments"===r||/^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(r)?_arrayLikeToArray(e,t):void 0}}function _arrayLikeToArray(e,t){(null==t||t>e.length)&&(t=e.length);for(var r=0,n=Array(t);r<t;r++)n[r]=e[r];return n}function _iterableToArrayLimit(e,t){var r=null==e?null:"undefined"!=typeof Symbol&&e[Symbol.iterator]||e["@@iterator"];if(null!=r){var n,o,a,c,i=[],u=!0,l=!1;try{if(a=(r=r.call(e)).next,0===t){if(Object(r)!==r)return;u=!1}else for(;!(u=(n=a.call(r)).done)&&(i.push(n.value),i.length!==t);u=!0);}catch(e){l=!0,o=e}finally{try{if(!u&&null!=r.return&&(c=r.return(),Object(c)!==c))return}finally{if(l)throw o}}return i}}function _arrayWithHoles(e){if(Array.isArray(e))return e}var defaultFop={fop:"Проценко Юрій Ігорович",sex:"m",inn:"3289205817",registrationDate:"06.12.2023",registrationNumber:"2005560000000181053",address:"65025, Одеська обл., місто Одеса, пр. Добровольського, будинок 137, квартира 55",accountNumber:"UA113071230000026004011398566",bank:"БАНК ВОСТОК",bankAbbreviation:"ПАТ"};function initializeEventListeners(){document.querySelector("#addFopBtn").onclick=function(){return tryCatch(showAddFopForm)},document.querySelector("#closeAddFopForm").onclick=function(){return tryCatch(closeAddFopForm)},document.getElementById("fop").addEventListener("blur",formatFopName),document.getElementById("contractDate").addEventListener("change",updateContractEndDate),document.getElementById("replaceForm").onsubmit=function(e){e.preventDefault(),console.log("Form submitted"),tryCatch(replaceData)},document.getElementById("addFopForm").onsubmit=function(e){e.preventDefault(),tryCatch(saveFormData)};var e=document.getElementById("contractNumber");e.value=loadLastContractNumber(),e.oninput=function(e){handleContractNumberChange(e)}}function loadLastContractNumber(){return localStorage.getItem("lastContractNumber")||""}function handleContractNumberChange(e){saveContractNumber(e.target.value)}function saveContractNumber(e){localStorage.setItem("lastContractNumber",e)}function showAddFopForm(){document.getElementById("addFopForm").classList.remove("hidden")}function closeAddFopForm(){document.getElementById("addFopForm").classList.add("hidden")}function formatFopName(){this.value=this.value.toLowerCase().split(" ").map((function(e){return e.charAt(0).toUpperCase()+e.slice(1)})).join(" ")}function updateContractEndDate(){var e=document.getElementById("contractDate"),t=document.getElementById("contractEndDate");if(e.value){var r=new Date(e.value),n=new Date(r.getFullYear()+1,r.getMonth(),r.getDate()+1);t.value=n.toISOString().split("T")[0]}else t.value=""}function validateFormData(e){for(var t=0,r=Object.entries(e);t<r.length;t++){var n=_slicedToArray(r[t],2);if(n[0],!n[1])return!1}return!!/^\d{10}$/.test(e.inn)&&!!/^UA\d{27}$/.test(e.accountNumber)}function saveFormData(){var e={fop:document.getElementById("fop").value,sex:document.getElementById("sex").value,inn:document.getElementById("inn").value,registrationDate:document.getElementById("registrationDate").value,registrationNumber:document.getElementById("registrationNumber").value,address:document.getElementById("address").value,accountNumber:document.getElementById("accountNumber").value,bank:document.getElementById("bank").value,bankAbbreviation:document.getElementById("bankAbbreviation").value};if(validateFormData(e)){var t=JSON.parse(localStorage.getItem("fopDataArray"))||[],r=t.findIndex((function(t){return t.inn===e.inn}));-1!==r?t[r]=e:t.push(e),localStorage.setItem("fopDataArray",JSON.stringify(t)),console.log("Settings saved."),document.getElementById("addFopForm").reset(),closeAddFopForm(),populateOurFopSelect()}}function getAllFops(){return(JSON.parse(localStorage.getItem("fopDataArray"))||[]).reduce((function(e,t){return e[t.inn]=t,e}),{})}function populateOurFopSelect(){var e=getAllFops(),t=document.getElementById("ourFop");t.innerHTML="";var r=!0;for(var n in e){var o=document.createElement("option");o.value=n,o.textContent=e[n].fop,r&&(o.selected=!0,r=!1),t.appendChild(o)}if(0===t.options.length){var a=document.createElement("option");a.value="",a.textContent="Наш ФОП",t.appendChild(a)}}function replaceData(){return _replaceData.apply(this,arguments)}function _replaceData(){return(_replaceData=_asyncToGenerator(_regeneratorRuntime().mark((function e(){var t,r,n,o,a,c;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return t=document.getElementById("ourFop").value,r=document.getElementById("contractNumber").value,n=document.getElementById("contractDate").value,o=document.getElementById("contractEndDate").value,a=getAllFops(),c=a[t]||defaultFop,e.next=8,replaceFopData(c);case 8:return e.next=10,replaceContractData(r,n,o);case 10:console.log("Data replacement completed.");case 11:case"end":return e.stop()}}),e)})))).apply(this,arguments)}function replaceFopData(e){return _replaceFopData.apply(this,arguments)}function _replaceFopData(){return(_replaceFopData=_asyncToGenerator(_regeneratorRuntime().mark((function e(t){return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.next=2,replaceText(defaultFop.fop,t.fop);case 2:return e.next=4,replaceText(defaultFop.inn,t.inn);case 4:return e.next=6,replaceText(proceedSex(defaultFop.sex),proceedSex(t.sex),!0);case 6:return e.next=8,replaceText(defaultFop.registrationDate,t.registrationDate);case 8:return e.next=10,replaceText(defaultFop.registrationNumber,t.registrationNumber);case 10:return e.next=12,replaceText(defaultFop.address,t.address);case 12:return e.next=14,replaceText(defaultFop.accountNumber,t.accountNumber);case 14:return e.next=16,replaceText(formatBankName(defaultFop.bank,defaultFop.bankAbbreviation),formatBankName(t.bank,t.bankAbbreviation));case 16:case"end":return e.stop()}}),e)})))).apply(this,arguments)}function replaceContractData(e,t,r){return _replaceContractData.apply(this,arguments)}function _replaceContractData(){return(_replaceContractData=_asyncToGenerator(_regeneratorRuntime().mark((function e(t,r,n){return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if(!t){e.next=3;break}return e.next=3,replaceTextRegex("([0-9]@)/24",t+"/24");case 3:if(!r){e.next=6;break}return e.next=6,replaceTextRegex("«[0-9]{2}» ([!0-9]@) 2024",formatDateUkrainian(r));case 6:if(!n){e.next=9;break}return e.next=9,replaceTextRegex("«[0-9]{2}» ([!0-9]@) 2025",formatDateUkrainian(n));case 9:case"end":return e.stop()}}),e)})))).apply(this,arguments)}function proceedSex(e){return"m"===e?"який":"яка"}function formatBankName(e,t){return"".concat(t," «").concat(e,"»")}function formatDateUkrainian(e){var t=new Date(e),r=t.getDate().toString().padStart(2,"0"),n=t.getFullYear();return"«".concat(r,"» ").concat(["січня","лютого","березня","квітня","травня","червня","липня","серпня","вересня","жовтня","листопада","грудня"][t.getMonth()]," ").concat(n)}function replaceText(e,t){return _replaceText.apply(this,arguments)}function _replaceText(){return _replaceText=_asyncToGenerator(_regeneratorRuntime().mark((function e(t,r){var n,o=arguments;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return n=o.length>2&&void 0!==o[2]&&o[2],e.next=3,Word.run(function(){var e=_asyncToGenerator(_regeneratorRuntime().mark((function e(o){var a,c;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,a=o.document.body,(c=a.search(t,{matchCase:!1,matchWholeWord:!0})).load("items"),e.next=6,o.sync();case 6:return c.items.length>0?(n?c.items[0].insertText(r,Word.InsertLocation.replace):c.items.forEach((function(e){e.insertText(r,Word.InsertLocation.replace)})),console.log("Replaced ".concat(n?"first occurrence":c.items.length+" occurrences",' of "').concat(t,'" with "').concat(r,'"'))):console.log('No matches found for "'.concat(t,'"')),e.next=9,o.sync();case 9:e.next=14;break;case 11:e.prev=11,e.t0=e.catch(0),console.log("Error in replaceText:"+e.t0);case 14:case"end":return e.stop()}}),e,null,[[0,11]])})));return function(t){return e.apply(this,arguments)}}());case 3:case"end":return e.stop()}}),e)}))),_replaceText.apply(this,arguments)}function replaceTextRegex(e,t){return _replaceTextRegex.apply(this,arguments)}function _replaceTextRegex(){return _replaceTextRegex=_asyncToGenerator(_regeneratorRuntime().mark((function e(t,r){var n,o=arguments;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return n=o.length>2&&void 0!==o[2]&&o[2],e.next=3,Word.run(function(){var e=_asyncToGenerator(_regeneratorRuntime().mark((function e(o){var a,c;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,a=o.document.body,(c=a.search(t,{matchWildcards:!0})).load("items"),e.next=6,o.sync();case 6:return c.items.length>0?(n?c.items[0].insertText(r,Word.InsertLocation.replace):c.items.forEach((function(e){e.insertText(r,Word.InsertLocation.replace)})),console.log("Replaced ".concat(n?"first occurrence":c.items.length+" occurrences",' matching "').concat(t,'" with "').concat(r,'"'))):console.log('No matches found for "'.concat(t,'"')),e.next=9,o.sync();case 9:e.next=14;break;case 11:e.prev=11,e.t0=e.catch(0),console.log("Error in replaceTextRegex: "+e.t0);case 14:case"end":return e.stop()}}),e,null,[[0,11]])})));return function(t){return e.apply(this,arguments)}}());case 3:case"end":return e.stop()}}),e)}))),_replaceTextRegex.apply(this,arguments)}function tryCatch(e){return _tryCatch.apply(this,arguments)}function _tryCatch(){return(_tryCatch=_asyncToGenerator(_regeneratorRuntime().mark((function e(t){return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,t();case 3:console.log("Action completed."),e.next=9;break;case 6:e.prev=6,e.t0=e.catch(0),console.error(e.t0);case 9:case"end":return e.stop()}}),e,null,[[0,6]])})))).apply(this,arguments)}Office.onReady((function(e){e.host===Office.HostType.Word&&(initializeEventListeners(),populateOurFopSelect())}));