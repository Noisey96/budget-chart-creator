!function(){"use strict";function e(e){console.log(e.message);var t=JSON.parse(e.message),n=document.getElementById("errorName"),c=document.getElementById("errorInfo");n.textContent=t.code,c.textContent=JSON.stringify(t.debugInfo)}function t(e){e.status===Office.AsyncResultStatus.Succeeded?Office.context.ui.messageParent("connected"):Office.context.ui.messageParent("failed")}Office.onReady().then((function(){Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived,e,t)}))}();