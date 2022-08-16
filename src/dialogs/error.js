/*
    Name: John Freeman
    Date: 6/20/22
    File: error.js
    File History: Created on 6/19/22. Edited on 6/20/22 to display more error information.
*/

// when ready, function messages the taskpane to send the error, and it receives and displays that error
(function() {
    "use strict";
    Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, displayError, registered);
    });

    // function displays the error
    function displayError(arg) {
        console.log(arg.message);
        let error = JSON.parse(arg.message);

        let errorNameElem = document.getElementById("errorName");
        let errorInfoElem = document.getElementById("errorInfo");
        errorNameElem.textContent = error.code;
        errorInfoElem.textContent = JSON.stringify(error.debugInfo);
    }

    // function sends a message to the taskpane when ready
    function registered(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) Office.context.ui.messageParent("connected");
        else Office.context.ui.messageParent("failed");
    }
}());