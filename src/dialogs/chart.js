/*
    Name: John Freeman
    Date: 6/21/22
    File: chart.js
    File History: Created on 6/14/22. Edited on 6/17/22 to rewrite comments. Edited on 6/20/22 and 6/21/22 to add functionality to update item categories and handle null limits.
*/

// function finds inputted category and limit in the dialog and sends them back to the taskpane
(function() {
    "use strict";
    Office.onReady()
    .then(function() {
        // when ready, function messages the taskpane to send the item categories, and it receives them
        Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, receiveItemCategoryList, registered);
        
        // adds event handler to complete the dialog
        document.getElementById("ok-button").onclick = sendSelectedToParentPage;
    });

    // receives the item category list by populating select elements with the relevant item categories
    let itemCategoryList = null;
    function receiveItemCategoryList(arg) {
        itemCategoryList = JSON.parse(arg.message);

        // adds inactive item categories to its relevant select element
        let selectElem = document.getElementById("select-item-category");
        for (let itemCategory of itemCategoryList.active) {
            if (newOption(selectElem, itemCategory)) {
                let option = document.createElement("option");
                option.value = itemCategory;
                option.textContent = itemCategory;
                selectElem.appendChild(option);
            }
        }
    }

    // helper function to identify new options being added to select elements
    function newOption(selectElem, value) {
        for (let child of selectElem.children) {
            if (child.value === value) return false;
        }
        return true;
    }

    // sends a message to the taskpane when ready
    function registered(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) Office.context.ui.messageParent("connected");
        else Office.context.ui.messageParent("failed");
    }

    // this function performs the bulk of the work
    function sendSelectedToParentPage() {
        let selectElem = document.getElementById("select-item-category");
        let category = selectElem.options[selectElem.selectedIndex].value;
        let limit = document.getElementById("input-limit").value || 0;
        let message = category + "_" + limit;
        Office.context.ui.messageParent(message);
    }
}());
