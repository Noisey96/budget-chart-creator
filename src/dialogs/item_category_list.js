/*
    Name: John Freeman
    Date: 6/21/22
    File: item_category_list.js
    File History: Created on 6/20/22. Edited on 6/21/22 to finish its basic functionality.
*/

(function() {
    "use strict";
    Office.onReady()
    .then(function() {
        // when ready, function messages the taskpane to send the item categories, and it receives them
        Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, receiveItemCategoryList, registered);

        // adds event handlers to show the correct form
        document.getElementById("addForm").onclick = showForm;
        document.getElementById("editForm").onclick = showForm;
        document.getElementById("deactivateForm").onclick = showForm;

        // adds event handler to handle editing an item category
        document.getElementById("active-to-edit").onchange = selectItemCategory;

        // adds event handlers to send the updated item category list back to the taskpane
        document.getElementById("add").onclick = addItemCategory;
        document.getElementById("edit").onclick = editItemCategory;
        document.getElementById("deactivate").onclick = deactivateItemCategory;
    });

    // receives the item category list by populating select elements with the relevant item categories
    let itemCategoryList = null;
    function receiveItemCategoryList(arg) {
        itemCategoryList = JSON.parse(arg.message);

        // adds inactive item categories to its relevant select element
        let inactiveToAdd = document.getElementById("inactive-to-add");
        for (let itemCategory of itemCategoryList.inactive) {
            if (newOption(inactiveToAdd, itemCategory)) {
                let addOption = document.createElement("option");
                addOption.value = itemCategory;
                addOption.textContent = itemCategory;
                inactiveToAdd.appendChild(addOption);
            }
        }

        // adds active item categories to its relevant select elements
        let activeToEdit = document.getElementById("active-to-edit");
        let activeToDeactivate = document.getElementById("active-to-deactivate");
        for (let itemCategory of itemCategoryList.active) {
            if (newOption(activeToEdit,itemCategory)) {
                let editOption = document.createElement("option");
                editOption.value = itemCategory;
                editOption.textContent = itemCategory;
                activeToEdit.appendChild(editOption);
            }
            if (newOption(activeToDeactivate, itemCategory)) {
                let deactivateOption = document.createElement("option");
                deactivateOption.value = itemCategory;
                deactivateOption.textContent = itemCategory;
                activeToDeactivate.appendChild(deactivateOption);
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

    // shows the correct form based on the button's name and the form's ID
    function showForm(event) {
        let name = event.target.name;
        console.log(name);
        let forms = document.getElementById("option-forms");
        for (let child of forms.children) {
            child.className = "ms-Form";
        }
        let addForm = document.getElementById(name);
        addForm.className += " active";
    }

    // takes the selected item category and adds it to the editing area
    function selectItemCategory() {
        // identifies the selected item category
        let selectElem = document.getElementById("active-to-edit");
        let itemCategory = selectElem.options[selectElem.selectedIndex].value;

        // adds it to the editing area
        let inputElem = document.getElementById("editing-area");
        inputElem.value = itemCategory;
    }

    // potentially adds a new active item category and sends the new item category list
    function addItemCategory() {
        // finds the item category to add
        let selectElem = document.getElementById("inactive-to-add");
        let selectedItemCategory = selectElem.options[selectElem.selectedIndex].value; 
        let inputElem = document.getElementById("new");
        let inputtedItemCategory = inputElem.value;
        
        // if the user wants to add two item categories, then send them an error
        if (selectedItemCategory && inputtedItemCategory) {
            let errorMsgElem = document.getElementById("addError");
            errorMsgElem.textContent = "You have selected " + selectedItemCategory + " and inputted " + inputtedItemCategory + ". Please select or input one of them at a time.";
        }
        // if the user wants to add a new item category...
        else if (!selectedItemCategory && inputtedItemCategory) {
            // then we validate the new item category and either...
            let errorMsg = invalidItemCategory(inputtedItemCategory);
            // send them an error...
            if (errorMsg) {
                let errorMsgElem = document.getElementById("addError");
                errorMsgElem.textContent = errorMsg;
            } 
            // or send the new item category list to the taskpane
            else {
                itemCategoryList.active.push(inputtedItemCategory);
                itemCategoryList.active = sortItemCategories(itemCategoryList.active);
                Office.context.ui.messageParent(JSON.stringify(itemCategoryList));
            }
        }
        // if the user wants to activate an inactive item category...
        else if (selectedItemCategory && !inputtedItemCategory) {
            // then we remove the item category from the inactive list,...
            let inactiveIndex = itemCategoryList.inactive.indexOf(selectedItemCategory);
            itemCategoryList.inactive = itemCategoryList.inactive.slice(0, inactiveIndex).concat(itemCategoryList.inactive.slice(inactiveIndex + 1));

            // add the item category to the active list,...
            itemCategoryList.active.push(selectedItemCategory);
            itemCategoryList.active = sortItemCategories(itemCategoryList.active);

            // and send the item category list to the taskpane
            Office.context.ui.messageParent(JSON.stringify(itemCategoryList));
        }
    }

    // generates the new item category list of active and inactive categories and sends it to the taskpane
    function editItemCategory() {
        // find the old and new item categories
        let selectElem = document.getElementById("active-to-edit");
        let inputElem = document.getElementById("editing-area");
        let oldItemCategory = selectElem.options[selectElem.selectedIndex].value;
        let newItemCategory = inputElem.value;

        // remove the old item category from the list, so we can properly validate the new item category
        let oldIndex = itemCategoryList.active.indexOf(oldItemCategory);
        itemCategoryList.active = itemCategoryList.active.slice(0, oldIndex).concat(itemCategoryList.active.slice(oldIndex + 1));
        let errorMsg = invalidItemCategory(newItemCategory);
        
        // if the new item category is invalid, add the old item category back
        if (errorMsg) {
            itemCategoryList.active.push(oldItemCategory);
            itemCategoryList.active = sortItemCategories(itemCategoryList.active);

            let errorMsgElem = document.getElementById("editError");
            errorMsgElem.textContent = errorMsg;
        }
        // otherwise, add the new item category and send the list to the taskpane
        else {
            itemCategoryList.active.push(newItemCategory);
            itemCategoryList.active = sortItemCategories(itemCategoryList.active);
            Office.context.ui.messageParent(JSON.stringify(itemCategoryList));
        }
    }

    // generates the new item category list of active and inactive categories and sends it to the taskpane
    function deactivateItemCategory() {
        // find the item category to deactivate
        let selectElem = document.getElementById("active-to-deactivate");
        let itemCategory = selectElem.options[selectElem.selectedIndex].value;

        // deactivates the item category
        itemCategoryList.active = itemCategoryList.active.filter(e => e !== itemCategory);
        itemCategoryList.inactive.push(itemCategory);
        itemCategoryList.inactive = sortItemCategories(itemCategoryList.inactive);

        // send new item category list to taskpane
        Office.context.ui.messageParent(JSON.stringify(itemCategoryList));
    }

    // helper function to validate a new item category
    function invalidItemCategory(itemCategory) {
        if (/_/.test(itemCategory)) return itemCategory + " has underscores (i.e. _) in it. Please remove all underscores from the new item category.";
        if (!/^[a-zA-Z]/.test(itemCategory)) return itemCategory + " does not start with a letter. All new item categories must start with a letter.";

        let strippedItemCategory = removeInvalidCharacters(itemCategory);
        let activeList = itemCategoryList.active;
        for (let category of activeList) {
            let strippedCategory = removeInvalidCharacters(category);
            if (strippedCategory === strippedItemCategory) return itemCategory + " is too similar to " + category + ".";
        }
        let inactiveList = itemCategoryList.inactive;
        for (let category of inactiveList) {
            let strippedCategory = removeInvalidCharacters(category);
            if (strippedCategory === strippedItemCategory) return itemCategory + " is too similar to " + category + ".";
        }
        return null;
    }

    // helper function to sort item categories
    function sortItemCategories(list) {
        return list.sort((a, b) => {
            // convert array elements to uppercase
            a = a.toUpperCase();
            b = b.toUpperCase();

            // always sort Other as last
            if (a === "OTHER") return 1;
            else if (b === "OTHER") return -1;

            // otherwise, sort alphabetically
            else if (a < b) return -1;
            else if (a > b) return 1;
            return 0;
        });
    }

    // helper function to remove invalid characters (for table names)
    function removeInvalidCharacters(name) {
        return name.replace(/[^A-Za-z0-9]/g, "");
    }
}());