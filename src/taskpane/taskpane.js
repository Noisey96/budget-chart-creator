/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.
*/
/*
  Name: John Freeman
  Date: 6/21/22
  File: taskpane.js
  File History: Created on 6/13/22. Edited on 6/17/22 to add in functionality to create the charts. Edited on 6/18/22 to add in functionality to format the chart, refresh the  
                data, and position the charts. Edited on 6/19/22 to decrease complexity of the code. Edited on 6/20/22 and 6/21/22 to add in functionality to update limits, 
                handle null limits, handle invalid characters, and handle add/removing item categories.
*/

console.log("test");

// once ready, link the HTML elements to the JavaScript
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("select-month").onchange = selectMonth;
    document.getElementById("add-update-chart").onclick = openChartDialog;
    document.getElementById("refresh-charts").onclick = refreshCharts;
    document.getElementById("edit-item-category-list").onclick = openItemCategoryListDialog;
  }
});

// global variables used to store data needed across different functions
let month, categoryColumn, amountColumn, error, charts = 0, itemCategoryList = {
  active: ["Charity & Gifts", 
           "Cleaning Products & Services", 
           "Clothing",
           "Electronics & Software",
           "Entertainment",
           "Food",
           "Furniture & Home Accessories",
           "Housing",
           "Insurance & Loan Payments",
           "Kitchen & Dining",
           "Medical Products & Services",
           "Work Supplies",
           "Other"],
  inactive: [],
}

// function to load in the expense data from the specified month
async function selectMonth() {
  await Excel.run(async (context) => {
    // translate specified month into its worksheet name
    let selectElem = document.getElementById("select-month");
    month = selectElem.options[selectElem.selectedIndex].value;
    
    // refresh each chart's name and title using a for-loop to identify each chart
    let chartSheet = context.workbook.worksheets.getActiveWorksheet();
    chartSheet.charts.load("count");
    await context.sync();
    let chartCount = chartSheet.charts.count;
    for (let i = 0; i < chartCount; i++) {
      let chart = chartSheet.charts.getItemAt(i);
      chart.load("name");
      let chartTitle = chart.title.load("text");
      await context.sync();

      // updates each chart's name and title
      let chartName = chart.name.split("_");
      chartName[1] = month;
      chart.name = chartName.join("_");
      let chartTitleText = chartTitle.text.split(" ");
      chartTitleText[chartTitleText.length - 1] = month;
      chartTitle.text = chartTitleText.join(" ");
      await context.sync();
    }

    // refresh each table and the transaction dependent properties of its relevant chart
    await refreshCharts();
  })
  .catch(handleErrors);
}

// pulls relevant transaction data and opens a dialog to add/update a chart
let chartDialog = null;
async function openChartDialog() {
  // pull in relevant transaction data to update item category list before the dialog
  await Excel.run(async (context) => {
    await findTransactionData(context);
    await context.sync();
  })
  .catch(handleErrors);

  // opens the dialog to add/update a chart
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/chart.html',
    {height: 45, width: 55},
    function (result) {
      chartDialog = result.value;
      chartDialog.addEventHandler(Office.EventType.DialogMessageReceived, addUpdateChart);
    }
  );
}

// function to add/update a chart to the active worksheet
async function addUpdateChart(arg) {
  if (arg.message === "connected") chartDialog.messageChild(JSON.stringify(itemCategoryList));
  else if (arg.message === "failed") {
    console.log("Connection to item category list dialog failed!")
    chartDialog.close();
    return null;
  } else {
    // receives the item category and limit from the dialog
    let fullCategory = arg.message.split("_")[0];
    let category = removeInvalidCharacters(fullCategory);
    let limit = Number(arg.message.split("_")[1]);
    chartDialog.close();

    await Excel.run(async (context) => {
      // calculate amount spent and unspent, and whether the user overspent in that item category
      let [spent, unspent, overspent] = calculateAmounts(category, limit);

      // if this is the first chart, intializes the chart sheet and the table sheet behind it
      let chartSheet = context.workbook.worksheets.getActiveWorksheet(), tableSheet;
      if (!charts) {
        // initializes the chart sheet
        chartSheet.charts.onAdded.add(repositionCharts);
        chartSheet.charts.onDeleted.add(deletedChart);

        // initializes the table sheet behind the chart sheet
        tableSheet = context.workbook.worksheets.add("Tables");
        chartSheet.load("position");
        tableSheet.load("position");
        await context.sync();
        tableSheet.position = chartSheet.position + 1;
        tableSheet.visibility = Excel.SheetVisibility.hidden;
        await context.sync();
      }

      // add/update the underlying table and its chart
      tableSheet = context.workbook.worksheets.getItem("Tables");
      let table = tableSheet.tables.getItemOrNullObject(category + "_" + month);
      await context.sync();
      let found = !table.isNullObject;
      // if relevant table is found, update the table and its chart
      if (found) {
        // updates the table
        table.getDataBodyRange().values = [[spent, unspent]];
        tableSheet.getUsedRange().format.autofitColumns();
        tableSheet.getUsedRange().format.autofitRows();
        await context.sync();

        // updates its chart
        let chart = chartSheet.charts.getItem(category + "_" + month);
        await dependentChartFormat(context, chart, limit, spent, overspent);
        await context.sync();
      }
      // otherwise, create the table and its chart
      else {
        // create the table
        tableSheet = context.workbook.worksheets.getItem("Tables");
        let tableRange = tableSheet.getRange("A" + ((charts * 3) + 1) + ":B" + ((charts * 3) + 2));
        let nextTable = tableSheet.tables.add(tableRange, true);
        nextTable.name = category + "_" + month;
        nextTable.getHeaderRowRange().values = [["Spent", "Unspent"]];
        nextTable.getDataBodyRange().values = [[spent, unspent]];
        tableSheet.getUsedRange().format.autofitColumns();
        tableSheet.getUsedRange().format.autofitRows();
        await context.sync();

        // create its chart and format it
        let chart = chartSheet.charts.add("BarStacked", tableRange, "Columns");
        await independentChartFormat(context, chart, fullCategory, category);
        await dependentChartFormat(context, chart, limit, spent, overspent);
        await context.sync();
        charts++;
      }
    })
    .catch(handleErrors);
  }
}

// function to refresh each table and the transaction dependent properties of its relevant chart
async function refreshCharts() {
  await Excel.run(async (context) => {
    // refresh all the tables
    let tableSheet = context.workbook.worksheets.getItemOrNullObject("Tables");
    await context.sync();
    if (!tableSheet.isNullObject) {
      // if there are pre-existing tables, refresh relevant transaction data behind all the tables
      await findTransactionData(context);

      // use a for-loop to find each table
      tableSheet.tables.load("count");
      await context.sync();
      let tableCount = tableSheet.tables.count;
      for (let i = 0; i < tableCount; i++) {
        let table = tableSheet.tables.getItemAt(i);

        // identify each table's category
        table.load("name");
        await context.sync();
        let category = table.name.split("_")[0];

        // calculate limit, amount spent, amount unspent, and whether the user overspent in that item category
        let tableData = table.getDataBodyRange().load("values");
        await context.sync();
        let amounts = tableData.values;
        let limit = amounts[0][0] + amounts[0][1];
        let [spent, unspent, overspent] = calculateAmounts(category, limit);

        // update table
        tableData.values = [[spent, unspent]];
        await context.sync();

        // refresh the relevant chart
        let chartSheet = context.workbook.worksheets.getActiveWorksheet();
        let chart = chartSheet.charts.getItem(category + "_" + month);
        await dependentChartFormat(context, chart, limit, spent, overspent);
      }
    }
  })
  .catch(handleErrors);
}

// function to reposition charts after one is added or deleted
async function repositionCharts() {
  await Excel.run(async (context) => {
    let chartSheet = context.workbook.worksheets.getActiveWorksheet();
    chartSheet.charts.load("items");
    await context.sync();

    let allCharts = chartSheet.charts.items;
    for (let i = 0; i < allCharts.length; i++) {
      let chart = allCharts[i];
      let row = Math.floor(i / 3);
      let column = i  % 3;
      chart.left = 10 + (380 * column);
      chart.top = 10 + (175 * row);
    }
  })
  .catch(handleErrors);
}

// helper function to find the relevant transaction data and adds new valid item categories to the item category list
async function findTransactionData(context) {
  // generate the name to the worksheet with the relevant transaction data
  let year = new Date().getFullYear();
  let worksheetName = month + " " + year;

  // finds the relevant transaction data
  let sheet = context.workbook.worksheets.getItem(worksheetName);
  let expensesTable = sheet.tables.getItem(month);
  categoryColumn = expensesTable.columns.getItem("Item Category").getDataBodyRange().load("values");
  amountColumn = expensesTable.columns.getItem("Cost").getDataBodyRange().load("values");
  await context.sync();

  // using the relevant transaction data, finds and adds new valid item categories to the item category list
  let newItemCategories = [...new Set(categoryColumn.values.map(e => e[0]))].filter(e => e !== "");
  for (let i = 0; i < newItemCategories.length; i++) {
    let newItemCategory = newItemCategories[i];
    let invalid = invalidItemCategory(newItemCategory);
    if (!invalid) {
      let active = itemCategoryList.active.includes(newItemCategory);
      let inactive = itemCategoryList.inactive.includes(newItemCategory);
      if (!active && !inactive) {
        itemCategoryList.active.push(newItemCategory);
        itemCategoryList.active = sortItemCategories(itemCategoryList.active);
      }
    }
  }
}

// helper function to calculate amount spent, unspent, and overspent in that item category
function calculateAmounts(category, limit) {
  // calculate amounts
  let spent = amountColumn.values.reduce((total, _, i) => {
    if (removeInvalidCharacters(categoryColumn.values[i][0]) === category) total += amountColumn.values[i][0];
    return total;
  }, 0);
  let unspent = limit - spent, overspent = false;
  if (unspent < 0) overspent = true;

  // reformats amounts
  spent = spent.toFixed(2);
  unspent = unspent.toFixed(2);

  return [spent, unspent, overspent];
}

// helper function to handle errors
let errorDialog = null;
function handleErrors(err) {
  error = err;
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/error.html',
    {height: 45, width: 55},
    function (result) {
      errorDialog = result.value;
      errorDialog.addEventHandler(Office.EventType.DialogMessageReceived, sendError);
    }
  );
}

// helper function to send the error to the error dialog
function sendError(arg) {
  if (arg.message === "connected") errorDialog.messageChild(JSON.stringify(error));
}

// helper function to format the chart's transaction independent properties
async function independentChartFormat(context, chart, fullCategory, category) {
  // format the chrt's size
  chart.height = 155;
  chart.width = 360;

  // give the chart a name
  chart.name = category + "_" + month;

  // format the chart's title
  chart.title.text = fullCategory + " Spending in " + month;
  chart.title.format.font.size = 16;
  chart.title.format.font.color = "#000000";

  // format the chart's legend
  chart.legend.position = "bottom";
  chart.legend.format.font.size = 10;
  chart.legend.format.font.color = "#000000";

  // format the chart's axes
  chart.axes.valueAxis.format.font.size = 10;
  chart.axes.valueAxis.format.font.color = "#000000";
  chart.axes.valueAxis.numberFormat = "$0.00";
  chart.axes.valueAxis.minimum = 0;
  chart.axes.categoryAxis.visible = false;

  // format the chart's data series
  chart.series.load("items");
  await context.sync();
  for (let i = 0; i < chart.series.items.length; i++) {
    let bar = chart.series.items[i];
    // format bar height
    bar.gapWidth = 0;
    // format bar data labels
    bar.dataLabels.showValue = true;
    bar.dataLabels.numberFormat = "$0.00";
    bar.dataLabels.format.font.size = 12;
    bar.dataLabels.format.font.color = "#000000";
  }
  await context.sync();
}

// helper function to format the chart's transaction dependent properties
async function dependentChartFormat(context, chart, limit, spent, overspent) {
  // format the chart's axes
  chart.axes.valueAxis.maximum = overspent ? Math.ceil(spent) : limit;
  chart.axes.valueAxis.majorUnit = limit || Math.ceil(spent);

  // format the chart's data series's colors
  chart.series.load("items", "items/name");
  await context.sync();
  for (let i = 0; i < chart.series.items.length; i++) {
    let bar = chart.series.items[i];
    if (overspent && bar.name === "Spent") bar.format.fill.setSolidColor("#FF0000");
    else if (bar.name === "Spent") bar.format.fill.setSolidColor("#0070C0");
    else bar.format.fill.setSolidColor("#00B050");
  }
  await context.sync();
}

// helper function to remove invalid characters (for table names)
function removeInvalidCharacters(name) {
  return name.replace(/[^A-Za-z0-9]/g, "");
}

// function to open the dialog for editing the item category list
let itemCategoryListDialog = null;
function openItemCategoryListDialog() {
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/item_category_list.html',
    {height: 45, width: 55},
    function (result) {
      itemCategoryListDialog = result.value;
      itemCategoryListDialog.addEventHandler(Office.EventType.DialogMessageReceived, sendAndReceiveItemCategoryList);
    }
  );
}

// function to send and receive item category lists to and from its dialog
function sendAndReceiveItemCategoryList(arg) {
  if (arg.message === "connected") itemCategoryListDialog.messageChild(JSON.stringify(itemCategoryList));
  else if (arg.message === "failed") {
    console.log("Connection to item category list dialog failed!")
    itemCategoryListDialog.close();
    return null;
  } else {
    itemCategoryList = JSON.parse(arg.message);
    itemCategoryListDialog.close();
  }
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

// when a chart is deleted, repositions other charts, and finds and deletes tables with no associated charts
async function deletedChart() {
  await repositionCharts();

  await Excel.run(async (context) => {
    // run through each table's name...
    let tableSheet = context.workbook.worksheets.getItemOrNullObject("Tables");
    await context.sync();
    if (!tableSheet.isNullObject) {
      tableSheet.tables.load("count");
      await context.sync();
      let tableCount = tableSheet.tables.count, pending = [];
      for (let i = 0; i < tableCount; i++) {
        let table = tableSheet.tables.getItemAt(i);
        table.load("name");
        await context.sync();
        let name = table.name;
  
        // and find whether there is an associated chart 
        let chartSheet = context.workbook.worksheets.getActiveWorksheet();
        let chart = chartSheet.charts.getItemOrNullObject(name);
        await context.sync();
        // if so, continue
        if (!chart.isNullObject) continue;
        // if not, add the table to the pending array
        else {
          pending.push(table);
        }
      }
      
      // afterwards, delete each table in the pending array
      for (let table of pending) {
        table.delete();
      }
      await context.sync();
    }
  })
  .catch(handleErrors);
}