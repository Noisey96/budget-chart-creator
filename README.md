# Budget Chart Creator

This project serves as a Microsoft Excel Add-in that creates charts to help with budgeting. The charts created by this project all look similar. Each chart will be a bar graph where one can evaluate their previous transaction data in a specified item category against a specified monetary limit. This kind of chart is to help avoid overspending in a particular item category for the current month.

This project is still in development. This project has an incomplete UI, makes too many assumptions, and is prone to errors.

## What?

This project contains an assets folder, a src folder, a manifest.xml file, a webpack.config file, and some other files. This project was started with [these directions](https://docs.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator).

The assets folder contains the images for the project. At the moment, it has the Microsoft-provided images. The src folder contains the core HTML, CSS, and JS files. The manifest.xml file defines how the Add-In is added to the Excel toolbar. The webpack.config file bundles the src code.

## How?

1. If needed, add your transaction data into Microsoft Excel.
2. Format your transaction data. At the moment, your transaction data needs to be split by month across multiple worksheets where each month's transaction data needs to be in a separate table on a separate worksheet. Each worksheet needs a title similar to "January 2022" and each table needs a title similar to "January". Finally, each table needs to have an "Item Category" column and a "Cost" column.
3. Clone or download the code.
4. Run "npm install" to install the dependencies.
5. Run "npm start" to start the application.
6. Select a month and you are all set to create a chart!

## Why?

As I mentioned earlier, this project creates a kind of chart to help avoid overspending in a particular item category for the current month. Originally, I wanted to add this kind of chart to my already-existing Excel spreadsheet for my budget. At the same time, I discovered Microsoft Add-Ins and the online resources Microsoft provides to assist with developing Microsoft Add-Ins. Therefore, I started this project to fill my personal need and improve my coding skills!
