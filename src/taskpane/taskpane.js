/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office, OfficeExtension */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    //const fs = require("fs");
    //var access = fs.createWriteStream("C:Users\\ocaml\\Codes\\Warder-Plugin\\.log");
    //process.stdout.write = process.stderr.write = access.write.bind(access);

    createTable();
    document.getElementById("create-table").onclick = createTable;
    document.getElementById("warder-analysis").onclick = warderAnalysis;

    document.getElementById("sideloading-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function createTable() {
  Excel.run(function(context) {
    console.log("create Table");
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:E1", true /* hasHeaders */);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount", "Ratio"]];
    expensesTable.rows.add(null /* add at the end */, [
      ["1/1/2017", "The Phone Company", "Communications", "120", "=D2 * 2"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33", "=D3 * 2"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9", "=D4 * 2"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33", "=D5 * 2"],
      ["1/11/2017", "Bellows College", "Education", "350.1", "=D6 * 2"],
      ["1/15/2017", "Trey Research", "Other", "135", "=D7 * 2"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88", "=D8 * 2"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88", "=D9 * 2"],
    ]);

    expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();

    return context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function warderAnalysis() {
  Excel.run(
    async function(context) {
      //why console.log cannot be seen, it must be somewhere???
      console.log("warder analysis");

      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var usedRange = currentWorksheet.getUsedRange();
      usedRange.load();
      await context.sync();

      //first stage
      var foundRange = usedRange.find("=", {
        completeMatch: false,
        matchCase: false,
        searchDirection: Excel.SearchDirection.forward
      });
      foundRange.load();
      await context.sync();
      foundRange.format.fill.color = "blue";
      foundRange.format.font.color = "white";

      //second stage

      //report

      //end

      return context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
