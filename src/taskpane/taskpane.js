/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/*eslint no-undef: "error"*/
/*eslint-env node*/

import { tokenize } from "excel-formula-tokenizer"
import { buildTree, visit } from "excel-formula-ast"

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Excel, Office, OfficeExtension */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    createTable();
    document.getElementById("preprocess").onclick = preprocess;
    document.getElementById("firststage").onclick = firststage;
    document.getElementById("secondstage").onclick = secondstage;
    document.getElementById("detection").onclick = detection;
    document.getElementById("postprocess").onclick = postprocess;

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

function highlightCell(range, color) {
  //range.format.font.color = "white";
  range.format.fill.color = color;
}

function preprocess() {
  Excel.run(
    async function(context) {

      /* first step: filter out all kinds of cells, get all formula cells, and cluster them somehow
      */
      var worksheet = context.workbook.worksheets.getActiveWorksheet()
      var usedRange = worksheet.getUsedRange()
      usedRange.load()
      var lastCell = usedRange.getLastCell()
      lastCell.load() 
      await context.sync()

      var formulas = []
      var numbers = []
      var strings = []
      for (let rowIndex = usedRange.rowIndex; rowIndex <= lastCell.rowIndex; rowIndex++) {
        for (let colIndex = usedRange.columnIndex; colIndex <= lastCell.columnIndex; colIndex++) {
          var cell = worksheet.getCell(rowIndex, colIndex)
          cell.load()
          await context.sync()
          var formula = cell.formulas[0][0]
          if (typeof formula === "string" && formula.indexOf('=') == 0) {
            highlightCell(cell, "blue") // formula cell 
            formulas.push(cell)
          }
          else if (typeof formula === "number") {
            highlightCell(cell, "red") // number cell
            numbers.push(cell)
          }
          else if (typeof formula === "string") {
            highlightCell(cell, "purple") // string cell
            strings.push(cell)
          }
          else {
            // could be error cell, or anything else
          }
        }
      }

      await context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function firststage() {
  Excel.run(async function(context) {


    await context.sync()
  }).catch(function(error) {
    console.log("Error: " + error)
  })
}

function secondstage() {
  Excel.run(async function(context) {


    await context.sync()
  }).catch(function(error) {
    console.log("Error: " + error)
  })
}

function detection() {
  Excel.run(async function(context) {


    await context.sync()
  }).catch(function(error) {
    console.log("Error: " + error)
  })
}

function postprocess() {
  Excel.run(async function(context) {


    await context.sync()
  }).catch(function(error) {
    console.log("Error: " + error)
  })
}