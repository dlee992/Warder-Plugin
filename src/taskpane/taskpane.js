/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/*eslint no-undef: "error"*/
/*eslint-env node*/


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
    
    document.getElementById("recreateTable").onclick = createTable;
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
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = currentWorksheet.getUsedRange()
    usedRange.delete("Left")
    //await context.sync()

    const expensesTable = currentWorksheet.tables.add("A1:E1", true /* hasHeaders */);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Amount", "Ratio", "Ratio2"]];
    expensesTable.rows.add(null /* add at the end */, [
      ["1/1/2017", "The Phone Company", "120", "240", "=D2 * 2"],
      ["1/2/2017", "Northwind Electric Cars", "142.33", "=C3 * 2", "=D3 * 2"],
      ["1/5/2017", "Best For You Organics Company", "27.9", "=C4 * 2", "=D4 * 2"],
      ["1/10/2017", "Coho Vineyard", "33", "=C5 * 2", "=D5 * 2"],
      ["1/11/2017", "Bellows College", "350.1", "=C6 * 2", "=D6 * 2"],
      ["1/15/2017", "Trey Research", "135", "270", "=D7 * 2"],
      ["1/15/2017", "Best For You Organics Company", "97.88", "=C8 * 2", "=D8 * 2"],
    ]);

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
  range.format.font.color = color;
  //range.format.fill.color = color;
}

var formulas = []
var numbers = []
var strings = []

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
            //highlightCell(cell, "red") // number cell
            numbers.push(cell)
          }
          else if (typeof formula === "string") {
            //highlightCell(cell, "purple") // string cell
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

    const {tokenize} = require("excel-formula-tokenizer")
    const {buildTree} = require("excel-formula-ast")
    const tree = buildTree(tokenize("= A1 * 1 + SUM(C1:F4)"))
    const {visitNode} = require("./firstStage/ast-visit")
    var ast = visitNode(tree)
    console.log(JSON.stringify(ast))

    await context.sync()
  }).catch(function(error) {
    console.log("Error: " + error)
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
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