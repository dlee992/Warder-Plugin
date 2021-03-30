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

    return context.sync().then(console.log("----- Create Table : done -----"));
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

function preprocess() {
  Excel.run(
    async function(context) {
    await context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

class MyCell {
  constructor(excel_cell, syntax_tree, astString) {
    this.excel_cell = excel_cell
    this.syntax_tree = syntax_tree
    this.astString = astString
  }

}


function firststage() {
  Excel.run(async function(context) {
    console.log("----- first stage : start -----")

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

    const {tokenize} = require("excel-formula-tokenizer")
    const {buildTree} = require("excel-formula-ast")
    const {buildAstTree, buildCdtTree, astSize} = require("./firstStage/tree-visit")

    for (let rowIndex = usedRange.rowIndex; rowIndex <= lastCell.rowIndex; rowIndex++) {
      for (let colIndex = usedRange.columnIndex; colIndex <= lastCell.columnIndex; colIndex++) {
        var cell = worksheet.getCell(rowIndex, colIndex)
        cell.load()
        
        await context.sync()
        var formula = cell.formulasR1C1[0][0]
        if (typeof formula === "string" && formula.indexOf('=') == 0) {
          //console.log("--- find a formula cell ---") 
          const syntax_tree = buildTree(tokenize(formula))
          
          //console.log("--- build ast tree ---") 
          const astString = buildAstTree(syntax_tree)
          //console.log(JSON.stringify(ast))
          
          //console.log("--- get tree size --- " + astSize(syntax_tree))
          
          formulas.push(new MyCell(cell, syntax_tree, astString))

          console.log("--- push a formula ---")
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

    for (let index = 0; index < formulas.length; index++) {
      var formula_cell_2 = formulas[index]
      // get cdt tree somehow 
    }

    var simMatrix = []
    for (let index_i = 0; index_i < formulas.length; index_i++) {
      var formula_cell_1 = formulas[index_i]
      simMatrix[index_i] = []
      for (let index_j = 0; index_j < formulas.length; index_j++) {
        if (index_i == index_j) continue
        var formula_cell_2 = formulas[index_j]
        //
        var red = 0
        simMatrix[index_i][index_j] = 1 - red/(astSize(formula_cell_1.syntax_tree) + astSize(formula_cell_2.syntax_tree))
      }
    }
    
    console.log("--- HAClustring: start ---")



    console.log("--- HAClustring: end ---")

    await context.sync().then(console.log("----- first stage : end -----"))
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