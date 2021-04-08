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
      ["1/1/2017", "The Phone Company", "120", "240", "= SUM(C2:D2)"],
      ["1/2/2017", "Northwind Electric Cars", "142.33", "=C3 * 2", "= SUM(C3:D3)"],
      ["1/5/2017", "Best For You Organics Company", "27.9", "=C4 * 2", "= SUM(C4:D4)"],
      ["1/10/2017", "Coho Vineyard", "33", "=C5 * 2", "= SUM(C5:D5)"],
      ["1/11/2017", "Bellows College", "350.1", "=C6 * 2", "= SUM(C6:D6)"],
      ["1/15/2017", "Trey Research", "135", "270", "= SUM(C7:D7)"],
      //["1/15/2017", "Best For You Organics Company", "97.88", "=C8 * 2", "= SUM(C8:D8)"],
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
  constructor(excel_cell, syntax_tree, astString, index) {
    this.excel_cell = excel_cell
    this.syntax_tree = syntax_tree
    this.astString = astString
    this.index = index
  }
}

class FirstCluster {
  constructor(myCell) {
    this.firstCell = myCell
    this.myCellSet = new Set()
    this.myCellSet.add(myCell)
  }
}

var colors = ["yellow", "blue", "red", "green", "grey", "orange"]

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
          highlightCell(cell, "red")
          //console.log("--- find a formula cell ---") 
          const syntax_tree = buildTree(tokenize(formula))
          
          //console.log("--- build ast tree ---") 
          const astString = buildAstTree(syntax_tree)
          //console.log(JSON.stringify(ast))
          
          //console.log("--- get tree size --- " + astSize(syntax_tree))
        
          formulas.push(new MyCell(cell, syntax_tree, astString, formulas.length))

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

    var ed = require('edit-distance')
    // Define cost functions.
    var insert, remove, update
    insert = remove = function(node) { return 1; }
    update = function(nodeA, nodeB) { return nodeA.id !== nodeB.id ? 1 : 0; }
    var children = function(node) { return node.children; }

    var simMatrix = new Array(formulas.length)
    for (let j = 0; j < formulas.length; j++) {
      simMatrix[j] = new Array(formulas.length)
    }

    for (let i = 0; i < formulas.length; i++) {
      var cell_i = formulas[i]
      
      for (let j = i+1; j < formulas.length; j++) {
        var cell_j = formulas[j]
        var ted = ed.ted(cell_i.astString, cell_j.astString, children, insert, remove, update)
        const sim = 1 - (ted.distance / (astSize(cell_i.syntax_tree) + astSize(cell_j.syntax_tree)))
        simMatrix[i][j] = simMatrix[j][i] = sim
        //console.log("--- ted --- ", cell_i.excel_cell.address, cell_j.excel_cell.address, sim)
      }
    }

    console.log("  --- HAClustring: start ---")
    /**
     * todo: nearest neighbor chain algorithm
     * Set of active clusters, one for each input point
     * Stack, empty
     * while (halting condition): 
     *    if Stack empty, push any cluster into Stack
     *    C = top of Stack
     *    exist D = nearest other cluster of C
     *    if push D into S
     *    otherwise, D in stack, must be father of C, pop C and D, merge them
     **/
    var firstClusterSet = new Set()
    var stack = []
    for (let index = 0; index < formulas.length; index++) {
      const myCell = formulas[index]
      var firstCluster = new FirstCluster(myCell)
      firstClusterSet.add(firstCluster)
      if (index == 0) 
        stack.push(firstCluster)  
    }

    while (firstClusterSet.size > 1) {
      if (stack.length == 0) stack.push(firstClusterSet.values().next().value)

      var currentCluster = stack.pop()
      var minimumDistance = Number.MAX_SAFE_INTEGER
      var pairCluster = undefined
      for (const cluster of firstClusterSet) {
        if (cluster === currentCluster) continue
        var dis = 0
        for (const cellInCluster of cluster.myCellSet) {
          for (const cellInCurrentCluster of currentCluster.myCellSet) {
            dis += (1 - simMatrix[cellInCluster.index][cellInCurrentCluster.index])
          }
        }
        dis = dis / (cluster.myCellSet.size * currentCluster.myCellSet.size)
        if (dis < minimumDistance) {
          minimumDistance = dis
          pairCluster = cluster
          console.log("minimumDistance:", dis)
        }
      }

      if (stack.indexOf(pairCluster) == -1) {
        stack.push(currentCluster)
        stack.push(pairCluster)
      } else {
        if (minimumDistance > 0.02) break 
        stack.pop()
        //debugging
        for (const myCell of currentCluster.myCellSet) {
          console.log("merged cells in currentCluster:", myCell.excel_cell.address)
        }
        for (const myCell of pairCluster.myCellSet) {
          console.log("merging cells in pairCluster:", myCell.excel_cell.address)
          currentCluster.myCellSet.add(myCell)
        }
        firstClusterSet.delete(pairCluster)
      }
    }

    var index = 0
    for (const cluster of firstClusterSet) {
      for (const myCell of cluster.myCellSet) {
        highlightCell(myCell.excel_cell, colors[index % colors.length])
      }
      index++
    }

    console.log("  --- HAClustring: end ---")

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