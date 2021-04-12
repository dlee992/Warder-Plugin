export class FeatureExtraction {
  constructor(sheetWrapper, rowBase, columnBase, finalFirstClusters, formulaCells, numberCells, stringCells) {
    this.sheetWrapper = sheetWrapper
    this.rowBase = rowBase
    this.columnBase = columnBase
    this.finalFirstClusters = finalFirstClusters
    this.formulaCells = formulaCells
    this.numberCells = numberCells
    this.stringCells = stringCells

    this.tableSet = new Set()

    this.extractCellAddress()

    this.extractTable()
    for (const table of this.tableSet) console.log(table.toString())

    this.extractLabel()

    //this.extractAlliance() // wait for second development

    this.extractCellArray()

    //this.extractGapTemplate() // wait for second development
  }

  /**
   * f1: row index, column index
   * f2?.: a cell can reference many cells, whether referenced cells are in the same row or column with the cell
   */
  extractCellAddress() {
    // nothing to do, implement in class CellWrapper
    console.log("  --- extract cell address: start ---")
    console.log("  --- extract cell address: end ---")
  }

  /**
   * label: is header
   */
  extractLabel() {
    console.log("  --- extract label: start ---")

    for (let rowIndex = 0; rowIndex < this.sheetWrapper.length; rowIndex++) {
      for (let colIndex = 0; colIndex < this.sheetWrapper[rowIndex].length; colIndex++) {
        var cellWrapper = this.sheetWrapper[rowIndex][colIndex]
        if (cellWrapper.cellType == "empty" || cellWrapper.cellType == "string") continue
        //console.log(rowIndex, colIndex)

        //find left header in the same row
        var minus = 1
        var labelRow = ""
        var copyRow = false //rewrite: "var copyRow". First time, copyRow is undefined, second time, it is defined!!! 
        while (colIndex >= minus) {
          var headerCellWrapper = this.sheetWrapper[rowIndex][colIndex - minus]
          if (headerCellWrapper.cellType == "empty") {
            if (labelRow !== "") break
            minus++
            continue
          }
          if (headerCellWrapper.cellType == "formula" || headerCellWrapper.cellType == "number") {
            if (labelRow !== "") break
            //console.log("copy", headerCellWrapper.ft_labelRow)
            cellWrapper.ft_labelRow = headerCellWrapper.ft_labelRow
            copyRow = true
            break
          }
          //console.log("go 1")
          if (headerCellWrapper.cellType == "string") {
            //console.log("string", headerCellWrapper.excel_cell.formulas[0][0])
            labelRow = headerCellWrapper.excel_cell.formulas[0][0] + "&" + labelRow
          }
          minus ++
        }
        
        if (copyRow == false) 
          cellWrapper.ft_labelRow = labelRow
        //console.log("go 2")
        //find up header in the same column, same structure
        minus = 1
        var labelColumn = ""
        var copyColumn = false
        while (rowIndex >= minus) {
          var headerCellWrapper = this.sheetWrapper[rowIndex - minus][colIndex]
          if (headerCellWrapper.cellType == "empty") {
            if (labelColumn !== "") break
            minus++
            continue
          }
          if (headerCellWrapper.cellType == "formula" || headerCellWrapper.cellType == "number") {
            if (labelColumn !== "") break
            cellWrapper.ft_labelColumn = headerCellWrapper.ft_labelColumn
            copyColumn = true
            break
          }
          //console.log("go 4")
          if (headerCellWrapper.cellType == "string") {
            labelColumn = headerCellWrapper.excel_cell.formulas[0][0] + "&" + labelColumn
          }
          minus ++
        }
        
        if (copyColumn == false) 
          cellWrapper.ft_labelColumn = labelColumn
        //console.log("go 5")
        console.log(cellWrapper.excel_cell.address, cellWrapper.ft_labelRow, cellWrapper.ft_labelColumn)
      }
    }

    console.log("  --- extract label: end ---")
  }

  /**
   * example: SUM(A2:E2), then A2,B2...E2 are alliances
   */
  extractAlliance() {

  }

  /**
   * table: means a contiguous cell range, in which, no empty row or column
   */
  extractTable() {
    /**
     * breadth-first search
     * a bug: potential bug
     * ******--*
     * *-------*
     * *--******
     * but rarely happen, for now, don't handle this case
     */
    console.log("  --- extract table: start ---")
    var visitedCells = new Array(this.sheetWrapper.length)
    for (let rowIndex = 0; rowIndex < this.sheetWrapper.length; rowIndex++) {
      var row = this.sheetWrapper[rowIndex]
      visitedCells[rowIndex] = new Array(row.length)
    }

    var moveRow = [1, 0, -1, 0]
    var moveCol = [0, 1, 0, -1]

    for (let rowIndex = 0; rowIndex < this.sheetWrapper.length; rowIndex++) {
      for (let colIndex = 0; colIndex < this.sheetWrapper[rowIndex].length; colIndex++) {
        var cellWrapper = this.sheetWrapper[rowIndex][colIndex]

        //two sub-conditions can be combined with cellWrapper?.cellType == "string" ???
        if (cellWrapper.cellType == "empty" || cellWrapper.cellType == "string" || visitedCells[rowIndex][colIndex] == true) continue
        //console.log(rowIndex, colIndex)
        
        
        var table = new Table(colIndex, rowIndex, colIndex, rowIndex)
        //console.log(table.toString())

        var queue = []
        queue.push(cellWrapper)
        visitedCells[rowIndex][colIndex] = true

        while (queue.length > 0) {
          //console.log("enter again")
          var firstCellWrapper = queue.shift()
          var firstCell = firstCellWrapper.excel_cell
          firstCellWrapper.ft_table = table
          
          //console.log(queue.length)
          for (let index = 0; index < 4; index++) {
            var newRowIndex = firstCell.rowIndex - this.rowBase + moveRow[index]
            var newColIndex = firstCell.columnIndex - this.columnBase + moveCol[index]
            //console.log(newRowIndex, newColIndex)
            if (newRowIndex < 0 || newColIndex < 0 || newRowIndex >= this.sheetWrapper.length || newColIndex >= this.sheetWrapper[0].length) continue
            //console.log("go 1")
            if (visitedCells[newRowIndex][newColIndex] == true) continue
            //console.log("go 2")
            if (this.sheetWrapper[newRowIndex][newColIndex].cellType !== "empty") {
              //console.log("go 3")
              queue.push(this.sheetWrapper[newRowIndex][newColIndex])
              visitedCells[newRowIndex][newColIndex] = true
              
              //update table
              table.up = table.up > newRowIndex? newRowIndex : table.up
              table.bottom = table.bottom < newRowIndex? newRowIndex : table.bottom
              table.left = table.left > newColIndex? newColIndex : table.left
              table.right = table.right < newColIndex? newColIndex : table.right
              //console.log("update", table.toString())
            }
          }
        }

        //change table to real sheet not relative sheet
        table.up += this.rowBase
        table.bottom += this.rowBase
        table.left += this.columnBase
        table.right += this.columnBase

        this.tableSet.add(table)
      }
      
    }

    //debug
    //for (let index = 0; index < this.sheetWrapper.length; index++) {
      //for (let index2 = 0; index2 < this.sheetWrapper[index].length; index2++) {
        //const cellWrapper = this.sheetWrapper[index][index2];
        //if (cellWrapper !== undefined)
          //console.log(cellWrapper.excel_cell.address, cellWrapper.ft_table.toString())
      //}
    //}


    console.log("  --- extract table: end ---")    
  }

  /**
   * cell array is amcheck's cell array
   */
  extractCellArray() {
    
  }

  /**
   * gap template: what is gap template
   */
  extractGapTemplate() {

  }
}

class Table {
  constructor(left, up, right, bottom) {
    this.left = left
    this.up = up
    this.right = right
    this.bottom = bottom
  }

  toString() {
    return "Table:(" + this.up + "," + this.left + ")--(" + this.bottom + "," + this.right + ")"
  }
}