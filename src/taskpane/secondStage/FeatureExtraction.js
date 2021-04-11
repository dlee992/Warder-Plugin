module.exports = {FeatureExtraction}

class Table {
  constructor(left, up, right, bottom) {
    this.left = left
    this.up = up
    this.right = right
    this.bottom = bottom
  }

  toString() {
    return "Table:left" + this.left + "up" + this.up + "right" + this.right + "bottom" + this.bottom
  }

  
}

class FeatureExtraction {
  constructor(finalFirstClusters, formulaCells, numberCells, stringCells) {
    this.finalFirstClusters = finalFirstClusters
    this.formulaCells = formulaCells
    this.numberCells = numberCells
    this.stringCells = stringCells

    extractCellAddress()

    extractTable()
    extractLabel()

    extractAlliance()

    extractCellArray()

    extractGapTemplate()

  }

  /**
   * f1: row index, column index
   * f2?.: a cell can reference many cells, whether referenced cells are in the same row or column with the cell
   */
  extractCellAddress() {
    // nothing to do
  }

  /**
   * label: is header
   */
  extractLabel() {

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