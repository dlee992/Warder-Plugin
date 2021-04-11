module.exports = {extractFeature}

function extractFeature(finalFirstClusters, formulaCells, numberCells, stringCells) {

    extractCellAddress()
    extractLabel()
    extractAlliance()
    extractTable()
    extractCellArray()
    extractGapTemplate()

}

/**
 * f1: row index, column index
 * f2?.: a cell can reference many cells, whether referenced cells are in the same row or column with the cell
 */
function extractCellAddress() {

}

/**
 * label: is header
 */
function extractLabel() {

}

/**
 * example: SUM(A2:E2), then A2,B2...E2 are alliances
 */
function extractAlliance() {

}

/**
 * table: means a contiguous cell range, in which, no empty row or column
 */
function extractTable() {

}

/**
 * cell array is amcheck's cell array
 */
function extractCellArray() {

}

/**
 * gap template: what is gap template
 */
function extractGapTemplate() {

}