export class CellWrapper {
    constructor(excel_cell, syntax_tree, astString, index, cellType) {
      this.excel_cell = excel_cell
      this.syntax_tree = syntax_tree
      this.astString = astString
      this.index = index
      this.cellType = cellType
      this.ft_cellAddressRow = "Row" + excel_cell.rowIndex
      this.ft_cellAddressColumn = "Column" + excel_cell.columnIndex
    }
  }