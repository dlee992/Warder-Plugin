
export class TreeVisitor {
  constructor() {
  }

  cdtSize() {
  }

  buildCdtTree() {
  }

  astSize(node) {
    switch (node.type) {
      case 'cell':
        return 1
      case 'cell-range':
        return 1 + this.astSize(node.left) + this.astSize(node.right)
      case 'function':
        var sum = 1
        node.arguments.forEach(arg => {
          sum += this.astSize(arg)
        })
        return sum
      case 'number':
        return 1 
      case 'text':
        return 1
      case 'logical':
        return 1
      case 'binary-expression':
        return 1 + this.astSize(node.left) + this.astSize(node.right)
      case 'unary-expression':
        return 1 + this.astSize(node.operand)
    } 
  }

  setSheetWrapper(sheetWrapper, rowBase, columnBase) {
    this.sheetWrapper = sheetWrapper
    this.rowBase = rowBase
    this.columnBase = columnBase
    // console.log(rowBase, columnBase)
  }

  computeReferenceSet(cellWrapper) {
    console.log(cellWrapper.excel_cell.formulas[0][0])
    this.cellWrapper = cellWrapper
    var referenceList = this.getReference(cellWrapper.syntax_tree)
    cellWrapper.referenceSet = new Set()
    var references = ""
    for (let index = 0; index < referenceList.length; index++) {
      const rCellWrapper = referenceList[index]
      cellWrapper.referenceSet.add(rCellWrapper)
      references += " & " + rCellWrapper.excel_cell.address
    }
    console.log(cellWrapper.excel_cell.address, references)
  }

  getReference(node) {
    var references = []
    switch (node.type) {
      case 'cell':
        var curCellWrapper = this.getCellWrapper(node.key)
        // console.log(curCellWrapper)
        references.push(curCellWrapper)
        return references

      case 'cell-range':
        var firstCellWrapper = this.getReference(node.left)[0]
        var lastCellWrapper =  this.getReference(node.right)[0]
        // console.log(firstCellWrapper)
        //todo: push a list into 
        // console.log('add cell range')
        var firstCell = firstCellWrapper.excel_cell
        var lastCell = lastCellWrapper.excel_cell
        for(let rowIndex = firstCell.rowIndex; rowIndex <= lastCell.rowIndex; rowIndex++) {
          for(let colIndex = firstCell.columnIndex; colIndex <= lastCell.columnIndex; colIndex++) {
            const wrapper = this.sheetWrapper[rowIndex - this.rowBase][colIndex - this.columnBase]
            references.push(wrapper)
            // console.log(wrapper.excel_cell.address)
          }
        }
        return references

      case 'function':
        node.arguments.forEach(arg => {
          references = references.concat(this.getReference(arg))
        })
        return references
      case 'number':
        return []
      case 'text':
        return []
      case 'logical':
        return []
      case 'binary-expression':
        references = references.concat(this.getReference(node.left))
        references = references.concat(this.getReference(node.right))
        return references
      case 'unary-expression':
        references = references.concat(this.getReference(node.operand))
        return references
    }
  }
  
  getCellWrapper(cellR1C1) {
    //console.log(cellR1C1)
    var numbers = cellR1C1.match(/-?\d+/g)
    for (let index = 0; index < numbers.length; index++) {
      numbers[index] = parseInt(numbers[index], 10)
    }
    //console.log(numbers)

    var rowIndex
    var colIndex
    //compute row
    var rowIdentify = cellR1C1.charAt(1)
    if (rowIdentify === '[') {
      // console.log('go 1')
      rowIndex = this.cellWrapper.excel_cell.rowIndex + numbers[0]
      // console.log('go 2')
    }
    else if (rowIdentify === 'C') {
      rowIndex = this.cellWrapper.excel_cell.rowIndex
    }
    else {
      rowIndex = numbers[0]
    }
    // console.log('go 2.5')
    //compute column
    var colIdentify = cellR1C1.charAt(cellR1C1.length - 1)
    // console.log(colIdentify)
    if (colIdentify === ']') {
      // console.log('go 5')
      colIndex = this.cellWrapper.excel_cell.columnIndex + numbers[numbers.length - 1]
    }
    else if (colIdentify === 'C') {
      // console.log('go 3')
      colIndex = this.cellWrapper.excel_cell.columnIndex
      // console.log('go 4')
    }
    else {
      // console.log('go 6')
      colIndex = numbers[numbers.length - 1]
    }
    // console.log(rowIndex, colIndex)

    rowIndex -= this.rowBase
    colIndex -= this.columnBase
    // console.log(rowIndex, colIndex) 

    // console.log(this.sheetWrapper.length, this.sheetWrapper[0].length)
    var curCellWrapper = this.sheetWrapper[rowIndex][colIndex]
    // console.log('go end')
    return curCellWrapper
  }

  buildAstTree(node) {
    switch (node.type) {
      case 'cell':
        return this.visitCell(node)
      case 'cell-range':
        return this.visitCellRange(node)
      case 'function':
        return this.visitFunction(node)
      case 'number':
        return this.visitNumber(node)
      case 'text':
        return this.visitText(node)
      case 'logical':
        return this.visitLogical(node)
      case 'binary-expression':
        return this.visitBinaryExpression(node)
      case 'unary-expression':
        return this.visitUnaryExpression(node)
    }
  }

  visitCell(node) {
      return {id: node.key}
  }

  visitCellRange(node) {
      var ast = {id: node.type}
      ast.children = []
      ast.children.push(this.buildAstTree(node.left))
      ast.children.push(this.buildAstTree(node.right))
      return ast
  }

  visitFunction(node) {
      var ast = {id: node.name}
      ast.children = []
      node.arguments.forEach(arg => ast.children.push(this.buildAstTree(arg)))
      return ast
  }

  visitNumber(node) {
      return {id: node.value}
  }

  visitText(node) {
      return {id: node.value}
  }

  visitLogical(node) {
      return {id: node.boolean}
  }

  visitBinaryExpression(node) {
      var ast = {id: node.operator}
      ast.children = []
      ast.children.push(this.buildAstTree(node.left))
      ast.children.push(this.buildAstTree(node.right))
      return ast
  }

  visitUnaryExpression(node) {
      var ast = {id: node.operator}
      ast.children = []
      ast.children.push(this.buildAstTree(node.operand))
      return ast
  }
}