
module.exports = {buildAstTree, buildCdtTree, astSize, cdtSize}


function cdtSize() {

}

function buildCdtTree() {

}


function astSize(node) {
  switch (node.type) {
    case 'cell':
      return 1
    case 'cell-range':
      return 1 + astSize(node.left) + astSize(node.right)
    case 'function':
      var sum = 1
      node.arguments.forEach(arg => {
        sum += astSize(arg)
      })
      return sum
    case 'number':
      return 1 
    case 'text':
      return 1
    case 'logical':
      return 1
    case 'binary-expression':
      return 1 + astSize(node.left) + astSize(node.right)
    case 'unary-expression':
      return 1 + astSize(node.operand)
  } 
}

function buildAstTree(node) {
  switch (node.type) {
    case 'cell':
      return visitCell(node)
    case 'cell-range':
      return visitCellRange(node)
    case 'function':
      return visitFunction(node)
    case 'number':
      return visitNumber(node)
    case 'text':
      return visitText(node)
    case 'logical':
      return visitLogical(node)
    case 'binary-expression':
      return visitBinaryExpression(node)
    case 'unary-expression':
      return visitUnaryExpression(node)
  }
}

function visitCell(node) {
    return {id: node.key}
}

function visitCellRange(node) {
    var ast = {id: node.type}
    ast.children = []
    ast.children.push(buildAstTree(node.left))
    ast.children.push(buildAstTree(node.right))
    return ast
}

function visitFunction(node) {
    var ast = {id: node.name}
    ast.children = []
    node.arguments.forEach(arg => ast.children.push(buildAstTree(arg)))
    return ast
}

function visitNumber(node) {
    return {id: node.value}
}

function visitText(node) {
    return {id: node.value}
}

function visitLogical(node) {
    return {id: node.boolean}
}

function visitBinaryExpression(node) {
    var ast = {id: node.operator}
    ast.children = []
    ast.children.push(buildAstTree(node.left))
    ast.children.push(buildAstTree(node.right))
    return ast
}

function visitUnaryExpression(node) {
    var ast = {id: node.operator}
    ast.children = []
    ast.children.push(buildAstTree(node.operand))
    return ast
}
