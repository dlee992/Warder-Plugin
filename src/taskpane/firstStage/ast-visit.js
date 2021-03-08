
module.exports = {visitNode}

function visitNode(node) {
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
    ast.children.push(visitNode(node.left))
    ast.children.push(visitNode(node.right))
    return ast
}

function visitFunction(node) {
    var ast = {id: node.name}
    ast.children = []
    node.arguments.forEach(arg => ast.children.push(visitNode(arg)))
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
    ast.children.push(visitNode(node.left))
    ast.children.push(visitNode(node.right))
    return ast
}

function visitUnaryExpression(node) {
    var ast = {id: node.operator}
    ast.children = []
    ast.children.push(visitNode(node.operand))
    return ast
}
