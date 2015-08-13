'use strict';

var fs = require('fs');

var _ = require('lodash');
var XLSX = require('xlsx');

var utils = require('./utils');

/**
 * Write
 */
var Workbook = require('./workbook');
var wb = new Workbook();
var lines = 10;

/**
 * Fibonacci
 */
var data = _.times(lines, function(n) {
  if(n < 2) return [n, {v:1}];

  var formula = '=SUM(' +
    Workbook.encodeCell({c:1, r:n-2}) +
    ':' +
    Workbook.encodeCell({c:1, r:n-1}) +
    ')';

  return [{v:n}, {f:formula}];
});
wb.pushSheet('Fibonacci', data);

/**
 * Suite géométrique
 */
var data = _.times(lines, function(n) {
  if(n < 1) return [n, 1];

  var formula = '=FACT(' +
    '$' + Workbook.encodeCol(0) + '$' + Workbook.encodeRow(1) +
    ':' +
    Workbook.encodeCell({c:0, r:n}) +
    ')';

  return [{v:n}, {f:formula}];
});
wb.pushSheet('Factorielle', data);

fs.writeFileSync('out.xlsx', wb.write());

/**
 * Read
 */
var book = XLSX.readFile('out.xlsx', {
  cellStyles: true
});

/**
 * Rewrite
 */

var fibo = book.Sheets.Fibonacci;
utils.fillBackground(fibo,  'F79646');
utils.drawRect(fibo);

var fact = book.Sheets.Factorielle;
utils.fillBackground(fact,  'AED07C');
utils.drawRect(fact);

XLSX.writeFile(book, 'out-stylised.xlsx');

console.log('end');