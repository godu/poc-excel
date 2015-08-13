'use strict';

var XLSX = require('xlsx');

function forEachCell(sheet, fun) {
  Object.keys(sheet).filter(function(key) {
    return [
      '!ref',
      '!cols',
      '!merges'
    ].indexOf(key) === -1;
  }).forEach(function(key) {
    fun(sheet[key], key, sheet);
  });
}

function fillBackground(sheet, color) {
  forEachCell(sheet, function(cell) {
      var style = cell.s = (cell.s || {});
      var fill = style.fill = (style.fill || {});
      fill.fgColor = {
        rgb: color
      };
  });
}

function drawRect(sheet, range) {
  range = range || sheet['!ref'];
  range = XLSX.utils.decode_range(range);
  
  function genBorder() {
    return {
      style: 'medium',
      color: {auto: 1}
    };
  }
  
  forEachCell(sheet, function(cell, key) {
    var coord = XLSX.utils.decode_cell(key);
    
    var style = cell.s = cell.s || {};
    var border = style.border = style.border || {};
    
    if(coord.r === range.s.r) border.top = genBorder();
    if(coord.r === range.e.r) border.bottom = genBorder();
    if(coord.c === range.s.c) border.left = genBorder();
    if(coord.c === range.e.c) border.right = genBorder();     
  });
}

module.exports = {
	forEachCell: forEachCell,
	fillBackground: fillBackground,
	drawRect: drawRect,
}