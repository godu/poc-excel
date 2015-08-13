'use strict';

var XLSX = require('xlsx');
var _ = require('lodash');

function datenum(v, date1904) {
	if(date1904) v+=1462;
	var epoch = Date.parse(v);
	return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}



function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}

Workbook.prototype.pushSheet = function(ws_name, ws, options) {
	this.Sheets[ws_name] = this.sheetFromArrayOfArrays(ws, options);
};

Workbook.prototype.sheetFromArrayOfArrays = function(data, opts) {
	var ws = {};
	var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			
			var d = data[R][C];
			var cell;
			if(d && (d.v || d.f)) cell = d;
			else cell = {v: d};
			var cell_ref = Workbook.encodeCell({c:C,r:R});
			
			if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n'; cell.z = XLSX.SSF._table[14];
				cell.v = datenum(cell.v);
			}
			else if(typeof cell.v === 'string') cell.t = 's';
			
			ws[cell_ref] = cell;
		}
	}
	if(range.s.c < 10000000) ws['!ref'] = Workbook.encodeRange(range);
	return ws;
};

Workbook.encodeCell = XLSX.utils.encode_cell.bind(XLSX.utils);
Workbook.encodeRange = XLSX.utils.encode_range.bind(XLSX.utils);
Workbook.encodeCol = XLSX.utils.encode_col.bind(XLSX.utils);
Workbook.encodeRow = XLSX.utils.encode_row.bind(XLSX.utils);

Workbook.prototype.write = function(options) {
	var defaults = {
      bookType:'xlsx',
      bookSST: false,
      type:'binary'
    };
	
	this.SheetNames = _.keys(this.Sheets);
	
	var data = XLSX.write(this, _.defaults(options || {}, defaults));
	if(!data) return false;
    var buffer = new Buffer(data, 'binary');
    return buffer;
}

module.exports = Workbook;