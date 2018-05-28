//----------------------------------------------------------------------------
// common
//----------------------------------------------------------------------------

function __isNull(obj)
{
	return (obj == null);
}

function __isEmpty(obj)
{
	if (__isNull(obj)) {
		return true;
	}
	return (__tostr(obj) == "");
}

function __isString(obj)
{
	if (obj == null) {
		return false;
	}
	return (typeof(obj) == "string" || obj instanceof String);
}

function __isNumber(obj)
{
	if (obj == null) {
		return false;
	}
	return (typeof(obj) == "number" || obj instanceof Number);
}

function __isObject(obj)
{
	if (obj == null) {
		return false;
	}
	return (typeof(obj) == "object" || obj instanceof Object);
}

function __tostr(obj)
{
	return "" + obj;
}

function __trim(text)
{
	return (__tostr(text).replace(/(^\s+|\s+$)/g, ""));
}

function __trimright(text)
{
	return(__tostr(text).replace(/\s+$/, ""));
}

function __toNumber(obj)
{
	if (obj == null) {
		return 0;
	}
	if (typeof(obj) == "number") {
		return obj;
	}
	if (obj instanceof Number) {
		return obj.valueOf();
	}
	if (__isString(obj)) {
		var n = new Number(__trim(obj));
		if (! isNaN(n)) {
			return n.valueOf();
		}
	}
	return 0;
}

function __toFixNumber(obj)
{
	var n = __toNumber(obj);
	if (n >= 0) {
		return Math.floor(n);
	} else {
		return Math.ceil(n);
	}
}

function __toNumString(value, len)
{
	var text = new Number(__toFixNumber(value)).toString();
	if (! __isNumber(len)) {
		return text;
	}
	if (len <= text.length) {
		return text;
	}
	if (len > 16) {
		len = 16;
	}
	return ("000000000000000" + text).match(new RegExp("................".substr(0, len) + "$"));
}

function __numstr(value, len)
{
	return __toNumString(value, len);
}

function __toHexString(value, len)
{
	var text = new Number(__toFixNumber(value)).toString(16);
	if (! __isNumber(len)) {
		return text;
	}
	if (len <= text.length) {
		return text;
	}
	return __zerostr(len - text.length) + text;
}

function __toOctString(value, len)
{
	var text = new Number(__toFixNumber(value)).toString(8);
	if (! __isNumber(len)) {
		return text;
	}
	if (len <= text.length) {
		return text;
	}
	return __zerostr(len - text.length) + text;
}

function __toBinString(value, len)
{
	var text = new Number(__toFixNumber(value)).toString(2);
	if (! __isNumber(len)) {
		return text;
	}
	if (len <= text.length) {
		return text;
	}
	return __zerostr(len - text.length) + text;
}

function __hexstr_lc(value, len)
{
	return __toHexString(value, len);
}

function __hexstr_uc(value, len)
{
	return __toHexString(value, len).toUpperCase();
}

function __zerostr(n) {
	var s = "";
	for (var i = 0; i < n; i++) {
		s += "0";
	}
	return s;
}

function __formatstr(fmt, objs)
{
	var fmtx = (""+fmt).replace(/\n/g, String.fromCharCode(0xff));
	var msg = "";
	if (objs instanceof Array) {
		while (fmtx != "") {
			if (fmtx.match(/^(.*?)([{](\d+)[}])(.*)$/)) {
				var m1 = RegExp.$1;
				var m2 = RegExp.$2;
				var m3 = RegExp.$3;
				var m4 = RegExp.$4;
				msg += m1;
				var i = new Number(m3).valueOf();
				if (typeof(objs[i]) != "undefined") {
					msg += ""+objs[i];
				} else {
					msg += m2;
				}
				fmtx = m4;
			} else {
				msg += fmtx;
				fmtx = "";
			}
		}
	}
	else if (objs != null) {
		msg = fmtx.replace(/[{]0[}]/g, ""+objs);
	}
	else {
		msg = fmtx;
	}
	return msg.replace(new RegExp(String.fromCharCode(0xff), "g"), "\n");
}
