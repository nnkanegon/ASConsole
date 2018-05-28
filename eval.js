//----------------------------------------------------------------------------
// eval helper (JavaScript)
//----------------------------------------------------------------------------

// Host-Script interface
// utils

var _result = "";
var _n = 0;
var __narray = [];

var __remoteEval = null;

function Console()
{
}
Console.prototype.log = function(msg)
{
	utils.println(msg);
}
var console = new Console();

///////////////////////////////////////////////////////////////////////
// eval helper

function ASConsole_execEval(_text)
{
	if (_text == "__detach()") {
		__detach();
		return null;
	}
	if (__remoteEval == null) {
		_text = _text.replace(/^(hex|dec|oct|bin)$/, "$1()"); // special convert
	}
	var _r;
	try {
		if (__remoteEval != null) {
			_r = __remoteEval(_text);
		}
		else {
			_r = eval(_text);
		}
	}
	catch (_e) {
		utils.setError(_e);
		if (__remoteEval != null) {
			__detach();
		}
	}
	if (utils.getError() != null) {
		return null;
	}
	_result = _r;
	if (__isNumber(_r)) {
		_n = _r;
		__narray = [_n];
	}
	return _r;
}


///////////////////////////////////////////////////////////////////////

function exportVariable(name, value)
{
	if (! __isString(name)) {
		throw new Error(utils.getText("BAD_PARAMETER"));
	}
	utils.setVariable(name, value);
}

function importVariable(name)
{
	if (! __isString(name)) {
		throw new Error(utils.getText("BAD_PARAMETER"));
	}
	return utils.getVariable(name);
}


///////////////////////////////////////////////////////////////////////

function __print(s)
{
	utils.print(s);
}

function __println(s)
{
	utils.println(s);
}


///////////////////////////////////////////////////////////////////////

function comInfo(obj, member)
{
	utils.comInfo(obj, member);
}

function comMethods(obj)
{
	utils.comMethods(obj);
}

function comProperties(obj)
{
	utils.comProperties(obj);
}

function comProps(obj)
{
	utils.comProperties(obj);
}


///////////////////////////////////////////////////////////////////////

function hex()
{
	var ary = __getNumeralParams.apply(this, arguments);
	var a = [];
	for (var i = 0; i < ary.length; i++) {
		// mask 0xffffffff
		var n = ary[i] | 0;
		if (n >= 0) {
			a.push(__formatstr("0x{0}", __hexstr_uc(n)));
		}
		else {
			a.push(__formatstr("0x{0}", __hexstr_uc(0x100000000+n)));
		}
	}
	__narray = ary;
	return a.join(", ");
}

function dec()
{
	var ary = __getNumeralParams.apply(this, arguments);
	var a = [];
	for (var i = 0; i < ary.length; i++) {
		var n = __numstr(Math.abs(ary[i]));
		var prefix = (ary[i] < 0) ? "-" : ""
		a.push(__formatstr("{0}{1}", [prefix, n]));
	}
	__narray = ary;
	return a.join(", ");
}

function oct()
{
	var ary = __getNumeralParams.apply(this, arguments);
	var a = [];
	for (var i = 0; i < ary.length; i++) {
		// mask 0xffffffff
		var n = ary[i] | 0;
		if (n == 0) {
			a.push("0");
		}
		else if (n >= 0) {
			a.push(__formatstr("0{0}", __toOctString(n)));
		}
		else {
			a.push(__formatstr("0{0}", __toOctString(0x100000000+n)));
		}
	}
	__narray = ary;
	return a.join(", ");
}

function bin()
{
	var ary = __getNumeralParams.apply(this, arguments);
	var a = [];
	for (var i = 0; i < ary.length; i++) {
		// mask 0xffffffff
		var n = ary[i] | 0;
		if (n >= 0) {
			a.push(__formatstr("0b{0}", __toBinString(n)));
		}
		else {
			a.push(__formatstr("0b{0}", __toBinString(0x100000000+n)));
		}
	}
	__narray = ary;
	return a.join(", ");
}


///////////////////////////////////////////////////////////////////////
// internal functions (interface)

function ASConsole_onActive()
{
}


///////////////////////////////////////////////////////////////////////
// internal functions

function __getNumeralParams()
{
	if (arguments.length == 0) {
		return __narray;
	}
	var ary = [];
	for (var i = 0; i < arguments.length; i++) {
		ary.push(arguments[i]);
	}
	return ary;
}
