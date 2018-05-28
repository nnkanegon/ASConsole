//----------------------------------------------------------------------------
// sample script of ASConsole user library
//----------------------------------------------------------------------------
// 
// IP calculator
// 
// usage:
// ipcalc(<subnet)
// 
// ex)
// ipcalc("192.168.250.25/24")                    # IP & mask(bit-length)
// ipcalc("192.168.250.25/255.255.255.0")         # IP & mask
// ipcalc("192.168.250.25,256")                   # IP & number of ip
// ipcalc("192.168.250.0-192.168.250.255")        # IP range
//
// example:
// js> ipcalc("192.168.0.10/16")
// NETWORK:         192.168.0.0/16
// START_ADDR:      192.168.0.0
// END_ADDR:        192.168.255.255
// NET_MASK:        255.255.0.0
// IP_COUNT:        65536
// START_ADDR(BIT): 11000000.10101000.00000000.00000000
// NET_MASK(BIT):   11111111.11111111.00000000.00000000
// INPUT_ADDR(BIT): 11000000.10101000.00000000.00001010
//----------------------------------------------------------------------------


utils.addHelp("ipcalc", ["ipcalc(<subnet>)","IP calculator"]);

function IPv4Calc(input)
{
	this.input_ip = 0;
	this.ip = 0;
	this.mask = 0;

	if (input) {
		this.parse(input);
	}
}

IPv4Calc.prototype.ip_from_array = function(ary)
{
	return ary[0] * (1<<24) + ary[1] * (1<<16) + ary[2] * (1<<8) + ary[3];
}

IPv4Calc.prototype.mask_from_array = function(ary)
{
	var mask_bit = ary[0] * (1<<24) + ary[1] * (1<<16) + ary[2] * (1<<8) + ary[3];
	return this.mask_from_bit(mask_bit);
}

IPv4Calc.prototype.ip_to_array = function(ip)
{
	return [(ip & 0xff000000)>>>24, (ip & 0x00ff0000)>>>16, (ip & 0x0000ff00)>>>8, (ip & 0x000000ff)];
}

IPv4Calc.prototype.ip_from_text = function(iptext)
{
	if (/^(\d+)\.(\d+)\.(\d+)\.(\d+)$/.exec(iptext) == null) {
		throw "wrong ip";
	}
	var d1 = new Number(__trim(RegExp.$1)).valueOf();
	var d2 = new Number(__trim(RegExp.$2)).valueOf();
	var d3 = new Number(__trim(RegExp.$3)).valueOf();
	var d4 = new Number(__trim(RegExp.$4)).valueOf();
	if (d1>255 || d2>255 || d3>255 || d4>255) {
		throw "wrong ip";
	}
	return this.ip_from_array([d1, d2, d3, d4]);
}

IPv4Calc.prototype.mask_from_text = function(masktext)
{
	masktext = __trim(masktext);

	var mask = 0;
	if (masktext == "") {
		return 32;
	}
	if (/^\d+$/.exec(masktext) != null) {
		mask = new Number(__trim(masktext)).valueOf();
		if (mask > 32) {
			throw "wrong mask";
		}
	}
	else if (/^(\d+)\.(\d+)\.(\d+)\.(\d+)$/.exec(masktext) != null) {
		var d1 = new Number(__trim(RegExp.$1)).valueOf();
		var d2 = new Number(__trim(RegExp.$2)).valueOf();
		var d3 = new Number(__trim(RegExp.$3)).valueOf();
		var d4 = new Number(__trim(RegExp.$4)).valueOf();
		if (d1>255 || d2>255 || d3>255 || d4>255) {
			throw "wrong mask";
		}
		mask = this.mask_from_array([d1, d2, d3, d4]);
	}
	else {
		throw "wrong mask";
	}
	return mask;
}

IPv4Calc.prototype.mask_from_count = function(masktext)
{
	var mask = null;

	masktext = __trim(masktext);
	if (masktext == "") {
		throw "wrong ip count";
	}
	if (/^\d+$/.exec(masktext) == null) {
		throw "wrong ip count";
	}
	ip_count = new Number(masktext).valueOf();
	for (var i = 0; i <= 32; i++) {
		if (ip_count == Math.pow(2, i)) {
			mask = 32 - i;
		}
	}
	if (mask == null) {
		throw "wrong ip count";
	}
	return mask;
}

IPv4Calc.prototype.mask_to_count = function(mask)
{
	return Math.pow(2, 32 - mask);
}

IPv4Calc.prototype.mask_to_bit = function(mask)
{
	var mask_bit;
	if (mask == 0) {
		return 0x00000000;
	}
	return (0x80000000 >> (mask - 1)) >>> 0;
}

IPv4Calc.prototype.mask_from_bit = function(mask_bit)
{
	if (mask_bit == 0) {
		return 0;
	}
	for (var i = 1; i <= 32; i++) {
		if (((0x80000000 >> (i - 1)) >>> 0) == mask_bit) {
			return i;
		}
	}
	throw new Error("wrong mask");
}

IPv4Calc.prototype.adjust_ip = function(ip, mask)
{
	var bit = this.mask_to_bit(mask);
	ip = (ip & bit) >>> 0;
	return ip;
}

IPv4Calc.prototype.format_ip = function(ip)
{
	var ip_ary = this.ip_to_array(ip);
	return "" + ip_ary[0] + "." + ip_ary[1] + "." + ip_ary[2] + "." + ip_ary[3];
}

IPv4Calc.prototype.format_ip_bit = function(ip)
{
	var ip_ary = this.ip_to_array(ip);
	return "" + __toBinString(ip_ary[0], 8) + "." + __toBinString(ip_ary[1], 8) + "." +
				__toBinString(ip_ary[2], 8) + "." + __toBinString(ip_ary[3], 8);
}

IPv4Calc.prototype.format_network = function(ip, mask)
{
	var text = this.format_ip(ip) + "/" + mask;
	return text;
}

IPv4Calc.prototype.dump = function()
{
	utils.println("NETWORK:	" + this.format_network(this.ip, this.mask));
	utils.println("START_ADDR:	" + this.format_ip(this.ip));
	utils.println("END_ADDR:	" + this.format_ip(this.ip + this.mask_to_count(this.mask) - 1));
	utils.println("NET_MASK:	" + this.format_ip(this.mask_to_bit(this.mask)));
	utils.println("IP_COUNT:	" + this.mask_to_count(this.mask));
	utils.println("START_ADDR(BIT):	" + this.format_ip_bit(this.ip));
	utils.println("NET_MASK(BIT):	" + this.format_ip_bit(this.mask_to_bit(this.mask)));
	utils.println("INPUT_ADDR(BIT):	" + this.format_ip_bit(this.input_ip));
}

IPv4Calc.prototype.parse = function(text)
{
	if (text == null) {
		throw "no param";
	}
	var ip = 0;
	var mask = 0;
	if (/([\d.]+)(([-\/,])(.*))?/.exec(__trim(text)) == null) {
		throw "wrong param";
	}
	var ip_text = RegExp.$1;
	var sep = RegExp.$3;
	var mask_text = RegExp.$4;

	ip = this.ip_from_text(ip_text);
	this.input_ip = ip;

	if (sep == "") {
		mask = 32;
	}
	else if (sep == "/") {
		mask = this.mask_from_text(mask_text);
	}
	else if (sep == ",") {
		mask = this.mask_from_count(mask_text);
	}
	else if (sep == "-") {
		var ip_to = this.ip_from_text(mask_text);
		if (ip > ip_to) {
			throw "wrong ip range";
		}
		var ip_count = ip_to - ip + 1;
		mask = this.mask_from_count(""+ip_count);
		if (mask == null) {
			throw "wrong ip range";
		}
		var ip_from = this.adjust_ip(ip, mask);
		if (ip != ip_from || (ip + ip_count - 1) != ip_to) {
			throw "wrong ip range";
		}
	}
	this.ip = this.adjust_ip(ip, mask);
	this.mask = mask;
}


///////////////////////////////////////////////////////////////////////
// helper

function ipcalc(input)
{
	try {
		var ipcalc = new IPv4Calc(input);
		ipcalc.dump();
	}
	catch(e) {
		utils.println(__tostr(e));
	}
}

