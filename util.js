//----------------------------------------------------------------------------
// utilities
//----------------------------------------------------------------------------

var objShell = new ActiveXObject("WScript.Shell");
var objFS = new ActiveXObject("Scripting.FileSystemObject");

var objXmlHttp;
try {
	objXmlHttp = new ActiveXObject('Msxml2.XMLHTTP');
}
catch (e) {
	objXmlHttp = new ActiveXObject('Microsoft.XMLHTTP');
}

var adTypeBinary = 1;
var adTypeText = 2;
var adSaveCreateOverWrite = 2;

var ForReading = 1;
var ForWriting = 2;
var ForAppending = 8;

var Log_Error = 0;
var Log_Trace = 1;
var Log_Debug = 2;
var LogPrefix = ["[ERROR]", "[TRACE]", "[DEBUG]"];


///////////////////////////////////////////////////////////////////////
// Utils

function Utils()
{
	if (typeof(WScript) != "undefined") {
		this.appDir = objFS.GetParentFolderName(WScript.ScriptFullName);
	}
	else if (typeof(document) != "undefined") {
		this.appDir = unescape(objFS.GetParentFolderName(document.location.pathname));
	}
	else {
		this.appDir = objShell.CurrentDirectory;
	}
	if (this.appDir.indexOf('/') === 0) {
		this.appDir = this.appDir.substring(1);
	}
	this.userDir = this.appDir;
	this.logLevel = Log_Error;
	this.sharedVariables = {}
	this.logPath = null;
	this.itemSeparator = ", "
	this.messages = messageTable["en"];
	this.language = "en";
	this.objError = null;
	this.postAction = null;
	this.scRubyScript = null;
	this.jsBrowserAttached = false;
	this.jsBrowserAttachedPrev = false;
	this.jsBrowserTitle = null;
}

// for debug (temporary method)
Utils.prototype.typeCheck = function(v)
{
	utils.println(typeof(v));
	utils.println(__tostr(v));
	utils.println((v == null) ? "v is null" : "v is not null");
	utils.println((typeof(v) == "undefined") ? "v is undefined" : "v is defined");
}

Utils.prototype.setSeparator = function(s)
{
	this.itemSeparator = __tostr(s);
}

Utils.prototype.getSeparator = function()
{
	return this.itemSeparator;
}

Utils.prototype.getAppDir = function()
{
	return this.appDir;
}

Utils.prototype.getUserDir = function()
{
	return this.userDir;
}

Utils.prototype.getLanguage = function()
{
	return this.language;
}

Utils.prototype.getScriptManager = function()
{
	return scriptManager;
}

Utils.prototype.setVariable = function(name, value)
{
	if (name == null) {
		return;
	}
	this.sharedVariables[name] = value;
}

Utils.prototype.getVariable = function(name)
{
	if (name == null) {
		return null;
	}
	if (! (name in this.sharedVariables)) {
		return null;
	}
	return this.sharedVariables[name];
}

Utils.prototype.createError = function(code, text)
{
	return new Error(code, text);
}

Utils.prototype.clearError = function()
{
	this.objError = null;
}

Utils.prototype.setError = function(obj)
{
	if (__isNull(this.objError)) {
		this.objError = obj;
	}
}

Utils.prototype.getError = function()
{
	return this.objError;
}

Utils.prototype.isErrorEmpty = function()
{
	return __isNull(this.objError);
}

Utils.prototype.clearPostAction = function()
{
	this.postAction = null;
}

Utils.prototype.setPostAction = function(action)
{
	this.postAction = action;
}

Utils.prototype.getPostAction = function()
{
	return __trim(__tostr(this.postAction));
}

Utils.prototype.setLogLevel = function(level)
{
	this.logLevel = level;
}

Utils.prototype.log = function(text, logLevel)
{
	if (logLevel > this.logLevel) {
		return;
	}
	if (typeof(WScript) != "undefined") {
		WScript.Echo(text);
	}
	if (this.logPath == null) {
		this.logPath = objFS.BuildPath(this.userDir, "ascon.log");
	}
	var ts = objFS.OpenTextFile(this.logPath, ForAppending, true);
	var prefix = (logLevel >= 0 && logLevel < LogPrefix.length) ? LogPrefix[logLevel] : "";
	ts.WriteLine(formatDate(new Date()) + " " + prefix + " " + text);
	ts.Close();
}

Utils.prototype.logError = function(text)
{
	this.log(text, Log_Error);
}

Utils.prototype.logTrace = function(text)
{
	this.log(text, Log_Trace);
}

Utils.prototype.logDebug = function(text)
{
	this.log(text, Log_Debug);
}

Utils.prototype.alert = function(msg)
{
	if (typeof(window) != "undefined") {
		window.alert(msg);
	}
}

Utils.prototype.getJavaScriptObjectName = function(obj)
{
	if (__isObject(obj)) {
		if ((typeof(obj.constructor) != "undefined") && obj.constructor != null) {
			if (__tostr(obj.constructor).match(/^\s*function +(\w+)/)) {
				return RegExp.$1;
			}
		}
	}
	return typeof(obj);
}

Utils.prototype.getJavaScriptProperties = function(obj)
{
	if (! __isObject(obj)) {
		return "";
	}
	try {
		var ary = [];
		for (var member in obj) {
			if (typeof(obj[member]) != "function") {
				ary.push(member);
			}
		}
		return ary.join(this.getSeparator());
	}
	catch (e) {
	}
	return "";
}

Utils.prototype.getJavaScriptMethods = function(obj)
{
	if (! __isObject(obj)) {
		return "";
	}
	try {
		var ary = [];
		for (var member in obj) {
			if (typeof(obj[member]) == "function") {
				ary.push(member+"()");
			}
		}
		return ary.join(this.getSeparator());
	}
	catch (e) {
	}
	return "";
}

Utils.prototype.println = function(text)
{
	consoleBuffer.println(text);
}

Utils.prototype.print = function(text)
{
	consoleBuffer.print(text);
}

Utils.prototype.printError = function(text)
{
	consoleBuffer.println("[ERROR] " + text);
}

Utils.prototype.flushConsole = function()
{
	updateConsole();
}

Utils.prototype.setLanguage = function(lang)
{
	this.language = lang.substr(0, 2);
	if (messageTable[this.language]) {
		this.messages = messageTable[this.language];
	}
	var scriptNames = scriptManager.getScriptNames();
	for (i in scriptNames) {
		var name = scriptNames[i];
		var sc = scriptManager.getScript(name).getScriptControl();
		try {
			sc.Run("ASConsole_setLocale")
		}
		catch (e) {}
	}
}

Utils.prototype.getText = function(id)
{
	if (! this.messages[id]) {
		return null;
	}
	return this.messages[id];
}

Utils.prototype.isRubyScriptEnabled = function()
{
	try {
		return this.validateRubyScript();
	}
	catch (e) {}
	return false;
}

Utils.prototype.comInfo = function(obj, member)
{
	try {
		this.validateRubyScript();
		this.scRubyScript.Run("ASConsole_comInfo", obj, member);
	}
	catch (e) {
		this.setError(e);
	}
}

Utils.prototype.comMethods = function(obj)
{
	try {
		this.validateRubyScript();
		this.scRubyScript.Run("ASConsole_comMethods", obj);
	}
	catch (e) {
		this.setError(e);
	}
}

Utils.prototype.comProperties = function(obj)
{
	try {
		this.validateRubyScript();
		this.scRubyScript.Run("ASConsole_comProperties", obj);
	}
	catch (e) {
		this.setError(e);
	}
}

Utils.prototype.validateRubyScript = function()
{
	try {
		if (this.scRubyScript == null) {
			this.scRubyScript = scriptManager.getScript(defaultRubyScript).getScriptControl();
		}
		if (this.scRubyScript != null) {
			return this.scRubyScript.Run("ASConsole_isRubyEnabled");
		}
	}
	catch(e) {}
	throw new Error(utils.getText("COM_FUNCTION_NEED_RUBY"));
}

Utils.prototype.loadText = function(filename, charset)
{
	var path = this.getAbsolutePath(filename);
	if (__isEmpty(charset)) {
		charset = "utf-8";
	}
	var stm = new ActiveXObject("ADODB.Stream");
	stm.Type = adTypeText;
	stm.Charset = charset;
	stm.Open();
	stm.LoadFromFile(path);
	var text = stm.ReadText();
	return text;
}

Utils.prototype.saveText = function(filename, text, charset)
{
	var path = this.getAbsolutePath(filename);
	if (__isEmpty(charset)) {
		charset = "utf-8";
	}
	var stm = new ActiveXObject("ADODB.Stream");
	stm.Type = adTypeText;
	stm.Charset = charset;
	stm.Open();
	stm.WriteText(text);
	stm.SaveToFile(path, adSaveCreateOverWrite);
	stm.Close();
}

Utils.prototype.load = function(path, charset)
{
	if (path.match(/^(https?:|file:)/)) {
		// ignore charset
		objXmlHttp.open('GET', path, false);
		try {
			objXmlHttp.send(null);
		}
		catch (e) {
			this.logError("load() failed.");
			this.logError(formatError(e));
			return null;
		}
		if (path.match(/file:/)) {
			// no support charset
			return objXmlHttp.responseText;
		}
		if (objXmlHttp.status != 200) {
			this.logError("load() failed.");
			this.logError(__formatstr("status code: {0}", objXmlHttp.status));
			return null;
		}
		if (__isEmpty(charset)) {
			return objXmlHttp.responseText;
		}
		return this.binaryToText(objXmlHttp.responseBody, charset);
	}
	return this.loadText(path, charset);
}

Utils.prototype.getAbsolutePath = function(path)
{
	if (path.match(/^[A-Za-z]:\\/)) {
		return path;
	}
	else if (path.match(/^\\\\/)) {
		return path;
	}
	return objFS.BuildPath(this.appDir, path);
}

Utils.prototype.binaryToText = function(buffer, charset)
{
	var stm = new ActiveXObject("ADODB.Stream");
	stm.Type = adTypeBinary;
	stm.Open();
	stm.Write(buffer);
	stm.Position = 0;
	stm.Type = adTypeText;
	stm.Charset = charset;
	return stm.ReadText();
}

Utils.prototype.fromJSON = function(text)
{
	try {
		return JSON.parse(text);
	} catch (e) {}
	return eval('(' + text + ')');
}

Utils.prototype.toJSON = function(obj)
{
	return JSON.stringify(obj, null, 2);
}

Utils.prototype.addHelp = function(cmd, obj)
{
	helpText["user library"][cmd] = obj;
}


var utils = new Utils();


///////////////////////////////////////////////////////////////////////
// Console buffer

function ConsoleBuffer(size)
{
	this.bufferSize = size;
	this.clear();
}

ConsoleBuffer.prototype.println = function(text)
{
	var s = this.currentLine;
	if (text != null) {
		s += __tostr(text);
	}
	// broken charactor replace to '_'
	s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]+/g, "_");
	s = s.replace(/\r*\n/g, "\r\n");
	if (s.charAt(s.length - 1) != "\n") {
		s += "\r\n";
	}
	this.lines.push(s);
	if (this.lines.length > consoleBufferSize) {
		var removeText = this.lines.shift();
		this.buffer = this.buffer.substring(removeText.length);
	}
	this.buffer += s;
	this.currentLine = "";
}

ConsoleBuffer.prototype.print = function(text)
{
	if (__isEmpty(text)) {
		return;
	}
	var s = __tostr(text);
	this.currentLine += s;
	if (this.currentLine.charAt(this.currentLine.length-1) == "\n") {
		this.println();
	}
}

ConsoleBuffer.prototype.clear = function()
{
	this.buffer = "";
	this.lines = [];
	this.currentLine = "";
}

ConsoleBuffer.prototype.flush = function()
{
	if (this.currentLine != "") {
		this.println();
	}
}

ConsoleBuffer.prototype.getBuffer = function()
{
	return this.buffer;
}

var consoleBuffer = new ConsoleBuffer(consoleBufferSize);


///////////////////////////////////////////////////////////////////////

function HistoryBuffer(size)
{
	this.commandList = [];
	this.labelList = [];
	this.idCount = 0;
	this.currentIndex = -1;
	this.bufferSize = size;
}

HistoryBuffer.prototype.restore = function(list)
{
	this.commandList = [];
	this.labelList = [];
	this.idCount = 0;
	for (var i = list.length - 1; i >= 0; i--) {
		this.addText(list[i]);
	}
}

HistoryBuffer.prototype.addText = function(text)
{
	if (__isEmpty(text)) {
		return;
	}
	text = this.getCorrectText(text);
	this.removeRepeatedItem(text);

	this.commandList.unshift(text);
	this.labelList.unshift(this.getLabelText(text));

	if (this.commandList.length > this.bufferSize) {
		this.commandList.length = this.bufferSize;
		this.labelList.length = this.bufferSize;
	}

	this.idCount++;
	this.resetIndex();
}

HistoryBuffer.prototype.getLabel = function(index)
{
	if (index == null) {
		index = this.currentIndex;
	}
	if (index < 0 || index >= this.commandList.length) {
		return "";
	}
	return this.labelList[index];
}

HistoryBuffer.prototype.getText = function(index)
{
	if (index == null) {
		index = this.currentIndex;
	}
	if (index < 0 || index >= this.commandList.length) {
		return "";
	}
	return this.commandList[index];
}

HistoryBuffer.prototype.getCorrectText = function(text)
{
	var identifier = this.getMultilineIdentifier(text);
	if (identifier == null) {
		return text;
	}
	for (var i = 0; i < this.commandList.length; i++) {
		if (identifier == this.getMultilineIdentifier(this.labelList[i])) {
			return this.commandList[i];
		}
	}
	return text;
}

HistoryBuffer.prototype.getLabelText = function(text)
{
	if (__isEmpty(text)) {
		return "";
	}
	if (! isMultiLineText(text)) {
		return text;
	}
	var label = __formatstr("#multiline#{0}  ({1})",
				[__numstr(this.idCount % 1000, 3), this.getTextHeader(text)]);

	return label;
}

HistoryBuffer.prototype.getMultilineIdentifier = function(text)
{
	if (text.match(/^(#multiline#\d+)/)) {
		return RegExp.$1;
	}
	return null;
}

HistoryBuffer.prototype.removeRepeatedItem = function(text)
{
	var index = -1;
	for (var i = 0; i < this.commandList.length; i++) {
		if (text == this.commandList[i]) {
			index = i;
			break;
		}
	}
	if (index != -1) {
		this.commandList.splice(index, 1);
		this.labelList.splice(index, 1);
	}
}

HistoryBuffer.prototype.resetIndex = function()
{
	this.currentIndex = -1;
}

HistoryBuffer.prototype.nextIndex = function()
{
	if (this.currentIndex + 1 < this.commandList.length) {
		this.currentIndex++;
		return true;
	}
	return false;
}

HistoryBuffer.prototype.prevIndex = function()
{
	if (this.currentIndex >= 0) {
		this.currentIndex--;
		return true;
	}
	return false;
}

HistoryBuffer.prototype.getTextHeader = function(text)
{
	var s = text.replace(/\s+/g, " ");
	if (s.length > 40) {
		s = s.substr(0, 40) + " ...";
	}
	return s;
}

HistoryBuffer.prototype.getUsed = function()
{
	return this.commandList.length;
}

HistoryBuffer.prototype.dump = function()
{
	var i;
	for (i = this.labelList.length - 1; i >= 0; i--) {
		utils.println(__formatstr("{0}: {1}", [i+1, this.getLabel(i)]));
	}
}

var historyBuffer = new HistoryBuffer(historyBufferSize);

///////////////////////////////////////////////////////////////////////

function AliasManager()
{
	this.aliasList = {};
}

AliasManager.prototype.add = function(alias, text)
{
	this.aliasList[alias] = text;
}

AliasManager.prototype.get = function(alias)
{
	if (! (alias in this.aliasList)) {
		return null;
	}
	return this.aliasList[alias];
}

AliasManager.prototype.remove = function(alias, text)
{
	delete this.aliasList[alias];
}

AliasManager.prototype.applyText = function(text)
{
	var s = __trim(text);
	if (s.match(/^([A-Za-z_]\w*)(?:|\s+.*)/m) == null) {
		return s;
	}
	var keyword = RegExp.$1;
	var replaceText = this.get(keyword);
	if (replaceText == null) {
		return s;
	}
	return s.replace(keyword, replaceText);
}

AliasManager.prototype.dump = function()
{
	var aliases = [];
	for (alias in this.aliasList) {
		aliases.push(alias);
	}
	aliases.sort();

	for (i in aliases) {
		var s = aliases[i] + " = " + this.get(aliases[i]).replace(/\r*\n/g, " ") + "\n";
		utils.println(s);
	}
}

var aliasManager = new AliasManager();


///////////////////////////////////////////////////////////////////////

function UserSetting()
{
	this.params = {};
}

UserSetting.prototype.restore = function()
{
	var path = objFS.BuildPath(utils.getUserDir(), "user.dat");
	try {
		if (objFS.FileExists(path)) {
			var text = utils.loadText(path, "utf-8");
			this.params = fromJSON(text);
			if (this.params.version != "2.0") {
				throw new Error("bad file format. [user.dat]");
			}
		}
	}
	catch (e) {
		utils.printError("UserSetting.restore() failed.");
		utils.printError(formatError(e));
		utils.logError("UserSetting.restore() failed.");
		utils.logError(formatError(e));
		this.backupUserSetting(path);
	}
	this.defaultParams();
	aliasManager.aliasList = this.params.aliasList;
	historyBuffer.restore(this.params.commandHistory);
}

UserSetting.prototype.defaultParams = function()
{
	this.params.version = "2.0";
	if (!("currentScript" in this.params))    { this.params.currentScript = "JavaScript"; }
	if (!("commandHistory" in this.params))   { this.params.commandHistory = []; }
	if (!("aliasList" in this.params))        { this.params.aliasList = {}; }
}

UserSetting.prototype.save = function()
{
	this.params.currentScript = currentScript.getName();
	this.params.commandHistory = historyBuffer.commandList;

	path = objFS.BuildPath(utils.getUserDir(), "user.dat");
	try {
		var text = toJSON(this.params);
		utils.saveText(path, text, "utf-8");
	}
	catch (e) {
		utils.logError("UserSetting.save() failed.");
		utils.logError(formatError(e));
	}
}

UserSetting.prototype.backupUserSetting = function(path)
{
	if (! objFS.FileExists(path)) {
		return;
	}
	var backupNum = 0;
	var newPath = null;
	for (;;) {
		backupNum += 1;
		var newPath = path + "." + backupNum;
		if (! objFS.FileExists(newPath)) {
			break;
		}
		if (backupNum >= 10) {
			return;
		}
	}
	try {
		objFS.MoveFile(path, newPath);
	}
	catch (e) {
		utils.printError("backupUserSetting() failed.");
		utils.printError(formatError(e));
		utils.logError("backupUserSetting() failed.");
		utils.logError(formatError(e));
	}
}

var userSetting = new UserSetting();

function fromJSON(text)
{
	try {
		return JSON.parse(text);
	} catch (e) {}
	return eval('(' + text + ')');
}

function toJSON(obj)
{
	return JSON.stringify(obj, null, 2);
}


///////////////////////////////////////////////////////////////////////

function isMultiLineText(text)
{
	if (! __isString(text)) {
		return false;
	}
	return (text.indexOf("\n") != -1);
}

function formatError(e)
{
	var s = "";
	if (e == null) {
		s = "(unknown error)";
	}
	//if (e instanceof Error) {
	else if (! __isNull(e.number) && ! __isNull(e.description)) {
		s = "" + (e.number & 0xffff) + ": " + e.description;
	}
	else if (! __isNull(e.description)) {
		s = "" + e.description;
	}
	else {
		s = __tostr(e);
	}
	return s;
}

function formatDate(dt)
{
	return __formatstr("{0}/{1}/{2} {3}:{4}:{5}.{6}",[
		__numstr(dt.getFullYear(), 4),__numstr(dt.getMonth()+1, 2),__numstr(dt.getDate(), 2),
		__numstr(dt.getHours(), 2),__numstr(dt.getMinutes(), 2),__numstr(dt.getSeconds(), 2),
		__numstr(dt.getMilliseconds(), 3)]);
}
