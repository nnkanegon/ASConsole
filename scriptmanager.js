//----------------------------------------------------------------------------
// script manager
//----------------------------------------------------------------------------

var JavaScript = "JavaScript";
var VBScript = "VBScript";
var RubyScript = "RubyScript";
var RubyScript18 = "RubyScript.1";
var RubyScript24 = "RubyScript.2.4";
var PerlScript = "PerlScript";
var PythonScript = "Python";
var PHPScript = "PHPScript";

var defaultScript = JavaScript;
var defaultRubyScript = RubyScript;		// use comXXX function


var scriptList = [
	[JavaScript,    ["common.js", "eval.js", "browserHelper.js"], "js"],
	[VBScript,      ["eval.vbs"], "vbs"],
	[RubyScript,    ["eval.rb"], "rb"],
//	[RubyScript18,  ["eval.rb"], "rb18"],
//	[RubyScript24,  ["eval.rb"], "rb24"],
	[PerlScript,    ["eval.pl"], "pl"],
	[PythonScript,  ["eval.py"], "py"],
	[PHPScript,     ["eval.php"], "php"]];

var scriptAlias = [
	["javascript",   JavaScript],
	["jscript",      JavaScript],
	["js",           JavaScript],
	["vbscript",     VBScript],
	["vbs",          VBScript],
	["rubyscript",   RubyScript],
	["ruby",         RubyScript],
	["rubyscript18", RubyScript18],
	["ruby18",       RubyScript18],
	["rubyscript24", RubyScript24],
	["ruby24",       RubyScript24],
	["perlscript",   PerlScript],
	["perl",         PerlScript],
	["python",       PythonScript],
	["py",           PythonScript],
	["phpscript",    PHPScript],
	["php",          PHPScript]];


///////////////////////////////////////////////////////////////////////
// ScriptManager

function ScriptManager()
{
	this.scriptNames = [];
	this.scripts = {};
	this.alias = {};
}

ScriptManager.prototype.setAlias = function(keyword, scriptName)
{
	this.alias[keyword] = scriptName;
}

ScriptManager.prototype.getAlias = function(keyword)
{
	if (! (keyword in this.alias)) {
		return null;
	}
	return this.alias[keyword];
}

ScriptManager.prototype.setScript = function(script)
{
	var name = script.getName();
	this.scriptNames.push(name);
	this.scripts[name] = script;
}

ScriptManager.prototype.getScript = function(name)
{
	var nm = this.getAlias(name);
	if (nm != null) {
		name = nm;
	}
	if (! (name in this.scripts)) {
		return null;
	}
	return this.scripts[name];
}

ScriptManager.prototype.getScriptNames = function()
{
	return this.scriptNames;
}

var scriptManager = new ScriptManager();

///////////////////////////////////////////////////////////////////////
// Script

function Script(name, scriptControl, shortName)
{
	this.name = name;
	this.shortName = shortName;
	this.scriptControl = scriptControl;
	this.enableStdIO = false;
}

Script.prototype.getName = function()
{
	return this.name;
}

Script.prototype.getShortName = function()
{
	return this.shortName;
}

Script.prototype.getScriptControl = function()
{
	return this.scriptControl;
}

Script.prototype.eval = function(text)
{
	var result = null;
	try {
		this.initStdBuffer();
		utils.clearError();
		utils.clearPostAction();
		if (text.indexOf(";;") == 0) {
			result = this.scriptControl.Eval(text.substr(2));
		}
		else if (text.indexOf(";") == 0) {
			this.scriptControl.ExecuteStatement(text.substr(1));
		}
		else if (text.indexOf("<<") == 0) {
			var s = __trim(text.substr(2));
			if (! s.match(/^([^,]+?)(?:,(.*))?$/)) {
				throw new Error("ASConsole: wrong parameter");
			}
			var path = __trim(RegExp.$1);
			var charset = __trim(RegExp.$2);
			var code = utils.load(path, charset);
			this.addCode(code);
		}
		else if (text.indexOf("<") == 0) {
			this.addCode(text.substr(1));
		}
		else {
			result = this.scriptControl.Run("ASConsole_execEval", text);
		}
		var err = this.scriptControl.Error;
		if (! __isNull(err.number)) {
			if ((err.number & 0xffff) != 0) {
				throw err;
			}
		}
		if (! __isEmpty(utils.getPostAction())) {
			eval(utils.getPostAction());
		}
	}
	catch(e) {
		utils.setError(e);
	}
	var err = utils.getError();
	if (err != null) {
		return {status: "error", result: "[ERROR] " + formatError(err)}
	}
	else {
		this.printStdBuffer();
		if (typeof(result) == "undefined") {
			return {status: "success", result: ""}
		}
		return {status: "success", result: result}
	}
}

Script.prototype.addCode = function(text)
{
	this.scriptControl.AddCode(text);
	if (! __isNull(this.scriptControl.Error.number)) {
		if ((this.scriptControl.Error.number & 0xffff) != 0) {
			throw this.scriptControl.Error;
		}
	}
}

Script.prototype.addCodeFromText = function(filename, charset)
{
	var text = utils.loadText(filename, charset);
	this.addCode(text);
}

Script.prototype.initStdBuffer = function()
{
	if (! this.enableStdIO) {
		return;
	}
	try {
		this.scriptControl.Run("ASConsole_initStdBuffer");
	}
	catch(e) {}
}

Script.prototype.printStdBuffer = function()
{
	if (! this.enableStdIO) {
		return;
	}
	var result = "";
	try {
		utils.clearError();
		result = this.scriptControl.Run("ASConsole_getStdBuffer");
	}
	catch(e) {
		utils.setError(e);
	}
	var err = utils.getError();
	if (err != null) {
		utils.printError(formatError(err));
	}
	else {
		var s = __tostr(result);
		if (s != "") {
			utils.println(s);
		}
	}
}

Script.prototype.onActive = function()
{
	try {
		this.scriptControl.Run("ASConsole_onActive");
	}
	catch(e) {}
	this.printStdBuffer();
}

///////////////////////////////////////////////////////////////////////

function getScriptControl(lang, scriptFiles)
{
	try {
		var sc = new ActiveXObject("ScriptControl");
		try {
			sc.Language = lang;
		}
		catch (e) {
			// script engine not found.
			utils.logTrace(formatError(e));
			return null;
		}
		sc.AddObject("utils", utils);
		for (i in scriptFiles) {
			var path = objFS.BuildPath(utils.getAppDir(), scriptFiles[i]);
			var text = utils.loadText(path, "utf-8");
			sc.AddCode(text);
			if (! __isNull(sc.Error.number)) {
				if ((sc.Error.number & 0xffff) != 0) {
					throw sc.Error;
				}
			}
		}
		return sc;
	}
	catch (e) {
		utils.logError("failed to script initialize : "+lang);
		utils.logError(formatError(e));
	}
	return null;
}

function initScripts()
{
	for (i in scriptList) {
		var elm = scriptList[i];
		var name = elm[0];
		var scriptFiles = elm[1];
		var scriptShortName = elm[2];
		var sc = getScriptControl(name, scriptFiles);
		if (sc != null) {
			var script = new Script(name, sc, scriptShortName);
			try {
				script.enableStdIO = sc.Run("ASConsole_isEnableStdIO");
				if (script.enableStdIO == 1) script.enableStdIO = true;
			}
			catch(e) {}
			scriptManager.setScript(script);
		}
	}
	for (i in scriptAlias) {
		var elm = scriptAlias[i];
		var alias = elm[0];
		var scriptName = elm[1];
		scriptManager.setAlias(alias, scriptName);
	}
}

function initUserLib()
{
	// initialize user scripts  (javascript only)
	// target: <appdir>/lib/*.js
	var sc = scriptManager.getScript(JavaScript).scriptControl;
	var userDir = objFS.BuildPath(utils.getAppDir(), "lib");
	var objdir = objFS.GetFolder(userDir);
	var fc = new Enumerator(objdir.files);
	for (; ! fc.atEnd(); fc.moveNext()) {
		var path = __tostr(fc.item());
		var file = fc.item().name;
		if (path.match(/\.js$/) == null) {
			continue;
		}
		try {
			sc.AddCode(utils.loadText(path, "utf-8"));
			if (! __isNull(sc.Error.number)) {
				if ((sc.Error.number & 0xffff) != 0) {
					throw sc.Error;
				}
			}
			utils.println("loaded: "+file);
		}
		catch (e) {
			utils.println("failed to load user script: "+file);
			utils.println(formatError(e));
			utils.logError("failed to load user script: "+file);
			utils.logError(formatError(e));
		}
	}
	utils.println();
}


///////////////////////////////////////////////////////////////////////

initScripts();

