//----------------------------------------------------------------------------
// help
//----------------------------------------------------------------------------

helpText = {
"system command":{
	"#alias":["#alias [<keyword> <text>]","set command alias"],
	"#clear":["#clear","clear console"],
	"#copy":["#copy","copy console text to clipborard"],
	"#cs":["#cs [<scriptname>]","change script engine"],
	"#edit":["#edit","edit console text"],
	"#editor":["#editor [<editor>]","set text editor"],
	"#enterexec":["#enterexec [true|false]","set operation of enter key in multi-line mode"],
	"#exit":["#exit","exit application"],
	"#unalias":["#unalias","delete command alias"],
	"#version":["#version","show version"],
	"#!<command>":["#!<command>","exec system application"],
	"##":["##[<n>]","show/exec command history"]
},
"utility command":{
	"comInfo":["comInfo(obj [,name])","show information of COM object"],
	"comMethods":["comMethods(obj)","show list of function of COM object"],
	"comProperties":["comProperties(obj)","show list of property of COM object"],
	"exportVariable":["exportVariable(<name>, <object>)","export variable"],
	"importVariable":["importVariable(<name>)","import variable"],
	"bin":["bin([n,...])","convert radix"],
	"oct":["oct([n,...])","convert radix"],
	"dec":["dec([n,...])","convert radix"],
	"hex":["hex([n,...])","convert radix"],
	"__attach":["__attach(<browserobject>, <initscript>)","attach to web browser"],
	"__detach":["__detach()","detach web browser"],
	"__createIE":["__createIE()","create IE instance"],
	"__internalIE":["__internalIE()","create internal IE object"]
},
"user library":{
}
};


///////////////////////////////////////////////////////////////////////
// HelpManager

function HelpManager()
{
}

HelpManager.prototype.help = function(cmd)
{
	if (__isEmpty(cmd)) {
		this.helpSummary();
	}
	else {
		this.helpCommand(cmd);
	}
}

HelpManager.prototype.helpSummary = function()
{
	for (category in helpText) {
		consoleBuffer.println(__formatstr("[{0}]", [category]));
		for (cmd in helpText[category]) {
			consoleBuffer.println(__formatstr("{0}", [cmd]));
		}
		consoleBuffer.println();
	}
}

HelpManager.prototype.helpCommand = function(cmd)
{
	var cmdhelp = null;
	for (category in helpText) {
		for (cmd2 in helpText[category]) {
			if(cmd2.toLowerCase().indexOf(cmd.toLowerCase()) === 0){
				cmdhelp = helpText[category][cmd2];
				break;
			}
		}
		if (cmdhelp != null) {
			break;
		}
	}
	if (cmdhelp != null) {
		consoleBuffer.println(cmdhelp[0]);
		consoleBuffer.println(cmdhelp[1]);
	}
}

helpManager = new HelpManager();
