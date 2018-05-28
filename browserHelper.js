//----------------------------------------------------------------------------
// browser attach (for jQuery/prototype.js)
//----------------------------------------------------------------------------

function __attach(doc, script)
{
	var internalDocument = false;
	if (doc == null) {
		doc = __internalIE();
		internalDocument = true;
	}
	doc.parentWindow.__asc_self = this;
	doc.parentWindow.__asc_utils = utils;
	doc.parentWindow.__detach = function() { __detach(); };
	var enableScript = false;
	try {
		enableScript = (doc.parentWindow.eval("1") == 1);
	}
	catch (e) {}
	if (! enableScript) {
		var sc = doc.createElement("script");
		sc.Type = "text/javascript";
		doc.body.appendChild(sc);
	}
	doc.parentWindow.eval("__asc_remoteEval=function(__exp){try{return eval(__exp);}catch(__e){__asc_utils.setError(__e);}}");
	__remoteEval = doc.parentWindow.__asc_remoteEval;
	if (script != null) {
		doc.parentWindow.eval(script);
	}
	utils.jsBrowserAttached = true;
	try {
		utils.jsBrowserTitle = doc.title;
	}
	catch (e) {}
	// for jQuery
	try {
		if (internalDocument) {
			if (typeof(doc.parentWindow.jQuery.support.cors) != "undefined") {
				doc.parentWindow.jQuery.support.cors = true;
			}
		}
	}
	catch(e){}
}

function __detach()
{
	__remoteEval = null;
	utils.jsBrowserAttached = false;
	utils.jsBrowserTitle = null;
}

function __createIE()
{
	var ie = new ActiveXObject('InternetExplorer.Application');
	ie.visible = true;
	ie.gohome();
	return ie;
}

function __internalIE()
{
	var doc = new ActiveXObject("htmlfile");
	doc.write('<!DOCTYPE html><html><head><meta charset="utf-8"><meta http-equiv="X-UA-Compatible" content="IE=9"><title>internal</title><script></script></head><body></body></html>');
	return doc;
}
