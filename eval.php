#-----------------------------------------------------------------------------
# eval helper (PHPScript)
#-----------------------------------------------------------------------------

# Host-Script interface
# $utils

$GLOBALS["user_variables"] = array();


#######################################################################
# eval helper

function ASConsole_execEval($__text)
{
	# special convert
	if (ereg('^[0-9]', $__text) || ereg('^[$][A-Za-z_][A-Za-z0-9_]*$', $__text)) {
		$__text = "print($__text);";
	}
	else {
		$__text = ereg_replace('^p +([^(].*)$', 'p(\1);', $__text);
	}

	extract($GLOBALS["user_variables"]);

	$__result = eval($__text);

	$__vars = get_defined_vars();
	$__keys = array_keys($__vars);
	for ($__i = 0; $__i < count($__keys); $__i++) {
		if (ereg('^__', $__keys[$__i])) {
			unset($__vars[$__keys[$__i]]);
		}
	}
	$GLOBALS["user_variables"] = $__vars;

	return $__result;
}

#######################################################################

function exportVariable($name, $value)
{
	global $utils;
	$utils->setVariable($name, $value);
}

function importVariable($name)
{
	global $utils;
	return $utils->getVariable($name);
}


#######################################################################

function p()
{
	for($i = 0; $i < func_num_args(); $i++) {
		var_dump(func_get_arg($i));
	}
}


#######################################################################

function comInfo($obj, $member=NULL)
{
	global $utils;
	$utils->comInfo($obj, $member);
}

function comMethods($obj)
{
	global $utils;
	$utils->comMethods($obj);
}

function comProperties($obj)
{
	global $utils;
	$utils->comProperties($obj);
}

function comProps($obj)
{
	global $utils;
	$utils->comProperties($obj);
}


#######################################################################
# internal functions (interface)

function ASConsole_onActive()
{
}

function ASConsole_setLocale()
{
}

function ASConsole_isEnableStdIO()
{
	return true;
}

function ASConsole_initStdBuffer()
{
	ob_clean();
}

function ASConsole_getStdBuffer()
{
	$buf = ob_get_contents();
	ASConsole_initStdBuffer();
	return $buf;
}


#######################################################################
# internal functions

function __error_handler($err_code, $err_string, $filename, $line, $symbols)
{
	$utils->setError($utils->createError($err_code, $err_string));
}


#######################################################################

#error_reporting(E_ALL & ~E_NOTICE);
#set_error_handler('__error_handler');

ob_start();
