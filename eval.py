#-----------------------------------------------------------------------------
# eval helper (Python)
#-----------------------------------------------------------------------------
import sys
import re
import codecs
import StringIO

# Host-Script interface
# utils

_result = ""
_n = 0

__stdBuffer = None


#######################################################################
# eval helper

def ASConsole_execEval(_text):
	global _result, _n
	_r = None
	try:
		if _text[0:1] == ":":
			_r = eval(_text[1:], globals())
		else:
			exec(_text, globals())
			return None
	except:
		utils.setError(repr(sys.exc_info()))
		return None
	if utils.getError() is not None:
		return None
	_result = _r
	if isinstance(_r, (int, long, float)):
		_n = _r
	if isinstance(_r, unicode):
		return _r
	elif isinstance(_r, str):
		return _r.decode('utf-8')
	else:
		return repr(_r)


#######################################################################

def exportVariable(name, value):
	utils.setVariable(name, value)

def importVariable(name):
	return utils.getVariable(name)


#######################################################################

def comInfo(obj, member=None):
	utils.comInfo(obj, member)

def comMethods(obj):
	utils.comMethods(obj)

def comProperties(obj):
	utils.comProperties(obj)

def comProps(obj):
	utils.comProperties(obj)


#######################################################################
# internal functions (interface)

def ASConsole_onActive():
	pass

def ASConsole_setLocale():
	pass

def ASConsole_isEnableStdIO():
	return True

def ASConsole_initStdBuffer():
	global __stdBuffer
	__stdBuffer = StringIO.StringIO()
	sys.stdout = __stdBuffer
	sys.stderr = __stdBuffer

def ASConsole_getStdBuffer():
	global __stdBuffer
	try:
		buf = __stdBuffer
		ASConsole_initStdBuffer()
		result = buf.getvalue()
		buf.close()
		return __getUnicodeText(result)
	except:
		utils.setError(repr(sys.exc_info()))
		return None


#######################################################################
# internal functions

# temporary step
def __getUnicodeText(s):
	try: return s.decode('utf-8')
	except: pass
	try: return s.decode('shift_jis')
	except: pass
	return s


#######################################################################

ASConsole_initStdBuffer()
