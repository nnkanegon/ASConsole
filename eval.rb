#-----------------------------------------------------------------------------
# eval helper (RubyScript)
#-----------------------------------------------------------------------------
require 'win32ole'
require 'stringio'

# Host-Script interface
# Utils

if (RUBY_VERSION < '1.9')
	$KCODE = 'NONE'            # default encoding
else
	$__encoding = 'ASCII-8BIT' # encoding of COM I/F (Windows ANSI codepage base?)
	Encoding.default_external = $__encoding
end

$utils = Utils
$__as = self
$__main = eval('self', TOPLEVEL_BINDING)
$__stdBuffer = ''
$__stdio = nil
$__result = nil


#######################################################################
# eval helper

def ASConsole_execEval(text)
	if (RUBY_VERSION >= '1.9')
		text.force_encoding($__encoding)
	end
	begin
		$__result = eval(text, TOPLEVEL_BINDING, 'ASConsole', 1)
	rescue SyntaxError
		Utils.setError($!.inspect)
	rescue
		Utils.setError($!.inspect)
	end
	unless (Utils.getError('dummy').nil?)
		return nil
	end
	if (RUBY_VERSION < '1.9')
		eval('_result = $__result', TOPLEVEL_BINDING)
	else
		eval('_result = $__result'.encode($__encoding), TOPLEVEL_BINDING)
	end
	if ($__result.kind_of?(Numeric))
		eval('_n = $__result', TOPLEVEL_BINDING)
	end
	if ($__result.kind_of?(String))
		$__result
	else
		$__result.inspect
	end
end


#######################################################################

class <<$__main

def exportVariable(name, value)
	if (value.nil?)
		$utils.setVariable(name)
	else
		$utils.setVariable(name, value)
	end
	nil
end

def importVariable(name)
	v = $utils.getVariable(name)
	if (RUBY_VERSION >= '1.9')
		if (v.kind_of?(String))
			v.force_encoding($__encoding)
		end
	end
end

end

#######################################################################

class <<$__main

def comInfo(obj, member=nil)
	$__as.ASConsole_comInfo(obj, member)
end

def comMethods(obj)
	$__as.ASConsole_comMethods(obj)
end

def comProperties(obj)
	$__as.ASConsole_comProperties(obj)
end

def comProps(obj)
	comProperties(obj)
end

end


#######################################################################
# internal functions (interface)

def ASConsole_onActive
end

def ASConsole_setLocale()
	locale = Utils.getLanguage('dummy')
	if (locale == "ja")
		if (RUBY_VERSION < '1.9')
			$KCODE = 'SJIS'
		else
			$__encoding = 'Windows-31J'
			Encoding.default_external = $__encoding
		end
	end
end

def ASConsole_isEnableStdIO()
	true
end

def ASConsole_initStdBuffer()
	if (RUBY_VERSION < '1.9')
		$__stdBuffer = ''
	else
		$__stdBuffer = ''.encode($__encoding)
	end
	$__stdio = StringIO.new($__stdBuffer)
	$stderr = $stdout = $__stdio
end

def ASConsole_getStdBuffer()
	result = $__stdBuffer
	ASConsole_initStdBuffer()
	result
end

def ASConsole_isRubyEnabled()
	true
end

# COM description
def ASConsole_comInfo(obj, member=nil)
	begin
		if (! $__main.__isCOMObject(obj))
			# type error ?
			#return obj.class.to_s
			Utils.println(obj.class.to_s)
			return nil
		end
		unless (member.nil?)
			if (! member.kind_of?(String))
				raise TypeError.new("wrong parameter")
			end
			Utils.println($__main.__comMemberInfo(obj, member))
		else
			Utils.println($__main.__comObjectInfo(obj))
		end
	rescue
		Utils.setError($!.inspect);
	end
	nil
end

# COM method list
def ASConsole_comMethods(obj)
	begin
		unless ($__main.__isCOMObject(obj))
			raise TypeError.new("wrong object.")
		end
		Utils.println($__main.__comGetMembers(obj, 'FUNC'))
	rescue
		Utils.setError($!.inspect)
	end
	nil
end

# COM property list
def ASConsole_comProperties(obj)
	begin
		unless ($__main.__isCOMObject(obj))
			raise TypeError.new("wrong object.")
		end
		Utils.println($__main.__comGetMembers(obj, 'PROPERTY'))
	rescue
		Utils.setError($!.inspect)
	end
	nil
end


#######################################################################
# internal functions

class <<$__main

# COM description (COM object)
def __comObjectInfo(obj)
	unless (__isCOMObject(obj))
		return "wrong object. (not COM object)"
	end
	return <<"EOS"
[guid(#{obj.ole_obj_help.guid}),
helpstring(\"#{obj.ole_obj_help.helpstring}\")]
#{obj.ole_obj_help.ole_type} #{obj.ole_obj_help.name}
EOS
end

# COM description (COM member)
def __comMemberInfo(obj, member)
	unless (__isCOMObject(obj))
		return "wrong object. (not COM object)"
	end
	unless (member.kind_of?(String))
		raise TypeError.new("wrong parameter")
	end
	ss = ""
	method = __comGetMember(obj, member, 'FUNC')
	if (method != nil)
		ss += sprintf("[id(0x%08x), method, helpstring(\"%s\")]\n", method.dispid.to_s, method.helpstring)
		ss += method.return_type + " " + method.name + "(" + __comFormatParams(method) + ")\n"
	end
	method = __comGetMember(obj, member, 'PROPERTYGET')
	if (method != nil)
		ss += "\n" unless (ss.empty?)
		ss += sprintf("[id(0x%08x), propget, helpstring(\"%s\")]\n", method.dispid.to_s, method.helpstring)
		ss += method.return_type + " " + method.name + "(" + __comFormatParams(method) + ")\n"
	end
	method = __comGetMember(obj, member, "PROPERTYPUT")
	if (method != nil)
		ss += "\n" unless (ss.empty?)
		ss += sprintf("[id(0x%08x), propput, helpstring(\"%s\")]\n", method.dispid.to_s, method.helpstring)
		ss += method.return_type + " " + method.name + "(" + __comFormatParams(method) + ")\n"
	end
	ss
end

# format text of comInfo
def __comFormatParams(method)
	ss = ""
	if (method.params.size > 0)
		tmp = ""
		method.params.each do |p|
			unless (tmp.empty?)
				tmp += ",\n"
			end
			a = []
			a.push("in") if (p.input?)
			a.push("out") if (p.output?)
			a.push("optional") if (p.optional?)
			tmp += "    " + "[" + a.join(', ') + "] " + p.ole_type + " " + p.name
		end
		ss += "\n" + tmp
	end
	ss
end

# member
# kind = FUNC, PROPERTYGET, PROPERTYPUT
def __comGetMember(obj, name, kind)
	obj.ole_methods().each do |o|
		next unless (o.kind_of?(WIN32OLE_METHOD))
		next if (o.name.downcase != name.downcase)
		next if (o.invoke_kind != kind)
		return o
	end
	nil
end

# member list
# kind = FUNC, PROPERTY
def __comGetMembers(obj, kind)
	a = []
	obj.ole_methods().each do |o|
		next unless (o.name)
		next unless (o.kind_of?(WIN32OLE_METHOD))
		case o.invoke_kind
		when 'FUNC'
			if (kind == 'FUNC')
				a.push(o.name + '()')
			end
		when 'PROPERTYGET'
			if (kind == 'PROPERTY')
				a.push(o.name)
			end
		when 'PROPERTYPUT'
			if (kind == 'PROPERTY')
				a.push(o.name + '=')
			end
		end
	end
	a.join($utils.getSeparator('dummy'))
end

def __isCOMObject(obj)
	return false if (obj.nil?)
	return false unless (obj.kind_of?(WIN32OLE))
	return false unless (defined?(obj.ole_obj_help))
	true
end

end


#######################################################################

ASConsole_initStdBuffer()
