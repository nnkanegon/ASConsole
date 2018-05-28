#-----------------------------------------------------------------------------
# eval helper (PerlScript)
#-----------------------------------------------------------------------------
use Win32::OLE;
use Data::Dumper;

# Host-Script interface
# $utils

$Data::Dumper::Terse = 1;
$Data::Dumper::Indent = 0;

@_result = '';
$_n = 0;

$__buf = '';


#######################################################################
# eval helper

sub ASConsole_execEval {
	my ($_text) = @_;
	my @_r = eval($_text);
	if ($@) {
		$utils->setError($@);
		return '';
	}
	if ($utils->getError('dummy')) {
		return '';
	}
	@_result = @_r;
	if ((defined $_r[0]) && ($_r[0] =~ /^(\d+\.?\d*|\.\d+)$/)) {
		$_n = $_r[0];
	}
	Dumper(@_r);
}


#######################################################################

sub exportVariable {
	my ($name, $value) = @_;
	if (scalar(@_) == 0) {
		$utils->setError("wrong parameter.");
		return undef;
	}
	if (defined($value)) {
		$utils->setVariable($name, $value);
	}
	else {
		$utils->setVariable($name);
	}
	undef
}

sub importVariable {
	my ($name) = @_;
	if (scalar(@_) != 1) {
		$utils->setError("wrong parameter.");
		return undef;
	}
	return $utils->getVariable($name);
}


#######################################################################

sub comInfo {
	my ($obj, $member) = @_;
	if (scalar(@_) == 0) {
		$utils->setError("wrong parameter.");
		return undef;
	}
	if (defined($member)) {
		$utils->comInfo($obj, $member);
	}
	else {
		$utils->comInfo($obj);
	}
	undef
}

sub comMethods {
	my ($obj) = @_;
	if (scalar(@_) == 0) {
		$utils->setError("wrong parameter.");
		return undef;
	}
	$utils->comMethods($obj);
	undef
}

sub comProperties {
	my ($obj) = @_;
	if (scalar(@_) == 0) {
		$utils->setError("wrong parameter.");
		return undef;
	}
	$utils->comProperties($obj);
	undef
}

sub comProps {
	my ($obj) = @_;
	if (scalar(@_) == 0) {
		$utils->setError("wrong parameter.");
		return undef;
	}
	$utils->comProperties($obj);
}

#######################################################################
# internal functions (interface)

sub ASConsole_onActive {
}

sub ASConsole_isEnableStdIO {
	return 1
}

sub ASConsole_initStdBuffer {
	open($_vf, '>', \$__buf);
	select($_vf);
}

sub ASConsole_getStdBuffer {
	my $result = $__buf;
	&ASConsole_initStdBuffer;
	return $result;
}


#######################################################################
# internal functions


#######################################################################

&ASConsole_initStdBuffer();
