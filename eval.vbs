'-----------------------------------------------------------------------------
' eval helper (VBScript)
'-----------------------------------------------------------------------------
'Option Explicit

' Host-Script interface
' Utils

Dim AS_result
Dim AS_n

AS_result = ""
AS_n = 0


'-----------------------------------------------------------------------------
' eval helper

Function ASConsole_execEval(text)
	On Error Resume Next
	If Left(text, 1) = ":" Then
		text = Mid(text, 2)
		ASConsole_execEval = Eval(text)
	Else
		Execute(text)
		ASConsole_execEval = Empty
	End If
	If Err.Number <> 0 Then
		Utils.setError(Utils.createError(Err.Number, Err.Description))
	End If
	Err.Clear
	On Error GoTo 0
	If IsNull(Utils.getError()) Then
		AS_result = ASConsole_execEval
		If IsNumeric(ASConsole_execEval) Then
			AS_n = ASConsole_execEval
		End If
	End If
End Function


'-----------------------------------------------------------------------------

Sub exportVariable(name, value)
	Utils.setVariable name, value
End Sub

Function importVariable(name)
	If IsObject(Utils.getVariable(name)) Then
		Set importVariable = Utils.getVariable(name)
	Else
		importVariable = Utils.getVariable(name)
	End If
End Function


'-----------------------------------------------------------------------------

Sub print(s)
	Utils.print(s)
End Sub

Sub println(s)
	Utils.println(s)
End Sub


'-----------------------------------------------------------------------------

Sub comInfo(obj, member)
	Utils.comInfo obj, member
End Sub

Sub comMethods(obj)
	Utils.comMethods obj
End Sub

Sub comProperties(obj)
	Utils.comProperties obj
End Sub

Sub comProps(obj)
	Utils.comProperties obj
End Sub


'-----------------------------------------------------------------------------
' internal functions (interface)

Sub ASConsole_onActive()
End Sub


'-----------------------------------------------------------------------------
' internal functions

