'Parameters Es: -> DBA_Crypt_API_DOS.vbs ACTION=E MyPWD=TestEncrypt
'Parameters Es: -> DBA_Crypt_API_DOS.vbs ACTION=D MyPWD=41231339041403372043155328442737360543309384043301138151903643112
'Parameters Es: -> DBA_Crypt_API_DOS.vbs ACTION=E MyPWD=TestEncrypt BaseCode=My 
'Parameters Es: -> DBA_Crypt_API_DOS.vbs ACTION=D MyPWD=31231339041403372643185328342767360543309384043361138181903645114133264349 BaseCode=My

' Questo programma effettua da riga di comando il cript e decript di una password passata come parametro

Dim PWD: PWD = ""
Dim BaseCode: BaseCode = ""
Dim ACTION: ACTION = ""

' Include Library whith Class DBA_Crypt vbs
set objFSO = CreateObject("Scripting.FileSystemObject")
Dim CurrentDirectory: CurrentDirectory = objFSO.GetAbsolutePathName(".")
Path_Lib =  Left(WScript.ScriptFullName, InStr(WScript.ScriptFullName, WScript.ScriptName) - 2)
' msgbox Path_Lib
includeFile Path_Lib & "\DBA_Crypt_Lib.vbs", objFSO

Sub includeFile (fSpec, objFSO)
    dim file, fileData
    On Error Resume Next
    set file = objFSO.openTextFile (fSpec)
    If err Then
    	MsgBox "Errore in " & Wscript.ScriptFullName & vbcr & err.number & " " & err.Description & " " & err.Source & vbcr & fSpec
    End If
    fileData = file.readAll ()
    file.close
    executeGlobal fileData
    set file = nothing
end Sub

If VerifyInputParameter then
	txtOut = "Error Parameter ACTION"
    Set oDBA_CRYPT = New DBA_Crypt
    oDBA_CRYPT.OutPutNumeric = false
     
    If ACTION = "D" Then
    	txtOut = oDBA_CRYPT.DBA_f_DecryptKey(PWD, BaseCode)	
		WScript.Echo txtOut 
		WScript.Quit 0
    End If
    If ACTION = "E" Then
    	txtOut = oDBA_CRYPT.DBA_f_EnCryptKey(PWD, BaseCode)	
		WScript.Echo txtOut 
		WScript.Quit 0
    End If    
	WScript.Echo txtOut 
	WScript.Quit 30
Else
	txtOut = "Error Parameter"
	WScript.Echo txtOut 
	WScript.Quit 50
End If

Function VerifyInputParameter()
	Dim command_line_args
	Dim Element_Arg
	
	VerifyInputParameter = False
	Set command_line_args = WScript.Arguments      
	
	' Verifica esistenza DBA_ROOT_INPUT
	If command_line_args.Count > 0 Then
		For Each Element_Arg In WScript.Arguments
			CurrentArg = UCase(Element_Arg)
			If InStr(CurrentArg, "MYPWD=") > 0 Then
				PWD = mid(Element_Arg, len("MyPWD=")+1)
				VerifyInputParameter = True
			End If
			If InStr(CurrentArg, UCase("BaseCode=")) > 0 Then
				BaseCode = mid(Element_Arg, len("BaseCode=")+1)
				VerifyInputParameter = True
			End If
			If InStr(CurrentArg, "ACTION=") > 0 Then
				ACTION = mid(Element_Arg, len("ACTION=")+1)
				VerifyInputParameter = True
			End If
		Next
	End If
End Function