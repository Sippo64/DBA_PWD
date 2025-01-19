@ECHO OFF
 SET Password_in_chiaro=La_mela_Buona_123$!

 echo Password_in_chiaro = %Password_in_chiaro%

 FOR /f "tokens=*" %%i in ('cscript //nologo DBA_Pwd.vbs "MyPWD=%Password_in_chiaro%" "ACTION=E" "BaseCode=@m@"') do SET ESITO=%%i
 
 
 SET Password_Criptata=%ESITO%
 
 echo Password Criptata = %ESITO%

 FOR /f "tokens=*" %%i in ('cscript //nologo DBA_Pwd.vbs "MyPWD=%Password_Criptata%" "ACTION=D" "BaseCode=@m@"') do SET ESITO=%%i

 echo Password Decriptata = %ESITO%

pause