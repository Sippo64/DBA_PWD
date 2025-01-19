' Questa libresia contiene la logica per criptare e decriptare password con o senza MasterKey
Class DBA_Crypt
	Dim m_arrDirty 
	Dim m_DirtyLastCharPosition 
	Dim m_OutPutNumeric
	Dim m_DirtyValue

	Private Sub Class_Initialize()
	    m_DirtyLastCharPosition = 1
	    m_OutPutNumeric = True
	    m_arrDirty = 12345678
		m_DirtyValue = 1
	End Sub
	
 	Public Property Let OutPutNumeric (value)
 		If Value = True Then
 			m_OutPutNumeric = True
 		Else
 			m_OutPutNumeric = False
 		End If
	End Property
		
	Sub Sleep(sec)
	    On Error Resume Next
	    TimerToLoopEnd = Timer + sec / 1000
	    Err.Clear
	    Do While TimerToLoopEnd > Timer
	        DoEvents
	    Loop
	End Sub
	
	Public Function DBA_f_DecryptKeyCode(KeyEncrypted, Pwd)
	    On Error Resume Next
	    DBA_f_DecryptKeyCode = "Pwd for Decrypt KeyCode not valid"
	       
	    If IsPwdOK(KeyEncrypted, Pwd) Then
	        DBA_f_DecryptKeyCode = m_DBA_f_Decrypt(KeyEncrypted, 1)
	    End If
	End Function
	
	Public Function DBA_f_EnCryptKey(Pwd, Codekey)
	    On Error Resume Next
	    DBA_f_EnCryptKey = m_DBA_f_EnCrypt(Pwd) & m_DBA_f_GetPoint() & m_DBA_f_EnCrypt(Codekey)
	    Sleep 5
	End Function
	
	Public Function DBA_f_DecryptKey(KeyEncrypted, Codekey)
	    On Error Resume Next
	    DBA_f_DecryptKey = "Key for Decrypt password not valid"
	       
	    If IsCodeKeyOK(KeyEncrypted, Codekey) Then
	        DBA_f_DecryptKey = m_DBA_f_Decrypt(KeyEncrypted, 0)
	    End If
	    
	End Function
	
	Public Function IsCodeKeyOK(KeyEncrypted, Codekey)
	    CodekeyDecrypted = m_DBA_f_Decrypt(KeyEncrypted, 1)
	    If Codekey = CodekeyDecrypted Then
	        IsCodeKeyOK = True
	    Else
	        IsCodeKeyOK = False
	    End If
	End Function
	
	Public Function IsPwdOK(KeyEncrypted, Pwd)
	    On Error Resume Next
	    PwdDecrypted = m_DBA_f_Decrypt(KeyEncrypted, 0)
	    If Pwd = PwdDecrypted Then
	        IsPwdOK = True
	    Else
	        IsPwdOK = False
	    End If
	End Function
	
	Private Function m_DBA_f_GetPoint()
	    m_DBA_f_GetPoint = "190364"
	End Function
	
	Private Function m_GetDirty()
	    On Error Resume Next
	    'Dim Randomx As Random
	    'Dim Randomx_Next
	
	    m_arrDirty = Replace(Timer, ",", "") & Replace(Timer, ",", "")
	    
	    If m_DirtyLastCharPosition > Len(m_arrDirty) Then
	        m_arrDirty = Replace(Timer, ",", "") & Replace(Timer, ",", "")
	        m_DirtyLastCharPosition = 1
	    Else
	        m_DirtyLastCharPosition = m_DirtyLastCharPosition + 1
	    End If
	  
	    m_GetDirty = Mid(m_arrDirty, m_DirtyLastCharPosition, 1)
	    
	    If Len(m_GetDirty) = 0 or m_GetDirty = "," or m_GetDirty = "." Then
	        m_GetDirty = "9"
	    End If
		
		'm_DirtyValue = m_DirtyValue + 1
		'if m_DirtyValue > 10 Then
		'	m_DirtyValue = 1
		'End If
		'm_GetDirty = m_DirtyValue
	End Function
	
	Private Function m_DBA_f_EnCrypt(Pwd)
	    On Error Resume Next
	    Dim AddSignal
	    Dim k
	    Dim x
	    Dim i
	    Dim iascii
	    Dim imaskedascii
	    Dim cdirty
	    
	    k = ""
	    x = Pwd
	    i = 1
	    
	    Do While (i <= Len(Pwd))
	        AddSignal = 311
	        If (i Mod 2) Then
	            k = k & i
	            AddSignal = 254
	        End If
	        
	        iascii = Asc(Mid(Pwd, i, 1))
	        
	        imaskedascii = iascii + AddSignal + i
	        
	        cdirty = m_GetDirty()
	        
	        k = k & imaskedascii & cdirty
	        'msgbox Mid(Pwd, i, 1) & vbcr & _
			'	    "i=" & i & vbcr & _
			'	   "k=" & k & vbcr & _
			'	   "iascii=" & iascii & vbcr & _
			'	   "cdirty=" & cdirty & vbcr & _
			'	   "imaskedascii=" & imaskedascii & vbcr & _
			'	   "AddSignal=" & AddSignal 
	        i = i + 1
	    Loop
	    
		x = m_GetDirty() & (112 + Len(Pwd)) & k
	    		
	    If m_OutPutNumeric = False then
	    	OutPwd = m_DBA_f_ToChar(x)
	    Else
	    	OutPwd = x
	    End if
    
	    m_DBA_f_EnCrypt = OutPwd
	    
	End Function
	
	Private Function m_DBA_f_Decrypt(KeyEncrypted, xQuote)
	    On Error Resume Next
	    Dim FoundMark
	    Dim KeyToDecrypt
	    Dim LenFor
	    Dim AddSignal
	    Dim i, y, p, pp
	    Dim iEncryptedChar
	    Dim iCodeChar
	    Dim cCodeChar
	    
	    FoundMark = InStr(KeyEncrypted, m_DBA_f_GetPoint())
	    
	    If FoundMark = 0 Then
	        m_DBA_f_Decrypt = ""
	        Exit Function
	    End If
	    
	    If xQuote = 0 Then
	        KeyToDecrypt = Left(KeyEncrypted, FoundMark - 1)
	    Else
	        KeyToDecrypt = Mid(KeyEncrypted, FoundMark + Len(m_DBA_f_GetPoint()), Len(KeyEncrypted))
	    End If
	
	   	KeyToDecrypt = m_DBA_f_ToNumber(KeyToDecrypt)
	   	
	    KeyToDecrypt = Right(KeyToDecrypt, Len(KeyToDecrypt) - 1)   

	    y = 1
	    
	    LenFor = Mid(KeyToDecrypt, y, 3) - 112
	    p = 1
	    pp = ""
	    y = 4
	    i = 1
	    Do While (i <= LenFor)
	        AddSignal = 311
	        If i Mod 2 Then
	            y = y + Len(i)
	            AddSignal = 254
	        End If
	        
	        iEncryptedChar = Mid(KeyToDecrypt, y, Len(AddSignal))
	        iCodeChar = iEncryptedChar - AddSignal - p
	        cCodeChar = Chr(iCodeChar)
	        pp = pp & cCodeChar
	        y = y + 4
	        p = p + 1
	        
	        i = i + 1
	    Loop
	    	    
	    m_DBA_f_Decrypt = pp
	
	End Function

	Private Function m_DBA_f_ToChar(KeyEncrypted)
		Dim Pwd: Pwd = KeyEncrypted
		
		LenStr= Len(Pwd)
		y = 1
		MyPwd = ""
		
	    Do While (y <= Len(Pwd))
	    	MySet = Mid(Pwd, y, 2)
	        If MySet > 10 And MySet < 62 Then '75 K /up to/ 126 ~
	        	NewMySet = MySet+65
	        	' exclude \ ^ ' | [DEL]
	        	If NewMySet = 92 Or NewMySet = 94 Or NewMySet = 96 Or NewMySet = 124 Or NewMySet = 127 Then
	        		MyPwd = MyPwd & MySet
	        	Else
		        	MyPwd = MyPwd & chr(NewMySet)	        	
	        	End If
	        Else	        	
	        	MyPwd = MyPwd & MySet
	        End If
	        y = y + 2
	    Loop
	     
		m_DBA_f_ToChar = MyPwd
	End Function
	
	Private Function m_DBA_f_ToNumber(KeyEncrypted)
		Dim Pwd: Pwd = KeyEncrypted
		
		LenStr= Len(Pwd)
		y = 1
		MyPwd = ""
		
	    Do While (y <= Len(Pwd))
	    	MyChar = Mid(Pwd, y, 1)
	    	
	    	MyAsci = Asc(MyChar)
	    	
	        'If MyAsci < 48 Or MyAsci > 57 Then
			If MyAsci < 48 Or MyAsci > 57 Then
	        	t = (MyAsci - 65)
	        	MyPwd = MyPwd & t
	        Else	        	
	        	MyPwd = MyPwd & MyChar
	        End If
	        y = y + 1
	    Loop
	     
		m_DBA_f_ToNumber = MyPwd
	End Function
		
End Class
