'Log出力クラス
'
'DATE       VER  NAME    COMMENT
'2016/06/22 0.00 HANTANI 新規作成

Option Explicit

Class ClsLog

    Private m_LogFileName
    Private m_Enable


    Public Property Get LogFileName
        LogFileName = m_LogFileName
    End Property

    Public Property Let SerialNumber(vData)
        m_LogFileName = vData
    End Property

    Public Property Get Enable
        Enable = m_Enable
    End Property

    Public Property Let Enable(vData)
        m_Enable = vData
    End Property

	function apppath
	    dim fso
	    set fso = createObject("Scripting.FileSystemObject")
	    apppath = fso.getParentFolderName(WScript.ScriptFullName)
	end function

	sub logPrint(s)

		if m_LogFileName<>"" and Enable = True then

			dim objFsoWR
			dim objFileWR
			Set objFsoWR = CreateObject("Scripting.FileSystemObject")
			Set objFileWR = objFsoWR.OpenTextFile(m_LogFileName, 8, True)

			If Err.Number > 0 Then
			    WScript.Echo "Open Error"
			Else
				objFileWR.WriteLine s
			End If

			objFileWR.Close
			Set objFileWR = Nothing
			Set objFsoWR = Nothing
		end if
	end sub

	sub logPrintLn(s)
		logPrint(s & vbcrlf)
	end sub


    Private Sub Class_Initialize()
		m_LogFileName = apppath "\Log.Log" 
		Enable = True
    End Sub

    Private Sub Class_Terminate()
        'MsgBox ("終了")
    End Sub

End Class
