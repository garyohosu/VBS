
msgbox(OpenFileDialog("ファイルを開く"))

function OpenFileDialog(title)

    Dim obj, filename
    Set obj = CreateObject("Excel.Application")
    filename = obj.GetOpenFilename("ALL File,*.*",1,title)
    obj.Quit
    Set obj = Nothing
    If filename <> False Then
          OpenFileDialog = filename
    End If

end function


