Option Explicit
'===========================================t
' Wish/Washフォルダ生成マクロ
'
'改定履歴
'DATE       VER  NAME    COMMENT
'2016/08/18 0.00 HANTANI 新規作成
'===========================================t

call main

sub main
	dim YD
	dim Ret
	dim Kisyu
	dim Mukesaki

	YD=inputbox("出荷年月を入れて下さい 例 1608")
	if len(YD) = 4 then
		Ret=inputbox("機種を選択してください。1:Wish,2:Wash")
		if Ret <> "1" and Ret <> "2" then
			msgbox("1または2を入力して下さい")
		else
			if Ret="1" then
				kisyu="WI"
			else
				if Ret = "2" then
					kisyu=""
				end if
			end if
			Ret=inputbox("向け先を選択してください。1:Secom,2:Pana,3:Minerish")
			if Ret <> "1" and Ret <> "2" and Ret <> "3" then
				msgbox("1または2または3を入力して下さい")
			else
				if Ret="1" then
					Mukesaki="SK"
				elseif Ret = "2" then
					Mukesaki="PD"
				elseif Ret = "3" then
					Mukesaki="MI"
				end if
				
				if Kisyu = "WI" then
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_AT_TRU1"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_AT_TRU2"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_出M_通常"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_工M_通常"
					msgbox("不良一覧は[不良一覧_■" & Kisyu & Mukesaki & YD & ".xls]にして下さい")
				else
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_AT_TRU1"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_AT_TRU2"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_CAL_TRU1"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_CAL_TRU2"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_出M_通常"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_工M_通常"
					msgbox("不良一覧は[不良一覧_■" & Kisyu & Mukesaki & YD & ".xls]にして下さい")
				end if
			end if
		end if
	else
		msgbox("長さが異常です")
	end if

end sub

sub mkdir(path)

	Dim ObjFso

	Set ObjFso=WScript.CreateObject("Scripting.FileSystemObject")

	If ObjFso.FolderExists(path) = False Then
		ObjFso.Createfolder(path)
	End If

	set ObjFso = Nothing

end sub

function apppath
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    apppath = fso.getParentFolderName(WScript.ScriptFullName)
end function
