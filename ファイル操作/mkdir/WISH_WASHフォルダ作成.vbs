Option Explicit
'===========================================t
' Wish/Wash�t�H���_�����}�N��
'
'���藚��
'DATE       VER  NAME    COMMENT
'2016/08/18 0.00 HANTANI �V�K�쐬
'===========================================t

call main

sub main
	dim YD
	dim Ret
	dim Kisyu
	dim Mukesaki

	YD=inputbox("�o�הN�������ĉ����� �� 1608")
	if len(YD) = 4 then
		Ret=inputbox("�@���I�����Ă��������B1:Wish,2:Wash")
		if Ret <> "1" and Ret <> "2" then
			msgbox("1�܂���2����͂��ĉ�����")
		else
			if Ret="1" then
				kisyu="WI"
			else
				if Ret = "2" then
					kisyu=""
				end if
			end if
			Ret=inputbox("�������I�����Ă��������B1:Secom,2:Pana,3:Minerish")
			if Ret <> "1" and Ret <> "2" and Ret <> "3" then
				msgbox("1�܂���2�܂���3����͂��ĉ�����")
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
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_�oM_�ʏ�"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_�HM_�ʏ�"
					msgbox("�s�ǈꗗ��[�s�ǈꗗ_��" & Kisyu & Mukesaki & YD & ".xls]�ɂ��ĉ�����")
				else
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_AT_TRU1"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_AT_TRU2"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_CAL_TRU1"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_CAL_TRU2"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_�oM_�ʏ�"
					mkdir apppath & "\" & Kisyu & Mukesaki & YD & "_�HM_�ʏ�"
					msgbox("�s�ǈꗗ��[�s�ǈꗗ_��" & Kisyu & Mukesaki & YD & ".xls]�ɂ��ĉ�����")
				end if
			end if
		end if
	else
		msgbox("�������ُ�ł�")
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
