'http://qiita.com/asterisk9101/items/a7310c0e4ec33352835f
'�t�H���_���̃t�@�C���ꗗ��Ԃ�

dim fso
set fso = createObject("Scripting.FileSystemObject")

dim folder
set folder = fso.getFolder("C:\")

' �t�@�C���ꗗ
dim file
for each file in folder.files
    msgbox file.name
next 

' �T�u�t�H���_�ꗗ
dim subfolder
for each subfolder in folder.subfolders
    msgbox subfolder.name
next
