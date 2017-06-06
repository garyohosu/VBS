'http://qiita.com/asterisk9101/items/a7310c0e4ec33352835f
'フォルダ内のファイル一覧を返す

dim fso
set fso = createObject("Scripting.FileSystemObject")

dim folder
set folder = fso.getFolder("C:\")

' ファイル一覧
dim file
for each file in folder.files
    msgbox file.name
next 

' サブフォルダ一覧
dim subfolder
for each subfolder in folder.subfolders
    msgbox subfolder.name
next
