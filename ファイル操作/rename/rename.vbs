option explicit
dim sfo, gfile

const filepath = "C:\test\test.txt" 'ファイルのパス

set sfo = createobject("scripting.filesystemobject")
set gfile = sfo.getfile(filepath)

gfile.name = "test001.txt" '名前をtest.txtからtest001.txtに変更

set sfo = nothing
set gfile = nothing



sub rename(filepath,filename)
	dim sfo, gfile
	
	set sfo = createobject("scripting.filesystemobject")
	set gfile = sfo.getfile(filepath)

	gfile.name = filename

	set sfo = nothing
	set gfile = nothing
end sub
