option explicit
dim sfo, gfile

const filepath = "C:\test\test.txt" '�t�@�C���̃p�X

set sfo = createobject("scripting.filesystemobject")
set gfile = sfo.getfile(filepath)

gfile.name = "test001.txt" '���O��test.txt����test001.txt�ɕύX

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
