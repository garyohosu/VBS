
dim objStdIn,objStdOut

set objStdIn =  Wscript.stdIn
set objStdOut = Wscript.stdout

dim mes

do 
	mes = objStdIn.readLine
	objStdOut.writeLine ucase(mes)
loop while mes <> ""

objStdOut.writeLine "bye"

objStdOut.close
objStdIn.close

set objStdIn = Nothing
set objStdOut = Nothing