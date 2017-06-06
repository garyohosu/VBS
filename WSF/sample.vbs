'http://blog.codebook-10000.com/entry/20140425/1398418913

Option Explicit

Function testFunction()
  Dim i
  Dim total

  For i = 1 To 100
    total = total + i
  Next

  testFunction = total

End Function