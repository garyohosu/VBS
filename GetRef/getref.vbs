Option Explicit

Dim pFunc             ' 関数ポインタ
Dim intCnt            ' カウンタ
Dim aryFunc(1)        ' 呼び出す関数名
aryFunc(0) = "GetDate"
aryFunc(1) = "GetTime"

for intCnt = 0 to UBound(aryFunc)
  Set pFunc = GetRef(aryFunc(intCnt))
  Msgbox pFunc()
  Set pFunc = Nothing
next

'システム日付取得
Function GetDate()
  GetDate = Date()
End Function

'システム時間取得
Function GetTime()
  GetTime = TIME()
End Function
