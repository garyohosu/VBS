Option Explicit

Dim pFunc             ' �֐��|�C���^
Dim intCnt            ' �J�E���^
Dim aryFunc(1)        ' �Ăяo���֐���
aryFunc(0) = "GetDate"
aryFunc(1) = "GetTime"

for intCnt = 0 to UBound(aryFunc)
  Set pFunc = GetRef(aryFunc(intCnt))
  Msgbox pFunc()
  Set pFunc = Nothing
next

'�V�X�e�����t�擾
Function GetDate()
  GetDate = Date()
End Function

'�V�X�e�����Ԏ擾
Function GetTime()
  GetTime = TIME()
End Function
