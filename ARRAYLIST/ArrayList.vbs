
' ArrayList�̍쐬
Dim myArrayList
Set myArrayList = CreateObject("System.Collections.ArrayList")

' �v�f��ǉ����Ă���
myArrayList.add "Alice"
myArrayList.add "Bob"
myArrayList.add "Charlie"

' For Each���ŗv�f��\��
Dim elem
For Each elem In myArrayList
   Wscript.echo elem
Next

' For���ŗv�f��\��
Dim i
Dim iMax : iMax = myArrayList.Count - 1
For i=0 To iMax
   Wscript.echo myArrayList(i)
Next

' �V�����v�f��}��
myArrayList.Insert 1, "Dave"
myArrayList.Insert 2, "Ellen"

' �v�f���\�[�g
myArrayList.Sort

' ����̗v�f���폜
myArrayList.Remove "Dave"

' �ʏ�̔z��ɕϊ�
Dim myArray : myArray = myArrayList.ToArray()

' �S�v�f���폜
myArrayList.Clear
