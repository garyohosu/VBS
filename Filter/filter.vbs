Dim aryIndex
Dim aryStrings
aryStrings = Array("�R�c", "����", "���R", "�R�{", "���")
aryIndex = Filter(aryStrings, "�R", True, 1)     '"�R�c", "���R", "�R�{" �̂R�v�f���Z�b�g����܂��B
msgbox(Join(aryIndex, "/"))
aryIndex = Filter(aryStrings, "�R", False, 1)    '"����", "���" �̂Q�v�f���Z�b�g����܂��B
msgbox(Join(aryIndex , "/"))
