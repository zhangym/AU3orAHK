Run("notepad.exe")
WinWaitActive("�ޱ��� - ���±�")
Send("This is some text.")
WinClose("�ޱ��� - ���±�")
WinWaitActive("���±�", "����")
;WinWaitActive("Notepad", "Do you want to save") ; When running under Windows XP
Send("!n")