Run("notepad.exe")
WinWaitActive("无标题 - 记事本")
Send("This is some text.")
WinClose("无标题 - 记事本")
WinWaitActive("记事本", "保存")
;WinWaitActive("Notepad", "Do you want to save") ; When running under Windows XP
Send("!n")