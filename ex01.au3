;;  变量及表达式
$a=100
$b=100
$c=$a+$b
$str="AutoIt Script"
ConsoleWrite("The Sum is: " & $c &"." &@CRLF)
ConsoleWrite("计算 by " & $str &"." &@CRLF)
MsgBox(0,"计算结果","计算结果是：" & $c)
