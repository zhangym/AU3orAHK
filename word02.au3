#include <array.au3>
#include <Word.au3>
Const $msoPatternDarkDownwardDiagonal = 15
Const $msoShapeRectangle = 1
Const $msoGradientHorizontal = 1

_WordErrorHandlerRegister()
$oWordApp = _WordCreate()
$oDocActive = _WordDocGetCollection($oWordApp, 0)

;;;下面示例为活动文档添加图案线条。

With $oDocActive.Shapes.AddLine(200, 80, 350, 80).Line
        .Weight = 8
        .ForeColor.RGB = 255
        .BackColor.RGB = 128;;;线条的背景色
        .Pattern = $msoPatternDarkDownwardDiagonal;;;加图案
EndWith
;;;下面示例为活动的文档添加一条线段。直线的起点有一个短且窄的椭圆，终点有一个长且宽的三角形。
Const $msoArrowheadShort = 1
Const $msoArrowheadOval = 6
Const $msoArrowheadNarrow = 1
Const $msoArrowheadLong = 3
Const $msoArrowheadTriangle = 2
Const $msoArrowheadWide = 3