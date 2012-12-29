#include <array.au3>
#include <Word.au3>
Const $msoPatternDarkDownwardDiagonal = 15
Const $msoShapeRectangle = 1
Const $msoGradientHorizontal = 1

_WordErrorHandlerRegister()
$oWordApp = _WordCreate()
$oDocActive = _WordDocGetCollection($oWordApp, 0)

;;画一条红直线

_WordDocLine($oDocActive, 510, 50, 90, 50, 1.5, 255)

Func _WordDocLine($o_DocActive, $s_Left, $s_left_H, $s_Line, $s__right_H, $s_Weight, $s_Color)
        With $o_DocActive.Shapes.AddLine($s_Left, $s_left_H, $s_Line, $s__right_H).Line
                .Weight = $s_Weight
                .ForeColor.RGB = $s_Color
                ;.BackColor.RGB = 128
        EndWith
	EndFunc   ;==>_WordDocLine

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

With $oDocActive.Shapes.AddLine(100, 200, 200, 300).Line
        .BeginArrowheadLength = $msoArrowheadShort
        .BeginArrowheadStyle = $msoArrowheadOval
        .BeginArrowheadWidth = $msoArrowheadNarrow
        .EndArrowheadLength = $msoArrowheadLong
        .EndArrowheadStyle = $msoArrowheadTriangle
        .EndArrowheadWidth = $msoArrowheadWide
EndWith

;;;下面示例将一个具有五个顶点的任意多边形添加至活动文档中。
Const $msoSegmentCurve = 1
Const $msoEditingCorner = 1
Const $msoEditingAuto = 0
Const $msoSegmentLine = 0
With $oDocActive.Shapes.BuildFreeform($msoEditingCorner, 360, 200)
        .AddNodes($msoSegmentCurve, $msoEditingCorner, 380, 230, 400, 250, 450, 300)
        .AddNodes($msoSegmentCurve, $msoEditingAuto, 480, 200)
        .AddNodes($msoSegmentLine, $msoEditingAuto, 480, 400)
        .AddNodes($msoSegmentLine, $msoEditingAuto, 360, 200)
        .ConvertToShape
EndWith

;;;下面例在当前文档中添加一个标注，然后设置标注的角度。
#cs
        在文档中添加画布。返回代表该画布的 Shape 对象，并将其添加到 Shapes 集合。
        .AddCanvas(Left, Top, Width, Height, Anchor)
        Left  Single 类型，     必需。画布左侧边缘相对于锁定标记的位置，以磅为单位。
        Top  Single 类型，      必需。画布上部边缘相对于锁定标记的位置，以磅为单位。
        Width  Single 类型，    必需。画布的宽度，以磅为单位。
        Height  Single 类型，   必需。画布的高度，以磅为单位。
        Anchor  Variant 类型，  可选。代表画布绑定文本的 Range 对象。
        如果指定 Anchor，则锁定标记将出现在锁定区域第一段的开头。
        如果省略该参数，将自动选定锁定区域，而画布将相对于页面的上部和左侧边缘进行定位。
#ce
;示例
Const $msoCalloutTwo = 2
Const $msoCalloutAngle30 = 2
_WordDocNewCallout($oDocActive)
Func _WordDocNewCallout($o_DocActive)

        With $o_DocActive.Shapes.AddCallout($msoCalloutTwo, 250, 180, 100, 80)
                .TextFrame.TextRange.Text = "在当前文档中添加一个标注."
                .Callout.Angle = $msoCalloutAngle30
        EndWith

EndFunc   ;==>_WordDocNewCallout

;;;下列示例在新文档中添加画布，然后在画布上添加两个图形，并设置填充和线条属性。

Const $wdWrapInline = 7
Const $msoShapeHeart = 21

_WordDocAddInlineCanvas($oDocActive)
Func _WordDocAddInlineCanvas($o_DocActive)


        $shpCanvas = $o_DocActive.Shapes.AddCanvas(150, 150, 70, 70)
        $shpCanvas.WrapFormat.Type = $wdWrapInline

        With $shpCanvas.CanvasItems
                .AddShape($msoShapeHeart, 10, 10, 50, 60)
                .AddLine(0, 0, 70, 70)
        EndWith
        With $shpCanvas
                .CanvasItems(1).Fill.ForeColor = 255
                .CanvasItems(2).Line.EndArrowheadStyle = $msoArrowheadTriangle
        EndWith


EndFunc   ;==>_WordDocAddInlineCanvas
;;;下列示例实现的功能是：将两个图形添至 myDocument，并组合这两个新图形，设置图形组合的填充格式，旋转此组合并将其置于绘图层的下面。
Const $msoShapeCan = 13
Const $msoShapeCube = 14
Const $msoTextureBlueTissuePaper = 17 ;(&H11)
Const $msoSendToBack = 1
$oDocActive.Shapes.AddShape($msoShapeCan, 150, 350, 100, 200).Name = "shpOne"
$oDocActive.Shapes.AddShape($msoShapeCube, 150, 550, 100, 200).Name = "shpTwo"
With $oDocActive.Shapes.Range(_ArrayCreate("shpOne", "shpTwo") ).Group
        .Fill.PresetTextured($msoTextureBlueTissuePaper)
        .Rotation = 45
        .ZOrder($msoSendToBack)
EndWith


;;;下列示例将一个用绿色大理石纹理填充的矩形添至活动文档中。
Const $msoTextureGreenMarble = 9
$oDocActive.Shapes.AddShape($msoShapeCan, 490, 490, 40, 80).Fill.PresetTextured($msoTextureGreenMarble)

;;;下列示例将包含文本“Test”的“艺术字”添加到活动文档中，并将文字由横排（指定“艺术字”样式的默认值，即 msoTextEffect1）转换为纵排。

Const $msoTextEffect1 = 0
$newWordArt = $oDocActive.Shapes.AddTextEffect($msoTextEffect1, "Test", "Arial Black", 36, False, False, 350, 100)
$newWordArt.TextEffect.ToggleVerticalText

;;;下列示例向活动文档添加三个三角形，并加以组合，为整个组合设置一个颜色，然后只更改第二个三角形的颜色。
Const $msoShapeIsoscelesTriangle = 7
;Const $msoTextureBlueTissuePaper = 17 ;(&H11)
;Const $msoTextureGreenMarble = 9
$oDocActive.Shapes.AddShape($msoShapeIsoscelesTriangle, _
                100, 600, 100, 100).Name = "shpOne"
$oDocActive.Shapes.AddShape($msoShapeIsoscelesTriangle, _
                240, 600, 100, 100).Name = "shpTwo"
$oDocActive.Shapes.AddShape($msoShapeIsoscelesTriangle, _
                390, 600, 100, 100).Name = "shpThree"
With $oDocActive.Shapes.Range(_ArrayCreate("shpOne", "shpTwo", "shpThree") ).Group
        .Fill.PresetTextured($msoTextureBlueTissuePaper)
        .GroupItems(2).Fill.PresetTextured($msoTextureGreenMarble)
EndWith


;;;本示例向活动文档添加两个十字形形状，并为每一个十字形设置第一种调整值（对于此类“自选图形”，该调整方式是唯一的）。
Const $msoShapeCross = 11

With $oDocActive.Shapes
    .AddShape($msoShapeCross, _
        10, 10, 100, 100).Adjustments.Item(1) = 0.4
    .AddShape($msoShapeCross, _
        150, 10, 100, 100).Adjustments.Item(1) = 0.2
EndWith

#cs
        ;;;向活动文档中添加一个矩形，然后设置矩形填充的前景色、背景色和过渡。
        .AddShape(Type, Left, Top, Width, Height, Anchor)
        expression   必需。该表达式返回一个 Shapes 对象。
        Type  Long 类型，     必需。要返回的图形类型。可以是任何 MsoAutoShapeType 常量。
        Left  Single 类型，   必需。“自选图形”对象左侧边缘的位置，以磅为单位。
        Top  Single 类型，    必需。“自选图形”对象上部边缘的位置，以磅为单位。
        Width  Single 类型，  必需。“自选图形”对象的宽度，以磅为单位。
        Height  Single 类型， 必需。“自选图形”对象的高度，以磅为单位。
        Anchor  Variant 类型，可选。代表该“自选图形”所连接文本的 Range 对象。
        如果指定 Anchor，则锁定标记位于锁定区域第一段的起始位置。
        如果忽略该参数，则 Word 将自动选定锁定区域，
        而自选图形将相对于页面的上部和左侧边缘进行定位。
#ce
_WordDocRectangle($oDocActive);向活动文档中添加一个矩形，
Func _WordDocRectangle($o_DocActive)

        With $o_DocActive.Shapes.AddShape($msoShapeRectangle, 250, 120, 90, 50).Fill
                .ForeColor.RGB = 128
                .BackColor.RGB = 170
                .TwoColorGradient($msoGradientHorizontal, 1)
        EndWith
EndFunc   ;==>_WordDocRectangle