#include <array.au3>
#include <Word.au3>
Const $msoPatternDarkDownwardDiagonal = 15
Const $msoShapeRectangle = 1
Const $msoGradientHorizontal = 1

_WordErrorHandlerRegister()
$oWordApp = _WordCreate()
$oDocActive = _WordDocGetCollection($oWordApp, 0)

;;��һ����ֱ��

_WordDocLine($oDocActive, 510, 50, 90, 50, 1.5, 255)

Func _WordDocLine($o_DocActive, $s_Left, $s_left_H, $s_Line, $s__right_H, $s_Weight, $s_Color)
        With $o_DocActive.Shapes.AddLine($s_Left, $s_left_H, $s_Line, $s__right_H).Line
                .Weight = $s_Weight
                .ForeColor.RGB = $s_Color
                ;.BackColor.RGB = 128
        EndWith
	EndFunc   ;==>_WordDocLine

;;;����ʾ��Ϊ��ĵ����ͼ��������

With $oDocActive.Shapes.AddLine(200, 80, 350, 80).Line
        .Weight = 8
        .ForeColor.RGB = 255
        .BackColor.RGB = 128;;;�����ı���ɫ
        .Pattern = $msoPatternDarkDownwardDiagonal;;;��ͼ��
EndWith
;;;����ʾ��Ϊ����ĵ����һ���߶Ρ�ֱ�ߵ������һ������խ����Բ���յ���һ�����ҿ�������Ρ�
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

;;;����ʾ����һ�����������������������������ĵ��С�
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

;;;�������ڵ�ǰ�ĵ������һ����ע��Ȼ�����ñ�ע�ĽǶȡ�
#cs
        ���ĵ�����ӻ��������ش���û����� Shape ���󣬲�������ӵ� Shapes ���ϡ�
        .AddCanvas(Left, Top, Width, Height, Anchor)
        Left  Single ���ͣ�     ���衣��������Ե�����������ǵ�λ�ã��԰�Ϊ��λ��
        Top  Single ���ͣ�      ���衣�����ϲ���Ե�����������ǵ�λ�ã��԰�Ϊ��λ��
        Width  Single ���ͣ�    ���衣�����Ŀ�ȣ��԰�Ϊ��λ��
        Height  Single ���ͣ�   ���衣�����ĸ߶ȣ��԰�Ϊ��λ��
        Anchor  Variant ���ͣ�  ��ѡ�����������ı��� Range ����
        ���ָ�� Anchor����������ǽ����������������һ�εĿ�ͷ��
        ���ʡ�Ըò��������Զ�ѡ���������򣬶������������ҳ����ϲ�������Ե���ж�λ��
#ce
;ʾ��
Const $msoCalloutTwo = 2
Const $msoCalloutAngle30 = 2
_WordDocNewCallout($oDocActive)
Func _WordDocNewCallout($o_DocActive)

        With $o_DocActive.Shapes.AddCallout($msoCalloutTwo, 250, 180, 100, 80)
                .TextFrame.TextRange.Text = "�ڵ�ǰ�ĵ������һ����ע."
                .Callout.Angle = $msoCalloutAngle30
        EndWith

EndFunc   ;==>_WordDocNewCallout

;;;����ʾ�������ĵ�����ӻ�����Ȼ���ڻ������������ͼ�Σ������������������ԡ�

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
;;;����ʾ��ʵ�ֵĹ����ǣ�������ͼ������ myDocument���������������ͼ�Σ�����ͼ����ϵ�����ʽ����ת����ϲ��������ڻ�ͼ������档
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


;;;����ʾ����һ������ɫ����ʯ�������ľ���������ĵ��С�
Const $msoTextureGreenMarble = 9
$oDocActive.Shapes.AddShape($msoShapeCan, 490, 490, 40, 80).Fill.PresetTextured($msoTextureGreenMarble)

;;;����ʾ���������ı���Test���ġ������֡���ӵ���ĵ��У����������ɺ��ţ�ָ���������֡���ʽ��Ĭ��ֵ���� msoTextEffect1��ת��Ϊ���š�

Const $msoTextEffect1 = 0
$newWordArt = $oDocActive.Shapes.AddTextEffect($msoTextEffect1, "Test", "Arial Black", 36, False, False, 350, 100)
$newWordArt.TextEffect.ToggleVerticalText

;;;����ʾ�����ĵ�������������Σ���������ϣ�Ϊ�����������һ����ɫ��Ȼ��ֻ���ĵڶ��������ε���ɫ��
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


;;;��ʾ�����ĵ��������ʮ������״����Ϊÿһ��ʮ�������õ�һ�ֵ���ֵ�����ڴ��ࡰ��ѡͼ�Ρ����õ�����ʽ��Ψһ�ģ���
Const $msoShapeCross = 11

With $oDocActive.Shapes
    .AddShape($msoShapeCross, _
        10, 10, 100, 100).Adjustments.Item(1) = 0.4
    .AddShape($msoShapeCross, _
        150, 10, 100, 100).Adjustments.Item(1) = 0.2
EndWith

#cs
        ;;;���ĵ������һ�����Σ�Ȼ�����þ�������ǰ��ɫ������ɫ�͹��ɡ�
        .AddShape(Type, Left, Top, Width, Height, Anchor)
        expression   ���衣�ñ��ʽ����һ�� Shapes ����
        Type  Long ���ͣ�     ���衣Ҫ���ص�ͼ�����͡��������κ� MsoAutoShapeType ������
        Left  Single ���ͣ�   ���衣����ѡͼ�Ρ���������Ե��λ�ã��԰�Ϊ��λ��
        Top  Single ���ͣ�    ���衣����ѡͼ�Ρ������ϲ���Ե��λ�ã��԰�Ϊ��λ��
        Width  Single ���ͣ�  ���衣����ѡͼ�Ρ�����Ŀ�ȣ��԰�Ϊ��λ��
        Height  Single ���ͣ� ���衣����ѡͼ�Ρ�����ĸ߶ȣ��԰�Ϊ��λ��
        Anchor  Variant ���ͣ���ѡ������á���ѡͼ�Ρ��������ı��� Range ����
        ���ָ�� Anchor�����������λ�����������һ�ε���ʼλ�á�
        ������Ըò������� Word ���Զ�ѡ����������
        ����ѡͼ�ν������ҳ����ϲ�������Ե���ж�λ��
#ce
_WordDocRectangle($oDocActive);���ĵ������һ�����Σ�
Func _WordDocRectangle($o_DocActive)

        With $o_DocActive.Shapes.AddShape($msoShapeRectangle, 250, 120, 90, 50).Fill
                .ForeColor.RGB = 128
                .BackColor.RGB = 170
                .TwoColorGradient($msoGradientHorizontal, 1)
        EndWith
EndFunc   ;==>_WordDocRectangle