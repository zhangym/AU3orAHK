#include <array.au3>
#include <Word.au3>
Const $msoPatternDarkDownwardDiagonal = 15
Const $msoShapeRectangle = 1
Const $msoGradientHorizontal = 1

_WordErrorHandlerRegister()
$oWordApp = _WordCreate()
$oDocActive = _WordDocGetCollection($oWordApp, 0)

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