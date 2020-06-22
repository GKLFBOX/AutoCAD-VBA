Attribute VB_Name = "AlignTextPosition"
'------------------------------------------------------------------------------
' ## �R�[�f�B���O�K�C�h���C��
'
' [You.Activate|VBA�R�[�f�B���O�K�C�h���C��]�ɏ�������
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## �����ʒu����
'
' �w������2�_�ƃI�t�Z�b�g�W�����當���ʒu�𒲐�(��������)
'------------------------------------------------------------------------------
Public Sub AlignTextPosition()
    
    On Error GoTo Error_Handler
    
    ' �Ώۂ̑I��/�w��/����
    Dim pickPoint As Variant
    Dim targetText As ZcadEntity    ' TODO: target�͌^�w���v�ۂ���������
    ThisDrawing.Utility.GetEntity targetText, pickPoint, _
        "�����I�u�W�F�N�g��I�� [Cancel(ESC)]"
    
    If Not CommonFunction.IsTextObject(targetText) Then
        ThisDrawing.Utility.Prompt "�G���[�F������I�����Ă��������B" & vbCrLf
        Exit Sub
    End If
    
    targetText.Highlight True
    
    Dim firstPoint As Variant
    Dim secondPoint As Variant
    firstPoint = ThisDrawing.Utility.GetPoint _
        (, "1�_�ڂ��w�� [Cancel(ESC)]")
    secondPoint = ThisDrawing.Utility.GetPoint _
        (firstPoint, "2�_�ڂ��w�� [Cancel(ESC)]")
    
    Dim offsetFactor As Double
    offsetFactor = ThisDrawing.Utility.GetReal _
        ("�I�t�Z�b�g�W�������(���������ɑ΂��銄��(x/10) " & _
         "[�ʏ�(2)/�L��(3)/����(1)/���L��(5)]:")
    offsetFactor = offsetFactor * 0.1
    
    Dim underFlag As String
    underFlag = ThisDrawing.Utility.GetString _
        (0, "���t���ɂ��܂���? [�͂�(Y)/������(N)]:")
    
    If underFlag = "Y" Then offsetFactor = offsetFactor * -1
    
    targetText.Rotation = 0 ' �I�t�Z�b�g�ʂ̓K�p�ȗ����̂��ߊp�x�v�f�폜
    
    ' ���_�ʒu�̎擾
    Dim textCenter() As Double
    textCenter = getTextCenter(targetText)
    textCenter(1) = textCenter(1) - targetText.Height * Abs(offsetFactor)
    
    ' �����ʒu�����̎��s
    Dim alignPoint() As Double
    Dim alignRad As Double
    alignPoint = getMiddlePoint(firstPoint, secondPoint)
    alignRad = calculateAngle(firstPoint, secondPoint)
    
    targetText.Move textCenter, alignPoint
    targetText.Rotate alignPoint, alignRad
    
    ' ���t������Ə���
    If offsetFactor < 0 Then
        Dim mirrorText As ZcadEntity
        Set mirrorText = targetText.Mirror(firstPoint, secondPoint)
        targetText.Delete
    Else
        Call CommonSub.ResetHighlight(targetText)
    End If
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ResetHighlight(targetText)
    ThisDrawing.Utility.Prompt "�G���[�F�R�}���h���I�����܂��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## 2�_�̊p�x�v�Z
'------------------------------------------------------------------------------
Private Function calculateAngle(ByRef first_point As Variant, _
                                ByRef second_point As Variant) As Double
    
    calculateAngle = Atn2(second_point(0) - first_point(0), _
                          second_point(1) - first_point(1))
    
End Function

'------------------------------------------------------------------------------
' ## �S�p�x�Ή�Atn�֐�
'------------------------------------------------------------------------------
Private Function Atn2(delta_x As Double, delta_y As Double) As Double
    
    Dim pi As Double
    pi = 4 * Atn(1)
    
    If delta_x = 0 And delta_y = 0 Then
        Atn2 = 0
        
    ElseIf delta_x > 0 And delta_y = 0 Then ' ��=0
        Atn2 = (pi / 2) * 0
        
    ElseIf delta_x = 0 And delta_y > 0 Then ' ��=90
        Atn2 = (pi / 2) * 1
        
    ElseIf delta_x < 0 And delta_y = 0 Then ' ��=180
        Atn2 = (pi / 2) * 2
        
    ElseIf delta_x = 0 And delta_y < 0 Then ' ��=270
        Atn2 = (pi / 2) * 3
        
    ElseIf delta_x > 0 And delta_y > 0 Then ' 0<��<90
        Atn2 = Atn(Abs(delta_y) / Abs(delta_x)) + ((pi / 2) * 0)
        
    ElseIf delta_x < 0 And delta_y > 0 Then ' 90<��<180
        Atn2 = ((pi / 2) - Atn(Abs(delta_y) / Abs(delta_x))) + ((pi / 2) * 1)
        
    ElseIf delta_x < 0 And delta_y < 0 Then ' 180<��<270
        Atn2 = Atn(Abs(delta_y) / Abs(delta_x)) + ((pi / 2) * 2)
        
    ElseIf delta_x > 0 And delta_y < 0 Then ' 90<��<180
        Atn2 = ((pi / 2) - Atn(Abs(delta_y) / Abs(delta_x))) + ((pi / 2) * 3)
        
    End If
    
End Function

'------------------------------------------------------------------------------
' ## 2�_�̒��_�擾
'------------------------------------------------------------------------------
Private Function getMiddlePoint(ByRef first_point As Variant, _
                                ByRef second_point As Variant) As Double()
    
    Dim i As Long
    Dim tmp(0 To 2) As Double
    
    For i = 0 To 2
        tmp(i) = (first_point(i) + second_point(i)) / 2
    Next i
    
    getMiddlePoint = tmp()
    
End Function

'------------------------------------------------------------------------------
' ## �����̉����S�擾
'------------------------------------------------------------------------------
Private Function getTextCenter(ByVal target_object As ZcadEntity) As Double()
    
    Dim minExtent As Variant, maxExtent As Variant
    target_object.GetBoundingBox minExtent, maxExtent
    
    ' BoundingBox�̎d�l�ύX�ɔ����X�Ίp�x���l������
    Dim targetOblique As Double
    Dim boxHeight As Double
    Dim exAmount As Double
    targetOblique = target_object.ObliqueAngle
    boxHeight = maxExtent(1) - minExtent(1)
    exAmount = boxHeight * Tan(targetOblique)
    maxExtent(0) = maxExtent(0) + exAmount
    
    Dim startPoint(0 To 2) As Double, endPoint(0 To 2) As Double
    startPoint(0) = minExtent(0): endPoint(0) = maxExtent(0)
    startPoint(1) = minExtent(1): endPoint(1) = minExtent(1)
    startPoint(2) = 0: endPoint(2) = 0
    
    Dim i As Long
    Dim tmp(0 To 2) As Double
    
    For i = 0 To 2
        tmp(i) = (startPoint(i) + endPoint(i)) / 2
    Next i
    
    getTextCenter = tmp()
    
End Function
