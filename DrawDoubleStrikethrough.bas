Attribute VB_Name = "DrawDoubleStrikethrough"
Option Explicit

'------------------------------------------------------------------------------
' ## ��d���������̍�}
'
' �w�肵�������I�u�W�F�N�g�ɓ���w�œ�d������������}����
'------------------------------------------------------------------------------
Public Sub DrawDoubleStrikethrough()
    
    On Error GoTo Error_Handler
    
    ' �Ώۂ̑I��
    Dim pickPoint As Variant
    Dim targetText As ZcadEntity
    ThisDrawing.Utility.GetEntity targetText, pickPoint, _
        "�����I�u�W�F�N�g��I�� [Cancel(ESC)]"
    
    If Not CommonFunction.IsTextObject(targetText) Then
        ThisDrawing.Utility.Prompt "�G���[�F������I�����Ă��������B" & vbCrLf
        Exit Sub
    End If
    
    ' FIXME: ����ă}���`�e�L�X�g��I������Ɗp�x�v�f�̍폜�݂̂��s����
    ' ��}�ȗ����̂��߂Ɋ�_�Ɗp�x���L�����p�x�v�f�폜
    Dim targetPoint As Variant
    Dim targetAngle As Double
    targetPoint = targetText.insertionPoint
    targetAngle = targetText.Rotation
    targetText.Rotate targetPoint, targetAngle * -1
    
    ' ���������̎n�I�[�v�Z
    Dim minExtent As Variant, maxExtent As Variant
    targetText.GetBoundingBox minExtent, maxExtent
    
    ' BoundingBox�̎d�l�ύX�ɔ����X�Ίp�x���l������
    Dim targetOblique As Double
    Dim boxHeight As Double
    Dim exAmount As Double
    targetOblique = targetText.ObliqueAngle
    boxHeight = maxExtent(1) - minExtent(1)
    exAmount = boxHeight * Tan(targetOblique)
    maxExtent(0) = maxExtent(0) + exAmount
    
    Dim startPoint(0 To 2) As Double, endPoint(0 To 2) As Double
    
    startPoint(0) = minExtent(0) - ((maxExtent(1) - minExtent(1)) * 0.15)
    startPoint(1) = minExtent(1) + ((maxExtent(1) - minExtent(1)) / 3)
    startPoint(2) = 0
    
    endPoint(0) = maxExtent(0) + ((maxExtent(1) - minExtent(1)) * 0.15)
    endPoint(1) = minExtent(1) + ((maxExtent(1) - minExtent(1)) / 3)
    endPoint(2) = 0
    
    ' ����������}
    Dim strikeThrough(0 To 1) As ZcadLine, offsetLine As Variant
    
    Set strikeThrough(0) = ThisDrawing.ModelSpace.AddLine(startPoint, endPoint)
    offsetLine = strikeThrough(0).Offset((maxExtent(1) - minExtent(1)) / 3)
    
    Set strikeThrough(1) = offsetLine(0)
    
    ' ��������ю��������̊p�x��߂�
    targetText.Rotate targetPoint, targetAngle
    
    strikeThrough(0).Layer = targetText.Layer
    strikeThrough(0).Rotate targetPoint, targetAngle
    
    strikeThrough(1).Layer = targetText.Layer
    strikeThrough(1).Rotate targetPoint, targetAngle
    
    Call CommonSub.ResetHighlight(targetText)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ResetHighlight(targetText)
    ThisDrawing.Utility.Prompt "�G���[�F�R�}���h���I�����܂��B" & vbCrLf
    
End Sub
