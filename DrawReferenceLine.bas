Attribute VB_Name = "DrawReferenceLine"
'------------------------------------------------------------------------------
' ## �R�[�f�B���O�K�C�h���C��
'
' [You.Activate|VBA�R�[�f�B���O�K�C�h���C��]�ɏ�������
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## �Q�Ɛ��̍�}
'
' �w�肵�������ƃI�t�Z�b�g�W������Q�Ɛ�����}����
'------------------------------------------------------------------------------
Public Sub DrawReferenceLine()
    
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
    
    ' ��}�ȗ����̂��߂Ɋ�_�Ɗp�x���L�����p�x�v�f�폜
    Dim targetPoint As Variant
    Dim targetAngle As Double
    targetPoint = targetText.InsertionPoint
    targetAngle = targetText.Rotation
    targetText.Rotate targetPoint, targetAngle * -1
    
    ' �Q�Ɛ��̎n�I�[�v�Z
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
    startPoint(1) = minExtent(1) - targetText.Height * 0.2
    startPoint(2) = 0
    
    endPoint(0) = maxExtent(0) + ((maxExtent(1) - minExtent(1)) * 0.15)
    endPoint(1) = minExtent(1) - targetText.Height * 0.2
    endPoint(2) = 0
    
    ' �Q�Ɛ���}
    Dim referenceLine As ZcadLine
    Set referenceLine = ThisDrawing.ModelSpace.AddLine(startPoint, endPoint)
    
    ' ��������ю��������̊p�x��߂�
    targetText.Rotate targetPoint, targetAngle
    referenceLine.Rotate targetPoint, targetAngle
    
    Call CommonSub.ResetHighlight(targetText)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ResetHighlight(targetText)
    ThisDrawing.Utility.Prompt "�G���[�F�R�}���h���I�����܂��B" & vbCrLf
    
End Sub
