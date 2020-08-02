Attribute VB_Name = "DrawReferenceLine"
Option Explicit

'------------------------------------------------------------------------------
' ## �����n�I�u�W�F�N�g�ւ̎Q�Ɛ���}
'
' �w�肵�������ƃI�t�Z�b�g�W������Q�Ɛ�����}����
'------------------------------------------------------------------------------
Public Sub DrawReferenceLine()
    
    On Error GoTo Error_Handler
    
    Dim pickPoint As Variant
    Dim targetEntity As ZcadEntity
    Dim configLayer As String
    Dim configLength As Single
    Dim configOffset As Single
    
    ' �ʒu�������镶���n�I�u�W�F�N�g�̑I��
    ThisDrawing.Utility.GetEntity targetEntity, pickPoint, _
        "�ʒu�������镶���܂��̓u���b�N��������I�� [Cancel(ESC)]"
    
    Dim configData As Variant
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.REFERENCELINE_CONFIG), vbCrLf)
    If UBound(configData) = 2 Then
        configLayer = configData(0)
        configLength = configData(1) / 2
        configOffset = configData(2)
    End If
    
    ' �e�L�X�g�܂��̓u���b�N�Q�Ƃ̔���
    If CommonFunction.IsTextObject(targetEntity) Then
        Call addTextReferenceLine _
            (targetEntity, pickPoint, configLayer, configLength, configOffset)
        Exit Sub
    ElseIf TypeOf targetEntity Is ZcadBlockReference Then
        
        Exit Sub
    Else
        ThisDrawing.Utility.Prompt _
            "�����܂��̓u���b�N���������I������܂���ł����B" & vbCrLf
        Exit Sub
    End If
    
    Call CommonSub.ResetHighlight(targetEntity)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ResetHighlight(targetEntity)
    ThisDrawing.Utility.Prompt "�Ȃ�炩�̃G���[�ł��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## �����ւ̎Q�Ɛ���}
'------------------------------------------------------------------------------
Private Sub addTextReferenceLine(ByRef target_text As ZcadEntity, _
                                 ByVal pick_point As Variant, _
                                 ByVal config_layer As String, _
                                 ByVal config_length As Single, _
                                 ByVal config_offset As Single)
    
    Dim targetAngle As Double
    Dim startPoint(0 To 2) As Double
    Dim endPoint(0 To 2) As Double
    
    ' ��}�ȗ����̂��߂Ɋ�_�Ɗp�x���L�����p�x�v�f�폜
    targetAngle = target_text.Rotation
    target_text.Rotate pick_point, targetAngle * -1
    
    Call calculateEdgePoints _
        (target_text, startPoint, endPoint, config_length, config_offset)
    
    ' �Q�Ɛ���}
    Dim referenceLine As ZcadLine
    Set referenceLine = ThisDrawing.ModelSpace.AddLine(startPoint, endPoint)
    
    ' ��������ю��������̊p�x��߂�
    target_text.Rotate pick_point, targetAngle
    referenceLine.Rotate pick_point, targetAngle
    
    ' ��w�K�p
    referenceLine.Layer = config_layer
    
End Sub

'------------------------------------------------------------------------------
' ## �Q�Ɛ��n�I�[�v�Z
'------------------------------------------------------------------------------
Private Sub calculateEdgePoints(ByVal target_text As ZcadEntity, _
                                ByRef start_point() As Double, _
                                ByRef end_point() As Double, _
                                ByVal config_length As Single, _
                                ByVal config_offset As Single)
    
    Dim minExtent As Variant, maxExtent As Variant
    
    target_text.GetBoundingBox minExtent, maxExtent
    
    ' ZWCAD2020�ł�GetBondingBox�������̌X�Ίp�x�𖳎����Ă��܂�����
    ' �Α̕������l�����X�Ίp�x����MaxPoint���œK������
    Dim textOblique As Double
    Dim deltaX As Double
    Dim deltaY As Double
    textOblique = target_text.ObliqueAngle
    deltaY = maxExtent(1) - minExtent(1)
    deltaX = deltaY * Tan(textOblique)
    If textOblique > 0 Then
        maxExtent(0) = maxExtent(0) + deltaX
    ElseIf textOblique < 0 Then
        minExtent(0) = minExtent(0) - deltaX
    End If
    
    ' �n�[�v�Z
    start_point(0) = minExtent(0) _
        - ((maxExtent(1) - minExtent(1)) * config_length)
    start_point(1) = minExtent(1) - target_text.Height * config_offset
    start_point(2) = 0
    
    ' �I�[�v�Z
    end_point(0) = maxExtent(0) _
        + ((maxExtent(1) - minExtent(1)) * config_length)
    end_point(1) = minExtent(1) - target_text.Height * config_offset
    end_point(2) = 0
    
End Sub
