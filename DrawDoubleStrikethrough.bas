Attribute VB_Name = "DrawDoubleStrikethrough"
Option Explicit

'------------------------------------------------------------------------------
' ## �����n�I�u�W�F�N�g�ւ̓�d����������}   2020/08/03 G.O.
'
' �w�肵�������I�u�W�F�N�g�ɓ���w�œ�d������������}����
'------------------------------------------------------------------------------
Public Sub DrawDoubleStrikethrough()
    
    On Error GoTo Error_Handler
    
    Dim targetEntity As ZcadEntity
    Dim pickPoint As Variant
    Dim configLayerOn As Boolean
    Dim configLayer As String
    Dim configLength As Single
    Dim configRed As Boolean
    Dim configTargetLayer As Boolean
    
    ' �ݒ�l�ǂݍ���
    Dim configData As Variant
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.STRIKETHROUGH_CONFIG), vbCrLf)
    configLayerOn = configData(0)
    configLayer = configData(1)
    configLength = configData(2) / 2
    configRed = configData(3)
    configTargetLayer = configData(4)
    
    ' �Ώە����n�I�u�W�F�N�g�̑I��
    ThisDrawing.Utility.GetEntity targetEntity, pickPoint, _
        "��d�������������������܂��̓u���b�N��������I�� [Cancel(ESC)]"
    
    ' �e�L�X�g�܂��̓u���b�N�Q�Ƃ̔���
    If CommonFunction.IsTextObject(targetEntity) Then
        Call addTextStrikethrough _
            (targetEntity, pickPoint, configLayerOn, _
            configLayer, configLength, configRed, configTargetLayer)
    ElseIf TypeOf targetEntity Is ZcadBlockReference Then
        Call addBlockStrikethrough _
            (targetEntity, pickPoint, configLayerOn, _
            configLayer, configLength, configRed, configTargetLayer)
    Else
        ThisDrawing.Utility.Prompt _
            "�����܂��̓u���b�N���������I������܂���ł����B" & vbCrLf
    End If
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "�Ȃ�炩�̃G���[�ł��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## �����ւ̓�d����������}
'------------------------------------------------------------------------------
Private Sub addTextStrikethrough(ByRef target_text As ZcadEntity, _
                                 ByVal pick_point As Variant, _
                                 ByVal config_layeron As Boolean, _
                                 ByVal config_layer As String, _
                                 ByVal config_length As Single, _
                                 ByVal config_red As Boolean, _
                                 ByVal config_targetlayer As Boolean)
    
    Dim targetAngle As Double
    Dim startPoint1(0 To 2) As Double, endPoint1(0 To 2) As Double
    Dim startPoint2(0 To 2) As Double, endPoint2(0 To 2) As Double
    Dim strikeThrough1 As ZcadLine, strikeThrough2 As ZcadLine
    
    ' ��}�ȗ����̂��߂Ɋ�_�Ɗp�x���L�����p�x�v�f�폜
    targetAngle = target_text.Rotation
    target_text.Rotate pick_point, targetAngle * -1
    
    ' ���������n�I�[�Z�o
    Call getStrikethroughEdge(target_text, startPoint1, endPoint1, _
        startPoint2, endPoint2, config_length)
    
    ' ����������}
    Set strikeThrough1 = ThisDrawing.ModelSpace.AddLine(startPoint1, endPoint1)
    Set strikeThrough2 = ThisDrawing.ModelSpace.AddLine(startPoint2, endPoint2)
    
    ' ��������ю��������̊p�x��߂�
    target_text.Rotate pick_point, targetAngle
    strikeThrough1.Rotate pick_point, targetAngle
    strikeThrough2.Rotate pick_point, targetAngle
    
    ' ��}�ݒ�̓K�p
    Call applyDrawingConfig(target_text, strikeThrough1, strikeThrough2, _
        config_layeron, config_layer, config_red, config_targetlayer)
    
End Sub

'------------------------------------------------------------------------------
' ## �u���b�N�ւ̓�d����������}
'------------------------------------------------------------------------------
Private Sub addBlockStrikethrough(ByRef target_block As ZcadBlockReference, _
                                  ByVal pick_point As Variant, _
                                  ByVal config_layeron As Boolean, _
                                  ByVal config_layer As String, _
                                  ByVal config_length As Single, _
                                  ByVal config_red As Boolean, _
                                  ByVal config_targetlayer As Boolean)
    
    On Error GoTo Error_Handler
    
    Dim replicaEntities As Variant
    Dim targetReplica As ZcadEntity
    Dim targetAngle As Double
    Dim startPoint1(0 To 2) As Double, endPoint1(0 To 2) As Double
    Dim startPoint2(0 To 2) As Double, endPoint2(0 To 2) As Double
    Dim strikeThrough1 As ZcadLine, strikeThrough2 As ZcadLine
    
    replicaEntities = target_block.Explode
    
    ' �����I�u�W�F�N�g�̑�����`���̒u��
    Call CommonSub.ReplaceAttributeTag(target_block, replicaEntities)
    
    ' �w��_�̕����I�u�W�F�N�g���擾
    Call CommonSub.GrabReplicaEntity(pick_point, targetReplica)
    
    ' ��ʏ�ł͕����I�u�W�F�N�g���\����
    Call CommonSub.HideReplica(replicaEntities)
    
    ' �e�L�X�g�������̔���
    If Not CommonFunction.IsTextObject(targetReplica) _
    And Not TypeOf targetReplica Is ZcadAttribute Then
        Call CommonSub.DeleteReplica(replicaEntities)
        ThisDrawing.Utility.Prompt _
            "�u���b�N���������I������܂���ł����B" & vbCrLf
        Exit Sub
    End If
    
    ' ��}�ȗ����̂��߂Ɋ�_�Ɗp�x���L�����p�x�v�f�폜
    targetAngle = targetReplica.Rotation
    targetReplica.Rotate pick_point, targetAngle * -1
    target_block.Rotate pick_point, targetAngle * -1
    
    ' ���������n�I�[�Z�o
    Call getStrikethroughEdge(targetReplica, startPoint1, endPoint1, _
        startPoint2, endPoint2, config_length)
    
    ' ����������}
    Set strikeThrough1 = ThisDrawing.ModelSpace.AddLine(startPoint1, endPoint1)
    Set strikeThrough2 = ThisDrawing.ModelSpace.AddLine(startPoint2, endPoint2)
    
    ' ��������ю��������̊p�x��߂�
    target_block.Rotate pick_point, targetAngle
    strikeThrough1.Rotate pick_point, targetAngle
    strikeThrough2.Rotate pick_point, targetAngle
    
    ' ��}�ݒ�̓K�p
    Call applyDrawingConfig(target_block, strikeThrough1, strikeThrough2, _
        config_layeron, config_layer, config_red, config_targetlayer)
    
    Call CommonSub.DeleteReplica(replicaEntities)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.DeleteReplica(replicaEntities)
    
End Sub

'------------------------------------------------------------------------------
' ## �Q�Ɛ��n�I�[�v�Z
'------------------------------------------------------------------------------
Private Sub getStrikethroughEdge(ByVal target_text As ZcadEntity, _
                                 ByRef start_point1() As Double, _
                                 ByRef end_point1() As Double, _
                                 ByRef start_point2() As Double, _
                                 ByRef end_point2() As Double, _
                                 ByVal config_length As Single)
    
    Dim minExtent As Variant, maxExtent As Variant
    
    ' �g����GetBoundingBox
    Call CommonSub.GetEnhancedBoundingBox(target_text, minExtent, maxExtent)
    
    ' �n�[�v�Z
    start_point1(0) = minExtent(0) _
        - ((maxExtent(1) - minExtent(1)) * config_length)
    start_point1(1) = minExtent(1) + ((maxExtent(1) - minExtent(1)) / 3)
    start_point1(2) = 0
    
    ' �I�[�v�Z
    end_point1(0) = maxExtent(0) _
        + ((maxExtent(1) - minExtent(1)) * config_length)
    end_point1(1) = minExtent(1) + ((maxExtent(1) - minExtent(1)) / 3)
    end_point1(2) = 0
    
    ' �n�[�v�Z2
    start_point2(0) = minExtent(0) _
        - ((maxExtent(1) - minExtent(1)) * config_length)
    start_point2(1) = minExtent(1) + ((maxExtent(1) - minExtent(1)) * 2 / 3)
    start_point2(2) = 0
    
    ' �I�[�v�Z2
    end_point2(0) = maxExtent(0) _
        + ((maxExtent(1) - minExtent(1)) * config_length)
    end_point2(1) = minExtent(1) + ((maxExtent(1) - minExtent(1)) * 2 / 3)
    end_point2(2) = 0
    
End Sub

'------------------------------------------------------------------------------
' ## ��}�ݒ�̓K�p
'------------------------------------------------------------------------------
Private Sub applyDrawingConfig(ByRef target_text As ZcadEntity, _
                               ByRef strike_through1 As ZcadLine, _
                               ByRef strike_through2 As ZcadLine, _
                               ByVal config_layeron As Boolean, _
                               ByVal config_layer As String, _
                               ByVal config_red As Boolean, _
                               ByVal config_targetlayer As Boolean)
    
    ' ��}��w�̓K�p����ю������Ώۉ�w�ύX�̓K�p
    If config_layeron Then
        strike_through1.Layer = config_layer
        strike_through2.Layer = config_layer
        If config_targetlayer Then target_text.Layer = config_layer
    Else
        strike_through1.Layer = target_text.Layer
        strike_through2.Layer = target_text.Layer
    End If
    
    ' ���������Ԓ��F�̓K�p
    Dim changeColor As ZcadZcCmColor
    If config_red Then
        Set changeColor = New ZcadZcCmColor
        changeColor.ColorIndex = zcRed
        strike_through1.TrueColor = changeColor
        strike_through2.TrueColor = changeColor
    End If
    
End Sub
