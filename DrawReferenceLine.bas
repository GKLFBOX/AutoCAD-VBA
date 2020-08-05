Attribute VB_Name = "DrawReferenceLine"
Option Explicit

'------------------------------------------------------------------------------
' ## �����n�I�u�W�F�N�g�ւ̎Q�Ɛ���}   2020/08/03 G.O.
'
' �w�肵�������ƃI�t�Z�b�g�W������Q�Ɛ�����}����
'------------------------------------------------------------------------------
Public Sub DrawReferenceLine()
    
    On Error GoTo Error_Handler
    
    Dim targetEntity As ZcadEntity
    Dim pickPoint As Variant
    Dim configLayerOn As Boolean
    Dim configLayer As String
    Dim configLength As Single
    Dim configOffset As Single
    
    ' �ݒ�l�ǂݍ���
    Dim configData As Variant
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.REFERENCELINE_CONFIG), vbCrLf)
    configLayerOn = configData(0)
    configLayer = configData(1)
    configLength = configData(2) / 2
    configOffset = configData(3)
    
    ' �Ώە����n�I�u�W�F�N�g�̑I��
    ThisDrawing.Utility.GetEntity targetEntity, pickPoint, _
        "�Q�Ɛ������������܂��̓u���b�N��������I�� [Cancel(ESC)]"
    
    ' �e�L�X�g�܂��̓u���b�N�Q�Ƃ̔���
    If CommonFunction.IsTextObject(targetEntity) Then
        Call addTextReferenceLine _
            (targetEntity, pickPoint, configLayerOn, _
            configLayer, configLength, configOffset)
    ElseIf TypeOf targetEntity Is ZcadBlockReference Then
        Call addBlockReferenceLine _
            (targetEntity, pickPoint, configLayerOn, _
            configLayer, configLength, configOffset)
    Else
        ThisDrawing.Utility.Prompt _
            "�����܂��̓u���b�N���������I������܂���ł����B" & vbCrLf
    End If
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "�Ȃ�炩�̃G���[�ł��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## �����ւ̎Q�Ɛ���}
'------------------------------------------------------------------------------
Private Sub addTextReferenceLine(ByRef target_text As ZcadEntity, _
                                 ByVal pick_point As Variant, _
                                 ByVal config_layeron As Boolean, _
                                 ByVal config_layer As String, _
                                 ByVal config_length As Single, _
                                 ByVal config_offset As Single)
    
    Dim targetAngle As Double
    Dim startPoint(0 To 2) As Double
    Dim endPoint(0 To 2) As Double
    Dim referenceLine As ZcadLine
    
    ' ��}�ȗ����̂��߂Ɋ�_�Ɗp�x���L�����p�x�v�f�폜
    targetAngle = target_text.Rotation
    target_text.Rotate pick_point, targetAngle * -1
    
    ' �Q�Ɛ��n�I�[�Z�o
    Call getReferenceLineEdge _
        (target_text, startPoint, endPoint, config_length, config_offset)
    
    ' �Q�Ɛ���}
    Set referenceLine = ThisDrawing.ModelSpace.AddLine(startPoint, endPoint)
    
    ' ��������юQ�Ɛ��̊p�x��߂�
    target_text.Rotate pick_point, targetAngle
    referenceLine.Rotate pick_point, targetAngle
    
    ' ��w�K�p
    If config_layeron Then referenceLine.Layer = config_layer
    
End Sub

'------------------------------------------------------------------------------
' ## �u���b�N�ւ̎Q�Ɛ���}
'------------------------------------------------------------------------------
Private Sub addBlockReferenceLine(ByRef target_block As ZcadBlockReference, _
                                  ByVal pick_point As Variant, _
                                  ByVal config_layeron As Boolean, _
                                  ByVal config_layer As String, _
                                  ByVal config_length As Single, _
                                  ByVal config_offset As Single)
    
    On Error GoTo Error_Handler
    
    Dim replicaEntities As Variant
    Dim targetReplica As ZcadEntity
    Dim targetAngle As Double
    Dim startPoint(0 To 2) As Double, endPoint(0 To 2) As Double
    Dim referenceLine As ZcadLine
    
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
    
    ' �Q�Ɛ��n�I�[�Z�o
    Call getReferenceLineEdge _
        (targetReplica, startPoint, endPoint, config_length, config_offset)
    
    ' �Q�Ɛ���}
    Set referenceLine = ThisDrawing.ModelSpace.AddLine(startPoint, endPoint)
    
    ' ��������юQ�Ɛ��̊p�x��߂�
    target_block.Rotate pick_point, targetAngle
    referenceLine.Rotate pick_point, targetAngle
    
    ' ��w�K�p
    If config_layeron Then referenceLine.Layer = config_layer
    
    Call CommonSub.DeleteReplica(replicaEntities)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.DeleteReplica(replicaEntities)
    
End Sub

'------------------------------------------------------------------------------
' ## �Q�Ɛ��n�I�[�v�Z
'------------------------------------------------------------------------------
Private Sub getReferenceLineEdge(ByVal target_text As ZcadEntity, _
                                 ByRef start_point() As Double, _
                                 ByRef end_point() As Double, _
                                 ByVal config_length As Single, _
                                 ByVal config_offset As Single)
    
    Dim minExtent As Variant, maxExtent As Variant
    
    ' �g����GetBoundingBox
    Call CommonSub.GetEnhancedBoundingBox(target_text, minExtent, maxExtent)
    
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
