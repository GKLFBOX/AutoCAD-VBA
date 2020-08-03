Attribute VB_Name = "AlignTextGroupCenter"
Option Explicit

'------------------------------------------------------------------------------
' ## �����n�I�u�W�F�N�g�ʒu����   2020/08/02 G.O.
'
' �w�肵��2�_�ƃI�t�Z�b�g�W�����當���n�I�u�W�F�N�g�̈ʒu�𒆉������ɒ�������
'------------------------------------------------------------------------------
Public Sub AlignTextGroupCenter()
    
    On Error GoTo Error_Handler
    
    Dim targetEntity As ZcadEntity
    Dim pickPoint As Variant
    
    ' �ʒu�������镶���n�I�u�W�F�N�g�̑I��
    ThisDrawing.Utility.GetEntity targetEntity, pickPoint, _
        "�ʒu�������镶���܂��̓u���b�N��������I�� [Cancel(ESC)]"
    
    targetEntity.Highlight True
    
    ' �e�L�X�g�܂��̓u���b�N�Q�Ƃ̔���
    If CommonFunction.IsTextObject(targetEntity) Then
        Call alignTextCenter(targetEntity, pickPoint)
    ElseIf TypeOf targetEntity Is ZcadBlockReference Then
        Call alignBlockCenter(targetEntity, pickPoint)
    Else
        ThisDrawing.Utility.Prompt _
            "�����܂��̓u���b�N���������I������܂���ł����B" & vbCrLf
    End If
    
    Call CommonSub.ResetHighlight(targetEntity)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ResetHighlight(targetEntity)
    ThisDrawing.Utility.Prompt "�Ȃ�炩�̃G���[�ł��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## �����ʒu����
'------------------------------------------------------------------------------
Private Sub alignTextCenter(ByRef target_text As ZcadEntity, _
                            ByVal pick_point As Variant)
    
    On Error GoTo Error_Handler
    
    Dim firstPoint As Variant, secondPoint As Variant
    Dim offsetFactor As String
    Dim underFlag As String
    Dim textCenter() As Double
    
    ' �����l�̃��[�U�[����
    Call inputAlignValue(firstPoint, secondPoint, offsetFactor, underFlag)
    
    ' �I�t�Z�b�g�v�Z�ȗ����̂��ߊp�x�v�f�폜
    Dim targetAngle As Double
    targetAngle = target_text.Rotation
    target_text.Rotate pick_point, targetAngle * -1
    
    ' �����̏㉺���S�擾����ю擾�ʒu�̃I�t�Z�b�g
    textCenter = getTextCenter(target_text, underFlag)
    Call offsetTextCenter(textCenter, underFlag, offsetFactor, target_text)
    
    ' �����ʒu�����̎��s
    Call doAlignment(firstPoint, secondPoint, textCenter, target_text)
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "�Ȃ�炩�̃G���[�ł��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## �u���b�N�ʒu����
'------------------------------------------------------------------------------
Private Sub alignBlockCenter(ByRef target_block As ZcadBlockReference, _
                             ByVal pick_point As Variant)
    
    On Error GoTo Error_Handler
    
    Dim replicaEntities As Variant
    Dim targetReplica As ZcadEntity
    Dim firstPoint As Variant, secondPoint As Variant
    Dim offsetFactor As String
    Dim underFlag As String
    Dim textCenter() As Double
    
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
    
    ' �����l�̃��[�U�[����
    Call inputAlignValue(firstPoint, secondPoint, offsetFactor, underFlag)
    
    ' �I�t�Z�b�g�v�Z�ȗ����̂��ߊp�x�v�f�폜
    Dim targetAngle As Double
    targetAngle = targetReplica.Rotation
    targetReplica.Rotate pick_point, targetAngle * -1
    target_block.Rotate pick_point, targetAngle * -1
    
    ' �����̏㉺���S�擾����ю擾�ʒu�̃I�t�Z�b�g
    textCenter = getTextCenter(targetReplica, underFlag)
    Call offsetTextCenter(textCenter, underFlag, offsetFactor, targetReplica)
    
    ' �����ʒu�����̎��s
    Call doAlignment(firstPoint, secondPoint, textCenter, target_block)
    
    Call CommonSub.DeleteReplica(replicaEntities)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.DeleteReplica(replicaEntities)
    ThisDrawing.Utility.Prompt "�Ȃ�炩�̃G���[�ł��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## �����l�̃��[�U�[����
'------------------------------------------------------------------------------
Private Sub inputAlignValue(ByRef first_point As Variant, _
                            ByRef second_point As Variant, _
                            ByRef offset_factor As String, _
                            ByRef under_flag As String)
    
    ' �������2�_���w��
    first_point = ThisDrawing.Utility.GetPoint _
        (, "1�_�ڂ��w�� [Cancel(ESC)]")
    second_point = ThisDrawing.Utility.GetPoint _
        (first_point, "2�_�ڂ��w�� [Cancel(ESC)]")
    
    ' �I�t�Z�b�g�W���̓���
    ' ZWCAD�̕s���Get�n��Prompt�ɑg�ݍ��񂾒l��
    ' ���R���܂��͉p��(�啶��)��������ɓ��͂���Ȃ����Ƃ��l�����Ă���
    offset_factor = ThisDrawing.Utility.GetString _
        (0, "�I�t�Z�b�g�W�������(���������ɑ΂��銄��(x/10)) " & _
        "[�ʏ�(2)/�L��(3)/����(1)/���L��(5)]:")
    offset_factor = offset_factor * 0.1
    
    ' ���t���̑I��
    under_flag = ThisDrawing.Utility.GetString _
        (0, "���t���ɂ��܂���? [�͂�(Y)/������(N)]:")
    
End Sub

'------------------------------------------------------------------------------
' ## �����̏㉺���S�擾
'------------------------------------------------------------------------------
Private Function getTextCenter(ByVal target_text As ZcadEntity, _
                               ByVal under_flag As String) As Double()
    
    Dim minExtent As Variant, maxExtent As Variant
    Dim leftPoint(0 To 2) As Double, rightPoint(0 To 2) As Double
    
    target_text.GetBoundingBox minExtent, maxExtent
    
    ' �g����GetBoundingBox
    Call CommonSub.GetEnhancedBoundingBox(target_text, minExtent, maxExtent)
    
    ' �㋫�E�܂��͉����E�̎擾
    If UCase(under_flag) = "Y" Then
        leftPoint(0) = minExtent(0): leftPoint(1) = maxExtent(1)
        rightPoint(0) = maxExtent(0): rightPoint(1) = maxExtent(1)
    Else
        leftPoint(0) = minExtent(0): leftPoint(1) = minExtent(1)
        rightPoint(0) = maxExtent(0): rightPoint(1) = minExtent(1)
    End If
    leftPoint(2) = 0
    rightPoint(2) = 0
    
    getTextCenter = getMiddlePoint(leftPoint, rightPoint)
    
End Function

'------------------------------------------------------------------------------
' ## ���S�ʒu�̏㉺�I�t�Z�b�g
'------------------------------------------------------------------------------
Private Sub offsetTextCenter(ByRef text_center() As Double, _
                             ByVal under_flag As String, _
                             ByVal offset_factor As String, _
                             ByVal target_text As ZcadEntity)
    
    If UCase(under_flag) = "Y" Then
        text_center(1) = text_center(1) _
            + target_text.Height * Abs(offset_factor)
    Else
        text_center(1) = text_center(1) _
            - target_text.Height * Abs(offset_factor)
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## �ʒu�������s
'------------------------------------------------------------------------------
Private Sub doAlignment(ByVal first_point As Variant, _
                        ByVal second_point As Variant, _
                        ByRef text_center() As Double, _
                        ByRef target_entity As ZcadEntity)
    
    Dim alignPoint() As Double
    Dim alignRadian As Double
    
    alignPoint = getMiddlePoint(first_point, second_point)
    alignRadian = calculateAngle(first_point, second_point)
    
    target_entity.Move text_center, alignPoint
    target_entity.Rotate alignPoint, alignRadian
    
End Sub

'------------------------------------------------------------------------------
' ## 2�_�̒��_�擾
'------------------------------------------------------------------------------
Private Function getMiddlePoint(ByRef first_point As Variant, _
                                ByRef second_point As Variant) As Double()
    
    Dim i As Long
    Dim middlePoint(0 To 2) As Double
    
    For i = 0 To 2
        middlePoint(i) = (first_point(i) + second_point(i)) / 2
    Next i
    
    getMiddlePoint = middlePoint()
    
End Function

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
    ElseIf delta_x > 0 And delta_y = 0 Then
        ' ��=0
        Atn2 = (pi / 2) * 0
    ElseIf delta_x = 0 And delta_y > 0 Then
        ' ��=90
        Atn2 = (pi / 2) * 1
    ElseIf delta_x < 0 And delta_y = 0 Then
        ' ��=180
        Atn2 = (pi / 2) * 2
    ElseIf delta_x = 0 And delta_y < 0 Then
        ' ��=270
        Atn2 = (pi / 2) * 3
    ElseIf delta_x > 0 And delta_y > 0 Then
        ' 0<��<90
        Atn2 = Atn(Abs(delta_y) / Abs(delta_x)) + ((pi / 2) * 0)
    ElseIf delta_x < 0 And delta_y > 0 Then
        ' 90<��<180
        Atn2 = ((pi / 2) - Atn(Abs(delta_y) / Abs(delta_x))) + ((pi / 2) * 1)
    ElseIf delta_x < 0 And delta_y < 0 Then
        ' 180<��<270
        Atn2 = Atn(Abs(delta_y) / Abs(delta_x)) + ((pi / 2) * 2)
    ElseIf delta_x > 0 And delta_y < 0 Then
        ' 90<��<180
        Atn2 = ((pi / 2) - Atn(Abs(delta_y) / Abs(delta_x))) + ((pi / 2) * 3)
    End If
    
End Function
