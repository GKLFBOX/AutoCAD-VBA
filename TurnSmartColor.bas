Attribute VB_Name = "TurnSmartColor"
Option Explicit

'------------------------------------------------------------------------------
' ## �I�u�W�F�N�g�̎�ނɉ������œK�ȐF�؂�ւ�   2020/08/09 G.O.
'
' �F���ԈȊO�̏ꍇ�͐ԂɕύX���Ԃ̏ꍇ��ByLayer�ɕύX����
' ���L5�O���[�v��Ώۂɂ��ꂼ��ɉ����ĐF�ύX���s��
' 1.[�~��,�~,�ȉ~,�n�b�`���O,2D�|�����C��,��,�}���`�e�L�X�g,
'   �|�����C��,���ː�,�X�v���C��,����,�\�z��]
' 2.[�u���b�N�Q��]
' 3.[3�_�p�x���@,���s���@,�p�x���@,�~�ʂ̒������@,�������@]
' 4.[���a���@,���a���@]
' 5.[���o��]
'------------------------------------------------------------------------------
Public Sub TurnSmartColor()
    
    On Error GoTo Error_Handler
    
    Dim targetSelectionSet As ZcadSelectionSet
    Dim targetEntity As ZcadEntity
    Dim targetLayer As ZcadLayer
    Dim noChange As Long
    
    ThisDrawing.Utility.Prompt _
        "�F�ύX����I�u�W�F�N�g��I�����Ă��������B" & vbCrLf
    
    ' �o�͑Ώۂ�͈͑I��
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    If targetSelectionSet.Count = 0 Then
        Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
        Exit Sub
    End If
    
    noChange = 0
    For Each targetEntity In targetSelectionSet
        
        Set targetLayer = ThisDrawing.Layers.Item(targetEntity.Layer)
        If targetLayer.Lock Then GoTo Continue_targetEntity
        
        If isGroup1(targetEntity) Then
            ' ��,������
            Call turnObjectColor(targetEntity)
        ElseIf isGroup2(targetEntity) Then
            ' �u���b�N�Q��
            Call turnGroup2Color(targetEntity, noChange)
        ElseIf isGroup3(targetEntity) Then
            ' �������@,�p�x���@��
            Call turnGroup3Color(targetEntity)
        ElseIf isGroup4(targetEntity) Then
            ' ���a���@,���a���@
            Call turnGroup4Color(targetEntity)
        ElseIf isGroup5(targetEntity) Then
            ' ���o��
            Call turnGroup5Color(targetEntity)
        Else
            noChange = noChange + 1
        End If
        
Continue_targetEntity:
        
    Next targetEntity
    
    ' �������ʕ\��
    Call displayResult(noChange)
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "�Ȃ�炩�̃G���[�ł��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## �O���[�v1����
'------------------------------------------------------------------------------
Private Function isGroup1(ByVal target_entity As ZcadEntity) As Boolean
    
    If TypeOf target_entity Is ZcadArc _
    Or TypeOf target_entity Is ZcadCircle _
    Or TypeOf target_entity Is ZcadEllipse _
    Or TypeOf target_entity Is ZcadHatch _
    Or TypeOf target_entity Is ZcadLWPolyline _
    Or TypeOf target_entity Is ZcadLine _
    Or TypeOf target_entity Is ZcadMText _
    Or TypeOf target_entity Is ZcadPolyline _
    Or TypeOf target_entity Is ZcadRay _
    Or TypeOf target_entity Is ZcadSpline _
    Or TypeOf target_entity Is ZcadText _
    Or TypeOf target_entity Is ZcadXline Then
        isGroup1 = True
    Else
        isGroup1 = False
    End If
    
End Function

'------------------------------------------------------------------------------
' ## �O���[�v2����
'------------------------------------------------------------------------------
Private Function isGroup2(ByVal target_entity As ZcadEntity) As Boolean
    
    If TypeOf target_entity Is ZcadBlockReference Then
        isGroup2 = True
    Else
        isGroup2 = False
    End If
    
End Function

'------------------------------------------------------------------------------
' ## �O���[�v3����
'------------------------------------------------------------------------------
Private Function isGroup3(ByVal target_entity As ZcadEntity) As Boolean
    
    If TypeOf target_entity Is ZcadDim3PointAngular _
    Or TypeOf target_entity Is ZcadDimAligned _
    Or TypeOf target_entity Is ZcadDimAngular _
    Or TypeOf target_entity Is ZcadDimArcLength _
    Or TypeOf target_entity Is ZcadDimRotated Then
        isGroup3 = True
    Else
        isGroup3 = False
    End If
    
End Function

'------------------------------------------------------------------------------
' ## �O���[�v4����
'------------------------------------------------------------------------------
Private Function isGroup4(ByVal target_entity As ZcadEntity) As Boolean
    
    If TypeOf target_entity Is ZcadDimDiametric _
    Or TypeOf target_entity Is ZcadDimRadial Then
        isGroup4 = True
    Else
        isGroup4 = False
    End If
    
End Function

'------------------------------------------------------------------------------
' ## �O���[�v5����
'------------------------------------------------------------------------------
Private Function isGroup5(ByVal target_entity As ZcadEntity) As Boolean
    
    If TypeOf target_entity Is ZcadLeader Then
        isGroup5 = True
    Else
        isGroup5 = False
    End If
    
End Function

'------------------------------------------------------------------------------
' ## �I�u�W�F�N�g�F�̕ύX
'------------------------------------------------------------------------------
Private Sub turnObjectColor(ByRef target_entity As ZcadEntity)
    
    Dim changeColor As ZcadZcCmColor
    
    ' �F���ԈȊO�̏ꍇ�͐Ԃɂ��Ԃ̏ꍇ��ByLayer�ɂ���
    Set changeColor = New ZcadZcCmColor
    If target_entity.TrueColor.ColorIndex = zcRed Then
        changeColor.ColorIndex = zcByLayer
    Else
        changeColor.ColorIndex = zcRed
    End If
    
    target_entity.TrueColor = changeColor
    
End Sub

'------------------------------------------------------------------------------
' ## �O���[�v2�̐F�ύX
'------------------------------------------------------------------------------
Private Sub turnGroup2Color(ByRef target_block As ZcadBlockReference, _
                            ByRef no_change As Long)
    
    Dim i As Long
    Dim replicaEntities As Variant
    Dim extractEntity As ZcadEntity
    Dim extractLayer As ZcadLayer
    Dim colorFlag As Boolean
    Dim colorByLayer As ZcadZcCmColor
    
    replicaEntities = target_block.Explode
    
    ' �u���b�N���I�u�W�F�N�g�̑���
    colorFlag = False
    For i = 0 To UBound(replicaEntities)
        Set extractEntity = replicaEntities(i)
        ' �F��ByBlock�̃I�u�W�F�N�g����ł�����ꍇ�͐F�ύX���s��
        If extractEntity.TrueColor.ColorIndex = zcByBlock Then colorFlag = True
        Set extractLayer = ThisDrawing.Layers.Item(extractEntity.Layer)
        With extractLayer
            If .Lock Then
                .Lock = False
                extractEntity.Delete
                .Lock = True
            Else
                extractEntity.Delete
            End If
        End With
    Next i
    
    Set colorByLayer = New ZcadZcCmColor
    
    If colorFlag Then
        Call turnObjectColor(target_block)
    Else
        colorByLayer.ColorIndex = zcByLayer
        target_block.TrueColor = colorByLayer
        no_change = no_change + 1
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## �O���[�v3�̐F�ύX
'------------------------------------------------------------------------------
Private Sub turnGroup3Color(ByRef target_entity As ZcadEntity)
    
    Dim colorByLayer As ZcadZcCmColor
    
    Set colorByLayer = New ZcadZcCmColor
    
    ' �⏕���L�萡�@�n�I�u�W�F�N�g�p�̐F�ύX
    With target_entity
        
        ' �����̐F�̂݃I�u�W�F�N�g�F���p��������
        If .TextColor = zcByBlock Then
            ' ByBlock�̏ꍇ�͐F�ύX
            Call turnObjectColor(target_entity)
        Else
            ' ByBlock�ȊO�̏ꍇ�I�u�W�F�N�g�F���p������悤�ɕύX
            If .TextColor = zcByLayer Then
                ' ByLayer�̏ꍇ�I�u�W�F�N�g�F��Ԃɂ���
                colorByLayer.ColorIndex = zcRed
            Else
                colorByLayer.ColorIndex = zcByLayer
            End If
            .TrueColor = colorByLayer
            .TextColor = zcByBlock
        End If
        
        ' �����̐F�ȊO�͉�w�F���p������悤�ɕύX
        .DimensionLineColor = zcByLayer
        .ExtensionLineColor = zcByLayer
        
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## �O���[�v4�̐F�ύX
'------------------------------------------------------------------------------
Private Sub turnGroup4Color(ByRef target_entity As ZcadEntity)
    
    Dim colorByLayer As ZcadZcCmColor
    
    Set colorByLayer = New ZcadZcCmColor
    
    ' �⏕���������@�n�I�u�W�F�N�g�p�̐F�ύX
    With target_entity
        
        ' �����̐F�̂݃I�u�W�F�N�g�F���p��������
        If .TextColor = zcByBlock Then
            ' ByBlock�̏ꍇ�͐F�ύX
            Call turnObjectColor(target_entity)
        Else
            ' ByBlock�ȊO�̏ꍇ�I�u�W�F�N�g�F���p������悤�ɕύX
            If .TextColor = zcByLayer Then
                ' ByLayer�̏ꍇ�I�u�W�F�N�g�F��Ԃɂ���
                colorByLayer.ColorIndex = zcRed
            Else
                colorByLayer.ColorIndex = zcByLayer
            End If
            .TrueColor = colorByLayer
            .TextColor = zcByBlock
        End If
        
        ' �����̐F�ȊO�͉�w�F���p������悤�ɕύX
        .DimensionLineColor = zcByLayer
        
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## �O���[�v5�̐F�ύX
'------------------------------------------------------------------------------
Private Sub turnGroup5Color(ByRef target_entity As ZcadLeader)
    
    Dim colorByLayer As ZcadZcCmColor
    
    Set colorByLayer = New ZcadZcCmColor
    
    ' ���o���I�u�W�F�N�g�p�̐F�ύX
    With target_entity
        
        ' �I�u�W�F�N�g�F���p��������
        If .DimensionLineColor = zcByBlock Then
            ' ByBlock�̏ꍇ�͐F�ύX
            Call turnObjectColor(target_entity)
        Else
            ' ByBlock�ȊO�̏ꍇ�I�u�W�F�N�g�F���p������悤�ɕύX
            If .DimensionLineColor = zcByLayer Then
                ' ByLayer�̏ꍇ�I�u�W�F�N�g�F��Ԃɂ���
                colorByLayer.ColorIndex = zcRed
            Else
                colorByLayer.ColorIndex = zcByLayer
            End If
            .TrueColor = colorByLayer
            .DimensionLineColor = zcByBlock
        End If
        
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## �������ʕ\��
'------------------------------------------------------------------------------
Private Sub displayResult(ByVal no_change As Long)
    
    Dim resultText As String
    
    resultText = "�I���I�u�W�F�N�g�̐F��؂�ւ��܂����B" & vbCrLf
    
    If no_change > 0 Then
        resultText = resultText _
            & "(�ΏۊO�I�u�W�F�N�g��" & no_change & "����܂����B)" & vbCrLf
    End If
    
    ThisDrawing.Utility.Prompt resultText
    
End Sub
