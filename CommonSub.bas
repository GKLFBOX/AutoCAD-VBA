Attribute VB_Name = "CommonSub"
Option Explicit

'------------------------------------------------------------------------------
' ## �n�C���C�g�̉���
'------------------------------------------------------------------------------
Public Sub ResetHighlight(ByVal target_object As ZcadEntity)

    If Not target_object Is Nothing Then target_object.Highlight False

End Sub

'------------------------------------------------------------------------------
' ## �I���Z�b�g�̍폜
'------------------------------------------------------------------------------
Public Sub ReleaseSelectionSet(ByVal target_selectionset As ZcadSelectionSet)

    If Not target_selectionset Is Nothing Then target_selectionset.Delete

End Sub

'------------------------------------------------------------------------------
' ## �g�u���b�N�̑I��
'------------------------------------------------------------------------------
Public Sub PickFrameBlock(ByRef frame_block As ZcadBlockReference)
    
    Dim pickEntity As ZcadEntity
    Dim pickPoint As Variant
    
    ThisDrawing.Utility.GetEntity pickEntity, pickPoint, _
        "�g�u���b�N��I�� [Cancel(ESC)]"
        
    If TypeOf pickEntity Is ZcadBlockReference Then
        Set frame_block = pickEntity
    Else
        ThisDrawing.Utility.Prompt "�u���b�N�ȊO���I������܂����B"
        Exit Sub
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## �g���̂��擾
'------------------------------------------------------------------------------
Public Sub FetchFrameName(ByVal frame_block As ZcadBlockReference, _
                          ByVal frame_tag As String, _
                          ByRef frame_name As String)
    
    Dim frameAttributes As Variant
    Dim currentAttribute As ZcadAttributeReference
    
    frameAttributes = frame_block.GetAttributes
    
    ' �w�葮���̌���
    Dim i As Long
    For i = 0 To UBound(frameAttributes)
        Set currentAttribute = frameAttributes(i)
        If currentAttribute.TagString = frame_tag Then
            frame_name = currentAttribute.TextString
            Exit Sub
        End If
    Next i
    
    ' �w�葮�������������ꍇ�̓Y�[�����Ē��ړ��͂𑣂�
    Dim minExtent As Variant, maxExtent As Variant
    If frame_name = "" Then
        frame_block.GetBoundingBox minExtent, maxExtent
        ThisDrawing.Application.ZoomWindow minExtent, maxExtent
        frame_name = ThisDrawing.Utility.GetString _
            (0, "�p���g���̂�������Ȃ����ߒ��ړ��� [Cancel(ESC)]:")
        Exit Sub
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## ������`�����������m�ȗp���g�T�C�Y�擾
'------------------------------------------------------------------------------
Public Sub FetchCorrectSize(ByVal frame_block As ZcadBlockReference, _
                            ByRef min_framepoint As Variant, _
                            ByRef max_framepoint As Variant)
    
    Dim i As Long, j As Long
    Dim replicaEntities As Variant
    Dim extractEntity As ZcadEntity
    Dim currentMin As Variant, currentMax As Variant
    
    replicaEntities = frame_block.Explode
    
    ' ������`�������u���b�N�\���v�f����p���g�T�C�Y���擾
    For i = 0 To UBound(replicaEntities)
        
        Set extractEntity = replicaEntities(i)
        If TypeOf extractEntity Is ZcadAttribute Then GoTo Continue_i
        
        ' ��r�X�V���s���ŊO���T�C�Y���擾����
        If IsEmpty(min_framepoint) And IsEmpty(max_framepoint) Then
            extractEntity.GetBoundingBox min_framepoint, max_framepoint
        Else
            extractEntity.GetBoundingBox currentMin, currentMax
            For j = 0 To 1
                If currentMin(j) <= min_framepoint(j) Then
                    min_framepoint(j) = currentMin(j)
                End If
                If currentMax(j) >= max_framepoint(j) Then
                    max_framepoint(j) = currentMax(j)
                End If
            Next j
        End If
        
Continue_i:
        extractEntity.Delete
        
    Next i
    
End Sub

'------------------------------------------------------------------------------
' ## �r���[�|�[�g�̕\������
'------------------------------------------------------------------------------
Public Sub ApplyViewportProperty(ByRef target_viewport As ZcadPViewport, _
                                 ByVal target_layer As String, _
                                 ByVal min_framepoint As Variant, _
                                 ByVal max_framepoint As Variant)
    
    Dim changeColor As ZcadZcCmColor
    
    Set changeColor = New ZcadZcCmColor
    changeColor.ColorIndex = zcByLayer
    
    ' �v���p�e�B�ݒ�
    With target_viewport
        
        .Layer = target_layer
        .TrueColor = changeColor
        .Linetype = "ByLayer"
        .LinetypeScale = 1
        .Lineweight = zcLnWtByLayer
        
    End With
    
    ' �r���[�|�[�g���̕\������
    With ThisDrawing
        
        ' �����Ńy�[�p�[��Ԃ̑S�̕\�������Ȃ���
        ' �Ȃ����r���[�|�[�g���̕\�������Ŏړx���኱�Y���Ă��܂�
        .Application.ZoomExtents
        
        .MSpace = True
        
        .ActivePViewport = target_viewport
        .Application.ZoomWindow min_framepoint, max_framepoint
        
        .MSpace = False
        
    End With
    
End Sub

