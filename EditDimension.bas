Attribute VB_Name = "EditDimension"
'------------------------------------------------------------------------------
' ## �R�[�f�B���O�K�C�h���C��
'
' [You.Activate|VBA�R�[�f�B���O�K�C�h���C��]�ɏ�������
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## ���@�I�u�W�F�N�g�̕���(���@�l)�̐F�݂̂�Ԃɂ���
'
' ���@�X�^�C���Ɉˑ������ɐ��@�l�̕����F�݂̂�F�ύX�Ŗ߂����ԂŐԂɂ���
' �I�u�W�F�N�g=��, ����(���@�l)=ByBlock, ���@������ѐ��@�⏕��=ByLayer
'------------------------------------------------------------------------------
Public Sub TurnRedDimension()
    
    On Error GoTo Error_Handler
    
    ' TODO: �\�ł����Lisp�𗘗p���Ď��O�I������������
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    If Not targetSelectionSet.Count = 0 Then
        
        Dim returnObject As ZcadEntity
        For Each returnObject In targetSelectionSet
            If TypeOf returnObject Is ZcadDimension Then
                returnObject.TextColor = zcByBlock
                returnObject.Color = zcRed  ' TODO: color�͎g���ׂ��łȂ�
            End If
        Next returnObject
        
    End If
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "�G���[�F�R�}���h���I�����܂��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## ���@���܂��͈��o���̃T�C�Y�ύX
'
' ���@���܂��͈��o���̑S�̂̐��@�ړx��ύX����
'------------------------------------------------------------------------------
Public Sub ResizeDimensionSize()
    
    On Error GoTo Error_Handler
    
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    If Not targetSelectionSet.Count = 0 Then
        
        Dim returnObject As ZcadEntity
        For Each returnObject In targetSelectionSet
            If TypeOf returnObject Is ZcadDimension _
            Or TypeOf returnObject Is ZcadLeader Then
                returnObject.Highlight True
            End If
        Next returnObject
        
        Dim sizeFactor As Long
        sizeFactor = ThisDrawing.Utility.GetInteger _
            ("�ύX�ړx����� �܂��� [25/50/80/100]:")
        
        For Each returnObject In targetSelectionSet
            If TypeOf returnObject Is ZcadDimension _
            Or TypeOf returnObject Is ZcadLeader Then
                returnObject.ScaleFactor = sizeFactor
            End If
        Next returnObject
        
    End If
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    Exit Sub
    
Error_Handler:
    For Each returnObject In targetSelectionSet
        Call CommonSub.ResetHighlight(returnObject)
    Next returnObject
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "�G���[�F�R�}���h���I�����܂��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## ���@���̕����I�t�Z�b�g�ʕύX
'
' ���@���̕����I�t�Z�b�g��(���@���ƕ����̗����)��ύX����
'------------------------------------------------------------------------------
Public Sub AdjustDimensionOffset()
    
    On Error GoTo Error_Handler
    
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    If Not targetSelectionSet.Count = 0 Then
        
        Dim returnObject As ZcadEntity
        For Each returnObject In targetSelectionSet
            If TypeOf returnObject Is ZcadDimension Then
                returnObject.Highlight True
            End If
        Next returnObject
        
        Dim offsetAmount As Double
        offsetAmount = ThisDrawing.Utility.GetInteger _
            ("�ύX�I�t�Z�b�g�ʂ���� �܂��� [�f�t�H���g(8)/������(5)]:")
        offsetAmount = offsetAmount * 0.1
        
        For Each returnObject In targetSelectionSet
            If TypeOf returnObject Is ZcadDimension Then
                returnObject.TextGap = offsetAmount
            End If
        Next returnObject
        
    End If
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    Exit Sub
    
Error_Handler:
    For Each returnObject In targetSelectionSet
        Call CommonSub.ResetHighlight(returnObject)
    Next returnObject
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "�G���[�F�R�}���h���I�����܂��B" & vbCrLf
    
End Sub
