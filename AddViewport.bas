Attribute VB_Name = "AddViewport"
Option Explicit

'------------------------------------------------------------------------------
' ## �g�u���b�N�I���ɂ��r���[�|�[�g�ǉ�   2020/07/26 G.O.
'
' ���C�A�E�g�g�u���b�N����Ώۃ��C�A�E�g�Ƀr���[�|�[�g��ǉ�����
'------------------------------------------------------------------------------
Public Sub AddViewport(ByVal frame_tag As String, _
                       ByVal scale_factor As Single, _
                       ByVal viewport_layer As String, _
                       ByVal custom_scale As Single)
    
    On Error GoTo Error_Handler
    
    Dim LayoutFrame As ZcadBlockReference
    Dim minFramePoint As Variant, maxFramePoint As Variant
    Dim frameWidth As Double, frameHeight As Double
    Dim targetLayout As ZcadLayout
    
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts("Model")
    
    ' �g�u���b�N�̑I��
    Call CommonSub.PickFrameBlock(LayoutFrame)
    If LayoutFrame Is Nothing Then Exit Sub
    
    ' �g���̎擾����ёΏۃ��C�A�E�g�̎擾
    Dim frameName As String
    Call CommonSub.FetchFrameName(LayoutFrame, frame_tag, frameName)
    Call fetchTargetLayout(frameName, targetLayout)
    If targetLayout Is Nothing Then Exit Sub
    
    ' �p���g�T�C�Y�擾
    Dim customScale As Single
    customScale = scale_factor / custom_scale
    Call CommonSub.FetchCorrectSize(LayoutFrame, minFramePoint, maxFramePoint)
    frameWidth = (maxFramePoint(0) - minFramePoint(0)) * customScale
    frameHeight = (maxFramePoint(1) - minFramePoint(1)) * customScale
    
    ' �Ώۃ��C�A�E�g�A�N�e�B�u��
    ThisDrawing.ActiveLayout = targetLayout
    
    ' �r���[�|�[�g�}������ђ���
    Dim layoutViewport As ZcadPViewport
    Call insertViewport(frameWidth, frameHeight, layoutViewport)
    Call CommonSub.ApplyViewportProperty _
        (layoutViewport, viewport_layer, minFramePoint, maxFramePoint)
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "�Ȃ�炩�̃G���[�ł��B"
    
End Sub

'------------------------------------------------------------------------------
' ## �Ώۃ��C�A�E�g�̎擾
'------------------------------------------------------------------------------
Private Sub fetchTargetLayout(ByVal frame_name As String, _
                              ByRef target_layout As ZcadLayout)
    
    Dim currentLayout As ZcadLayout
    For Each currentLayout In ThisDrawing.Layouts
        If currentLayout.Name = frame_name Then
            Set target_layout = currentLayout
            Exit Sub
        End If
    Next currentLayout
    
    ThisDrawing.Utility.Prompt "�Ώۃ��C�A�E�g�����݂��܂���B"
    
End Sub

'------------------------------------------------------------------------------
' ## �r���[�|�[�g�}��
'------------------------------------------------------------------------------
Private Sub insertViewport(ByVal frame_width As Double, _
                           ByVal frame_height As Double, _
                           ByRef layout_viewport As ZcadPViewport)
    
    Dim targetPoint As Variant
    Dim viewCenter(0 To 2) As Double
    
    targetPoint = ThisDrawing.Utility.GetPoint _
        (, "�}���_(����)���w�� [Cancel(ESC)]:")
    
    viewCenter(0) = targetPoint(0) + frame_width / 2
    viewCenter(1) = targetPoint(1) + frame_height / 2
    viewCenter(2) = 0
    
    Set layout_viewport = ThisDrawing.PaperSpace.AddPViewport _
        (viewCenter, frame_width, frame_height)
    
End Sub
