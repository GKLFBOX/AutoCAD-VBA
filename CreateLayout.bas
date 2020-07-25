Attribute VB_Name = "CreateLayout"
Option Explicit

'------------------------------------------------------------------------------
' ## �g�u���b�N�I���ɂ�郌�C�A�E�g�쐬   2020/07/25 G.O.
'
' �p���g�u���b�N���烌�C�A�E�g�𐶐�����
' �g�O�Ƀf�[�^��ł̎��ʗp�}�[�N���������Ă��邱�Ƃ��l������
' �p���g�T�C�Y�͑�����`���������T�C�Y���擾����
'------------------------------------------------------------------------------
Public Sub CreateLayout(ByVal frame_tag As String, _
                        ByVal scale_factor As Single, _
                        ByVal viewport_layer As String, _
                        ByVal style_name As String, _
                        ByVal printer_name As String, _
                        ByVal a3_paper As String, _
                        ByVal a4_paper As String, _
                        ByVal offset_x As Single, _
                        ByVal offset_y As Single)
    
    On Error GoTo Error_Handler
    
    ' �g�̕�����э����͐��}�덷���l������Single�Ƃ��Ă���
    Dim paperFrame As ZcadBlockReference
    Dim minFramePoint As Variant, maxFramePoint As Variant
    Dim frameWidth As Single, frameHeight As Single
    Dim newLayout As ZcadLayout
    
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts("Model")
    
    ' �g�u���b�N�̑I��
    Call CommonSub.PickFrameBlock(paperFrame)
    If paperFrame Is Nothing Then Exit Sub
    
    ' �g���̎擾����у��C�A�E�g�쐬
    Dim frameName As String
    Call CommonSub.FetchFrameName(paperFrame, frame_tag, frameName)
    Set newLayout = ThisDrawing.Layouts.Add(frameName)
    
    ' �p���g�T�C�Y�擾
    Call CommonSub.FetchCorrectSize(paperFrame, minFramePoint, maxFramePoint)
    frameWidth = maxFramePoint(0) - minFramePoint(0)
    frameHeight = maxFramePoint(1) - minFramePoint(1)
    
    ' �V�K���C�A�E�g�A�N�e�B�u��
    ThisDrawing.ActiveLayout = newLayout
    
    ' �r���[�|�[�g����
    ' �V�KPaperSpace��Item(0)�̓��C�A�E�g�̉�ʂ��̂��̂ł���
    ' ���[�U�[���F�����Ă���r���[�|�[�g�I�u�W�F�N�g��Item(1)�̂��ߒ���
    Dim paperViewport As ZcadPViewport
    Set paperViewport = ThisDrawing.PaperSpace.Item(1)
    Call adjustViewportShape(paperViewport, frameWidth, frameHeight)
    Call CommonSub.ApplyViewportProperty _
        (paperViewport, viewport_layer, minFramePoint, maxFramePoint)
    
    ' ����ݒ�
    Call configurePrintSettings(newLayout, frameWidth, frameHeight, _
        scale_factor, style_name, printer_name, a3_paper, a4_paper, _
        offset_x, offset_y)
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "�Ȃ�炩�̃G���[�ł��B"
    
End Sub

'------------------------------------------------------------------------------
' ## �r���[�|�[�g�̈ʒu����уT�C�Y����
'------------------------------------------------------------------------------
Private Sub adjustViewportShape(ByRef paper_viewport As ZcadPViewport, _
                                ByVal frame_width As Single, _
                                ByVal frame_height As Single)
    
    Dim frameCenter(0 To 2) As Double
    
    ' �p���g���S�Z�o
    frameCenter(0) = frame_width / 2
    frameCenter(1) = frame_height / 2
    frameCenter(2) = 0
    
    ' �ʒu�y�уT�C�Y����
    With paper_viewport
        .Center = frameCenter
        .Width = frame_width
        .Height = frame_height
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## ���C�A�E�g�̈���ݒ�
'------------------------------------------------------------------------------
Private Sub configurePrintSettings(ByRef new_layout As ZcadLayout, _
                                   ByVal frame_width As Single, _
                                   ByVal frame_height As Single, _
                                   ByVal scale_factor As Variant, _
                                   ByVal style_name As Variant, _
                                   ByVal printer_name As Variant, _
                                   ByVal a3_paper As Variant, _
                                   ByVal a4_paper As Variant, _
                                   ByVal offset_x As Variant, _
                                   ByVal offset_y As Variant)
    
    Dim plotOffset(0 To 1) As Double
    
    With new_layout
        
        ' ����̈�
        .PlotType = zcLayout
        
        ' �ړx
        .UseStandardScale = False
        .SetCustomScale 1, scale_factor
        
        ' ����X�^�C��
        .StyleSheet = style_name
        
        ' �v�����^����
        .ConfigName = printer_name
        
        ' ����I�t�Z�b�g(XY����ʂƋt�̂��ߒ���)
        plotOffset(0) = offset_y: plotOffset(1) = offset_x
        .PlotOrigin = plotOffset
        
        ' �p���T�C�Y����ѕ���
        frame_width = frame_width / scale_factor
        frame_height = frame_height / scale_factor
        If frame_width = 420 And frame_height = 297 Then
            .CanonicalMediaName = a3_paper
            .PlotRotation = zc90degrees
        ElseIf frame_width = 297 And frame_height = 420 Then
            .CanonicalMediaName = a3_paper
            .PlotRotation = zc0degrees
        ElseIf frame_width = 210 And frame_height = 297 Then
            .CanonicalMediaName = a4_paper
            .PlotRotation = zc0degrees
        Else
            MsgBox "�p���T�C�Y�����o�ł��Ȃ��������ߎ蓮�Őݒ肵�ĉ������B"
        End If
        
    End With
    
    ' ��ʏ�̔��f���s�����߃��C�A�E�g�^�u�̐؂�ւ�
    ' ZWCAD�̕s���Regen�ł͔��f����Ȃ����Ƃɑ΂���΍�
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts("Model")
    ThisDrawing.ActiveLayout = new_layout
    
End Sub
