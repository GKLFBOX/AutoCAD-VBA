Attribute VB_Name = "CreatLayout"
Option Explicit

'------------------------------------------------------------------------------
' ## �ݒ�l���\��̒萔�S
'------------------------------------------------------------------------------
Private Const PAPER_TAG As String = "CPN"
Private Const PRINTER_NAME As String = "DWG to PDF.pc5"
Private Const SCALE_FACTOR As Single = 100
Private Const PAPER_NAME_A3 As String = "ISO_A3_(297.00_x_420.00_MM)"
Private Const PAPER_NAME_A4 As String = "ISO_A4_(210.00_x_297.00_MM)"
Private Const STYLE_NAME As String = "Monochrome.ctb"
Private Const OFFSET_X As Double = "-17"
Private Const OFFSET_Y As Double = "-5"

'------------------------------------------------------------------------------
' ## �g�u���b�N�I���ɂ�郌�C�A�E�g����
'
' �w��̃}�[�N�𑮐��Ƃ��Ď��p���g�u���b�N���烌�C�A�E�g�𐶐�����
'------------------------------------------------------------------------------
Public Sub CreateLayout()
    
    'On Error GoTo Error_Handler
    
    Dim paperFrame As ZcadBlockReference
    Dim minFramePoint As Variant, maxFramePoint As Variant
    Dim frameWidth As Single, frameHeight As Single
    Dim frameName As String
    Dim paperViewport As ZcadPViewport
    
    Dim newLayout As zcadLayout
    
    ' �p���g�u���b�N�̑I��
    Call pickPaperFrame(paperFrame)
    If paperFrame Is Nothing Then Exit Sub
    
    ' �p���g�T�C�Y�擾
    paperFrame.GetBoundingBox minFramePoint, maxFramePoint
    frameWidth = maxFramePoint(0) - minFramePoint(0)
    frameHeight = maxFramePoint(1) - minFramePoint(1)
    
    ' �p���g���̎擾
    Call fetchFrameName(paperFrame, frameName)
    If frameName = "" Then Exit Sub
    
    ' �V�K���C�A�E�g�̍쐬����уA�N�e�B�u��
    Set newLayout = ThisDrawing.Layouts.Add(frameName)
    ThisDrawing.ActiveLayout = newLayout
    
    ' �V�KPaperSpace��Item(0)�̓��C�A�E�g�̉�ʂ��̂��̂ł���
    ' ���[�U�[���F�����Ă���r���[�|�[�g�I�u�W�F�N�g��Item(1)�̂��ߒ���
    Set paperViewport = ThisDrawing.PaperSpace.Item(1)
    
    ' �r���[�|�[�g�̈ʒu����уT�C�Y����
    Call adjustViewportShape(paperViewport, frameWidth, frameHeight)
    
    ' �\�������O�ɑS�̕\�������Ȃ��ƂȂ����r���[�|�[�g���S�̕\�����o�O��
    ' �ړx���኱�Y���Ă��܂����߂����ōs���Ă���(�����s��)
    ZoomExtents
    
    ' �r���[�|�[�g�̕\������
    Call adjustViewportDisplay(paperViewport, minFramePoint, maxFramePoint)
    
    ' ����ݒ�
    Call configurePrintSettings(newLayout, frameWidth, frameHeight)
    
    ' ��ʏ�̔��f���s�����߃��C�A�E�g�^�u�̐؂�ւ�
    ' ZWCAD�̕s���Regen�ł͔��f����Ȃ����Ƃɑ΂���΍�
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts("Model")
    ThisDrawing.ActiveLayout = newLayout
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.prompt "�Ȃ�炩�̃G���[�ł��B"
    
End Sub

'------------------------------------------------------------------------------
' ## �p���g�u���b�N�̑I��
'------------------------------------------------------------------------------
Private Sub pickPaperFrame(ByRef paper_frame As ZcadBlockReference)
    
    Dim pickEntity As ZcadEntity
    Dim pickPoint As Variant
    
    ThisDrawing.Utility.GetEntity pickEntity, pickPoint, _
        "�p���g��I�� [Cancel(ESC)]"
        
    If TypeOf pickEntity Is ZcadBlockReference Then
        Set paper_frame = pickEntity
    Else
        ThisDrawing.Utility.prompt "�u���b�N�ȊO���I������܂����B"
        Exit Sub
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## �p���g���̂��擾
'------------------------------------------------------------------------------
Private Sub fetchFrameName(ByVal paper_frame As ZcadBlockReference, _
                           ByRef frame_name As String)
    
    Dim frameAttributes As Variant
    Dim currentAttribute As ZcadAttributeReference
    
    frameAttributes = paper_frame.GetAttributes
    
    Dim i As Long
    For i = 0 To UBound(frameAttributes)
        Set currentAttribute = frameAttributes(i)
        If currentAttribute.TagString = PAPER_TAG Then
            frame_name = currentAttribute.TextString
            Exit For
        End If
    Next i
    
    If frame_name = "" Then
        ThisDrawing.Utility.prompt "�p���g���̂�������܂���B"
        Exit Sub
    End If
    
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
' ## �r���[�|�[�g�̕\������
'------------------------------------------------------------------------------
Private Sub adjustViewportDisplay(ByRef paper_viewport As ZcadPViewport, _
                                  ByVal min_framepoint As Variant, _
                                  ByVal max_framepoint As Variant)
    
    ThisDrawing.MSpace = True
    
    ThisDrawing.ActivePViewport = paper_viewport
    ZoomWindow min_framepoint, max_framepoint
    
    ThisDrawing.MSpace = False
    
End Sub

'------------------------------------------------------------------------------
' ## ���C�A�E�g�̈���ݒ�
'------------------------------------------------------------------------------
Private Sub configurePrintSettings(ByRef new_layout As zcadLayout, _
                                   ByVal frame_width As Single, _
                                   ByVal frame_height As Single)
    
    Dim plotOffset(0 To 1) As Double
    
    ' �I�t�Z�b�g��(XY����ʂƋt�̂��ߒ���)
    plotOffset(0) = OFFSET_Y: plotOffset(1) = OFFSET_X
    
    ' �ړx����
    frame_width = frame_width / SCALE_FACTOR
    frame_height = frame_height / SCALE_FACTOR
    
    With new_layout
        
        ' �v�����^����
        .ConfigName = PRINTER_NAME
        
        ' �p���T�C�Y����ѕ���
        If frame_width = 420 And frame_height = 297 Then
            .CanonicalMediaName = PAPER_NAME_A3
            .PlotRotation = zc90degrees
        ElseIf frame_width = 297 And frame_height = 420 Then
            .CanonicalMediaName = PAPER_NAME_A3
            .PlotRotation = zc0degrees
        ElseIf frame_width = 210 And frame_height = 297 Then
            .CanonicalMediaName = PAPER_NAME_A4
            .PlotRotation = zc0degrees
        Else
            MsgBox "�p���T�C�Y�����o�ł��Ȃ��������ߎ蓮�ݒ���s���ĉ������B"
            Exit Sub
        End If
        
        ' ����̈�
        .PlotType = zcLayout
        
        ' ����I�t�Z�b�g
        .PlotOrigin = plotOffset
        
        ' �ړx(����Ȃ��Ƃ����ǉ�����)
        .UseStandardScale = False
        .SetCustomScale 1, SCALE_FACTOR
        
        ' ����X�^�C��
        .StyleSheet = STYLE_NAME
        
    End With
    
End Sub
