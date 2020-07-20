Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------------------------------------------------
' ## �ݒ�l���\��̒萔�S
'------------------------------------------------------------------------------
Private Const PAPER_TAG As String = "CPN"       ' �p���g���肷�鑮���}�[�N
Private Const PRINTER_NAME As String = "DWG to PDF.pc5"   ' �v�����^����
Private Const SCALE_FACTOR As Single = 1        ' �}�ʂ̎ړx(1:n)
Private Const PAPER_NAME_A3 As String = "ISO_A3_(297.00_x_420.00_MM)"
Private Const PAPER_NAME_A4 As String = "ISO_A4_(297.00_x_210.00_MM)"
Private Const STYLE_NAME As String = "Monochrome.ctb"

'------------------------------------------------------------------------------
' ## �g�u���b�N�I���ɂ�郌�C�A�E�g����
'
' �y�[�W�ԍ��𑮐��Ƃ��Ď��p���g�u���b�N���烌�C�A�E�g�𐶐�����
'------------------------------------------------------------------------------
Public Sub CreateLayout()
    
    'On Error GoTo Error_Handler
    
    ' �p���g�u���b�N�̑I��
    Dim pickEntity As ZcadEntity
    Dim pickPoint As Variant
    Dim paperFrame As ZcadBlockReference    ' �X�R�[�v����������
    ThisDrawing.Utility.GetEntity pickEntity, pickPoint, _
        "�p���g��I�� [Cancel(ESC)]"
    If TypeOf pickEntity Is ZcadBlockReference Then
        Set paperFrame = pickEntity
    Else
        ThisDrawing.Utility.prompt "�u���b�N�ȊO���I������܂����B"
        Exit Sub
    End If
    
    ' �p���g�T�C�Y�擾(�X�R�[�v����������)
    Dim minExtent As Variant, maxExtent As Variant
    Dim frameWidth As Single, frameHight As Single
    paperFrame.GetBoundingBox minExtent, maxExtent
    frameWidth = maxExtent(0) - minExtent(0)
    frameHight = maxExtent(1) - minExtent(1)
    
    ' ���C�A�E�g���̎擾
    Dim frameAttributes As Variant
    Dim currentAttribute As ZcadAttributeReference
    Dim frameName As String
    Dim i As Long
    frameAttributes = paperFrame.GetAttributes
    For i = 0 To UBound(frameAttributes)
        Set currentAttribute = frameAttributes(i)
        If currentAttribute.TagString = PAPER_TAG Then
            frameName = currentAttribute.TextString
            Exit For
        End If
    Next i
    If frameName = "" Then
        ThisDrawing.Utility.prompt "���C�A�E�g���̂�������܂���B"
        Exit Sub
    End If
    
    ' �V�K���C�A�E�g�̍쐬����уA�N�e�B�u��
    Dim newLayout As zcadLayout
    Set newLayout = ThisDrawing.Layouts.Add(frameName)
    ThisDrawing.ActiveLayout = newLayout
    
    ' ����ݒ�
    newLayout.ConfigName = PRINTER_NAME
    
    Dim flg As Long: flg = 0
    If frameHight = 297 * SCALE_FACTOR Then
        If frameWidth = 420 * SCALE_FACTOR Then
            newLayout.CanonicalMediaName = PAPER_NAME_A3
            newLayout.PlotRotation = zc90degrees
            flg = 1
        ElseIf frameWidth = 210 * SCALE_FACTOR Then
            newLayout.CanonicalMediaName = PAPER_NAME_A4
            newLayout.PlotRotation = zc0degrees
            flg = 1
        End If
    End If
    If flg = 0 Then
        MsgBox "�p���T�C�Y�����o�ł��Ȃ��������ߎ蓮�ݒ���s���ĉ������B"
    End If
    
    ' ���C�A�E�g�r���[�|�[�g�쐬�ƃT�C�Y����
    Dim paperViewport As ZcadPViewport
    Set paperViewport = ThisDrawing.PaperSpace.Item(1)  ' (0)�̓��f����ԁH
    
    paperViewport.Width = frameWidth
    paperViewport.height = frameHight
    
    Dim frameCenter(0 To 2) As Double
    frameCenter(0) = frameWidth / 2
    frameCenter(1) = frameHight / 2
    frameCenter(2) = 0
    
    paperViewport.center = frameCenter
    
    ZoomExtents
    
    ' �r���[�|�[�g���Ɏw��p����\��
    ThisDrawing.MSpace = True
    
    ThisDrawing.ActivePViewport = paperViewport
    ZoomWindow minExtent, maxExtent
    
    ThisDrawing.MSpace = False
    
    ' ���̑��̈���ݒ�
    Dim plotOffset(0 To 1) As Double
    plotOffset(0) = -5: plotOffset(1) = -17
    
    With newLayout
        .PlotType = zcLayout
        .UseStandardScale = False
        .SetCustomScale 1, SCALE_FACTOR
        .StyleSheet = STYLE_NAME
        .PlotOrigin = plotOffset
    End With
    
    ' �\�����X�V���邽�߂Ƀ^�u�؂�ւ����s��
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts("Model")
    ThisDrawing.ActiveLayout = newLayout
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.prompt "�Ȃ�炩�̃G���[�ł��B"
    
End Sub
