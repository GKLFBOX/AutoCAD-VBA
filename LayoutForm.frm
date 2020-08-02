VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LayoutForm 
   Caption         =   "���C�A�E�g�ҏW"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   OleObjectBlob   =   "LayoutForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "LayoutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' ## �ݒ�t�@�C���̃t�@�C����
'------------------------------------------------------------------------------
Private Const LAYOUT_CONFIG As String = "\LayoutSetting.config"

'------------------------------------------------------------------------------
' ## �t�H�[��������
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Dim i As Long
    Dim tempPlotConfig As ZcadPlotConfiguration
    
    ' �Ăяo���p�ꎞ�y�[�W�ݒ�
    Set tempPlotConfig = ThisDrawing.PlotConfigurations.Add("TempPlotConfig")
    tempPlotConfig.RefreshPlotDeviceInfo
    
    ' ���C���[���̌Ăяo��
    For i = 0 To ThisDrawing.Layers.Count - 1
        LayoutLayerBox.AddItem ThisDrawing.Layers.Item(i).Name
    Next i
    
    ' ����X�^�C�����̌Ăяo��
    Dim styleList As Variant
    styleList = tempPlotConfig.GetPlotStyleTableNames
    For i = 0 To UBound(styleList)
        StyleNameBox.AddItem styleList(i)
    Next i
    
    ' �v�����^���̌Ăяo��
    Dim printerList As Variant
    printerList = tempPlotConfig.GetPlotDeviceNames
    For i = 0 To UBound(printerList)
        PrinterNameBox.AddItem printerList(i)
    Next i
    
    ' �ݒ�l�ǂݍ���
    Dim configData As Variant
    configData = Split(CommitConfig.LoadConfig(LAYOUT_CONFIG), vbCrLf)
    If Not UBound(configData) = 9 Then Exit Sub
    
    FrameTagBox.Value = configData(0)
    ScaleFactorBox.Value = configData(1)
    LayoutLayerBox.Value = configData(2)
    StyleNameBox.Value = configData(3)
    PrinterNameBox.Value = configData(4)
    A3PaperBox.Value = configData(5)
    A4PaperBox.Value = configData(6)
    OffsetXBox.Value = configData(7)
    OffsetYBox.Value = configData(8)
    CustomScaleBox.Value = configData(9)
    
    ' �p���ݒ�Ăяo��
    Dim paperList As Variant
    tempPlotConfig.ConfigName = PrinterNameBox.Value
    paperList = tempPlotConfig.GetCanonicalMediaNames
    For i = 0 To UBound(paperList)
        If paperList(i) Like "*A3*" Then A3PaperBox.AddItem paperList(i)
        If paperList(i) Like "*A4*" Then A4PaperBox.AddItem paperList(i)
    Next i
    
End Sub

'------------------------------------------------------------------------------
' ## �v�����^���̍X�V��
'------------------------------------------------------------------------------
Private Sub PrinterNameBox_Change()
    
    Dim tempPlotConfig As ZcadPlotConfiguration
    
    ' �Ăяo���p�ꎞ�y�[�W�ݒ�
    Set tempPlotConfig = ThisDrawing.PlotConfigurations.Item("TempPlotConfig")
    
    ' �p���ݒ胊�Z�b�g
    A3PaperBox.Value = ""
    A4PaperBox.Value = ""
    A3PaperBox.Clear
    A4PaperBox.Clear
    
    ' �v�����^���̂̑��݊m�F
    Dim printerList As Variant
    printerList = tempPlotConfig.GetPlotDeviceNames
    If Not CommonFunction.IsMatchList _
        (printerList, PrinterNameBox.Value) Then Exit Sub
    
    ' �p���ݒ�Ăяo������ѕ⊮
    Dim i As Long
    Dim paperList As Variant
    tempPlotConfig.ConfigName = PrinterNameBox.Value
    paperList = tempPlotConfig.GetCanonicalMediaNames
    For i = 0 To UBound(paperList)
        If paperList(i) Like "A3" Then A3PaperBox.Value = paperList(i)
        If paperList(i) Like "A4" Then A4PaperBox.Value = paperList(i)
        If paperList(i) Like "*A3*" Then A3PaperBox.AddItem paperList(i)
        If paperList(i) Like "*A4*" Then A4PaperBox.AddItem paperList(i)
    Next i
    
End Sub

'------------------------------------------------------------------------------
' ## �I�t�Z�b�g�ʓ��̓{�^��
'------------------------------------------------------------------------------
Private Sub InputOffsetButton_Click()
    
    Dim tempPlotConfig As ZcadPlotConfiguration
    
    ' �Ăяo���p�ꎞ�y�[�W�ݒ�
    Set tempPlotConfig = ThisDrawing.PlotConfigurations.Item("TempPlotConfig")
    
    ' A3�p���ݒ�̑��݊m�F
    Dim paperList As Variant
    paperList = tempPlotConfig.GetCanonicalMediaNames
    If Not CommonFunction.IsMatchList _
        (paperList, A3PaperBox.Value) Then Exit Sub
    
    ' �I�t�Z�b�g�ʓ���(XY����ʂƋt�̂��ߒ���)
    ' �������l���擾�ł��Ȃ��p���ݒ肪���݂��邽�ߒ���
    Dim offsetLowerLeft As Variant, offsetUpperRight As Variant
    tempPlotConfig.GetPaperMargins offsetLowerLeft, offsetUpperRight
    OffsetXBox.Value = CSng((offsetLowerLeft(1) + offsetUpperRight(1)) / -2)
    OffsetYBox.Value = CSng((offsetLowerLeft(0) + offsetUpperRight(0)) / -2)
    
End Sub

'------------------------------------------------------------------------------
' ## �V�K���C�A�E�g�쐬�{�^��
'------------------------------------------------------------------------------
Private Sub CreateLayoutButton_Click()
    
    Dim configData As Variant
    
    ' �ݒ�l�̓��͊m�F
    If Not validateConfiguration() Then Exit Sub
    
    LayoutForm.Hide
    
    ' �V�K���C�A�E�g�쐬���s
    Call CreateLayout.CreateLayout(FrameTagBox.Value, _
                                   ScaleFactorBox.Value, _
                                   LayoutLayerBox.Value, _
                                   StyleNameBox.Value, _
                                   PrinterNameBox.Value, _
                                   A3PaperBox.Value, _
                                   A4PaperBox.Value, _
                                   OffsetXBox.Value, _
                                   OffsetYBox.Value)
    
    ' �ݒ�l�ۑ�����
    configData = FrameTagBox.Value & vbCrLf _
               & ScaleFactorBox.Value & vbCrLf _
               & LayoutLayerBox.Value & vbCrLf _
               & StyleNameBox.Value & vbCrLf _
               & PrinterNameBox.Value & vbCrLf _
               & A3PaperBox.Value & vbCrLf _
               & A4PaperBox.Value & vbCrLf _
               & OffsetXBox.Value & vbCrLf _
               & OffsetYBox.Value & vbCrLf _
               & CustomScaleBox.Value
    
    ' �ݒ�t�H���_�̏���
    Call CommitConfig.PrepareConfigFolder
    
    ' �ݒ�l�ۑ�
    Call CommitConfig.SaveConfig(LAYOUT_CONFIG, configData)
    
    LayoutForm.Show
    
End Sub

'------------------------------------------------------------------------------
' ## �ݒ�l�̓��͊m�F
'------------------------------------------------------------------------------
Private Function validateConfiguration() As Boolean
    
    validateConfiguration = False
    
    Dim tempPlotConfig As ZcadPlotConfiguration
    
    ' �Ăяo���p�ꎞ�y�[�W�ݒ�
    Set tempPlotConfig = ThisDrawing.PlotConfigurations.Item("TempPlotConfig")
    
    ' �ړx���͊m�F
    If Not IsNumeric(ScaleFactorBox.Value) Or ScaleFactorBox.Value <= 0 Then
        MsgBox "�ړx�̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    ' ���C�A�E�g��w���͊m�F
    Dim i As Long
    Dim layerList() As Variant
    For i = 0 To ThisDrawing.Layers.Count - 1
        ReDim Preserve layerList(i)
        layerList(i) = ThisDrawing.Layers.Item(i).Name
    Next i
    If Not CommonFunction.IsMatchList(layerList, LayoutLayerBox.Value) Then
        MsgBox "���C�A�E�g��w�̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    ' ����X�^�C���̑��݊m�F
    Dim styleList As Variant
    styleList = tempPlotConfig.GetPlotStyleTableNames
    If Not CommonFunction.IsMatchList(styleList, StyleNameBox.Value) Then
        MsgBox "����X�^�C���̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    ' �v�����^���̂̑��݊m�F
    Dim printerList As Variant
    printerList = tempPlotConfig.GetPlotDeviceNames
    If Not CommonFunction.IsMatchList(printerList, PrinterNameBox.Value) Then
        MsgBox "�v�����^���̂̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    ' A3�p���ݒ肨���A4�p���ݒ�̑��݊m�F
    Dim paperList As Variant
    paperList = tempPlotConfig.GetCanonicalMediaNames
    If Not CommonFunction.IsMatchList(paperList, A3PaperBox.Value) _
    Or Not CommonFunction.IsMatchList(paperList, A4PaperBox.Value) Then
        MsgBox "�p���ݒ�̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    ' �I�t�Z�b�g�ʓ��͊m�F
    If Not IsNumeric(OffsetXBox.Value) _
    Or Not IsNumeric(OffsetYBox.Value) Then
        MsgBox "�I�t�Z�b�g�ʂ̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    ' �k�ړ��͊m�F
    If Not IsNumeric(CustomScaleBox.Value) Or CustomScaleBox.Value <= 0 Then
        MsgBox "�k�ڂ̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    validateConfiguration = True
    
End Function

'------------------------------------------------------------------------------
' ## �r���[�|�[�g�ǉ��{�^��
'------------------------------------------------------------------------------
Private Sub AddViewportButton_Click()
    
    Dim configData As Variant
    
    ' �ݒ�l�̓��͊m�F
    If Not validateConfiguration() Then Exit Sub
    
    LayoutForm.Hide
    
    ' �r���[�|�[�g�ǉ����s
    Call AddViewport.AddViewport(FrameTagBox.Value, _
                                 ScaleFactorBox.Value, _
                                 LayoutLayerBox.Value, _
                                 CustomScaleBox.Value)
    
    ' �ݒ�l�ۑ�����
    configData = FrameTagBox.Value & vbCrLf _
               & ScaleFactorBox.Value & vbCrLf _
               & LayoutLayerBox.Value & vbCrLf _
               & StyleNameBox.Value & vbCrLf _
               & PrinterNameBox.Value & vbCrLf _
               & A3PaperBox.Value & vbCrLf _
               & A4PaperBox.Value & vbCrLf _
               & OffsetXBox.Value & vbCrLf _
               & OffsetYBox.Value & vbCrLf _
               & CustomScaleBox.Value
    
    ' �ݒ�t�H���_�̏���
    Call CommitConfig.PrepareConfigFolder
    
    ' �ݒ�l�ۑ�
    Call CommitConfig.SaveConfig(LAYOUT_CONFIG, configData)
    
    LayoutForm.Show
    
End Sub

'------------------------------------------------------------------------------
' ## �t�H�[���I�����Ɉꎞ�y�[�W�ݒ���폜
'------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    
    Dim tempPlotConfig As ZcadPlotConfiguration
    Set tempPlotConfig = ThisDrawing.PlotConfigurations.Item("TempPlotConfig")
    tempPlotConfig.Delete
    
End Sub
