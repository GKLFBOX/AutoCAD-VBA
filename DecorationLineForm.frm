VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DecorationLineForm 
   Caption         =   "�����������ݒ�"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   OleObjectBlob   =   "DecorationLineForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "DecorationLineForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' ## �t�H�[��������
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Dim i As Long
    Dim configData As Variant
    
    ' �Q�Ɛ���}�ݒ背�C���[���̌Ăяo��
    For i = 0 To ThisDrawing.Layers.Count - 1
        ReferenceLineLayerBox.AddItem ThisDrawing.Layers.Item(i).Name
    Next i
    
    ' �������}�ݒ背�C���[���̌Ăяo��
    For i = 0 To ThisDrawing.Layers.Count - 1
        StrikethroughLayerBox.AddItem ThisDrawing.Layers.Item(i).Name
    Next i
    
    ' �Q�Ɛ���}�ݒ�l�ǂݍ���
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.REFERENCELINE_CONFIG), vbCrLf)
    If UBound(configData) = 2 Then
        ReferenceLineLayerBox.Value = configData(0)
        ReferenceLineLengthBox.Value = configData(1)
        ReferenceLineOffsetBox.Value = configData(2)
    End If
    
    ' �������}�ݒ�l�ǂݍ���
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.STRIKETHROUGH_CONFIG), vbCrLf)
    If UBound(configData) = 2 Then
        StrikethroughLayerBox.Value = configData(0)
        StrikethroughRedBox.Value = _
            IIf(configData(1) = "True", "True", "False")
        TargetEntityLayerBox.Value = _
            IIf(configData(2) = "True", "True", "False")
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## �ݒ�l�ۑ�
'------------------------------------------------------------------------------
Private Sub DecorationSaveButton_Click()
    
    Dim configData As Variant
    
    ' ��w���X�g�擾
    Dim i As Long
    Dim layerList() As Variant
    For i = 0 To ThisDrawing.Layers.Count - 1
        ReDim Preserve layerList(i)
        layerList(i) = ThisDrawing.Layers.Item(i).Name
    Next i
    
    ' �Q�Ɛ���}��w���͊m�F
    If Not CommonFunction.IsMatchList _
        (layerList, ReferenceLineLayerBox.Value) Then
        MsgBox "�Q�Ɛ���}��w�̓��͂��s���ł��B", vbCritical
        Exit Sub
    End If
    
    ' �Q�Ɛ������W�����͊m�F
    If Not IsNumeric(ReferenceLineLengthBox.Value) Then
        MsgBox "�Q�Ɛ������W���̓��͂��s���ł��B", vbCritical
        Exit Sub
    End If
    
    ' �Q�Ɛ��I�t�Z�b�g�W�����͊m�F
    If Not IsNumeric(ReferenceLineOffsetBox.Value) Then
        MsgBox "�Q�Ɛ��I�t�Z�b�g�W���̓��͂��s���ł��B", vbCritical
        Exit Sub
    End If
    
    ' �������}��w���͊m�F
    If Not CommonFunction.IsMatchList _
        (layerList, StrikethroughLayerBox.Value) Then
        MsgBox "�������}��w�̓��͂��s���ł��B", vbCritical
        Exit Sub
    End If
    
    ' �ݒ�t�H���_�̏���
    Call CommitConfig.PrepareConfigFolder
    
    ' �Q�Ɛ���}�ݒ�l�ۑ�
    configData = ReferenceLineLayerBox.Value & vbCrLf _
               & ReferenceLineLengthBox.Value & vbCrLf _
               & ReferenceLineOffsetBox.Value
    
    Call CommitConfig.SaveConfig _
        (FormDisplay.REFERENCELINE_CONFIG, configData)
    
    ' �������}�ݒ�l�ۑ�
    configData = StrikethroughLayerBox.Value & vbCrLf _
               & StrikethroughRedBox.Value & vbCrLf _
               & TargetEntityLayerBox.Value
    
    Call CommitConfig.SaveConfig _
        (FormDisplay.STRIKETHROUGH_CONFIG, configData)
    
End Sub
