VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DecorationLineForm 
   Caption         =   "�����������ݒ�"
   ClientHeight    =   5790
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
    
    ' �I�v�V�����{�^��������
    ReferenceLineLayerOffButton.Value = True
    StrikethroughLayerOffButton.Value = True
    
    ' �Q�Ɛ���}�ݒ背�C���[���̌Ăяo��
    For i = 0 To ThisDrawing.Layers.Count - 1
        ReferenceLineLayerBox.AddItem ThisDrawing.Layers.Item(i).Name
    Next i
    
    ' ����������}�ݒ背�C���[���̌Ăяo��
    For i = 0 To ThisDrawing.Layers.Count - 1
        StrikethroughLayerBox.AddItem ThisDrawing.Layers.Item(i).Name
    Next i
    
    ' �Q�Ɛ���}�ݒ�l�ǂݍ���
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.REFERENCELINE_CONFIG), vbCrLf)
    If UBound(configData) = 3 Then
        ReferenceLineLayerOnButton.Value = _
            IIf(configData(0) = "True", "True", "False")
        ReferenceLineLayerBox.Value = configData(1)
        ReferenceLineLengthBox.Value = configData(2)
        ReferenceLineOffsetBox.Value = configData(3)
    End If
    
    ' ����������}�ݒ�l�ǂݍ���
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.STRIKETHROUGH_CONFIG), vbCrLf)
    If UBound(configData) = 4 Then
        StrikethroughLayerOnButton.Value = _
            IIf(configData(0) = "True", "True", "False")
        StrikethroughLayerBox.Value = configData(1)
        StrikethroughLengthBox.Value = configData(2)
        StrikethroughRedBox.Value = _
            IIf(configData(3) = "True", "True", "False")
        TargetEntityLayerBox.Value = _
            IIf(configData(4) = "True", "True", "False")
    End If
    
    ' �Q�Ɛ���}��w�̎w��؂�ւ�
    Call ReferenceLineLayerOnButton_Change
    
    ' ����������}��w�̎w��؂�ւ�
    Call StrikethroughLayerOnButton_Change
    
End Sub

'------------------------------------------------------------------------------
' ## �Q�Ɛ���}��w�̎w��؂�ւ�
'------------------------------------------------------------------------------
Private Sub ReferenceLineLayerOnButton_Change()
    
    If ReferenceLineLayerOnButton.Value Then
        ReferenceLineLayerBox.Enabled = True
    Else
        ReferenceLineLayerBox.Enabled = False
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## ����������}��w�̎w��؂�ւ�
'------------------------------------------------------------------------------
Private Sub StrikethroughLayerOnButton_Change()
    
    If StrikethroughLayerOnButton.Value Then
        TargetEntityLayerBox.Enabled = True
        StrikethroughLayerBox.Enabled = True
    Else
        TargetEntityLayerBox.Value = False
        TargetEntityLayerBox.Enabled = False
        StrikethroughLayerBox.Enabled = False
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## �ݒ�l�ۑ�
'------------------------------------------------------------------------------
Private Sub DecorationSaveButton_Click()
    
    Dim configData As Variant
    
    ' �ݒ�l���͂̊m�F
    If Not validateDecorationConfig() Then Exit Sub
    
    ' �ݒ�t�H���_�̏���
    Call CommitConfig.PrepareConfigFolder
    
    ' �Q�Ɛ���}�ݒ�l�ۑ�
    configData = ReferenceLineLayerOnButton.Value & vbCrLf _
               & ReferenceLineLayerBox.Value & vbCrLf _
               & ReferenceLineLengthBox.Value & vbCrLf _
               & ReferenceLineOffsetBox.Value
    
    Call CommitConfig.SaveConfig _
        (FormDisplay.REFERENCELINE_CONFIG, configData)
    
    ' �������}�ݒ�l�ۑ�
    configData = StrikethroughLayerOnButton.Value & vbCrLf _
               & StrikethroughLayerBox.Value & vbCrLf _
               & StrikethroughLengthBox.Value & vbCrLf _
               & StrikethroughRedBox.Value & vbCrLf _
               & TargetEntityLayerBox.Value
    
    Call CommitConfig.SaveConfig _
        (FormDisplay.STRIKETHROUGH_CONFIG, configData)
    
End Sub

'------------------------------------------------------------------------------
' ## �ݒ�l���͂̊m�F
'------------------------------------------------------------------------------
Private Function validateDecorationConfig() As Boolean
    
    validateDecorationConfig = False
    
    ' ��w���X�g�擾
    Dim i As Long
    Dim layerList() As Variant
    For i = 0 To ThisDrawing.Layers.Count - 1
        ReDim Preserve layerList(i)
        layerList(i) = ThisDrawing.Layers.Item(i).Name
    Next i
    
    ' �Q�Ɛ���}��w���͊m�F
    If ReferenceLineLayerOnButton.Value _
    And Not CommonFunction.IsMatchList _
        (layerList, ReferenceLineLayerBox.Value) Then
        MsgBox "�Q�Ɛ���}��w�̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    ' �Q�Ɛ������W�����͊m�F
    If Not IsNumeric(ReferenceLineLengthBox.Value) Then
        MsgBox "�Q�Ɛ������W���̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    ' �Q�Ɛ��I�t�Z�b�g�W�����͊m�F
    If Not IsNumeric(ReferenceLineOffsetBox.Value) Then
        MsgBox "�Q�Ɛ��I�t�Z�b�g�W���̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    ' �������}��w���͊m�F
    If StrikethroughLayerOnButton.Value _
    And Not CommonFunction.IsMatchList _
        (layerList, StrikethroughLayerBox.Value) Then
        MsgBox "����������}��w�̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    ' �Q�Ɛ������W�����͊m�F
    If Not IsNumeric(StrikethroughLengthBox.Value) Then
        MsgBox "�������������W���̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    validateDecorationConfig = True
    
End Function
