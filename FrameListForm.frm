VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrameListForm 
   Caption         =   "�p���g���X�gcsv�o��"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "FrameListForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FrameListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' ## �ݒ�t�@�C���̃t�@�C����
'------------------------------------------------------------------------------
Private Const FRAMELIST_CONFIG As String = "\FrameList.config"

'------------------------------------------------------------------------------
' ## �t�H�[��������
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    ' �u���b�N���̌Ăяo��
    Dim i As Long
    Dim buf As String
    For i = 0 To ThisDrawing.Blocks.Count - 1
        buf = ThisDrawing.Blocks.Item(i).Name
        If Left(buf, 1) <> "*" Then FrameBlockNameBox.AddItem buf
    Next i
    
    ' �ݒ�l�ǂݍ���
    Dim configData As Variant
    configData = Split(CommitConfig.LoadConfig(FRAMELIST_CONFIG), vbCrLf)
    If Not UBound(configData) = 1 Then Exit Sub
    
    FrameBlockNameBox.Value = configData(0)
    FrameTagBox.Value = configData(1)
    
End Sub

'------------------------------------------------------------------------------
' ## �p���g���X�g�o�̓{�^��
'------------------------------------------------------------------------------
Private Sub OutputFrameListButton_Click()
    
    Dim configData As Variant
    
    ' �ݒ�l�̓��͊m�F
    If Not validateFrameListConfig() Then Exit Sub
    
    FrameListForm.Hide
    
    ' �p���g���X�g�o�͎��s
    Call OutputFrameList.OutputFrameList(FrameBlockNameBox.Value, _
                                         FrameTagBox.Value)
    
    ' �ݒ�l�ۑ�����
    configData = FrameBlockNameBox.Value & vbCrLf _
               & FrameTagBox.Value
    
    ' �ݒ�t�H���_�̏���
    Call CommitConfig.PrepareConfigFolder
    
    ' �ݒ�l�ۑ�
    Call CommitConfig.SaveConfig(FRAMELIST_CONFIG, configData)
    
    Unload FrameListForm
    
End Sub

'------------------------------------------------------------------------------
' ## �ݒ�l���͂̊m�F
'------------------------------------------------------------------------------
Private Function validateFrameListConfig() As Boolean
    
    validateFrameListConfig = False
    
    ' �p���g�u���b�N�����͊m�F
    Dim i As Long
    Dim blockList() As Variant
    For i = 0 To ThisDrawing.Blocks.Count - 1
        ReDim Preserve blockList(i)
        blockList(i) = ThisDrawing.Blocks.Item(i).Name
    Next i
    If Not CommonFunction.IsMatchList(blockList, FrameBlockNameBox.Value) Then
        MsgBox "�p���g�u���b�N���̓��͂��s���ł��B", vbCritical
        Exit Function
    End If
    
    validateFrameListConfig = True
    
End Function
