Attribute VB_Name = "OutputTextList"
Option Explicit

'------------------------------------------------------------------------------
' ## �����I�u�W�F�N�g��csv�o�̓v���O����   2020/07/26 G.O.
'
' �}�育�ƂɑI��͈͂̕����I�u�W�F�N�g��csv�`���̃��X�g�ŏo�͂���
' �o�͂�����͉�w,�F,�t�H���g,���e,����,���W
'------------------------------------------------------------------------------
Public Sub OutputTextList()
    
    On Error GoTo Error_Handler
    
    Dim outputFile As String
    Dim outputData As String
    Dim figureList() As String
    Dim figureText As String
    
    ' �}��̑I��
    Dim targetFigure As ZcadEntity
    Dim pickPoint As Variant
    ThisDrawing.Utility.GetEntity targetFigure, pickPoint, _
        "�}���I�� [Cancel(ESC)]"
    Call CommonSub.ResetHighlight(targetFigure)
    If Not CommonFunction.IsTextObject(targetFigure) Then
        ThisDrawing.Utility.Prompt "�������I������܂���ł����B" & vbCrLf
        Exit Sub
    End If
    
    figureText = targetFigure.TextString
    ReDim Preserve figureList(0)
    
    ' �o�̓f�[�^�w�b�_�����܂��͐}��d���������
    outputFile = CommonFunction.MakeFilePath("_�e�L�X�g�f�[�^", ".csv")
    If Dir(outputFile) = "" Then
        outputData = makeHeader()
    Else
        Call makeFigureList(outputFile, figureList())
    End If
    
    Call avoidDuplicateFigure(figureText, figureList())
    
    ' �o�͑Ώۂ�͈͑I��
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    ' �o�̓f�[�^�̍쐬����я����o��
    Call makeTextData(targetSelectionSet, figureText, outputData)
    If outputData = "" Then
        ThisDrawing.Utility.Prompt "�e�L�X�g���I������܂���ł����B" & vbCrLf
    Else
        Call CommonSub.OutputCSV(outputFile, outputData)
        ThisDrawing.Utility.Prompt "�e�L�X�g���o���������܂����B" & vbCrLf
    End If
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "�Ȃ�炩�̃G���[�ł��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## �w�b�_����
'------------------------------------------------------------------------------
Private Function makeHeader() As String
    
    makeHeader = """�}��""," _
               & """��w""," _
               & """�F""," _
               & """�X�^�C��""," _
               & """���e""," _
               & """��������""," _
               & """X���W""," _
               & """Y���W""," _
               & """Z���W""" & vbCrLf
    
End Function

'------------------------------------------------------------------------------
' ## �}�胊�X�g�̐���
'------------------------------------------------------------------------------
Private Sub makeFigureList(ByVal output_file As String, _
                           ByRef figure_list() As String)
    
    Dim i As Long
    Dim bufferData As String
    
    Open output_file For Input As #1
        i = 0
        Line Input #1, bufferData
        Do Until EOF(1)
            Line Input #1, bufferData
            ReDim Preserve figure_list(0 To i)
            figure_list(i) = Left(bufferData, InStr(bufferData, """,") - 1)
            figure_list(i) = Right(figure_list(i), Len(figure_list(i)) - 1)
            i = i + 1
        Loop
    Close #1
    
End Sub

'------------------------------------------------------------------------------
' ## �}��d���������
'------------------------------------------------------------------------------
Private Sub avoidDuplicateFigure(ByRef figure_text As String, _
                                 ByRef figure_list() As String)
    
    Dim i As Long
    Dim buffer_text As String
    
    i = 1
    buffer_text = figure_text
    Do While CommonFunction.IsMatchList(figure_list, figure_text)
        i = i + 1
        figure_text = buffer_text & " (" & i & ")"
    Loop
    
    ' �}��m�F�v�����v�g
    ThisDrawing.Utility.Prompt _
        "�}��́u" & figure_text & "�v�ł��B" & vbCrLf
    ThisDrawing.Utility.Prompt _
        "��肪������Ώo�͔͈͂�I�����Ă��������B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## csv�`���̃f�[�^�쐬
'------------------------------------------------------------------------------
Private Sub makeTextData(ByVal target_selectionset As ZcadSelectionSet, _
                         ByVal figure_text As String, _
                         ByRef output_data As String)
    
    Dim bufferData As String
    Dim extractEntity As ZcadEntity
    Dim extractLayer As String
    Dim extractColor As Long
    Dim extractStyle As String
    Dim extractText As String
    Dim extractHeight As Double
    Dim extractCoordinate As Variant
    
    ' �f�[�^�쐬�O�̒l��ۑ�
    bufferData = output_data
    
    ' �����񉻂����csv�`���f�[�^�쐬
    figure_text = CommonFunction.FormatString(figure_text)
    For Each extractEntity In target_selectionset
        
        If Not CommonFunction.IsTextObject(extractEntity) Then _
            GoTo Continue_extractEntity
            
        With extractEntity
            extractLayer = CommonFunction.FormatString(.Layer)
            extractColor = .TrueColor.ColorIndex
            extractStyle = CommonFunction.FormatString(.StyleName)
            extractText = CommonFunction.FormatString(.TextString)
            extractHeight = .Height
            extractCoordinate = .insertionPoint
        End With
        
        output_data = output_data _
                    & figure_text & "," _
                    & extractLayer & "," _
                    & extractColor & "," _
                    & extractStyle & "," _
                    & extractText & "," _
                    & extractHeight & "," _
                    & extractCoordinate(0) & "," _
                    & extractCoordinate(1) & "," _
                    & extractCoordinate(2) & vbCrLf
        
Continue_extractEntity:
        
    Next extractEntity
    
    ' �l�̍폜�܂��͍ŏI�s�̉��s�폜
    If bufferData = output_data Then
        output_data = ""
    Else
        output_data = Left(output_data, Len(output_data) - 2)
    End If
    
End Sub
