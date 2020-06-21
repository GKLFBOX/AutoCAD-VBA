Attribute VB_Name = "OutputText"
'------------------------------------------------------------------------------
' ## �R�[�f�B���O�K�C�h���C��
'
' [You.Activate|VBA�R�[�f�B���O�K�C�h���C��]�ɏ�������
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## �����I�u�W�F�N�g��csv�o�̓v���O����
'
' �����I�u�W�F�N�g�̓��e�Ƒ�����csv�`���ŏo�͂���
'------------------------------------------------------------------------------
Public Sub OutputText()
    
    On Error GoTo Error_Handler
    
    ' �}��̑I��
    Dim pickPoint As Variant
    Dim targetFigure As ZcadEntity
    ThisDrawing.Utility.GetEntity targetFigure, pickPoint, _
        "�}���I�� [Cancel(ESC)]"
    
    Call CommonSub.ResetHighlight(targetFigure)
    
    If Not CommonFunction.IsTextObject(targetFigure) Then
        ThisDrawing.Utility.Prompt "�G���[�F������I�����Ă��������B" & vbCrLf
        Exit Sub
    End If
    
    ThisDrawing.Utility.Prompt _
        "�}��́u" & targetFigure.TextString & "�v�ł��B" & vbCrLf
    ThisDrawing.Utility.Prompt _
        "��肪������Ώo�͔͈͂�I�����Ă��������B" & vbCrLf
    
    ' �o�͑Ώۂ�͈͑I��
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    ' �o�̓f�[�^�̏���
    Dim outputFile As String
    outputFile = CommonFunction.MakeFilePath("_�e�L�X�g�f�[�^", ".csv")
    
    Dim outputData As String
    If Dir(outputFile) = "" Then
        outputData = _
            "�}��,��w,�F,�X�^�C��,���e,��������,X���W,Y���W,Z���W" & vbCrLf
    End If
    
    ' �o�̓f�[�^�̍쐬
    If Not targetSelectionSet.Count = 0 Then
        Call makeTextData(targetSelectionSet, targetFigure, outputData)
    End If
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    ' �o�̓f�[�^�̏����o��
    Call outputCSV(outputFile, outputData)
    ThisDrawing.Utility.Prompt "�e�L�X�g���o���������܂����B" & vbCrLf
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "�G���[�F�R�}���h���I�����܂��B" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## csv�t�@�C���ւ̏o��
'------------------------------------------------------------------------------
Private Sub outputCSV(ByVal output_file As String, ByVal output_data As String)
    
    Open output_file For Append As #1
    Print #1, output_data
    Close #1
    
End Sub

'------------------------------------------------------------------------------
' ## csv�`���̃f�[�^�쐬
'------------------------------------------------------------------------------
Private Sub makeTextData(ByVal target_selectionset As ZcadSelectionSet, _
                         ByVal target_figure As ZcadEntity, _
                         ByRef output_data As String)
    
    ' �}���csv�p�����񉻐��`����
    Dim figureText As String
    figureText = formatString(target_figure.TextString)
    
    ' �����񉻏�����csv�`���f�[�^�쐬
    Dim extractObject As ZcadEntity
    Dim exLayer As String
    Dim exColor As Long
    Dim exStyle As String
    Dim exText As String
    Dim exHeight As Double
    Dim exCoordinate As Variant
    
    For Each extractObject In target_selectionset
        If CommonFunction.IsTextObject(extractObject) Then
            
            With extractObject
                exLayer = formatString(.Layer)
                exColor = .TrueColor.ColorIndex
                exStyle = formatString(.StyleName)
                exText = formatString(.TextString)
                exHeight = .Height
                exCoordinate = .InsertionPoint
            End With
            
            output_data = output_data & _
                figureText & "," & _
                exLayer & "," & _
                exColor & "," & _
                exStyle & "," & _
                exText & "," & _
                exHeight & "," & _
                exCoordinate(0) & "," & _
                exCoordinate(1) & "," & _
                exCoordinate(2) & vbCrLf
            
        End If
    Next extractObject
    
    output_data = Left(output_data, Len(output_data) - 2) ' �ŏI�s�̉��s�폜
    
End Sub

'------------------------------------------------------------------------------
' ## csv�p�̕����񐮌`(�_�u���N�H�[�e�[�V�����̕t���ƕ�����)
'------------------------------------------------------------------------------
Private Function formatString(ByVal target_text As String) As String
    
    formatString = """" & Replace(target_text, """", """""") & """"
    
End Function
