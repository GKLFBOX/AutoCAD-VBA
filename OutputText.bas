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
Sub OutputText()
    
    'On Error GoTo Error_Handler
    
    ' �}��̑I��
    Dim pickPoint As Variant
    Dim targetFigure As ZcadEntity
    ThisDrawing.Utility.GetEntity targetFigure, pickPoint, _
        "�}��(�����I�u�W�F�N�g)��I�� [Cancel(ESC)]"
    
    targetFigure.Highlight True
    
    ' csv�����ɐ}��̕����񉻏���
    Dim figureText As String
    figureText = _
    """" & Replace(targetFigure.TextString, """", """""") & """"
    
    If Not GeneralRoutine.isTextObject(targetFigure) Then
        targetFigure.Highlight False
        ThisDrawing.Utility.Prompt "������I�����Ă��������B" & vbCrLf
        Exit Sub
    End If
    
    targetFigure.Highlight False
    
    ' �o�͑Ώۂ�͈͑I��
    Dim targetSelectionSet As ZcadSelectionSet
    
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    ' �o�̓f�[�^�̍쐬
    If Not targetSelectionSet.Count = 0 Then
        
        Dim extractObject As ZcadEntity  ' �ϐ������v�����Ȃ����ߗv����
        Dim exLayer As String
        Dim exColor As Long
        Dim exStyle As String
        Dim exText As String
        Dim exHeight As Double
        Dim exCoordinate As Variant
        
        Dim outputData As String
        
        outputData = _
        "�}��,��w,�F,�X�^�C��,���e,��������,X���W,Y���W,Z���W" & vbCrLf
        
        For Each extractObject In targetSelectionSet
            If GeneralRoutine.isTextObject(extractObject) Then
                
                exLayer = """" & Replace(extractObject.Layer, """", """""") & """"
                exColor = extractObject.TrueColor.ColorIndex
                exStyle = """" & Replace(extractObject.StyleName, """", """""") & """"
                exText = """" & Replace(extractObject.TextString, """", """""") & """"
                exHeight = extractObject.Height
                exCoordinate = extractObject.InsertionPoint
                
                outputData = outputData & figureText & "," & exLayer & "," & _
                exColor & "," & exStyle & "," & exText & "," & exHeight & "," & _
                exCoordinate(0) & "," & exCoordinate(1) & "," & exCoordinate(2) & vbCrLf
                
            End If
        Next extractObject
        
    End If
    
    targetSelectionSet.Delete
    
    Dim filePath As String
    Dim outputFile As String
    
    filePath = Left(ThisDrawing.FullName, Len(ThisDrawing.FullName) - 4)
    outputFile = filePath & "_�e�L�X�g�f�[�^.csv"
    Open outputFile For Output As #1
    Print #1, outputData
    Close #1
    
    MsgBox "�e�L�X�g���o���������܂����B"
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "�G���[�F�R�}���h���I�����܂��B" & vbCrLf
    targetSelectionSet.Delete
    
End Sub
