Attribute VB_Name = "OutputFrameList"
Option Explicit

'------------------------------------------------------------------------------
' ## �p���g���X�g��csv�o�̓v���O����   2020/08/06 G.O.
'
' �I��͈̗͂p���g�f�[�^��csv�`���̃��X�g�ŏo�͂���
' �o�͂�����̓y�[�W�ԍ�,��w,�F,���W
'------------------------------------------------------------------------------
Public Sub OutputFrameList(ByVal frame_blockname As String, _
                           ByVal frame_tag As String)
    
    On Error GoTo Error_Handler
    
    Dim outputFile As String
    Dim outputData As String
    
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts("Model")
    
    ' �o�̓f�[�^�w�b�_����
    outputFile = CommonFunction.MakeFilePath("_�p���g�f�[�^", ".csv")
    If Dir(outputFile) = "" Then outputData = makeHeader()
    
    ThisDrawing.Utility.Prompt "�o�͔͈͂�I�����Ă��������B" & vbCrLf
    
    ' �o�͑Ώۂ�͈͑I��
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    ' �o�̓f�[�^�̍쐬����я����o��
    Call makeFrameData _
        (targetSelectionSet, frame_blockname, frame_tag, outputData)
    If outputData = "" Then
        ThisDrawing.Utility.Prompt "�p���g���I������܂���ł����B" & vbCrLf
    Else
        Call CommonSub.OutputCSV(outputFile, outputData)
        ThisDrawing.Utility.Prompt "�p���g�f�[�^�o�͂��������܂����B" & vbCrLf
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
    
    makeHeader = """�y�[�W�ԍ�""," _
               & """��w""," _
               & """�F""," _
               & """����X���W""," _
               & """����Y���W""," _
               & """����Z���W""," _
               & """�E��X���W""," _
               & """�E��Y���W""," _
               & """�E��Z���W""" & vbCrLf
    
End Function

'------------------------------------------------------------------------------
' ## csv�`���̃f�[�^�쐬
'------------------------------------------------------------------------------
Private Sub makeFrameData(ByVal target_selectionset As ZcadSelectionSet, _
                          ByVal frame_blockname As String, _
                          ByVal frame_tag As String, _
                          ByRef output_data As String)
    
    Dim bufferData As String
    Dim extractEntity As ZcadEntity
    Dim extractBlock As ZcadBlockReference
    Dim extractPageNo As String
    Dim extractLayer As String
    Dim extractColor As Long
    Dim extractMinPoint As Variant, extractMaxPoint As Variant
    
    ' �f�[�^�쐬�O�̒l��ۑ�
    bufferData = output_data
    
    ' �����񉻂����csv�`���f�[�^�쐬
    For Each extractEntity In target_selectionset
        
        ' �w��u���b�N�łȂ��ꍇ�̓X�L�b�v
        If Not TypeOf extractEntity Is ZcadBlockReference Then _
            GoTo Continue_extractEntity
        
        Set extractBlock = extractEntity
        
        If Not extractBlock.Name = frame_blockname Then _
            GoTo Continue_extractEntity
        
        ' �y�[�W�ԍ��擾����уy�[�W���W�擾
        Call CommonSub.FetchFrameName _
            (extractBlock, frame_tag, extractPageNo)
        Call CommonSub.FetchCorrectSize _
            (extractBlock, extractMinPoint, extractMaxPoint)
        
        extractPageNo = CommonFunction.FormatString(extractPageNo)
        extractLayer = CommonFunction.FormatString(extractBlock.Layer)
        extractColor = extractBlock.TrueColor.ColorIndex
        
        output_data = output_data _
                    & extractPageNo & "," _
                    & extractLayer & "," _
                    & extractColor & "," _
                    & extractMinPoint(0) & "," _
                    & extractMinPoint(1) & "," _
                    & extractMinPoint(2) & "," _
                    & extractMaxPoint(0) & "," _
                    & extractMaxPoint(1) & "," _
                    & extractMaxPoint(2) & vbCrLf
        
        Erase extractMinPoint, extractMaxPoint
        
Continue_extractEntity:
        
    Next extractEntity
    
    ' �l�̍폜�܂��͍ŏI�s�̉��s�폜
    If bufferData = output_data Then
        output_data = ""
    Else
        output_data = Left(output_data, Len(output_data) - 2)
    End If
    
End Sub
