Attribute VB_Name = "CommonFunction"
Option Explicit

'------------------------------------------------------------------------------
' ## �e�L�X�g����
'------------------------------------------------------------------------------
Public Function IsTextObject(ByVal target_object As ZcadEntity) As Boolean
    
    IsTextObject = False
    
    If TypeOf target_object Is ZcadText _
    Or TypeOf target_object Is ZcadMText Then
        IsTextObject = True
    End If
    
End Function

'------------------------------------------------------------------------------
' ## ���t�@�C����+�C�Ӗ��̂ɂ��t�@�C���p�X�����֐�
'------------------------------------------------------------------------------
Public Function MakeFilePath(ByVal addition_name As String, _
                             ByVal file_extension As String) As String
    
    MakeFilePath = Left(ThisDrawing.fullName, Len(ThisDrawing.fullName) - 4)
    MakeFilePath = MakeFilePath & addition_name & file_extension
    
End Function

'------------------------------------------------------------------------------
' ## ������̃��X�g�ƍ�
'------------------------------------------------------------------------------
Public Function IsMatchList(ByVal target_list As Variant, _
                            ByVal target_value As String) As Boolean
    
    IsMatchList = False
    
    Dim i As Long
    For i = 0 To UBound(target_list)
        If target_value = target_list(i) Then
            IsMatchList = True
            Exit Function
        End If
    Next i
    
End Function

'------------------------------------------------------------------------------
' ## �z���IsEmpty
'------------------------------------------------------------------------------
Public Function IsEmptyArray(ByRef confirmation_array As Variant) As Boolean
    
    On Error GoTo Error_Handler
    
    ' �G���[�܂��͍ő�v�f����0�����̏ꍇ�͋�
    IsEmptyArray = IIf(UBound(confirmation_array) < 0, True, False)
    
    Exit Function
    
Error_Handler:
    IsEmptyArray = True
    
End Function

'------------------------------------------------------------------------------
' ## csv�p�̕�����(�_�u���N�H�[�e�[�V�����̕t��)
'------------------------------------------------------------------------------
Public Function FormatString(ByVal target_text As String) As String
    
    FormatString = """" & Replace(target_text, """", """""") & """"
    
End Function
