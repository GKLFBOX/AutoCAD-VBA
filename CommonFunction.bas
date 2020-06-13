Attribute VB_Name = "CommonFunction"
'------------------------------------------------------------------------------
' ## �R�[�f�B���O�K�C�h���C��
'
' [You.Activate|VBA�R�[�f�B���O�K�C�h���C��]�ɏ�������
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## �e�L�X�g����֐�
'------------------------------------------------------------------------------
Public Function IsTextObject(ByVal target_object As ZcadEntity) As Boolean
    
    IsTextObject = False
    
    If (TypeOf target_object Is ZcadText) _
    Or (TypeOf target_object Is ZcadMText) Then
        IsTextObject = True
    End If
    
End Function

'------------------------------------------------------------------------------
' ## ���t�@�C����+�C�Ӗ��̂ɂ��t�@�C���p�X�����֐�
'------------------------------------------------------------------------------
Public Function MakeFilePath(ByVal addition_name As String, _
                             ByVal file_extension As String) As String
    
    MakeFilePath = Left(ThisDrawing.FullName, Len(ThisDrawing.FullName) - 4)
    MakeFilePath = MakeFilePath & addition_name & file_extension
    
End Function
