Attribute VB_Name = "GeneralRoutine"
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
Public Function isTextObject(ByVal target_object As ZcadEntity) As Boolean
    
    isTextObject = False
    
    If (TypeOf target_object Is ZcadText) _
    Or (TypeOf target_object Is ZcadMText) Then
        isTextObject = True
    End If
    
End Function

'------------------------------------------------------------------------------
' ## �n�C���C�g�̉���
'------------------------------------------------------------------------------
'Public Sub ResetHighlight(ByVal target_object As ZcadEntity)
'
'    If Not target_object Is Nothing Then target_object.Highlight False
'
'End Sub
