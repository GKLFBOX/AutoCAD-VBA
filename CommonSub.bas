Attribute VB_Name = "CommonSub"
'------------------------------------------------------------------------------
' ## �R�[�f�B���O�K�C�h���C��
'
' [You.Activate|VBA�R�[�f�B���O�K�C�h���C��]�ɏ�������
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## �n�C���C�g�̉���
'------------------------------------------------------------------------------
Public Sub ResetHighlight(ByVal target_object As ZcadEntity)

    If Not target_object Is Nothing Then target_object.Highlight False

End Sub

'------------------------------------------------------------------------------
' ## �I���Z�b�g�̍폜
'------------------------------------------------------------------------------
Public Sub ReleaseSelectionSet(ByVal target_selectionset As ZcadSelectionSet)

    If Not target_selectionset Is Nothing Then target_selectionset.Delete

End Sub
