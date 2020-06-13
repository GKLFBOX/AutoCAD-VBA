Attribute VB_Name = "CommonSub"
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## ハイライトの解除
'------------------------------------------------------------------------------
Public Sub ResetHighlight(ByVal target_object As ZcadEntity)

    If Not target_object Is Nothing Then target_object.Highlight False

End Sub

'------------------------------------------------------------------------------
' ## 選択セットの削除
'------------------------------------------------------------------------------
Public Sub ReleaseSelectionSet(ByVal target_selectionset As ZcadSelectionSet)

    If Not target_selectionset Is Nothing Then target_selectionset.Delete

End Sub
