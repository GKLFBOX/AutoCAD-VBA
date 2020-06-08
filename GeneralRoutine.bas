Attribute VB_Name = "GeneralRoutine"
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## テキスト判定関数
'------------------------------------------------------------------------------
Public Function isTextObject(ByVal target_object As ZcadEntity) As Boolean
    
    isTextObject = False
    
    If (TypeOf target_object Is ZcadText) _
    Or (TypeOf target_object Is ZcadMText) Then
        isTextObject = True
    End If
    
End Function

'------------------------------------------------------------------------------
' ## ハイライトの解除
'------------------------------------------------------------------------------
'Public Sub ResetHighlight(ByVal target_object As ZcadEntity)
'
'    If Not target_object Is Nothing Then target_object.Highlight False
'
'End Sub
