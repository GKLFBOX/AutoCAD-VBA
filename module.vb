'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## 部分的な図面比較
'
' 図面の一部と基点を指定し2図面の部分的な比較を行う
'------------------------------------------------------------------------------
Public Sub ComparePart()
    
    Dim pickPoint As Variant
    
    
    
    Dim partialSet As ZcadSelectionSet
    Set partialSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    
    
End Sub
