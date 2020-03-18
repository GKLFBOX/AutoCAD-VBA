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
    Dim pickObject As ZcadEntity
    ThisDrawing.Utility.GetEntity pickObject, pickPoint, _
        "境界オブジェクトを選択 [Cancel(ESC)]"
    
    pickObject.Highlight True
    
    If TypeOf pickObject Is ZcadLWPolyline Then
        ' 矩形に指定する
        Dim boundaryLine As ZcadLWPolyline
        Set boundaryLine = pickObject
        
        Dim boundaryPoints As Variant
        boundaryPoints = boundaryLine.Coordinates
        
        
    Else
        ThisDrawing.Utility.Prompt "ポリラインを選択してくだしあ" & vbCrLf
        'pickObject.Highlight False
        Exit Sub
    End If
    
    Dim partialSet As ZcadSelectionSet
    Set partialSet = ThisDrawing.SelectionSets.Add("aaa")
    
    partialSet.SelectByPolygon zcSelectionSetFence, boundaryPoints
    
    partialSet.Erase
    
    partialSet.Delete
    
End Sub
