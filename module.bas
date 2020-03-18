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
        ' 4頂点のポリライン/水平および並行の判定
        Dim boundaryLine As ZcadLWPolyline
        Set boundaryLine = pickObject
        
        Dim boundaryPoints As Variant
        boundaryPoints = boundaryLine.Coordinates
        
        If UBound(boundaryPoints) = 7 Then
            boundaryPoints(0) = boundaryPoints(0)
            
            ' xy軸なりの長方形の証明
            ' 点1,点2の線分がxまたはyに並行かつ対角線の長さが等しい
            
            
            
        Else
            ThisDrawing.Utility.Prompt "四角形を選択してくだしあ" & vbCrLf
            pickObject.Highlight False
            Exit Sub
        End If
        
    Else
        ThisDrawing.Utility.Prompt "ポリラインを選択してくだしあ" & vbCrLf
        pickObject.Highlight False
        Exit Sub
    End If
    
    ' Coodinatesにz軸要素をつけて3点の配列に変換する
    
    Dim partialSet As ZcadSelectionSet
    Set partialSet = ThisDrawing.SelectionSets.Add("aaa")
    
    partialSet.SelectByPolygon zcSelectionSetFence, boundaryPoints
    
    partialSet.Erase
    
    partialSet.Delete
    
End Sub
