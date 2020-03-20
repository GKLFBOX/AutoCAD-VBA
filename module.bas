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
        
        Dim boundaryLine As ZcadLWPolyline
        Set boundaryLine = pickObject
        
        Dim boundaryPoints As Variant
        boundaryPoints = boundaryLine.Coordinates
        
        ' 4頂点のポリライン判定
        If UBound(boundaryPoints) = 7 Then
            
            Dim n As Long
            Dim point_x(0 To 3) As Double
            Dim point_y(0 To 3) As Double
            
            For n = 0 To 3
                point_x(n) = boundaryPoints(n + n)
                point_y(n) = boundaryPoints(n + n + 1)
            Next n
            
            ' 対角線の長さが等しく、それぞれの中点で交わるとき長方形になる
            
            
            
        Else
            ThisDrawing.Utility.Prompt "(test)四角形を選択してくだしあ" & vbCrLf
            pickObject.Highlight False
            Exit Sub
        End If
        
    Else
        ThisDrawing.Utility.Prompt "(test)ポリラインを選択してくだしあ" & vbCrLf
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
