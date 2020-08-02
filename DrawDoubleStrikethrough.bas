Attribute VB_Name = "DrawDoubleStrikethrough"
Option Explicit

'------------------------------------------------------------------------------
' ## 二重取り消し線の作図
'
' 指定した文字オブジェクトに同画層で二重取り消し線を作図する
'------------------------------------------------------------------------------
Public Sub DrawDoubleStrikethrough()
    
    On Error GoTo Error_Handler
    
    ' 対象の選択
    Dim pickPoint As Variant
    Dim targetText As ZcadEntity
    ThisDrawing.Utility.GetEntity targetText, pickPoint, _
        "文字オブジェクトを選択 [Cancel(ESC)]"
    
    If Not CommonFunction.IsTextObject(targetText) Then
        ThisDrawing.Utility.Prompt "エラー：文字を選択してください。" & vbCrLf
        Exit Sub
    End If
    
    ' FIXME: 誤ってマルチテキストを選択すると角度要素の削除のみが行われる
    ' 作図簡略化のために基点と角度を記憶し角度要素削除
    Dim targetPoint As Variant
    Dim targetAngle As Double
    targetPoint = targetText.insertionPoint
    targetAngle = targetText.Rotation
    targetText.Rotate targetPoint, targetAngle * -1
    
    ' 取り消し線の始終端計算
    Dim minExtent As Variant, maxExtent As Variant
    targetText.GetBoundingBox minExtent, maxExtent
    
    ' BoundingBoxの仕様変更に伴い傾斜角度を考慮する
    Dim targetOblique As Double
    Dim boxHeight As Double
    Dim exAmount As Double
    targetOblique = targetText.ObliqueAngle
    boxHeight = maxExtent(1) - minExtent(1)
    exAmount = boxHeight * Tan(targetOblique)
    maxExtent(0) = maxExtent(0) + exAmount
    
    Dim startPoint(0 To 2) As Double, endPoint(0 To 2) As Double
    
    startPoint(0) = minExtent(0) - ((maxExtent(1) - minExtent(1)) * 0.15)
    startPoint(1) = minExtent(1) + ((maxExtent(1) - minExtent(1)) / 3)
    startPoint(2) = 0
    
    endPoint(0) = maxExtent(0) + ((maxExtent(1) - minExtent(1)) * 0.15)
    endPoint(1) = minExtent(1) + ((maxExtent(1) - minExtent(1)) / 3)
    endPoint(2) = 0
    
    ' 取り消し線作図
    Dim strikeThrough(0 To 1) As ZcadLine, offsetLine As Variant
    
    Set strikeThrough(0) = ThisDrawing.ModelSpace.AddLine(startPoint, endPoint)
    offsetLine = strikeThrough(0).Offset((maxExtent(1) - minExtent(1)) / 3)
    
    Set strikeThrough(1) = offsetLine(0)
    
    ' 文字および取り消し線の角度を戻す
    targetText.Rotate targetPoint, targetAngle
    
    strikeThrough(0).Layer = targetText.Layer
    strikeThrough(0).Rotate targetPoint, targetAngle
    
    strikeThrough(1).Layer = targetText.Layer
    strikeThrough(1).Rotate targetPoint, targetAngle
    
    Call CommonSub.ResetHighlight(targetText)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ResetHighlight(targetText)
    ThisDrawing.Utility.Prompt "エラー：コマンドを終了します。" & vbCrLf
    
End Sub
