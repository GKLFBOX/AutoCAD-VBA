Attribute VB_Name = "EditDimension"
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## 寸法オブジェクトの文字(寸法値)の色のみを赤にする
'
' 寸法スタイルに依存せずに寸法値の文字色のみを色変更で戻せる状態で赤にする
' オブジェクト=赤, 文字(寸法値)=ByBlock, 寸法線および寸法補助線=ByLayer
'------------------------------------------------------------------------------
Public Sub TurnRedDimension()
    
    On Error GoTo Error_Handler
    
    ' TODO: 可能であればLispを利用して事前選択を実装する
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    If Not targetSelectionSet.Count = 0 Then
        
        Dim returnObject As ZcadEntity
        For Each returnObject In targetSelectionSet
            If TypeOf returnObject Is ZcadDimension Then
                returnObject.TextColor = zcByBlock
                returnObject.Color = zcRed  ' TODO: colorは使うべきでない
            End If
        Next returnObject
        
    End If
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "エラー：コマンドを終了します。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## 寸法線または引出線のサイズ変更
'
' 寸法線または引出線の全体の寸法尺度を変更する
'------------------------------------------------------------------------------
Public Sub ResizeDimensionSize()
    
    On Error GoTo Error_Handler
    
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    If Not targetSelectionSet.Count = 0 Then
        
        Dim returnObject As ZcadEntity
        For Each returnObject In targetSelectionSet
            If TypeOf returnObject Is ZcadDimension _
            Or TypeOf returnObject Is ZcadLeader Then
                returnObject.Highlight True
            End If
        Next returnObject
        
        Dim sizeFactor As Long
        sizeFactor = ThisDrawing.Utility.GetInteger _
            ("変更尺度を入力 または [25/50/80/100]:")
        
        For Each returnObject In targetSelectionSet
            If TypeOf returnObject Is ZcadDimension _
            Or TypeOf returnObject Is ZcadLeader Then
                returnObject.ScaleFactor = sizeFactor
            End If
        Next returnObject
        
    End If
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    Exit Sub
    
Error_Handler:
    For Each returnObject In targetSelectionSet
        Call CommonSub.ResetHighlight(returnObject)
    Next returnObject
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "エラー：コマンドを終了します。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## 寸法線の文字オフセット量変更
'
' 寸法線の文字オフセット量(寸法線と文字の離れ量)を変更する
'------------------------------------------------------------------------------
Public Sub AdjustDimensionOffset()
    
    On Error GoTo Error_Handler
    
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    If Not targetSelectionSet.Count = 0 Then
        
        Dim returnObject As ZcadEntity
        For Each returnObject In targetSelectionSet
            If TypeOf returnObject Is ZcadDimension Then
                returnObject.Highlight True
            End If
        Next returnObject
        
        Dim offsetAmount As Double
        offsetAmount = ThisDrawing.Utility.GetInteger _
            ("変更オフセット量を入力 または [デフォルト(8)/小さめ(5)]:")
        offsetAmount = offsetAmount * 0.1
        
        For Each returnObject In targetSelectionSet
            If TypeOf returnObject Is ZcadDimension Then
                returnObject.TextGap = offsetAmount
            End If
        Next returnObject
        
    End If
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    Exit Sub
    
Error_Handler:
    For Each returnObject In targetSelectionSet
        Call CommonSub.ResetHighlight(returnObject)
    Next returnObject
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "エラー：コマンドを終了します。" & vbCrLf
    
End Sub
