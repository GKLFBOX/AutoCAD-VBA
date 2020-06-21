Attribute VB_Name = "OutputText"
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## 文字オブジェクトのcsv出力プログラム
'
' 文字オブジェクトの内容と属性をcsv形式で出力する
'------------------------------------------------------------------------------
Public Sub OutputText()
    
    On Error GoTo Error_Handler
    
    ' 図題の選択
    Dim pickPoint As Variant
    Dim targetFigure As ZcadEntity
    ThisDrawing.Utility.GetEntity targetFigure, pickPoint, _
        "図題を選択 [Cancel(ESC)]"
    
    Call CommonSub.ResetHighlight(targetFigure)
    
    If Not CommonFunction.IsTextObject(targetFigure) Then
        ThisDrawing.Utility.Prompt "エラー：文字を選択してください。" & vbCrLf
        Exit Sub
    End If
    
    ThisDrawing.Utility.Prompt _
        "図題は「" & targetFigure.TextString & "」です。" & vbCrLf
    ThisDrawing.Utility.Prompt _
        "問題が無ければ出力範囲を選択してください。" & vbCrLf
    
    ' 出力対象を範囲選択
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    ' 出力データの準備
    Dim outputFile As String
    outputFile = CommonFunction.MakeFilePath("_テキストデータ", ".csv")
    
    Dim outputData As String
    If Dir(outputFile) = "" Then
        outputData = _
            "図題,画層,色,スタイル,内容,文字高さ,X座標,Y座標,Z座標" & vbCrLf
    End If
    
    ' 出力データの作成
    If Not targetSelectionSet.Count = 0 Then
        Call makeTextData(targetSelectionSet, targetFigure, outputData)
    End If
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    ' 出力データの書き出し
    Call outputCSV(outputFile, outputData)
    ThisDrawing.Utility.Prompt "テキスト抽出が完了しました。" & vbCrLf
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "エラー：コマンドを終了します。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## csvファイルへの出力
'------------------------------------------------------------------------------
Private Sub outputCSV(ByVal output_file As String, ByVal output_data As String)
    
    Open output_file For Append As #1
    Print #1, output_data
    Close #1
    
End Sub

'------------------------------------------------------------------------------
' ## csv形式のデータ作成
'------------------------------------------------------------------------------
Private Sub makeTextData(ByVal target_selectionset As ZcadSelectionSet, _
                         ByVal target_figure As ZcadEntity, _
                         ByRef output_data As String)
    
    ' 図題のcsv用文字列化整形処理
    Dim figureText As String
    figureText = formatString(target_figure.TextString)
    
    ' 文字列化処理とcsv形式データ作成
    Dim extractObject As ZcadEntity
    Dim exLayer As String
    Dim exColor As Long
    Dim exStyle As String
    Dim exText As String
    Dim exHeight As Double
    Dim exCoordinate As Variant
    
    For Each extractObject In target_selectionset
        If CommonFunction.IsTextObject(extractObject) Then
            
            With extractObject
                exLayer = formatString(.Layer)
                exColor = .TrueColor.ColorIndex
                exStyle = formatString(.StyleName)
                exText = formatString(.TextString)
                exHeight = .Height
                exCoordinate = .InsertionPoint
            End With
            
            output_data = output_data & _
                figureText & "," & _
                exLayer & "," & _
                exColor & "," & _
                exStyle & "," & _
                exText & "," & _
                exHeight & "," & _
                exCoordinate(0) & "," & _
                exCoordinate(1) & "," & _
                exCoordinate(2) & vbCrLf
            
        End If
    Next extractObject
    
    output_data = Left(output_data, Len(output_data) - 2) ' 最終行の改行削除
    
End Sub

'------------------------------------------------------------------------------
' ## csv用の文字列整形(ダブルクォーテーションの付加と文字化)
'------------------------------------------------------------------------------
Private Function formatString(ByVal target_text As String) As String
    
    formatString = """" & Replace(target_text, """", """""") & """"
    
End Function
