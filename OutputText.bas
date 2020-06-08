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
Sub OutputText()
    
    'On Error GoTo Error_Handler
    
    ' 図題の選択
    Dim pickPoint As Variant
    Dim targetFigure As ZcadEntity
    ThisDrawing.Utility.GetEntity targetFigure, pickPoint, _
        "図題(文字オブジェクト)を選択 [Cancel(ESC)]"
    
    targetFigure.Highlight True
    
    ' csv向けに図題の文字列化処理
    Dim figureText As String
    figureText = _
    """" & Replace(targetFigure.TextString, """", """""") & """"
    
    If Not GeneralRoutine.isTextObject(targetFigure) Then
        targetFigure.Highlight False
        ThisDrawing.Utility.Prompt "文字を選択してください。" & vbCrLf
        Exit Sub
    End If
    
    targetFigure.Highlight False
    
    ' 出力対象を範囲選択
    Dim targetSelectionSet As ZcadSelectionSet
    
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    ' 出力データの作成
    If Not targetSelectionSet.Count = 0 Then
        
        Dim extractObject As ZcadEntity  ' 変数名が思いつかないため要検討
        Dim exLayer As String
        Dim exColor As Long
        Dim exStyle As String
        Dim exText As String
        Dim exHeight As Double
        Dim exCoordinate As Variant
        
        Dim outputData As String
        
        outputData = _
        "図題,画層,色,スタイル,内容,文字高さ,X座標,Y座標,Z座標" & vbCrLf
        
        For Each extractObject In targetSelectionSet
            If GeneralRoutine.isTextObject(extractObject) Then
                
                exLayer = """" & Replace(extractObject.Layer, """", """""") & """"
                exColor = extractObject.TrueColor.ColorIndex
                exStyle = """" & Replace(extractObject.StyleName, """", """""") & """"
                exText = """" & Replace(extractObject.TextString, """", """""") & """"
                exHeight = extractObject.Height
                exCoordinate = extractObject.InsertionPoint
                
                outputData = outputData & figureText & "," & exLayer & "," & _
                exColor & "," & exStyle & "," & exText & "," & exHeight & "," & _
                exCoordinate(0) & "," & exCoordinate(1) & "," & exCoordinate(2) & vbCrLf
                
            End If
        Next extractObject
        
    End If
    
    targetSelectionSet.Delete
    
    Dim filePath As String
    Dim outputFile As String
    
    filePath = Left(ThisDrawing.FullName, Len(ThisDrawing.FullName) - 4)
    outputFile = filePath & "_テキストデータ.csv"
    Open outputFile For Output As #1
    Print #1, outputData
    Close #1
    
    MsgBox "テキスト抽出が完了しました。"
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "エラー：コマンドを終了します。" & vbCrLf
    targetSelectionSet.Delete
    
End Sub
