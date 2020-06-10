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
    
    On Error GoTo Error_Handler
    
    ' 図題の選択
    Dim pickPoint As Variant
    Dim targetFigure As ZcadEntity
    ThisDrawing.Utility.GetEntity targetFigure, pickPoint, _
        "図題を選択 [Cancel(ESC)]"
    
    targetFigure.Highlight False
    
    If Not GeneralRoutine.isTextObject(targetFigure) Then
        ThisDrawing.Utility.Prompt "エラー：文字を選択してください。" & vbCrLf
        Exit Sub
    End If
    
    ThisDrawing.Utility.Prompt _
        "図題は「" & targetFigure.TextString & "」です。" & vbCrLf
    ThisDrawing.Utility.Prompt _
        "問題が無ければ出力範囲を選択してください。" & vbCrLf
    
    ' 図題のcsv用文字列化処理
    Dim figureText As String
    figureText = _
    """" & Replace(targetFigure.TextString, """", """""") & """"
    
    ' 出力対象を範囲選択
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    ' 出力データの準備
    Dim filePath As String
    Dim outputFile As String
    
    filePath = Left(ThisDrawing.FullName, Len(ThisDrawing.FullName) - 4)
    outputFile = filePath & "_テキストデータ.csv"
    
    Dim outputData As String
    If Dir(outputFile) = "" Then
        outputData = _
            "図題,画層,色,スタイル,内容,文字高さ,X座標,Y座標,Z座標" & vbCrLf
    End If
    
    ' 出力データの作成
    ' TODO: プロシージャを分割する
    If Not targetSelectionSet.Count = 0 Then
        
        Dim extractObject As ZcadEntity
        Dim exLayer As String
        Dim exColor As Long
        Dim exStyle As String
        Dim exText As String
        Dim exHeight As Double
        Dim exCoordinate As Variant
        
        For Each extractObject In targetSelectionSet
            If GeneralRoutine.isTextObject(extractObject) Then
                
                With extractObject
                    exLayer = """" & Replace(.Layer, """", """""") & """"
                    exColor = .TrueColor.ColorIndex
                    exStyle = """" & Replace(.StyleName, """", """""") & """"
                    exText = """" & Replace(.TextString, """", """""") & """"
                    exHeight = .Height
                    exCoordinate = .InsertionPoint
                End With
                
                outputData = outputData & _
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
        
    End If
    
    outputData = Left(outputData, Len(outputData) - 2)
    targetSelectionSet.Delete
    
    ' 出力データの書き出し
    Open outputFile For Append As #1
    Print #1, outputData
    Close #1
    
    ThisDrawing.Utility.Prompt "テキスト抽出が完了しました。" & vbCrLf
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "エラー：コマンドを終了します。" & vbCrLf
    targetSelectionSet.Delete
    
End Sub
