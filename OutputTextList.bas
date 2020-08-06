Attribute VB_Name = "OutputTextList"
Option Explicit

'------------------------------------------------------------------------------
' ## 文字オブジェクトのcsv出力プログラム   2020/07/26 G.O.
'
' 図題ごとに選択範囲の文字オブジェクトをcsv形式のリストで出力する
' 出力する情報は画層,色,フォント,内容,高さ,座標
'------------------------------------------------------------------------------
Public Sub OutputTextList()
    
    On Error GoTo Error_Handler
    
    Dim outputFile As String
    Dim outputData As String
    Dim figureList() As String
    Dim figureText As String
    
    ' 図題の選択
    Dim targetFigure As ZcadEntity
    Dim pickPoint As Variant
    ThisDrawing.Utility.GetEntity targetFigure, pickPoint, _
        "図題を選択 [Cancel(ESC)]"
    Call CommonSub.ResetHighlight(targetFigure)
    If Not CommonFunction.IsTextObject(targetFigure) Then
        ThisDrawing.Utility.Prompt "文字が選択されませんでした。" & vbCrLf
        Exit Sub
    End If
    
    figureText = targetFigure.TextString
    ReDim Preserve figureList(0)
    
    ' 出力データヘッダ生成または図題重複回避処理
    outputFile = CommonFunction.MakeFilePath("_テキストデータ", ".csv")
    If Dir(outputFile) = "" Then
        outputData = makeHeader()
    Else
        Call makeFigureList(outputFile, figureList())
    End If
    
    Call avoidDuplicateFigure(figureText, figureList())
    
    ' 出力対象を範囲選択
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    ' 出力データの作成および書き出し
    Call makeTextData(targetSelectionSet, figureText, outputData)
    If outputData = "" Then
        ThisDrawing.Utility.Prompt "テキストが選択されませんでした。" & vbCrLf
    Else
        Call CommonSub.OutputCSV(outputFile, outputData)
        ThisDrawing.Utility.Prompt "テキスト抽出が完了しました。" & vbCrLf
    End If
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "なんらかのエラーです。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## ヘッダ生成
'------------------------------------------------------------------------------
Private Function makeHeader() As String
    
    makeHeader = """図題""," _
               & """画層""," _
               & """色""," _
               & """スタイル""," _
               & """内容""," _
               & """文字高さ""," _
               & """X座標""," _
               & """Y座標""," _
               & """Z座標""" & vbCrLf
    
End Function

'------------------------------------------------------------------------------
' ## 図題リストの生成
'------------------------------------------------------------------------------
Private Sub makeFigureList(ByVal output_file As String, _
                           ByRef figure_list() As String)
    
    Dim i As Long
    Dim bufferData As String
    
    Open output_file For Input As #1
        i = 0
        Line Input #1, bufferData
        Do Until EOF(1)
            Line Input #1, bufferData
            ReDim Preserve figure_list(0 To i)
            figure_list(i) = Left(bufferData, InStr(bufferData, """,") - 1)
            figure_list(i) = Right(figure_list(i), Len(figure_list(i)) - 1)
            i = i + 1
        Loop
    Close #1
    
End Sub

'------------------------------------------------------------------------------
' ## 図題重複回避処理
'------------------------------------------------------------------------------
Private Sub avoidDuplicateFigure(ByRef figure_text As String, _
                                 ByRef figure_list() As String)
    
    Dim i As Long
    Dim buffer_text As String
    
    i = 1
    buffer_text = figure_text
    Do While CommonFunction.IsMatchList(figure_list, figure_text)
        i = i + 1
        figure_text = buffer_text & " (" & i & ")"
    Loop
    
    ' 図題確認プロンプト
    ThisDrawing.Utility.Prompt _
        "図題は「" & figure_text & "」です。" & vbCrLf
    ThisDrawing.Utility.Prompt _
        "問題が無ければ出力範囲を選択してください。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## csv形式のデータ作成
'------------------------------------------------------------------------------
Private Sub makeTextData(ByVal target_selectionset As ZcadSelectionSet, _
                         ByVal figure_text As String, _
                         ByRef output_data As String)
    
    Dim bufferData As String
    Dim extractEntity As ZcadEntity
    Dim extractLayer As String
    Dim extractColor As Long
    Dim extractStyle As String
    Dim extractText As String
    Dim extractHeight As Double
    Dim extractCoordinate As Variant
    
    ' データ作成前の値を保存
    bufferData = output_data
    
    ' 文字列化およびcsv形式データ作成
    figure_text = CommonFunction.FormatString(figure_text)
    For Each extractEntity In target_selectionset
        
        If Not CommonFunction.IsTextObject(extractEntity) Then _
            GoTo Continue_extractEntity
            
        With extractEntity
            extractLayer = CommonFunction.FormatString(.Layer)
            extractColor = .TrueColor.ColorIndex
            extractStyle = CommonFunction.FormatString(.StyleName)
            extractText = CommonFunction.FormatString(.TextString)
            extractHeight = .Height
            extractCoordinate = .insertionPoint
        End With
        
        output_data = output_data _
                    & figure_text & "," _
                    & extractLayer & "," _
                    & extractColor & "," _
                    & extractStyle & "," _
                    & extractText & "," _
                    & extractHeight & "," _
                    & extractCoordinate(0) & "," _
                    & extractCoordinate(1) & "," _
                    & extractCoordinate(2) & vbCrLf
        
Continue_extractEntity:
        
    Next extractEntity
    
    ' 値の削除または最終行の改行削除
    If bufferData = output_data Then
        output_data = ""
    Else
        output_data = Left(output_data, Len(output_data) - 2)
    End If
    
End Sub
