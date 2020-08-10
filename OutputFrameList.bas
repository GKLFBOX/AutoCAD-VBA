Attribute VB_Name = "OutputFrameList"
Option Explicit

'------------------------------------------------------------------------------
' ## 用紙枠リストのcsv出力プログラム   2020/08/06 G.O.
'
' 選択範囲の用紙枠データをcsv形式のリストで出力する
' 出力する情報はページ番号,画層,色,座標
'------------------------------------------------------------------------------
Public Sub OutputFrameList(ByVal frame_blockname As String, _
                           ByVal frame_tag As String)
    
    On Error GoTo Error_Handler
    
    Dim outputFile As String
    Dim outputData As String
    
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts("Model")
    
    ' 出力データヘッダ生成
    outputFile = CommonFunction.MakeFilePath("_用紙枠データ", ".csv")
    If Dir(outputFile) = "" Then outputData = makeHeader()
    
    ThisDrawing.Utility.Prompt "出力範囲を選択してください。" & vbCrLf
    
    ' 出力対象を範囲選択
    Dim targetSelectionSet As ZcadSelectionSet
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    ' 出力データの作成および書き出し
    Call makeFrameData _
        (targetSelectionSet, frame_blockname, frame_tag, outputData)
    If outputData = "" Then
        ThisDrawing.Utility.Prompt "用紙枠が選択されませんでした。" & vbCrLf
    Else
        Call CommonSub.OutputCSV(outputFile, outputData)
        ThisDrawing.Utility.Prompt "用紙枠データ出力が完了しました。" & vbCrLf
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
    
    makeHeader = """ページ番号""," _
               & """画層""," _
               & """色""," _
               & """左下X座標""," _
               & """左下Y座標""," _
               & """左下Z座標""," _
               & """右上X座標""," _
               & """右上Y座標""," _
               & """右上Z座標""" & vbCrLf
    
End Function

'------------------------------------------------------------------------------
' ## csv形式のデータ作成
'------------------------------------------------------------------------------
Private Sub makeFrameData(ByVal target_selectionset As ZcadSelectionSet, _
                          ByVal frame_blockname As String, _
                          ByVal frame_tag As String, _
                          ByRef output_data As String)
    
    Dim bufferData As String
    Dim extractEntity As ZcadEntity
    Dim extractBlock As ZcadBlockReference
    Dim extractPageNo As String
    Dim extractLayer As String
    Dim extractColor As Long
    Dim extractMinPoint As Variant, extractMaxPoint As Variant
    
    ' データ作成前の値を保存
    bufferData = output_data
    
    ' 文字列化およびcsv形式データ作成
    For Each extractEntity In target_selectionset
        
        ' 指定ブロックでない場合はスキップ
        If Not TypeOf extractEntity Is ZcadBlockReference Then _
            GoTo Continue_extractEntity
        
        Set extractBlock = extractEntity
        
        If Not extractBlock.Name = frame_blockname Then _
            GoTo Continue_extractEntity
        
        ' ページ番号取得およびページ座標取得
        Call CommonSub.FetchFrameName _
            (extractBlock, frame_tag, extractPageNo)
        Call CommonSub.FetchCorrectSize _
            (extractBlock, extractMinPoint, extractMaxPoint)
        
        extractPageNo = CommonFunction.FormatString(extractPageNo)
        extractLayer = CommonFunction.FormatString(extractBlock.Layer)
        extractColor = extractBlock.TrueColor.ColorIndex
        
        output_data = output_data _
                    & extractPageNo & "," _
                    & extractLayer & "," _
                    & extractColor & "," _
                    & extractMinPoint(0) & "," _
                    & extractMinPoint(1) & "," _
                    & extractMinPoint(2) & "," _
                    & extractMaxPoint(0) & "," _
                    & extractMaxPoint(1) & "," _
                    & extractMaxPoint(2) & vbCrLf
        
        Erase extractMinPoint, extractMaxPoint
        
Continue_extractEntity:
        
    Next extractEntity
    
    ' 値の削除または最終行の改行削除
    If bufferData = output_data Then
        output_data = ""
    Else
        output_data = Left(output_data, Len(output_data) - 2)
    End If
    
End Sub
