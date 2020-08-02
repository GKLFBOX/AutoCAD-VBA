VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LayoutForm 
   Caption         =   "レイアウト編集"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   OleObjectBlob   =   "LayoutForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "LayoutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' ## 設定ファイルのファイル名
'------------------------------------------------------------------------------
Private Const LAYOUT_CONFIG As String = "\LayoutSetting.config"

'------------------------------------------------------------------------------
' ## フォーム初期化
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Dim i As Long
    Dim tempPlotConfig As ZcadPlotConfiguration
    
    ' 呼び出し用一時ページ設定
    Set tempPlotConfig = ThisDrawing.PlotConfigurations.Add("TempPlotConfig")
    tempPlotConfig.RefreshPlotDeviceInfo
    
    ' レイヤー名称呼び出し
    For i = 0 To ThisDrawing.Layers.Count - 1
        LayoutLayerBox.AddItem ThisDrawing.Layers.Item(i).Name
    Next i
    
    ' 印刷スタイル名称呼び出し
    Dim styleList As Variant
    styleList = tempPlotConfig.GetPlotStyleTableNames
    For i = 0 To UBound(styleList)
        StyleNameBox.AddItem styleList(i)
    Next i
    
    ' プリンタ名称呼び出し
    Dim printerList As Variant
    printerList = tempPlotConfig.GetPlotDeviceNames
    For i = 0 To UBound(printerList)
        PrinterNameBox.AddItem printerList(i)
    Next i
    
    ' 設定値読み込み
    Dim configData As Variant
    configData = Split(CommitConfig.LoadConfig(LAYOUT_CONFIG), vbCrLf)
    If Not UBound(configData) = 9 Then Exit Sub
    
    FrameTagBox.Value = configData(0)
    ScaleFactorBox.Value = configData(1)
    LayoutLayerBox.Value = configData(2)
    StyleNameBox.Value = configData(3)
    PrinterNameBox.Value = configData(4)
    A3PaperBox.Value = configData(5)
    A4PaperBox.Value = configData(6)
    OffsetXBox.Value = configData(7)
    OffsetYBox.Value = configData(8)
    CustomScaleBox.Value = configData(9)
    
    ' 用紙設定呼び出し
    Dim paperList As Variant
    tempPlotConfig.ConfigName = PrinterNameBox.Value
    paperList = tempPlotConfig.GetCanonicalMediaNames
    For i = 0 To UBound(paperList)
        If paperList(i) Like "*A3*" Then A3PaperBox.AddItem paperList(i)
        If paperList(i) Like "*A4*" Then A4PaperBox.AddItem paperList(i)
    Next i
    
End Sub

'------------------------------------------------------------------------------
' ## プリンタ名称更新時
'------------------------------------------------------------------------------
Private Sub PrinterNameBox_Change()
    
    Dim tempPlotConfig As ZcadPlotConfiguration
    
    ' 呼び出し用一時ページ設定
    Set tempPlotConfig = ThisDrawing.PlotConfigurations.Item("TempPlotConfig")
    
    ' 用紙設定リセット
    A3PaperBox.Value = ""
    A4PaperBox.Value = ""
    A3PaperBox.Clear
    A4PaperBox.Clear
    
    ' プリンタ名称の存在確認
    Dim printerList As Variant
    printerList = tempPlotConfig.GetPlotDeviceNames
    If Not CommonFunction.IsMatchList _
        (printerList, PrinterNameBox.Value) Then Exit Sub
    
    ' 用紙設定呼び出しおよび補完
    Dim i As Long
    Dim paperList As Variant
    tempPlotConfig.ConfigName = PrinterNameBox.Value
    paperList = tempPlotConfig.GetCanonicalMediaNames
    For i = 0 To UBound(paperList)
        If paperList(i) Like "A3" Then A3PaperBox.Value = paperList(i)
        If paperList(i) Like "A4" Then A4PaperBox.Value = paperList(i)
        If paperList(i) Like "*A3*" Then A3PaperBox.AddItem paperList(i)
        If paperList(i) Like "*A4*" Then A4PaperBox.AddItem paperList(i)
    Next i
    
End Sub

'------------------------------------------------------------------------------
' ## オフセット量入力ボタン
'------------------------------------------------------------------------------
Private Sub InputOffsetButton_Click()
    
    Dim tempPlotConfig As ZcadPlotConfiguration
    
    ' 呼び出し用一時ページ設定
    Set tempPlotConfig = ThisDrawing.PlotConfigurations.Item("TempPlotConfig")
    
    ' A3用紙設定の存在確認
    Dim paperList As Variant
    paperList = tempPlotConfig.GetCanonicalMediaNames
    If Not CommonFunction.IsMatchList _
        (paperList, A3PaperBox.Value) Then Exit Sub
    
    ' オフセット量入力(XYが一般と逆のため注意)
    ' 正しく値が取得できない用紙設定が存在するため注意
    Dim offsetLowerLeft As Variant, offsetUpperRight As Variant
    tempPlotConfig.GetPaperMargins offsetLowerLeft, offsetUpperRight
    OffsetXBox.Value = CSng((offsetLowerLeft(1) + offsetUpperRight(1)) / -2)
    OffsetYBox.Value = CSng((offsetLowerLeft(0) + offsetUpperRight(0)) / -2)
    
End Sub

'------------------------------------------------------------------------------
' ## 新規レイアウト作成ボタン
'------------------------------------------------------------------------------
Private Sub CreateLayoutButton_Click()
    
    Dim configData As Variant
    
    ' 設定値の入力確認
    If Not validateConfiguration() Then Exit Sub
    
    LayoutForm.Hide
    
    ' 新規レイアウト作成実行
    Call CreateLayout.CreateLayout(FrameTagBox.Value, _
                                   ScaleFactorBox.Value, _
                                   LayoutLayerBox.Value, _
                                   StyleNameBox.Value, _
                                   PrinterNameBox.Value, _
                                   A3PaperBox.Value, _
                                   A4PaperBox.Value, _
                                   OffsetXBox.Value, _
                                   OffsetYBox.Value)
    
    ' 設定値保存準備
    configData = FrameTagBox.Value & vbCrLf _
               & ScaleFactorBox.Value & vbCrLf _
               & LayoutLayerBox.Value & vbCrLf _
               & StyleNameBox.Value & vbCrLf _
               & PrinterNameBox.Value & vbCrLf _
               & A3PaperBox.Value & vbCrLf _
               & A4PaperBox.Value & vbCrLf _
               & OffsetXBox.Value & vbCrLf _
               & OffsetYBox.Value & vbCrLf _
               & CustomScaleBox.Value
    
    ' 設定フォルダの準備
    Call CommitConfig.PrepareConfigFolder
    
    ' 設定値保存
    Call CommitConfig.SaveConfig(LAYOUT_CONFIG, configData)
    
    LayoutForm.Show
    
End Sub

'------------------------------------------------------------------------------
' ## 設定値の入力確認
'------------------------------------------------------------------------------
Private Function validateConfiguration() As Boolean
    
    validateConfiguration = False
    
    Dim tempPlotConfig As ZcadPlotConfiguration
    
    ' 呼び出し用一時ページ設定
    Set tempPlotConfig = ThisDrawing.PlotConfigurations.Item("TempPlotConfig")
    
    ' 尺度入力確認
    If Not IsNumeric(ScaleFactorBox.Value) Or ScaleFactorBox.Value <= 0 Then
        MsgBox "尺度の入力が不正です。", vbCritical
        Exit Function
    End If
    
    ' レイアウト画層入力確認
    Dim i As Long
    Dim layerList() As Variant
    For i = 0 To ThisDrawing.Layers.Count - 1
        ReDim Preserve layerList(i)
        layerList(i) = ThisDrawing.Layers.Item(i).Name
    Next i
    If Not CommonFunction.IsMatchList(layerList, LayoutLayerBox.Value) Then
        MsgBox "レイアウト画層の入力が不正です。", vbCritical
        Exit Function
    End If
    
    ' 印刷スタイルの存在確認
    Dim styleList As Variant
    styleList = tempPlotConfig.GetPlotStyleTableNames
    If Not CommonFunction.IsMatchList(styleList, StyleNameBox.Value) Then
        MsgBox "印刷スタイルの入力が不正です。", vbCritical
        Exit Function
    End If
    
    ' プリンタ名称の存在確認
    Dim printerList As Variant
    printerList = tempPlotConfig.GetPlotDeviceNames
    If Not CommonFunction.IsMatchList(printerList, PrinterNameBox.Value) Then
        MsgBox "プリンタ名称の入力が不正です。", vbCritical
        Exit Function
    End If
    
    ' A3用紙設定およびA4用紙設定の存在確認
    Dim paperList As Variant
    paperList = tempPlotConfig.GetCanonicalMediaNames
    If Not CommonFunction.IsMatchList(paperList, A3PaperBox.Value) _
    Or Not CommonFunction.IsMatchList(paperList, A4PaperBox.Value) Then
        MsgBox "用紙設定の入力が不正です。", vbCritical
        Exit Function
    End If
    
    ' オフセット量入力確認
    If Not IsNumeric(OffsetXBox.Value) _
    Or Not IsNumeric(OffsetYBox.Value) Then
        MsgBox "オフセット量の入力が不正です。", vbCritical
        Exit Function
    End If
    
    ' 縮尺入力確認
    If Not IsNumeric(CustomScaleBox.Value) Or CustomScaleBox.Value <= 0 Then
        MsgBox "縮尺の入力が不正です。", vbCritical
        Exit Function
    End If
    
    validateConfiguration = True
    
End Function

'------------------------------------------------------------------------------
' ## ビューポート追加ボタン
'------------------------------------------------------------------------------
Private Sub AddViewportButton_Click()
    
    Dim configData As Variant
    
    ' 設定値の入力確認
    If Not validateConfiguration() Then Exit Sub
    
    LayoutForm.Hide
    
    ' ビューポート追加実行
    Call AddViewport.AddViewport(FrameTagBox.Value, _
                                 ScaleFactorBox.Value, _
                                 LayoutLayerBox.Value, _
                                 CustomScaleBox.Value)
    
    ' 設定値保存準備
    configData = FrameTagBox.Value & vbCrLf _
               & ScaleFactorBox.Value & vbCrLf _
               & LayoutLayerBox.Value & vbCrLf _
               & StyleNameBox.Value & vbCrLf _
               & PrinterNameBox.Value & vbCrLf _
               & A3PaperBox.Value & vbCrLf _
               & A4PaperBox.Value & vbCrLf _
               & OffsetXBox.Value & vbCrLf _
               & OffsetYBox.Value & vbCrLf _
               & CustomScaleBox.Value
    
    ' 設定フォルダの準備
    Call CommitConfig.PrepareConfigFolder
    
    ' 設定値保存
    Call CommitConfig.SaveConfig(LAYOUT_CONFIG, configData)
    
    LayoutForm.Show
    
End Sub

'------------------------------------------------------------------------------
' ## フォーム終了時に一時ページ設定を削除
'------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    
    Dim tempPlotConfig As ZcadPlotConfiguration
    Set tempPlotConfig = ThisDrawing.PlotConfigurations.Item("TempPlotConfig")
    tempPlotConfig.Delete
    
End Sub
