VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrameListForm 
   Caption         =   "用紙枠リストcsv出力"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "FrameListForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FrameListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' ## 設定ファイルのファイル名
'------------------------------------------------------------------------------
Private Const FRAMELIST_CONFIG As String = "\FrameList.config"

'------------------------------------------------------------------------------
' ## フォーム初期化
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    ' ブロック名称呼び出し
    Dim i As Long
    Dim buf As String
    For i = 0 To ThisDrawing.Blocks.Count - 1
        buf = ThisDrawing.Blocks.Item(i).Name
        If Left(buf, 1) <> "*" Then FrameBlockNameBox.AddItem buf
    Next i
    
    ' 設定値読み込み
    Dim configData As Variant
    configData = Split(CommitConfig.LoadConfig(FRAMELIST_CONFIG), vbCrLf)
    If Not UBound(configData) = 1 Then Exit Sub
    
    FrameBlockNameBox.Value = configData(0)
    FrameTagBox.Value = configData(1)
    
End Sub

'------------------------------------------------------------------------------
' ## 用紙枠リスト出力ボタン
'------------------------------------------------------------------------------
Private Sub OutputFrameListButton_Click()
    
    Dim configData As Variant
    
    ' 設定値の入力確認
    If Not validateFrameListConfig() Then Exit Sub
    
    FrameListForm.Hide
    
    ' 用紙枠リスト出力実行
    Call OutputFrameList.OutputFrameList(FrameBlockNameBox.Value, _
                                         FrameTagBox.Value)
    
    ' 設定値保存準備
    configData = FrameBlockNameBox.Value & vbCrLf _
               & FrameTagBox.Value
    
    ' 設定フォルダの準備
    Call CommitConfig.PrepareConfigFolder
    
    ' 設定値保存
    Call CommitConfig.SaveConfig(FRAMELIST_CONFIG, configData)
    
    Unload FrameListForm
    
End Sub

'------------------------------------------------------------------------------
' ## 設定値入力の確認
'------------------------------------------------------------------------------
Private Function validateFrameListConfig() As Boolean
    
    validateFrameListConfig = False
    
    ' 用紙枠ブロック名入力確認
    Dim i As Long
    Dim blockList() As Variant
    For i = 0 To ThisDrawing.Blocks.Count - 1
        ReDim Preserve blockList(i)
        blockList(i) = ThisDrawing.Blocks.Item(i).Name
    Next i
    If Not CommonFunction.IsMatchList(blockList, FrameBlockNameBox.Value) Then
        MsgBox "用紙枠ブロック名の入力が不正です。", vbCritical
        Exit Function
    End If
    
    validateFrameListConfig = True
    
End Function
