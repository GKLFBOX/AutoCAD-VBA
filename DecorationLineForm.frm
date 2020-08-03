VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DecorationLineForm 
   Caption         =   "文字装飾線設定"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   OleObjectBlob   =   "DecorationLineForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "DecorationLineForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------------------------------
' ## フォーム初期化
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Dim i As Long
    Dim configData As Variant
    
    ' 参照線作図設定レイヤー名称呼び出し
    For i = 0 To ThisDrawing.Layers.Count - 1
        ReferenceLineLayerBox.AddItem ThisDrawing.Layers.Item(i).Name
    Next i
    
    ' 取り消し線作図設定レイヤー名称呼び出し
    For i = 0 To ThisDrawing.Layers.Count - 1
        StrikethroughLayerBox.AddItem ThisDrawing.Layers.Item(i).Name
    Next i
    
    ' 参照線作図設定値読み込み
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.REFERENCELINE_CONFIG), vbCrLf)
    If UBound(configData) = 3 Then
        ReferenceLineLayerOnBox.Value = configData(0)
        ReferenceLineLayerBox.Value = configData(1)
        ReferenceLineLengthBox.Value = configData(2)
        ReferenceLineOffsetBox.Value = configData(3)
    End If
    
    ' 取り消し線作図設定値読み込み
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.STRIKETHROUGH_CONFIG), vbCrLf)
    If UBound(configData) = 4 Then
        StrikethroughLayerOnBox.Value = _
            IIf(configData(0) = "True", "True", "False")
        StrikethroughLayerBox.Value = configData(1)
        StrikethroughLengthBox.Value = configData(2)
        StrikethroughRedBox.Value = _
            IIf(configData(3) = "True", "True", "False")
        TargetEntityLayerBox.Value = _
            IIf(configData(4) = "True", "True", "False")
    End If
    
    ' 参照線作図画層の指定切り替え
    If ReferenceLineLayerOnBox.Value Then
        ReferenceLineLayerBox.Enabled = True
    Else
        ReferenceLineLayerBox.Enabled = False
    End If
    
    ' 取り消し線作図画層の指定切り替え
    If StrikethroughLayerOnBox.Value Then
        StrikethroughLayerBox.Enabled = True
    Else
        StrikethroughLayerBox.Enabled = False
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## 参照線作図画層の指定切り替え
'------------------------------------------------------------------------------
Private Sub ReferenceLineLayerOnBox_Change()
    
    If ReferenceLineLayerOnBox.Value Then
        ReferenceLineLayerBox.Enabled = True
    Else
        ReferenceLineLayerBox.Enabled = False
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## 取り消し線作図画層の指定切り替え
'------------------------------------------------------------------------------
Private Sub StrikethroughLayerOnBox_Change()
    
    If StrikethroughLayerOnBox.Value Then
        StrikethroughLayerBox.Enabled = True
    Else
        StrikethroughLayerBox.Enabled = False
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## 設定値保存
'------------------------------------------------------------------------------
Private Sub DecorationSaveButton_Click()
    
    Dim configData As Variant
    
    ' 設定値入力の確認
    If Not validateDecorationConfign() Then Exit Sub
    
    ' 設定フォルダの準備
    Call CommitConfig.PrepareConfigFolder
    
    ' 参照線作図設定値保存
    configData = ReferenceLineLayerOnBox.Value & vbCrLf _
               & ReferenceLineLayerBox.Value & vbCrLf _
               & ReferenceLineLengthBox.Value & vbCrLf _
               & ReferenceLineOffsetBox.Value
    
    Call CommitConfig.SaveConfig _
        (FormDisplay.REFERENCELINE_CONFIG, configData)
    
    ' 取消線作図設定値保存
    configData = StrikethroughLayerOnBox.Value & vbCrLf _
               & StrikethroughLayerBox.Value & vbCrLf _
               & StrikethroughLengthBox.Value & vbCrLf _
               & StrikethroughRedBox.Value & vbCrLf _
               & TargetEntityLayerBox.Value
    
    Call CommitConfig.SaveConfig _
        (FormDisplay.STRIKETHROUGH_CONFIG, configData)
    
End Sub

'------------------------------------------------------------------------------
' ## 設定値入力の確認
'------------------------------------------------------------------------------
Private Function validateDecorationConfign() As Boolean
    
    validateDecorationConfign = False
    
    ' 画層リスト取得
    Dim i As Long
    Dim layerList() As Variant
    For i = 0 To ThisDrawing.Layers.Count - 1
        ReDim Preserve layerList(i)
        layerList(i) = ThisDrawing.Layers.Item(i).Name
    Next i
    
    ' 参照線作図画層入力確認
    If ReferenceLineLayerOnBox.Value _
    And Not CommonFunction.IsMatchList _
        (layerList, ReferenceLineLayerBox.Value) Then
        MsgBox "参照線作図画層の入力が不正です。", vbCritical
        Exit Function
    End If
    
    ' 参照線線長係数入力確認
    If Not IsNumeric(ReferenceLineLengthBox.Value) Then
        MsgBox "参照線線長係数の入力が不正です。", vbCritical
        Exit Function
    End If
    
    ' 参照線オフセット係数入力確認
    If Not IsNumeric(ReferenceLineOffsetBox.Value) Then
        MsgBox "参照線オフセット係数の入力が不正です。", vbCritical
        Exit Function
    End If
    
    ' 取消線作図画層入力確認
    If StrikethroughLayerOnBox.Value _
    And Not CommonFunction.IsMatchList _
        (layerList, StrikethroughLayerBox.Value) Then
        MsgBox "取り消し線作図画層の入力が不正です。", vbCritical
        Exit Function
    End If
    
    ' 参照線線長係数入力確認
    If Not IsNumeric(StrikethroughLengthBox.Value) Then
        MsgBox "取り消し線線長係数の入力が不正です。", vbCritical
        Exit Function
    End If
    
    validateDecorationConfign = True
    
End Function
