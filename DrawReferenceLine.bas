Attribute VB_Name = "DrawReferenceLine"
Option Explicit

'------------------------------------------------------------------------------
' ## 文字系オブジェクトへの参照線作図   2020/08/03 G.O.
'
' 指定した文字とオフセット係数から参照線を作図する
'------------------------------------------------------------------------------
Public Sub DrawReferenceLine()
    
    On Error GoTo Error_Handler
    
    Dim targetEntity As ZcadEntity
    Dim pickPoint As Variant
    Dim configLayerOn As Boolean
    Dim configLayer As String
    Dim configLength As Single
    Dim configOffset As Single
    
    ' 設定値読み込み
    Dim configData As Variant
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.REFERENCELINE_CONFIG), vbCrLf)
    configLayerOn = configData(0)
    configLayer = configData(1)
    configLength = configData(2) / 2
    configOffset = configData(3)
    
    ' 対象文字系オブジェクトの選択
    ThisDrawing.Utility.GetEntity targetEntity, pickPoint, _
        "参照線を引く文字またはブロック内文字を選択 [Cancel(ESC)]"
    
    ' テキストまたはブロック参照の判定
    If CommonFunction.IsTextObject(targetEntity) Then
        Call addTextReferenceLine _
            (targetEntity, pickPoint, configLayerOn, _
            configLayer, configLength, configOffset)
    ElseIf TypeOf targetEntity Is ZcadBlockReference Then
        Call addBlockReferenceLine _
            (targetEntity, pickPoint, configLayerOn, _
            configLayer, configLength, configOffset)
    Else
        ThisDrawing.Utility.Prompt _
            "文字またはブロック内文字が選択されませんでした。" & vbCrLf
    End If
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "なんらかのエラーです。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## 文字への参照線作図
'------------------------------------------------------------------------------
Private Sub addTextReferenceLine(ByRef target_text As ZcadEntity, _
                                 ByVal pick_point As Variant, _
                                 ByVal config_layeron As Boolean, _
                                 ByVal config_layer As String, _
                                 ByVal config_length As Single, _
                                 ByVal config_offset As Single)
    
    Dim targetAngle As Double
    Dim startPoint(0 To 2) As Double
    Dim endPoint(0 To 2) As Double
    Dim referenceLine As ZcadLine
    
    ' 作図簡略化のために基点と角度を記憶し角度要素削除
    targetAngle = target_text.Rotation
    target_text.Rotate pick_point, targetAngle * -1
    
    ' 参照線始終端算出
    Call getReferenceLineEdge _
        (target_text, startPoint, endPoint, config_length, config_offset)
    
    ' 参照線作図
    Set referenceLine = ThisDrawing.ModelSpace.AddLine(startPoint, endPoint)
    
    ' 文字および参照線の角度を戻す
    target_text.Rotate pick_point, targetAngle
    referenceLine.Rotate pick_point, targetAngle
    
    ' 画層適用
    If config_layeron Then referenceLine.Layer = config_layer
    
End Sub

'------------------------------------------------------------------------------
' ## ブロックへの参照線作図
'------------------------------------------------------------------------------
Private Sub addBlockReferenceLine(ByRef target_block As ZcadBlockReference, _
                                  ByVal pick_point As Variant, _
                                  ByVal config_layeron As Boolean, _
                                  ByVal config_layer As String, _
                                  ByVal config_length As Single, _
                                  ByVal config_offset As Single)
    
    On Error GoTo Error_Handler
    
    Dim replicaEntities As Variant
    Dim targetReplica As ZcadEntity
    Dim targetAngle As Double
    Dim startPoint(0 To 2) As Double, endPoint(0 To 2) As Double
    Dim referenceLine As ZcadLine
    
    replicaEntities = target_block.Explode
    
    ' 分解オブジェクトの属性定義名称置換
    Call CommonSub.ReplaceAttributeTag(target_block, replicaEntities)
    
    ' 指定点の分解オブジェクトを取得
    Call CommonSub.GrabReplicaEntity(pick_point, targetReplica)
    
    ' 画面上では分解オブジェクトを非表示化
    Call CommonSub.HideReplica(replicaEntities)
    
    ' テキスト内文字の判定
    If Not CommonFunction.IsTextObject(targetReplica) _
    And Not TypeOf targetReplica Is ZcadAttribute Then
        Call CommonSub.DeleteReplica(replicaEntities)
        ThisDrawing.Utility.Prompt _
            "ブロック内文字が選択されませんでした。" & vbCrLf
        Exit Sub
    End If
    
    ' 作図簡略化のために基点と角度を記憶し角度要素削除
    targetAngle = targetReplica.Rotation
    targetReplica.Rotate pick_point, targetAngle * -1
    target_block.Rotate pick_point, targetAngle * -1
    
    ' 参照線始終端算出
    Call getReferenceLineEdge _
        (targetReplica, startPoint, endPoint, config_length, config_offset)
    
    ' 参照線作図
    Set referenceLine = ThisDrawing.ModelSpace.AddLine(startPoint, endPoint)
    
    ' 文字および参照線の角度を戻す
    target_block.Rotate pick_point, targetAngle
    referenceLine.Rotate pick_point, targetAngle
    
    ' 画層適用
    If config_layeron Then referenceLine.Layer = config_layer
    
    Call CommonSub.DeleteReplica(replicaEntities)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.DeleteReplica(replicaEntities)
    
End Sub

'------------------------------------------------------------------------------
' ## 参照線始終端計算
'------------------------------------------------------------------------------
Private Sub getReferenceLineEdge(ByVal target_text As ZcadEntity, _
                                 ByRef start_point() As Double, _
                                 ByRef end_point() As Double, _
                                 ByVal config_length As Single, _
                                 ByVal config_offset As Single)
    
    Dim minExtent As Variant, maxExtent As Variant
    
    ' 拡張版GetBoundingBox
    Call CommonSub.GetEnhancedBoundingBox(target_text, minExtent, maxExtent)
    
    ' 始端計算
    start_point(0) = minExtent(0) _
        - ((maxExtent(1) - minExtent(1)) * config_length)
    start_point(1) = minExtent(1) - target_text.Height * config_offset
    start_point(2) = 0
    
    ' 終端計算
    end_point(0) = maxExtent(0) _
        + ((maxExtent(1) - minExtent(1)) * config_length)
    end_point(1) = minExtent(1) - target_text.Height * config_offset
    end_point(2) = 0
    
End Sub
