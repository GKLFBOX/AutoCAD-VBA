Attribute VB_Name = "DrawDoubleStrikethrough"
Option Explicit

'------------------------------------------------------------------------------
' ## 文字系オブジェクトへの二重取り消し線作図   2020/08/03 G.O.
'
' 指定した文字オブジェクトに同画層で二重取り消し線を作図する
'------------------------------------------------------------------------------
Public Sub DrawDoubleStrikethrough()
    
    On Error GoTo Error_Handler
    
    Dim targetEntity As ZcadEntity
    Dim pickPoint As Variant
    Dim configLayerOn As Boolean
    Dim configLayer As String
    Dim configLength As Single
    Dim configRed As Boolean
    Dim configTargetLayer As Boolean
    
    ' 設定値読み込み
    Dim configData As Variant
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.STRIKETHROUGH_CONFIG), vbCrLf)
    configLayerOn = configData(0)
    configLayer = configData(1)
    configLength = configData(2) / 2
    configRed = configData(3)
    configTargetLayer = configData(4)
    
    ' 対象文字系オブジェクトの選択
    ThisDrawing.Utility.GetEntity targetEntity, pickPoint, _
        "二重取り消し線を引く文字またはブロック内文字を選択 [Cancel(ESC)]"
    
    ' テキストまたはブロック参照の判定
    If CommonFunction.IsTextObject(targetEntity) Then
        Call addTextStrikethrough _
            (targetEntity, pickPoint, configLayerOn, _
            configLayer, configLength, configRed, configTargetLayer)
    ElseIf TypeOf targetEntity Is ZcadBlockReference Then
        Call addBlockStrikethrough _
            (targetEntity, pickPoint, configLayerOn, _
            configLayer, configLength, configRed, configTargetLayer)
    Else
        ThisDrawing.Utility.Prompt _
            "文字またはブロック内文字が選択されませんでした。" & vbCrLf
    End If
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "なんらかのエラーです。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## 文字への二重取り消し線作図
'------------------------------------------------------------------------------
Private Sub addTextStrikethrough(ByRef target_text As ZcadEntity, _
                                 ByVal pick_point As Variant, _
                                 ByVal config_layeron As Boolean, _
                                 ByVal config_layer As String, _
                                 ByVal config_length As Single, _
                                 ByVal config_red As Boolean, _
                                 ByVal config_targetlayer As Boolean)
    
    Dim targetAngle As Double
    Dim startPoint1(0 To 2) As Double, endPoint1(0 To 2) As Double
    Dim startPoint2(0 To 2) As Double, endPoint2(0 To 2) As Double
    Dim strikeThrough1 As ZcadLine, strikeThrough2 As ZcadLine
    
    ' 作図簡略化のために基点と角度を記憶し角度要素削除
    targetAngle = target_text.Rotation
    target_text.Rotate pick_point, targetAngle * -1
    
    ' 取り消し線始終端算出
    Call getStrikethroughEdge(target_text, startPoint1, endPoint1, _
        startPoint2, endPoint2, config_length)
    
    ' 取り消し線作図
    Set strikeThrough1 = ThisDrawing.ModelSpace.AddLine(startPoint1, endPoint1)
    Set strikeThrough2 = ThisDrawing.ModelSpace.AddLine(startPoint2, endPoint2)
    
    ' 文字および取り消し線の角度を戻す
    target_text.Rotate pick_point, targetAngle
    strikeThrough1.Rotate pick_point, targetAngle
    strikeThrough2.Rotate pick_point, targetAngle
    
    ' 作図設定の適用
    Call applyDrawingConfig(target_text, strikeThrough1, strikeThrough2, _
        config_layeron, config_layer, config_red, config_targetlayer)
    
End Sub

'------------------------------------------------------------------------------
' ## ブロックへの二重取り消し線作図
'------------------------------------------------------------------------------
Private Sub addBlockStrikethrough(ByRef target_block As ZcadBlockReference, _
                                  ByVal pick_point As Variant, _
                                  ByVal config_layeron As Boolean, _
                                  ByVal config_layer As String, _
                                  ByVal config_length As Single, _
                                  ByVal config_red As Boolean, _
                                  ByVal config_targetlayer As Boolean)
    
    On Error GoTo Error_Handler
    
    Dim replicaEntities As Variant
    Dim targetReplica As ZcadEntity
    Dim targetAngle As Double
    Dim startPoint1(0 To 2) As Double, endPoint1(0 To 2) As Double
    Dim startPoint2(0 To 2) As Double, endPoint2(0 To 2) As Double
    Dim strikeThrough1 As ZcadLine, strikeThrough2 As ZcadLine
    
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
    
    ' 取り消し線始終端算出
    Call getStrikethroughEdge(targetReplica, startPoint1, endPoint1, _
        startPoint2, endPoint2, config_length)
    
    ' 取り消し線作図
    Set strikeThrough1 = ThisDrawing.ModelSpace.AddLine(startPoint1, endPoint1)
    Set strikeThrough2 = ThisDrawing.ModelSpace.AddLine(startPoint2, endPoint2)
    
    ' 文字および取り消し線の角度を戻す
    target_block.Rotate pick_point, targetAngle
    strikeThrough1.Rotate pick_point, targetAngle
    strikeThrough2.Rotate pick_point, targetAngle
    
    ' 作図設定の適用
    Call applyDrawingConfig(target_block, strikeThrough1, strikeThrough2, _
        config_layeron, config_layer, config_red, config_targetlayer)
    
    Call CommonSub.DeleteReplica(replicaEntities)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.DeleteReplica(replicaEntities)
    
End Sub

'------------------------------------------------------------------------------
' ## 参照線始終端計算
'------------------------------------------------------------------------------
Private Sub getStrikethroughEdge(ByVal target_text As ZcadEntity, _
                                 ByRef start_point1() As Double, _
                                 ByRef end_point1() As Double, _
                                 ByRef start_point2() As Double, _
                                 ByRef end_point2() As Double, _
                                 ByVal config_length As Single)
    
    Dim minExtent As Variant, maxExtent As Variant
    
    ' 拡張版GetBoundingBox
    Call CommonSub.GetEnhancedBoundingBox(target_text, minExtent, maxExtent)
    
    ' 始端計算
    start_point1(0) = minExtent(0) _
        - ((maxExtent(1) - minExtent(1)) * config_length)
    start_point1(1) = minExtent(1) + ((maxExtent(1) - minExtent(1)) / 3)
    start_point1(2) = 0
    
    ' 終端計算
    end_point1(0) = maxExtent(0) _
        + ((maxExtent(1) - minExtent(1)) * config_length)
    end_point1(1) = minExtent(1) + ((maxExtent(1) - minExtent(1)) / 3)
    end_point1(2) = 0
    
    ' 始端計算2
    start_point2(0) = minExtent(0) _
        - ((maxExtent(1) - minExtent(1)) * config_length)
    start_point2(1) = minExtent(1) + ((maxExtent(1) - minExtent(1)) * 2 / 3)
    start_point2(2) = 0
    
    ' 終端計算2
    end_point2(0) = maxExtent(0) _
        + ((maxExtent(1) - minExtent(1)) * config_length)
    end_point2(1) = minExtent(1) + ((maxExtent(1) - minExtent(1)) * 2 / 3)
    end_point2(2) = 0
    
End Sub

'------------------------------------------------------------------------------
' ## 作図設定の適用
'------------------------------------------------------------------------------
Private Sub applyDrawingConfig(ByRef target_text As ZcadEntity, _
                               ByRef strike_through1 As ZcadLine, _
                               ByRef strike_through2 As ZcadLine, _
                               ByVal config_layeron As Boolean, _
                               ByVal config_layer As String, _
                               ByVal config_red As Boolean, _
                               ByVal config_targetlayer As Boolean)
    
    ' 作図画層の適用および取り消し対象画層変更の適用
    If config_layeron Then
        strike_through1.Layer = config_layer
        strike_through2.Layer = config_layer
        If config_targetlayer Then target_text.Layer = config_layer
    Else
        strike_through1.Layer = target_text.Layer
        strike_through2.Layer = target_text.Layer
    End If
    
    ' 取り消し線赤着色の適用
    Dim changeColor As ZcadZcCmColor
    If config_red Then
        Set changeColor = New ZcadZcCmColor
        changeColor.ColorIndex = zcRed
        strike_through1.TrueColor = changeColor
        strike_through2.TrueColor = changeColor
    End If
    
End Sub
