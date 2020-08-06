Attribute VB_Name = "CommonSub"
Option Explicit

'------------------------------------------------------------------------------
' ## ハイライトの解除
'------------------------------------------------------------------------------
Public Sub ResetHighlight(ByVal target_object As ZcadEntity)

    If Not target_object Is Nothing Then target_object.Highlight False

End Sub

'------------------------------------------------------------------------------
' ## 選択セットの削除
'------------------------------------------------------------------------------
Public Sub ReleaseSelectionSet(ByVal target_selectionset As ZcadSelectionSet)

    If Not target_selectionset Is Nothing Then target_selectionset.Delete

End Sub

'------------------------------------------------------------------------------
' ## 枠ブロックの選択
'------------------------------------------------------------------------------
Public Sub PickFrameBlock(ByRef frame_block As ZcadBlockReference)
    
    Dim targetEntity As ZcadEntity
    Dim pickPoint As Variant
    
    ThisDrawing.Utility.GetEntity targetEntity, pickPoint, _
        "枠ブロックを選択 [Cancel(ESC)]"
        
    If TypeOf targetEntity Is ZcadBlockReference Then
        Set frame_block = targetEntity
    Else
        ThisDrawing.Utility.Prompt "ブロック以外が選択されました。"
        Exit Sub
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## 枠名称を取得
'------------------------------------------------------------------------------
Public Sub FetchFrameName(ByVal frame_block As ZcadBlockReference, _
                          ByVal frame_tag As String, _
                          ByRef frame_name As String)
    
    Dim frameAttributes As Variant
    Dim currentAttribute As ZcadAttributeReference
    
    frameAttributes = frame_block.GetAttributes
    
    ' 指定属性の検索
    Dim i As Long
    For i = 0 To UBound(frameAttributes)
        Set currentAttribute = frameAttributes(i)
        If currentAttribute.TagString = frame_tag Then
            frame_name = currentAttribute.TextString
            Exit Sub
        End If
    Next i
    
    ' 指定属性が無かった場合はズームして直接入力を促す
    Dim minExtent As Variant, maxExtent As Variant
    If frame_name = "" Then
        frame_block.GetBoundingBox minExtent, maxExtent
        ThisDrawing.Application.ZoomWindow minExtent, maxExtent
        frame_name = ThisDrawing.Utility.GetString _
            (0, "用紙枠名称が見つからないため直接入力 [Cancel(ESC)]:")
        Exit Sub
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## 属性定義を除いた正確な用紙枠サイズ取得
'------------------------------------------------------------------------------
Public Sub FetchCorrectSize(ByVal frame_block As ZcadBlockReference, _
                            ByRef min_framepoint As Variant, _
                            ByRef max_framepoint As Variant)
    
    Dim i As Long, j As Long
    Dim replicaEntities As Variant
    Dim extractEntity As ZcadEntity
    Dim currentMin As Variant, currentMax As Variant
    
    replicaEntities = frame_block.Explode
    
    ' 属性定義を除くブロック構成要素から用紙枠サイズを取得
    For i = 0 To UBound(replicaEntities)
        
        Set extractEntity = replicaEntities(i)
        If TypeOf extractEntity Is ZcadAttribute Then GoTo Continue_i
        
        ' 比較更新を行い最外周サイズを取得する
        If CommonFunction.IsEmptyArray(min_framepoint) _
        And CommonFunction.IsEmptyArray(max_framepoint) Then
            extractEntity.GetBoundingBox min_framepoint, max_framepoint
        Else
            extractEntity.GetBoundingBox currentMin, currentMax
            For j = 0 To 1
                If currentMin(j) <= min_framepoint(j) Then
                    min_framepoint(j) = currentMin(j)
                End If
                If currentMax(j) >= max_framepoint(j) Then
                    max_framepoint(j) = currentMax(j)
                End If
            Next j
        End If
        
Continue_i:
        extractEntity.Delete
        
    Next i
    
End Sub

'------------------------------------------------------------------------------
' ## ビューポートの表示調整
'------------------------------------------------------------------------------
Public Sub ApplyViewportProperty(ByRef target_viewport As ZcadPViewport, _
                                 ByVal target_layer As String, _
                                 ByVal min_framepoint As Variant, _
                                 ByVal max_framepoint As Variant)
    
    Dim changeColor As ZcadZcCmColor
    
    Set changeColor = New ZcadZcCmColor
    changeColor.ColorIndex = zcByLayer
    
    ' プロパティ設定
    With target_viewport
        
        .Layer = target_layer
        .TrueColor = changeColor
        .Linetype = "ByLayer"
        .LinetypeScale = 1
        .Lineweight = zcLnWtByLayer
        
    End With
    
    ' ビューポート内の表示調整
    With ThisDrawing
        
        ' ここでペーパー空間の全体表示をしないと
        ' なぜかビューポート内の表示調整で尺度が若干ズレてしまう
        .Application.ZoomExtents
        
        .MSpace = True
        
        .ActivePViewport = target_viewport
        .Application.ZoomWindow min_framepoint, max_framepoint
        
        .MSpace = False
        
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## 分解オブジェクトの属性定義名称置換
'------------------------------------------------------------------------------
Public Sub ReplaceAttributeTag(ByRef target_block As ZcadBlockReference, _
                               ByVal replica_entities As Variant)
    
    Dim targetAttributes As Variant
    
    ' 属性取得および有無確認
    targetAttributes = target_block.GetAttributes
    If CommonFunction.IsEmptyArray(targetAttributes) Then Exit Sub
    
    ' 分解オブジェクトの属性定義を検索
    Dim i As Long, j As Long
    Dim currentReplica As ZcadEntity
    Dim currentAttribute As ZcadAttributeReference
    For i = 0 To UBound(replica_entities)
        
        Set currentReplica = replica_entities(i)
        If Not TypeOf currentReplica Is ZcadAttribute Then _
            GoTo Continue_i
        
        ' 画面表示上の属性定義名称をブロックの対応する属性値に置換
        For j = 0 To UBound(targetAttributes)
            Set currentAttribute = targetAttributes(j)
            If currentAttribute.TagString = currentReplica.TagString Then
                currentReplica.TagString = currentAttribute.TextString
                Exit For
            End If
        Next j
        
Continue_i:
    
    Next i
    
End Sub

'------------------------------------------------------------------------------
' ## 指定点の分解オブジェクト取得
'------------------------------------------------------------------------------
Public Sub GrabReplicaEntity(ByVal pick_point As Variant, _
                             ByRef target_replica As ZcadEntity)
    
    Dim replicaSelectionSet As ZcadSelectionSet
    
    Set replicaSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    
    replicaSelectionSet.SelectAtPoint pick_point
    Set target_replica = replicaSelectionSet.Item(0)
    
    Call CommonSub.ReleaseSelectionSet(replicaSelectionSet)
    
End Sub

'------------------------------------------------------------------------------
' ## 分解オブジェクトの非表示
'------------------------------------------------------------------------------
Public Sub HideReplica(ByVal replica_entities As Variant)
    
    Dim i As Long
    Dim targetReplica As ZcadEntity
    
    For i = 0 To UBound(replica_entities)
        Set targetReplica = replica_entities(i)
        targetReplica.Visible = False
    Next i
    
End Sub

'------------------------------------------------------------------------------
' ## 分解オブジェクトの削除
'------------------------------------------------------------------------------
Public Sub DeleteReplica(ByVal replica_entities As Variant)
    
    Dim i As Long
    Dim targetReplica As ZcadEntity
    
    For i = 0 To UBound(replica_entities)
        Set targetReplica = replica_entities(i)
        targetReplica.Delete
    Next i
    
End Sub

'------------------------------------------------------------------------------
' ## 拡張版(ZWCAD2020斜体文字対応版)GetBoundingBox
'------------------------------------------------------------------------------
Public Sub GetEnhancedBoundingBox(ByVal target_text As ZcadEntity, _
                                  ByRef min_extent As Variant, _
                                  ByRef max_extent As Variant)
    
    target_text.GetBoundingBox min_extent, max_extent
    
    ' ZWCAD2020ではGetBondingBoxが文字の傾斜角度を無視してしまうため
    ' 斜体の文字を考慮し傾斜角度からMinPointまたはMaxPointを最適化する
    Dim textOblique As Double
    Dim deltaX As Double, deltaY As Double
    textOblique = target_text.ObliqueAngle
    deltaY = max_extent(1) - min_extent(1)
    deltaX = deltaY * Tan(textOblique)
    If textOblique > 0 Then
        max_extent(0) = max_extent(0) + deltaX
    ElseIf textOblique < 0 Then
        min_extent(0) = min_extent(0) - deltaX
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## csvファイルへの出力
'------------------------------------------------------------------------------
Public Sub OutputCSV(ByVal output_file As String, ByVal output_data As String)
    
    Open output_file For Append As #1
    Print #1, output_data
    Close #1
    
End Sub
