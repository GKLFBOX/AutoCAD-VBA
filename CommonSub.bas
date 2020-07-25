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
    
    Dim pickEntity As ZcadEntity
    Dim pickPoint As Variant
    
    ThisDrawing.Utility.GetEntity pickEntity, pickPoint, _
        "枠ブロックを選択 [Cancel(ESC)]"
        
    If TypeOf pickEntity Is ZcadBlockReference Then
        Set frame_block = pickEntity
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
        If IsEmpty(min_framepoint) And IsEmpty(max_framepoint) Then
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

