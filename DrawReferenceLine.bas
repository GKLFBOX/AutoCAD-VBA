Attribute VB_Name = "DrawReferenceLine"
Option Explicit

'------------------------------------------------------------------------------
' ## 文字系オブジェクトへの参照線作図
'
' 指定した文字とオフセット係数から参照線を作図する
'------------------------------------------------------------------------------
Public Sub DrawReferenceLine()
    
    On Error GoTo Error_Handler
    
    Dim pickPoint As Variant
    Dim targetEntity As ZcadEntity
    Dim configLayer As String
    Dim configLength As Single
    Dim configOffset As Single
    
    ' 位置調整する文字系オブジェクトの選択
    ThisDrawing.Utility.GetEntity targetEntity, pickPoint, _
        "位置調整する文字またはブロック内文字を選択 [Cancel(ESC)]"
    
    Dim configData As Variant
    configData = Split(CommitConfig.LoadConfig _
        (FormDisplay.REFERENCELINE_CONFIG), vbCrLf)
    If UBound(configData) = 2 Then
        configLayer = configData(0)
        configLength = configData(1) / 2
        configOffset = configData(2)
    End If
    
    ' テキストまたはブロック参照の判定
    If CommonFunction.IsTextObject(targetEntity) Then
        Call addTextReferenceLine _
            (targetEntity, pickPoint, configLayer, configLength, configOffset)
        Exit Sub
    ElseIf TypeOf targetEntity Is ZcadBlockReference Then
        
        Exit Sub
    Else
        ThisDrawing.Utility.Prompt _
            "文字またはブロック内文字が選択されませんでした。" & vbCrLf
        Exit Sub
    End If
    
    Call CommonSub.ResetHighlight(targetEntity)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ResetHighlight(targetEntity)
    ThisDrawing.Utility.Prompt "なんらかのエラーです。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## 文字への参照線作図
'------------------------------------------------------------------------------
Private Sub addTextReferenceLine(ByRef target_text As ZcadEntity, _
                                 ByVal pick_point As Variant, _
                                 ByVal config_layer As String, _
                                 ByVal config_length As Single, _
                                 ByVal config_offset As Single)
    
    Dim targetAngle As Double
    Dim startPoint(0 To 2) As Double
    Dim endPoint(0 To 2) As Double
    
    ' 作図簡略化のために基点と角度を記憶し角度要素削除
    targetAngle = target_text.Rotation
    target_text.Rotate pick_point, targetAngle * -1
    
    Call calculateEdgePoints _
        (target_text, startPoint, endPoint, config_length, config_offset)
    
    ' 参照線作図
    Dim referenceLine As ZcadLine
    Set referenceLine = ThisDrawing.ModelSpace.AddLine(startPoint, endPoint)
    
    ' 文字および取り消し線の角度を戻す
    target_text.Rotate pick_point, targetAngle
    referenceLine.Rotate pick_point, targetAngle
    
    ' 画層適用
    referenceLine.Layer = config_layer
    
End Sub

'------------------------------------------------------------------------------
' ## 参照線始終端計算
'------------------------------------------------------------------------------
Private Sub calculateEdgePoints(ByVal target_text As ZcadEntity, _
                                ByRef start_point() As Double, _
                                ByRef end_point() As Double, _
                                ByVal config_length As Single, _
                                ByVal config_offset As Single)
    
    Dim minExtent As Variant, maxExtent As Variant
    
    target_text.GetBoundingBox minExtent, maxExtent
    
    ' ZWCAD2020ではGetBondingBoxが文字の傾斜角度を無視してしまうため
    ' 斜体文字を考慮し傾斜角度からMaxPointを最適化する
    Dim textOblique As Double
    Dim deltaX As Double
    Dim deltaY As Double
    textOblique = target_text.ObliqueAngle
    deltaY = maxExtent(1) - minExtent(1)
    deltaX = deltaY * Tan(textOblique)
    If textOblique > 0 Then
        maxExtent(0) = maxExtent(0) + deltaX
    ElseIf textOblique < 0 Then
        minExtent(0) = minExtent(0) - deltaX
    End If
    
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
