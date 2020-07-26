Attribute VB_Name = "AddViewport"
Option Explicit

'------------------------------------------------------------------------------
' ## 枠ブロック選択によるビューポート追加   2020/07/26 G.O.
'
' レイアウト枠ブロックから対象レイアウトにビューポートを追加する
'------------------------------------------------------------------------------
Public Sub AddViewport(ByVal frame_tag As String, _
                       ByVal scale_factor As Single, _
                       ByVal viewport_layer As String, _
                       ByVal custom_scale As Single)
    
    On Error GoTo Error_Handler
    
    Dim LayoutFrame As ZcadBlockReference
    Dim minFramePoint As Variant, maxFramePoint As Variant
    Dim frameWidth As Double, frameHeight As Double
    Dim targetLayout As ZcadLayout
    
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts("Model")
    
    ' 枠ブロックの選択
    Call CommonSub.PickFrameBlock(LayoutFrame)
    If LayoutFrame Is Nothing Then Exit Sub
    
    ' 枠名称取得および対象レイアウトの取得
    Dim frameName As String
    Call CommonSub.FetchFrameName(LayoutFrame, frame_tag, frameName)
    Call fetchTargetLayout(frameName, targetLayout)
    If targetLayout Is Nothing Then Exit Sub
    
    ' 用紙枠サイズ取得
    Dim customScale As Single
    customScale = scale_factor / custom_scale
    Call CommonSub.FetchCorrectSize(LayoutFrame, minFramePoint, maxFramePoint)
    frameWidth = (maxFramePoint(0) - minFramePoint(0)) * customScale
    frameHeight = (maxFramePoint(1) - minFramePoint(1)) * customScale
    
    ' 対象レイアウトアクティブ化
    ThisDrawing.ActiveLayout = targetLayout
    
    ' ビューポート挿入および調整
    Dim layoutViewport As ZcadPViewport
    Call insertViewport(frameWidth, frameHeight, layoutViewport)
    Call CommonSub.ApplyViewportProperty _
        (layoutViewport, viewport_layer, minFramePoint, maxFramePoint)
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "なんらかのエラーです。"
    
End Sub

'------------------------------------------------------------------------------
' ## 対象レイアウトの取得
'------------------------------------------------------------------------------
Private Sub fetchTargetLayout(ByVal frame_name As String, _
                              ByRef target_layout As ZcadLayout)
    
    Dim currentLayout As ZcadLayout
    For Each currentLayout In ThisDrawing.Layouts
        If currentLayout.Name = frame_name Then
            Set target_layout = currentLayout
            Exit Sub
        End If
    Next currentLayout
    
    ThisDrawing.Utility.Prompt "対象レイアウトが存在しません。"
    
End Sub

'------------------------------------------------------------------------------
' ## ビューポート挿入
'------------------------------------------------------------------------------
Private Sub insertViewport(ByVal frame_width As Double, _
                           ByVal frame_height As Double, _
                           ByRef layout_viewport As ZcadPViewport)
    
    Dim targetPoint As Variant
    Dim viewCenter(0 To 2) As Double
    
    targetPoint = ThisDrawing.Utility.GetPoint _
        (, "挿入点(左下)を指定 [Cancel(ESC)]:")
    
    viewCenter(0) = targetPoint(0) + frame_width / 2
    viewCenter(1) = targetPoint(1) + frame_height / 2
    viewCenter(2) = 0
    
    Set layout_viewport = ThisDrawing.PaperSpace.AddPViewport _
        (viewCenter, frame_width, frame_height)
    
End Sub
