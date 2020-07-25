Attribute VB_Name = "CreateLayout"
Option Explicit

'------------------------------------------------------------------------------
' ## 枠ブロック選択によるレイアウト作成   2020/07/25 G.O.
'
' 用紙枠ブロックからレイアウトを生成する
' 枠外にデータ上での識別用マークを持たせていることを考慮して
' 用紙枠サイズは属性定義を除いたサイズを取得する
'------------------------------------------------------------------------------
Public Sub CreateLayout(ByVal frame_tag As String, _
                        ByVal scale_factor As Single, _
                        ByVal viewport_layer As String, _
                        ByVal style_name As String, _
                        ByVal printer_name As String, _
                        ByVal a3_paper As String, _
                        ByVal a4_paper As String, _
                        ByVal offset_x As Single, _
                        ByVal offset_y As Single)
    
    On Error GoTo Error_Handler
    
    ' 枠の幅および高さは製図誤差を考慮してSingleとしている
    Dim paperFrame As ZcadBlockReference
    Dim minFramePoint As Variant, maxFramePoint As Variant
    Dim frameWidth As Single, frameHeight As Single
    Dim newLayout As ZcadLayout
    
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts("Model")
    
    ' 枠ブロックの選択
    Call CommonSub.PickFrameBlock(paperFrame)
    If paperFrame Is Nothing Then Exit Sub
    
    ' 枠名称取得およびレイアウト作成
    Dim frameName As String
    Call CommonSub.FetchFrameName(paperFrame, frame_tag, frameName)
    Set newLayout = ThisDrawing.Layouts.Add(frameName)
    
    ' 用紙枠サイズ取得
    Call CommonSub.FetchCorrectSize(paperFrame, minFramePoint, maxFramePoint)
    frameWidth = maxFramePoint(0) - minFramePoint(0)
    frameHeight = maxFramePoint(1) - minFramePoint(1)
    
    ' 新規レイアウトアクティブ化
    ThisDrawing.ActiveLayout = newLayout
    
    ' ビューポート調整
    ' 新規PaperSpaceのItem(0)はレイアウトの画面そのものであり
    ' ユーザーが認識しているビューポートオブジェクトはItem(1)のため注意
    Dim paperViewport As ZcadPViewport
    Set paperViewport = ThisDrawing.PaperSpace.Item(1)
    Call adjustViewportShape(paperViewport, frameWidth, frameHeight)
    Call CommonSub.ApplyViewportProperty _
        (paperViewport, viewport_layer, minFramePoint, maxFramePoint)
    
    ' 印刷設定
    Call configurePrintSettings(newLayout, frameWidth, frameHeight, _
        scale_factor, style_name, printer_name, a3_paper, a4_paper, _
        offset_x, offset_y)
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "なんらかのエラーです。"
    
End Sub

'------------------------------------------------------------------------------
' ## ビューポートの位置およびサイズ調整
'------------------------------------------------------------------------------
Private Sub adjustViewportShape(ByRef paper_viewport As ZcadPViewport, _
                                ByVal frame_width As Single, _
                                ByVal frame_height As Single)
    
    Dim frameCenter(0 To 2) As Double
    
    ' 用紙枠中心算出
    frameCenter(0) = frame_width / 2
    frameCenter(1) = frame_height / 2
    frameCenter(2) = 0
    
    ' 位置及びサイズ調整
    With paper_viewport
        .Center = frameCenter
        .Width = frame_width
        .Height = frame_height
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## レイアウトの印刷設定
'------------------------------------------------------------------------------
Private Sub configurePrintSettings(ByRef new_layout As ZcadLayout, _
                                   ByVal frame_width As Single, _
                                   ByVal frame_height As Single, _
                                   ByVal scale_factor As Variant, _
                                   ByVal style_name As Variant, _
                                   ByVal printer_name As Variant, _
                                   ByVal a3_paper As Variant, _
                                   ByVal a4_paper As Variant, _
                                   ByVal offset_x As Variant, _
                                   ByVal offset_y As Variant)
    
    Dim plotOffset(0 To 1) As Double
    
    With new_layout
        
        ' 印刷領域
        .PlotType = zcLayout
        
        ' 尺度
        .UseStandardScale = False
        .SetCustomScale 1, scale_factor
        
        ' 印刷スタイル
        .StyleSheet = style_name
        
        ' プリンタ名称
        .ConfigName = printer_name
        
        ' 印刷オフセット(XYが一般と逆のため注意)
        plotOffset(0) = offset_y: plotOffset(1) = offset_x
        .PlotOrigin = plotOffset
        
        ' 用紙サイズおよび方向
        frame_width = frame_width / scale_factor
        frame_height = frame_height / scale_factor
        If frame_width = 420 And frame_height = 297 Then
            .CanonicalMediaName = a3_paper
            .PlotRotation = zc90degrees
        ElseIf frame_width = 297 And frame_height = 420 Then
            .CanonicalMediaName = a3_paper
            .PlotRotation = zc0degrees
        ElseIf frame_width = 210 And frame_height = 297 Then
            .CanonicalMediaName = a4_paper
            .PlotRotation = zc0degrees
        Else
            MsgBox "用紙サイズが検出できなかったため手動で設定して下さい。"
        End If
        
    End With
    
    ' 画面上の反映を行うためレイアウトタブの切り替え
    ' ZWCADの不具合でRegenでは反映されないことに対する対策
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts("Model")
    ThisDrawing.ActiveLayout = new_layout
    
End Sub
