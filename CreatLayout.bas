Attribute VB_Name = "CreatLayout"
Option Explicit

'------------------------------------------------------------------------------
' ## 設定値化予定の定数郡
'------------------------------------------------------------------------------
Private Const PAPER_TAG As String = "CPN"
Private Const PRINTER_NAME As String = "DWG to PDF.pc5"
Private Const SCALE_FACTOR As Single = 100
Private Const PAPER_NAME_A3 As String = "ISO_A3_(297.00_x_420.00_MM)"
Private Const PAPER_NAME_A4 As String = "ISO_A4_(210.00_x_297.00_MM)"
Private Const STYLE_NAME As String = "Monochrome.ctb"
Private Const OFFSET_X As Double = "-17"
Private Const OFFSET_Y As Double = "-5"

'------------------------------------------------------------------------------
' ## 枠ブロック選択によるレイアウト生成
'
' 指定のマークを属性として持つ用紙枠ブロックからレイアウトを生成する
'------------------------------------------------------------------------------
Public Sub CreateLayout()
    
    'On Error GoTo Error_Handler
    
    Dim paperFrame As ZcadBlockReference
    Dim minFramePoint As Variant, maxFramePoint As Variant
    Dim frameWidth As Single, frameHeight As Single
    Dim frameName As String
    Dim paperViewport As ZcadPViewport
    
    Dim newLayout As zcadLayout
    
    ' 用紙枠ブロックの選択
    Call pickPaperFrame(paperFrame)
    If paperFrame Is Nothing Then Exit Sub
    
    ' 用紙枠サイズ取得
    paperFrame.GetBoundingBox minFramePoint, maxFramePoint
    frameWidth = maxFramePoint(0) - minFramePoint(0)
    frameHeight = maxFramePoint(1) - minFramePoint(1)
    
    ' 用紙枠名称取得
    Call fetchFrameName(paperFrame, frameName)
    If frameName = "" Then Exit Sub
    
    ' 新規レイアウトの作成およびアクティブ化
    Set newLayout = ThisDrawing.Layouts.Add(frameName)
    ThisDrawing.ActiveLayout = newLayout
    
    ' 新規PaperSpaceのItem(0)はレイアウトの画面そのものであり
    ' ユーザーが認識しているビューポートオブジェクトはItem(1)のため注意
    Set paperViewport = ThisDrawing.PaperSpace.Item(1)
    
    ' ビューポートの位置およびサイズ調整
    Call adjustViewportShape(paperViewport, frameWidth, frameHeight)
    
    ' 表示調整前に全体表示をしないとなぜかビューポート内全体表示がバグり
    ' 尺度が若干ズレてしまうためここで行っている(原因不明)
    ZoomExtents
    
    ' ビューポートの表示調整
    Call adjustViewportDisplay(paperViewport, minFramePoint, maxFramePoint)
    
    ' 印刷設定
    Call configurePrintSettings(newLayout, frameWidth, frameHeight)
    
    ' 画面上の反映を行うためレイアウトタブの切り替え
    ' ZWCADの不具合でRegenでは反映されないことに対する対策
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts("Model")
    ThisDrawing.ActiveLayout = newLayout
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.prompt "なんらかのエラーです。"
    
End Sub

'------------------------------------------------------------------------------
' ## 用紙枠ブロックの選択
'------------------------------------------------------------------------------
Private Sub pickPaperFrame(ByRef paper_frame As ZcadBlockReference)
    
    Dim pickEntity As ZcadEntity
    Dim pickPoint As Variant
    
    ThisDrawing.Utility.GetEntity pickEntity, pickPoint, _
        "用紙枠を選択 [Cancel(ESC)]"
        
    If TypeOf pickEntity Is ZcadBlockReference Then
        Set paper_frame = pickEntity
    Else
        ThisDrawing.Utility.prompt "ブロック以外が選択されました。"
        Exit Sub
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## 用紙枠名称を取得
'------------------------------------------------------------------------------
Private Sub fetchFrameName(ByVal paper_frame As ZcadBlockReference, _
                           ByRef frame_name As String)
    
    Dim frameAttributes As Variant
    Dim currentAttribute As ZcadAttributeReference
    
    frameAttributes = paper_frame.GetAttributes
    
    Dim i As Long
    For i = 0 To UBound(frameAttributes)
        Set currentAttribute = frameAttributes(i)
        If currentAttribute.TagString = PAPER_TAG Then
            frame_name = currentAttribute.TextString
            Exit For
        End If
    Next i
    
    If frame_name = "" Then
        ThisDrawing.Utility.prompt "用紙枠名称が見つかりません。"
        Exit Sub
    End If
    
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
' ## ビューポートの表示調整
'------------------------------------------------------------------------------
Private Sub adjustViewportDisplay(ByRef paper_viewport As ZcadPViewport, _
                                  ByVal min_framepoint As Variant, _
                                  ByVal max_framepoint As Variant)
    
    ThisDrawing.MSpace = True
    
    ThisDrawing.ActivePViewport = paper_viewport
    ZoomWindow min_framepoint, max_framepoint
    
    ThisDrawing.MSpace = False
    
End Sub

'------------------------------------------------------------------------------
' ## レイアウトの印刷設定
'------------------------------------------------------------------------------
Private Sub configurePrintSettings(ByRef new_layout As zcadLayout, _
                                   ByVal frame_width As Single, _
                                   ByVal frame_height As Single)
    
    Dim plotOffset(0 To 1) As Double
    
    ' オフセット量(XYが一般と逆のため注意)
    plotOffset(0) = OFFSET_Y: plotOffset(1) = OFFSET_X
    
    ' 尺度調整
    frame_width = frame_width / SCALE_FACTOR
    frame_height = frame_height / SCALE_FACTOR
    
    With new_layout
        
        ' プリンタ名称
        .ConfigName = PRINTER_NAME
        
        ' 用紙サイズおよび方向
        If frame_width = 420 And frame_height = 297 Then
            .CanonicalMediaName = PAPER_NAME_A3
            .PlotRotation = zc90degrees
        ElseIf frame_width = 297 And frame_height = 420 Then
            .CanonicalMediaName = PAPER_NAME_A3
            .PlotRotation = zc0degrees
        ElseIf frame_width = 210 And frame_height = 297 Then
            .CanonicalMediaName = PAPER_NAME_A4
            .PlotRotation = zc0degrees
        Else
            MsgBox "用紙サイズが検出できなかったため手動設定を行って下さい。"
            Exit Sub
        End If
        
        ' 印刷領域
        .PlotType = zcLayout
        
        ' 印刷オフセット
        .PlotOrigin = plotOffset
        
        ' 尺度(足りないところを追加する)
        .UseStandardScale = False
        .SetCustomScale 1, SCALE_FACTOR
        
        ' 印刷スタイル
        .StyleSheet = STYLE_NAME
        
    End With
    
End Sub
