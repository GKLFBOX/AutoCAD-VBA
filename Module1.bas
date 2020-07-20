Attribute VB_Name = "Module1"
Option Explicit

'------------------------------------------------------------------------------
' ## 設定値化予定の定数郡
'------------------------------------------------------------------------------
Private Const PAPER_TAG As String = "CPN"       ' 用紙枠判定する属性マーク
Private Const PRINTER_NAME As String = "DWG to PDF.pc5"   ' プリンタ名称
Private Const SCALE_FACTOR As Single = 1        ' 図面の尺度(1:n)
Private Const PAPER_NAME_A3 As String = "ISO_A3_(297.00_x_420.00_MM)"
Private Const PAPER_NAME_A4 As String = "ISO_A4_(297.00_x_210.00_MM)"
Private Const STYLE_NAME As String = "Monochrome.ctb"

'------------------------------------------------------------------------------
' ## 枠ブロック選択によるレイアウト生成
'
' ページ番号を属性として持つ用紙枠ブロックからレイアウトを生成する
'------------------------------------------------------------------------------
Public Sub CreateLayout()
    
    'On Error GoTo Error_Handler
    
    ' 用紙枠ブロックの選択
    Dim pickEntity As ZcadEntity
    Dim pickPoint As Variant
    Dim paperFrame As ZcadBlockReference    ' スコープがおかしい
    ThisDrawing.Utility.GetEntity pickEntity, pickPoint, _
        "用紙枠を選択 [Cancel(ESC)]"
    If TypeOf pickEntity Is ZcadBlockReference Then
        Set paperFrame = pickEntity
    Else
        ThisDrawing.Utility.prompt "ブロック以外が選択されました。"
        Exit Sub
    End If
    
    ' 用紙枠サイズ取得(スコープがおかしい)
    Dim minExtent As Variant, maxExtent As Variant
    Dim frameWidth As Single, frameHight As Single
    paperFrame.GetBoundingBox minExtent, maxExtent
    frameWidth = maxExtent(0) - minExtent(0)
    frameHight = maxExtent(1) - minExtent(1)
    
    ' レイアウト名称取得
    Dim frameAttributes As Variant
    Dim currentAttribute As ZcadAttributeReference
    Dim frameName As String
    Dim i As Long
    frameAttributes = paperFrame.GetAttributes
    For i = 0 To UBound(frameAttributes)
        Set currentAttribute = frameAttributes(i)
        If currentAttribute.TagString = PAPER_TAG Then
            frameName = currentAttribute.TextString
            Exit For
        End If
    Next i
    If frameName = "" Then
        ThisDrawing.Utility.prompt "レイアウト名称が見つかりません。"
        Exit Sub
    End If
    
    ' 新規レイアウトの作成およびアクティブ化
    Dim newLayout As zcadLayout
    Set newLayout = ThisDrawing.Layouts.Add(frameName)
    ThisDrawing.ActiveLayout = newLayout
    
    ' 印刷設定
    newLayout.ConfigName = PRINTER_NAME
    
    Dim flg As Long: flg = 0
    If frameHight = 297 * SCALE_FACTOR Then
        If frameWidth = 420 * SCALE_FACTOR Then
            newLayout.CanonicalMediaName = PAPER_NAME_A3
            newLayout.PlotRotation = zc90degrees
            flg = 1
        ElseIf frameWidth = 210 * SCALE_FACTOR Then
            newLayout.CanonicalMediaName = PAPER_NAME_A4
            newLayout.PlotRotation = zc0degrees
            flg = 1
        End If
    End If
    If flg = 0 Then
        MsgBox "用紙サイズが検出できなかったため手動設定を行って下さい。"
    End If
    
    ' レイアウトビューポート作成とサイズ調整
    Dim paperViewport As ZcadPViewport
    Set paperViewport = ThisDrawing.PaperSpace.Item(1)  ' (0)はモデル空間？
    
    paperViewport.Width = frameWidth
    paperViewport.height = frameHight
    
    Dim frameCenter(0 To 2) As Double
    frameCenter(0) = frameWidth / 2
    frameCenter(1) = frameHight / 2
    frameCenter(2) = 0
    
    paperViewport.center = frameCenter
    
    ZoomExtents
    
    ' ビューポート内に指定用紙を表示
    ThisDrawing.MSpace = True
    
    ThisDrawing.ActivePViewport = paperViewport
    ZoomWindow minExtent, maxExtent
    
    ThisDrawing.MSpace = False
    
    ' その他の印刷設定
    Dim plotOffset(0 To 1) As Double
    plotOffset(0) = -5: plotOffset(1) = -17
    
    With newLayout
        .PlotType = zcLayout
        .UseStandardScale = False
        .SetCustomScale 1, SCALE_FACTOR
        .StyleSheet = STYLE_NAME
        .PlotOrigin = plotOffset
    End With
    
    ' 表示を更新するためにタブ切り替えを行う
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts("Model")
    ThisDrawing.ActiveLayout = newLayout
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.prompt "なんらかのエラーです。"
    
End Sub
