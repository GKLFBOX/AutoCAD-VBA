Attribute VB_Name = "AlignTextPosition"
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## 文字位置調整
'
' 指示した2点とオフセット係数から文字位置を調整(中央揃え)
'------------------------------------------------------------------------------
Public Sub AlignTextPosition()
    
    On Error GoTo Error_Handler
    
    ' 対象の選択/指定/入力
    Dim pickPoint As Variant
    Dim targetText As ZcadEntity    ' TODO: targetは型指定を要否を検討する
    ThisDrawing.Utility.GetEntity targetText, pickPoint, _
        "文字オブジェクトを選択 [Cancel(ESC)]"
    
    If Not CommonFunction.IsTextObject(targetText) Then
        ThisDrawing.Utility.Prompt "エラー：文字を選択してください。" & vbCrLf
        Exit Sub
    End If
    
    targetText.Highlight True
    
    Dim firstPoint As Variant
    Dim secondPoint As Variant
    firstPoint = ThisDrawing.Utility.GetPoint _
        (, "1点目を指定 [Cancel(ESC)]")
    secondPoint = ThisDrawing.Utility.GetPoint _
        (firstPoint, "2点目を指定 [Cancel(ESC)]")
    
    Dim offsetFactor As Double
    offsetFactor = ThisDrawing.Utility.GetReal _
        ("オフセット係数を入力(文字高さに対する割合(x/10) " & _
         "[通常(2)/広め(3)/狭い(1)/超広め(5)]:")
    offsetFactor = offsetFactor * 0.1
    
    Dim underFlag As String
    underFlag = ThisDrawing.Utility.GetString _
        (0, "下付きにしますか? [はい(Y)/いいえ(N)]:")
    
    If underFlag = "Y" Then offsetFactor = offsetFactor * -1
    
    targetText.Rotation = 0 ' オフセット量の適用簡略化のため角度要素削除
    
    ' 中点位置の取得
    Dim textCenter() As Double
    textCenter = getTextCenter(targetText)
    textCenter(1) = textCenter(1) - targetText.Height * Abs(offsetFactor)
    
    ' 文字位置調整の実行
    Dim alignPoint() As Double
    Dim alignRad As Double
    alignPoint = getMiddlePoint(firstPoint, secondPoint)
    alignRad = calculateAngle(firstPoint, secondPoint)
    
    targetText.Move textCenter, alignPoint
    targetText.Rotate alignPoint, alignRad
    
    ' 下付き判定と処理
    If offsetFactor < 0 Then
        Dim mirrorText As ZcadEntity
        Set mirrorText = targetText.Mirror(firstPoint, secondPoint)
        targetText.Delete
    Else
        Call CommonSub.ResetHighlight(targetText)
    End If
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ResetHighlight(targetText)
    ThisDrawing.Utility.Prompt "エラー：コマンドを終了します。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## 2点の角度計算
'------------------------------------------------------------------------------
Private Function calculateAngle(ByRef first_point As Variant, _
                                ByRef second_point As Variant) As Double
    
    calculateAngle = Atn2(second_point(0) - first_point(0), _
                          second_point(1) - first_point(1))
    
End Function

'------------------------------------------------------------------------------
' ## 全角度対応Atn関数
'------------------------------------------------------------------------------
Private Function Atn2(delta_x As Double, delta_y As Double) As Double
    
    Dim pi As Double
    pi = 4 * Atn(1)
    
    If delta_x = 0 And delta_y = 0 Then
        Atn2 = 0
        
    ElseIf delta_x > 0 And delta_y = 0 Then ' θ=0
        Atn2 = (pi / 2) * 0
        
    ElseIf delta_x = 0 And delta_y > 0 Then ' θ=90
        Atn2 = (pi / 2) * 1
        
    ElseIf delta_x < 0 And delta_y = 0 Then ' θ=180
        Atn2 = (pi / 2) * 2
        
    ElseIf delta_x = 0 And delta_y < 0 Then ' θ=270
        Atn2 = (pi / 2) * 3
        
    ElseIf delta_x > 0 And delta_y > 0 Then ' 0<θ<90
        Atn2 = Atn(Abs(delta_y) / Abs(delta_x)) + ((pi / 2) * 0)
        
    ElseIf delta_x < 0 And delta_y > 0 Then ' 90<θ<180
        Atn2 = ((pi / 2) - Atn(Abs(delta_y) / Abs(delta_x))) + ((pi / 2) * 1)
        
    ElseIf delta_x < 0 And delta_y < 0 Then ' 180<θ<270
        Atn2 = Atn(Abs(delta_y) / Abs(delta_x)) + ((pi / 2) * 2)
        
    ElseIf delta_x > 0 And delta_y < 0 Then ' 90<θ<180
        Atn2 = ((pi / 2) - Atn(Abs(delta_y) / Abs(delta_x))) + ((pi / 2) * 3)
        
    End If
    
End Function

'------------------------------------------------------------------------------
' ## 2点の中点取得
'------------------------------------------------------------------------------
Private Function getMiddlePoint(ByRef first_point As Variant, _
                                ByRef second_point As Variant) As Double()
    
    Dim i As Long
    Dim tmp(0 To 2) As Double
    
    For i = 0 To 2
        tmp(i) = (first_point(i) + second_point(i)) / 2
    Next i
    
    getMiddlePoint = tmp()
    
End Function

'------------------------------------------------------------------------------
' ## 文字の下中心取得
'------------------------------------------------------------------------------
Private Function getTextCenter(ByVal target_object As ZcadEntity) As Double()
    
    Dim minExtent As Variant, maxExtent As Variant
    target_object.GetBoundingBox minExtent, maxExtent
    
    ' BoundingBoxの仕様変更に伴い傾斜角度を考慮する
    Dim targetOblique As Double
    Dim boxHeight As Double
    Dim exAmount As Double
    targetOblique = target_object.ObliqueAngle
    boxHeight = maxExtent(1) - minExtent(1)
    exAmount = boxHeight * Tan(targetOblique)
    maxExtent(0) = maxExtent(0) + exAmount
    
    Dim startPoint(0 To 2) As Double, endPoint(0 To 2) As Double
    startPoint(0) = minExtent(0): endPoint(0) = maxExtent(0)
    startPoint(1) = minExtent(1): endPoint(1) = minExtent(1)
    startPoint(2) = 0: endPoint(2) = 0
    
    Dim i As Long
    Dim tmp(0 To 2) As Double
    
    For i = 0 To 2
        tmp(i) = (startPoint(i) + endPoint(i)) / 2
    Next i
    
    getTextCenter = tmp()
    
End Function
