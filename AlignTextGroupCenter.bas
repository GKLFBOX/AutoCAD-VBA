Attribute VB_Name = "AlignTextGroupCenter"
Option Explicit

'------------------------------------------------------------------------------
' ## 文字系オブジェクト位置調整   2020/08/02 G.O.
'
' 指定した2点とオフセット係数から文字系オブジェクトの位置を中央揃えに調整する
'------------------------------------------------------------------------------
Public Sub AlignTextGroupCenter()
    
    On Error GoTo Error_Handler
    
    Dim targetEntity As ZcadEntity
    Dim pickPoint As Variant
    
    ' 位置調整する文字系オブジェクトの選択
    ThisDrawing.Utility.GetEntity targetEntity, pickPoint, _
        "位置調整する文字またはブロック内文字を選択 [Cancel(ESC)]"
    
    targetEntity.Highlight True
    
    ' テキストまたはブロック参照の判定
    If CommonFunction.IsTextObject(targetEntity) Then
        Call alignTextCenter(targetEntity, pickPoint)
    ElseIf TypeOf targetEntity Is ZcadBlockReference Then
        Call alignBlockCenter(targetEntity, pickPoint)
    Else
        ThisDrawing.Utility.Prompt _
            "文字またはブロック内文字が選択されませんでした。" & vbCrLf
    End If
    
    Call CommonSub.ResetHighlight(targetEntity)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ResetHighlight(targetEntity)
    ThisDrawing.Utility.Prompt "なんらかのエラーです。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## 文字位置調整
'------------------------------------------------------------------------------
Private Sub alignTextCenter(ByRef target_text As ZcadEntity, _
                            ByVal pick_point As Variant)
    
    On Error GoTo Error_Handler
    
    Dim firstPoint As Variant, secondPoint As Variant
    Dim offsetFactor As String
    Dim underFlag As String
    Dim textCenter() As Double
    
    ' 調整値のユーザー入力
    Call inputAlignValue(firstPoint, secondPoint, offsetFactor, underFlag)
    
    ' オフセット計算簡略化のため角度要素削除
    Dim targetAngle As Double
    targetAngle = target_text.Rotation
    target_text.Rotate pick_point, targetAngle * -1
    
    ' 文字の上下中心取得および取得位置のオフセット
    textCenter = getTextCenter(target_text, underFlag)
    Call offsetTextCenter(textCenter, underFlag, offsetFactor, target_text)
    
    ' 文字位置調整の実行
    Call doAlignment(firstPoint, secondPoint, textCenter, target_text)
    
    Exit Sub
    
Error_Handler:
    ThisDrawing.Utility.Prompt "なんらかのエラーです。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## ブロック位置調整
'------------------------------------------------------------------------------
Private Sub alignBlockCenter(ByRef target_block As ZcadBlockReference, _
                             ByVal pick_point As Variant)
    
    On Error GoTo Error_Handler
    
    Dim replicaEntities As Variant
    Dim targetReplica As ZcadEntity
    Dim firstPoint As Variant, secondPoint As Variant
    Dim offsetFactor As String
    Dim underFlag As String
    Dim textCenter() As Double
    
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
    
    ' 調整値のユーザー入力
    Call inputAlignValue(firstPoint, secondPoint, offsetFactor, underFlag)
    
    ' オフセット計算簡略化のため角度要素削除
    Dim targetAngle As Double
    targetAngle = targetReplica.Rotation
    targetReplica.Rotate pick_point, targetAngle * -1
    target_block.Rotate pick_point, targetAngle * -1
    
    ' 文字の上下中心取得および取得位置のオフセット
    textCenter = getTextCenter(targetReplica, underFlag)
    Call offsetTextCenter(textCenter, underFlag, offsetFactor, targetReplica)
    
    ' 文字位置調整の実行
    Call doAlignment(firstPoint, secondPoint, textCenter, target_block)
    
    Call CommonSub.DeleteReplica(replicaEntities)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.DeleteReplica(replicaEntities)
    ThisDrawing.Utility.Prompt "なんらかのエラーです。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## 調整値のユーザー入力
'------------------------------------------------------------------------------
Private Sub inputAlignValue(ByRef first_point As Variant, _
                            ByRef second_point As Variant, _
                            ByRef offset_factor As String, _
                            ByRef under_flag As String)
    
    ' 調整先の2点を指定
    first_point = ThisDrawing.Utility.GetPoint _
        (, "1点目を指定 [Cancel(ESC)]")
    second_point = ThisDrawing.Utility.GetPoint _
        (first_point, "2点目を指定 [Cancel(ESC)]")
    
    ' オフセット係数の入力
    ' ZWCADの不具合でGet系のPromptに組み込んだ値は
    ' 自然数または英字(大文字)しか正常に入力されないことを考慮している
    offset_factor = ThisDrawing.Utility.GetString _
        (0, "オフセット係数を入力(文字高さに対する割合(x/10)) " & _
        "[通常(2)/広め(3)/狭い(1)/超広め(5)]:")
    offset_factor = offset_factor * 0.1
    
    ' 下付きの選択
    under_flag = ThisDrawing.Utility.GetString _
        (0, "下付きにしますか? [はい(Y)/いいえ(N)]:")
    
End Sub

'------------------------------------------------------------------------------
' ## 文字の上下中心取得
'------------------------------------------------------------------------------
Private Function getTextCenter(ByVal target_text As ZcadEntity, _
                               ByVal under_flag As String) As Double()
    
    Dim minExtent As Variant, maxExtent As Variant
    Dim leftPoint(0 To 2) As Double, rightPoint(0 To 2) As Double
    
    target_text.GetBoundingBox minExtent, maxExtent
    
    ' 拡張版GetBoundingBox
    Call CommonSub.GetEnhancedBoundingBox(target_text, minExtent, maxExtent)
    
    ' 上境界または下境界の取得
    If UCase(under_flag) = "Y" Then
        leftPoint(0) = minExtent(0): leftPoint(1) = maxExtent(1)
        rightPoint(0) = maxExtent(0): rightPoint(1) = maxExtent(1)
    Else
        leftPoint(0) = minExtent(0): leftPoint(1) = minExtent(1)
        rightPoint(0) = maxExtent(0): rightPoint(1) = minExtent(1)
    End If
    leftPoint(2) = 0
    rightPoint(2) = 0
    
    getTextCenter = getMiddlePoint(leftPoint, rightPoint)
    
End Function

'------------------------------------------------------------------------------
' ## 中心位置の上下オフセット
'------------------------------------------------------------------------------
Private Sub offsetTextCenter(ByRef text_center() As Double, _
                             ByVal under_flag As String, _
                             ByVal offset_factor As String, _
                             ByVal target_text As ZcadEntity)
    
    If UCase(under_flag) = "Y" Then
        text_center(1) = text_center(1) _
            + target_text.Height * Abs(offset_factor)
    Else
        text_center(1) = text_center(1) _
            - target_text.Height * Abs(offset_factor)
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## 位置調整実行
'------------------------------------------------------------------------------
Private Sub doAlignment(ByVal first_point As Variant, _
                        ByVal second_point As Variant, _
                        ByRef text_center() As Double, _
                        ByRef target_entity As ZcadEntity)
    
    Dim alignPoint() As Double
    Dim alignRadian As Double
    
    alignPoint = getMiddlePoint(first_point, second_point)
    alignRadian = calculateAngle(first_point, second_point)
    
    target_entity.Move text_center, alignPoint
    target_entity.Rotate alignPoint, alignRadian
    
End Sub

'------------------------------------------------------------------------------
' ## 2点の中点取得
'------------------------------------------------------------------------------
Private Function getMiddlePoint(ByRef first_point As Variant, _
                                ByRef second_point As Variant) As Double()
    
    Dim i As Long
    Dim middlePoint(0 To 2) As Double
    
    For i = 0 To 2
        middlePoint(i) = (first_point(i) + second_point(i)) / 2
    Next i
    
    getMiddlePoint = middlePoint()
    
End Function

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
    ElseIf delta_x > 0 And delta_y = 0 Then
        ' θ=0
        Atn2 = (pi / 2) * 0
    ElseIf delta_x = 0 And delta_y > 0 Then
        ' θ=90
        Atn2 = (pi / 2) * 1
    ElseIf delta_x < 0 And delta_y = 0 Then
        ' θ=180
        Atn2 = (pi / 2) * 2
    ElseIf delta_x = 0 And delta_y < 0 Then
        ' θ=270
        Atn2 = (pi / 2) * 3
    ElseIf delta_x > 0 And delta_y > 0 Then
        ' 0<θ<90
        Atn2 = Atn(Abs(delta_y) / Abs(delta_x)) + ((pi / 2) * 0)
    ElseIf delta_x < 0 And delta_y > 0 Then
        ' 90<θ<180
        Atn2 = ((pi / 2) - Atn(Abs(delta_y) / Abs(delta_x))) + ((pi / 2) * 1)
    ElseIf delta_x < 0 And delta_y < 0 Then
        ' 180<θ<270
        Atn2 = Atn(Abs(delta_y) / Abs(delta_x)) + ((pi / 2) * 2)
    ElseIf delta_x > 0 And delta_y < 0 Then
        ' 90<θ<180
        Atn2 = ((pi / 2) - Atn(Abs(delta_y) / Abs(delta_x))) + ((pi / 2) * 3)
    End If
    
End Function
