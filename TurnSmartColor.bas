Attribute VB_Name = "TurnSmartColor"
Option Explicit

'------------------------------------------------------------------------------
' ## オブジェクトの種類に応じた最適な色切り替え   2020/08/09 G.O.
'
' 色が赤以外の場合は赤に変更し赤の場合はByLayerに変更する
' 下記5グループを対象にそれぞれに応じて色変更を行う
' 1.[円弧,円,楕円,ハッチング,2Dポリライン,線,マルチテキスト,
'   ポリライン,放射線,スプライン,文字,構築線]
' 2.[ブロック参照]
' 3.[3点角度寸法,平行寸法,角度寸法,円弧の長さ寸法,長さ寸法]
' 4.[直径寸法,半径寸法]
' 5.[引出線]
'------------------------------------------------------------------------------
Public Sub TurnSmartColor()
    
    On Error GoTo Error_Handler
    
    Dim targetSelectionSet As ZcadSelectionSet
    Dim targetEntity As ZcadEntity
    Dim targetLayer As ZcadLayer
    Dim noChange As Long
    
    ThisDrawing.Utility.Prompt _
        "色変更するオブジェクトを選択してください。" & vbCrLf
    
    ' 出力対象を範囲選択
    Set targetSelectionSet = ThisDrawing.SelectionSets.Add("NewSelectionSet")
    targetSelectionSet.SelectOnScreen
    
    If targetSelectionSet.Count = 0 Then
        Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
        Exit Sub
    End If
    
    noChange = 0
    For Each targetEntity In targetSelectionSet
        
        Set targetLayer = ThisDrawing.Layers.Item(targetEntity.Layer)
        If targetLayer.Lock Then GoTo Continue_targetEntity
        
        If isGroup1(targetEntity) Then
            ' 線,文字等
            Call turnObjectColor(targetEntity)
        ElseIf isGroup2(targetEntity) Then
            ' ブロック参照
            Call turnGroup2Color(targetEntity, noChange)
        ElseIf isGroup3(targetEntity) Then
            ' 長さ寸法,角度寸法等
            Call turnGroup3Color(targetEntity)
        ElseIf isGroup4(targetEntity) Then
            ' 直径寸法,半径寸法
            Call turnGroup4Color(targetEntity)
        ElseIf isGroup5(targetEntity) Then
            ' 引出線
            Call turnGroup5Color(targetEntity)
        Else
            noChange = noChange + 1
        End If
        
Continue_targetEntity:
        
    Next targetEntity
    
    ' 処理結果表示
    Call displayResult(noChange)
    
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    
    Exit Sub
    
Error_Handler:
    Call CommonSub.ReleaseSelectionSet(targetSelectionSet)
    ThisDrawing.Utility.Prompt "なんらかのエラーです。" & vbCrLf
    
End Sub

'------------------------------------------------------------------------------
' ## グループ1判定
'------------------------------------------------------------------------------
Private Function isGroup1(ByVal target_entity As ZcadEntity) As Boolean
    
    If TypeOf target_entity Is ZcadArc _
    Or TypeOf target_entity Is ZcadCircle _
    Or TypeOf target_entity Is ZcadEllipse _
    Or TypeOf target_entity Is ZcadHatch _
    Or TypeOf target_entity Is ZcadLWPolyline _
    Or TypeOf target_entity Is ZcadLine _
    Or TypeOf target_entity Is ZcadMText _
    Or TypeOf target_entity Is ZcadPolyline _
    Or TypeOf target_entity Is ZcadRay _
    Or TypeOf target_entity Is ZcadSpline _
    Or TypeOf target_entity Is ZcadText _
    Or TypeOf target_entity Is ZcadXline Then
        isGroup1 = True
    Else
        isGroup1 = False
    End If
    
End Function

'------------------------------------------------------------------------------
' ## グループ2判定
'------------------------------------------------------------------------------
Private Function isGroup2(ByVal target_entity As ZcadEntity) As Boolean
    
    If TypeOf target_entity Is ZcadBlockReference Then
        isGroup2 = True
    Else
        isGroup2 = False
    End If
    
End Function

'------------------------------------------------------------------------------
' ## グループ3判定
'------------------------------------------------------------------------------
Private Function isGroup3(ByVal target_entity As ZcadEntity) As Boolean
    
    If TypeOf target_entity Is ZcadDim3PointAngular _
    Or TypeOf target_entity Is ZcadDimAligned _
    Or TypeOf target_entity Is ZcadDimAngular _
    Or TypeOf target_entity Is ZcadDimArcLength _
    Or TypeOf target_entity Is ZcadDimRotated Then
        isGroup3 = True
    Else
        isGroup3 = False
    End If
    
End Function

'------------------------------------------------------------------------------
' ## グループ4判定
'------------------------------------------------------------------------------
Private Function isGroup4(ByVal target_entity As ZcadEntity) As Boolean
    
    If TypeOf target_entity Is ZcadDimDiametric _
    Or TypeOf target_entity Is ZcadDimRadial Then
        isGroup4 = True
    Else
        isGroup4 = False
    End If
    
End Function

'------------------------------------------------------------------------------
' ## グループ5判定
'------------------------------------------------------------------------------
Private Function isGroup5(ByVal target_entity As ZcadEntity) As Boolean
    
    If TypeOf target_entity Is ZcadLeader Then
        isGroup5 = True
    Else
        isGroup5 = False
    End If
    
End Function

'------------------------------------------------------------------------------
' ## オブジェクト色の変更
'------------------------------------------------------------------------------
Private Sub turnObjectColor(ByRef target_entity As ZcadEntity)
    
    Dim changeColor As ZcadZcCmColor
    
    ' 色が赤以外の場合は赤にし赤の場合はByLayerにする
    Set changeColor = New ZcadZcCmColor
    If target_entity.TrueColor.ColorIndex = zcRed Then
        changeColor.ColorIndex = zcByLayer
    Else
        changeColor.ColorIndex = zcRed
    End If
    
    target_entity.TrueColor = changeColor
    
End Sub

'------------------------------------------------------------------------------
' ## グループ2の色変更
'------------------------------------------------------------------------------
Private Sub turnGroup2Color(ByRef target_block As ZcadBlockReference, _
                            ByRef no_change As Long)
    
    Dim i As Long
    Dim replicaEntities As Variant
    Dim extractEntity As ZcadEntity
    Dim extractLayer As ZcadLayer
    Dim colorFlag As Boolean
    Dim colorByLayer As ZcadZcCmColor
    
    replicaEntities = target_block.Explode
    
    ' ブロック内オブジェクトの走査
    colorFlag = False
    For i = 0 To UBound(replicaEntities)
        Set extractEntity = replicaEntities(i)
        ' 色がByBlockのオブジェクトが一つでもある場合は色変更を行う
        If extractEntity.TrueColor.ColorIndex = zcByBlock Then colorFlag = True
        Set extractLayer = ThisDrawing.Layers.Item(extractEntity.Layer)
        With extractLayer
            If .Lock Then
                .Lock = False
                extractEntity.Delete
                .Lock = True
            Else
                extractEntity.Delete
            End If
        End With
    Next i
    
    Set colorByLayer = New ZcadZcCmColor
    
    If colorFlag Then
        Call turnObjectColor(target_block)
    Else
        colorByLayer.ColorIndex = zcByLayer
        target_block.TrueColor = colorByLayer
        no_change = no_change + 1
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## グループ3の色変更
'------------------------------------------------------------------------------
Private Sub turnGroup3Color(ByRef target_entity As ZcadEntity)
    
    Dim colorByLayer As ZcadZcCmColor
    
    Set colorByLayer = New ZcadZcCmColor
    
    ' 補助線有り寸法系オブジェクト用の色変更
    With target_entity
        
        ' 文字の色のみオブジェクト色を継承させる
        If .TextColor = zcByBlock Then
            ' ByBlockの場合は色変更
            Call turnObjectColor(target_entity)
        Else
            ' ByBlock以外の場合オブジェクト色を継承するように変更
            If .TextColor = zcByLayer Then
                ' ByLayerの場合オブジェクト色を赤にする
                colorByLayer.ColorIndex = zcRed
            Else
                colorByLayer.ColorIndex = zcByLayer
            End If
            .TrueColor = colorByLayer
            .TextColor = zcByBlock
        End If
        
        ' 文字の色以外は画層色を継承するように変更
        .DimensionLineColor = zcByLayer
        .ExtensionLineColor = zcByLayer
        
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## グループ4の色変更
'------------------------------------------------------------------------------
Private Sub turnGroup4Color(ByRef target_entity As ZcadEntity)
    
    Dim colorByLayer As ZcadZcCmColor
    
    Set colorByLayer = New ZcadZcCmColor
    
    ' 補助線無し寸法系オブジェクト用の色変更
    With target_entity
        
        ' 文字の色のみオブジェクト色を継承させる
        If .TextColor = zcByBlock Then
            ' ByBlockの場合は色変更
            Call turnObjectColor(target_entity)
        Else
            ' ByBlock以外の場合オブジェクト色を継承するように変更
            If .TextColor = zcByLayer Then
                ' ByLayerの場合オブジェクト色を赤にする
                colorByLayer.ColorIndex = zcRed
            Else
                colorByLayer.ColorIndex = zcByLayer
            End If
            .TrueColor = colorByLayer
            .TextColor = zcByBlock
        End If
        
        ' 文字の色以外は画層色を継承するように変更
        .DimensionLineColor = zcByLayer
        
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## グループ5の色変更
'------------------------------------------------------------------------------
Private Sub turnGroup5Color(ByRef target_entity As ZcadLeader)
    
    Dim colorByLayer As ZcadZcCmColor
    
    Set colorByLayer = New ZcadZcCmColor
    
    ' 引出線オブジェクト用の色変更
    With target_entity
        
        ' オブジェクト色を継承させる
        If .DimensionLineColor = zcByBlock Then
            ' ByBlockの場合は色変更
            Call turnObjectColor(target_entity)
        Else
            ' ByBlock以外の場合オブジェクト色を継承するように変更
            If .DimensionLineColor = zcByLayer Then
                ' ByLayerの場合オブジェクト色を赤にする
                colorByLayer.ColorIndex = zcRed
            Else
                colorByLayer.ColorIndex = zcByLayer
            End If
            .TrueColor = colorByLayer
            .DimensionLineColor = zcByBlock
        End If
        
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## 処理結果表示
'------------------------------------------------------------------------------
Private Sub displayResult(ByVal no_change As Long)
    
    Dim resultText As String
    
    resultText = "選択オブジェクトの色を切り替えました。" & vbCrLf
    
    If no_change > 0 Then
        resultText = resultText _
            & "(対象外オブジェクトが" & no_change & "個ありました。)" & vbCrLf
    End If
    
    ThisDrawing.Utility.Prompt resultText
    
End Sub
