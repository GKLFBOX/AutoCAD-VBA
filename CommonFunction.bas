Attribute VB_Name = "CommonFunction"
Option Explicit

'------------------------------------------------------------------------------
' ## テキスト判定
'------------------------------------------------------------------------------
Public Function IsTextObject(ByVal target_object As ZcadEntity) As Boolean
    
    IsTextObject = False
    
    If (TypeOf target_object Is ZcadText) _
    Or (TypeOf target_object Is ZcadMText) Then
        IsTextObject = True
    End If
    
End Function

'------------------------------------------------------------------------------
' ## 元ファイル名+任意名称によるファイルパス生成関数
'------------------------------------------------------------------------------
Public Function MakeFilePath(ByVal addition_name As String, _
                             ByVal file_extension As String) As String
    
    MakeFilePath = Left(ThisDrawing.fullName, Len(ThisDrawing.fullName) - 4)
    MakeFilePath = MakeFilePath & addition_name & file_extension
    
End Function

'------------------------------------------------------------------------------
' ## 文字列のリスト照合
'------------------------------------------------------------------------------
Public Function IsMatchList(ByVal target_list As Variant, _
                            ByVal target_value As String) As Boolean
    
    IsMatchList = False
    
    Dim i As Long
    For i = 0 To UBound(target_list)
        If target_value = target_list(i) Then
            IsMatchList = True
            Exit Function
        End If
    Next i
    
End Function

