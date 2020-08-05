Attribute VB_Name = "CommonFunction"
Option Explicit

'------------------------------------------------------------------------------
' ## テキスト判定
'------------------------------------------------------------------------------
Public Function IsTextObject(ByVal target_object As ZcadEntity) As Boolean
    
    IsTextObject = False
    
    If TypeOf target_object Is ZcadText _
    Or TypeOf target_object Is ZcadMText Then
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

'------------------------------------------------------------------------------
' ## 配列版IsEmpty
'------------------------------------------------------------------------------
Public Function IsEmptyArray(ByRef confirmation_array As Variant) As Boolean
    
    On Error GoTo Error_Handler
    
    ' エラーまたは最大要素数が0未満の場合は空
    IsEmptyArray = IIf(UBound(confirmation_array) < 0, True, False)
    
    Exit Function
    
Error_Handler:
    IsEmptyArray = True
    
End Function

'------------------------------------------------------------------------------
' ## csv用の文字列化(ダブルクォーテーションの付加)
'------------------------------------------------------------------------------
Public Function FormatString(ByVal target_text As String) As String
    
    FormatString = """" & Replace(target_text, """", """""") & """"
    
End Function
