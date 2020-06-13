Attribute VB_Name = "CommonFunction"
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## テキスト判定関数
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
    
    MakeFilePath = Left(ThisDrawing.FullName, Len(ThisDrawing.FullName) - 4)
    MakeFilePath = MakeFilePath & addition_name & file_extension
    
End Function
