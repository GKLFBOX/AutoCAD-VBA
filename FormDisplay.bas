Attribute VB_Name = "FormDisplay"
Option Explicit

'------------------------------------------------------------------------------
' ## レイアウト編集フォーム表示
'------------------------------------------------------------------------------
Public Sub DisplayLayoutForm()
    
    ' モードレス表示はフォーカスが取れないため使用していない
    Load LayoutForm
    LayoutForm.Show
    
End Sub
