Attribute VB_Name = "FormDisplay"
Option Explicit

'------------------------------------------------------------------------------
' ## 文字装飾線設定ファイルのファイル名
'------------------------------------------------------------------------------
Public Const REFERENCELINE_CONFIG As String = "\ReferenceLine.config"
Public Const STRIKETHROUGH_CONFIG As String = "\Strikethrough.config"

'------------------------------------------------------------------------------
' ## レイアウト編集フォーム表示
'------------------------------------------------------------------------------
Public Sub DisplayLayoutForm()
    
    ' モードレス表示はフォーカスが取れないため使用していない
    Load LayoutForm
    LayoutForm.Show
    
End Sub

'------------------------------------------------------------------------------
' ## 文字装飾設定フォーム表示
'------------------------------------------------------------------------------
Public Sub DisplayDecorationLineForm()
    
    Load DecorationLineForm
    DecorationLineForm.Show
    
End Sub
