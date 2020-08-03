Attribute VB_Name = "CommitConfig"
Option Explicit

'------------------------------------------------------------------------------
' ## 設定フォルダのフォルダ名
'------------------------------------------------------------------------------
Private Const CONFIG_FOLDER As String = "\config"

'------------------------------------------------------------------------------
' ## 設定ファイルの読み込み
'------------------------------------------------------------------------------
Public Function LoadConfig(ByVal config_filename As String) As String
    
    LoadConfig = ""
    
    Dim configFilePath As String
    
    configFilePath = getProjectPath() & CONFIG_FOLDER & config_filename
    
    If Dir(configFilePath) = "" Then Exit Function
    
    Dim bufferData As String
    Open configFilePath For Input As #1
        Do Until EOF(1)
            Line Input #1, bufferData
            LoadConfig = LoadConfig & bufferData & vbCrLf
        Loop
    Close #1
    
    If LoadConfig <> "" Then
        LoadConfig = Left(LoadConfig, Len(LoadConfig) - Len(vbCrLf))
    End If
    
End Function

'------------------------------------------------------------------------------
' ## プロジェクトファイルのパス取得
'------------------------------------------------------------------------------
Private Function getProjectPath() As String
    
    Dim fullName As String
    fullName = ThisDrawing.Application.VBE.ActiveVBProject.FileName
    getProjectPath = Left(fullName, InStrRev(fullName, "\") - 1)
    
End Function

'------------------------------------------------------------------------------
' ## 設定フォルダの存在確認および作成
'------------------------------------------------------------------------------
Public Sub PrepareConfigFolder()
    
    Dim configFolderPath As String
    
    configFolderPath = getProjectPath() & CONFIG_FOLDER
    
    If Dir(configFolderPath, vbDirectory) = "" Then
        MkDir configFolderPath
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## 設定ファイルへの書き出し
'------------------------------------------------------------------------------
Public Sub SaveConfig(ByVal config_filename As String, _
                      ByVal config_data As String)
    
    Dim configFilePath As String
    configFilePath = getProjectPath() & CONFIG_FOLDER & config_filename
    
    Open configFilePath For Output As #1
        Print #1, config_data
    Close #1
    
End Sub
