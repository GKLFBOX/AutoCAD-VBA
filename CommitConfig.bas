Attribute VB_Name = "CommitConfig"
Option Explicit

'------------------------------------------------------------------------------
' ## �ݒ�t�H���_�̃t�H���_��
'------------------------------------------------------------------------------
Private Const CONFIG_FOLDER As String = "\config"

'------------------------------------------------------------------------------
' ## �ݒ�t�@�C���̓ǂݍ���
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
' ## �v���W�F�N�g�t�@�C���̃p�X�擾
'------------------------------------------------------------------------------
Private Function getProjectPath() As String
    
    Dim fullName As String
    fullName = ThisDrawing.Application.VBE.ActiveVBProject.FileName
    getProjectPath = Left(fullName, InStrRev(fullName, "\") - 1)
    
End Function

'------------------------------------------------------------------------------
' ## �ݒ�t�H���_�̑��݊m�F����э쐬
'------------------------------------------------------------------------------
Public Sub PrepareConfigFolder()
    
    Dim configFolderPath As String
    
    configFolderPath = getProjectPath() & CONFIG_FOLDER
    
    If Dir(configFolderPath, vbDirectory) = "" Then
        MkDir configFolderPath
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## �ݒ�t�@�C���ւ̏����o��
'------------------------------------------------------------------------------
Public Sub SaveConfig(ByVal config_filename As String, _
                      ByVal config_data As String)
    
    Dim configFilePath As String
    configFilePath = getProjectPath() & CONFIG_FOLDER & config_filename
    
    Open configFilePath For Output As #1
        Print #1, config_data
    Close #1
    
End Sub
