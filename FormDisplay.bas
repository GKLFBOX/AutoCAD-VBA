Attribute VB_Name = "FormDisplay"
Option Explicit

'------------------------------------------------------------------------------
' ## �����������ݒ�t�@�C���̃t�@�C����
'------------------------------------------------------------------------------
Public Const REFERENCELINE_CONFIG As String = "\ReferenceLine.config"
Public Const STRIKETHROUGH_CONFIG As String = "\Strikethrough.config"

'------------------------------------------------------------------------------
' ## ���C�A�E�g�ҏW�t�H�[���\��
'------------------------------------------------------------------------------
Public Sub DisplayLayoutForm()
    
    ' ���[�h���X�\���̓t�H�[�J�X�����Ȃ����ߎg�p���Ă��Ȃ�
    Load LayoutForm
    LayoutForm.Show
    
End Sub

'------------------------------------------------------------------------------
' ## ���C�A�E�g�ҏW�t�H�[���\��
'------------------------------------------------------------------------------
Public Sub DisplayDecorationLineForm()
    
    ' ���[�h���X�\���̓t�H�[�J�X�����Ȃ����ߎg�p���Ă��Ȃ�
    Load DecorationLineForm
    DecorationLineForm.Show
    
End Sub
