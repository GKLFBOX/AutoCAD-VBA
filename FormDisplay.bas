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
' ## ���������ݒ�t�H�[���\��
'------------------------------------------------------------------------------
Public Sub DisplayDecorationLineForm()
    
    Load DecorationLineForm
    DecorationLineForm.Show
    
End Sub

'------------------------------------------------------------------------------
' ## �p���g���X�gcsv�o�̓t�H�[���\��
'------------------------------------------------------------------------------
Public Sub DisplayFrameListForm()
    
    Load FrameListForm
    FrameListForm.Show
    
End Sub
