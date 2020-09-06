Attribute VB_Name = "ConstantsError"
Option Explicit

' *********************************************************
' �G���[�Ɋ֘A�����萔���W���[��
'
' �쐬�ҁ@�FIson
' �����@�@�F2009/03/31�@�V�K�쐬
'
' ���L�����F
'
' *********************************************************

Public Const ERR_NUMBER_PROC_CANCEL            As Long = 1 + vbObjectError + 512
Public Const ERR_NUMBER_SQL_EXECUTE_FAILED     As Long = 2 + vbObjectError + 512
Public Const ERR_NUMBER_OUT_OF_RANGE_SHEET     As Long = 3 + vbObjectError + 512
Public Const ERR_NUMBER_OUT_OF_RANGE_SELECTION As Long = 4 + vbObjectError + 512
Public Const ERR_NUMBER_DISCONNECT_DB          As Long = 5 + vbObjectError + 512
Public Const ERR_NUMBER_NOT_SELECTED_SCHEMA    As Long = 6 + vbObjectError + 512
Public Const ERR_NUMBER_NOT_SELECTED_TABLE     As Long = 7 + vbObjectError + 512
Public Const ERR_NUMBER_DUPLICATE_SELECTION_CELL As Long = 8 + vbObjectError + 512
Public Const ERR_NUMBER_OVER_SELECT_COND_CONTROL As Long = 9 + vbObjectError + 512
Public Const ERR_NUMBER_IS_NOT_TABLE_SHEET      As Long = 10 + vbObjectError + 512
Public Const ERR_NUMBER_UNSUPPORT_DB           As Long = 11 + vbObjectError + 512
Public Const ERR_NUMBER_NON_ACTIVE_BOOK        As Long = 12 + vbObjectError + 512
Public Const ERR_NUMBER_NOT_EXIST_TABLE_INFO   As Long = 13 + vbObjectError + 512
Public Const ERR_NUMBER_DLL_FUNCTION_WARNING   As Long = 14 + vbObjectError + 512
Public Const ERR_NUMBER_SHORTCUT_SETTING_FAILED As Long = 15 + vbObjectError + 512
Public Const ERR_NUMBER_POPUP_SETTING_FAILED As Long = 16 + vbObjectError + 512
Public Const ERR_NUMBER_RCLICKMENU_SETTING_FAILED As Long = 17 + vbObjectError + 512
Public Const ERR_NUMBER_FILE_OUTPUT_FAILED As Long = 18 + vbObjectError + 512
Public Const ERR_NUMBER_SQL_EMPTY            As Long = 19 + vbObjectError + 512
Public Const ERR_NUMBER_IS_NOT_SQL_DEFINE_SHEET      As Long = 20 + vbObjectError + 512
Public Const ERR_NUMBER_PK_COLUMN_NOT_FOUND   As Long = 21 + vbObjectError + 512
Public Const ERR_NUMBER_SNAP_DIFF__EXEC_ERROR   As Long = 22 + vbObjectError + 512
Public Const ERR_NUMBER_NOT_SELECTED_DB_CONNECT   As Long = 23 + vbObjectError + 512
Public Const ERR_NUMBER_NOT_SELECTED_TABLE_SHEET  As Long = 25 + vbObjectError + 512
Public Const ERR_NUMBER_CREATE_WORKSHEET_FAILED  As Long = 26 + vbObjectError + 512
Public Const ERR_NUMBER_ACTIVE_ADDIN_BOOK        As Long = 27 + vbObjectError + 512
Public Const ERR_NUMBER_SHEET_MISSING        As Long = 28 + vbObjectError + 512
Public Const ERR_NUMBER_CELL_MAX_LENGTH_OVER As Long = 29 + vbObjectError + 512
Public Const ERR_NUMBER_CELL_MAX_LENGTH_OVER_REFLECT As Long = 30 + vbObjectError + 512

Public Const ERR_NUMBER_REG_EXP_NOT_CREATED   As Long = 997 + vbObjectError + 512
Public Const ERR_NUMBER_REGISTRY_ACCESS_FAILED   As Long = 998 + vbObjectError + 512
Public Const ERR_NUMBER_DLL_FUNCTION_FAILED      As Long = 999 + vbObjectError + 512

Public Const ERR_DESC_PROC_CANCEL              As String = "�������L�����Z������܂����B"
Public Const ERR_DESC_SQL_EXECUTE_FAILED       As String = "SQL���s���ɃG���[���������܂����B"
Public Const ERR_DESC_OUT_OF_RANGE_SHEET       As String = "���R�[�h�����������߁A�S�Ẵ��R�[�h���V�[�g�Ɏ�荞�ނ��Ƃ��ł��܂���ł����B"
Public Const ERR_DESC_OUT_OF_RANGE_SELECTION   As String = "�Z���̑I��̈悪���͔͈͊O�ɂ���܂��B"
Public Const ERR_DESC_DISCONNECT_DB            As String = "�f�[�^�x�[�X�ɐڑ�����Ă��܂���B"
Public Const ERR_DESC_NOT_SELECTED_SCHEMA      As String = "�X�L�[�}��1�ȏ�I�����Ă��������B"
Public Const ERR_DESC_NOT_SELECTED_TABLE       As String = "�e�[�u����1�ȏ�I�����Ă��������B"
Public Const ERR_DESC_DUPLICATE_SELECTION_CELL As String = "�I�������Z�����d�����Ă��܂��B"
Public Const ERR_DESC_OVER_SELECT_COND_CONTROL As String = "�v���C�}���L�[���R���g���[����葽�����ߐ������ݒ�ł��܂���ł����B"
Public Const ERR_DESC_IS_NOT_TABLE_SHEET       As String = "�e�[�u���V�[�g�ł͂Ȃ����ߎ��s�ł��܂���B"
Public Const ERR_DESC_UNSUPPORT_DB             As String = "���Ή���DB�ɐڑ�����Ă��܂��B"
Public Const ERR_DESC_NON_ACTIVE_BOOK          As String = "���[�N�u�b�N���A�N�e�B�u�ɂȂ��Ă��Ȃ����ߎ��s�ł��܂���B"
Public Const ERR_DESC_NOT_EXIST_TABLE_INFO     As String = "�e�[�u����񂪎擾�ł��܂���ł����B" & vbNewLine & _
                                                           "�ڑ�����DB�ɑΏۃe�[�u�������݂��Ă��邩�m�F���Ă��������B"
Public Const ERR_DESC_REG_EXP_NOT_CREATED   As String = "���K�\���I�u�W�F�N�g�̐����Ɏ��s���܂����BPC��IE5.0�ȏオ�C���X�g�[������Ă���K�v������܂��B"
Public Const ERR_DESC_DLL_FUNCTION_WARNING     As String = "DLL�̌Ăяo���Ɏ��s���܂����B"
Public Const ERR_DESC_SHORTCUT_SETTING_FAILED  As String = "�V���[�g�J�b�g�L�[�̐ݒ�Ɏ��s���܂����B"
Public Const ERR_DESC_POPUP_SETTING_FAILED     As String = "�|�b�v�A�b�v���j���[�̐ݒ�Ɏ��s���܂����B"
Public Const ERR_DESC_RCLICKMENU_SETTING_FAILED As String = "�E�N���b�N���j���[�̐ݒ�Ɏ��s���܂����B"
Public Const ERR_DESC_FILE_OUTPUT_FAILED As String = "�t�@�C���o�͂Ɏ��s���܂����B"
Public Const ERR_DESC_SQL_EMPTY                As String = "SQL�������͂ł��B"
Public Const ERR_DESC_IS_NOT_SQL_DEFINE_SHEET  As String = "SQL��`�V�[�g�ł͂Ȃ����ߎ��s�ł��܂���B"
Public Const ERR_DESC_PK_COLUMN_NOT_FOUND      As String = "PK�J������������܂���ł����B"
Public Const ERR_DESC_SNAP_DIFF__EXEC_ERROR    As String = "�X�i�b�v�V���b�g��r���s���ɃG���[���������܂����B"

Public Const ERR_DESC_NOT_SELECTED_DB_CONNECT  As String = "�ڑ�����I�����Ă��������B"
Public Const ERR_DESC_NOT_SELECTED_TABLE_SHEET As String = "�V�[�g��1�ȏ�I�����Ă��������B"
Public Const ERR_DESC_CREATE_WORKSHEET_FAILED  As String = "���[�N�u�b�N�̍쐬�Ɏ��s���܂����B"
Public Const ERR_DESC_ACTIVE_ADDIN_BOOK        As String = "�A�h�C���u�b�N���A�N�e�B�u�ɂȂ��Ă��܂��B" & vbNewLine & "���̃u�b�N���A�N�e�B�u�ɂ��čēx���s���Ă��������B"
Public Const ERR_DESC_SHEET_MISSING            As String = "�ΏۂƂȂ�V�[�g�������ǂݎ��܂���B" & vbNewLine & "�폜���ꂽ�\��������܂��B"
Public Const ERR_DESC_CELL_MAX_LENGTH_OVER     As String = "�Z���ւ̓��͉\�ȍő啶�����i32767�����j�𒴂��āA�f�[�^���������܂����B"
Public Const ERR_DESC_CELL_MAX_LENGTH_OVER_REFLECT    As String = "�Z���ւ̃f�[�^���f���ɍő啶�����i32767�����j�𒴂��āA�f�[�^���������܂����B"

Public Const ERR_DESC_REGISTRY_ACCESS_FAILED   As String = "���W�X�g���̃A�N�Z�X�Ɏ��s���܂����B"
Public Const ERR_DESC_DLL_FUNCTION_FAILED      As String = "DLL�̌Ăяo���Ɏ��s���܂����B"

Public Const ERR_DESC_COLUMN_SIZE_OVER_SHEET_SIZE As String = "�J���������������߁A�S�ẴJ�������V�[�g�Ɏ�荞�ނ��Ƃ��ł��܂���ł����B"

Public Const ERR_MSG_ERROR_LEVEL               As String = "�G���[���������܂����B"

Public Const VALID_ERR_NUMERIC                 As String = "���l����͂��Ă��������B"
Public Const VALID_ERR_INTEGER                 As String = "���l����͂��Ă��������B(���������܂܂�)"
Public Const VALID_ERR_NO_LIST_ITEM            As String = "���X�g���獀�ڂ�I�����Ă��������B"
Public Const VALID_ERR_REQUIRED                As String = "�K�{���͂ł��B"
Public Const VALID_ERR_AND_OVER                As String = "{1}�ȏ�̐��l����͂��Ă��������B"
Public Const VALID_ERR_AND_LESS                As String = "{1}�ȉ��̐��l����͂��Ă��������B"
Public Const VALID_ERR_INVALID                 As String = "���͒l���s���ł��B"
Public Const VALID_ERR_INVALID_SIZE            As String = "���͒l�̃T�C�Y���s���ł��B"
Public Const VALID_ERR_NOT_ALPHA_NUM_MARK_FULL As String = "�p�����܂��͋L��( -   �̂�_)�ƑS�p�����݂̂���͂��Ă��������B"

' =========================================================
' ���A�v���P�[�V�����G���[�`�F�b�N
'
' �T�v�@�@�@�F�{�A�v���P�[�V�����Ŕ��������G���[�ł��邩���`�F�b�N����B
' �����@�@�@�Fnum �G���[�ԍ�
' �߂�l�@�@�FTrue �{�A�v���P�[�V�����Ŕ��������G���[
'
' =========================================================
Public Function isApplicationError(ByVal num As Long) As Boolean

    If 1 + vbObjectError + 512 <= num And num <= 900 + vbObjectError + 512 Then
    
        isApplicationError = True
    Else
    
        isApplicationError = False
    End If
    
End Function
