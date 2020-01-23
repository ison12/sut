#pragma once

class CExcelDispatchWrapper
{
public:

    /**
     * �R���X�g���N�^�B
     */
    CExcelDispatchWrapper(IDispatch* applicationDisp);

    /**
     * �f�X�g���N�^�B
     */
    ~CExcelDispatchWrapper(void);

protected:

    /**
     * Excel.Application�I�u�W�F�N�g
     */
    COleDispatchDriver excelApplication;

public:

    /**
     * �߂�l�R�[�h ����
     */
    static const int SUCCESS = 0;

    /**
     * �߂�l�R�[�h ���s�iDISPID�̎擾�Ɏ��s�j
     */
    static const int ERROR_DISPID_NOT_FOUND = 999;

    /**
     * �߂�l�R�[�h ���s�i�\�����ʃG���[�j
     */
    static const int ERROR_UNEXPECTED = 1000;

    /**
     * �߂�l�R�[�h ���s�i���ڂ�������Ȃ��j
     */
    static const int ERROR_ITEM_NOT_FOUND = 2000;

    /**
     * �߂�l�R�[�h ���s�i���ڂ����������j
     */
    static const int ERROR_ITEM_EXIST = 2001;

    /**
     * �I�u�W�F�N�g����Excel.Application�����O���B
     *
     * @return Excel.Application IDispatch
     */
    IDispatch* detachIDispatch();

    /**
     * �o�[�W�������擾����B
     *
     * @return Excel�o�[�W����
     */
    CString getVersion();

    /**
     * Excel�A�v���̌x���\���L����ύX����B
     *
     * @param b TRUE�̏ꍇ�A�x����\������
     * @return SUCCESS ����
     */
    int displayAlerts(BOOL b);

    /**
     * Excel�A�v���̕\���X�e�[�^�X��ύX����B
     *
     * @param visible TRUE�̏ꍇ�AExcel�A�v����\������
     * @return SUCCESS ����
     */
    int appVisible(BOOL visible);

    /**
     * Excel�A�v�����I������B
     *
     * @return SUCCESS ����
     */
    int appQuit();

    /**
     * Addins�I�u�W�F�N�g���擾����B
     *
     * @return Addins�I�u�W�F�N�g
     */
    IDispatch* getAddinsObject();

    /**
     * �A�h�C����ǉ�����B
     *
     * @param addinPath �A�h�C���p�X
     * @return SUCCESS ����
     */
    int attachAddin(CString addinPath);

    /**
     * �A�h�C�����폜����B
     *
     * @param addinPath �A�h�C���p�X
     * @return SUCCESS ����
     */
    int removeAddin(CString addinPath);

};
