#pragma once

class CComComponentRegister
{
public:

    static const int FUNC_SUCCESS = 0;
    static const int FUNC_FAILED  = 1;

    /**
     * �R���X�g���N�^�B
     */
    CComComponentRegister(void);

    /**
     * �f�X�g���N�^�B
     */
    virtual ~CComComponentRegister(void);

    /**
     * Com�R���|�[�l���g�̓o�^�B
     *
     * @param filePath �t�@�C���p�X
     */
    int regist(CString filePath, CString curDir);

    /**
     * Com�R���|�[�l���g�̓o�^�����B
     *
     * @param filePath �t�@�C���p�X
     */
    int unregist(CString filePath, CString curDir);

protected:

    /**
     * regsvr�̎��s�t�@�C����
     */
    static LPCTSTR REG_SVR_EXE;

    /**
     * regsvr32.exe�̎��s�iCOM��o�^���邽�߂̎��s�t�@�C���j
     *
     * @param option �I�v�V����
     * @param filePath �t�@�C���p�X
     */
    int execRegSvr(CString option, CString filePath, CString curDir);

};
