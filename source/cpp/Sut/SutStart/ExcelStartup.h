#pragma once

// �W�F�l���b�N�L�����ɑΉ�����string�N���X
typedef std::basic_string<TCHAR> tstring;

class CExcelStartup
{

public:

    /**
     * �R���X�g���N�^�B
     *
     * @param path �p�X
     */
    CExcelStartup(CString path);

    /**
     * �f�X�g���N�^�B
     */
    ~CExcelStartup(void);

    /**
     * Excel���N������B
     *
     * @return Excel�A�v���P�[�V�����̃I�[�g���[�V�����I�u�W�F�N�g
     */
    IDispatch* startUp();

protected:

    /**
     * Excel�A�v���P�[�V�����̃p�X
     */
    CString excelPath;

};
