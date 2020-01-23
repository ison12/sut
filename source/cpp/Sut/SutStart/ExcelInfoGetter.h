#pragma once

#include <afxwin.h>

// �W�F�l���b�N�L�����ɑΉ�����string�N���X
typedef std::basic_string<TCHAR> tstring;

class CExcelInfo
{

public:

    /**
     * �R���X�g���N�^�B
     */
    CExcelInfo(void) {}

    /**
     * �f�X�g���N�^�B
     */
    ~CExcelInfo(void) {}

    /**
     * �A�v���P�[�V������
     */
    tstring appName;

    /**
     * �A�v���P�[�V�����p�X
     */
    tstring appPath;

};

class CExcelInfoGetter
{
public:

    /**
     * �R���X�g���N�^�B
     */
    CExcelInfoGetter(void);

    /**
     * �f�X�g���N�^�B
     */
    ~CExcelInfoGetter(void);

protected:

	// 
	static std::vector<std::map<tstring, tstring>> EXCEL_INFO_LIST;

    /**
     * �o�[�W������ Excel��CLSID
     */
    static LPCTSTR CLSID_COMPONENT_EXCEL[];

    /**
     * �o�[�W������ Excel�̃A�v���P�[�V������
     */
    static LPCTSTR COMPONENT_EXCEL_NAME[];

    /**
     * Excel�̃o�[�W����
     */
    static LPCTSTR EXCEL_VERSION[];

    /**
     * CLSID���i�[���Ă���z��̒���
     */
    static const int CLSID_ARRAY_LENGTH;

    /**
     * �C���X�g�[���ς݂�Excel�A�v���P�[�V������񃊃X�g
     */
    std::vector<CExcelInfo*> installedExcelList;

    /**
     * ���W�X�g���p�X�@Excel�̃Z�L�����e�B�ݒ���̃p�X�̌����E�u���Ώە�����
     */
    static LPCTSTR REG_PATH_EXCEL_PARAM_VERSION;

    /**
     * ���W�X�g���p�X�@Excel�̃Z�L�����e�B�ݒ���
     */
    static LPCTSTR REG_PATH_EXCEL_SECURITY_SETTING;

    /**
     * ���W�X�g���̒l�̖��O Excel�Z�L�����e�B��� �C���X�g�[�����ꂽ�A�h�C����M�p���Ȃ��t���O
     */
    static LPCTSTR REG_VALUE_NAME_EXCEL_SECURITY_DONTTRUSTINSTALLEDFILES;

    /**
     * ���W�X�g���̒l�̖��O Excel�Z�L�����e�B��� �S�ẴA�h�C���𖳌��ɂ���t���O
     */
    static LPCTSTR REG_VALUE_NAME_EXCEL_SECURITY_DISABLEALLADDINS;

    /**
     * ���W�X�g���̒l�̖��O Excel�Z�L�����e�B��� �����ς݃A�h�C���̂ݗL���ɂ���t���O
     */
    static LPCTSTR REG_VALUE_NAME_EXCEL_SECURITY_REQUIREDADDINSIG; 

    /**
     * ���W�X�g���p�X�@Excel�̃A�h�C���Ǘ��f�B���N�g��
     */
    static LPCTSTR REG_PATH_EXCEL_ADDIN_MANAGER;

public:

    static const int POSSIBLE_ADDIN_INSTALL_OK         = 0;
    static const int POSSIBLE_ADDIN_INSTALL_NG         = 1;
    static const int POSSIBLE_ADDIN_INSTALL_UNEXPECTED = 2;

    static const int COMPLETELY_DELETE_ADDIN_OK = 0;
    static const int COMPLETELY_DELETE_ADDIN_TARGET_KEY_NOT_FOUND = 1;
    static const int COMPLETELY_DELETE_ADDIN_UNEXPECTED = 2;
    static const int COMPLETELY_DELETE_ADDIN_SUSPEND = 3;

    /**
     * �C���X�g�[���ς݂�Excel�A�v���P�[�V������񃊃X�g���擾����B
     *
     * @return �C���X�g�[���ς݂�Excel�A�v���P�[�V������񃊃X�g�ithis->installedExcelList��Ԃ��j
     */
    std::vector<CExcelInfo*>& getInstalledExcelApplication();

	/**
	 * �G�N�Z��������o�[�W�������擾����
	 */
	void getExcelVersionByName(const CString& excelName, CString& excelVersion);

    /**
     * Excel�v���Z�X�����݂��邩���`�F�b�N����B
     *
     * @return true �v���Z�X�����݂���
     */
    bool existExcelProcess();

    /**
     * �A�h�C���̃C���X�g�[�����\�����`�F�b�N����B
     *
     * @param excelVersion Excel�̃o�[�W����
     * @return POSSIBLE_ADDIN_INSTALL_OK �C���X�g�[���\
     *         POSSIBLE_ADDIN_INSTALL_NG �C���X�g�[���s��
     *         POSSIBLE_ADDIN_INSTALL_UNEXPECTED �\�����ʃG���[
     */
    int isPossibleAddinInstall(CString excelVersion);

    /**
     * �A�h�C���̃C���X�g�[�����\�����`�F�b�N����B
     * Excel2000�ȍ~
     *
     * @param excelVersion Excel�̃o�[�W����
     * @return POSSIBLE_ADDIN_INSTALL_OK �C���X�g�[���\
     *         POSSIBLE_ADDIN_INSTALL_NG �C���X�g�[���s��
     *         POSSIBLE_ADDIN_INSTALL_UNEXPECTED �\�����ʃG���[
     */
    int isPossibleAddinInstallForOverExcel2000(CString excelVersion);

    /**
     * �A�h�C���̃C���X�g�[�����\�����`�F�b�N����B
     * Excel2007�ȍ~
     *
     * @param excelVersion Excel�̃o�[�W����
     * @return POSSIBLE_ADDIN_INSTALL_OK �C���X�g�[���\
     *         POSSIBLE_ADDIN_INSTALL_NG �C���X�g�[���s��
     *         POSSIBLE_ADDIN_INSTALL_UNEXPECTED �\�����ʃG���[
     */
    int isPossibleAddinInstallForOverExcel2007(CString excelVersion);

    /**
     * �A�h�C�������S�ɍ폜����B
     *
     * �A�h�C���͈ȉ��̃��W�X�g���p�X�Ƀt�@�C���p�X���ۑ�����Ă���A�ȉ�����p�X���폜���Ȃ���
     * �����̃A�h�C�����ăC���X�g�[�����Ă��������L���ɂȂ�Ȃ��B
     *
     * HKEY_CURRENT_USER\Software\Microsoft\Office\�o�[�W������\Excel\Add-in Manager
     *
     * �{���\�b�h�͏�q�����p�X����addinFileName�Ƀ}�b�`���������폜����B
     * �܂��A��q�����p�X��Addin.Install�v���p�e�B��False�ɐݒ肵�AExcel�I����ɂ͂��߂ď������܂��B
     * �]���āA����ȑO�ɍ폜���悤�Ƃ��Ă����݂��Ȃ��B
     *
     * @param excelVersion Excel�̃o�[�W����
     * @param addinFileName �A�h�C���t�@�C����
     */
    int completelyDeleteAddin(CString excelVersion, CString addinFileName);

    /**
     * �I�I�I���쐬
     */
    void searchRunningObjectTable();

};
