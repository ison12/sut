#pragma once

class ExcelAddinRegistry
{
public:

	/**
	 * �R���X�g���N�^�B
	 */
	ExcelAddinRegistry(void);

	/**
	 * �f�X�g���N�^�B
	 */
	virtual ~ExcelAddinRegistry(void);

private:

    /**
     * ���W�X�g���p�X�@Excel�̃Z�L�����e�B�ݒ���̃p�X�̌����E�u���Ώە�����
     */
    static LPCTSTR REG_PATH_EXCEL_PARAM_VERSION;

    /**
     * ���W�X�g���p�X�@Excel�̃A�h�C���Ǘ��f�B���N�g��
     */
    static LPCTSTR REG_PATH_EXCEL_ADDIN_MANAGER;

    /**
     * ���W�X�g���p�X�@Excel�̃A�h�C���I�v�V�����f�B���N�g��
     */
    static LPCTSTR REG_PATH_EXCEL_ADDIN_OPTIONS;

public:

	/**
	 * Excel�A�h�C����ǉ�����B
	 */
	int installAddin(CString& version, CString& addinFilePath);

	/**
	 * Excel�A�h�C�����폜����B
	 */
	int uninstallAddin(CString& version, CString& addinFilePath);

	/**
	 * Excel�A�h�C�����폜����B
	 */
	int uninstallAddinAtAddinManager(CString& version, CString& addinFilePath);

};
