
// SutStartDlg.h : �w�b�_�[ �t�@�C��
//

#pragma once
#include "afxcmn.h"
#include "ExcelInfoGetter.h"
#include "ExcelStartup.h"
#include "ExcelDispatchWrapper.h"
#include "ComComponentRegister.h"
#include "CommandExecutor.h"
#include "afxwin.h"

// CSutStartDlg �_�C�A���O
class CSutStartDlg : public CDialog
{
// �R���X�g���N�V����
public:
	CSutStartDlg(CWnd* pParent = NULL);	// �W���R���X�g���N�^

// �_�C�A���O �f�[�^
	enum { IDD = IDD_SUTSTART_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV �T�|�[�g


// ����
protected:
	HICON m_hIcon;

	// �������ꂽ�A���b�Z�[�W���蓖�Ċ֐�
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:

protected:

    /**
     * �C���X�g�[���ς�Excel�A�v���P�[�V�������X�g
     */
    CListCtrl m_installedExcelList;

	/**
	 * COM�R���|�[�l���g�폜�`�F�b�N�{�b�N�X
	 */
	CButton m_comDelete;

	/**
	 * ���W�X�g���폜�`�F�b�N�{�b�N�X
	 */
	CButton m_regDelete;

    /**
     * �C���X�g�[���{�^��
     */
    CButton m_install;
    
    /**
     * �A���C���X�g�[���{�^��
     */
    CButton m_uninstall;

    /**
     * Excel�̏����擾����I�u�W�F�N�g
     */
    CExcelInfoGetter m_excelInfoGetter;

    /**
     * �C���X�g�[���ς�Excel�A�v���P�[�V�������X�g�R���g���[�������������鏈��
     */
    void initInstalledExcelList();

    /**
     * Excel���N������B
     *
     * @return Excel.Application�I�u�W�F�N�g
     */
    IDispatch* launchExcelBySelectedListItem(int itemIndex);

    /**
     * Sut.xlam�̃p�X���擾����B
     *
     * @return Sut.xlam�̃p�X
     */
    CString getSutAddinPath();

    /**
     * �t�@�C�������݂��邩���m�F����B
     *
     * @param filePath �t�@�C���p�X
     * @return true �t�@�C�������݂���
     */
    bool existFile(CString filePath);

    CString getExeFilePath();

    /**
     * Com�R���|�[�l���g��o�^����B
     */
    bool registComComponent();

    /**
     * Com�R���|�[�l���g�̓o�^����������B
     */
    bool unregistComComponent();

    /**
     * �Z�L�����e�B�u���b�N�̉����B
     */
    bool releaseSecurityBlock();

    /**
     * �C���X�g�[�������B
     */
	void processInstall();

    /**
     * �A���C���X�g�[�������B
     */
	void processUninstall();

public:
	afx_msg HBRUSH OnCtlColor(CDC *pDC, CWnd *pWnd, UINT nCtlColor);
    afx_msg void OnNMDblclkInstalledExcelList(NMHDR *pNMHDR, LRESULT *pResult);
    afx_msg void OnBnClickedInstall();
    afx_msg void OnBnClickedUninstall();
};
