
// SutStartDlg.cpp : �����t�@�C��
//

#include "stdafx.h"
#include "SutStart.h"
#include "SutStartDlg.h"
#include "ProgressDlg.h"
#include "ExcelAddinRegistry.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CSutStartDlg �_�C�A���O




CSutStartDlg::CSutStartDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CSutStartDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CSutStartDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialog::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_INSTALLED_EXCEL_LIST, m_installedExcelList);
	DDX_Control(pDX, IDC_CHK_COM_DELETE, m_comDelete);
    DDX_Control(pDX, IDC_CHK_REG_DELETE, m_regDelete);
    DDX_Control(pDX, IDC_INSTALL, m_install);
    DDX_Control(pDX, IDC_UNINSTALL, m_uninstall);
}

BEGIN_MESSAGE_MAP(CSutStartDlg, CDialog)
	ON_WM_PAINT()
	ON_WM_CTLCOLOR()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
    ON_NOTIFY(NM_DBLCLK, IDC_INSTALLED_EXCEL_LIST, &CSutStartDlg::OnNMDblclkInstalledExcelList)
    ON_BN_CLICKED(IDC_INSTALL, &CSutStartDlg::OnBnClickedInstall)
    ON_BN_CLICKED(IDC_UNINSTALL, &CSutStartDlg::OnBnClickedUninstall)
END_MESSAGE_MAP()


// CSutStartDlg ���b�Z�[�W �n���h��

BOOL CSutStartDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ���̃_�C�A���O�̃A�C�R����ݒ肵�܂��B�A�v���P�[�V�����̃��C�� �E�B���h�E���_�C�A���O�łȂ��ꍇ�A
	//  Framework �́A���̐ݒ�������I�ɍs���܂��B
	SetIcon(m_hIcon, TRUE);			// �傫���A�C�R���̐ݒ�
	SetIcon(m_hIcon, FALSE);		// �������A�C�R���̐ݒ�

    // �C���X�g�[���ς�Excel���X�g�R���g���[��������������
    initInstalledExcelList();

    m_excelInfoGetter.existExcelProcess();

	return TRUE;  // �t�H�[�J�X���R���g���[���ɐݒ肵���ꍇ�������ATRUE ��Ԃ��܂��B
}

// �_�C�A���O�ɍŏ����{�^����ǉ�����ꍇ�A�A�C�R����`�悷�邽�߂�
//  ���̃R�[�h���K�v�ł��B�h�L�������g/�r���[ ���f�����g�� MFC �A�v���P�[�V�����̏ꍇ�A
//  ����́AFramework �ɂ���Ď����I�ɐݒ肳��܂��B

void CSutStartDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // �`��̃f�o�C�X �R���e�L�X�g

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// �N���C�A���g�̎l�p�`�̈���̒���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// �A�C�R���̕`��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

HBRUSH CSutStartDlg::OnCtlColor(CDC *pDC, CWnd *pWnd, UINT nCtlColor)
{

	CWnd* target = NULL;
	
	// �^�C�g���R���g���[�����擾����
	target = GetDlgItem(IDC_STATIC_TITLE);

	if (target->GetSafeHwnd() == pWnd->GetSafeHwnd()) {
		// �^�C�g���R���g���[���̐F��ύX����
		pDC->SetBkColor(RGB(0xFF, 0xFF, 0xFF));

		return (HBRUSH)GetStockObject(WHITE_BRUSH);
	}

	return NULL;
}

// ���[�U�[���ŏ��������E�B���h�E���h���b�O���Ă���Ƃ��ɕ\������J�[�\�����擾���邽�߂ɁA
//  �V�X�e�������̊֐����Ăяo���܂��B
HCURSOR CSutStartDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CSutStartDlg::initInstalledExcelList()
{

    // ���X�g�X�^�C�����擾����
    DWORD dwStyle = ListView_GetExtendedListViewStyle(m_installedExcelList.GetSafeHwnd());

    
    dwStyle = dwStyle
		      |  LVS_EX_FULLROWSELECT // �s�S�̂̑I�����\�ɂ���
		      |  LVS_EX_CHECKBOXES;   // �s���`�F�b�N�{�b�N�X�ɂ���

    // �V�����X�^�C����K�p����
    ListView_SetExtendedListViewStyle(m_installedExcelList.GetSafeHwnd(), dwStyle);

    // �C���X�g�[���ς݃G�N�Z���A�v���P�[�V���������擾����
    std::vector<CExcelInfo*>& installedList = m_excelInfoGetter.getInstalledExcelApplication();
    // �T�C�Y���擾
    int installedListSize = installedList.size();

    // ---------------------------------------------
    /*
     * ���X�g�ɃJ������ݒ肷��
     */
    // ---------------------------------------------
    CString colNameVersion;
    CString colNamePath;
    colNameVersion.LoadString(IDS_EXCEL_LIST_COLUMN_VERSION);
    colNamePath.LoadString(IDS_EXCEL_LIST_COLUMN_PATH);

    LVCOLUMN lvColumn;

    lvColumn.mask  = LVCF_FMT
                   | LVCF_WIDTH
                   | LVCF_TEXT
                   | LVCF_SUBITEM;

    lvColumn.fmt      = LVCFMT_LEFT;
    lvColumn.cx       = 130;
    lvColumn.pszText  = (LPTSTR)((LPCTSTR)colNameVersion);
    lvColumn.iSubItem = 0;

    if (m_installedExcelList.InsertColumn(0, &lvColumn) == -1) {

        // �G���[����
    }

    lvColumn.cx       = 500;
    lvColumn.pszText  = (LPTSTR)((LPCTSTR)colNamePath);
    lvColumn.iSubItem = 1;

    if (m_installedExcelList.InsertColumn(1, &lvColumn) == -1) {

        // �G���[����
    }
    // ---------------------------------------------

    // ---------------------------------------------
    /*
     * ���ڂ�ǉ�����
     */
    // ---------------------------------------------
    for (int i = 0; i < installedListSize; i++) {

        CExcelInfo* excelInfo = installedList.at(i);

        LVITEM lvItem;
        lvItem.mask = LVIF_TEXT;
        lvItem.iItem = i;
        lvItem.iSubItem = 0;
        lvItem.pszText = (LPTSTR)excelInfo->appName.c_str();

        int result = m_installedExcelList.InsertItem(&lvItem);
        if (result == -1) {

        }

        lvItem.iSubItem = 1;
        lvItem.pszText = (LPTSTR)excelInfo->appPath.c_str();

        result = m_installedExcelList.SetItem(&lvItem);
        if (result == -1) {

        }

    }
    // ---------------------------------------------

}

IDispatch* CSutStartDlg::launchExcelBySelectedListItem(int itemIndex)
{
    // ���X�g�ɂđI������Ă��鍀�ڂ���
    // �A�v���P�[�V�������ƃA�v���P�[�V�����p�X���擾
    TCHAR appName[256];
    TCHAR appPath[_MAX_PATH];
    m_installedExcelList.GetItemText(itemIndex, 0, appName, 256);
    m_installedExcelList.GetItemText(itemIndex, 1, appPath, _MAX_PATH);

    // Excel�N���p�I�u�W�F�N�g�𐶐�����
    CExcelStartup startup(appPath);
    // Excel���N����Application�I�u�W�F�N�g���擾����
    IDispatch* excelApplicationDisp = startup.startUp();

    if (excelApplicationDisp == NULL) {

        // Excel�̋N���Ɏ��s�����|�̃��b�Z�[�W��\������
        CString infoMessage;
        infoMessage.LoadString(IDS_INFO_EXCEL_LAUNCHED_FAILED);
        CString infoMessage2;
        infoMessage2.LoadString(IDS_INFO_EXCEL_REACTION_PROCESS);
        infoMessage.Append(_T("\n"));
        infoMessage.Append(infoMessage2);
        AfxMessageBox((LPCTSTR)infoMessage, 0, MB_OK | MB_ICONEXCLAMATION);
        
        return NULL;
    }

    // Excel��IDispatch�o�R�ő��삷�邽�߂Ƀ��b�p�[�I�u�W�F�N�g�𐶐�����
    CExcelDispatchWrapper excelDisp(excelApplicationDisp);
    // �x���\�����Ȃ�
    excelDisp.displayAlerts(FALSE);
    // Excel���\���ɂ��Ă���
    excelDisp.appVisible(FALSE);

    int refCnt = excelApplicationDisp->AddRef();

    ATLTRACE2("Excel.Application of refer count is %d\n", refCnt);

    return excelApplicationDisp;
}

CString CSutStartDlg::getSutAddinPath()
{

    // EXE�t�@�C���̔z�u�ꏊ���擾����
    CString exePath = getExeFilePath();

    CString addinFileName;
    CString addinPath;

    //// �p�X�̌������p�X��؂蕶������������
    //pos = exePath.ReverseFind('\\');

    //// �������ꍇ
    //if (pos != -1) {

    //    addinPath = exePath.Mid(0, pos);

    //// �������Ȃ������ꍇ
    //} else {

    //    return CString();
    //}

    addinFileName.LoadString(IDS_ADDIN_FILE_NAME);

    addinPath = exePath;
    addinPath.Append(_T("\\"));
    addinPath.Append(addinFileName);

#ifdef _DEBUG

    addinPath = _T("C:\\Users\\hideki.isobe\\Documents\\sut_work\\Release\\Sut.xlam");
#endif

    return addinPath;

}

bool CSutStartDlg::existFile(CString filePath)
{

    // �t�@�C���X�e�[�^�X
    CFileStatus fileStatus;

    // �t�@�C���X�e�[�^�X�̎擾
    if (CFile::GetStatus(filePath, fileStatus)) {

        // �擾�ł����ꍇ
        return true;

    } else {

        // �擾�ł��Ȃ������ꍇ
        return false;
    }

}

CString CSutStartDlg::getExeFilePath()
{

    // exe�̃C���X�^���X�n���h�����擾����
    HMODULE hModule = AfxGetInstanceHandle();

    // exe����Ƃ����x�[�X�p�X���擾����
    // �x�[�X�p�X
    TCHAR temp[_MAX_PATH];
    // �x�[�X�t�@�C���p�X���擾����
    GetModuleFileName(hModule, temp, _MAX_PATH);

    // SutWhite�̃p�X��tstring�Ɉڂ��ς���
    CString exePath(temp);

    // �p�X�̌������p�X��؂蕶������������
    int pos = exePath.ReverseFind('\\');

    // �������ꍇ
    if (pos != -1) {

        exePath = exePath.Mid(0, pos);
        return exePath;

    // �������Ȃ������ꍇ
    } else {

        return CString();
    }

}


bool CSutStartDlg::registComComponent()
{

    int registSuccess = 0;

    // EXE���z�u����Ă���t�@�C���p�X���擾����
    CString exePath = getExeFilePath();

    // -----------------------------------
    // �ΏۂƂȂ�COM�R���|�[�l���g�t�@�C��
    // SutRed.dll
    CString comSutRed; comSutRed.LoadString(IDS_COM_FILE_SUT_RED);
    // -----------------------------------

    // COM�R���|�[�l���g�o�^�E�������̃J�����g�p�X
    CString comCurDir;
    comCurDir.LoadStringW(IDS_COM_CUR_DIR);

    CString curDir(exePath);
    curDir.Append(comCurDir);

    // COM�R���|�[�l���g�t�@�C���p�X
    CString targetFilePath;

    // COM�R���|�[�l���g�o�^�p�I�u�W�F�N�g
    CComComponentRegister comReg;

    // SutRed.dll��o�^����
    #ifdef _DEBUG
        targetFilePath.Append(_T("D:\\documents\\sut\\CPP\\Sut"));
        targetFilePath.Append(_T("\\Debug ASM"));
        targetFilePath.Append(_T("\\SutRed.dll"));
    #else
        targetFilePath.Append(exePath);
        targetFilePath.Append(comSutRed);
    #endif
    registSuccess  = comReg.regist(targetFilePath, curDir);

    if (registSuccess == CComComponentRegister::FUNC_SUCCESS) {

        return true;
    } else {

        return false;
    }
}

bool CSutStartDlg::unregistComComponent()
{

    int registSuccess = 0;

    // EXE���z�u����Ă���t�@�C���p�X���擾����
    CString exePath = getExeFilePath();

    // -----------------------------------
    // �ΏۂƂȂ�COM�R���|�[�l���g�t�@�C��
    // SutRed.dll
    CString comSutRed; comSutRed.LoadString(IDS_COM_FILE_SUT_RED);
    // -----------------------------------

    // COM�R���|�[�l���g�o�^�E�������̃J�����g�p�X
    CString comCurDir;
    comCurDir.LoadStringW(IDS_COM_CUR_DIR);

    CString curDir(exePath);
    curDir.Append(comCurDir);

    // COM�R���|�[�l���g�t�@�C���p�X
    CString targetFilePath;

    // COM�R���|�[�l���g�o�^�p�I�u�W�F�N�g
    CComComponentRegister comReg;

    // SutRed.dll��o�^����
    #ifdef _DEBUG
        targetFilePath.Append(_T("C:\\Users\\hideki.isobe\\Documents\\sut_work\\CPP\\Sut\\SutRed"));
        targetFilePath.Append(_T("\\Debug ASM"));
        targetFilePath.Append(_T("\\SutRed.dll"));
    #else
        targetFilePath.Append(exePath);
        targetFilePath.Append(comSutRed);
    #endif
    registSuccess  = comReg.unregist(targetFilePath, curDir);

    if (registSuccess == CComComponentRegister::FUNC_SUCCESS
		|| registSuccess == 5) {

        return true;
    } else {

        return false;
    }

}


bool CSutStartDlg::releaseSecurityBlock()
{

    int registSuccess = 0;

    // EXE���z�u����Ă���t�@�C���p�X���擾����
    CString exePath = getExeFilePath();

    CString curDir(exePath);

    // PowerShell�R�}���h
    CString powerShell = _T("powershell");
    CString powerShellCommand;
	powerShellCommand.Format(_T("-Command \"Get-ChildItem '%s\\*.*' -Recurse | Unblock-File\""), (LPCTSTR)exePath);

    // COM�R���|�[�l���g�o�^�p�I�u�W�F�N�g
    CCommandExecutor command;

    #ifdef _DEBUG

		registSuccess  = command.exec(powerShell, powerShellCommand, curDir);

		if (registSuccess == CCommandExecutor::FUNC_SUCCESS) {

			return true;
		} else {

			return false;
		}

    #else

		registSuccess  = command.exec(_T("powershell"), powerShellCommand, curDir);

		if (registSuccess == CCommandExecutor::FUNC_SUCCESS) {

			return true;
		} else {

			return false;
		}

    #endif
}

void CSutStartDlg::OnNMDblclkInstalledExcelList(NMHDR *pNMHDR, LRESULT *pResult)
{
    LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
    // TODO: �����ɃR���g���[���ʒm�n���h�� �R�[�h��ǉ����܂��B

    *pResult = 0;
}

void CSutStartDlg::OnBnClickedInstall()
{
	BeginWaitCursor();

	processInstall();

	EndWaitCursor();
}

void CSutStartDlg::OnBnClickedUninstall()
{
	BeginWaitCursor();

	processUninstall();

	EndWaitCursor();
}

void CSutStartDlg::processInstall()
{
	// �`�F�b�N����
	int checkedCount = 0;

	// ���X�g�{�b�N�X�̃A�C�e�������擾
	int count = m_installedExcelList.GetItemCount();

	for (int i = 0; i < count; i++) {

		// �`�F�b�N����Ă���ꍇ
		if (m_installedExcelList.GetCheck(i)) {

			checkedCount++;

            // �C���X�g�[���̗���͈ȉ��̂Ƃ���Ƃ���B
            // (1). �Z�L�����e�B�u���b�N�̉���
            // (3). Excel�A�h�C���̃C���X�g�[��

			bool releasedSecutiryBlock = releaseSecurityBlock();

            if (!releasedSecutiryBlock) {

                // �G���[����
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_SECURITY_BLOCK_RELEASE_FAILED);
				AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONWARNING);
            }

            // COM�R���|�[�l���g��o�^����
            bool registedCom = registComComponent();

            if (!registedCom) {

                // �G���[����
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_COM_INSTALLED_FAILED);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONERROR);
                return;
            }

            // �A�h�C���t�@�C���p�X���擾���t�@�C���̑��݃`�F�b�N�����{����            
            CString addinFilePath = getSutAddinPath();
            
            if (!existFile(addinFilePath)) {

                // ���s�����ꍇ
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_ADDIN_FILE_NOT_FOUND);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONHAND);
                
                return;
            }

            // Excel���N�����Ă���ꍇ
            if (m_excelInfoGetter.existExcelProcess()) {

                // ���s�����ꍇ
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_EXCEL_PROCESS_EXIST);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONEXCLAMATION);
                
                return;
            }

            // COM�R���|�[�l���g�o�^�Ȍ�A�G���[�����������ꍇ
            // ���[���o�b�N���邽�߂�COM�̓o�^����������B
            // ������

            // ���ݑI������Ă��鍀�ڂ��擾����
            int itemIndex = i;

			// ���X�g�ɂđI������Ă��鍀�ڂ���
			// �A�v���P�[�V�������ƃA�v���P�[�V�����p�X���擾
			TCHAR appName[256];
			m_installedExcelList.GetItemText(itemIndex, 0, appName, 256);

			// Excel�̃o�[�W�������擾����
			CString ver;
			m_excelInfoGetter.getExcelVersionByName(appName, ver);

            // Excel�A�h�C�����C���X�g�[���\�����`�F�b�N����
            int isPossibleAddinInstall = m_excelInfoGetter.isPossibleAddinInstall(ver);

            if (isPossibleAddinInstall == CExcelInfoGetter::POSSIBLE_ADDIN_INSTALL_UNEXPECTED) {

                // �\�����ʃG���[
                // ���s�����ꍇ
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_CHECK_ADDIN_INSTALL_UNEXPECTED);
				infoMessage.Append(_T(" ver"));
				infoMessage.Append(ver);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONHAND);
                return;
            }
            else if (isPossibleAddinInstall == CExcelInfoGetter::POSSIBLE_ADDIN_INSTALL_NG) {

                // NG�̏ꍇ
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_CHECK_ADDIN_INSTALL_NG);
                CString infoMessage2;
                infoMessage2.LoadString(IDS_INFO_CHECK_ADDIN_INSTALL_NG_DETAIL);
                infoMessage.Append(_T("\n"));
                infoMessage.Append(infoMessage2);

                // �m�F�_�C�A���O��\������
                int messResult = AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONEXCLAMATION);

                return;
            }

			ExcelAddinRegistry addinReg;
			if (addinReg.installAddin(ver, addinFilePath) != ERROR_SUCCESS) {

                // �\�����ʃG���[
                // ���s�����ꍇ
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_CHECK_ADDIN_INSTALL_UNEXPECTED);
				infoMessage.Append(_T(" ver:"));
				infoMessage.Append(ver);
				infoMessage.Append(_T(" addin:"));
				infoMessage.Append(addinFilePath);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONHAND);
                return;
			}

            CString infoMessage;
			infoMessage.LoadString(IDS_INFO_EXCEL_ADDED_ADDIN_SUCCESS);
			AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONINFORMATION);
		}
	}

	// �`�F�b�N���ꂽ������0�̏ꍇ
	if (checkedCount == 0) {

		CString infoMessage;
		infoMessage.LoadString(IDS_SELECTED_ONE_MORE);
		AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONEXCLAMATION);
	}
}

void CSutStartDlg::processUninstall()
{
	// com�R���|�[�l���g�폜
	BOOL isComDelete = m_comDelete.GetCheck();
	// ���W�X�g���폜
	BOOL isRegDelete = m_regDelete.GetCheck();

	ATLTRACE2("Dose user delete com component? %s\n", isComDelete == TRUE ? "true" : "false");
	ATLTRACE2("Dose user delete registry? %s\n", isRegDelete == TRUE ? "true" : "false");

	// �`�F�b�N����
	int checkedCount = 0;

	// ���X�g�{�b�N�X�̃A�C�e�������擾
	int count = m_installedExcelList.GetItemCount();

	for (int i = 0; i < count; i++) {

		// �`�F�b�N����Ă���ꍇ
		if (m_installedExcelList.GetCheck(i)) {

			checkedCount++;

            // Excel���N�����Ă���ꍇ
            if (m_excelInfoGetter.existExcelProcess()) {

                // ���s�����ꍇ
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_EXCEL_PROCESS_EXIST);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONEXCLAMATION);
                
                return;
            }

            // ���ݑI������Ă��鍀�ڂ��擾����
            int itemIndex = i;

            // �A�h�C���t�@�C���p�X���擾���t�@�C���̑��݃`�F�b�N�����{����            
            CString addinFilePath = getSutAddinPath();

			// ���X�g�ɂđI������Ă��鍀�ڂ���
			// �A�v���P�[�V�������ƃA�v���P�[�V�����p�X���擾
			TCHAR appName[256];
			m_installedExcelList.GetItemText(itemIndex, 0, appName, 256);

			// Excel�̃o�[�W�������擾����
			CString ver;
			m_excelInfoGetter.getExcelVersionByName(appName, ver);

			ExcelAddinRegistry addinReg;
			if (addinReg.uninstallAddin(ver, addinFilePath) != ERROR_SUCCESS) {

                // �\�����ʃG���[
                // ���s�����ꍇ
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_CHECK_ADDIN_INSTALL_UNEXPECTED);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONHAND);
                return;
			}
			if (addinReg.uninstallAddinAtAddinManager(ver, addinFilePath) != ERROR_SUCCESS) {

                // �\�����ʃG���[
                // ���s�����ꍇ
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_CHECK_ADDIN_INSTALL_UNEXPECTED);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONHAND);
                return;
			}

            CString infoMessage;
            infoMessage.LoadString(IDS_INFO_COMPLETELY_DELETE_ADDIN_OK);
            AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONINFORMATION);
		}
	}

	// com�R���|�[�l���g���폜����
	if (isComDelete) {
		checkedCount++;

        // �A�h�C�������S�ɍ폜�ł����ꍇ��COM�R���|�[�l���g���폜����
        // COM�R���|�[�l���g��o�^��������
        bool unregistedCom = unregistComComponent();

        if (!unregistedCom) {

            // �G���[����
            CString infoMessage;
            infoMessage.LoadString(IDS_INFO_COM_UNINSTALLED_FAILED);
            AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONERROR);
            return;
        }

        CString infoMessage;
        infoMessage.LoadString(IDS_INFO_COM_UNINSTALLED_SUCCESS);
		AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONINFORMATION);

	}

	// ���W�X�g�����폜����
	if (isRegDelete) {
		checkedCount++;

		// �֐��̖߂�l���i�[����
		LONG lResult;

		CString key(LPCTSTR(_T("Software\\ison\\Sut")));
		lResult = AfxGetApp()->DelRegTree(HKEY_CURRENT_USER, key);

		if (lResult != ERROR_SUCCESS && lResult != ERROR_FILE_NOT_FOUND) {

            // �G���[����
            CString infoMessage;
			infoMessage.LoadString(IDS_INFO_REG_DELETE_FAILED);
            AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONERROR);
            return;
		}

        CString infoMessage;
        infoMessage.LoadString(IDS_INFO_REG_DELETE_SUCCESS);
		AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONINFORMATION);
	}

	// �`�F�b�N���ꂽ������0�̏ꍇ
	if (checkedCount == 0) {

		CString infoMessage;
		infoMessage.LoadString(IDS_SELECTED_ONE_MORE_OR_UNINSTALL_OPTION);
		AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONEXCLAMATION);
	}

}
