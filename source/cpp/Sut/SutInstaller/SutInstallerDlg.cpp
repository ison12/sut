
// SutInstallerDlg.cpp : 実装ファイル
//

#include "pch.h"
#include "framework.h"
#include "SutInstaller.h"
#include "SutInstallerDlg.h"
#include "afxdialogex.h"
#include "ExcelAddinRegistry.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CSutInstallerDlg ダイアログ



CSutInstallerDlg::CSutInstallerDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_SUTINSTALLER_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CSutInstallerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_INSTALLED_EXCEL_LIST, m_installedExcelList);
    DDX_Control(pDX, IDC_CHK_REG_DELETE, m_regDelete);
    DDX_Control(pDX, IDC_INSTALL, m_install);
    DDX_Control(pDX, IDC_UNINSTALL, m_uninstall);
}

BEGIN_MESSAGE_MAP(CSutInstallerDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_INSTALL, &CSutInstallerDlg::OnBnClickedInstall)
	ON_BN_CLICKED(IDC_UNINSTALL, &CSutInstallerDlg::OnBnClickedUninstall)
END_MESSAGE_MAP()


// CSutInstallerDlg メッセージ ハンドラー

BOOL CSutInstallerDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// このダイアログのアイコンを設定します。アプリケーションのメイン ウィンドウがダイアログでない場合、
	//  Framework は、この設定を自動的に行います。
	SetIcon(m_hIcon, TRUE);			// 大きいアイコンの設定
	SetIcon(m_hIcon, FALSE);		// 小さいアイコンの設定

	// インストール済みExcelリストコントロールを初期化する
	initInstalledExcelList();

	return TRUE;  // フォーカスをコントロールに設定した場合を除き、TRUE を返します。
}

void CSutInstallerDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	CDialogEx::OnSysCommand(nID, lParam);
}

// ダイアログに最小化ボタンを追加する場合、アイコンを描画するための
//  下のコードが必要です。ドキュメント/ビュー モデルを使う MFC アプリケーションの場合、
//  これは、Framework によって自動的に設定されます。

void CSutInstallerDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 描画のデバイス コンテキスト

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// クライアントの四角形領域内の中央
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// アイコンの描画
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

HBRUSH CSutInstallerDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{

    CWnd* target = NULL;

    // タイトルコントロールを取得する
    target = GetDlgItem(IDC_STATIC_TITLE);

    if (target->GetSafeHwnd() == pWnd->GetSafeHwnd()) {
        // タイトルコントロールの色を変更する
        pDC->SetBkColor(RGB(0xFF, 0xFF, 0xFF));

        return (HBRUSH)GetStockObject(WHITE_BRUSH);
    }

    return NULL;
}

// ユーザーが最小化したウィンドウをドラッグしているときに表示するカーソルを取得するために、
//  システムがこの関数を呼び出します。
HCURSOR CSutInstallerDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CSutInstallerDlg::OnBnClickedInstall()
{
    BeginWaitCursor();

    processInstall();

    EndWaitCursor();
}

void CSutInstallerDlg::OnBnClickedUninstall()
{
    BeginWaitCursor();

    processUninstall();

    EndWaitCursor();
}

void CSutInstallerDlg::initInstalledExcelList()
{

    // リストスタイルを取得する
    DWORD dwStyle = ListView_GetExtendedListViewStyle(m_installedExcelList.GetSafeHwnd());

    dwStyle = dwStyle
        | LVS_EX_FULLROWSELECT // 行全体の選択を可能にする
        | LVS_EX_CHECKBOXES;   // 行をチェックボックスにする

    // 新しいスタイルを適用する
    ListView_SetExtendedListViewStyle(m_installedExcelList.GetSafeHwnd(), dwStyle);

    // インストール済みエクセルアプリケーション情報を取得する
    std::vector<CExcelInfo*>& installedList = m_excelInfoGetter.getInstalledExcelApplication();
    // サイズを取得
    size_t installedListSize = installedList.size();

    // ---------------------------------------------
    /*
     * リストにカラムを設定する
     */
     // ---------------------------------------------
    CString colNameVersion;
    CString colNamePath;
    colNameVersion.LoadString(IDS_EXCEL_LIST_COLUMN_VERSION);
    colNamePath.LoadString(IDS_EXCEL_LIST_COLUMN_PATH);

    LVCOLUMN lvColumn;

    lvColumn.mask = LVCF_FMT
                  | LVCF_WIDTH
                  | LVCF_TEXT
                  | LVCF_SUBITEM;

    lvColumn.fmt = LVCFMT_LEFT;
    lvColumn.cx = 130;
    lvColumn.pszText = (LPTSTR)((LPCTSTR)colNameVersion);
    lvColumn.iSubItem = 0;

    if (m_installedExcelList.InsertColumn(0, &lvColumn) == -1) {

        // エラー処理
    }

    lvColumn.cx = 500;
    lvColumn.pszText = (LPTSTR)((LPCTSTR)colNamePath);
    lvColumn.iSubItem = 1;

    if (m_installedExcelList.InsertColumn(1, &lvColumn) == -1) {

        // エラー処理
    }
    // ---------------------------------------------

    // ---------------------------------------------
    /*
     * 項目を追加する
     */
     // ---------------------------------------------
    for (size_t i = 0; i < installedListSize; i++) {

        CExcelInfo* excelInfo = installedList.at(i);

        LVITEM lvItem;
        lvItem.mask = LVIF_TEXT;
        lvItem.iItem = static_cast<int>(i);
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

IDispatch* CSutInstallerDlg::launchExcelBySelectedListItem(int itemIndex)
{
    // リストにて選択されている項目から
    // アプリケーション名とアプリケーションパスを取得
    TCHAR appName[256];
    TCHAR appPath[_MAX_PATH];
    m_installedExcelList.GetItemText(itemIndex, 0, appName, 256);
    m_installedExcelList.GetItemText(itemIndex, 1, appPath, _MAX_PATH);

    // Excel起動用オブジェクトを生成する
    CExcelStartup startup(appPath);
    // Excelを起動しApplicationオブジェクトを取得する
    IDispatch* excelApplicationDisp = startup.startUp();

    if (excelApplicationDisp == NULL) {

        // Excelの起動に失敗した旨のメッセージを表示する
        CString infoMessage;
        infoMessage.LoadString(IDS_INFO_EXCEL_LAUNCHED_FAILED);
        CString infoMessage2;
        infoMessage2.LoadString(IDS_INFO_EXCEL_REACTION_PROCESS);
        infoMessage.Append(_T("\n"));
        infoMessage.Append(infoMessage2);
        AfxMessageBox((LPCTSTR)infoMessage, 0, MB_OK | MB_ICONEXCLAMATION);

        return NULL;
    }

    // ExcelをIDispatch経由で操作するためにラッパーオブジェクトを生成する
    CExcelDispatchWrapper excelDisp(excelApplicationDisp);
    // 警告表示しない
    excelDisp.displayAlerts(FALSE);
    // Excelを非表示にしておく
    excelDisp.appVisible(FALSE);

    int refCnt = excelApplicationDisp->AddRef();

    ATLTRACE2("Excel.Application of refer count is %d\n", refCnt);

    return excelApplicationDisp;
}

CString CSutInstallerDlg::getSutAddinPath()
{

    // EXEファイルの配置場所を取得する
    CString exePath = getExeFilePath();

    CString addinFileName;
    CString addinPath;

    addinFileName.LoadString(IDS_ADDIN_FILE_NAME);

    addinPath = exePath;
    addinPath.Append(_T("\\"));
    addinPath.Append(addinFileName);

#ifdef _DEBUG

    addinPath = _T("C:\\Users\\hideki.isobe\\Documents\\sut_work\\Release\\Sut.xlam");
#endif

    return addinPath;
}

bool CSutInstallerDlg::existFile(CString filePath)
{
    // ファイルステータス
    CFileStatus fileStatus;

    // ファイルステータスの取得
    if (CFile::GetStatus(filePath, fileStatus)) {

        // 取得できた場合
        return true;
    }
    else {

        // 取得できなかった場合
        return false;
    }

}

CString CSutInstallerDlg::getExeFilePath()
{
    // exeのインスタンスハンドルを取得する
    HMODULE hModule = AfxGetInstanceHandle();

    // exeを基準としたベースパスを取得する
    // ベースパス
    TCHAR temp[_MAX_PATH];
    // ベースファイルパスを取得する
    GetModuleFileName(hModule, temp, _MAX_PATH);

    // SutWhiteのパスをtstringに移し変える
    CString exePath(temp);

    // パスの後方からパス区切り文字を検索する
    int pos = exePath.ReverseFind('\\');

    // 見つけた場合
    if (pos != -1) {

        exePath = exePath.Mid(0, pos);
        return exePath;

        // 見つけられなかった場合
    }
    else {

        return CString();
    }

}

bool CSutInstallerDlg::releaseSecurityBlock()
{

    int registSuccess = 0;

    // EXEが配置されているファイルパスを取得する
    CString exePath = getExeFilePath();

    CString curDir(exePath);

    // PowerShellコマンド
    CString powerShell = _T("powershell");
    CString powerShellCommand;
    powerShellCommand.Format(_T("-Command \"Get-ChildItem '%s\\*.*' -Recurse | Unblock-File\""), (LPCTSTR)exePath);

    // COMコンポーネント登録用オブジェクト
    CCommandExecutor command;

#ifdef _DEBUG

    registSuccess = command.exec(powerShell, powerShellCommand, curDir);

    if (registSuccess == CCommandExecutor::FUNC_SUCCESS) {

        return true;
    }
    else {

        return false;
    }

#else

    registSuccess = command.exec(_T("powershell"), powerShellCommand, curDir);

    if (registSuccess == CCommandExecutor::FUNC_SUCCESS) {

        return true;
    }
    else {

        return false;
    }

#endif
}

void CSutInstallerDlg::processInstall()
{
    // チェック件数
    int checkedCount = 0;

    // リストボックスのアイテム数を取得
    int count = m_installedExcelList.GetItemCount();

    for (int i = 0; i < count; i++) {

        // チェックされている場合
        if (m_installedExcelList.GetCheck(i)) {

            checkedCount++;

            // インストールの流れは以下のとおりとする。
            // (1). セキュリティブロックの解除
            // (2). Excelアドインのインストール

            bool releasedSecutiryBlock = releaseSecurityBlock();

            if (!releasedSecutiryBlock) {

                // エラー発生
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_SECURITY_BLOCK_RELEASE_FAILED);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONWARNING);
            }

            // アドインファイルパスを取得しファイルの存在チェックを実施する            
            CString addinFilePath = getSutAddinPath();

            if (!existFile(addinFilePath)) {

                // 失敗した場合
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_ADDIN_FILE_NOT_FOUND);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONHAND);

                return;
            }

            // Excelが起動している場合
            if (m_excelInfoGetter.existExcelProcess()) {

                // 失敗した場合
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_EXCEL_PROCESS_EXIST);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONEXCLAMATION);

                return;
            }

            // 現在選択されている項目を取得する
            int itemIndex = i;

            // リストにて選択されている項目から
            // アプリケーション名とアプリケーションパスを取得
            TCHAR appName[256];
            m_installedExcelList.GetItemText(itemIndex, 0, appName, 256);

            // Excelのバージョンを取得する
            CString ver;
            m_excelInfoGetter.getExcelVersionByName(appName, ver);

            // Excelアドインがインストール可能かをチェックする
            int isPossibleAddinInstall = m_excelInfoGetter.isPossibleAddinInstall(ver);

            if (isPossibleAddinInstall == CExcelInfoGetter::POSSIBLE_ADDIN_INSTALL_UNEXPECTED) {

                // 予期せぬエラー
                // 失敗した場合
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_CHECK_ADDIN_INSTALL_UNEXPECTED);
                infoMessage.Append(_T(" ver"));
                infoMessage.Append(ver);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONHAND);
                return;
            }
            else if (isPossibleAddinInstall == CExcelInfoGetter::POSSIBLE_ADDIN_INSTALL_NG) {

                // NGの場合
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_CHECK_ADDIN_INSTALL_NG);
                CString infoMessage2;
                infoMessage2.LoadString(IDS_INFO_CHECK_ADDIN_INSTALL_NG_DETAIL);
                infoMessage.Append(_T("\n"));
                infoMessage.Append(infoMessage2);

                // 確認ダイアログを表示する
                int messResult = AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONEXCLAMATION);

                return;
            }

            ExcelAddinRegistry addinReg;
            if (addinReg.installAddin(ver, addinFilePath) != ERROR_SUCCESS) {

                // 予期せぬエラー
                // 失敗した場合
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

    // チェックされた件数が0の場合
    if (checkedCount == 0) {

        CString infoMessage;
        infoMessage.LoadString(IDS_SELECTED_ONE_MORE);
        AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONEXCLAMATION);
    }
}

void CSutInstallerDlg::processUninstall()
{
    // レジストリ削除
    BOOL isRegDelete = m_regDelete.GetCheck();

    ATLTRACE2("Dose user delete registry? %s\n", isRegDelete == TRUE ? "true" : "false");

    // チェック件数
    int checkedCount = 0;

    // リストボックスのアイテム数を取得
    int count = m_installedExcelList.GetItemCount();

    for (int i = 0; i < count; i++) {

        // チェックされている場合
        if (m_installedExcelList.GetCheck(i)) {

            checkedCount++;

            // Excelが起動している場合
            if (m_excelInfoGetter.existExcelProcess()) {

                // 失敗した場合
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_EXCEL_PROCESS_EXIST);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONEXCLAMATION);

                return;
            }

            // 現在選択されている項目を取得する
            int itemIndex = i;

            // アドインファイルパスを取得しファイルの存在チェックを実施する            
            CString addinFilePath = getSutAddinPath();

            // リストにて選択されている項目から
            // アプリケーション名とアプリケーションパスを取得
            TCHAR appName[256];
            m_installedExcelList.GetItemText(itemIndex, 0, appName, 256);

            // Excelのバージョンを取得する
            CString ver;
            m_excelInfoGetter.getExcelVersionByName(appName, ver);

            ExcelAddinRegistry addinReg;
            if (addinReg.uninstallAddin(ver, addinFilePath) != ERROR_SUCCESS) {

                // 予期せぬエラー
                // 失敗した場合
                CString infoMessage;
                infoMessage.LoadString(IDS_INFO_CHECK_ADDIN_INSTALL_UNEXPECTED);
                AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONHAND);
                return;
            }
            if (addinReg.uninstallAddinAtAddinManager(ver, addinFilePath) != ERROR_SUCCESS) {

                // 予期せぬエラー
                // 失敗した場合
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

    // レジストリを削除する
    if (isRegDelete) {
        checkedCount++;

        // 関数の戻り値を格納する
        LONG lResult;

        CString key(LPCTSTR(_T("Software\\ison\\Sut")));
        lResult = AfxGetApp()->DelRegTree(HKEY_CURRENT_USER, key);

        if (lResult != ERROR_SUCCESS && lResult != ERROR_FILE_NOT_FOUND) {

            // エラー発生
            CString infoMessage;
            infoMessage.LoadString(IDS_INFO_REG_DELETE_FAILED);
            AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONERROR);
            return;
        }

        CString infoMessage;
        infoMessage.LoadString(IDS_INFO_REG_DELETE_SUCCESS);
        AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONINFORMATION);
    }

    // チェックされた件数が0の場合
    if (checkedCount == 0) {

        CString infoMessage;
        infoMessage.LoadString(IDS_SELECTED_ONE_MORE_OR_UNINSTALL_OPTION);
        AfxMessageBox((LPCTSTR)infoMessage, MB_OK | MB_ICONEXCLAMATION);
    }

}
