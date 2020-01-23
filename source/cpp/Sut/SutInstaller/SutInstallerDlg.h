
// SutInstallerDlg.h : ヘッダー ファイル
//

#pragma once

#include "ExcelInfoGetter.h"
#include "ExcelStartup.h"
#include "ExcelDispatchWrapper.h"
#include "ComComponentRegister.h"
#include "CommandExecutor.h"


// CSutInstallerDlg ダイアログ
class CSutInstallerDlg : public CDialogEx
{
// コンストラクション
public:
	CSutInstallerDlg(CWnd* pParent = nullptr);	// 標準コンストラクター

// ダイアログ データ
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_SUTINSTALLER_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV サポート


// 実装
protected:
	HICON m_hIcon;

	// 生成された、メッセージ割り当て関数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
    afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
    afx_msg void OnBnClickedInstall();
	afx_msg void OnBnClickedUninstall();

protected:

    /**
     * インストール済みExcelアプリケーションリスト
     */
    CListCtrl m_installedExcelList;

    /**
     * COMコンポーネント削除チェックボックス
     */
    CButton m_comDelete;

    /**
     * レジストリ削除チェックボックス
     */
    CButton m_regDelete;

    /**
     * インストールボタン
     */
    CButton m_install;

    /**
     * アンインストールボタン
     */
    CButton m_uninstall;

    /**
     * Excelの情報を取得するオブジェクト
     */
    CExcelInfoGetter m_excelInfoGetter;

    /**
     * インストール済みExcelアプリケーションリストコントロールを初期化する処理
     */
    void initInstalledExcelList();

    /**
     * Excelを起動する。
     *
     * @return Excel.Applicationオブジェクト
     */
    IDispatch* launchExcelBySelectedListItem(int itemIndex);

    /**
     * Sut.xlamのパスを取得する。
     *
     * @return Sut.xlamのパス
     */
    CString getSutAddinPath();

    /**
     * ファイルが存在するかを確認する。
     *
     * @param filePath ファイルパス
     * @return true ファイルが存在する
     */
    bool existFile(CString filePath);

    CString getExeFilePath();

    /**
     * セキュリティブロックの解除。
     */
    bool releaseSecurityBlock();

    /**
     * インストール処理。
     */
    void processInstall();

    /**
     * アンインストール処理。
     */
    void processUninstall();

};
