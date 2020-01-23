#pragma once

#include <afxwin.h>

// ジェネリックキャラに対応したstringクラス
typedef std::basic_string<TCHAR> tstring;

class CExcelInfo
{

public:

    /**
     * コンストラクタ。
     */
    CExcelInfo(void) {}

    /**
     * デストラクタ。
     */
    ~CExcelInfo(void) {}

    /**
     * アプリケーション名
     */
    tstring appName;

    /**
     * アプリケーションパス
     */
    tstring appPath;

};

class CExcelInfoGetter
{
public:

    /**
     * コンストラクタ。
     */
    CExcelInfoGetter(void);

    /**
     * デストラクタ。
     */
    ~CExcelInfoGetter(void);

protected:

    /**
     * バージョン別 ExcelのCLSID
     */
    static LPCTSTR CLSID_COMPONENT_EXCEL[];

    /**
     * バージョン別 Excelのアプリケーション名
     */
    static LPCTSTR COMPONENT_EXCEL_NAME[];

    /**
     * CLSIDを格納している配列の長さ
     */
    static const int CLSID_ARRAY_LENGTH;

    /**
     * インストール済みのExcelアプリケーション情報リスト
     */
    std::vector<CExcelInfo*> installedExcelList;

    /**
     * レジストリパス　Excelのセキュリティ設定情報のパスの検索・置換対象文字列
     */
    static LPCTSTR REG_PATH_EXCEL_PARAM_VERSION;

    /**
     * レジストリパス　Excelのセキュリティ設定情報
     */
    static LPCTSTR REG_PATH_EXCEL_SECURITY_SETTING;

    /**
     * レジストリの値の名前 Excelセキュリティ情報 インストールされたアドインを信用しないフラグ
     */
    static LPCTSTR REG_VALUE_NAME_EXCEL_SECURITY_DONTTRUSTINSTALLEDFILES;

    /**
     * レジストリの値の名前 Excelセキュリティ情報 全てのアドインを無効にするフラグ
     */
    static LPCTSTR REG_VALUE_NAME_EXCEL_SECURITY_DISABLEALLADDINS;

    /**
     * レジストリの値の名前 Excelセキュリティ情報 署名済みアドインのみ有効にするフラグ
     */
    static LPCTSTR REG_VALUE_NAME_EXCEL_SECURITY_REQUIREDADDINSIG; 

    /**
     * レジストリパス　Excelのアドイン管理ディレクトリ
     */
    static LPCTSTR REG_PATH_EXCEL_ADDIN_MANAGER;
	static LPCTSTR REG_PATH_EXCEL_ADDIN_MANAGER2;

public:

    static const int POSSIBLE_ADDIN_INSTALL_OK         = 0;
    static const int POSSIBLE_ADDIN_INSTALL_NG         = 1;
    static const int POSSIBLE_ADDIN_INSTALL_UNEXPECTED = 2;

    static const int DEL_ADDIN_OK = 0;
    static const int DEL_ADDIN_TARGET_KEY_NOT_FOUND = 1;
    static const int DEL_ADDIN_UNEXPECTED = 2;
    static const int DEL_ADDIN_SUSPEND = 3;

    static const int ADD_ADDIN_OK = 0;
    static const int ADD_ADDIN_TARGET_KEY_NOT_FOUND = 1;
    static const int ADD_ADDIN_UNEXPECTED = 2;

    /**
     * インストール済みのExcelアプリケーション情報リストを取得する。
     *
     * @return インストール済みのExcelアプリケーション情報リスト（this->installedExcelListを返す）
     */
    std::vector<CExcelInfo*>& getInstalledExcelApplication();

    /**
     * Excelプロセスが存在するかをチェックする。
     *
     * @return true プロセスが存在する
     */
    bool existExcelProcess();

    /**
     * アドインのインストールが可能かをチェックする。
     *
     * @param excelVersion Excelのバージョン
     * @return POSSIBLE_ADDIN_INSTALL_OK インストール可能
     *         POSSIBLE_ADDIN_INSTALL_NG インストール不可
     *         POSSIBLE_ADDIN_INSTALL_UNEXPECTED 予期せぬエラー
     */
    int isPossibleAddinInstall(CString excelVersion);

    /**
     * アドインのインストールが可能かをチェックする。
     * Excel2000以降
     *
     * @param excelVersion Excelのバージョン
     * @return POSSIBLE_ADDIN_INSTALL_OK インストール可能
     *         POSSIBLE_ADDIN_INSTALL_NG インストール不可
     *         POSSIBLE_ADDIN_INSTALL_UNEXPECTED 予期せぬエラー
     */
    int isPossibleAddinInstallForOverExcel2000(CString excelVersion);

    /**
     * アドインのインストールが可能かをチェックする。
     * Excel2007以降
     *
     * @param excelVersion Excelのバージョン
     * @return POSSIBLE_ADDIN_INSTALL_OK インストール可能
     *         POSSIBLE_ADDIN_INSTALL_NG インストール不可
     *         POSSIBLE_ADDIN_INSTALL_UNEXPECTED 予期せぬエラー
     */
    int isPossibleAddinInstallForOverExcel2007(CString excelVersion);

    /**
     * アドインを完全に削除する。
     *
     * アドインは以下のレジストリパスにファイルパスが保存されており、以下からエントリを削除することで完全に削除することが可能。
     *
     * HKEY_CURRENT_USER\Software\Microsoft\Office\バージョン毎\Excel\Options
     * HKEY_CURRENT_USER\Software\Microsoft\Office\バージョン毎\Excel\Add-in Manager
     *
     * 前者は、Excelでアドインがオンの状態でのみエントリが存在する。
	 * 後者は、Excelでアドインがオフの状態でのみエントリが存在する。
	 * 両者ともに、排他的である。
     *
     * @param excelVersion Excelのバージョン
     * @param addinFileName アドインファイル名
     */
    int delAddin(CString excelVersion, CString addinFileName);

	/**
	 * アドインを追加する。
     *
     * アドインを追加するには、以下のレジストリパスに"OPEN"という接頭辞でエントリを作成する。
	 * 接頭辞がOPENであれば、続く文字列は何でも良い。
     *
     * HKEY_CURRENT_USER\Software\Microsoft\Office\バージョン毎\Excel\Options
     *
     * 本メソッドは上述したパスにOPENエントリを追加する。
     *
	 * @param excelVersion Excelのバージョン
	 * @param addinFileName アドインファイル名
	 */
	int addAddin(CString excelVersion, CString addinFileName);

	/**
	 * バージョンを取得する。
	 * @param 製品名 Excel2000やExcel2003等
	 * @return バージョン番号 Excel2000なら9.0
	 */
	CString getVersion(CString& productName);

    /**
     * ！！！未作成
     */
    void searchRunningObjectTable();

};
