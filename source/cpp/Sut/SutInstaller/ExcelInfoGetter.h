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

	// 
	static std::vector<std::map<tstring, tstring>> EXCEL_INFO_LIST;

    /**
     * バージョン別 ExcelのCLSID
     */
    static LPCTSTR CLSID_COMPONENT_EXCEL[];

    /**
     * バージョン別 Excelのアプリケーション名
     */
    static LPCTSTR COMPONENT_EXCEL_NAME[];

    /**
     * Excelのバージョン
     */
    static LPCTSTR EXCEL_VERSION[];

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

public:

    static const int POSSIBLE_ADDIN_INSTALL_OK         = 0;
    static const int POSSIBLE_ADDIN_INSTALL_NG         = 1;
    static const int POSSIBLE_ADDIN_INSTALL_UNEXPECTED = 2;

    static const int COMPLETELY_DELETE_ADDIN_OK = 0;
    static const int COMPLETELY_DELETE_ADDIN_TARGET_KEY_NOT_FOUND = 1;
    static const int COMPLETELY_DELETE_ADDIN_UNEXPECTED = 2;
    static const int COMPLETELY_DELETE_ADDIN_SUSPEND = 3;

    /**
     * インストール済みのExcelアプリケーション情報リストを取得する。
     *
     * @return インストール済みのExcelアプリケーション情報リスト（this->installedExcelListを返す）
     */
    std::vector<CExcelInfo*>& getInstalledExcelApplication();

	/**
	 * エクセル名からバージョンを取得する
	 */
	void getExcelVersionByName(const CString& excelName, CString& excelVersion);

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
     * アドインは以下のレジストリパスにファイルパスが保存されており、以下からパスを削除しないと
     * 同名のアドインを再インストールしても正しく有効にならない。
     *
     * HKEY_CURRENT_USER\Software\Microsoft\Office\バージョン毎\Excel\Add-in Manager
     *
     * 本メソッドは上述したパスからaddinFileNameにマッチした情報を削除する。
     * また、上述したパスはAddin.InstallプロパティをFalseに設定し、Excel終了後にはじめて書き込まれる。
     * 従って、それ以前に削除しようとしても存在しない。
     *
     * @param excelVersion Excelのバージョン
     * @param addinFileName アドインファイル名
     */
    int completelyDeleteAddin(CString excelVersion, CString addinFileName);

    /**
     * ！！！未作成
     */
    void searchRunningObjectTable();

};
