#pragma once

class ExcelAddinRegistry
{
public:

	/**
	 * コンストラクタ。
	 */
	ExcelAddinRegistry(void);

	/**
	 * デストラクタ。
	 */
	virtual ~ExcelAddinRegistry(void);

private:

    /**
     * レジストリパス　Excelのセキュリティ設定情報のパスの検索・置換対象文字列
     */
    static LPCTSTR REG_PATH_EXCEL_PARAM_VERSION;

    /**
     * レジストリパス　Excelのアドイン管理ディレクトリ
     */
    static LPCTSTR REG_PATH_EXCEL_ADDIN_MANAGER;

    /**
     * レジストリパス　Excelのアドインオプションディレクトリ
     */
    static LPCTSTR REG_PATH_EXCEL_ADDIN_OPTIONS;

public:

	/**
	 * Excelアドインを追加する。
	 */
	int installAddin(CString& version, CString& addinFilePath);

	/**
	 * Excelアドインを削除する。
	 */
	int uninstallAddin(CString& version, CString& addinFilePath);

	/**
	 * Excelアドインを削除する。
	 */
	int uninstallAddinAtAddinManager(CString& version, CString& addinFilePath);

};
