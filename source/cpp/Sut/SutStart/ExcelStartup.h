#pragma once

// ジェネリックキャラに対応したstringクラス
typedef std::basic_string<TCHAR> tstring;

class CExcelStartup
{

public:

    /**
     * コンストラクタ。
     *
     * @param path パス
     */
    CExcelStartup(CString path);

    /**
     * デストラクタ。
     */
    ~CExcelStartup(void);

    /**
     * Excelを起動する。
     *
     * @return Excelアプリケーションのオートメーションオブジェクト
     */
    IDispatch* startUp();

protected:

    /**
     * Excelアプリケーションのパス
     */
    CString excelPath;

};
