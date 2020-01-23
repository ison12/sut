#pragma once

class CExcelDispatchWrapper
{
public:

    /**
     * コンストラクタ。
     */
    CExcelDispatchWrapper(IDispatch* applicationDisp);

    /**
     * デストラクタ。
     */
    ~CExcelDispatchWrapper(void);

protected:

    /**
     * Excel.Applicationオブジェクト
     */
    COleDispatchDriver excelApplication;

public:

    /**
     * 戻り値コード 成功
     */
    static const int SUCCESS = 0;

    /**
     * 戻り値コード 失敗（DISPIDの取得に失敗）
     */
    static const int ERROR_DISPID_NOT_FOUND = 999;

    /**
     * 戻り値コード 失敗（予期せぬエラー）
     */
    static const int ERROR_UNEXPECTED = 1000;

    /**
     * 戻り値コード 失敗（項目が見つからない）
     */
    static const int ERROR_ITEM_NOT_FOUND = 2000;

    /**
     * 戻り値コード 失敗（項目が見つかった）
     */
    static const int ERROR_ITEM_EXIST = 2001;

    /**
     * オブジェクトからExcel.Applicationを取り外す。
     *
     * @return Excel.Application IDispatch
     */
    IDispatch* detachIDispatch();

    /**
     * バージョンを取得する。
     *
     * @return Excelバージョン
     */
    CString getVersion();

    /**
     * Excelアプリの警告表示有無を変更する。
     *
     * @param b TRUEの場合、警告を表示する
     * @return SUCCESS 成功
     */
    int displayAlerts(BOOL b);

    /**
     * Excelアプリの表示ステータスを変更する。
     *
     * @param visible TRUEの場合、Excelアプリを表示する
     * @return SUCCESS 成功
     */
    int appVisible(BOOL visible);

    /**
     * Excelアプリを終了する。
     *
     * @return SUCCESS 成功
     */
    int appQuit();

    /**
     * Addinsオブジェクトを取得する。
     *
     * @return Addinsオブジェクト
     */
    IDispatch* getAddinsObject();

    /**
     * アドインを追加する。
     *
     * @param addinPath アドインパス
     * @return SUCCESS 成功
     */
    int attachAddin(CString addinPath);

    /**
     * アドインを削除する。
     *
     * @param addinPath アドインパス
     * @return SUCCESS 成功
     */
    int removeAddin(CString addinPath);

};
