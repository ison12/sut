#pragma once

class CComComponentRegister
{
public:

    static const int FUNC_SUCCESS = 0;
    static const int FUNC_FAILED  = 1;

    /**
     * コンストラクタ。
     */
    CComComponentRegister(void);

    /**
     * デストラクタ。
     */
    virtual ~CComComponentRegister(void);

    /**
     * Comコンポーネントの登録。
     *
     * @param filePath ファイルパス
     */
    int regist(CString filePath, CString curDir);

    /**
     * Comコンポーネントの登録解除。
     *
     * @param filePath ファイルパス
     */
    int unregist(CString filePath, CString curDir);

protected:

    /**
     * regsvrの実行ファイル名
     */
    static LPCTSTR REG_SVR_EXE;

    /**
     * regsvr32.exeの実行（COMを登録するための実行ファイル）
     *
     * @param option オプション
     * @param filePath ファイルパス
     */
    int execRegSvr(CString option, CString filePath, CString curDir);

};
