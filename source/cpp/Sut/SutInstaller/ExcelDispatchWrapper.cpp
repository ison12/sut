#include "pch.h"
#include "Resource.h"
#include "ExcelDispatchWrapper.h"

CExcelDispatchWrapper::CExcelDispatchWrapper(IDispatch* applicationDisp) : excelApplication(applicationDisp)
{
}

CExcelDispatchWrapper::~CExcelDispatchWrapper(void)
{
}

IDispatch* CExcelDispatchWrapper::detachIDispatch()
{
    // ディスパッチインターフェースを取得
    IDispatch* ret = excelApplication.m_lpDispatch;
    // 解放する
    excelApplication.ReleaseDispatch();

    return ret;
}

CString CExcelDispatchWrapper::getVersion()
{
    try {

        // プロパティ・メソッドの実行結果
        HRESULT result;

        // 戻り値
        VARIANT vResult;
        VariantInit(&vResult);

        // Invoke名
        LPOLESTR name = L"Version";

        // InvokeするディスパッチID
        DISPID dispid;

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            // 通常有り得ない
            // 空文字列を返す
            return CString();
        }

        // バージョン文字列を取得する
        excelApplication.GetProperty(dispid, VT_VARIANT, (void*)&vResult);

        // 戻り値がBSTRの場合
        if (vResult.vt & VT_BSTR) {

            return vResult.bstrVal;
        
        // 戻り値がBSTR以外の場合
        } else {

            // 通常有り得ない
            // 空文字列を返す
            return CString();
        }

    }  // End try.

    catch(COleException *e)
    {

        // エラーメッセージ取得
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // メッセージを表示する
        AfxMessageBox(errorMess, MB_ICONHAND);

        // 通常有り得ない
        // 空文字列を返す
        return CString();
    }

    catch(COleDispatchException *e)
    {

        // エラーメッセージ取得
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // メッセージを表示する
        AfxMessageBox(errorMess, MB_ICONHAND);

        // 通常有り得ない
        // 空文字列を返す
        return CString();
    }
    catch(...)
    {
        // メッセージを表示する
        AfxMessageBox(_T("Unexpected error"), MB_ICONHAND);

        // 通常有り得ない
        // 空文字列を返す
        return CString();
    }

}

int CExcelDispatchWrapper::displayAlerts(BOOL b)
{

    try {

        // プロパティ・メソッドの実行結果
        HRESULT result;

        // 戻り値
        VARIANT vResult;
        VariantInit(&vResult);

        // Invoke名
        LPOLESTR name = L"DisplayAlerts";

        // InvokeするディスパッチID
        DISPID dispid;

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            // 通常有り得ない
            return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
        }

        // Application.DisplayAlertsを設定
        excelApplication.SetProperty(dispid, VT_BOOL, b);

    }  // End try.

    catch(COleException *e)
    {

        // エラーメッセージ取得
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // メッセージを表示する
        AfxMessageBox(errorMess, MB_ICONHAND);

        // 通常有り得ない
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    catch(COleDispatchException *e)
    {

        // エラーメッセージ取得
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // メッセージを表示する
        AfxMessageBox(errorMess, MB_ICONHAND);

        // 通常有り得ない
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }
    catch(...)
    {
        // メッセージを表示する
        AfxMessageBox(_T("Unexpected error"), MB_ICONHAND);

        // 通常有り得ない
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    return CExcelDispatchWrapper::SUCCESS;
}

int CExcelDispatchWrapper::appVisible(BOOL visible)
{
    try {

        // プロパティ・メソッドの実行結果
        HRESULT result;

        // 戻り値
        VARIANT vResult;
        VariantInit(&vResult);

        // Invoke名
        LPOLESTR name = L"Visible";

        // InvokeするディスパッチID
        DISPID dispid;

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            // 通常有り得ない
            return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
        }

        // Application.Visibleを設定
        excelApplication.SetProperty(dispid, VT_BOOL, visible);

    }  // End try.

    catch(COleException *e)
    {
        // エラーメッセージ取得
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // メッセージを表示する
        AfxMessageBox(errorMess, MB_ICONHAND);


        // 通常有り得ない
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    catch(COleDispatchException *e)
    {
        // エラーメッセージ取得
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // メッセージを表示する
        AfxMessageBox(errorMess, MB_ICONHAND);


        // 通常有り得ない
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }
    catch(...)
    {
        // メッセージを表示する
        AfxMessageBox(_T("Unexpected error"), MB_ICONHAND);

        // 通常有り得ない
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    return CExcelDispatchWrapper::SUCCESS;
}

int CExcelDispatchWrapper::appQuit()
{

    try {

        // プロパティ・メソッドの実行結果
        HRESULT result;

        // 戻り値
        VARIANT vResult;
        VariantInit(&vResult);

        // Invoke名
        LPOLESTR name = L"Quit";

        // InvokeするディスパッチID
        DISPID dispid;

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            // 通常有り得ない
            return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
        }

        // Application.Visibleを設定
        excelApplication.InvokeHelper(dispid, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);

    }  // End try.

    catch(COleException *e)
    {
        // エラーメッセージ取得
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // メッセージを表示する
        AfxMessageBox(errorMess, MB_ICONHAND);

        // 通常有り得ない
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    catch(COleDispatchException *e)
    {
        // エラーメッセージ取得
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // メッセージを表示する
        AfxMessageBox(errorMess, MB_ICONHAND);

        // 通常有り得ない
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }
    catch(...)
    {
        // メッセージを表示する
        AfxMessageBox(_T("Unexpected error"), MB_ICONHAND);

        // 通常有り得ない
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    return CExcelDispatchWrapper::SUCCESS;
}

IDispatch* CExcelDispatchWrapper::getAddinsObject()
{

    // プロパティ・メソッドの実行結果
    HRESULT result;

    // InvokeするディスパッチID
    DISPID dispid;

    // IDispatchオブジェクト
    IDispatch* pDisp = NULL;

    // -------------------------------------------------------------------
    // Application.Workbooksオブジェクトの取得
    // Workbooksオブジェクト
    COleDispatchDriver workbooks;
    {
        // Invoke名
        LPOLESTR name = L"Workbooks";

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            return NULL;
        }

        // 戻り値を初期化
        pDisp = NULL;

        // Addinsオブジェクトを取得する
        excelApplication.GetProperty(dispid, VT_DISPATCH, (void*)&pDisp);

        if (pDisp != NULL) {

            workbooks.AttachDispatch(pDisp);
        
        } else {

            return NULL;
        }
    }

    // --------------------------------------------------------------------------------------
    // ActiveなWorkbookの存在を判定し、あればWorkbookを取得し、なければ新たに追加する
    // ※ActiveWorkbookが存在しない場合、Addinsに対する操作が失敗することがある
    // -------------------------------------------------------------------
    // Workbookオブジェクト
    COleDispatchDriver workbook;
    // 存在有無
    bool isActiveWorkbook = false;

    // Application.ActiveWorkbookのチェック
    {
        // Invoke名
        LPOLESTR name = L"ActiveWorkbook";

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            return NULL;
        }

        // 戻り値を初期化
        pDisp = NULL;

        // Workbookを追加しオブジェクトを取得する
        excelApplication.GetProperty(dispid, VT_DISPATCH, (void*)&pDisp);

        if (pDisp != NULL) {

            workbook.AttachDispatch(pDisp);
            isActiveWorkbook = true;
        
        }
    }


    // -------------------------------------------------------------------
    // Workbooks.Addオブジェクトの取得
    if (!isActiveWorkbook) {

        // Invoke名
        LPOLESTR name = L"Add";

        result = workbooks.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            return NULL;
        }

        // 戻り値を初期化
        pDisp = NULL;

        static BYTE paramInfo[] = VTS_VARIANT;

        // パラメータ
        VARIANT paramVar;
        VariantInit(&paramVar);
        paramVar.vt = VT_ERROR;
        paramVar.scode = DISP_E_PARAMNOTFOUND;

        // Workbookを追加しオブジェクトを取得する
        workbooks.InvokeHelper(dispid, DISPATCH_METHOD, VT_DISPATCH, (void*)&pDisp, paramInfo, &paramVar);

        if (pDisp != NULL) {

            workbook.AttachDispatch(pDisp);

            // Applicationオブジェクトを取得
            // Visibleを実行する
        
        } else {

            return NULL;
        }
    }

    // -------------------------------------------------------------------
    // Application.AddInsオブジェクトの取得
    {
        // Invoke名
        LPOLESTR name = L"AddIns";

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            return NULL;
        }

        // 戻り値を初期化
        pDisp = NULL;

        // Addinsオブジェクトを取得する
        excelApplication.GetProperty(dispid, VT_DISPATCH, (void*)&pDisp);

    }

    return pDisp;
}

int CExcelDispatchWrapper::attachAddin(CString addinPath)
{

    try {
        // プロパティ・メソッドの実行結果
        HRESULT result;

        // InvokeするディスパッチID
        DISPID dispid;

        // IDispatchオブジェクト
        IDispatch* pDisp = NULL;

        // -------------------------------------------------------------------
        // Application.AddInsオブジェクトの取得
        // Addinsオブジェクト
        COleDispatchDriver addins;
        pDisp = getAddinsObject();

        if (pDisp != NULL) {

            addins.AttachDispatch(pDisp);
        
        } else {

            return CExcelDispatchWrapper::ERROR_UNEXPECTED;
        }


        // アドインが既にインストールされているかを確認する
        bool isAddinInstalled = false;

        // Addinオブジェクト
        COleDispatchDriver addin;

        // -------------------------------------------------------------------
        // Application.AddInsオブジェクトのItemメソッドの実行
        {
            // Invoke名
            LPOLESTR name = L"Item";

            result = addins.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
            if (FAILED(result)) {

                return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
            }

            // 戻り値を初期化
            pDisp = NULL;

            static BYTE paramInfo[] = VTS_VARIANT;

            // アドイン名を使ってアドインオブジェクトを検索する
            CString addinName;
            addinName.LoadString(IDS_ADDIN_NAME);

            BSTR paramStr = addinName.AllocSysString();

            VARIANT paramVar;
            VariantInit(&paramVar);
            paramVar.vt = VT_BSTR;
            paramVar.bstrVal = paramStr;

            try {

                // アドインを追加する
                addins.InvokeHelper(dispid, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&pDisp, paramInfo, &paramVar);

            }
            catch(COleDispatchException *e)
            {
                // -2147352565 (8002000B)    Invalid index.
                if (-2147352565 == e->m_scError) {

                    isAddinInstalled = false;
                } else {

                    throw e;
                }
            }

            SysFreeString(paramStr);

            if (pDisp != NULL) {

                isAddinInstalled = true;
                addin.AttachDispatch(pDisp);
            
            }
        }

        // アドインが既にインストールされている場合
        if (isAddinInstalled) {

            return ERROR_ITEM_EXIST;
        }

        // -------------------------------------------------------------------
        // Application.AddInsオブジェクトのAddメソッドの実行
        {
            // Invoke名
            LPOLESTR name = L"Add";

            result = addins.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
            if (FAILED(result)) {

                return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
            }

            // 戻り値を初期化
            pDisp = NULL;

            static BYTE paramInfo[] = VTS_BSTR VTS_VARIANT;

            BSTR paramStr = addinPath.AllocSysString();

            VARIANT paramVar;
            VariantInit(&paramVar);
            paramVar.vt = VT_BOOL;
            paramVar.boolVal = FALSE;

            // アドインを追加する
            addins.InvokeHelper(dispid, DISPATCH_METHOD, VT_DISPATCH, (void*)&pDisp, paramInfo, paramStr, &paramVar);

            SysFreeString(paramStr);

            if (pDisp != NULL) {

                addin.AttachDispatch(pDisp);
            
            } else {

                return CExcelDispatchWrapper::ERROR_UNEXPECTED;
            }
        }

        // -------------------------------------------------------------------
        // AddinオブジェクトのInstalledプロパティの変更
        {
            // Invoke名
            LPOLESTR name = L"Installed";

            result = addin.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
            if (FAILED(result)) {

                return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
            }

            // アドインを追加する
            addin.SetProperty(dispid, VT_BOOL, TRUE);

        }


    }  // End try.

    catch(COleException *e)
    {

        // エラーメッセージ取得
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // メッセージを表示する
        AfxMessageBox(errorMess, MB_ICONHAND);

        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    catch(COleDispatchException *e)
    {

        // エラーメッセージ取得
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // メッセージを表示する
        AfxMessageBox(errorMess, MB_ICONHAND);

        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }
    catch(...)
    {
        // メッセージを表示する
        AfxMessageBox(_T("Unexpected error"), MB_ICONHAND);

        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    return CExcelDispatchWrapper::SUCCESS;
}

int CExcelDispatchWrapper::removeAddin(CString addinPath)
{

    try {
        // プロパティ・メソッドの実行結果
        HRESULT result;

        // InvokeするディスパッチID
        DISPID dispid;

        // IDispatchオブジェクト
        IDispatch* pDisp = NULL;

        // -------------------------------------------------------------------
        // Application.AddInsオブジェクトの取得
        // Addinsオブジェクト
        COleDispatchDriver addins;
        pDisp = getAddinsObject();

        if (pDisp != NULL) {

            addins.AttachDispatch(pDisp);
        
        } else {

            return CExcelDispatchWrapper::ERROR_UNEXPECTED;
        }

        // -------------------------------------------------------------------
        // Application.AddInsオブジェクトのItemメソッドの実行
        // Addinオブジェクト
        COleDispatchDriver addin;
        {
            // Invoke名
            LPOLESTR name = L"Item";

            result = addins.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
            if (FAILED(result)) {

                return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
            }

            // 戻り値を初期化
            pDisp = NULL;

            static BYTE paramInfo[] = VTS_VARIANT;

            BSTR paramStr = addinPath.AllocSysString();

            VARIANT paramVar;
            VariantInit(&paramVar);
            paramVar.vt = VT_BSTR;
            paramVar.bstrVal = paramStr;

            // アドインを追加する
            addins.InvokeHelper(dispid, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&pDisp, paramInfo, &paramVar);

            SysFreeString(paramStr);

            if (pDisp != NULL) {

                addin.AttachDispatch(pDisp);
            
            } else {

                // 削除対象が見つからない場合
                return CExcelDispatchWrapper::ERROR_ITEM_NOT_FOUND;
            }

        }

        // -------------------------------------------------------------------
        // AddinオブジェクトのInstalledプロパティの変更
        {
            // Invoke名
            LPOLESTR name = L"Installed";

            result = addin.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
            if (FAILED(result)) {

                return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
            }

            // アドインを追加する
            addin.SetProperty(dispid, VT_BOOL, FALSE);

        }


    }  // End try.

    catch(COleException *e)
    {
        // エラーメッセージ取得
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // メッセージを表示する
        AfxMessageBox(errorMess, MB_ICONHAND);

        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    catch(COleDispatchException *e)
    {
        // -2147352565 (8002000B)    Invalid index.
        if (-2147352565 == e->m_scError) {

            return CExcelDispatchWrapper::ERROR_ITEM_NOT_FOUND;
        }

        // エラーメッセージ取得
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // メッセージを表示する
        AfxMessageBox(errorMess, MB_ICONHAND);

        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }
    catch(...)
    {
        // メッセージを表示する
        AfxMessageBox(_T("Unexpected error"), MB_ICONHAND);

        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    return CExcelDispatchWrapper::SUCCESS;
}
