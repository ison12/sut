#include "StdAfx.h"
#include "ExcelStartup.h"

CExcelStartup::CExcelStartup(CString path) : excelPath(path)
{
}

CExcelStartup::~CExcelStartup(void)
{
}

IDispatch* CExcelStartup::startUp()
{

    // Store the handle of the currently active window...
    HWND hwndCurrent = ::GetForegroundWindow();

    // Launch Excel and wait until it is waiting for
    // user input...

    STARTUPINFO Start;
    ZeroMemory(&Start,sizeof(STARTUPINFO));
    Start.cb=sizeof(Start);
    Start.dwFlags = STARTF_USESHOWWINDOW;
    Start.wShowWindow = SW_SHOWMINIMIZED;

    PROCESS_INFORMATION ProcInfo;

    // プロセスを生成する
    CreateProcess((LPCTSTR)excelPath
                , _T(" /e ") // /s = Xlstart および Xlstart 代替フォルダのファイルを一切開かず
                                // に起動します。また、ツール バー ファイル (Excel.xlb または 
                                // <ユーザー名>.xlb) を読み込まずに起動します。Excel のタイ
                                // トルバーには safe mode と表示されます。
                                // ※/sを有効にするとメニューバーが読み込まれなくなるので、それによりメニューの追加ができなくなるのでよろしくない
                                // /e = 起動画面の表示および新規ブック (Book1.xls) の作成を行わずに Excel を起動します。
                , 0
                , 0
                , 1
                , NORMAL_PRIORITY_CLASS
                , 0
                , NULL
                , &Start
                , &ProcInfo);

    if (WaitForInputIdle(ProcInfo.hProcess, 10000)==WAIT_TIMEOUT) {

        // デバッグ用メッセージ
        // リリース後も継続して出力する
        AfxMessageBox(_T("Timed out waiting for Excel."), MB_OK | MB_ICONINFORMATION);
    }

    // Restore the active window to the foreground...
    //  NOTE: If you comment out this line, the code will fail!
    SetForegroundWindow(hwndCurrent);

    // ----------------------------------------------------------------
    // 以下のコードを実行する前に
    // CoInitializeを予め実行しておく

    // Attach to the running instance...
    CLSID clsid;
    CLSIDFromProgID(L"Excel.Application", &clsid);

    IUnknown *pUnk = NULL;
    IDispatch *pDisp = NULL;

    for (int i = 1; i <= 5; i++) { // try attaching for up to 5 attempts
    
        HRESULT hr = GetActiveObject(clsid, NULL, (IUnknown**)&pUnk);

        if (SUCCEEDED(hr)) {

            hr = pUnk->QueryInterface(IID_IDispatch, (void **)&pDisp);
            break;
        }

        Sleep(1000);
    }
            
    //Release the no-longer-needed IUnknown...
    if (pUnk) 
        pUnk->Release();

    return pDisp;
}
