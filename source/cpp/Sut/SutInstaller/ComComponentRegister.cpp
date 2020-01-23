#include "pch.h"
#include "ComComponentRegister.h"

// static変数の初期化
LPCTSTR CComComponentRegister::REG_SVR_EXE = _T("regsvr32.exe");

CComComponentRegister::CComComponentRegister(void)
{
}

CComComponentRegister::~CComComponentRegister(void)
{
}

int CComComponentRegister::execRegSvr(CString option, CString filePath, CString curDir)
{

    CString commandLine;
    commandLine.Append(option);    // オプション

    commandLine.Append(_T("\""));   // ターゲットファイルの両端を(")で閉じる
    commandLine.Append(filePath);   // ターゲットファイル
    commandLine.Append(_T("\""));

    // システムディレクトリを取得する
    TCHAR systemDir[_MAX_PATH];
    if (!GetSystemDirectory(systemDir, _MAX_PATH)) {

        return FUNC_FAILED;
    }

    // システムディレクトリとregsvr32.exeを結合する（regsvr32.exeのフルパスを取得する）
    CString regsvrPath;
    regsvrPath.Append((LPCTSTR)systemDir);
    regsvrPath.Append(_T("\\"));
    regsvrPath.Append(REG_SVR_EXE);

    STARTUPINFO Start;
    ZeroMemory(&Start,sizeof(STARTUPINFO));
    Start.cb = sizeof(Start);
    Start.dwFlags = STARTF_USESHOWWINDOW;
    Start.wShowWindow = SW_SHOWMINIMIZED;

    PROCESS_INFORMATION ProcInfo;

    // regsvr32.exeを実行する
    BOOL ret = CreateProcess((LPCTSTR)regsvrPath
                            , (LPTSTR)((LPCTSTR)commandLine)
                            , 0
                            , 0
                            , 1
                            , NORMAL_PRIORITY_CLASS
                            , 0
                            , (LPTSTR)((LPCTSTR)curDir)
                            , &Start
                            , &ProcInfo);

    // CreateProcessに失敗
    if (!ret) {

        return FUNC_FAILED;
    }

    // スレッドハンドルを閉じる
    CloseHandle(ProcInfo.hThread);

    // プロセスが終了するまで待機する
    WaitForSingleObject(ProcInfo.hProcess, INFINITE);

    // 終了コードを取得
    DWORD dwExCode;
    GetExitCodeProcess(ProcInfo.hProcess, &dwExCode);

    // プロセスハンドルを閉じる
    CloseHandle(ProcInfo.hProcess);

    if (dwExCode == 0) {

        return FUNC_SUCCESS;
    } else {

        return dwExCode;
    }
}

int CComComponentRegister::regist(CString filePath, CString curDir)
{

    // オプション文字列を設定
    CString option(_T(" /s ")); // サイレントモード

    // regsvr32を実行する
    int ret = execRegSvr(option, filePath, curDir);

    return ret;
}

int CComComponentRegister::unregist(CString filePath, CString curDir)
{

    // オプション文字列を設定
    CString option(_T(" /s /u ")); // サイレントモードとアンインストール

    // regsvr32を実行する
    int ret = execRegSvr(option, filePath, curDir);

    return ret;
}
