#include "pch.h"
#include "CommandExecutor.h"

CCommandExecutor::CCommandExecutor(void)
{
}

CCommandExecutor::~CCommandExecutor(void)
{
}

int CCommandExecutor::exec(CString command, CString option, CString curDir)
{

    // システムディレクトリを取得する
    TCHAR systemDir[_MAX_PATH];
    if (!GetSystemDirectory(systemDir, _MAX_PATH)) {

        return FUNC_FAILED;
    }

    STARTUPINFO Start;
    ZeroMemory(&Start,sizeof(STARTUPINFO));
    Start.cb = sizeof(Start);
    Start.dwFlags = STARTF_USESHOWWINDOW;
    Start.wShowWindow = SW_SHOWMINIMIZED;

    PROCESS_INFORMATION ProcInfo;

    // regsvr32.exeを実行する
    BOOL ret = CreateProcess(NULL
                            , (LPTSTR)((LPCTSTR)(command + _T(" ") + option))
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
