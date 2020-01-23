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

    // �V�X�e���f�B���N�g�����擾����
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

    // regsvr32.exe�����s����
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

    // CreateProcess�Ɏ��s
    if (!ret) {

        return FUNC_FAILED;
    }

    // �X���b�h�n���h�������
    CloseHandle(ProcInfo.hThread);

    // �v���Z�X���I������܂őҋ@����
    WaitForSingleObject(ProcInfo.hProcess, INFINITE);

    // �I���R�[�h���擾
    DWORD dwExCode;
    GetExitCodeProcess(ProcInfo.hProcess, &dwExCode);

    // �v���Z�X�n���h�������
    CloseHandle(ProcInfo.hProcess);

    if (dwExCode == 0) {

        return FUNC_SUCCESS;
    } else {

        return dwExCode;
    }
}
