#include "pch.h"
#include "ComComponentRegister.h"

// static�ϐ��̏�����
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
    commandLine.Append(option);    // �I�v�V����

    commandLine.Append(_T("\""));   // �^�[�Q�b�g�t�@�C���̗��[��(")�ŕ���
    commandLine.Append(filePath);   // �^�[�Q�b�g�t�@�C��
    commandLine.Append(_T("\""));

    // �V�X�e���f�B���N�g�����擾����
    TCHAR systemDir[_MAX_PATH];
    if (!GetSystemDirectory(systemDir, _MAX_PATH)) {

        return FUNC_FAILED;
    }

    // �V�X�e���f�B���N�g����regsvr32.exe����������iregsvr32.exe�̃t���p�X���擾����j
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

    // regsvr32.exe�����s����
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

int CComComponentRegister::regist(CString filePath, CString curDir)
{

    // �I�v�V�����������ݒ�
    CString option(_T(" /s ")); // �T�C�����g���[�h

    // regsvr32�����s����
    int ret = execRegSvr(option, filePath, curDir);

    return ret;
}

int CComComponentRegister::unregist(CString filePath, CString curDir)
{

    // �I�v�V�����������ݒ�
    CString option(_T(" /s /u ")); // �T�C�����g���[�h�ƃA���C���X�g�[��

    // regsvr32�����s����
    int ret = execRegSvr(option, filePath, curDir);

    return ret;
}
