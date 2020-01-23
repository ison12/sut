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

    // �v���Z�X�𐶐�����
    CreateProcess((LPCTSTR)excelPath
                , _T(" /e ") // /s = Xlstart ����� Xlstart ��փt�H���_�̃t�@�C������؊J����
                                // �ɋN�����܂��B�܂��A�c�[�� �o�[ �t�@�C�� (Excel.xlb �܂��� 
                                // <���[�U�[��>.xlb) ��ǂݍ��܂��ɋN�����܂��BExcel �̃^�C
                                // �g���o�[�ɂ� safe mode �ƕ\������܂��B
                                // ��/s��L���ɂ���ƃ��j���[�o�[���ǂݍ��܂�Ȃ��Ȃ�̂ŁA����ɂ�胁�j���[�̒ǉ����ł��Ȃ��Ȃ�̂ł�낵���Ȃ�
                                // /e = �N����ʂ̕\������ѐV�K�u�b�N (Book1.xls) �̍쐬���s�킸�� Excel ���N�����܂��B
                , 0
                , 0
                , 1
                , NORMAL_PRIORITY_CLASS
                , 0
                , NULL
                , &Start
                , &ProcInfo);

    if (WaitForInputIdle(ProcInfo.hProcess, 10000)==WAIT_TIMEOUT) {

        // �f�o�b�O�p���b�Z�[�W
        // �����[�X����p�����ďo�͂���
        AfxMessageBox(_T("Timed out waiting for Excel."), MB_OK | MB_ICONINFORMATION);
    }

    // Restore the active window to the foreground...
    //  NOTE: If you comment out this line, the code will fail!
    SetForegroundWindow(hwndCurrent);

    // ----------------------------------------------------------------
    // �ȉ��̃R�[�h�����s����O��
    // CoInitialize��\�ߎ��s���Ă���

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
