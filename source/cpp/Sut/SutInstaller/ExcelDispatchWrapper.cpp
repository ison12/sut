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
    // �f�B�X�p�b�`�C���^�[�t�F�[�X���擾
    IDispatch* ret = excelApplication.m_lpDispatch;
    // �������
    excelApplication.ReleaseDispatch();

    return ret;
}

CString CExcelDispatchWrapper::getVersion()
{
    try {

        // �v���p�e�B�E���\�b�h�̎��s����
        HRESULT result;

        // �߂�l
        VARIANT vResult;
        VariantInit(&vResult);

        // Invoke��
        LPOLESTR name = L"Version";

        // Invoke����f�B�X�p�b�`ID
        DISPID dispid;

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            // �ʏ�L�蓾�Ȃ�
            // �󕶎����Ԃ�
            return CString();
        }

        // �o�[�W������������擾����
        excelApplication.GetProperty(dispid, VT_VARIANT, (void*)&vResult);

        // �߂�l��BSTR�̏ꍇ
        if (vResult.vt & VT_BSTR) {

            return vResult.bstrVal;
        
        // �߂�l��BSTR�ȊO�̏ꍇ
        } else {

            // �ʏ�L�蓾�Ȃ�
            // �󕶎����Ԃ�
            return CString();
        }

    }  // End try.

    catch(COleException *e)
    {

        // �G���[���b�Z�[�W�擾
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // ���b�Z�[�W��\������
        AfxMessageBox(errorMess, MB_ICONHAND);

        // �ʏ�L�蓾�Ȃ�
        // �󕶎����Ԃ�
        return CString();
    }

    catch(COleDispatchException *e)
    {

        // �G���[���b�Z�[�W�擾
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // ���b�Z�[�W��\������
        AfxMessageBox(errorMess, MB_ICONHAND);

        // �ʏ�L�蓾�Ȃ�
        // �󕶎����Ԃ�
        return CString();
    }
    catch(...)
    {
        // ���b�Z�[�W��\������
        AfxMessageBox(_T("Unexpected error"), MB_ICONHAND);

        // �ʏ�L�蓾�Ȃ�
        // �󕶎����Ԃ�
        return CString();
    }

}

int CExcelDispatchWrapper::displayAlerts(BOOL b)
{

    try {

        // �v���p�e�B�E���\�b�h�̎��s����
        HRESULT result;

        // �߂�l
        VARIANT vResult;
        VariantInit(&vResult);

        // Invoke��
        LPOLESTR name = L"DisplayAlerts";

        // Invoke����f�B�X�p�b�`ID
        DISPID dispid;

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            // �ʏ�L�蓾�Ȃ�
            return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
        }

        // Application.DisplayAlerts��ݒ�
        excelApplication.SetProperty(dispid, VT_BOOL, b);

    }  // End try.

    catch(COleException *e)
    {

        // �G���[���b�Z�[�W�擾
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // ���b�Z�[�W��\������
        AfxMessageBox(errorMess, MB_ICONHAND);

        // �ʏ�L�蓾�Ȃ�
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    catch(COleDispatchException *e)
    {

        // �G���[���b�Z�[�W�擾
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // ���b�Z�[�W��\������
        AfxMessageBox(errorMess, MB_ICONHAND);

        // �ʏ�L�蓾�Ȃ�
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }
    catch(...)
    {
        // ���b�Z�[�W��\������
        AfxMessageBox(_T("Unexpected error"), MB_ICONHAND);

        // �ʏ�L�蓾�Ȃ�
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    return CExcelDispatchWrapper::SUCCESS;
}

int CExcelDispatchWrapper::appVisible(BOOL visible)
{
    try {

        // �v���p�e�B�E���\�b�h�̎��s����
        HRESULT result;

        // �߂�l
        VARIANT vResult;
        VariantInit(&vResult);

        // Invoke��
        LPOLESTR name = L"Visible";

        // Invoke����f�B�X�p�b�`ID
        DISPID dispid;

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            // �ʏ�L�蓾�Ȃ�
            return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
        }

        // Application.Visible��ݒ�
        excelApplication.SetProperty(dispid, VT_BOOL, visible);

    }  // End try.

    catch(COleException *e)
    {
        // �G���[���b�Z�[�W�擾
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // ���b�Z�[�W��\������
        AfxMessageBox(errorMess, MB_ICONHAND);


        // �ʏ�L�蓾�Ȃ�
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    catch(COleDispatchException *e)
    {
        // �G���[���b�Z�[�W�擾
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // ���b�Z�[�W��\������
        AfxMessageBox(errorMess, MB_ICONHAND);


        // �ʏ�L�蓾�Ȃ�
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }
    catch(...)
    {
        // ���b�Z�[�W��\������
        AfxMessageBox(_T("Unexpected error"), MB_ICONHAND);

        // �ʏ�L�蓾�Ȃ�
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    return CExcelDispatchWrapper::SUCCESS;
}

int CExcelDispatchWrapper::appQuit()
{

    try {

        // �v���p�e�B�E���\�b�h�̎��s����
        HRESULT result;

        // �߂�l
        VARIANT vResult;
        VariantInit(&vResult);

        // Invoke��
        LPOLESTR name = L"Quit";

        // Invoke����f�B�X�p�b�`ID
        DISPID dispid;

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            // �ʏ�L�蓾�Ȃ�
            return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
        }

        // Application.Visible��ݒ�
        excelApplication.InvokeHelper(dispid, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);

    }  // End try.

    catch(COleException *e)
    {
        // �G���[���b�Z�[�W�擾
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // ���b�Z�[�W��\������
        AfxMessageBox(errorMess, MB_ICONHAND);

        // �ʏ�L�蓾�Ȃ�
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    catch(COleDispatchException *e)
    {
        // �G���[���b�Z�[�W�擾
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // ���b�Z�[�W��\������
        AfxMessageBox(errorMess, MB_ICONHAND);

        // �ʏ�L�蓾�Ȃ�
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }
    catch(...)
    {
        // ���b�Z�[�W��\������
        AfxMessageBox(_T("Unexpected error"), MB_ICONHAND);

        // �ʏ�L�蓾�Ȃ�
        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    return CExcelDispatchWrapper::SUCCESS;
}

IDispatch* CExcelDispatchWrapper::getAddinsObject()
{

    // �v���p�e�B�E���\�b�h�̎��s����
    HRESULT result;

    // Invoke����f�B�X�p�b�`ID
    DISPID dispid;

    // IDispatch�I�u�W�F�N�g
    IDispatch* pDisp = NULL;

    // -------------------------------------------------------------------
    // Application.Workbooks�I�u�W�F�N�g�̎擾
    // Workbooks�I�u�W�F�N�g
    COleDispatchDriver workbooks;
    {
        // Invoke��
        LPOLESTR name = L"Workbooks";

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            return NULL;
        }

        // �߂�l��������
        pDisp = NULL;

        // Addins�I�u�W�F�N�g���擾����
        excelApplication.GetProperty(dispid, VT_DISPATCH, (void*)&pDisp);

        if (pDisp != NULL) {

            workbooks.AttachDispatch(pDisp);
        
        } else {

            return NULL;
        }
    }

    // --------------------------------------------------------------------------------------
    // Active��Workbook�̑��݂𔻒肵�A�����Workbook���擾���A�Ȃ���ΐV���ɒǉ�����
    // ��ActiveWorkbook�����݂��Ȃ��ꍇ�AAddins�ɑ΂��鑀�삪���s���邱�Ƃ�����
    // -------------------------------------------------------------------
    // Workbook�I�u�W�F�N�g
    COleDispatchDriver workbook;
    // ���ݗL��
    bool isActiveWorkbook = false;

    // Application.ActiveWorkbook�̃`�F�b�N
    {
        // Invoke��
        LPOLESTR name = L"ActiveWorkbook";

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            return NULL;
        }

        // �߂�l��������
        pDisp = NULL;

        // Workbook��ǉ����I�u�W�F�N�g���擾����
        excelApplication.GetProperty(dispid, VT_DISPATCH, (void*)&pDisp);

        if (pDisp != NULL) {

            workbook.AttachDispatch(pDisp);
            isActiveWorkbook = true;
        
        }
    }


    // -------------------------------------------------------------------
    // Workbooks.Add�I�u�W�F�N�g�̎擾
    if (!isActiveWorkbook) {

        // Invoke��
        LPOLESTR name = L"Add";

        result = workbooks.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            return NULL;
        }

        // �߂�l��������
        pDisp = NULL;

        static BYTE paramInfo[] = VTS_VARIANT;

        // �p�����[�^
        VARIANT paramVar;
        VariantInit(&paramVar);
        paramVar.vt = VT_ERROR;
        paramVar.scode = DISP_E_PARAMNOTFOUND;

        // Workbook��ǉ����I�u�W�F�N�g���擾����
        workbooks.InvokeHelper(dispid, DISPATCH_METHOD, VT_DISPATCH, (void*)&pDisp, paramInfo, &paramVar);

        if (pDisp != NULL) {

            workbook.AttachDispatch(pDisp);

            // Application�I�u�W�F�N�g���擾
            // Visible�����s����
        
        } else {

            return NULL;
        }
    }

    // -------------------------------------------------------------------
    // Application.AddIns�I�u�W�F�N�g�̎擾
    {
        // Invoke��
        LPOLESTR name = L"AddIns";

        result = excelApplication.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(result)) {

            return NULL;
        }

        // �߂�l��������
        pDisp = NULL;

        // Addins�I�u�W�F�N�g���擾����
        excelApplication.GetProperty(dispid, VT_DISPATCH, (void*)&pDisp);

    }

    return pDisp;
}

int CExcelDispatchWrapper::attachAddin(CString addinPath)
{

    try {
        // �v���p�e�B�E���\�b�h�̎��s����
        HRESULT result;

        // Invoke����f�B�X�p�b�`ID
        DISPID dispid;

        // IDispatch�I�u�W�F�N�g
        IDispatch* pDisp = NULL;

        // -------------------------------------------------------------------
        // Application.AddIns�I�u�W�F�N�g�̎擾
        // Addins�I�u�W�F�N�g
        COleDispatchDriver addins;
        pDisp = getAddinsObject();

        if (pDisp != NULL) {

            addins.AttachDispatch(pDisp);
        
        } else {

            return CExcelDispatchWrapper::ERROR_UNEXPECTED;
        }


        // �A�h�C�������ɃC���X�g�[������Ă��邩���m�F����
        bool isAddinInstalled = false;

        // Addin�I�u�W�F�N�g
        COleDispatchDriver addin;

        // -------------------------------------------------------------------
        // Application.AddIns�I�u�W�F�N�g��Item���\�b�h�̎��s
        {
            // Invoke��
            LPOLESTR name = L"Item";

            result = addins.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
            if (FAILED(result)) {

                return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
            }

            // �߂�l��������
            pDisp = NULL;

            static BYTE paramInfo[] = VTS_VARIANT;

            // �A�h�C�������g���ăA�h�C���I�u�W�F�N�g����������
            CString addinName;
            addinName.LoadString(IDS_ADDIN_NAME);

            BSTR paramStr = addinName.AllocSysString();

            VARIANT paramVar;
            VariantInit(&paramVar);
            paramVar.vt = VT_BSTR;
            paramVar.bstrVal = paramStr;

            try {

                // �A�h�C����ǉ�����
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

        // �A�h�C�������ɃC���X�g�[������Ă���ꍇ
        if (isAddinInstalled) {

            return ERROR_ITEM_EXIST;
        }

        // -------------------------------------------------------------------
        // Application.AddIns�I�u�W�F�N�g��Add���\�b�h�̎��s
        {
            // Invoke��
            LPOLESTR name = L"Add";

            result = addins.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
            if (FAILED(result)) {

                return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
            }

            // �߂�l��������
            pDisp = NULL;

            static BYTE paramInfo[] = VTS_BSTR VTS_VARIANT;

            BSTR paramStr = addinPath.AllocSysString();

            VARIANT paramVar;
            VariantInit(&paramVar);
            paramVar.vt = VT_BOOL;
            paramVar.boolVal = FALSE;

            // �A�h�C����ǉ�����
            addins.InvokeHelper(dispid, DISPATCH_METHOD, VT_DISPATCH, (void*)&pDisp, paramInfo, paramStr, &paramVar);

            SysFreeString(paramStr);

            if (pDisp != NULL) {

                addin.AttachDispatch(pDisp);
            
            } else {

                return CExcelDispatchWrapper::ERROR_UNEXPECTED;
            }
        }

        // -------------------------------------------------------------------
        // Addin�I�u�W�F�N�g��Installed�v���p�e�B�̕ύX
        {
            // Invoke��
            LPOLESTR name = L"Installed";

            result = addin.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
            if (FAILED(result)) {

                return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
            }

            // �A�h�C����ǉ�����
            addin.SetProperty(dispid, VT_BOOL, TRUE);

        }


    }  // End try.

    catch(COleException *e)
    {

        // �G���[���b�Z�[�W�擾
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // ���b�Z�[�W��\������
        AfxMessageBox(errorMess, MB_ICONHAND);

        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    catch(COleDispatchException *e)
    {

        // �G���[���b�Z�[�W�擾
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // ���b�Z�[�W��\������
        AfxMessageBox(errorMess, MB_ICONHAND);

        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }
    catch(...)
    {
        // ���b�Z�[�W��\������
        AfxMessageBox(_T("Unexpected error"), MB_ICONHAND);

        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    return CExcelDispatchWrapper::SUCCESS;
}

int CExcelDispatchWrapper::removeAddin(CString addinPath)
{

    try {
        // �v���p�e�B�E���\�b�h�̎��s����
        HRESULT result;

        // Invoke����f�B�X�p�b�`ID
        DISPID dispid;

        // IDispatch�I�u�W�F�N�g
        IDispatch* pDisp = NULL;

        // -------------------------------------------------------------------
        // Application.AddIns�I�u�W�F�N�g�̎擾
        // Addins�I�u�W�F�N�g
        COleDispatchDriver addins;
        pDisp = getAddinsObject();

        if (pDisp != NULL) {

            addins.AttachDispatch(pDisp);
        
        } else {

            return CExcelDispatchWrapper::ERROR_UNEXPECTED;
        }

        // -------------------------------------------------------------------
        // Application.AddIns�I�u�W�F�N�g��Item���\�b�h�̎��s
        // Addin�I�u�W�F�N�g
        COleDispatchDriver addin;
        {
            // Invoke��
            LPOLESTR name = L"Item";

            result = addins.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
            if (FAILED(result)) {

                return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
            }

            // �߂�l��������
            pDisp = NULL;

            static BYTE paramInfo[] = VTS_VARIANT;

            BSTR paramStr = addinPath.AllocSysString();

            VARIANT paramVar;
            VariantInit(&paramVar);
            paramVar.vt = VT_BSTR;
            paramVar.bstrVal = paramStr;

            // �A�h�C����ǉ�����
            addins.InvokeHelper(dispid, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&pDisp, paramInfo, &paramVar);

            SysFreeString(paramStr);

            if (pDisp != NULL) {

                addin.AttachDispatch(pDisp);
            
            } else {

                // �폜�Ώۂ�������Ȃ��ꍇ
                return CExcelDispatchWrapper::ERROR_ITEM_NOT_FOUND;
            }

        }

        // -------------------------------------------------------------------
        // Addin�I�u�W�F�N�g��Installed�v���p�e�B�̕ύX
        {
            // Invoke��
            LPOLESTR name = L"Installed";

            result = addin.m_lpDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
            if (FAILED(result)) {

                return CExcelDispatchWrapper::ERROR_DISPID_NOT_FOUND;
            }

            // �A�h�C����ǉ�����
            addin.SetProperty(dispid, VT_BOOL, FALSE);

        }


    }  // End try.

    catch(COleException *e)
    {
        // �G���[���b�Z�[�W�擾
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // ���b�Z�[�W��\������
        AfxMessageBox(errorMess, MB_ICONHAND);

        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    catch(COleDispatchException *e)
    {
        // -2147352565 (8002000B)    Invalid index.
        if (-2147352565 == e->m_scError) {

            return CExcelDispatchWrapper::ERROR_ITEM_NOT_FOUND;
        }

        // �G���[���b�Z�[�W�擾
        TCHAR errorMess[256];
        e->GetErrorMessage(errorMess, 256);
        // ���b�Z�[�W��\������
        AfxMessageBox(errorMess, MB_ICONHAND);

        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }
    catch(...)
    {
        // ���b�Z�[�W��\������
        AfxMessageBox(_T("Unexpected error"), MB_ICONHAND);

        return CExcelDispatchWrapper::ERROR_UNEXPECTED;
    }

    return CExcelDispatchWrapper::SUCCESS;
}
