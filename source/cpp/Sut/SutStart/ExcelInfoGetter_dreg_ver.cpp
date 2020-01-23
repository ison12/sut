#include "StdAfx.h"
#include "Resource.h"
#include "ExcelInfoGetter.h"

// ----------------------------------------------------
// static�����o�̏�����
const int CExcelInfoGetter::CLSID_ARRAY_LENGTH = 5;

LPCTSTR CExcelInfoGetter::CLSID_COMPONENT_EXCEL[] = {
    // Office2000
    _T("{CC29E96F-7BC2-11D1-A921-00A0C91E2AA2}")
    // OfficeXP
   ,_T("{5572D282-F5E5-11D3-A8E8-0060083FD8D3}")
    // Office2003
   ,_T("{A2B280D4-20FB-4720-99F7-40C09FBCE10A}")
    // Office2007
   ,_T("{0638C49D-BB8B-4CD1-B191-052E8F325736}")
    // Office2010
   ,_T("{0638C49D-BB8B-4CD1-B191-052E8F325736}")
};

LPCTSTR CExcelInfoGetter::COMPONENT_EXCEL_NAME[] = {
    // Office2000
    _T("Excel2000")
    // OfficeXP
   ,_T("Excel2002")
    // Office2003
   ,_T("Excel2003")
    // Office2007
   ,_T("Excel2007")
    // Office2007
   ,_T("Excel2010")
};

LPCTSTR CExcelInfoGetter::REG_PATH_EXCEL_PARAM_VERSION = _T("${version}");

LPCTSTR CExcelInfoGetter::REG_PATH_EXCEL_SECURITY_SETTING = _T("Software\\Microsoft\\Office\\${version}\\Excel\\Security");
LPCTSTR CExcelInfoGetter::REG_VALUE_NAME_EXCEL_SECURITY_DONTTRUSTINSTALLEDFILES = _T("DontTrustInstalledFiles");
LPCTSTR CExcelInfoGetter::REG_VALUE_NAME_EXCEL_SECURITY_DISABLEALLADDINS = _T("DisableAllAddins");
LPCTSTR CExcelInfoGetter::REG_VALUE_NAME_EXCEL_SECURITY_REQUIREDADDINSIG = _T("RequireAddinSig");

LPCTSTR CExcelInfoGetter::REG_PATH_EXCEL_ADDIN_MANAGER  = _T("Software\\Microsoft\\Office\\${version}\\Excel\\Add-in Manager");
LPCTSTR CExcelInfoGetter::REG_PATH_EXCEL_ADDIN_MANAGER2 = _T("Software\\Microsoft\\Office\\${version}\\Excel\\Options");

// ----------------------------------------------------

CExcelInfoGetter::CExcelInfoGetter(void)
{
}

CExcelInfoGetter::~CExcelInfoGetter(void)
{

    // �C���X�g�[���ς�Excel���X�g��1�����������ă��������������
    int size = installedExcelList.size();

    for (int i = 0; i < size; i++) {

        delete installedExcelList.at(i);
    }
}

std::vector<CExcelInfo*>& CExcelInfoGetter::getInstalledExcelApplication()
{

    for (int i = 0; i < CLSID_ARRAY_LENGTH; i++) {

        // �C���X�g�[���p�X�̃T�C�Y
        DWORD installPathSize = _MAX_PATH;
        // �C���X�g�[���p�X
        LPTSTR installPath = NULL;

        // �C���X�g�[���p�X�̃o�b�t�@���m�ۂ���
        installPath = new TCHAR[installPathSize];

        // �v���_�N�g�R�[�h���擾����
        TCHAR productCode[256];
        MsiGetProductCode(CLSID_COMPONENT_EXCEL[i], productCode);

        // �C���X�g�[���X�e�[�^�X
        INSTALLSTATE installstate;
        // Excel�A�v���P�[�V�����̃p�X���擾����
        installstate = MsiGetComponentPath(productCode
                                          ,CLSID_COMPONENT_EXCEL[i]
                                          ,installPath
                                          ,&installPathSize);

        // MsiGetComponentPath �߂�l
        // Value / Meaning
        // INSTALLSTATE_NOTUSED / The component being requested is disabled on the computer.
        // INSTALLSTATE_ABSENT / The component is not installed.
        // INSTALLSTATE_INVALIDARG / One of the function parameters is invalid.
        // INSTALLSTATE_LOCAL / The component is installed locally.
        // INSTALLSTATE_SOURCE / The component is installed to run from source.
        // INSTALLSTATE_SOURCEABSENT / The component source is inaccessible.
        // INSTALLSTATE_UNKNOWN / The product code or component ID is unknown.
        if ((installstate == INSTALLSTATE_LOCAL) || (installstate == INSTALLSTATE_SOURCE)) {

            // �C���X�g�[������Ă���ꍇ�ACExcelInfo�ɏ���ݒ肵���X�g�ɒǉ�����
            CExcelInfo* excelInfo = new CExcelInfo();
            excelInfo->appName = COMPONENT_EXCEL_NAME[i];
            excelInfo->appPath = installPath;

            installedExcelList.push_back(excelInfo);

        }

        // �C���X�g�[���p�X�̃o�b�t�@���������
        delete installPath;

    }


    return this->installedExcelList;

}

bool CExcelInfoGetter::existExcelProcess()
{

    bool funcRet = false;

    // �X�i�b�v�V���b�g�n���h��
    HANDLE hSnapShot = NULL;

    // �v���Z�X�G���g���\���� (��1�v�f���T�C�Y�ŏ�����)
    PROCESSENTRY32 p32;
    p32.dwSize = sizeof(PROCESSENTRY32);

    // �X�i�b�v�V���b�g�̍쐬
    //hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0);
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);

    if (INVALID_HANDLE_VALUE != hSnapShot) {

        // �X�i�b�v�V���b�g�̐擪����G���g���̎擾
        BOOL ret = Process32First(hSnapShot, &p32);

        while (ret) {

            TRACE(p32.szExeFile);
            TRACE(_T("\n"));

            // �v���Z�X��Excel�A�v���P�[�V���������`�F�b�N
            if (_tcscmp(p32.szExeFile, _T("EXCEL.EXE")) == 0) {

                funcRet = true;
                break;
            }

            ret = Process32Next(hSnapShot, &p32);
        }

    }

    CloseHandle(hSnapShot);

    return funcRet;
}

int CExcelInfoGetter::isPossibleAddinInstall(CString excelVersion)
{

    // Excel�̃o�[�W�����𐔒l�ɕϊ�����
    int versionInt = _tstoi((LPCTSTR)excelVersion);

    if (versionInt >= 12) {

        return isPossibleAddinInstallForOverExcel2007(excelVersion);
    
    } else if (versionInt >= 9) {

        return isPossibleAddinInstallForOverExcel2000(excelVersion);
    }

    return POSSIBLE_ADDIN_INSTALL_UNEXPECTED;
}

int CExcelInfoGetter::isPossibleAddinInstallForOverExcel2000(CString excelVersion)
{

    // ���W�X�g���̃p�X��ݒ肷��
    CString regPath(REG_PATH_EXCEL_SECURITY_SETTING);
    // �p�X�̈ꕔ���ł��� version ������u������
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

    // �L�[�̃n���h��
    HKEY hkResult;
    // �������ʂ��󂯎��
    DWORD dwDisposition;
    // �֐��̖߂�l���i�[����
    LONG lResult;

    // HKEY_CURRENT_USER\Software\Microsoft\Office\[version]\Excel\Security ���I�[�v��
    lResult = RegCreateKeyEx(HKEY_CURRENT_USER
                            , (LPCTSTR)regPath
                            , 0
                            , NULL
                            , REG_OPTION_NON_VOLATILE
                            , KEY_READ
                            , NULL
                            , &hkResult
                            , &dwDisposition);

    if (lResult != ERROR_SUCCESS) {

        // �G���[�Ƃ��ĕԂ�
        return POSSIBLE_ADDIN_INSTALL_UNEXPECTED;
    }

    // ���W�X�g������擾�����l�̃f�[�^�^
    DWORD retType;
    // ���W�X�g������擾�����l�̃o�C�g�T�C�Y
    DWORD retSize;
    // ���W�X�g������擾�����l�iDWORD�^�j
    DWORD retInt;

    // DontTrustInstalledFiles���擾
    lResult = RegQueryValueEx(hkResult
                  , REG_VALUE_NAME_EXCEL_SECURITY_DONTTRUSTINSTALLEDFILES
                  , NULL
                  , &retType
                  , (LPBYTE)&retInt
                  , &retSize);

    // �l�����݂��Ȃ��ꍇ
    if (lResult == ERROR_FILE_NOT_FOUND) {

        // �L�[���N���[�Y����
        RegCloseKey(hkResult);
        // ���A�h�C���̓C���X�g�[���\�ł���Ɣ��f����
        return POSSIBLE_ADDIN_INSTALL_OK;

    // ���������G���[�����������ꍇ
    } else if (lResult != ERROR_SUCCESS) {

        // �L�[���N���[�Y����
        RegCloseKey(hkResult);
        // �G���[�Ƃ��ĕԂ�
        return POSSIBLE_ADDIN_INSTALL_UNEXPECTED;
    }

    // ����Ɏ擾�ł����̂ŁA�߂�l�𔻕ʂ���
    if (retInt == 0) {

        // �L�[���N���[�Y����
        RegCloseKey(hkResult);
        // ��0�̏ꍇ�A�A�h�C���̓C���X�g�[���\
        return POSSIBLE_ADDIN_INSTALL_OK;
    
    } else {

        // �L�[���N���[�Y����
        RegCloseKey(hkResult);
        // ��0�ȊO�̏ꍇ�A�A�h�C���̓C���X�g�[���ł��Ȃ��\��������
        return POSSIBLE_ADDIN_INSTALL_NG;
    }
}

int CExcelInfoGetter::isPossibleAddinInstallForOverExcel2007(CString excelVersion)
{

    // ���W�X�g���̃p�X��ݒ肷��
    CString regPath(REG_PATH_EXCEL_SECURITY_SETTING);
    // �p�X�̈ꕔ���ł��� version ������u������
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

    // �L�[�̃n���h��
    HKEY hkResult;
    // �������ʂ��󂯎��
    DWORD dwDisposition;
    // �֐��̖߂�l���i�[����
    LONG lResult;

    // HKEY_CURRENT_USER\Software\Microsoft\Office\[version]\Excel\Security ���I�[�v��
    lResult = RegCreateKeyEx(HKEY_CURRENT_USER
                            , (LPCTSTR)regPath
                            , 0
                            , NULL
                            , REG_OPTION_NON_VOLATILE
                            , KEY_READ
                            , NULL
                            , &hkResult
                            , &dwDisposition);

    if (lResult != ERROR_SUCCESS) {

        // �G���[�Ƃ��ĕԂ�
        return POSSIBLE_ADDIN_INSTALL_UNEXPECTED;
    }

    // ���W�X�g������擾�����l�̃f�[�^�^
    DWORD retType;
    // ���W�X�g������擾�����l�̃o�C�g�T�C�Y
    DWORD retSize;

    DWORD retDisableAllAddins = 0;
    DWORD retRequireAddinSig  = 0;

    // DisableAllAddins���擾
    lResult = RegQueryValueEx(hkResult
                  , REG_VALUE_NAME_EXCEL_SECURITY_DISABLEALLADDINS
                  , NULL
                  , &retType
                  , (LPBYTE)&retDisableAllAddins
                  , &retSize);

    // �l�����݂��Ȃ��ꍇ
    if (lResult == ERROR_FILE_NOT_FOUND) {

        // �f�t�H���g�l��ݒ肷��
        retDisableAllAddins = 0;

    // ���������G���[�����������ꍇ
    } else if (lResult != ERROR_SUCCESS) {

        // �L�[���N���[�Y����
        RegCloseKey(hkResult);
        // �G���[�Ƃ��ĕԂ�
        return POSSIBLE_ADDIN_INSTALL_UNEXPECTED;
    }

    // RequireAddinSig���擾
    lResult = RegQueryValueEx(hkResult
                  , REG_VALUE_NAME_EXCEL_SECURITY_REQUIREDADDINSIG
                  , NULL
                  , &retType
                  , (LPBYTE)&retRequireAddinSig
                  , &retSize);

    // �l�����݂��Ȃ��ꍇ
    if (lResult == ERROR_FILE_NOT_FOUND) {

        // �f�t�H���g�l��ݒ肷��
        retRequireAddinSig = 0;

    // ���������G���[�����������ꍇ
    } else if (lResult != ERROR_SUCCESS) {

        // �L�[���N���[�Y����
        RegCloseKey(hkResult);
        // �G���[�Ƃ��ĕԂ�
        return POSSIBLE_ADDIN_INSTALL_UNEXPECTED;
    }

    // �L�[���N���[�Y����
    RegCloseKey(hkResult);

    // ����Ɏ擾�ł����̂ŁA�߂�l�𔻕ʂ���
    if (retDisableAllAddins == 0 && retRequireAddinSig == 0) {

        // DisableAllAddins �� RequireAddinSig �������Ƃ�0�̏ꍇ
        // ���A�h�C���̓C���X�g�[���\
        return POSSIBLE_ADDIN_INSTALL_OK;
    
    } else {

        // ���A�h�C���̓C���X�g�[���ł��Ȃ��\��������
        return POSSIBLE_ADDIN_INSTALL_NG;
    }

}

int CExcelInfoGetter::delAddin(CString excelVersion, CString addinFileName)
{
	bool success = false;

	// ��̃��W�X�g���L�[�G���g������A�h�C���̃t�@�C���p�X���������āA�Y������ꍇ�ɍ폜����B
	// ���
	{
		success = false;

		// ���W�X�g���̃p�X��ݒ肷��
		CString regPath(REG_PATH_EXCEL_ADDIN_MANAGER2);
		// �p�X�̈ꕔ���ł��� version ������u������
		regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

		// �L�[�̃n���h��
		HKEY hkResult;
		// �������ʂ��󂯎��
		DWORD dwDisposition;
		// �֐��̖߂�l���i�[����
		LONG lResult;

		// Add-in Manager�̃��W�X�g���p�X�L�[���J��
		lResult = RegCreateKeyEx(HKEY_CURRENT_USER
								, (LPCTSTR)regPath
								, 0
								, NULL
								, REG_OPTION_NON_VOLATILE
								, KEY_READ | KEY_SET_VALUE // �ǂݎ��I�v�V�����ƃf�[�^�ݒ�I�v�V������t������
								, NULL
								, &hkResult
								, &dwDisposition);

		if (lResult != ERROR_SUCCESS) {

			// �\�����ʃG���[
			return DEL_ADDIN_UNEXPECTED;
		}

		// �C���f�b�N�X
		int index = 0;

		while (1) {

			// �l�̖��O�̒���
			DWORD valueNameLen = _MAX_PATH;
			// �l�̖��O
			TCHAR valueName[_MAX_PATH];

			// �l��񋓂���
			lResult = RegEnumValue(hkResult
								   , index
								   , valueName
								   , &valueNameLen
								   , NULL
								   , NULL
								   , NULL
								   , NULL);

			// �߂�l�`�F�b�N
			if (lResult == ERROR_NO_MORE_ITEMS) {

				// ����ȏ㍀�ڂ��Ȃ��ꍇ
				break;

			} else if (lResult != ERROR_SUCCESS) {

				// �L�[���N���[�Y����
				RegCloseKey(hkResult);
				// �G���[�����������ꍇ�́A�����𒆒f����
				// �\�����ʃG���[
				return DEL_ADDIN_UNEXPECTED;
			}

			// ���W�X�g���l�̖��O
			CString valueNameStr(valueName);
			// ���W�X�g���l�̖��O����t�@�C�����𒊏o
			CString valueNameStrOnlyFileName;

			// �p�X�̌������p�X��؂蕶������������
			int pos = valueNameStr.Find(_T("OPEN"));

			// �������ꍇ
			if (pos != -1) {

				TCHAR* buff = NULL;
				DWORD  buffSize = 0;
				lResult = RegQueryValueEx(hkResult
								, valueName
								, NULL
								, NULL
								, NULL
								, &buffSize);

				// �߂�l�`�F�b�N
				if (lResult != ERROR_SUCCESS) {
					// �L�[���N���[�Y����
					RegCloseKey(hkResult);
					// �\�����ʃG���[
					return DEL_ADDIN_UNEXPECTED;
				}

				buff = new TCHAR[buffSize];

				// �p�X���擾����
				lResult = RegQueryValueEx(hkResult
								, valueName
								, NULL
								, NULL
								, (LPBYTE)buff
								, &buffSize);

				CString filePath(buff);
				// �p�X�̌������p�X��؂蕶������������
				int pos = filePath.ReverseFind('\\');

				// �������ꍇ
				if (pos != -1) {

					filePath = filePath.Mid(pos + 1);
				}


				// ���W�X�g������擾�����l�ƃA�h�C���t�@�C��������v���Ă��邩���m�F����
				if (addinFileName.Compare(filePath) == 0) {

					// �A�h�C���t�@�C�����ƈ�v���Ă���̂Ń��W�X�g������폜����
					lResult = RegDeleteValue(hkResult, valueName);

					// �߂�l�`�F�b�N
					if (lResult != ERROR_SUCCESS) {

						// �L�[���N���[�Y����
						RegCloseKey(hkResult);
						// �\�����ʃG���[
						return DEL_ADDIN_UNEXPECTED;
					}

					success = true;
					break;
				}
			}

			index++;

		}

		// �L�[���N���[�Y����
		RegCloseKey(hkResult);

	}

	if (!success) {

		return DEL_ADDIN_TARGET_KEY_NOT_FOUND;
	}


	// ���
	{
		success = false;

		// ���W�X�g���̃p�X��ݒ肷��
		CString regPath(REG_PATH_EXCEL_ADDIN_MANAGER);
		// �p�X�̈ꕔ���ł��� version ������u������
		regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

		// �L�[�̃n���h��
		HKEY hkResult;
		// �������ʂ��󂯎��
		DWORD dwDisposition;
		// �֐��̖߂�l���i�[����
		LONG lResult;

		// Add-in Manager�̃��W�X�g���p�X�L�[���J��
		lResult = RegCreateKeyEx(HKEY_CURRENT_USER
								, (LPCTSTR)regPath
								, 0
								, NULL
								, REG_OPTION_NON_VOLATILE
								, KEY_READ | KEY_SET_VALUE // �ǂݎ��I�v�V�����ƃf�[�^�ݒ�I�v�V������t������
								, NULL
								, &hkResult
								, &dwDisposition);

		if (lResult != ERROR_SUCCESS) {

			// �\�����ʃG���[
			return DEL_ADDIN_UNEXPECTED;
		}

		// �C���f�b�N�X
		int index = 0;

		while (1) {

			// �l�̖��O�̒���
			DWORD valueNameLen = _MAX_PATH;
			// �l�̖��O
			TCHAR valueName[_MAX_PATH];

			// �l��񋓂���
			lResult = RegEnumValue(hkResult
								   , index
								   , valueName
								   , &valueNameLen
								   , NULL
								   , NULL
								   , NULL
								   , NULL);

			// �߂�l�`�F�b�N
			if (lResult == ERROR_NO_MORE_ITEMS) {

				// ����ȏ㍀�ڂ��Ȃ��ꍇ
				break;

			} else if (lResult != ERROR_SUCCESS) {

				// �L�[���N���[�Y����
				RegCloseKey(hkResult);
				// �G���[�����������ꍇ�́A�����𒆒f����
				// �\�����ʃG���[
				return DEL_ADDIN_UNEXPECTED;
			}

			// ���W�X�g���l�̖��O
			CString valueNameStr(valueName);
			// ���W�X�g���l�̖��O����t�@�C�����𒊏o
			CString valueNameStrOnlyFileName;

			// �p�X�̌������p�X��؂蕶������������
			int pos = valueNameStr.ReverseFind('\\');

			// �������ꍇ
			if (pos != -1) {

				valueNameStrOnlyFileName = valueNameStr.Mid(pos + 1);
			}

			// ���W�X�g������擾�����l�ƃA�h�C���t�@�C��������v���Ă��邩���m�F����
			if (valueNameStrOnlyFileName.Compare(addinFileName) == 0) {

				// �A�h�C���t�@�C�����ƈ�v���Ă���̂Ń��W�X�g������폜����
				lResult = RegDeleteValue(hkResult, valueName);

				// �߂�l�`�F�b�N
				if (lResult != ERROR_SUCCESS) {

					// �L�[���N���[�Y����
					RegCloseKey(hkResult);
					// �\�����ʃG���[
					return DEL_ADDIN_UNEXPECTED;
	            
				}

				success = true;
			}

			index++;

		}

		// �L�[���N���[�Y����
		RegCloseKey(hkResult);

	}

	if (!success) {

		return DEL_ADDIN_TARGET_KEY_NOT_FOUND;
	}

    // �ΏۂƂȂ�L�[��������Ȃ��ꍇ
    return DEL_ADDIN_OK;
}

int CExcelInfoGetter::addAddin(CString excelVersion, CString addinFileName)
{
    // ���W�X�g���̃p�X��ݒ肷��
    CString regPath(REG_PATH_EXCEL_ADDIN_MANAGER2);
    // �p�X�̈ꕔ���ł��� version ������u������
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

    // �L�[�̃n���h��
    HKEY hkResult;
    // �������ʂ��󂯎��
    DWORD dwDisposition;
    // �֐��̖߂�l���i�[����
    LONG lResult;

    // Add-in Manager�̃��W�X�g���p�X�L�[���J��
    lResult = RegCreateKeyEx(HKEY_CURRENT_USER
                            , (LPCTSTR)regPath
                            , 0
                            , NULL
                            , REG_OPTION_NON_VOLATILE
                            , KEY_READ | KEY_SET_VALUE // �ǂݎ��I�v�V�����ƃf�[�^�ݒ�I�v�V������t������
                            , NULL
                            , &hkResult
                            , &dwDisposition);

    if (lResult != ERROR_SUCCESS) {

        // �\�����ʃG���[
        return ADD_ADDIN_TARGET_KEY_NOT_FOUND;
    }

	// �A�h�C���t�@�C�����̃o�C�g�������߂�
	DWORD addinFileNameSize =  addinFileName.GetLength() * sizeof(TCHAR);

	// �t�@�C���p�X��o�^����
	lResult = RegSetValueEx(hkResult, _T("OPEN"), 0, REG_SZ, (BYTE*)((LPCTSTR)addinFileName), addinFileNameSize);

    // �߂�l�`�F�b�N
    if (lResult != ERROR_SUCCESS) {

        // �L�[���N���[�Y����
        RegCloseKey(hkResult);
        // �\�����ʃG���[
		return ADD_ADDIN_UNEXPECTED;
    
    }

    // �L�[���N���[�Y����
    RegCloseKey(hkResult);
    // �ΏۂƂȂ�L�[��������Ȃ��ꍇ
	return ADD_ADDIN_OK;
}

void CExcelInfoGetter::searchRunningObjectTable()
{

    //HRESULT hr;

    //// CreateBindCtx
    //// http://msdn.microsoft.com/en-us/library/ms678542(VS.85).aspx
    //// Get a BindCtx.
    //IBindCtx *pbc;
    //hr = CreateBindCtx(0, &pbc);
    //
    //if (FAILED(hr)) {
    //    //DoErr("CreateBindCtx()", hr);
    //    return;
    //}

    //// IBindCtx::GetRunningObjectTable
    //// http://msdn.microsoft.com/en-us/library/ms680065(VS.85).aspx
    //// Get running-object table.
    //IRunningObjectTable *prot;
    //hr = pbc->GetRunningObjectTable(&prot);
    //
    //if (FAILED(hr)) {
    //    //DoErr("GetRunningObjectTable()", hr);
    //    pbc->Release();
    //    return;
    //}

    //// IRunningObjectTable::EnumRunning
    //// http://msdn.microsoft.com/en-us/library/ms678491(VS.85).aspx
    //// Creates and returns a pointer to an enumerator that can list the monikers of all the objects currently registered in the running object table (ROT).
    //// Get enumeration interface.
    //IEnumMoniker *pem;
    //hr = prot->EnumRunning(&pem);

    //if (FAILED(hr)) {
    //    //DoErr("EnumRunning()", hr);
    //    prot->Release();
    //    pbc->Release();
    //    return;
    //}

    //// Start at the beginning.
    //pem->Reset();

    //// Churn through enumeration.
    //ULONG fetched;
    //IMoniker *pmon;
    //int n = 0;

    //while (pem->Next(1, &pmon, &fetched) == S_OK) {

    //    CreateAntiMoniker(&pmon);

    //    // Get DisplayName.
    //    LPOLESTR pName;
    //    pmon->GetDisplayName(pbc, NULL, &pName);

    //    // Convert it to ASCII.
    //    char szName[512];
    //    WideCharToMultiByte(CP_ACP, 0, pName, -1, szName, 512, NULL, NULL);

    //    std::cout << szName << std::endl;

    //    

    //    // Compare it against the name we got in SetHostNames().
    //    //if (!strcmp(szName, m_szDocName)) {

    //        //DoMsg("Found document in ROT!");

    //        // Bind to this ROT entry.
    //        IDispatch *pDisp = NULL;
    //        hr = pmon->BindToObject(pbc, NULL, IID_IDispatch, (void**)&pDisp);

    //        if (!FAILED(hr)) {

    //            // Remember IDispatch.
    //            //m_pDocDisp = pDisp;

    //            //m_pDocDisp);
    //            //DoMsg(buf);

    //            if (pDisp == NULL) {

    //                continue;
    //            }

    //            m_shellApplication.DetachDispatch();
    //            m_shellApplication.AttachDispatch(pDisp);
    //            CString str = NULL;
    //            {
    //                // Invoke��
    //                LPOLESTR name = L"Name";

    //                // Invoke����f�B�X�p�b�`ID
    //                DISPID dispid;

    //                hr = m_shellApplication.m_lpDispatch->GetIDsOfNames(IID_NULL
    //                                                             , &name
    //                                                             , 1
    //                                                             , LOCALE_USER_DEFAULT
    //                                                             , &dispid);

    //                if (FAILED(hr)) {

    //                    ::MessageBox(NULL, _T("error"), _T("error msg"), 0);
    //                
    //                } else {


    //                    m_shellApplication.GetProperty(dispid, VT_BSTR, (void*)&str);
    //                    std::cout << str << std::endl;
    //                }



    //            }


    //        }

    //        else {

    //            //DoErr("BindToObject()", hr);

    //        }
    //    //}

    //    // Release interfaces.
    //    pmon->Release();

    //    // Break out if we obtained the IDispatch successfully.
    //    //if (m_pDocDisp != NULL) break;
    //}

    //// Release interfaces.
    //pem->Release();
    //prot->Release();
    //pbc->Release();
}

CString CExcelInfoGetter::getVersion(CString& productName)
{
	if (productName.Compare(_T("Excel2000")) == 0) {

		return _T("9.0");
	} else if (productName.Compare(_T("Excel2002")) == 0) {

		return _T("10.0");
	} else if (productName.Compare(_T("Excel2003")) == 0) {

		return _T("11.0");
	} else if (productName.Compare(_T("Excel2007")) == 0) {

		return _T("12.0");
	} else if (productName.Compare(_T("Excel2010")) == 0) {

		return _T("13.0");
	}

	return _T("");
}