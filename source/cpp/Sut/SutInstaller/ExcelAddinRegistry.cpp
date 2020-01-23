#include "pch.h"
#include "ExcelAddinRegistry.h"

LPCTSTR ExcelAddinRegistry::REG_PATH_EXCEL_PARAM_VERSION = _T("${version}");

LPCTSTR ExcelAddinRegistry::REG_PATH_EXCEL_ADDIN_MANAGER = _T("Software\\Microsoft\\Office\\${version}\\Excel\\Add-in Manager");
LPCTSTR ExcelAddinRegistry::REG_PATH_EXCEL_ADDIN_OPTIONS = _T("Software\\Microsoft\\Office\\${version}\\Excel\\Options");

ExcelAddinRegistry::ExcelAddinRegistry(void)
{
}

ExcelAddinRegistry::~ExcelAddinRegistry(void)
{
}

int ExcelAddinRegistry::uninstallAddin(CString& excelVersion, CString& addinFilePath)
{
	// �A�h�C�������擾����
	CString addinName = PathFindFileName(addinFilePath);

    // ���W�X�g���̃p�X��ݒ肷��
    CString regPath(REG_PATH_EXCEL_ADDIN_OPTIONS);
    // �p�X�̈ꕔ���ł��� version ������u������
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

    // �L�[�̃n���h��
    HKEY hkResult;
    // �������ʂ��󂯎��
    DWORD dwDisposition;
    // �֐��̖߂�l���i�[����
    LONG lResult = 0;

    // HKEY_CURRENT_USER\Software\Microsoft\Office\[version]\Excel\Options ���I�[�v��
    lResult = RegCreateKeyEx(HKEY_CURRENT_USER
                            , (LPCTSTR)regPath
                            , 0
                            , NULL
                            , REG_OPTION_NON_VOLATILE
							, KEY_ALL_ACCESS
                            , NULL
                            , &hkResult
                            , &dwDisposition);

    if (lResult != ERROR_SUCCESS) {

        // �G���[�Ƃ��ĕԂ�
        return lResult;
    }

	DWORD keyIndex = 0;

	const int VALNAME_LENGTH = 1024;
	DWORD valnameLength = 0;
	const int VAL_LENGTH = 1024;
	TCHAR valName[VALNAME_LENGTH];
	TCHAR val[VAL_LENGTH];
	DWORD valLength = 0;

	// �L�[��񋓂���
	while (1) {

		// �L�[���擾����
		valnameLength = VALNAME_LENGTH;
		lResult = RegEnumValue(hkResult, keyIndex++, valName, &valnameLength, NULL, NULL, NULL, NULL);

		if (lResult == ERROR_NO_MORE_ITEMS) {
			lResult = 0;

			break;
		}

		// �L�[��CString�ɕϊ�����
		CString valNameStr = valName;

		// OPEN�Ŏn�܂�L�[�̏ꍇ
		if (valNameStr.Find(_T("OPEN")) == 0) {

			// �l���擾����
			DWORD type = NULL;
			valLength = VAL_LENGTH;
			lResult = RegQueryValueEx(hkResult, valName, 0, &type, (LPBYTE)val, &valLength);

			if (lResult != ERROR_SUCCESS) {

				RegCloseKey(hkResult);
				return lResult;
			}

			// ������ł��邩���m�F����
			if (type == REG_SZ) {

				// �擾�����t�@�C���p�X���A�h�C���Ɠ������O�ł���ꍇ
				CString fileName = PathFindFileName(val);
				fileName.Replace(_T("\""), _T(""));
				if (fileName == addinName) {

					lResult = RegDeleteValue(hkResult, valName);

					if (lResult != ERROR_SUCCESS) {

						RegCloseKey(hkResult);
						return lResult;
					}
				}
			}
		}
	}

	RegCloseKey(hkResult);

	return lResult;

}

int ExcelAddinRegistry::installAddin(CString& version, CString& addinFilePath)
{
    // �֐��̖߂�l���i�[����
    LONG lResult = 0;

	// �������񓯂����O�̃A�h�C�����폜����
	lResult = uninstallAddin(version, addinFilePath);

    if (lResult != ERROR_SUCCESS) {
        // �G���[�Ƃ��ĕԂ�
        return lResult;
    }

    // ���W�X�g���̃p�X��ݒ肷��
    CString regPath(REG_PATH_EXCEL_ADDIN_OPTIONS);
    // �p�X�̈ꕔ���ł��� version ������u������
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, version);

    // �L�[�̃n���h��
    HKEY hkResult;
    // �������ʂ��󂯎��
    DWORD dwDisposition;

    // HKEY_CURRENT_USER\Software\Microsoft\Office\[version]\Excel\Options ���I�[�v��
    lResult = RegCreateKeyEx(HKEY_CURRENT_USER
                            , (LPCTSTR)regPath
                            , 0
                            , NULL
                            , REG_OPTION_NON_VOLATILE
							, KEY_ALL_ACCESS
                            , NULL
                            , &hkResult
                            , &dwDisposition);

    if (lResult != ERROR_SUCCESS) {

		RegCloseKey(hkResult);
        // �G���[�Ƃ��ĕԂ�
        return lResult;
    }

	DWORD keyIndex = 0;

	const int KEY_LENGTH = 1024;
	const int VAL_LENGTH = 1024;
	TCHAR key[KEY_LENGTH];
	//TCHAR val[VAL_LENGTH];

	DWORD t = 0;

	while (1) {

		if (t == 0) {

			_stprintf_s(key, KEY_LENGTH, _T("OPEN"));
		} else {

			_stprintf_s(key, KEY_LENGTH, _T("OPEN%d"), t);
		}

		if (RegQueryValueEx(hkResult, key, 0, 0, NULL, NULL) == ERROR_FILE_NOT_FOUND) {

			addinFilePath = _T("\"") + addinFilePath + _T("\"");

			lResult = RegSetValueEx(hkResult, key, 0, REG_SZ, (LPBYTE)((LPCTSTR)addinFilePath), addinFilePath.GetLength() * 2);

			if (lResult != ERROR_SUCCESS) {

				RegCloseKey(hkResult);
				// �G���[�Ƃ��ĕԂ�
				return lResult;
			}

			break;
		}

		t++;
	}
	return 0;
}

int ExcelAddinRegistry::uninstallAddinAtAddinManager(CString& version, CString& addinFilePath)
{
	// �A�h�C�������擾����
	CString addinFileName = PathFindFileName(addinFilePath);

    // ���W�X�g���̃p�X��ݒ肷��
    CString regPath(REG_PATH_EXCEL_ADDIN_MANAGER);
    // �p�X�̈ꕔ���ł��� version ������u������
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, version);

    // �L�[�̃n���h��
    HKEY hkResult;
    // �������ʂ��󂯎��
    DWORD dwDisposition;
    // �֐��̖߂�l���i�[����
    LONG lResult = 0;

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
        return lResult;
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

			lResult = 0;
            // ����ȏ㍀�ڂ��Ȃ��ꍇ
            break;

        } else if (lResult != ERROR_SUCCESS) {

            // �L�[���N���[�Y����
            RegCloseKey(hkResult);
            // �G���[�����������ꍇ�́A�����𒆒f����
            // �\�����ʃG���[
            return lResult;
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
                return lResult;
            
            }

            // �L�[���N���[�Y����
            RegCloseKey(hkResult);
            // ����I��
            return lResult;
        }

        index++;

    }

    // �L�[���N���[�Y����
    RegCloseKey(hkResult);
    // �ΏۂƂȂ�L�[��������Ȃ��ꍇ
    return lResult;
}