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
	// アドイン名を取得する
	CString addinName = PathFindFileName(addinFilePath);

    // レジストリのパスを設定する
    CString regPath(REG_PATH_EXCEL_ADDIN_OPTIONS);
    // パスの一部分である version 部分を置換する
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

    // キーのハンドル
    HKEY hkResult;
    // 処理結果を受け取る
    DWORD dwDisposition;
    // 関数の戻り値を格納する
    LONG lResult = 0;

    // HKEY_CURRENT_USER\Software\Microsoft\Office\[version]\Excel\Options をオープン
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

        // エラーとして返す
        return lResult;
    }

	DWORD keyIndex = 0;

	const int VALNAME_LENGTH = 1024;
	DWORD valnameLength = 0;
	const int VAL_LENGTH = 1024;
	TCHAR valName[VALNAME_LENGTH];
	TCHAR val[VAL_LENGTH];
	DWORD valLength = 0;

	// キーを列挙する
	while (1) {

		// キーを取得する
		valnameLength = VALNAME_LENGTH;
		lResult = RegEnumValue(hkResult, keyIndex++, valName, &valnameLength, NULL, NULL, NULL, NULL);

		if (lResult == ERROR_NO_MORE_ITEMS) {
			lResult = 0;

			break;
		}

		// キーをCStringに変換する
		CString valNameStr = valName;

		// OPENで始まるキーの場合
		if (valNameStr.Find(_T("OPEN")) == 0) {

			// 値を取得する
			DWORD type = NULL;
			valLength = VAL_LENGTH;
			lResult = RegQueryValueEx(hkResult, valName, 0, &type, (LPBYTE)val, &valLength);

			if (lResult != ERROR_SUCCESS) {

				RegCloseKey(hkResult);
				return lResult;
			}

			// 文字列であるかを確認する
			if (type == REG_SZ) {

				// 取得したファイルパスがアドインと同じ名前である場合
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
    // 関数の戻り値を格納する
    LONG lResult = 0;

	// いったん同じ名前のアドインを削除する
	lResult = uninstallAddin(version, addinFilePath);

    if (lResult != ERROR_SUCCESS) {
        // エラーとして返す
        return lResult;
    }

    // レジストリのパスを設定する
    CString regPath(REG_PATH_EXCEL_ADDIN_OPTIONS);
    // パスの一部分である version 部分を置換する
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, version);

    // キーのハンドル
    HKEY hkResult;
    // 処理結果を受け取る
    DWORD dwDisposition;

    // HKEY_CURRENT_USER\Software\Microsoft\Office\[version]\Excel\Options をオープン
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
        // エラーとして返す
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
				// エラーとして返す
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
	// アドイン名を取得する
	CString addinFileName = PathFindFileName(addinFilePath);

    // レジストリのパスを設定する
    CString regPath(REG_PATH_EXCEL_ADDIN_MANAGER);
    // パスの一部分である version 部分を置換する
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, version);

    // キーのハンドル
    HKEY hkResult;
    // 処理結果を受け取る
    DWORD dwDisposition;
    // 関数の戻り値を格納する
    LONG lResult = 0;

    // Add-in Managerのレジストリパスキーを開く
    lResult = RegCreateKeyEx(HKEY_CURRENT_USER
                            , (LPCTSTR)regPath
                            , 0
                            , NULL
                            , REG_OPTION_NON_VOLATILE
                            , KEY_READ | KEY_SET_VALUE // 読み取りオプションとデータ設定オプションを付加する
                            , NULL
                            , &hkResult
                            , &dwDisposition);

    if (lResult != ERROR_SUCCESS) {

        // 予期せぬエラー
        return lResult;
    }

    // インデックス
    int index = 0;

    while (1) {

        // 値の名前の長さ
        DWORD valueNameLen = _MAX_PATH;
        // 値の名前
        TCHAR valueName[_MAX_PATH];

        // 値を列挙する
        lResult = RegEnumValue(hkResult
                               , index
                               , valueName
                               , &valueNameLen
                               , NULL
                               , NULL
                               , NULL
                               , NULL);

        // 戻り値チェック
        if (lResult == ERROR_NO_MORE_ITEMS) {

			lResult = 0;
            // これ以上項目がない場合
            break;

        } else if (lResult != ERROR_SUCCESS) {

            // キーをクローズする
            RegCloseKey(hkResult);
            // エラーが発生した場合は、処理を中断する
            // 予期せぬエラー
            return lResult;
        }

        // レジストリ値の名前
        CString valueNameStr(valueName);
        // レジストリ値の名前からファイル名を抽出
        CString valueNameStrOnlyFileName;

        // パスの後方からパス区切り文字を検索する
        int pos = valueNameStr.ReverseFind('\\');

        // 見つけた場合
        if (pos != -1) {

            valueNameStrOnlyFileName = valueNameStr.Mid(pos + 1);
        }

        // レジストリから取得した値とアドインファイル名が一致しているかを確認する
        if (valueNameStrOnlyFileName.Compare(addinFileName) == 0) {

            // アドインファイル名と一致しているのでレジストリから削除する
            lResult = RegDeleteValue(hkResult, valueName);

            // 戻り値チェック
            if (lResult != ERROR_SUCCESS) {

                // キーをクローズする
                RegCloseKey(hkResult);
                // 予期せぬエラー
                return lResult;
            
            }

            // キーをクローズする
            RegCloseKey(hkResult);
            // 正常終了
            return lResult;
        }

        index++;

    }

    // キーをクローズする
    RegCloseKey(hkResult);
    // 対象となるキーが見つからない場合
    return lResult;
}