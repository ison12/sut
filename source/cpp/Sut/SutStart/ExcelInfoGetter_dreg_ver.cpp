#include "StdAfx.h"
#include "Resource.h"
#include "ExcelInfoGetter.h"

// ----------------------------------------------------
// staticメンバの初期化
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

    // インストール済みExcelリストを1件ずつ処理してメモリを解放する
    int size = installedExcelList.size();

    for (int i = 0; i < size; i++) {

        delete installedExcelList.at(i);
    }
}

std::vector<CExcelInfo*>& CExcelInfoGetter::getInstalledExcelApplication()
{

    for (int i = 0; i < CLSID_ARRAY_LENGTH; i++) {

        // インストールパスのサイズ
        DWORD installPathSize = _MAX_PATH;
        // インストールパス
        LPTSTR installPath = NULL;

        // インストールパスのバッファを確保する
        installPath = new TCHAR[installPathSize];

        // プロダクトコードを取得する
        TCHAR productCode[256];
        MsiGetProductCode(CLSID_COMPONENT_EXCEL[i], productCode);

        // インストールステータス
        INSTALLSTATE installstate;
        // Excelアプリケーションのパスを取得する
        installstate = MsiGetComponentPath(productCode
                                          ,CLSID_COMPONENT_EXCEL[i]
                                          ,installPath
                                          ,&installPathSize);

        // MsiGetComponentPath 戻り値
        // Value / Meaning
        // INSTALLSTATE_NOTUSED / The component being requested is disabled on the computer.
        // INSTALLSTATE_ABSENT / The component is not installed.
        // INSTALLSTATE_INVALIDARG / One of the function parameters is invalid.
        // INSTALLSTATE_LOCAL / The component is installed locally.
        // INSTALLSTATE_SOURCE / The component is installed to run from source.
        // INSTALLSTATE_SOURCEABSENT / The component source is inaccessible.
        // INSTALLSTATE_UNKNOWN / The product code or component ID is unknown.
        if ((installstate == INSTALLSTATE_LOCAL) || (installstate == INSTALLSTATE_SOURCE)) {

            // インストールされている場合、CExcelInfoに情報を設定しリストに追加する
            CExcelInfo* excelInfo = new CExcelInfo();
            excelInfo->appName = COMPONENT_EXCEL_NAME[i];
            excelInfo->appPath = installPath;

            installedExcelList.push_back(excelInfo);

        }

        // インストールパスのバッファを解放する
        delete installPath;

    }


    return this->installedExcelList;

}

bool CExcelInfoGetter::existExcelProcess()
{

    bool funcRet = false;

    // スナップショットハンドル
    HANDLE hSnapShot = NULL;

    // プロセスエントリ構造体 (第1要素をサイズで初期化)
    PROCESSENTRY32 p32;
    p32.dwSize = sizeof(PROCESSENTRY32);

    // スナップショットの作成
    //hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0);
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);

    if (INVALID_HANDLE_VALUE != hSnapShot) {

        // スナップショットの先頭からエントリの取得
        BOOL ret = Process32First(hSnapShot, &p32);

        while (ret) {

            TRACE(p32.szExeFile);
            TRACE(_T("\n"));

            // プロセスがExcelアプリケーションかをチェック
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

    // Excelのバージョンを数値に変換する
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

    // レジストリのパスを設定する
    CString regPath(REG_PATH_EXCEL_SECURITY_SETTING);
    // パスの一部分である version 部分を置換する
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

    // キーのハンドル
    HKEY hkResult;
    // 処理結果を受け取る
    DWORD dwDisposition;
    // 関数の戻り値を格納する
    LONG lResult;

    // HKEY_CURRENT_USER\Software\Microsoft\Office\[version]\Excel\Security をオープン
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

        // エラーとして返す
        return POSSIBLE_ADDIN_INSTALL_UNEXPECTED;
    }

    // レジストリから取得した値のデータ型
    DWORD retType;
    // レジストリから取得した値のバイトサイズ
    DWORD retSize;
    // レジストリから取得した値（DWORD型）
    DWORD retInt;

    // DontTrustInstalledFilesを取得
    lResult = RegQueryValueEx(hkResult
                  , REG_VALUE_NAME_EXCEL_SECURITY_DONTTRUSTINSTALLEDFILES
                  , NULL
                  , &retType
                  , (LPBYTE)&retInt
                  , &retSize);

    // 値が存在しない場合
    if (lResult == ERROR_FILE_NOT_FOUND) {

        // キーをクローズする
        RegCloseKey(hkResult);
        // ★アドインはインストール可能であると判断する
        return POSSIBLE_ADDIN_INSTALL_OK;

    // そもそもエラーが発生した場合
    } else if (lResult != ERROR_SUCCESS) {

        // キーをクローズする
        RegCloseKey(hkResult);
        // エラーとして返す
        return POSSIBLE_ADDIN_INSTALL_UNEXPECTED;
    }

    // 正常に取得できたので、戻り値を判別する
    if (retInt == 0) {

        // キーをクローズする
        RegCloseKey(hkResult);
        // ★0の場合、アドインはインストール可能
        return POSSIBLE_ADDIN_INSTALL_OK;
    
    } else {

        // キーをクローズする
        RegCloseKey(hkResult);
        // ★0以外の場合、アドインはインストールできない可能性がある
        return POSSIBLE_ADDIN_INSTALL_NG;
    }
}

int CExcelInfoGetter::isPossibleAddinInstallForOverExcel2007(CString excelVersion)
{

    // レジストリのパスを設定する
    CString regPath(REG_PATH_EXCEL_SECURITY_SETTING);
    // パスの一部分である version 部分を置換する
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

    // キーのハンドル
    HKEY hkResult;
    // 処理結果を受け取る
    DWORD dwDisposition;
    // 関数の戻り値を格納する
    LONG lResult;

    // HKEY_CURRENT_USER\Software\Microsoft\Office\[version]\Excel\Security をオープン
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

        // エラーとして返す
        return POSSIBLE_ADDIN_INSTALL_UNEXPECTED;
    }

    // レジストリから取得した値のデータ型
    DWORD retType;
    // レジストリから取得した値のバイトサイズ
    DWORD retSize;

    DWORD retDisableAllAddins = 0;
    DWORD retRequireAddinSig  = 0;

    // DisableAllAddinsを取得
    lResult = RegQueryValueEx(hkResult
                  , REG_VALUE_NAME_EXCEL_SECURITY_DISABLEALLADDINS
                  , NULL
                  , &retType
                  , (LPBYTE)&retDisableAllAddins
                  , &retSize);

    // 値が存在しない場合
    if (lResult == ERROR_FILE_NOT_FOUND) {

        // デフォルト値を設定する
        retDisableAllAddins = 0;

    // そもそもエラーが発生した場合
    } else if (lResult != ERROR_SUCCESS) {

        // キーをクローズする
        RegCloseKey(hkResult);
        // エラーとして返す
        return POSSIBLE_ADDIN_INSTALL_UNEXPECTED;
    }

    // RequireAddinSigを取得
    lResult = RegQueryValueEx(hkResult
                  , REG_VALUE_NAME_EXCEL_SECURITY_REQUIREDADDINSIG
                  , NULL
                  , &retType
                  , (LPBYTE)&retRequireAddinSig
                  , &retSize);

    // 値が存在しない場合
    if (lResult == ERROR_FILE_NOT_FOUND) {

        // デフォルト値を設定する
        retRequireAddinSig = 0;

    // そもそもエラーが発生した場合
    } else if (lResult != ERROR_SUCCESS) {

        // キーをクローズする
        RegCloseKey(hkResult);
        // エラーとして返す
        return POSSIBLE_ADDIN_INSTALL_UNEXPECTED;
    }

    // キーをクローズする
    RegCloseKey(hkResult);

    // 正常に取得できたので、戻り値を判別する
    if (retDisableAllAddins == 0 && retRequireAddinSig == 0) {

        // DisableAllAddins と RequireAddinSig が両方とも0の場合
        // ★アドインはインストール可能
        return POSSIBLE_ADDIN_INSTALL_OK;
    
    } else {

        // ★アドインはインストールできない可能性がある
        return POSSIBLE_ADDIN_INSTALL_NG;
    }

}

int CExcelInfoGetter::delAddin(CString excelVersion, CString addinFileName)
{
	bool success = false;

	// 二つのレジストリキーエントリからアドインのファイルパスを検索して、該当する場合に削除する。
	// 一つ目
	{
		success = false;

		// レジストリのパスを設定する
		CString regPath(REG_PATH_EXCEL_ADDIN_MANAGER2);
		// パスの一部分である version 部分を置換する
		regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

		// キーのハンドル
		HKEY hkResult;
		// 処理結果を受け取る
		DWORD dwDisposition;
		// 関数の戻り値を格納する
		LONG lResult;

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
			return DEL_ADDIN_UNEXPECTED;
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

				// これ以上項目がない場合
				break;

			} else if (lResult != ERROR_SUCCESS) {

				// キーをクローズする
				RegCloseKey(hkResult);
				// エラーが発生した場合は、処理を中断する
				// 予期せぬエラー
				return DEL_ADDIN_UNEXPECTED;
			}

			// レジストリ値の名前
			CString valueNameStr(valueName);
			// レジストリ値の名前からファイル名を抽出
			CString valueNameStrOnlyFileName;

			// パスの後方からパス区切り文字を検索する
			int pos = valueNameStr.Find(_T("OPEN"));

			// 見つけた場合
			if (pos != -1) {

				TCHAR* buff = NULL;
				DWORD  buffSize = 0;
				lResult = RegQueryValueEx(hkResult
								, valueName
								, NULL
								, NULL
								, NULL
								, &buffSize);

				// 戻り値チェック
				if (lResult != ERROR_SUCCESS) {
					// キーをクローズする
					RegCloseKey(hkResult);
					// 予期せぬエラー
					return DEL_ADDIN_UNEXPECTED;
				}

				buff = new TCHAR[buffSize];

				// パスを取得する
				lResult = RegQueryValueEx(hkResult
								, valueName
								, NULL
								, NULL
								, (LPBYTE)buff
								, &buffSize);

				CString filePath(buff);
				// パスの後方からパス区切り文字を検索する
				int pos = filePath.ReverseFind('\\');

				// 見つけた場合
				if (pos != -1) {

					filePath = filePath.Mid(pos + 1);
				}


				// レジストリから取得した値とアドインファイル名が一致しているかを確認する
				if (addinFileName.Compare(filePath) == 0) {

					// アドインファイル名と一致しているのでレジストリから削除する
					lResult = RegDeleteValue(hkResult, valueName);

					// 戻り値チェック
					if (lResult != ERROR_SUCCESS) {

						// キーをクローズする
						RegCloseKey(hkResult);
						// 予期せぬエラー
						return DEL_ADDIN_UNEXPECTED;
					}

					success = true;
					break;
				}
			}

			index++;

		}

		// キーをクローズする
		RegCloseKey(hkResult);

	}

	if (!success) {

		return DEL_ADDIN_TARGET_KEY_NOT_FOUND;
	}


	// 二つ目
	{
		success = false;

		// レジストリのパスを設定する
		CString regPath(REG_PATH_EXCEL_ADDIN_MANAGER);
		// パスの一部分である version 部分を置換する
		regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

		// キーのハンドル
		HKEY hkResult;
		// 処理結果を受け取る
		DWORD dwDisposition;
		// 関数の戻り値を格納する
		LONG lResult;

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
			return DEL_ADDIN_UNEXPECTED;
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

				// これ以上項目がない場合
				break;

			} else if (lResult != ERROR_SUCCESS) {

				// キーをクローズする
				RegCloseKey(hkResult);
				// エラーが発生した場合は、処理を中断する
				// 予期せぬエラー
				return DEL_ADDIN_UNEXPECTED;
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
					return DEL_ADDIN_UNEXPECTED;
	            
				}

				success = true;
			}

			index++;

		}

		// キーをクローズする
		RegCloseKey(hkResult);

	}

	if (!success) {

		return DEL_ADDIN_TARGET_KEY_NOT_FOUND;
	}

    // 対象となるキーが見つからない場合
    return DEL_ADDIN_OK;
}

int CExcelInfoGetter::addAddin(CString excelVersion, CString addinFileName)
{
    // レジストリのパスを設定する
    CString regPath(REG_PATH_EXCEL_ADDIN_MANAGER2);
    // パスの一部分である version 部分を置換する
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

    // キーのハンドル
    HKEY hkResult;
    // 処理結果を受け取る
    DWORD dwDisposition;
    // 関数の戻り値を格納する
    LONG lResult;

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
        return ADD_ADDIN_TARGET_KEY_NOT_FOUND;
    }

	// アドインファイル名のバイト数を求める
	DWORD addinFileNameSize =  addinFileName.GetLength() * sizeof(TCHAR);

	// ファイルパスを登録する
	lResult = RegSetValueEx(hkResult, _T("OPEN"), 0, REG_SZ, (BYTE*)((LPCTSTR)addinFileName), addinFileNameSize);

    // 戻り値チェック
    if (lResult != ERROR_SUCCESS) {

        // キーをクローズする
        RegCloseKey(hkResult);
        // 予期せぬエラー
		return ADD_ADDIN_UNEXPECTED;
    
    }

    // キーをクローズする
    RegCloseKey(hkResult);
    // 対象となるキーが見つからない場合
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
    //                // Invoke名
    //                LPOLESTR name = L"Name";

    //                // InvokeするディスパッチID
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