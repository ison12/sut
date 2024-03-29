#include "pch.h"
#include "Resource.h"
#include "ExcelInfoGetter.h"
#include "Common.h"
#include "CsvReader.h"

// ----------------------------------------------------

std::vector<std::map<tstring, tstring>> CExcelInfoGetter::EXCEL_INFO_LIST;

LPCTSTR CExcelInfoGetter::REG_PATH_EXCEL_PARAM_VERSION = _T("${version}");

LPCTSTR CExcelInfoGetter::REG_PATH_EXCEL_SECURITY_SETTING = _T("Software\\Microsoft\\Office\\${version}\\Excel\\Security");
LPCTSTR CExcelInfoGetter::REG_VALUE_NAME_EXCEL_SECURITY_DONTTRUSTINSTALLEDFILES = _T("DontTrustInstalledFiles");
LPCTSTR CExcelInfoGetter::REG_VALUE_NAME_EXCEL_SECURITY_DISABLEALLADDINS = _T("DisableAllAddins");
LPCTSTR CExcelInfoGetter::REG_VALUE_NAME_EXCEL_SECURITY_REQUIREDADDINSIG = _T("RequireAddinSig");

LPCTSTR CExcelInfoGetter::REG_PATH_EXCEL_ADDIN_MANAGER = _T("Software\\Microsoft\\Office\\${version}\\Excel\\Add-in Manager");

// ----------------------------------------------------

CExcelInfoGetter::CExcelInfoGetter(void)
{

	// excel情報リストが読み込まれていない場合、ファイルから読み込みを実施する
	if (EXCEL_INFO_LIST.size() <= 0) {

		// exeディレクトリに配置されているSutInstaller-Excel.propファイルから読み込む
		tstring appPath = common::getApplicationPath();
		appPath.append(TEXT("\\"));
		appPath.append(TEXT("SutInstaller-Excel.prop"));

		CSVReader csv(appPath);

		std::map<tstring, tstring> excelInfo;
		std::vector<tstring> tokens;

		int row = 1;
		int col = 1;
		while (csv.Read(tokens) != -1) {

			excelInfo[_T("name")]    = tokens[0];
			excelInfo[_T("clsid")]   = tokens[1];
			excelInfo[_T("version")] = tokens[2];

			EXCEL_INFO_LIST.push_back(excelInfo);
		}
	}
}

CExcelInfoGetter::~CExcelInfoGetter(void)
{

    // インストール済みExcelリストを1件ずつ処理してメモリを解放する
    size_t size = installedExcelList.size();

    for (size_t i = 0; i < size; i++) {

        delete installedExcelList.at(i);
    }
}

void CExcelInfoGetter::getExcelVersionByName(const CString& excelName, CString& excelVersion)
{

	for (std::vector<std::map<tstring, tstring>>::iterator i = EXCEL_INFO_LIST.begin(); i != EXCEL_INFO_LIST.end(); i++) {

		if ((*i)[_T("name")].c_str() == excelName) {

			excelVersion = (*i)[_T("version")].c_str();
			return;
		}

	}

	return;
}

std::vector<CExcelInfo*>& CExcelInfoGetter::getInstalledExcelApplication()
{

	for (std::vector<std::map<tstring, tstring>>::iterator i = EXCEL_INFO_LIST.begin(); i != EXCEL_INFO_LIST.end(); i++) {

        // インストールパスのサイズ
        DWORD installPathSize = _MAX_PATH;
        // インストールパス
        LPTSTR installPath = NULL;

        // インストールパスのバッファを確保する
        installPath = new TCHAR[installPathSize];

        // プロダクトコードを取得する
        TCHAR productCode[256];
		MsiGetProductCode((*i)[_T("clsid")].c_str(), productCode);

        // インストールステータス
        INSTALLSTATE installstate;
        // Excelアプリケーションのパスを取得する
        installstate = MsiGetComponentPath(productCode
										  ,(*i)[_T("clsid")].c_str()
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
            excelInfo->appName = (*i)[_T("name")];
            excelInfo->appPath = installPath;

            installedExcelList.push_back(excelInfo);

        }

        // インストールパスのバッファを解放する
        delete installPath;

    }


	/*
	 * Excelの既定のバージョンを取得する
	 */
	tstring version;
	{
        HKEY hKey = NULL;
		DWORD dwResult = ::RegOpenKeyEx(
				  HKEY_CLASSES_ROOT
				, L"CLSID\\{00024500-0000-0000-C000-000000000046}\\ProgID"
				, 0
				, KEY_QUERY_VALUE
				, &hKey
			);

		// Excelのパスが格納されているレジストリから既定値のキーを取得する
		// データの読出しバッファ
		TCHAR waReadBuf[MAX_PATH];
		// データの読出しバッファのサイズ(文字数では無くバイト数) / 読みだしたサイズ
		DWORD dwReadSize = sizeof( waReadBuf );

		dwResult = RegQueryValueEx(
			hKey
			, 0 // 既定値
			, 0
			, 0
			, (LPBYTE)waReadBuf
			, &dwReadSize);

		if (ERROR_SUCCESS == dwResult) {

			// Excelのパス
			version = waReadBuf;

			// Excelのパスは2パターンあり
			// c:\program...\Excel.exe /automation → 省略形式
			// "c:\program...\Excel.exe" /automation → 完全形式（ダブルクォーテーションで囲まれている形式）
			size_t pos = version.find_last_of(_T("."));
			if (pos != tstring::npos) {
				// 完全形式（ダブルクォーテーションで囲まれている形式）
				version = version.substr(pos + 1);
				version.append(_T(".0"));
			}
		}

		RegCloseKey(hKey);

	}
	
	/*
	 * Excelの既定のバージョンのファイルのパスを取得する
	 */
	tstring path;
	{
		// Excelのパスが格納されているレジストリを開く
        HKEY hKey = NULL;
        DWORD dwResult = ::RegOpenKeyEx(
                  HKEY_CLASSES_ROOT
                , L"CLSID\\{00024500-0000-0000-C000-000000000046}\\LocalServer32"
                , 0
                , KEY_QUERY_VALUE
                , &hKey
            );
		// Excelのパスが格納されているレジストリから既定値のキーを取得する
		// データの読出しバッファ
		TCHAR waReadBuf[MAX_PATH];
		// データの読出しバッファのサイズ(文字数では無くバイト数) / 読みだしたサイズ
		DWORD dwReadSize = sizeof( waReadBuf );

		dwResult = RegQueryValueEx(
			hKey
			, 0 // 既定値
			, 0
			, 0
			, (LPBYTE)waReadBuf
			, &dwReadSize);

		if (ERROR_SUCCESS == dwResult) {

			// Excelのパス
			path = waReadBuf;

			// Excelのパスは2パターンあり
			// c:\program...\EXCEL.EXE /automation → 省略形式
			// "c:\program...\EXCEL.EXE" /automation → 完全形式（ダブルクォーテーションで囲まれている形式）
			size_t pos = path.find(_T("EXCEL.EXE"));
			if (pos != tstring::npos) {
				// EXCEL.EXEのパスを取得する
				path = path.substr(0, pos + 9);

				size_t posDoubleQuart = path.find(_T("\""));
				if (posDoubleQuart != tstring::npos) {
					// 前方にダブルクォートが付与されている場合は、ダブルクォートを除去する
					path = path.substr(1);
				}		
			}		
		}

		RegCloseKey(hKey);
	}

	if (version.length() > 0 &&  path.length() > 0) {

		// デフォルトのExcelファイルのパスを取得する
		CExcelInfo* excelInfo = new CExcelInfo();
		excelInfo->appName = _T("Excel (Default)");
		excelInfo->appPath = path;

		installedExcelList.push_back(excelInfo);

		std::map<tstring, tstring> excelInfoMap;
		excelInfoMap[_T("name")]    = excelInfo->appName;
		//excelInfoMap[_T("clsid")]   = 0;
		excelInfoMap[_T("version")] = version;

		EXCEL_INFO_LIST.push_back(excelInfoMap);
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

int CExcelInfoGetter::completelyDeleteAddin(CString excelVersion, CString addinFileName)
{

    // レジストリのパスを設定する
    CString regPath(REG_PATH_EXCEL_ADDIN_MANAGER);
    // パスの一部分である version 部分を置換する
    regPath.Replace(REG_PATH_EXCEL_PARAM_VERSION, excelVersion);

    // 確認メッセージを表示する
    CString confirmMessage;
    confirmMessage.LoadString(IDS_INFO_COMPLETELY_DELETE_ADDIN_CONFIRM);
    confirmMessage.Append(_T("\n"));
    confirmMessage.Append(_T("HKEY_CURRENT_USER\\"));
    confirmMessage.Append(regPath);

    int messCode = AfxMessageBox(confirmMessage, MB_YESNO | MB_ICONINFORMATION);

    if (messCode == IDNO) {

        // 処理を中断する
        return COMPLETELY_DELETE_ADDIN_SUSPEND;
    }

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
        return COMPLETELY_DELETE_ADDIN_UNEXPECTED;
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
            return COMPLETELY_DELETE_ADDIN_UNEXPECTED;
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

            // 確認メッセージを表示する
            CString confirmMessage2;
            confirmMessage2.LoadString(IDS_INFO_COMPLETELY_DELETE_ADDIN_CONFIRM2);
            confirmMessage2.Append(_T("\n"));
            confirmMessage2.Append(valueName);

            int messCode2 = AfxMessageBox(confirmMessage2, MB_YESNO | MB_ICONINFORMATION);

            if (messCode2 == IDNO) {

                // 処理を中断する
                return COMPLETELY_DELETE_ADDIN_SUSPEND;
            }

            // アドインファイル名と一致しているのでレジストリから削除する
            lResult = RegDeleteValue(hkResult, valueName);

            // 戻り値チェック
            if (lResult != ERROR_SUCCESS) {

                // キーをクローズする
                RegCloseKey(hkResult);
                // 予期せぬエラー
                return COMPLETELY_DELETE_ADDIN_UNEXPECTED;
            
            }

            // キーをクローズする
            RegCloseKey(hkResult);
            // 正常終了
            return COMPLETELY_DELETE_ADDIN_OK;
        }

        index++;

    }

    // キーをクローズする
    RegCloseKey(hkResult);
    // 対象となるキーが見つからない場合
    return COMPLETELY_DELETE_ADDIN_TARGET_KEY_NOT_FOUND;
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