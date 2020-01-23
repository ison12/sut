#include "pch.h"
#include "Common.h"

int common::choiceMinNum(int n1, int n2, int n3, int n4) {

    int ret = 0;

    if (n1 <= n2 && n1 <= n3 && n1 <= n4) {

        ret = n1;

    }
    else if (n2 <= n1 && n2 <= n3 && n2 <= n4) {

        ret = n2;

    }
    else if (n3 <= n1 && n3 <= n2 && n3 <= n4) {

        ret = n3;

    }
    else {

        ret = n4;
    }

    return ret;
}

int common::choiceMaxNum(int n1, int n2, int n3, int n4) {

    int ret = 0;

    if (n1 >= n2 && n1 >= n3 && n1 >= n4) {

        ret = n1;

    }
    else if (n2 >= n1 && n2 >= n3 && n2 >= n4) {

        ret = n2;

    }
    else if (n3 >= n1 && n3 >= n2 && n3 >= n4) {

        ret = n3;

    }
    else {

        ret = n4;
    }

    return ret;
}

int common::choiceMaxNum(int* var) {

    int ret = 0;

    int length = sizeof(var) / sizeof(var[0]);

    for (int i = 0; i < length; i++) {

        if (ret < var[i]) {

            ret = var[i];
        }
    }

    return ret;
}

std::basic_string<TCHAR> common::getApplicationPath() {

    TCHAR path[_MAX_PATH];

    // 実行ファイルのファイルパス（ファイル名含む）を取得する
    GetModuleFileName(NULL, path, sizeof(path));

    // string型に入れ替え
    std::basic_string<TCHAR> strPath(path);

    // ファイルパスのみ除外して、戻り値とする
    return strPath.substr(0, strPath.find_last_of(_T("\\")));

}

std::basic_string<TCHAR> common::getModulePath(HMODULE hModule) {

    TCHAR path[_MAX_PATH];

    // 実行ファイルのファイルパス（ファイル名含む）を取得する
    GetModuleFileName(hModule, path, sizeof(path));

    // string型に入れ替え
    std::basic_string<TCHAR> strPath(path);

    // ファイルパスのみ除外して、戻り値とする
    return strPath.substr(0, strPath.find_last_of(_T("\\")));

}

std::basic_string<TCHAR> common::getErrorMessage(DWORD errorNo) {

    // メッセージバッファ
    LPTSTR lpMsgBuf;

    // エラーNoからメッセージを取得する
    FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER
        | FORMAT_MESSAGE_FROM_SYSTEM
        | FORMAT_MESSAGE_IGNORE_INSERTS
        , NULL
        , errorNo
        , MAKELANGID(LANG_NEUTRAL
            , SUBLANG_DEFAULT)
        , (LPTSTR)&lpMsgBuf
        , 0
        , NULL);

    // 戻り値
    std::basic_string<TCHAR> message(lpMsgBuf);

    // メッセージバッファを解放する
    LocalFree(lpMsgBuf);

    return message;
}

std::basic_string<TCHAR> common::getLastErrorMessage() {

    // 最後に発生したエラーのエラーコードからメッセージを取得する
    return getErrorMessage(GetLastError());
}


void common::showErrorMessage(DWORD errorNo) {

    // メッセージを取得する
    std::basic_string<TCHAR> message = getErrorMessage(errorNo);

    // メッセージ表示
    MessageBox(NULL
        , message.c_str()
        , NULL
        , MB_OK
        | MB_ICONERROR);

}

void common::showLastErrorMessage() {

    // 最後に発生したエラーのメッセージを表示する
    showErrorMessage(GetLastError());
}

std::vector<DEVMODE> common::getDisplaySettingsInfo() {

    std::vector<DEVMODE> list;

    DEVMODE tmp;

    // EnumDisplaySettingsの戻り値
    int ret = 1;

    // 列挙するディスプレイ情報のインデックス
    int i = 0;

    while (ret) {

        ret = EnumDisplaySettings(
            NULL,
            i,
            &tmp
        );

        list.push_back(tmp);

        i++;
    }

    return list;
}

void common::outDisplaySettingsInfo() {

    std::vector<DEVMODE> list = common::getDisplaySettingsInfo();

    std::cout << "■ディスプレイデバイス情報の列挙 " << std::endl;

    // 全てのディスプレイ情報を取得する
    for (std::vector<DEVMODE>::iterator i = list.begin(); i != list.end(); i++) {

        // ディスプレイモード構造体を取得
        DEVMODE tmp = (*i);

        std::cout << "◆ディスプレイ " << tmp.dmDeviceName << std::endl;

        std::cout << "　幅　　：" << tmp.dmPelsWidth << std::endl;
        std::cout << "　高さ  ：" << tmp.dmPelsHeight << std::endl;
        std::cout << "　色深度：" << tmp.dmBitsPerPel << std::endl;

    }

    std::cout << "■■ " << std::endl;

}

COLORREF common::calcBgLight(int r, int g, int b, double bgLight) {

    // 0の場合、計算ができないため１にする
    if (r == 0) r = 1;
    if (g == 0) g = 1;
    if (b == 0) b = 1;

    int calcR = (int)((double)r * bgLight);
    int calcG = (int)((double)g * bgLight);
    int calcB = (int)((double)b * bgLight);

    if (calcR < 0)   calcR = 0;
    if (calcG < 0)   calcG = 0;
    if (calcB < 0)   calcB = 0;
    if (calcR > 255) calcR = 255;
    if (calcG > 255) calcG = 255;
    if (calcB > 255) calcB = 255;

    return RGB(calcR, calcG, calcB);
}

void common::calcBgLight(int r, int g, int b, double bgLight, int& calcR, int& calcG, int& calcB) {

    // 0の場合、計算ができないため１にする
    if (r == 0) r = 1;
    if (g == 0) g = 1;
    if (b == 0) b = 1;

    calcR = (int)((double)r * bgLight);
    calcG = (int)((double)g * bgLight);
    calcB = (int)((double)b * bgLight);

    if (calcR < 0)   calcR = 0;
    if (calcG < 0)   calcG = 0;
    if (calcB < 0)   calcB = 0;
    if (calcR > 255) calcR = 255;
    if (calcG > 255) calcG = 255;
    if (calcB > 255) calcB = 255;

}

COLORREF common::calcAlphaBlend(int sr, int sg, int sb, int dr, int dg, int db, double alpha) {

    int calcR = (int)(((double)dr * (1.0 - alpha)) + ((double)sr * alpha));
    int calcG = (int)(((double)dg * (1.0 - alpha)) + ((double)sg * alpha));
    int calcB = (int)(((double)db * (1.0 - alpha)) + ((double)sb * alpha));

    if (calcR < 0)   calcR = 0;
    if (calcG < 0)   calcG = 0;
    if (calcB < 0)   calcB = 0;
    if (calcR > 255) calcR = 255;
    if (calcG > 255) calcG = 255;
    if (calcB > 255) calcB = 255;

    return RGB(calcR, calcG, calcB);
}

void common::calcAlphaBlend(int sr, int sg, int sb, int dr, int dg, int db, double alpha, int& calcR, int& calcG, int& calcB) {

    calcR = (int)(((double)dr * (1.0 - alpha)) + ((double)sr * alpha));
    calcG = (int)(((double)dg * (1.0 - alpha)) + ((double)sg * alpha));
    calcB = (int)(((double)db * (1.0 - alpha)) + ((double)sb * alpha));

    if (calcR < 0)   calcR = 0;
    if (calcG < 0)   calcG = 0;
    if (calcB < 0)   calcB = 0;
    if (calcR > 255) calcR = 255;
    if (calcG > 255) calcG = 255;
    if (calcB > 255) calcB = 255;

}

SAFEARRAY* common::createSafeArrayOneDim(VARENUM type, int size)
{
    // 配列情報を設定する(一次元)
    // 配列情報構造体
    SAFEARRAYBOUND rgb[1];
    // サイズを設定する
    rgb[0].cElements = size;
    // 下限値を設定する
    rgb[0].lLbound = 0;

    // 配列を生成する
    SAFEARRAY* psa = SafeArrayCreate(type, 1, rgb);

    // 配列の生成に成功したかを判定する
    if (!psa) {

        return NULL;
    }

    return psa;
}

SAFEARRAY* common::createSafeArrayTwoDim(VARENUM type, int size1, int size2)
{
    // 配列情報を設定する(二次元)
    // 配列情報構造体
    SAFEARRAYBOUND rgb[2];
    // サイズを設定する
    rgb[0].cElements = size1;
    // 下限値を設定する
    rgb[0].lLbound = 0;
    // サイズを設定する
    rgb[1].cElements = size2;
    // 下限値を設定する
    rgb[1].lLbound = 0;

    // 配列を生成する
    SAFEARRAY* psa = SafeArrayCreate(type, 2, rgb);

    // 配列の生成に成功したかを判定する
    if (!psa) {

        return NULL;
    }

    return psa;
}

void common::initSafeArrayOneDim(SAFEARRAY* var)
{
    // 文字列であることを前提としてテストデータを設定する
    BSTR* sData;

    // データにアクセスする
    HRESULT hr = SafeArrayAccessData(var, (void**)&sData);

    if (S_OK != hr) {

        return;
    }

    // 配列情報構造体
    SAFEARRAYBOUND bound = var->rgsabound[0];

    for (ULONG i = 0; i < bound.cElements; i++) {

        // ワイド文字列ストリーム
        std::wstringstream wsstream;
        // 文字列を設定する
        wsstream << L"データ" << i << std::endl;
        // 配列に文字列を設定する
        sData[i] = SysAllocString(wsstream.str().c_str());
    }

    // データアクセスを解放する
    hr = SafeArrayUnaccessData(var);

    if (S_OK != hr) {

        return;
    }


}

void common::initSafeArrayTwoDim(SAFEARRAY* var)
{
    // 配列をロックする
    HRESULT hr = SafeArrayLock(var);

    if (S_OK != hr) {

        return;
    }

    // 配列情報構造体
    SAFEARRAYBOUND bound1 = var->rgsabound[0];
    SAFEARRAYBOUND bound2 = var->rgsabound[1];

    // 文字列であることを前提としてテストデータを設定する
    BSTR* sData;
    // 要素位置
    long indices[2];

    for (ULONG i = 0; i < bound1.cElements; i++) {

        for (ULONG j = 0; j < bound2.cElements; j++) {

            // ワイド文字列ストリーム
            std::wstringstream wsstream;

            // 文字列を設定する
            wsstream << L"データ" << i << L"-" << j << std::endl;

            //要素位置[i][j]への値セット
            indices[0] = j;
            indices[1] = i;

            // 配列からデータポインタを取得する
            SafeArrayPtrOfIndex(var, indices, (void HUGEP * FAR*) & sData);

            // 配列に文字列を設定する
            *sData = SysAllocString(wsstream.str().c_str());
        }

    }

    // 配列のロックを解放する
    hr = SafeArrayUnlock(var);

    if (S_OK != hr) {

        return;
    }


}
