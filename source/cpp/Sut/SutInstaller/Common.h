#ifndef SEZARWIN_COMMON_H_
#define SEZARWIN_COMMON_H_

#include <string>
#include <vector>
#include <iostream>
#include <comutil.h>
#include <windows.h>

// デバッグモードの有無で出力のON・OFFを切り替える
#ifdef _DEBUG
// デバッグモードの場合

    // UNICODE文字セットの場合
#ifdef UNICODE
#define C_OUT(message)    std::wcout << message
#define C_OUT_NL(message) std::wcout << message << std::endl

// マルチバイト文字セットの場合
#else
#define C_OUT(message)     std::cout << message
#define C_OUT_NL(message)  std::cout << message << std::endl

#endif

#else
// デバッグモードではない場合

    // 何も出力しない
#define C_OUT(message)
#define C_OUT_NL(message)
#endif

/**
 * [概要]     : 共通処理を定義。
 *
 *
 * [備考]     :
 *
 * [作成者]   : Sandora
 * [履歴]     : 2007/08/13   Sandora  新規作成
 *            :
 *
 * Copyright(c)2007 Sandora All rights reserved.
 *
 */
namespace common {

    /**
        * 文字列の置換。
        *
        * @param str 置換対象文字列
        * @param from 検索文字列
        * @param to   置換文字列
        */
    template<class T> void replaceStr(std::basic_string<T>& str
        , const std::basic_string<T>& from
        , const std::basic_string<T>& to)
    {

        std::basic_string<T>::size_type pos = 0;

        while (pos = str.find(from, pos), pos != std::string::npos) {

            str.replace(pos, from.length(), to);

            pos += to.length();
        }

    }

    /**
        * 与えられた引数の最小値を求める。
        *
        * @param n1
        * @param n2
        * @param n3
        * @param n4
        *
        * @return 最小値
        */
    int choiceMinNum(int n1, int n2, int n3, int n4);

    /**
        * 与えられた引数の最大値を求める。
        *
        * @param n1
        * @param n2
        * @param n3
        * @param n4
        *
        * @return 最大値
        */
    int choiceMaxNum(int n1, int n2, int n3, int n4);

    /**
        * 与えられた引数の最大値を求める。
        *
        * @param int配列
        *
        * @return 最大値
        */
    int choiceMaxNum(int*);

    /**
        * exeファイルが置かれているパスを取得する。
        *
        * @return exeファイルのパス
        */
    std::basic_string<TCHAR> getApplicationPath();

    /**
        * モジュールファイル（exeやdll）が置かれているパスを取得する。
        *
        * @param hModule モジュールハンドル
        * @return ファイルパス
        */
    std::basic_string<TCHAR> getModulePath(HMODULE hModule);

    /**
        * 最後に発生したエラーメッセージを表示する
        */
    std::basic_string<TCHAR> getErrorMessage(DWORD errorNo);

    /**
        * 最後に発生したエラーメッセージを表示する
        */
    std::basic_string<TCHAR> getLastErrorMessage();

    /**
        * 最後に発生したエラーメッセージを表示する
        */
    void showErrorMessage(DWORD errorNo);

    /**
        * 最後に発生したエラーメッセージを表示する
        */
    void showLastErrorMessage();

    /**
        * 全てのディスプレイデバイスのグラフィックスモードに関する情報を取得します。
        *
        */
    std::vector<DEVMODE> getDisplaySettingsInfo();

    /**
        * 全てのディスプレイデバイスのグラフィックスモードに関する情報を出力します。
        *
        */
    void outDisplaySettingsInfo();

    /**
        * 任意の色に対して、輝度を加味した色を計算する。
        *
        * @param r 元色（赤成分）
        * @param g 元色（緑成分）
        * @param b 元色（青成分）
        * @param bgLight 輝度 【0（真っ暗） <= 1（通常） <= ∞（明るい）】
        * @return 計算後の色
        */
    COLORREF calcBgLight(int r, int g, int b, double bgLight);

    /**
        * 任意の色に対して、輝度を加味した色を計算する。
        *
        * @param r 元色（赤成分）
        * @param g 元色（緑成分）
        * @param b 元色（青成分）
        * @param bgLight 輝度 【0（真っ暗） <= 1（通常） <= ∞（明るい）】
        * @param calcR 計算後の色（赤成分）
        * @param calcG 計算後の色（緑成分）
        * @param calcB 計算後の色（青成分）
        * @return 計算後の色
        */
    void calcBgLight(int r, int g, int b, double bgLight, int& calcR, int& calcG, int& calcB);

    /**
        * 二つの色を合成する。
        *
        * @param sr 前景（赤成分）
        * @param sg 前景（緑成分）
        * @param sb 前景（青成分）
        * @param dr 背景（赤成分）
        * @param dg 背景（緑成分）
        * @param db 背景（青成分）
        * @param alpha 透過率 0が透明 1が不透明
        * @return 計算後の色
        */
    COLORREF calcAlphaBlend(int sr, int sg, int sb, int dr, int dg, int db, double alpha);

    /**
        * 二つの色を合成する。
        *
        * @param sr 前景（赤成分）
        * @param sg 前景（緑成分）
        * @param sb 前景（青成分）
        * @param dr 背景（赤成分）
        * @param dg 背景（緑成分）
        * @param db 背景（青成分）
        * @param alpha 透過率 0が透明 1が不透明
        * @param calcR 計算後の色（赤成分）
        * @param calcG 計算後の色（緑成分）
        * @param calcB 計算後の色（青成分）
        * @return 計算後の色
        */
    void calcAlphaBlend(int sr, int sg, int sb, int dr, int dg, int db, double alpha, int& calcR, int& calcG, int& calcB);

    SAFEARRAY* createSafeArrayOneDim(VARENUM type, int size);

    SAFEARRAY* createSafeArrayTwoDim(VARENUM type, int size1, int size2);

    void initSafeArrayOneDim(SAFEARRAY* var);

    void initSafeArrayTwoDim(SAFEARRAY* var);

};

#endif /*COMMON_H_*/
