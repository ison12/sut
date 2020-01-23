
// SutInstaller.h : PROJECT_NAME アプリケーションのメイン ヘッダー ファイルです
//

#pragma once

#ifndef __AFXWIN_H__
	#error "PCH に対してこのファイルをインクルードする前に 'pch.h' をインクルードしてください"
#endif

#include "resource.h"		// メイン シンボル


// CSutInstallerApp:
// このクラスの実装については、SutInstaller.cpp を参照してください
//

class CSutInstallerApp : public CWinApp
{
public:
	CSutInstallerApp();

// オーバーライド
public:
	virtual BOOL InitInstance();

// 実装

	DECLARE_MESSAGE_MAP()
};

extern CSutInstallerApp theApp;
