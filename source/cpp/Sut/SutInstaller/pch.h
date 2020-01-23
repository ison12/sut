// pch.h: プリコンパイル済みヘッダー ファイルです。
// 次のファイルは、その後のビルドのビルド パフォーマンスを向上させるため 1 回だけコンパイルされます。
// コード補完や多くのコード参照機能などの IntelliSense パフォーマンスにも影響します。
// ただし、ここに一覧表示されているファイルは、ビルド間でいずれかが更新されると、すべてが再コンパイルされます。
// 頻繁に更新するファイルをここに追加しないでください。追加すると、パフォーマンス上の利点がなくなります。

#ifndef PCH_H
#define PCH_H

// プリコンパイルするヘッダーをここに追加します
#include "framework.h"

#include <iostream>
#include <sstream>
#include <string>
#include <vector>
#include <map>

#include <tlhelp32.h> // プロセスを操作する関数

#include <msi.h> // Msi.libのリンクが必要
#pragma comment(lib, "Msi.lib")

// boost
#include "boost\shared_ptr.hpp"
#include "boost\scoped_ptr.hpp"

typedef std::basic_string<TCHAR> tstring;

#endif //PCH_H
