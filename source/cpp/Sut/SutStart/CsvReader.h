/**
 * CSVファイル読み込みクラス
 * @author      台北猫々
 * @version     CVS $Id: CSVReader.h,v 1.1 2008/03/26 12:45:24 tamamo Exp $
 * @license     BSD license:
 * Copyright (c) 2008, Taipei Cat Project
 * All rights reserved.
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of the Taipei Cat Project nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

#ifndef _CSVREADER_H__
#define _CSVREADER_H__

#include <string>
#include <vector>
#include <fstream>
using namespace std;

typedef basic_string<TCHAR> tstring;
typedef basic_fstream<TCHAR> tfstream;

typedef boost::shared_ptr<tfstream> tfstreamPtr;

#define DEFAULT_SEPARATOR	','
#define DEFAULT_QUOTE_CHARACTER	'"'

class CSVReader
{
public:

	/**
		* コンストラクタ
		* @param stream ファイルストリーム
		* @comment セパレータ(,), エンクオート(")
		*/
	CSVReader(tfstream& stream);

	/**
		* コンストラクタ
		* @param stream ファイルストリーム
		* @param sep セパレータ
		* @comment エンクオート(")
		*/
	CSVReader(tfstream& stream, const TCHAR sep);

	/**
		* コンストラクタ
		* @param stream ファイルストリーム
		* @param sep セパレータ
		* @param quo エンクオート
		*/
	CSVReader(tfstream& stream, const TCHAR sep, const TCHAR quo);

	/**
		* コンストラクタ
		* @param filepath ファイルパス
		* @comment セパレータ(,), エンクオート(")
		*/
	CSVReader(const tstring& filepath);

	/**
		* コンストラクタ
		* @param filepath ファイルパス
		* @param sep セパレータ
		* @comment エンクオート(")
		*/
	CSVReader(const tstring& filepath, const TCHAR sep);

	/**
		* コンストラクタ
		* @param filepath ファイルパス
		* @param sep セパレータ
		* @param quo エンクオート
		*/
	CSVReader(const tstring& filepath, const TCHAR sep, const TCHAR quo);

	/**
		* デストラクタ
		*/
	virtual ~CSVReader(void);

	/**
		* CSVファイルを１行読み込んで、分割して配列で返します。
		* @param tokens トークン(OUT)
		* @return 0:正常 -1:EOF
		*/
	int Read(vector<tstring>& tokens);

	/**
		* ファイルストリームをクローズします。
		* @return 0:正常 -1:異常
		*/
	int Close(void);

private:

	/**
		* ファイルストリームをオープンします。
		* @param filepath ファイルパス
		* @return 0:正常 -1:異常
		*/
	int Open(const tstring& filepath);

	/**
		* ファイルから１行読み込みます。
		* @param line 行データ
		* @return >=0：読み込んだデータ長 -1：EOF
		*/
	int GetNextLine(tstring& line);

	/**
		* データをパースします。
		* @param nextLine 行データ
		* @param tokens パースしたデータの配列(OUT)
		* @return 0
		*/
	int Parse(tstring& nextLine, vector<tstring>& tokens);

	tfstreamPtr pstream;
	TCHAR SEPARATOR;
	TCHAR QUOTE;

};

#endif
