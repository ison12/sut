/**
 * CSVファイル読み込みクラス
 * @author      台北猫々
 * @version     CVS $Id: CSVReader.cpp,v 1.1 2008/03/26 12:45:24 tamamo Exp $
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

#include "pch.h"
#include "CSVReader.h"

CSVReader::CSVReader(tfstream& stream) :
	SEPARATOR(DEFAULT_SEPARATOR),
	QUOTE(DEFAULT_QUOTE_CHARACTER),
	pstream(&stream)
{
}

CSVReader::CSVReader(tfstream& stream, const TCHAR sep) :
	SEPARATOR(sep),
	QUOTE(DEFAULT_QUOTE_CHARACTER),
	pstream(&stream)
{
}

CSVReader::CSVReader(tfstream& stream, const TCHAR sep, const TCHAR quo) :
	SEPARATOR(sep),
	QUOTE(quo),
	pstream(&stream)
{
}

CSVReader::CSVReader(const tstring& filepath) :
	SEPARATOR(DEFAULT_SEPARATOR),
	QUOTE(DEFAULT_QUOTE_CHARACTER)
{
	this->Open(filepath);
}

CSVReader::CSVReader(const tstring& filepath, const TCHAR sep) :
	SEPARATOR(sep),
	QUOTE(DEFAULT_QUOTE_CHARACTER)
{
	Open(filepath);
}

CSVReader::CSVReader(const tstring& filepath, const TCHAR sep, const TCHAR quo) :
	SEPARATOR(sep),
	QUOTE(quo)
{
	Open(filepath);
}

CSVReader::~CSVReader(void)
{
	Close();
}

int CSVReader::Read(vector<tstring>& tokens) {
	tokens.clear();

	tstring nextLine;
	if (GetNextLine(nextLine) <= 0) {
		return -1;
	}
	Parse(nextLine, tokens);
	return 0;
}

int CSVReader::GetNextLine(tstring& line) {

	if (!pstream || pstream->eof()) {
		return -1;
	}
	std::getline(*pstream, line);
	return (int)line.length();
}

int CSVReader::Parse(tstring& nextLine, vector<tstring>& tokens) {
	tstring token;
	bool interQuotes = false;
	do {
		if (interQuotes) {
			token += _T('\n');
			if (GetNextLine(nextLine) < 0) {
				break;
			}
		}

		for (int i = 0; i < (int)nextLine.length(); i++) {

			TCHAR c = nextLine.at(i);
			if (c == QUOTE) {
				if (interQuotes
					&& (int)nextLine.length() > (i + 1)
					&& nextLine.at(i + 1) == QUOTE) {
					token += nextLine.at(i + 1);
					i++;
				}
				else {
					interQuotes = !interQuotes;
					if (i > 2
						&& nextLine.at(i - 1) != SEPARATOR
						&& (int)nextLine.length() > (i + 1)
						&& nextLine.at(i + 1) != SEPARATOR
						) {
						token += c;
					}
				}
			}
			else if (c == SEPARATOR && !interQuotes) {
				tokens.push_back(token);
				token.clear();
			}
			else {
				token += c;
			}
		}
	} while (interQuotes);
	tokens.push_back(token);
	return 0;
}

int CSVReader::Open(const tstring& filepath)
{
	int ret = 0;

	if (!pstream) {
		pstream = tfstreamPtr(new tfstream(filepath.c_str(), std::ios::in));

		ret = pstream->is_open();
		if (!ret) {
			pstream = tfstreamPtr();
		}
	}

	return ret;
}

int CSVReader::Close(void) {
	if (pstream) {
		pstream->close();
		pstream = tfstreamPtr();
	}
	return 0;
}
