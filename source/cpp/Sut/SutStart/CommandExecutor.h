#pragma once

class CCommandExecutor
{
public:
	
    static const int FUNC_SUCCESS = 0;
    static const int FUNC_FAILED  = 1;

	CCommandExecutor(void);
	virtual ~CCommandExecutor(void);

public:
	int exec(CString command, CString option, CString curDir);
};
