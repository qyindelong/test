
// MySQL_Test.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CMySQL_TestApp:
// �йش����ʵ�֣������ MySQL_Test.cpp
//

class CMySQL_TestApp : public CWinAppEx
{
public:
	CMySQL_TestApp();

// ��д
	public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CMySQL_TestApp theApp;