
// Excel2PDW.h : PROJECT_NAME ���� ���α׷��� ���� �� ��� �����Դϴ�.
//

#pragma once

#ifndef __AFXWIN_H__
	#error "PCH�� ���� �� ������ �����ϱ� ���� 'stdafx.h'�� �����մϴ�."
#endif

#include "resource.h"		// �� ��ȣ�Դϴ�.


// CExcel2PDWApp:
// �� Ŭ������ ������ ���ؼ��� Excel2PDW.cpp�� �����Ͻʽÿ�.
//

class CExcel2PDWApp : public CWinApp
{
public:
	CExcel2PDWApp();

// �������Դϴ�.
public:
	virtual BOOL InitInstance();

// �����Դϴ�.

	DECLARE_MESSAGE_MAP()
};

extern CExcel2PDWApp theApp;