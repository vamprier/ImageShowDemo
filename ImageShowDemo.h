// ImageShowDemo.h : main header file for the IMAGESHOWDEMO application
//

#if !defined(AFX_IMAGESHOWDEMO_H__0AFF0453_58E6_4E3C_94A1_94850E886EC3__INCLUDED_)
#define AFX_IMAGESHOWDEMO_H__0AFF0453_58E6_4E3C_94A1_94850E886EC3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CImageShowDemoApp:
// See ImageShowDemo.cpp for the implementation of this class
//

class CImageShowDemoApp : public CWinApp
{
public:
	CImageShowDemoApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CImageShowDemoApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL
public:
    void DivDisplay();
// Implementation

	//{{AFX_MSG(CImageShowDemoApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_IMAGESHOWDEMO_H__0AFF0453_58E6_4E3C_94A1_94850E886EC3__INCLUDED_)
