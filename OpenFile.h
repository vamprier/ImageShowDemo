#if !defined(AFX_OPENFILE_H__3A08C837_B764_4585_8D8A_C0C3E4965EEF__INCLUDED_)
#define AFX_OPENFILE_H__3A08C837_B764_4585_8D8A_C0C3E4965EEF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// OpenFile.h : header file
//
#include "Resource.h"
/////////////////////////////////////////////////////////////////////////////
// COpenFile dialog

class COpenFile : public CDialog
{
// Construction
public:
	COpenFile(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(COpenFile)
	enum { IDD = IDD_DIALOG1 };
		// NOTE: the ClassWizard will add data members here
	CString m_FilePath;
	//}}AFX_DATA
public:
	CString m_FileName;

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(COpenFile)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	afx_msg void OnBnClickedButtonOpen();
	// Generated message map functions
	//{{AFX_MSG(COpenFile)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_OPENFILE_H__3A08C837_B764_4585_8D8A_C0C3E4965EEF__INCLUDED_)
