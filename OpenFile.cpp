// OpenFile.cpp : implementation file
//

#include "stdafx.h"
#include "ImageShowDemo.h"
#include "OpenFile.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// COpenFile dialog


COpenFile::COpenFile(CWnd* pParent /*=NULL*/)
	: CDialog(COpenFile::IDD, pParent)
{
	//{{AFX_DATA_INIT(COpenFile)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	m_FilePath=_T("");
}


void COpenFile::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(COpenFile)
	DDX_Text(pDX, IDC_EDIT1, m_FilePath);
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(COpenFile, CDialog)
	//{{AFX_MSG_MAP(COpenFile)
	ON_BN_CLICKED(IDC_BUTTON1, OnBnClickedButtonOpen)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// COpenFile message handlers
void COpenFile::OnBnClickedButtonOpen()
{
	UpdateData();
	CFileDialog ExcDlg(TRUE,"*.xls||*.xlsx", NULL,OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT |OFN_ALLOWMULTISELECT|OFN_ENABLESIZING,
		"Excel2003 Files (*.xls)|*.xls|Excel2007 Files (*.xlsx)|*.xlsx||", NULL);
	ExcDlg.m_ofn.lpstrTitle="请选择分层文件";
	if(ExcDlg.DoModal()==IDCANCEL)  return;
	m_FilePath = ExcDlg.GetPathName();
	m_FileName = ExcDlg.GetFileName();
	UpdateData(FALSE);
}