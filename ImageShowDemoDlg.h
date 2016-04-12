// ImageShowDemoDlg.h : header file
//

#if !defined(AFX_IMAGESHOWDEMODLG_H__D5E7D11D_0664_4ACB_A396_A73540CEC803__INCLUDED_)
#define AFX_IMAGESHOWDEMODLG_H__D5E7D11D_0664_4ACB_A396_A73540CEC803__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "ximage.h"

/////////////////////////////////////////////////////////////////////////////
// CImageShowDemoDlg dialog
struct SingleImageInfo
{
	int m_Index;//图片索引号
	CString m_Name;//图片的名称
	int m_IsFather;//value =-1 没有父节点；value >0:某个具体的index;
	int m_IsSon;//value = 0:无子节点；value>0:有m_IsSon个子节点
	float** m_HotRect;//热点的矩形区域数组
	int* m_Nindex;//热点区域对应图片索引号
	int m_Direction[2];//m_Direction[0]该图的左图片;m_Direction[1]该图的右图片
};

class CImageShowDemoDlg : public CDialog
{
// Construction
public:
	CImageShowDemoDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CImageShowDemoDlg)
	enum { IDD = IDD_IMAGESHOWDEMO_DIALOG };
	CStatic	m_Picture;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CImageShowDemoDlg)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL
public:
	CxImage *image;
	//	CxImage *information;
	//	CxImage *home;
	CString lastimage;
	int lastnumber;
	int whichHot;
	CPoint m_DownPoint;
	BOOL IsRightCursor;
	BOOL IsLeftCursor;
	BOOL IsInfoCursor;
	BOOL IsHitHot;
	int m_ScreenX;
	int m_ScreenY;
	CString m_strFilePath;
	CStringArray *LineText;
	SingleImageInfo* m_SingleImageInfo;
	int TotalRowNumLayer;
	int TotalColNumLayer;
	int RecodeNumber;
	void ReadFileExcel(CString FilePath);
//	CStdioFile TheFile;
//	CString filestr;
	int itime;
// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CImageShowDemoDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
	afx_msg void OnTimer(UINT nIDEvent);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_IMAGESHOWDEMODLG_H__D5E7D11D_0664_4ACB_A396_A73540CEC803__INCLUDED_)
