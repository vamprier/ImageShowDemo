// ImageShowDemoDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ImageShowDemo.h"
#include "ImageShowDemoDlg.h"
#include "OpenFile.h"
#include "excel.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CImageShowDemoDlg dialog

CImageShowDemoDlg::CImageShowDemoDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CImageShowDemoDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CImageShowDemoDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	image=NULL;
	m_SingleImageInfo = NULL;
	IsRightCursor = FALSE;
	IsLeftCursor = FALSE;
	//	home=NULL;
	lastnumber = 0;
	TotalRowNumLayer=0;
	TotalColNumLayer=0;
	whichHot = -1;
	m_ScreenX = 0;
	m_ScreenY = 0;
	itime=1;
	RecodeNumber = 0;
}

void CImageShowDemoDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CImageShowDemoDlg)
	DDX_Control(pDX, IDC_STATIC_IMAGE, m_Picture);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CImageShowDemoDlg, CDialog)
	//{{AFX_MSG_MAP(CImageShowDemoDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_LBUTTONDOWN()
	ON_WM_MOUSEMOVE()
	ON_WM_SETCURSOR()
	ON_WM_TIMER()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CImageShowDemoDlg message handlers

BOOL CImageShowDemoDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	//ɾ��WS_CAPTION��WS_BORDER���
	ModifyStyle(WS_CAPTION, 0);
	ModifyStyle(WS_BORDER, 0);
	m_ScreenX   =   GetSystemMetrics(   SM_CXSCREEN   ); 
	m_ScreenY   =   GetSystemMetrics(   SM_CYSCREEN   );
	//���öԻ���λ�úʹ�С
	SetWindowPos(NULL, 0, 0 , m_ScreenX, m_ScreenY, SWP_NOZORDER);
	
	CRect rect;
	GetDlgItem(IDC_STATIC_IMAGE)->GetWindowRect(rect);
	GetDlgItem(IDC_STATIC_IMAGE)->MoveWindow(0, 0,m_ScreenX, m_ScreenY);
	GetDlgItem(IDC_STATIC_IMAGE)->ModifyStyle(0, WS_CLIPSIBLINGS, 0);
	//
	int i=0,j=0;
	COpenFile m_dlg;
	if (m_dlg.DoModal() == IDCANCEL)
	{
//		DestroyWindow();
		return TRUE;
	}
	if (m_dlg.m_FilePath == "")
	{
//		DestroyWindow();
		return TRUE;
	}
	int startl = m_dlg.m_FilePath.ReverseFind('\\');
	m_strFilePath = m_dlg.m_FilePath.Mid(0,startl);
	
	ReadFileExcel(m_dlg.m_FilePath);
	m_SingleImageInfo = new SingleImageInfo[TotalRowNumLayer];
	for (i=0;i<TotalRowNumLayer;i++)
	{
		m_SingleImageInfo[i].m_IsSon = 0;
	}
	i=1;
	do
	{
		m_SingleImageInfo[RecodeNumber].m_Index = atoi(LineText[i].GetAt(0));;//ͼƬ������
		m_SingleImageInfo[RecodeNumber].m_Name = m_strFilePath + "\\" + LineText[i].GetAt(1);//ͼƬ������
		m_SingleImageInfo[RecodeNumber].m_IsFather = atoi(LineText[i].GetAt(2));//value =-1 û�и���:��value >0:ĳ�������index;
		m_SingleImageInfo[RecodeNumber].m_Direction[0] = atoi(LineText[i].GetAt(9));
		m_SingleImageInfo[RecodeNumber].m_Direction[1] = atoi(LineText[i].GetAt(10));
		m_SingleImageInfo[RecodeNumber].m_IsSon = atoi(LineText[i].GetAt(3));
		m_SingleImageInfo[RecodeNumber].m_HotRect = new float*[m_SingleImageInfo[RecodeNumber].m_IsSon+2];
		for (int index = 0;index<m_SingleImageInfo[RecodeNumber].m_IsSon+2;index++)
		{
			m_SingleImageInfo[RecodeNumber].m_HotRect[index] = new float[4];
		}
		m_SingleImageInfo[RecodeNumber].m_Nindex = new int[m_SingleImageInfo[RecodeNumber].m_IsSon+2];
		if (m_SingleImageInfo[RecodeNumber].m_IsSon != 0)
		{
			for (j=i;j<i+m_SingleImageInfo[RecodeNumber].m_IsSon;j++)
			{
				m_SingleImageInfo[RecodeNumber].m_HotRect[j-i][0] = atof(LineText[j].GetAt(4));//Left
				m_SingleImageInfo[RecodeNumber].m_HotRect[j-i][1] = atof(LineText[j].GetAt(5));//Top
				m_SingleImageInfo[RecodeNumber].m_HotRect[j-i][2] = atof(LineText[j].GetAt(6));//Right
				m_SingleImageInfo[RecodeNumber].m_HotRect[j-i][3] = atof(LineText[j].GetAt(7));//Bottom
				m_SingleImageInfo[RecodeNumber].m_Nindex[j-i] = atof(LineText[j].GetAt(8));
			}
			i+=m_SingleImageInfo[RecodeNumber].m_IsSon;
		}
		else
		{
			i++;
		}
		m_SingleImageInfo[RecodeNumber].m_HotRect[m_SingleImageInfo[RecodeNumber].m_IsSon][0]=0;//Left
		m_SingleImageInfo[RecodeNumber].m_HotRect[m_SingleImageInfo[RecodeNumber].m_IsSon][1]=0.90;//Top
		m_SingleImageInfo[RecodeNumber].m_HotRect[m_SingleImageInfo[RecodeNumber].m_IsSon][2]=0.10;//Right
		m_SingleImageInfo[RecodeNumber].m_HotRect[m_SingleImageInfo[RecodeNumber].m_IsSon][3]=1.0;//Bottom
		m_SingleImageInfo[RecodeNumber].m_HotRect[m_SingleImageInfo[RecodeNumber].m_IsSon+1][0]=0.90;//Left
		m_SingleImageInfo[RecodeNumber].m_HotRect[m_SingleImageInfo[RecodeNumber].m_IsSon+1][1]=0.90;//Top
		m_SingleImageInfo[RecodeNumber].m_HotRect[m_SingleImageInfo[RecodeNumber].m_IsSon+1][2]=1.0;//Right
		m_SingleImageInfo[RecodeNumber].m_HotRect[m_SingleImageInfo[RecodeNumber].m_IsSon+1][3]=1.0;//Bottom
		m_SingleImageInfo[RecodeNumber].m_Nindex[m_SingleImageInfo[RecodeNumber].m_IsSon] = m_SingleImageInfo[0].m_Index;
		m_SingleImageInfo[RecodeNumber].m_Nindex[m_SingleImageInfo[RecodeNumber].m_IsSon+1] = m_SingleImageInfo[RecodeNumber].m_IsFather;
		m_SingleImageInfo[RecodeNumber].m_IsSon+=2;
		RecodeNumber++;
	}
	while(i<TotalRowNumLayer);
	
	lastimage = m_SingleImageInfo[0].m_Name;
	lastnumber = 0;
	int find=lastimage.ReverseFind('.');
	CString ext;
	ext=lastimage.Right(lastimage.GetLength()-find-1);
	int type = CxImage::GetTypeIdFromName(ext);
	image = new CxImage();
	image->Load(lastimage,type);
	SetTimer(0,3000,NULL);
//	TheFile.Open("E:\\log.txt",CFile::modeCreate | CFile::modeWrite);
//	SetTimer(0,500,NULL);
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CImageShowDemoDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CImageShowDemoDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		//���öԻ���λ�úʹ�С
		SetWindowPos(NULL, 0, 0 , m_ScreenX, m_ScreenY, SWP_NOZORDER);
		CRect rect;
		GetDlgItem(IDC_STATIC_IMAGE)->GetWindowRect(rect);
		GetDlgItem(IDC_STATIC_IMAGE)->MoveWindow(0, 0,m_ScreenX, m_ScreenY);
		GetDlgItem(IDC_STATIC_IMAGE)->ModifyStyle(0, WS_CLIPSIBLINGS, 0);
		if(image != NULL)
		{
			CRect rect;
		    CRect rect1;
			m_Picture.GetClientRect(&rect);
			CDC* pDC = m_Picture.GetWindowDC();
			CWnd *pWnd = GetDlgItem(IDC_STATIC_IMAGE);//����Ϊ�ؼ�ID
			pWnd->GetClientRect(&rect1);//rcΪ�ؼ��Ĵ�С��
			image->Draw( pDC ->m_hDC,rect);    
			ReleaseDC(pDC);
		 }
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CImageShowDemoDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

BOOL CImageShowDemoDlg::PreTranslateMessage(MSG* pMsg) 
{
	if(pMsg->message==WM_KEYDOWN && pMsg->wParam==VK_ESCAPE)
	{
		if (m_SingleImageInfo != NULL)
		{
			delete []m_SingleImageInfo;
		}
		DestroyWindow();
	}
	else
	return CDialog::PreTranslateMessage(pMsg);
}

void CImageShowDemoDlg::OnLButtonDown(UINT nFlags, CPoint point) 
{
	int i = 0;
	IsInfoCursor = FALSE;
	if (m_SingleImageInfo[lastnumber].m_IsSon > 0)
	{
		for (i=0;i<m_SingleImageInfo[lastnumber].m_IsSon;i++)
		{
			CRect TempRect;
			TempRect.left = m_SingleImageInfo[lastnumber].m_HotRect[i][0]*m_ScreenX;
			TempRect.top = m_SingleImageInfo[lastnumber].m_HotRect[i][1]*m_ScreenY;
			TempRect.right = m_SingleImageInfo[lastnumber].m_HotRect[i][2]*m_ScreenX;
			TempRect.bottom = m_SingleImageInfo[lastnumber].m_HotRect[i][3]*m_ScreenY;
			if (TempRect.PtInRect(point) == TRUE)
			{
				IsInfoCursor = TRUE;
				whichHot = i;
				break;
			}
		}
	}
	if(IsInfoCursor == TRUE)
	{
		if (m_SingleImageInfo[lastnumber].m_IsFather == -1 && m_SingleImageInfo[lastnumber].m_Nindex[whichHot] == -1)
		{
			lastnumber = lastnumber;
		}
		else
		{
			lastnumber = m_SingleImageInfo[lastnumber].m_Nindex[whichHot];
		}	
		lastimage = m_SingleImageInfo[lastnumber].m_Name;
		int find=lastimage.ReverseFind('.');
		CString ext;
		ext=lastimage.Right(lastimage.GetLength()-find-1);
		int type = CxImage::GetTypeIdFromName(ext);
		if (image != NULL)
		{
			delete image;
		}
		image = new CxImage();
		image->Load(lastimage,type);
		CRect rect;
	    CRect rect1;
	    m_Picture.GetClientRect(&rect);
	    CDC* pDC = m_Picture.GetWindowDC();
	    CWnd *pWnd = GetDlgItem(IDC_STATIC_IMAGE);//����Ϊ�ؼ�ID
	    pWnd->GetClientRect(&rect1);//rcΪ�ؼ��Ĵ�С��
	    image->Draw( pDC ->m_hDC,rect);
		//Sleep(1000);
		IsInfoCursor = FALSE;
	}
	else
	{
		if (m_DownPoint.x < m_ScreenX/2 && m_SingleImageInfo[lastnumber].m_Direction[0]>-1)
		{
			if (itime >0)
			{
				lastnumber = m_SingleImageInfo[lastnumber].m_Direction[0];
				lastimage = m_SingleImageInfo[lastnumber].m_Name;
			}
			else
			{
				lastimage = m_SingleImageInfo[lastnumber].m_Name;
				if (lastnumber > RecodeNumber)
				{
					lastnumber = 0;
				}
				else
				{
					lastnumber++;
				}
			}
			int find=lastimage.ReverseFind('.');
			CString ext;
			ext=lastimage.Right(lastimage.GetLength()-find-1);
			if (ext == _T("")) return ;	
			int type = CxImage::GetTypeIdFromName(ext);
			if (image != NULL)
			{
				delete image;
			}
			image = new CxImage();
			image->Load(lastimage,type);
			CRect rect;
	        CRect rect1;
	        m_Picture.GetClientRect(&rect);
	        CDC* pDC = m_Picture.GetWindowDC();
	        CWnd *pWnd = GetDlgItem(IDC_STATIC_IMAGE);//����Ϊ�ؼ�ID
	        pWnd->GetClientRect(&rect1);//rcΪ�ؼ��Ĵ�С��
	        image->Draw( pDC ->m_hDC,rect);   
			//Sleep(1000);
		}
		if (m_DownPoint.x >  m_ScreenX/2 && m_SingleImageInfo[lastnumber].m_Direction[1]>-1)
		{
			if (itime > 0)
			{
				lastnumber = m_SingleImageInfo[lastnumber].m_Direction[1];
				lastimage = m_SingleImageInfo[lastnumber].m_Name;
			}
			else
			{
				lastimage = m_SingleImageInfo[lastnumber].m_Name;
				if (lastnumber > RecodeNumber)
				{
					lastnumber = 0;
				}
				else
				{
					lastnumber++;
				}
			}
			int find=lastimage.ReverseFind('.');
			CString ext;
			ext=lastimage.Right(lastimage.GetLength()-find-1);
			if (ext == _T("")) return ;	
			int type = CxImage::GetTypeIdFromName(ext);
			if (image != NULL)
			{
				delete image;
			}
			image = new CxImage();
			image->Load(lastimage,type);
			CRect rect;
	        CRect rect1;
	        m_Picture.GetClientRect(&rect);
	        CDC* pDC = m_Picture.GetWindowDC();
	        CWnd *pWnd = GetDlgItem(IDC_STATIC_IMAGE);//����Ϊ�ؼ�ID
	        pWnd->GetClientRect(&rect1);//rcΪ�ؼ��Ĵ�С��
	        image->Draw( pDC ->m_hDC,rect); 
			//Sleep(1000);
		}
	}
	CDialog::OnLButtonDown(nFlags, point);
}

void CImageShowDemoDlg::OnMouseMove(UINT nFlags, CPoint point) 
{
	m_ScreenX   =   GetSystemMetrics(   SM_CXSCREEN   ); 
	m_ScreenY   =   GetSystemMetrics(   SM_CYSCREEN   );
	itime = 1;
	m_DownPoint = point;
	int i=0;
	IsInfoCursor = FALSE;
	IsRightCursor = FALSE;
	IsLeftCursor = FALSE;
	whichHot = -1;
	if (m_SingleImageInfo[lastnumber].m_IsSon > 0)
	{
		for (i=0;i<m_SingleImageInfo[lastnumber].m_IsSon;i++)
		{
			CRect TempRect;
			TempRect.left = m_SingleImageInfo[lastnumber].m_HotRect[i][0]*m_ScreenX;
			TempRect.top = m_SingleImageInfo[lastnumber].m_HotRect[i][1]*m_ScreenY;
			TempRect.right = m_SingleImageInfo[lastnumber].m_HotRect[i][2]*m_ScreenX;
			TempRect.bottom = m_SingleImageInfo[lastnumber].m_HotRect[i][3]*m_ScreenY;
			if (TempRect.PtInRect(point) == TRUE)
			{
				IsInfoCursor = TRUE;
				whichHot = i;
				break;
			}
		}
	}
	CDialog::OnMouseMove(nFlags, point);
}

BOOL CImageShowDemoDlg::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message) 
{
	if (IsInfoCursor == TRUE)
	{
		HCURSOR  hCursor = AfxGetApp()->LoadCursor(IDC_CURSOR_INFO);
		::SetCursor( hCursor);//m_hcursorΪǰ�����õĹ����
		return TRUE;
	}
	else
	{
		HCURSOR  hCursor = AfxGetApp()->LoadStandardCursor(IDC_ARROW);
		::SetCursor( hCursor);//m_hcursorΪǰ�����õĹ����
		return TRUE;
	}
	return CDialog::OnSetCursor(pWnd, nHitTest, message);
}

void CImageShowDemoDlg::ReadFileExcel(CString filename)
{
	int i,j;
	CStringArray TitName;
	LineText = NULL;
	//
	AfxEnableControlContainer();
	//����Excel ������(����Excel),����app��ȫ�ֻ��Ա����
	_Application app;//����ȫ�ֵ�excel����������
	// ����Excel ������
	if (!app.CreateDispatch(_T("Excel.Application")))   
	{   
		AfxMessageBox(_T("�޷�����Excel������"));   
		return ;   
	}
	//����Excel��״̬
	app.SetVisible(FALSE); // TRUE:ʹExcel�ɼ�,FALSE:ʹExcel���ɼ�   
	app.SetUserControl(TRUE); // ���������û�����Excel 

	//VC��Excel�Ĳ���
	// �������
	Workbooks books;   
	_Workbook book;   
	Worksheets sheets;   
	_Worksheet sheet;   
	//*****
	//�õ�Worksheets 
	books.AttachDispatch(app.GetWorkbooks()); 
	book.AttachDispatch(books.Add(_variant_t(filename)));
	//�õ�Worksheets 
	sheets.AttachDispatch(book.GetWorksheets());

	LPDISPATCH lpDisp;   
	Range   range;   
	Range iCell;
	COleVariant vResult;
	COleVariant vtMissing((long)DISP_E_PARAMNOTFOUND, VT_ERROR);//_variant_t(vtMissing)((long)DISP_E_PARAMNOTFOUND, VT_ERROR); 
	//���ļ�
	lpDisp = books.Open(_T(filename),      
		_variant_t(vtMissing) , _variant_t(vtMissing) , _variant_t(vtMissing) , _variant_t(vtMissing) , _variant_t(vtMissing) ,
		_variant_t(vtMissing) , _variant_t(vtMissing) , _variant_t(vtMissing) , _variant_t(vtMissing) , _variant_t(vtMissing) ,
	_variant_t(vtMissing) , _variant_t(vtMissing) , _variant_t(vtMissing) , _variant_t(vtMissing) );
	//*****
	//�õ���ǰ��Ծsheet
	//����е�Ԫ�������ڱ༭״̬�У��˲������ܷ��أ���һֱ�ȴ�
	//�õ�Workbook
	book.AttachDispatch(lpDisp);
	lpDisp = sheets.GetItem(COleVariant(long(1)));
	sheet.AttachDispatch(lpDisp);
	//*****
	//��ȡ�Ѿ�ʹ���������Ϣ�������Ѿ�ʹ�õ���������������ʼ�С���ʼ��
	Range usedRange;
	usedRange.AttachDispatch(sheet.GetUsedRange());
	range.AttachDispatch(usedRange.GetRows());
	TotalRowNumLayer=range.GetCount();                   //�Ѿ�ʹ�õ�����
	range.AttachDispatch(usedRange.GetColumns());
	TotalColNumLayer=range.GetCount();                   //�Ѿ�ʹ�õ�����
	long iStartRow=usedRange.GetRow();               //��ʹ���������ʼ�У���1��ʼ
	long iStartCol=usedRange.GetColumn();            //��ʹ���������ʼ�У���1��ʼ
	LineText = new CStringArray[TotalRowNumLayer];
	
	for(j=1;j<TotalRowNumLayer+1;j++)
	{
		for(i=0;i<TotalColNumLayer;i++)
		{
			CString str;
			range.AttachDispatch(sheet.GetCells()); 
			range.AttachDispatch(range.GetItem (COleVariant((long)j),COleVariant((long)(i+1))).pdispVal);
			vResult =range.GetValue2();
			if(vResult.vt == VT_BSTR)       //�ַ���
			{
				str=vResult.bstrVal;
			}
			else if (vResult.vt==VT_R8)     //8�ֽڵ����� 
			{
				str.Format(_T("%f"),vResult.dblVal);
			}
			else if(vResult.vt==VT_EMPTY)   //��Ԫ��յ�
			{
				str="";
			} 
			LineText[j-1].Add(str);
			int ij = i+j*TotalColNumLayer;
		}
	}
	//*****
	//�ر����е�book���˳�Excel  
	book.SetSaved(TRUE);     // ��Workbook�ı���״̬����Ϊ�ѱ��棬������ϵͳ��ʾ�Ƿ��˹�����
	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();    // �ͷ�Range����
	sheet.ReleaseDispatch();    // �ͷ�Sheet����
	sheets.ReleaseDispatch();    // �ͷ�Sheets����
	book.Close( _variant_t(vtMissing) , _variant_t(vtMissing) , _variant_t(vtMissing) );// �ر�Workbook����
	books.Close();           // �ر�Workbooks����
	book.ReleaseDispatch();     // �ͷ�Workbook����
	books.ReleaseDispatch();    // �ͷ�Workbooks����	
	app.Quit();          // �˳�_Application
	app.ReleaseDispatch(); // �ͷ�_Application
}

void CImageShowDemoDlg::OnTimer(UINT nIDEvent) 
{
	if (nIDEvent == 0)
	{
		if(itime>0)
		{
           itime++;		   
		}
		if (itime > 10)
		{  
			itime=0;			
		}
		if(!itime)
		{
			SendMessage(WM_LBUTTONDOWN);
		}
	}
	CDialog::OnTimer(nIDEvent);
}
