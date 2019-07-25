
// Excel2PDWDlg.cpp : ���� ����
//

#include "stdafx.h"
#include "Excel2PDW.h"
#include "Excel2PDWDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define DFD_FREQ_OFFSET		(1900)

FREQ_RESOL gFreqRes[ 6 ] =
{	// min, max, offset, res
	{     0,     0,											 0,    0  }, 
	{  2000,  6000,        DFD_FREQ_OFFSET,  1.25 },		/* ���뿪		*/
	{  5500, 10000,  12000-DFD_FREQ_OFFSET, -1.25 },		/* ��뿪1	*/
	{ 10000, 14000,  16000-DFD_FREQ_OFFSET, -1.25 },		/* ��뿪2	*/
	{ 14000, 18000,  12000+DFD_FREQ_OFFSET,  1.25 },		/* ��뿪3	*/
	{     0,  5000,   6300-DFD_FREQ_OFFSET, -1.25 }			/* C/D			*/
} ;

PA_RESOL gPaRes[ 6 ] =
{	// min, max, offset, res
	{ 0,    0  }, 
	{ (float) -100,  0.4 },		/* ���뿪		*/
	{ (float) -100,  0.4 },		/* ��뿪1	*/
	{ (float) -100,  0.4 },		/* ��뿪2	*/
	{ (float) -100,  0.4 },		/* ��뿪3	*/
	{ (float) -100,  0.14071, }		/* C/D			*/
} ;



// ���� ���α׷� ������ ���Ǵ� CAboutDlg ��ȭ �����Դϴ�.

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// ��ȭ ���� �������Դϴ�.
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV �����Դϴ�.

// �����Դϴ�.
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CExcel2PDWDlg ��ȭ ����




CExcel2PDWDlg::CExcel2PDWDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CExcel2PDWDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	//m_CEditOpenFile.SetWindowText( "" );
}

void CExcel2PDWDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_OPENFILE, m_CEditOpenFile);
	DDX_Control(pDX, IDC_BUTTON_CONVERT, m_CButtonConvert);
}

BEGIN_MESSAGE_MAP(CExcel2PDWDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_OPENFILE, &CExcel2PDWDlg::OnBnClickedButtonOpenfile)
	ON_BN_CLICKED(IDC_BUTTON_CONVERT, &CExcel2PDWDlg::OnBnClickedButtonConvert)
	ON_COMMAND(IDC_STATIC, &CExcel2PDWDlg::OnStaticOpenFileName)
END_MESSAGE_MAP()


// CExcel2PDWDlg �޽��� ó����

BOOL CExcel2PDWDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// �ý��� �޴��� "����..." �޴� �׸��� �߰��մϴ�.

	// IDM_ABOUTBOX�� �ý��� ��� ������ �־�� �մϴ�.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// �� ��ȭ ������ �������� �����մϴ�. ���� ���α׷��� �� â�� ��ȭ ���ڰ� �ƴ� ��쿡��
	//  �����ӿ�ũ�� �� �۾��� �ڵ����� �����մϴ�.
	SetIcon(m_hIcon, TRUE);			// ū �������� �����մϴ�.
	SetIcon(m_hIcon, FALSE);		// ���� �������� �����մϴ�.

	// TODO: ���⿡ �߰� �ʱ�ȭ �۾��� �߰��մϴ�.

	return TRUE;  // ��Ŀ���� ��Ʈ�ѿ� �������� ������ TRUE�� ��ȯ�մϴ�.
}

void CExcel2PDWDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// ��ȭ ���ڿ� �ּ�ȭ ���߸� �߰��� ��� �������� �׸�����
//  �Ʒ� �ڵ尡 �ʿ��մϴ�. ����/�� ���� ����ϴ� MFC ���� ���α׷��� ��쿡��
//  �����ӿ�ũ���� �� �۾��� �ڵ����� �����մϴ�.

void CExcel2PDWDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // �׸��⸦ ���� ����̽� ���ؽ�Ʈ�Դϴ�.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Ŭ���̾�Ʈ �簢������ �������� ����� ����ϴ�.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// �������� �׸��ϴ�.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// ����ڰ� �ּ�ȭ�� â�� ���� ���ȿ� Ŀ���� ǥ�õǵ��� �ý��ۿ���
//  �� �Լ��� ȣ���մϴ�.
HCURSOR CExcel2PDWDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CExcel2PDWDlg::OnBnClickedButtonOpenfile()
{
	// TODO: ���⿡ ��Ʈ�� �˸� ó���� �ڵ带 �߰��մϴ�.
	CString strFilter( "���� ����|*.xlsx|PDW ����|*.pdw|�������(*.*)|*.*|" );
	CFileDialog dlgFile( TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, strFilter );

	m_strFilename = "";
	if( IDOK == dlgFile.DoModal() )
	{
		if( dlgFile.GetPathName() != "" )
		{
			m_strFilename = dlgFile.GetPathName();
		}
	}

	m_CEditOpenFile.SetWindowTextA( m_strFilename );

}


void CExcel2PDWDlg::OnBnClickedButtonConvert()
{
	// TODO: ���⿡ ��Ʈ�� �˸� ó���� �ڵ带 �߰��մϴ�.
	UINT unCol;
	long l, lMaxRow;
	CString strText;

	UINT toa;
	double frq, pa, pw;

	_PDW stPDW;

	CString strConvertedFilename;

	m_CButtonConvert.EnableWindow( FALSE );

	if( m_strFilename.IsEmpty() == false ) {
		CFile theFile;

		CXlSimpleAutomation XL( m_strFilename ); 

		strConvertedFilename = m_strFilename + ".pdw";
		if( TRUE == theFile.Open( strConvertedFilename, CFile::modeCreate | CFile::typeBinary | CFile::modeWrite ) ) {
			lMaxRow = XL.GetRowNum();
			for( l=2 ; l <= lMaxRow; ++l ) {
				unCol = 2;

				GetCellValue( & XL, unCol, l, & toa, & pw, & pa, & frq );

				memset( & stPDW, 0, sizeof(_PDW) );
				stPDW.item.stat = PDW_NORMAL;
				stPDW.item.amplitude = EncodePA( pa );
				stPDW.item.freq = EncodeFreq( frq );
				stPDW.item.pw = EncodePW( pw );
				stPDW.item.toa = EncodeTOA( toa );
				stPDW.item.band = EncodeBand( frq );

				theFile.Write( & stPDW, sizeof(_PDW) );

			}

			theFile.Close();
		}
		else {
			CString strErrMsg;

			MessageBox( "AAAA" );
		}
	}

	m_CButtonConvert.EnableWindow( TRUE );
}

UINT CExcel2PDWDlg::EncodeTOA( double dValue )
{
	UINT uiValue;

	uiValue = (UINT) ( dValue / 25. + 0.5 );
	return uiValue;

}

UINT CExcel2PDWDlg::EncodePW( double dValue )
{
	UINT uiValue;

	uiValue = (UINT) ( dValue / 25. + 0.5 );
	return uiValue;
}

UINT CExcel2PDWDlg::EncodeFreq( double dValue )
{
	int i;
	UINT uiValue;
	double dTemp;

	uiValue = (UINT) ( dValue + 0.5 );
	for( i=1 ; i <= 4 ; ++i ) {
		if( uiValue >= gFreqRes[i].min && uiValue <= gFreqRes[i].max ) {
			break;
		}
	}

	if( i <= 4 ) {
		dTemp = ( dValue - gFreqRes[i].offset ) / gFreqRes[i].res;
		uiValue = (UINT) ( dTemp + 0.5 );
	}
	else {
		uiValue = 0;
	}

	return uiValue;
}

UINT CExcel2PDWDlg::EncodeBand( double dValue )
{
	int i;
	UINT uiValue;
	double dTemp;

	uiValue = (UINT) ( dValue + 0.5 );
	for( i=1 ; i <= 4 ; ++i ) {
		if( uiValue >= gFreqRes[i].min && uiValue <= gFreqRes[i].max ) {
			break;
		}
	}

	if( i <= 4 ) {
		uiValue = (UINT) i;
	}
	else {
		uiValue = 0;
	}

	return uiValue;
}

UINT CExcel2PDWDlg::EncodePA( double dValue )
{
	UINT uiValue;

	uiValue = dValue - gPaRes[ 1 ].offset;
	if( uiValue > 0 ) {
		uiValue = (UINT) ( uiValue / gPaRes[ 1 ].res + 0.5 );
	}
	else {
		uiValue = 0;
	}
	return uiValue;

}

void CExcel2PDWDlg::OnStaticOpenFileName()
{
	// TODO: ���⿡ ��� ó���� �ڵ带 �߰��մϴ�.


}

void CExcel2PDWDlg::GetCellValue( CXlSimpleAutomation *pXL, UINT unCol, long lRow, UINT *pToa, double *pPw, double *pPa, double *pFreq )
{
	CString strText;

	pXL->GetCellValue( unCol++, lRow, strText ); 
	*pToa = atoi( strText );

	pXL->GetCellValue( unCol++, lRow, strText ); 
	*pPw = atof( strText );

	pXL->GetCellValue( unCol++, lRow, strText ); 
	*pPa = atof( strText );

	pXL->GetCellValue( unCol++, lRow, strText ); 
	*pFreq = atoi( strText ) / 1000.;
}
