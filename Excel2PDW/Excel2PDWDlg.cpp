
// Excel2PDWDlg.cpp : 구현 파일
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
	{  2000,  6000,        DFD_FREQ_OFFSET,  1.25 },		/* 저대역		*/
	{  5500, 10000,  12000-DFD_FREQ_OFFSET, -1.25 },		/* 고대역1	*/
	{ 10000, 14000,  16000-DFD_FREQ_OFFSET, -1.25 },		/* 고대역2	*/
	{ 14000, 18000,  12000+DFD_FREQ_OFFSET,  1.25 },		/* 고대역3	*/
	{     0,  5000,   6300-DFD_FREQ_OFFSET, -1.25 }			/* C/D			*/
} ;

PA_RESOL gPaRes[ 6 ] =
{	// min, max, offset, res
	{ 0,    0  }, 
	{ (float) -100,  0.4 },		/* 저대역		*/
	{ (float) -100,  0.4 },		/* 고대역1	*/
	{ (float) -100,  0.4 },		/* 고대역2	*/
	{ (float) -100,  0.4 },		/* 고대역3	*/
	{ (float) -100,  0.14071, }		/* C/D			*/
} ;



// 응용 프로그램 정보에 사용되는 CAboutDlg 대화 상자입니다.

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 대화 상자 데이터입니다.
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 지원입니다.

// 구현입니다.
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


// CExcel2PDWDlg 대화 상자




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


// CExcel2PDWDlg 메시지 처리기

BOOL CExcel2PDWDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 시스템 메뉴에 "정보..." 메뉴 항목을 추가합니다.

	// IDM_ABOUTBOX는 시스템 명령 범위에 있어야 합니다.
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

	// 이 대화 상자의 아이콘을 설정합니다. 응용 프로그램의 주 창이 대화 상자가 아닐 경우에는
	//  프레임워크가 이 작업을 자동으로 수행합니다.
	SetIcon(m_hIcon, TRUE);			// 큰 아이콘을 설정합니다.
	SetIcon(m_hIcon, FALSE);		// 작은 아이콘을 설정합니다.

	// TODO: 여기에 추가 초기화 작업을 추가합니다.

	return TRUE;  // 포커스를 컨트롤에 설정하지 않으면 TRUE를 반환합니다.
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

// 대화 상자에 최소화 단추를 추가할 경우 아이콘을 그리려면
//  아래 코드가 필요합니다. 문서/뷰 모델을 사용하는 MFC 응용 프로그램의 경우에는
//  프레임워크에서 이 작업을 자동으로 수행합니다.

void CExcel2PDWDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 그리기를 위한 디바이스 컨텍스트입니다.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 클라이언트 사각형에서 아이콘을 가운데에 맞춥니다.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 아이콘을 그립니다.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 사용자가 최소화된 창을 끄는 동안에 커서가 표시되도록 시스템에서
//  이 함수를 호출합니다.
HCURSOR CExcel2PDWDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CExcel2PDWDlg::OnBnClickedButtonOpenfile()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	CString strFilter( "엑셀 파일|*.xlsx|PDW 파일|*.pdw|모든파일(*.*)|*.*|" );
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
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
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
	// TODO: 여기에 명령 처리기 코드를 추가합니다.


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
