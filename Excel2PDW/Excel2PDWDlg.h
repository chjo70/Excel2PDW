
// Excel2PDWDlg.h : 헤더 파일
//

#pragma once

#include "XlSimpleAutomation.h"
#include "afxwin.h"

#define PDW_NORMAL          (1)
#define PDW_CW              (2)
#define PDW_FMOP						(5)
#define PDW_CW_FMOP					(6)
#define PDW_SHORTP          (7)
#define PDW_ALL							(8)

typedef union {
	unsigned int wpdw[4];
	unsigned char bpdw[ 4 ][ 4 ];

	struct pdw_phase {
		// Phase 1
		unsigned toa				 : 32; 	// [0 ~ 4,294,967,295] Res. 20ns (0 ~ 85,899,345.900 us)

		// Phase 2
		unsigned freq				 : 14; 	// [0 ~ 8,191] Res. 1.25 MHz
		unsigned fdiff			 : 3;		// 주파수 최대 변이값 (<- Reserved)
		unsigned blank_tag	 : 1;		// Blanking Tag
		unsigned pw					 : 14; 	// [0 ~ 16,383] Res. 20ns (0 ~ 327.660 us)

		// Phase 3
		unsigned amplitude	 : 8;  	// [0 ~ 255] Res. 0.3125 dBm (-74.6875 ~ + 5 dBm)
		unsigned filter_tag	 : 8;  	// 전처리필터 번호 (<- Reserved)
		unsigned aoa				 : 9;  	// [0 ~ 511] Res. 0.703125 도 (0 ~ 359.297 도)
		unsigned dv					 : 1;  	// [Valid | Invalid ]
		unsigned max_channel : 2;		// Ch1(0x0), Ch2(0x1), Ch3(0x2), Ch4(0x3)
		unsigned stat			   : 4;  	// PMOP(0x8), FMOP(0x4), CW(0x2), Pulse(0x1), Short Pulse(0xF)


		// Phase 4 for dummy
		unsigned flag        : 28;	// 수집 완료 플레그
		unsigned band        : 4;

	} item ;

} _PDW;

typedef struct {			// frequency band code를 위한 구조체 
	UINT min;				// min frequency
	UINT max;				// max frequency
	int offset;       // max frequency
	float res;				// 각 구간에 따른 resolution

} FREQ_RESOL ;

typedef struct {
	float offset;      // max frequency
	float res;			// 각 구간에 따른 resolution
} PA_RESOL ;


// CExcel2PDWDlg 대화 상자
class CExcel2PDWDlg : public CDialogEx
{
// 생성입니다.
public:
	CExcel2PDWDlg(CWnd* pParent = NULL);	// 표준 생성자입니다.

	void GetCellValue( CXlSimpleAutomation *pXL, UINT unCol, long lRow, UINT *pToa, double *pPw, double *pPa, double *pFreq );
	UINT EncodePA( double dValue );
	UINT EncodeFreq( double dValue );
	UINT EncodeBand( double dValue );
	UINT EncodePW( double dValue );
	UINT EncodeTOA( double dValue );

// 대화 상자 데이터입니다.
	enum { IDD = IDD_EXCEL2PDW_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 지원입니다.


private:
	CString m_strFilename;


// 구현입니다.
protected:
	HICON m_hIcon;

	// 생성된 메시지 맵 함수
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtonOpenfile();
	afx_msg void OnBnClickedButtonConvert();
	afx_msg void OnStaticOpenFileName();
	CEdit m_CEditOpenFile;
	CButton m_CButtonConvert;
};
