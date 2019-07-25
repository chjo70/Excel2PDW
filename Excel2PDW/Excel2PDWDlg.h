
// Excel2PDWDlg.h : ��� ����
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
		unsigned fdiff			 : 3;		// ���ļ� �ִ� ���̰� (<- Reserved)
		unsigned blank_tag	 : 1;		// Blanking Tag
		unsigned pw					 : 14; 	// [0 ~ 16,383] Res. 20ns (0 ~ 327.660 us)

		// Phase 3
		unsigned amplitude	 : 8;  	// [0 ~ 255] Res. 0.3125 dBm (-74.6875 ~ + 5 dBm)
		unsigned filter_tag	 : 8;  	// ��ó������ ��ȣ (<- Reserved)
		unsigned aoa				 : 9;  	// [0 ~ 511] Res. 0.703125 �� (0 ~ 359.297 ��)
		unsigned dv					 : 1;  	// [Valid | Invalid ]
		unsigned max_channel : 2;		// Ch1(0x0), Ch2(0x1), Ch3(0x2), Ch4(0x3)
		unsigned stat			   : 4;  	// PMOP(0x8), FMOP(0x4), CW(0x2), Pulse(0x1), Short Pulse(0xF)


		// Phase 4 for dummy
		unsigned flag        : 28;	// ���� �Ϸ� �÷���
		unsigned band        : 4;

	} item ;

} _PDW;

typedef struct {			// frequency band code�� ���� ����ü 
	UINT min;				// min frequency
	UINT max;				// max frequency
	int offset;       // max frequency
	float res;				// �� ������ ���� resolution

} FREQ_RESOL ;

typedef struct {
	float offset;      // max frequency
	float res;			// �� ������ ���� resolution
} PA_RESOL ;


// CExcel2PDWDlg ��ȭ ����
class CExcel2PDWDlg : public CDialogEx
{
// �����Դϴ�.
public:
	CExcel2PDWDlg(CWnd* pParent = NULL);	// ǥ�� �������Դϴ�.

	void GetCellValue( CXlSimpleAutomation *pXL, UINT unCol, long lRow, UINT *pToa, double *pPw, double *pPa, double *pFreq );
	UINT EncodePA( double dValue );
	UINT EncodeFreq( double dValue );
	UINT EncodeBand( double dValue );
	UINT EncodePW( double dValue );
	UINT EncodeTOA( double dValue );

// ��ȭ ���� �������Դϴ�.
	enum { IDD = IDD_EXCEL2PDW_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV �����Դϴ�.


private:
	CString m_strFilename;


// �����Դϴ�.
protected:
	HICON m_hIcon;

	// ������ �޽��� �� �Լ�
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
