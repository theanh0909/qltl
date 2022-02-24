
// UpdatesmDlg.h : header file
//

#pragma once

// CUpdatesmDlg dialog
class CUpdatesmDlg : public CDialogEx
{
// Construction
public:
	CUpdatesmDlg(CWnd* pParent = nullptr);	// standard constructor
	~CUpdatesmDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_UPDATESM_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	DECLARE_MESSAGE_MAP()

};
