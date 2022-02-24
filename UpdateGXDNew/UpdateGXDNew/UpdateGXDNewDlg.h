#pragma once
#include "msolib.h"
#include "EasySize.h"
#include "Registry.h"
#include "ColorEdit.h"
#include "XPStyleButtonST.h"
#include "TextProgressCtrl.h"
#include "CDownloadFileSever.h"

// CUpdateGXDNewDlg dialog
class CUpdateGXDNewDlg : public CDialogEx
{
	DECLARE_EASYSIZE

// Construction
public:
	CUpdateGXDNewDlg(CWnd* pParent = nullptr);	// standard constructor
	~CUpdateGXDNewDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_UPDATEGXDNEW_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

// Implementation
protected:

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);

	DECLARE_MESSAGE_MAP()

public:
	Base64Ex *bb;
	CRegistry *reg;
	CDownloadFileSever *sv;

	CXPStyleButtonST m_btn_ok, m_btn_cancel;
	CThemeHelperST m_ThemeHelper;
	CTextProgressCtrl m_progressUpdate;
	CColorEdit m_txt_noidung;
	CColorEdit m_lbl_warning;
	CButton m_check_skipup;
	CFont m_myFont;
	
	CString pathFolder;
	CString szOffice;
	int iCheckUpSoft;

public:
	void _SetDefault();
	void _SaveRegistry();
	void _LoadLogUpdate();	
	void _QuitApplication(CString szAppName);

	CString _GetPathFolder(CString fileName);
	CString _GetNameOfPath(CString pathFile, int ipath = 0);
	int _ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine);

	BOOL _CreateNewFolder(CString directoryPath);
	bool _CreateDirectories(CString pathName);
	void _DownloadFileSmartStart();

	afx_msg void OnBnClickedButtonOk();
	afx_msg void OnBnClickedButtonCancel();

	afx_msg void OnBnClickedCheckSkipupdate();
};
