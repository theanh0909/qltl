#pragma once
#include "StdAfx.h"

class DlgCreateNewTemp : public CDialogEx
{
	DECLARE_DYNAMIC(DlgCreateNewTemp)

public:
	DlgCreateNewTemp(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgCreateNewTemp();

	// Dialog Data
	enum { IDD = DLG_CREATE_NEWTEMP };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CColorEdit m_txt_name;
	CXPStyleButtonST m_btn_ok, m_btn_cancel;
	CThemeHelperST m_ThemeHelper;

	Function *ff;

	_bstr_t bsFManager;
	_WorksheetPtr psFManager;
	RangePtr prFManager;

	int icotSTT, icotLV, icotMH, icotTenFile, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;
	int icotFolder, icotTen, icotType;

	struct MyStrSheet
	{
		CString szName;
		CString szCode;
		CString szKodau;
		CString szLinhvuc;
	};

public:
	void _XacdinhSheetLienquan();
	bool _CreateNewSheet(CString szNameSheet);	
	void _CopyStyleCategory(_WorksheetPtr pSheet);

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnEnChangeTxtName();
};
